"""Выгрузка сырых данных из МойСклад в один словарь."""
import re
from datetime import datetime, timedelta

from .api import ApiError, MoySklad

FMT = "%Y-%m-%d %H:%M:%S"
UUID = re.compile(r"[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}")


def _id(holder) -> str | None:
    m = UUID.search((holder or {}).get("meta", {}).get("href", ""))
    return m.group(0) if m else None


def _docs_with_positions(ms: MoySklad, kind: str, since: str) -> list[dict]:
    docs = ms.rows(f"entity/{kind}", {"filter": f"moment>={since}", "expand": "positions.assortment"})
    for d in docs:
        pos = d.get("positions") or {}
        if pos.get("meta", {}).get("size", 0) > len(pos.get("rows", [])):  # в документе больше 100 позиций
            pos["rows"] = ms.rows(f"entity/{kind}/{d['id']}/positions", {"expand": "assortment"})
    return [d for d in docs if d.get("applicable") is not False]


def sales_from_docs(demands, returns, costs, since: str, store_id=None, agent_ids=None) -> list[dict]:
    """Строки в формате отчёта «Прибыльность», собранные из отгрузок и возвратов покупателей."""
    out = {}
    for docs, sign in ((demands, 1), (returns, -1)):
        for d in docs:
            if d["moment"][:19] < since or (store_id and _id(d.get("store")) != store_id):
                continue
            if agent_ids is not None and _id(d.get("agent")) not in agent_ids:
                continue
            rate = (d.get("rate") or {}).get("value") or 1
            for p in (d.get("positions") or {}).get("rows", []):
                a = p.get("assortment") or {}
                pid = _id(a)
                if not pid or a.get("meta", {}).get("type") == "service":
                    continue
                row = out.setdefault(pid, {
                    "assortment": {"meta": a["meta"], "name": a.get("name", ""), "code": a.get("code", ""),
                                   "article": a.get("article", "")},
                    "sellQuantity": 0, "sellSum": 0, "sellCostSum": 0, "returnQuantity": 0, "returnSum": 0,
                    "returnCostSum": 0, "profit": 0})
                qty = p.get("quantity") or 0
                amount = qty * (p.get("price") or 0) * (1 - (p.get("discount") or 0) / 100) * rate
                cost = qty * (costs.get(pid) or (a.get("buyPrice") or {}).get("value") or 0)
                side = "sell" if sign > 0 else "return"
                row[side + "Quantity"] += qty
                row[side + "Sum"] += amount
                row[side + "CostSum"] += cost
                row["profit"] += sign * (amount - cost)
    return list(out.values())


def fetch_all(ms: MoySklad, cfg: dict, today: datetime, log=print) -> dict:
    t_to = today.strftime(FMT)
    t30 = (today - timedelta(days=30)).strftime(FMT)
    t90 = (today - timedelta(days=90)).strftime(FMT)
    t365 = (today - timedelta(days=365)).strftime(FMT)

    last_sale = {}
    log("Склады и контрагенты…")
    stores = ms.rows("entity/store")
    counterparties = [{"id": c["id"], "name": c.get("name", ""), "tags": c.get("tags", []), "meta": c["meta"]}
                      for c in ms.rows("entity/counterparty")]
    bazaar = [c for c in counterparties
              if any(w.lower() in c["name"].lower() for w in cfg["bazaar"]["agents"])]
    bazaar_ids = {c["id"] for c in bazaar}
    try:
        balances = {_id(r.get("counterparty")): r.get("balance") or 0 for r in ms.rows("report/counterparty")}
    except ApiError as e:
        if e.status != 403:
            raise
        log("  нет прав на отчёт по контрагентам — балансы заводов будут пустыми")
        balances = None
    log("Остатки…")
    stock_all = ms.rows("report/stock/all", {"stockMode": "all"})
    stock_bystore = ms.rows("report/stock/bystore")

    log("Продажи за 30 и 90 дней…")
    try:
        prof = lambda since, flt=None: ms.rows("report/profit/byvariant", {  # noqa: E731
            "momentFrom": since, "momentTo": t_to, **({"filter": flt} if flt else {})})
        sales30, sales90 = prof(t30), prof(t90)
        by_store = {st["id"]: prof(t90, f"store={st['meta']['href']}") for st in stores}
        baz30 = [r for c in bazaar for r in prof(t30, f"counterparty={c['meta']['href']}")]
        baz90 = [r for c in bazaar for r in prof(t90, f"counterparty={c['meta']['href']}")]
        source = "отчёт «Прибыльность»"
    except ApiError as e:
        if e.status != 403:
            raise
        log("  нет прав на отчёт «Прибыльность» — считаю продажи по отгрузкам и возвратам")
        costs = {_id(r): r.get("price") or 0 for r in stock_all}
        docs = [_docs_with_positions(ms, "demand", t90) + _docs_with_positions(ms, "retaildemand", t90),
                _docs_with_positions(ms, "salesreturn", t90) + _docs_with_positions(ms, "retailsalesreturn", t90)]
        sales30 = sales_from_docs(*docs, costs, t30)
        sales90 = sales_from_docs(*docs, costs, t90)
        by_store = {st["id"]: sales_from_docs(*docs, costs, t90, st["id"]) for st in stores}
        baz30 = sales_from_docs(*docs, costs, t30, agent_ids=bazaar_ids)
        baz90 = sales_from_docs(*docs, costs, t90, agent_ids=bazaar_ids)
        source = "отгрузки и возвраты (себестоимость — текущая из остатков)"
        last_sale = {}
        for d in docs[0]:
            for p in (d.get("positions") or {}).get("rows", []):
                pid = _id(p.get("assortment"))
                if pid and d["moment"] > last_sale.get(pid, ""):
                    last_sale[pid] = d["moment"]

    log("Приёмки и платежи за год…")
    supplies = ms.rows("entity/supply", {"filter": f"moment>={t365}"})
    payments = []
    for kind in ("paymentout", "cashout"):
        payments += ms.rows(f"entity/{kind}", {"filter": f"moment>={t365}", "expand": "expenseItem"})
    bazaar_in = 0
    for c in bazaar:
        for kind in ("paymentin", "cashin"):
            bazaar_in += sum(x.get("sum") or 0 for x in ms.rows(
                f"entity/{kind}", {"filter": f"moment>={t90};agent={c['meta']['href']}"}))

    log("Заказы поставщикам…")
    orders = ms.rows("entity/purchaseorder", {"filter": f"created>={cfg['orders']['since']} 00:00:00", "expand": "state"})
    for o in orders:
        o["positions_rows"] = ms.rows(f"entity/purchaseorder/{o['id']}/positions")

    log("Продажи по клиентам за два года…")
    from .sales import first_month, yoy_start
    s0 = first_month(today).strftime(FMT)
    s2 = yoy_start(today).strftime(FMT)
    sales_docs = {k: [{"moment": d["moment"], "sum": d.get("sum") or 0, "agent": {"meta": {"href": d["agent"]["meta"]["href"]}},
                       "applicable": d.get("applicable")} for d in rows_by_month(ms, k, yoy_start(today), today)]
                  for k in ("demand", "salesreturn", "retaildemand", "retailsalesreturn")}

    pay_in = [{"moment": d["moment"], "sum": d.get("sum") or 0, "agent": {"meta": {"href": d["agent"]["meta"]["href"]}},
               "applicable": d.get("applicable")}
              for k in ("paymentin", "cashin") for d in ms.rows(f"entity/{k}", {"filter": f"moment>={s0}"})]

    return {
        "last_sale": last_sale, "pay_in": pay_in, "sales_docs": sales_docs, "orders": orders, "stores": stores, "counterparties": counterparties, "balances": balances, "stock_all": stock_all,
        "stock_bystore": stock_bystore, "sales30": sales30, "sales90": sales90, "sales90_by_store": by_store,
        "bazaar30": baz30, "bazaar90": baz90, "bazaar_in90": bazaar_in, "bazaar_ids": sorted(bazaar_ids),
        "sales_source": source, "supplies": supplies, "payments": payments,
    }


def rows_by_month(ms: MoySklad, kind: str, since: datetime, until: datetime, params: dict | None = None) -> list[dict]:
    """Документы помесячно: короткие запросы без глубокого листания, на больших аккаунтах это в разы надёжнее."""
    out, a = [], datetime(since.year, since.month, 1)
    while a <= until:
        b = datetime(a.year + (a.month == 12), a.month % 12 + 1, 1)
        lo = max(a, since)
        out += ms.rows(f"entity/{kind}", {**(params or {}),
                                          "filter": f"moment>={lo.strftime(FMT)};moment<{b.strftime(FMT)}"})
        a = b
    return out


def fetch_sales_only(ms: MoySklad, cfg: dict, today: datetime, log=print) -> dict:
    """Только то, что нужно для отчётов «Продажи по клиентам» и «Дебиторка» (для второго аккаунта)."""
    from .sales import first_month, yoy_start
    log("Контрагенты и балансы…")
    counterparties = [{"id": c["id"], "name": c.get("name", ""), "tags": c.get("tags", []), "meta": c["meta"]}
                      for c in ms.rows("entity/counterparty", {"filter": "archived=true;archived=false"})]
    try:
        balances = {_id(r.get("counterparty")): r.get("balance") or 0 for r in ms.rows("report/counterparty")}
    except ApiError as e:
        if e.status != 403:
            raise
        log("  нет прав на отчёт по контрагентам — дебиторки не будет")
        balances = None
    log("Продажи за два года…")
    s0, s2 = first_month(today).strftime(FMT), yoy_start(today).strftime(FMT)
    pick = lambda d: {"moment": d["moment"], "sum": d.get("sum") or 0,  # noqa: E731
                      "agent": {"meta": {"href": d["agent"]["meta"]["href"]}}, "applicable": d.get("applicable")}
    forbidden = []

    def docs(kind, since):
        try:
            return [pick(d) for d in rows_by_month(ms, kind, datetime.strptime(since, FMT), today)]
        except ApiError as e:
            if e.status != 403:
                raise
            forbidden.append(kind)
            log(f"  нет прав на {kind} — пропускаю")
            return []

    if not cfg.get("_yoy", True):
        s2 = s0
    sales_docs = {k: docs(k, s2) for k in ("demand", "salesreturn", "retaildemand", "retailsalesreturn")}
    pay_in = [d for k in ("paymentin", "cashin") for d in docs(k, s0)]
    return {"counterparties": counterparties, "balances": balances, "sales_docs": sales_docs, "pay_in": pay_in,
            "forbidden": forbidden}
