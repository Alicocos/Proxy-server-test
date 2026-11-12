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


def sales_from_docs(demands, returns, costs, since: str, store_id=None) -> list[dict]:
    """Строки в формате отчёта «Прибыльность», собранные из отгрузок и возвратов покупателей."""
    out = {}
    for docs, sign in ((demands, 1), (returns, -1)):
        for d in docs:
            if d["moment"][:19] < since or (store_id and _id(d.get("store")) != store_id):
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
    d = cfg["documents"]
    t_to = today.strftime(FMT)
    t30 = (today - timedelta(days=30)).strftime(FMT)
    t90 = (today - timedelta(days=90)).strftime(FMT)

    log("Склады и контрагенты…")
    stores = ms.rows("entity/store")
    agents = {c["id"]: c.get("name", "") for c in ms.rows("entity/counterparty")}
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
        sales30 = ms.rows("report/profit/byvariant", {"momentFrom": t30, "momentTo": t_to})
        sales90 = ms.rows("report/profit/byvariant", {"momentFrom": t90, "momentTo": t_to})
        by_store = {st["id"]: ms.rows("report/profit/byvariant", {
            "momentFrom": t90, "momentTo": t_to, "filter": f"store={st['meta']['href']}"}) for st in stores}
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
        source = "отгрузки и возвраты (себестоимость — текущая из остатков)"

    log("Заказы поставщикам, приёмки, платежи…")
    since = (today - timedelta(days=d["orders_lookback_days"])).strftime(FMT)
    orders = ms.rows("entity/purchaseorder", {"filter": f"moment>={since}", "expand": "state,rate.currency"})
    for o in orders:  # позиции заказов — чтобы считать товар в пути
        o["positions_rows"] = ms.rows(f"entity/purchaseorder/{o['id']}/positions")
    supplies = ms.rows("entity/supply", {"filter": f"moment>={since}"})
    pay_since = (today - timedelta(days=d["payments_days"])).strftime(FMT)
    payments = []
    for kind in ("paymentout", "cashout"):
        payments += ms.rows(f"entity/{kind}", {"filter": f"moment>={pay_since}", "expand": "rate.currency,operations"})

    return {
        "stores": stores, "agents": agents, "balances": balances, "stock_all": stock_all,
        "stock_bystore": stock_bystore, "sales30": sales30, "sales90": sales90, "sales90_by_store": by_store,
        "sales_source": source, "orders": orders, "supplies": supplies, "payments": payments,
    }
