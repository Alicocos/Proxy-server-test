"""Расчёты: анализ закупа, заморозка, перемещения, заказы заводам, платежи, логистика."""
import math
import re
from collections import defaultdict
from datetime import datetime

UUID = re.compile(r"[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}")

URGENT, ORDER, OK, OVER, DEAD, NEW, TRANSIT = (
    "Срочно заказать", "Заказать", "Норма", "Излишек", "Неликвид", "Новый приход", "В пути")
STATUS_ORDER = {URGENT: 0, ORDER: 1, OVER: 2, DEAD: 3, OK: 4, NEW: 5, TRANSIT: 6}


def _id(meta_holder: dict) -> str | None:
    href = (meta_holder or {}).get("meta", {}).get("href", "")
    m = UUID.search(href)
    return m.group(0) if m else None


def _money(v) -> float:
    return (v or 0) / 100


def _dt(s) -> datetime | None:
    if not s:
        return None
    if isinstance(s, datetime):
        return s
    try:
        return datetime.strptime(str(s)[:19], "%Y-%m-%d %H:%M:%S")
    except ValueError:
        return None


def _sales(rows: list[dict]) -> dict:
    out = {}
    for r in rows:
        a = r.get("assortment") or {}
        pid = _id(a)
        if not pid:
            continue
        cur = out.setdefault(pid, {"qty": 0.0, "rev": 0.0, "profit": 0.0, "cogs": 0.0,
                                   "name": a.get("name", ""), "code": a.get("code", ""),
                                   "article": a.get("article", "")})
        cur["qty"] += (r.get("sellQuantity") or 0) - (r.get("returnQuantity") or 0)
        cur["rev"] += _money((r.get("sellSum") or 0) - (r.get("returnSum") or 0))
        cur["cogs"] += _money((r.get("sellCostSum") or 0) - (r.get("returnCostSum") or 0))
        cur["profit"] += _money(r.get("profit"))
    return out


def _is_excluded(name: str, cfg: dict) -> bool:
    return any(w.lower() in (name or "").lower() for w in cfg["purchase"].get("exclude_stores", []))


def _excluded_stock(raw: dict, cfg: dict) -> dict:
    """Остаток на исключённых складах (брак и т.п.) по каждому товару."""
    out = defaultdict(float)
    for r in raw["stock_bystore"]:
        for s in r.get("stockByStore", []):
            if _is_excluded(s.get("name"), cfg) and (s.get("stock") or 0) > 0:
                out[_id(r)] += s["stock"]
    return out


def _order_open(o: dict, cfg: dict) -> bool:
    closed = {x.lower() for x in cfg["documents"].get("closed_states", [])}
    total = _money(o.get("sum"))
    return (o.get("applicable") is not False and total > 0
            and _money(o.get("shippedSum")) < total - 0.005
            and ((o.get("state") or {}).get("name") or "").lower() not in closed)


def _order_transit(raw: dict, cfg: dict) -> dict:
    """Количество по позициям открытых заказов поставщикам, ещё не принятое на склад."""
    out = defaultdict(float)
    for o in raw["orders"]:
        if _order_open(o, cfg):
            for p in o.get("positions_rows", []):
                out[_id(p.get("assortment"))] += max(0, (p.get("quantity") or 0) - (p.get("shipped") or 0))
    return out


def build_products(raw: dict, cfg: dict) -> list[dict]:
    p = cfg["purchase"]
    lead, target = p["lead_time_days"], p["lead_time_days"] + p["safety_days"] + p["order_cycle_days"]
    s30, s90 = _sales(raw["sales30"]), _sales(raw["sales90"])

    excluded, ordered = _excluded_stock(raw, cfg), _order_transit(raw, cfg)
    items = {}
    for r in raw["stock_all"]:
        pid = _id(r)
        if not pid:
            continue
        folder = r.get("folder") or {}
        items[pid] = {
            "id": pid, "name": r.get("name", ""), "code": r.get("code", ""), "article": r.get("article", ""),
            "folder": "/".join(x for x in (folder.get("pathName"), folder.get("name")) if x),
            "stock": (r.get("stock") or 0) - excluded.get(pid, 0), "reserve": r.get("reserve") or 0, "transit": max(r.get("inTransit") or 0, ordered.get(pid, 0)),
            "cost": _money(r.get("price")), "sale_price": _money(r.get("salePrice")),
            "stock_days": r.get("stockDays"),
        }
    for pid, s in s90.items():  # проданное, но отсутствующее в отчёте остатков
        items.setdefault(pid, {"id": pid, "name": s["name"], "code": s["code"], "article": s["article"],
                               "folder": "", "stock": 0, "reserve": 0, "transit": 0,
                               "cost": s["cogs"] / s["qty"] if s["qty"] > 0 else 0,
                               "sale_price": 0, "stock_days": None})

    out = []
    for pid, it in items.items():
        q30 = max(s30.get(pid, {}).get("qty", 0), 0)
        q90 = max(s90.get(pid, {}).get("qty", 0), 0)
        daily = 0.5 * q30 / 30 + 0.5 * q90 / 90
        stock, transit, cost = max(it["stock"], 0), max(it["transit"], 0), it["cost"]
        if stock <= 0 and transit <= 0 and daily == 0:
            continue
        rec, frozen = 0, 0.0
        if daily == 0:
            if stock > 0:
                sd = it["stock_days"]
                status = NEW if sd is not None and sd < p["new_arrival_days"] else DEAD
                frozen = stock * cost if status == DEAD else 0.0
            else:
                status = TRANSIT
        else:
            rec = max(0, math.ceil(daily * target - stock - transit))
            if (stock + transit) / daily < lead:
                status = URGENT
            elif rec > 0:
                status = ORDER
            elif stock / daily > p["overstock_days"]:
                status = OVER
                frozen = (stock - daily * p["overstock_days"]) * cost
            else:
                status = OK
        it.update({
            "q30": q30, "q90": q90, "daily": daily,
            "rev90": s90.get(pid, {}).get("rev", 0.0), "profit90": s90.get(pid, {}).get("profit", 0.0),
            "cogs90": s90.get(pid, {}).get("cogs", 0.0),
            "cover": stock / daily if daily else None,
            "cover_transit": (stock + transit) / daily if daily else None,
            "stock_value": stock * cost, "transit_value": transit * cost,
            "rec": rec, "rec_value": rec * cost, "status": status, "frozen": frozen, "abc": "",
        })
        out.append(it)

    total = sum(i["rev90"] for i in out if i["rev90"] > 0)
    acc = 0.0
    for i in sorted(out, key=lambda x: -x["rev90"]):
        if i["rev90"] <= 0 or total == 0:
            i["abc"] = "-"
            continue
        i["abc"] = "A" if acc / total < 0.8 else "B" if acc / total < 0.95 else "C"
        acc += i["rev90"]
    out.sort(key=lambda x: (STATUS_ORDER[x["status"]], -x["rec_value"], -x["frozen"], -x["stock_value"]))
    return out


def build_stores(raw: dict, products: list[dict], cfg: dict):
    """Свод по складам и предложения по перемещениям между ними."""
    t = cfg["transfers"]
    by_id = {p["id"]: p for p in products}
    store_names = {s["id"]: s["name"] for s in raw["stores"]}
    stock = defaultdict(dict)  # pid -> store_id -> qty
    for r in raw["stock_bystore"]:
        pid = _id(r)
        for s in r.get("stockByStore", []):
            sid = _id(s)
            store_names.setdefault(sid, s.get("name", sid))
            if s.get("stock"):
                stock[pid][sid] = s["stock"]
    sold = defaultdict(dict)
    for sid, rows in raw["sales90_by_store"].items():
        for pid, s in _sales(rows).items():
            if s["qty"] > 0:
                sold[pid][sid] = s["qty"]

    summary = {sid: {"store": n, "stock_value": 0.0, "cogs90": 0.0, "no_sales_value": 0.0, "sku": 0}
               for sid, n in store_names.items()}
    transfers = []
    for pid in set(stock) | set(sold):
        prod = by_id.get(pid)
        cost = prod["cost"] if prod else 0
        donors, takers = [], []
        for sid in set(stock[pid]) | set(sold[pid]):
            qty, q90 = max(stock[pid].get(sid, 0), 0), sold[pid].get(sid, 0)
            daily = q90 / 90
            sm = summary[sid]
            sm["stock_value"] += qty * cost
            sm["cogs90"] += q90 * cost
            if qty > 0:
                sm["sku"] += 1
                if q90 == 0:
                    sm["no_sales_value"] += qty * cost
            cover = qty / daily if daily else None
            if _is_excluded(store_names[sid], cfg):
                continue
            if daily and cover < t["deficit_days"]:
                takers.append((sid, qty, q90, cover, math.ceil(daily * t["target_days"] - qty)))
            elif qty > 0 and (not daily or cover > t["surplus_days"]):
                donors.append([sid, qty, q90, cover, math.floor(qty - daily * t["keep_days"])])
        if not prod:
            continue
        donors.sort(key=lambda d: -d[4])
        for sid, qty, q90, cover, need in sorted(takers, key=lambda x: -x[4]):
            for d in donors:
                move = min(need, d[4])
                if move <= 0:
                    continue
                transfers.append({
                    "name": prod["name"], "code": prod["code"], "article": prod["article"],
                    "to": store_names[sid], "to_stock": qty, "to_q90": q90, "to_cover": cover,
                    "from": store_names[d[0]], "from_stock": d[1], "from_q90": d[2], "from_cover": d[3],
                    "qty": move, "value": move * cost,
                })
                d[4] -= move
                need -= move
                if need <= 0:
                    break
    for sm in summary.values():
        sm["turnover"] = sm["stock_value"] / (sm["cogs90"] / 90) if sm["cogs90"] > 0 else None
        sm["no_sales_share"] = sm["no_sales_value"] / sm["stock_value"] if sm["stock_value"] else 0
    stores = sorted((s for s in summary.values() if s["stock_value"] or s["cogs90"]),
                    key=lambda s: -s["stock_value"])
    transfers.sort(key=lambda x: -x["value"])
    return stores, transfers


def _currency(doc: dict):
    rate = doc.get("rate") or {}
    cur = rate.get("currency") or {}
    return cur.get("isoCode") or cur.get("name") or "", rate.get("value") or 1


def build_orders(raw: dict, cfg: dict, today: datetime):
    names = raw.get("agents") or {}
    blank = lambda name: {"agent": name, "orders": 0, "sum_base": 0.0, "to_pay_base": 0.0, "overdue_orders": 0,  # noqa: E731
                          "supplied_base": 0.0, "supplies": 0, "payments_base": 0.0, "payments": 0,
                          "deposit": None, "debt": None}
    suppliers, orders = {}, []

    def sup(doc):
        agent = doc.get("agent") or {}
        aid = _id(agent) or "?"
        name = names.get(aid) or agent.get("name") or "?"
        return aid, name, suppliers.setdefault(aid, blank(name))

    for o in raw["orders"]:
        total = _money(o.get("sum"))
        if o.get("applicable") is False or total <= 0:
            continue
        paid, recv = _money(o.get("payedSum")), _money(o.get("shippedSum"))
        cur, rate = _currency(o)
        aid, name, s = sup(o)
        moment, planned = _dt(o.get("moment")), _dt(o.get("deliveryPlannedMoment"))
        state = (o.get("state") or {}).get("name", "")
        received_all = not _order_open(o, cfg)
        row = {
            "name": o.get("name", ""), "moment": moment, "agent": name, "state": state, "currency": cur,
            "sum": total, "paid": paid, "received": recv, "to_pay": max(0.0, total - paid),
            "prepaid": max(0.0, paid - recv), "paid_share": min(paid / total, 1.0), "planned": planned,
            "age": (today - moment).days if moment else None,
            "overdue": (today - planned).days if planned and not received_all and planned < today else 0,
            "open": not received_all, "description": o.get("description", ""),
            "sum_base": total * rate, "to_pay_base": max(0.0, total - paid) * rate,
            "flags": "; ".join(f for f, bad in (("нет плановой даты", not planned), ("нет привязанных оплат", paid == 0),
                                                ("нет статуса", not state)) if bad),
        }
        orders.append(row)
        if row["open"]:
            s["orders"] += 1
            s["sum_base"] += row["sum_base"]
            s["to_pay_base"] += row["to_pay_base"]
            s["overdue_orders"] += 1 if row["overdue"] > 0 else 0

    supplies = [x for x in raw.get("supplies", []) if x.get("applicable") is not False]
    for x in supplies:
        _, _, s = sup(x)
        s["supplied_base"] += _money(x.get("sum")) * _currency(x)[1]
        s["supplies"] += 1

    payments = []
    for p in raw["payments"]:
        aid = _id(p.get("agent") or {}) or "?"
        if aid not in suppliers or p.get("applicable") is False:
            continue
        cur, rate = _currency(p)
        amount = _money(p.get("sum"))
        ops = p.get("operations") or []
        payments.append({
            "moment": _dt(p.get("moment")), "name": p.get("name", ""), "agent": suppliers[aid]["agent"],
            "kind": "Наличные" if p.get("meta", {}).get("type") == "cashout" else "Безнал",
            "currency": cur, "sum": amount, "sum_base": amount * rate,
            "linked": ", ".join(x.get("name", "") for x in ops if x.get("name")) or ("да" if ops else ""),
            "purpose": p.get("paymentPurpose") or p.get("description") or "",
        })
        suppliers[aid]["payments_base"] += amount * rate
        suppliers[aid]["payments"] += 1
    payments.sort(key=lambda x: x["moment"] or datetime.min, reverse=True)

    balances = raw.get("balances")
    if balances is not None:  # в МойСклад баланс < 0: контрагент должен нам (предоплата); > 0: должны мы
        for aid, s in suppliers.items():
            bal = _money(balances.get(aid))
            s["deposit"], s["debt"] = max(0.0, -bal), max(0.0, bal)

    open_orders = sorted((o for o in orders if o["open"]), key=lambda o: (-o["overdue"], -o["sum_base"]))
    sup_rows = sorted(suppliers.values(), key=lambda s: (-(s["deposit"] or 0), -(s["debt"] or 0), -s["payments_base"]))
    quality = {
        "orders": len(open_orders),
        "orders_no_pay": sum(1 for o in open_orders if o["paid"] == 0),
        "orders_no_plan": sum(1 for o in open_orders if not o["planned"]),
        "supplies": len(supplies),
        "supplies_unlinked": sum(1 for x in supplies if "purchaseOrder" not in x),
        "payments": len(payments),
        "payments_unlinked": sum(1 for p in payments if not p["linked"]),
    }
    return open_orders, sup_rows, payments, quality


def build_logistics(rows: list[dict], today: datetime) -> list[dict]:
    out = []
    for r in rows:
        r = dict(r)
        arrived, ext_ship, eta = r.get("fact"), r.get("ext_ship"), r.get("eta")
        if arrived:
            r["status"] = "Прибыл"
        elif ext_ship:
            r["status"] = "В пути (внешняя)"
        elif r.get("cn_arrive"):
            r["status"] = "На складе карго"
        elif r.get("cn_ship"):
            r["status"] = "В пути по Китаю"
        else:
            r["status"] = "Ждёт отгрузки с завода"
        start = r.get("cn_ship") or ext_ship
        r["days"] = ((arrived or today) - start).days if start else None
        r["overdue"] = (today - eta).days if eta and not arrived and eta < today else 0
        cn, ext = r.get("cn_cost") or 0, r.get("ext_cost") or 0
        r["ext_unpaid"] = max(0.0, ext - (r.get("ext_paid") or 0))
        vol, wt = r.get("volume") or 0, r.get("weight") or 0
        r["per_m3"] = ext / vol if vol and ext else None
        r["per_kg"] = ext / wt if wt and ext else None
        r["cn_cost"], r["ext_cost"] = cn, ext
        out.append(r)
    rank = {"В пути (внешняя)": 0, "На складе карго": 1, "В пути по Китаю": 2, "Ждёт отгрузки с завода": 3, "Прибыл": 4}
    out.sort(key=lambda r: (rank[r["status"]], -r["overdue"]))
    return out


def build_report(raw: dict, logistics_rows: list[dict], cfg: dict, today: datetime) -> dict:
    products = build_products(raw, cfg)
    stores, transfers = build_stores(raw, products, cfg)
    orders, suppliers, payments, quality = build_orders(raw, cfg, today)
    logistics = build_logistics(logistics_rows, today)

    by_cost = {p["id"]: p["cost"] for p in products}
    stock_value = sum(p["stock_value"] for p in products)
    cogs90 = sum(p["cogs90"] for p in products)
    frozen = sum(p["frozen"] for p in products)
    active = [r for r in logistics if r["status"] != "Прибыл"]
    unpaid = defaultdict(float)
    for r in logistics:
        if r["ext_unpaid"]:
            unpaid[r.get("ext_cur") or "?"] += r["ext_unpaid"]
    last30 = sum(p["sum_base"] for p in payments if p["moment"] and (today - p["moment"]).days <= 30)
    kpi = {
        "stock_value": stock_value,
        "sku_in_stock": sum(1 for p in products if p["stock"] > 0),
        "turnover_days": stock_value / (cogs90 / 90) if cogs90 > 0 else None,
        "frozen": frozen, "frozen_share": frozen / stock_value if stock_value else 0,
        "dead_value": sum(p["frozen"] for p in products if p["status"] == DEAD),
        "dead_sku": sum(1 for p in products if p["status"] == DEAD),
        "over_value": sum(p["frozen"] for p in products if p["status"] == OVER),
        "over_sku": sum(1 for p in products if p["status"] == OVER),
        "transit_value": sum(p["transit_value"] for p in products),
        "urgent_sku": sum(1 for p in products if p["status"] == URGENT),
        "rec_value": sum(p["rec_value"] for p in products),
        "rec_sku": sum(1 for p in products if p["rec"] > 0),
        "orders_open": len(orders),
        "orders_sum": sum(o["sum_base"] for o in orders),
        "orders_to_pay": sum(o["to_pay_base"] for o in orders),
        "deposits": sum(s["deposit"] for s in suppliers) if raw.get("balances") is not None else None,
        "debts": sum(s["debt"] for s in suppliers) if raw.get("balances") is not None else None,
        "excluded_value": sum(q * by_cost.get(pid, 0) for pid, q in _excluded_stock(raw, cfg).items()),
        "sales_source": raw.get("sales_source", ""),
        "orders_overdue": sum(1 for o in orders if o["overdue"] > 0),
        "payments_30": last30,
        "ship_active": len(active),
        "ship_overdue": sum(1 for r in active if r["overdue"] > 0),
        "ship_unpaid": ", ".join(f"{v:,.0f} {c}".replace(",", " ") for c, v in unpaid.items()) or "0",
        "transfers": len(transfers), "transfers_value": sum(t["value"] for t in transfers),
    }
    return {"today": today, "kpi": kpi, "quality": quality, "products": products, "stores": stores, "transfers": transfers,
            "orders": orders, "suppliers": suppliers, "payments": payments, "logistics": logistics}
