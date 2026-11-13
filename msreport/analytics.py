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


def build_orders(raw: dict, cfg: dict, today: datetime):
    """Открытые заказы поставщикам МойСклад → товар в пути по позициям, ближайший приход и строки для листа."""
    closed = {x.lower() for x in cfg["orders"].get("closed_states", [])}
    names = {c["id"]: c["name"] for c in raw.get("counterparties", [])}
    transit, etas, rows = defaultdict(float), {}, []
    for o in raw.get("orders", []):
        state = ((o.get("state") or {}).get("name") or "")
        if o.get("applicable") is False or state.lower() in closed:
            continue
        rate = (o.get("rate") or {}).get("value") or 1
        left_qty = left_sum = ordered = 0.0
        planned = _dt(o.get("deliveryPlannedMoment"))
        for p in o.get("positions_rows", []):
            pid = _id(p.get("assortment"))
            left = max(0.0, (p.get("quantity") or 0) - (p.get("shipped") or 0))
            ordered += p.get("quantity") or 0
            if not pid or not left:
                continue
            transit[pid] += left
            if planned and (pid not in etas or planned < etas[pid]):
                etas[pid] = planned
            left_qty += left
            left_sum += left * _money(p.get("price")) * (1 - (p.get("discount") or 0) / 100) * rate
        if not left_qty:
            continue
        moment = _dt(o.get("moment"))
        total = _money(o.get("sum")) * rate
        rows.append({
            "name": o.get("name", ""), "moment": moment, "agent": names.get(_id(o.get("agent")), ""),
            "state": state, "sum": total, "paid": _money(o.get("payedSum")) * rate,
            "ordered": ordered, "left_qty": left_qty, "left_sum": left_sum, "planned": planned,
            "age": (today - moment).days if moment else None,
            "overdue": (today - planned).days if planned and planned < today else 0,
            "status": "Просрочен" if planned and planned < today else "Нет плановой даты" if not planned else "В пути",
            "description": o.get("description", ""),
        })
    rows.sort(key=lambda r: (-r["overdue"], r["planned"] or datetime.max))
    return transit, etas, rows


def build_products(raw: dict, cfg: dict, ordered: dict | None = None) -> list[dict]:
    p = cfg["purchase"]
    lead, target = p["lead_time_days"], p["lead_time_days"] + p["safety_days"] + p["order_cycle_days"]
    s30, s90 = _sales(raw["sales30"]), _sales(raw["sales90"])
    b30, b90 = _sales(raw.get("bazaar30", [])), _sales(raw.get("bazaar90", []))

    excluded = _excluded_stock(raw, cfg)
    items = {}
    for r in raw["stock_all"]:
        pid = _id(r)
        if not pid:
            continue
        folder = r.get("folder") or {}
        items[pid] = {
            "id": pid, "name": r.get("name", ""), "code": r.get("code", ""), "article": r.get("article", ""),
            "folder": "/".join(x for x in (folder.get("pathName"), folder.get("name")) if x),
            "stock": (r.get("stock") or 0) - excluded.get(pid, 0), "reserve": r.get("reserve") or 0, "transit": max(r.get("inTransit") or 0, (ordered or {}).get(pid, 0)),
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
            "baz30": b30.get(pid, {}).get("qty", 0), "baz90": b90.get(pid, {}).get("qty", 0),
            "baz_sum90": b90.get(pid, {}).get("rev", 0.0),
            "baz_share": min(1.0, max(0.0, b90.get(pid, {}).get("qty", 0) / q90)) if q90 else None,
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


def build_money(raw: dict, cfg: dict, today: datetime):
    """Заводы по платежам и приёмкам: депозиты, долги, реестр платежей, помесячная динамика."""
    sc = cfg["suppliers"]
    tags = {t.lower() for t in sc["tags"]}
    bazaar = set(raw.get("bazaar_ids", []))
    logi = {x.lower() for x in cfg["logistics"]["expense_items"]}
    sup = {c["id"]: {"agent": c["name"], "deposit": None, "debt": None, "pay30": 0.0, "pay90": 0.0, "pay365": 0.0,
                     "sup90": 0.0, "sup365": 0.0, "payments": 0, "supplies": 0, "last_pay": None, "last_sup": None,
                     "flag": ""}
           for c in raw.get("counterparties", [])
           if c["id"] not in bazaar and tags & {t.lower() for t in c.get("tags", [])}}
    months = defaultdict(lambda: {"paid": 0.0, "received": 0.0, "logistics": 0.0})
    register = []

    for p in raw["payments"]:
        when = _dt(p.get("moment"))
        if p.get("applicable") is False or not when:
            continue
        amount = _money(p.get("sum")) * ((p.get("rate") or {}).get("value") or 1)
        item = (p.get("expenseItem") or {}).get("name", "")
        age = (today - when).days
        if item.lower() in logi:
            months[when.strftime("%Y-%m")]["logistics"] += amount
        s = sup.get(_id(p.get("agent") or {}))
        if not s:
            continue
        months[when.strftime("%Y-%m")]["paid"] += amount
        s["pay365"] += amount
        s["pay90"] += amount if age <= 90 else 0
        s["pay30"] += amount if age <= 30 else 0
        s["payments"] += 1
        s["last_pay"] = max(s["last_pay"] or when, when)
        if age <= sc["payments_days"]:
            register.append({"moment": when, "name": p.get("name", ""), "agent": s["agent"], "sum": amount,
                             "kind": "Наличные" if p.get("meta", {}).get("type") == "cashout" else "Безнал",
                             "item": item, "purpose": p.get("paymentPurpose") or p.get("description") or ""})
    for x in raw.get("supplies", []):
        when = _dt(x.get("moment"))
        s = sup.get(_id(x.get("agent") or {}))
        if x.get("applicable") is False or not when or not s:
            continue
        amount = _money(x.get("sum")) * ((x.get("rate") or {}).get("value") or 1)
        months[when.strftime("%Y-%m")]["received"] += amount
        s["sup365"] += amount
        s["sup90"] += amount if (today - when).days <= 90 else 0
        s["supplies"] += 1
        s["last_sup"] = max(s["last_sup"] or when, when)

    balances = raw.get("balances")
    for aid, s in sup.items():
        if balances is not None:  # в МойСклад баланс < 0: контрагент должен нам (предоплата); > 0: должны мы
            bal = _money(balances.get(aid))
            s["deposit"], s["debt"] = max(0.0, -bal), max(0.0, bal)
        s["days_no_sup"] = (today - s["last_sup"]).days if s["last_sup"] else None
        if s["deposit"] and (s["days_no_sup"] is None or s["days_no_sup"] > sc["stale_days"]):
            s["flag"] = "Предоплата без движения"
    rows = sorted((s for s in sup.values() if s["deposit"] or s["debt"] or s["pay365"] or s["sup365"]),
                  key=lambda s: (-(s["deposit"] or 0), -(s["debt"] or 0), -s["pay365"]))
    register.sort(key=lambda x: x["moment"], reverse=True)
    monthly = [{"month": m, **v, "diff": v["paid"] - v["received"],
                "logi_share": v["logistics"] / v["received"] if v["received"] else None}
               for m, v in sorted(months.items())]
    return rows, register, monthly


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
    ordered, etas, orders = build_orders(raw, cfg, today)
    products = build_products(raw, cfg, ordered)
    for p in products:
        p["eta"] = etas.get(p["id"]) if p["transit"] else None
    stores, transfers = build_stores(raw, products, cfg)
    suppliers, payments, monthly = build_money(raw, cfg, today)
    logistics = build_logistics(logistics_rows, today)

    by_cost = {p["id"]: p["cost"] for p in products}
    stock_value = sum(p["stock_value"] for p in products)
    cogs90 = sum(p["cogs90"] for p in products)
    rev90 = sum(p["rev90"] for p in products)
    frozen = sum(p["frozen"] for p in products)
    active = [r for r in logistics if r["status"] != "Прибыл"]
    unpaid = defaultdict(float)
    for r in logistics:
        if r["ext_unpaid"]:
            unpaid[r.get("ext_cur") or "?"] += r["ext_unpaid"]
    has_bal = raw.get("balances") is not None
    baz_ids = raw.get("bazaar_ids", [])
    baz_sum = sum(p["baz_sum90"] for p in products)
    last3 = monthly[-3:]
    kpi = {
        "stock_value": stock_value,
        "sku_in_stock": sum(1 for p in products if p["stock"] > 0),
        "turnover_days": stock_value / (cogs90 / 90) if cogs90 > 0 else None,
        "frozen": frozen, "frozen_share": frozen / stock_value if stock_value else 0,
        "dead_value": sum(p["frozen"] for p in products if p["status"] == DEAD),
        "dead_sku": sum(1 for p in products if p["status"] == DEAD),
        "over_value": sum(p["frozen"] for p in products if p["status"] == OVER),
        "over_sku": sum(1 for p in products if p["status"] == OVER),
        "excluded_value": sum(q * by_cost.get(pid, 0) for pid, q in _excluded_stock(raw, cfg).items()),
        "transfers": len(transfers), "transfers_value": sum(t["value"] for t in transfers),
        "transit_value": sum(p["transit_value"] for p in products),
        "urgent_sku": sum(1 for p in products if p["status"] == URGENT),
        "rec_value": sum(p["rec_value"] for p in products),
        "rec_sku": sum(1 for p in products if p["rec"] > 0),
        "deposits": sum(s["deposit"] for s in suppliers) if has_bal else None,
        "debts": sum(s["debt"] for s in suppliers) if has_bal else None,
        "stale": sum(s["deposit"] for s in suppliers if s["flag"]) if has_bal else None,
        "stale_n": sum(1 for s in suppliers if s["flag"]),
        "pay30": sum(s["pay30"] for s in suppliers), "pay90": sum(s["pay90"] for s in suppliers),
        "sup90": sum(s["sup90"] for s in suppliers),
        "logi3": sum(m["logistics"] for m in last3),
        "logi3_share": (sum(m["logistics"] for m in last3) / sum(m["received"] for m in last3)
                        if sum(m["received"] for m in last3) else None),
        "baz_sum": baz_sum, "baz_share": baz_sum / rev90 if rev90 else None,
        "baz_in": _money(raw.get("bazaar_in90")),
        "baz_balance": sum(_money(raw["balances"].get(i)) for i in baz_ids) if has_bal else None,
        "baz_sku": sum(1 for p in products if p["baz90"] > 0),
        "ship_active": len(active), "ship_overdue": sum(1 for r in active if r["overdue"] > 0),
        "ship_unpaid": ", ".join(f"{v:,.0f} {c}".replace(",", " ") for c, v in unpaid.items()) or "0",
        "sales_source": raw.get("sales_source", ""),
        "transit_sku": sum(1 for p in products if p["transit"] > 0),
        "eta_overdue": sum(1 for p in products if p["transit"] > 0 and p["eta"] and p["eta"] < today),
        "orders_open": len(orders), "orders_left": sum(o["left_sum"] for o in orders),
        "orders_overdue": sum(1 for o in orders if o["overdue"]),
        "orders_noplan": sum(1 for o in orders if not o["planned"]),
        "orders_since": cfg["orders"]["since"],
    }
    return {"today": today, "kpi": kpi, "orders": orders, "products": products, "stores": stores, "transfers": transfers,
            "suppliers": suppliers, "payments": payments, "monthly": monthly, "logistics": logistics,
            "bazaar": sorted((p for p in products if p["baz90"] > 0), key=lambda p: -p["baz_sum90"])}
