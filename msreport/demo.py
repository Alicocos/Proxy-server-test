"""Выдуманные данные в формате ответов API МойСклад — для проверки отчёта без токена."""
import random
import uuid
from datetime import datetime, timedelta

H = "https://api.moysklad.ru/api/remap/1.2/entity"


def _meta(kind, id_):
    return {"meta": {"href": f"{H}/{kind}/{id_}", "type": kind}}


def demo_raw(today: datetime) -> dict:
    rnd = random.Random(7)
    uid = lambda: str(uuid.UUID(int=rnd.getrandbits(128)))  # noqa: E731
    f = lambda d: (today - timedelta(days=d)).strftime("%Y-%m-%d %H:%M:%S.000")  # noqa: E731

    names = ["Основной склад 1", "Основной склад 2"] + [f"Базар, ангар {i}" for i in range(1, 8)] + ["Склад брак"]
    stores = [{"id": uid(), "name": n} for n in names]
    for s in stores:
        s.update(_meta("store", s["id"]))

    groups = ["Посуда", "Текстиль", "Инструмент", "Хозтовары", "Электрика"]
    stock_all, bystore, s30, s90 = [], [], [], []
    by_store = {s["id"]: [] for s in stores}
    for n in range(150):
        pid, cost = uid(), rnd.randint(200, 9000) * 100
        kind = rnd.choices(["fast", "normal", "slow", "dead", "new", "out"], [20, 40, 15, 12, 5, 8])[0]
        daily = {"fast": rnd.uniform(5, 20), "normal": rnd.uniform(0.5, 4), "slow": rnd.uniform(0.02, 0.3),
                 "dead": 0, "new": 0, "out": rnd.uniform(1, 6)}[kind]
        cover = {"fast": rnd.uniform(10, 90), "normal": rnd.uniform(20, 200), "slow": rnd.uniform(200, 900),
                 "out": rnd.uniform(0, 8)}.get(kind)
        total = round(daily * cover) if cover is not None else rnd.randint(20, 400)
        a = {"name": f"{rnd.choice(groups)} арт.{1000 + n}", "code": str(50000 + n), "article": f"CN-{1000 + n}"}
        assort = {**_meta("product", pid), **a}
        # остаток: половина на основных складах, остальное неравномерно по ангарам
        parts, left = [], total
        for s in stores:
            q = round(left * rnd.choice([0, 0, 0.2, 0.5])) if s is not stores[-1] else left
            left -= q
            parts.append({**_meta("store", s["id"]), "name": s["name"], "stock": q, "reserve": 0, "inTransit": 0})
        stock_all.append({
            "meta": {"href": f"{H}/product/{pid}?expand=supplier", "type": "product"}, **a,
            "stock": total, "reserve": rnd.choice([0, 0, 0, 5]),
            "inTransit": rnd.choice([0, 0, 0, round(daily * 40)]), "price": cost, "salePrice": round(cost * 1.5),
            "stockDays": 12 if kind == "new" else rnd.randint(40, 400),
            "folder": {"name": a["name"].split()[0], "pathName": "Китай"},
        })
        bystore.append({"meta": {"href": f"{H}/product/{pid}?expand=supplier"}, "stockByStore": parts})
        if daily:
            q90, q30 = round(daily * 90), round(daily * 30 * rnd.uniform(0.6, 1.4))
            row = lambda q: {"assortment": assort, "sellQuantity": q, "returnQuantity": 0, "sellSum": q * cost * 1.5,  # noqa: E731
                             "returnSum": 0, "sellCostSum": q * cost, "returnCostSum": 0, "profit": q * cost * 0.5}
            if q90:
                s90.append(row(q90))
                sellers = rnd.sample(stores, rnd.randint(1, 4))
                for s in sellers:
                    by_store[s["id"]].append(row(max(1, q90 // len(sellers))))
            if q30:
                s30.append(row(q30))

    factories = [{"id": uid(), "name": n, "tags": ["поставщики китай"]} for n in (
        "Yiwu Hongda Trading", "Guangzhou Meili Textile", "Ningbo Tools Co", "Foshan Ceramics", "Shenzhen Elec Ltd")]
    baz = {"id": uid(), "name": "Строи сити склад", "tags": ["базар"]}
    other = {"id": uid(), "name": "Аренда ангара", "tags": []}
    agent = lambda c: {"agent": _meta("counterparty", c["id"])}  # noqa: E731
    payments, supplies = [], []
    for fac in factories:
        for _ in range(rnd.randint(4, 12)):
            payments.append({"meta": {"type": "paymentout"}, "name": f"{rnd.randint(1, 9999):05d}",
                             "moment": f(rnd.randint(1, 360)), "sum": rnd.randint(20, 300) * 10000000, **agent(fac),
                             "expenseItem": {"name": "Закупка товаров"}, "description": "оплата заводу"})
        for _ in range(rnd.randint(3, 10)):
            supplies.append({"moment": f(rnd.randint(70 if fac is factories[0] else 1, 360)),
                             "sum": rnd.randint(20, 300) * 10000000, **agent(fac)})
    for n in range(40):
        payments.append({"meta": {"type": "paymentout"}, "name": f"L{n:04d}", "moment": f(rnd.randint(1, 360)),
                         "sum": rnd.randint(5, 60) * 1000000, **agent(other),
                         "expenseItem": {"name": rnd.choice(["Доставка", "Доставка завод", "Аренда"])}})
    balances = {fc["id"]: rnd.choice([-1, -1, 1]) * rnd.randint(5, 90) * 100000000 for fc in factories}
    balances[baz["id"]] = 45000000000
    bz = lambda rows: [{**r, "sellQuantity": r["sellQuantity"] // 3, "sellSum": r["sellSum"] // 3,  # noqa: E731
                        "sellCostSum": r["sellCostSum"] // 3, "profit": r["profit"] // 3} for r in rows[::2]]
    return {"counterparties": factories + [baz, other], "balances": balances, "supplies": supplies,
            "payments": payments, "bazaar30": bz(s30), "bazaar90": bz(s90), "bazaar_in90": 9000000000,
            "bazaar_ids": [baz["id"]], "sales_source": "демо",
            "stores": stores, "stock_all": stock_all, "stock_bystore": bystore, "sales30": s30, "sales90": s90,
            "sales90_by_store": by_store}


def demo_logistics(today: datetime) -> list[dict]:
    d = lambda n: today - timedelta(days=n)  # noqa: E731
    base = {"places": 40, "weight": 1200.0, "volume": 6.5, "cn_carrier": "Deppon", "cn_track": "DPK77120045",
            "cn_cost": 1800.0, "cn_cur": "CNY", "ext_carrier": "Карго Алатау", "ext_route": "Гуанчжоу – Хоргос – Алматы",
            "ext_cost": 2600.0, "ext_cur": "USD", "ext_paid": 0.0, "comment": ""}
    return [
        {**base, "shipment": "L-101", "order": "00301", "factory": "Yiwu Hongda Trading", "goods": "Посуда",
         "cn_ship": d(40), "cn_arrive": d(36), "ext_ship": d(33), "eta": d(8), "fact": None},
        {**base, "shipment": "L-102", "order": "00305", "factory": "Ningbo Tools Co", "goods": "Инструмент",
         "cn_ship": d(20), "cn_arrive": d(17), "ext_ship": d(14), "eta": d(-10), "fact": None, "ext_paid": 2600.0},
        {**base, "shipment": "L-103", "order": "00309", "factory": "Foshan Ceramics", "goods": "Керамика",
         "cn_ship": d(6), "cn_arrive": d(2), "ext_ship": None, "eta": None, "fact": None, "ext_cost": 0.0},
        {**base, "shipment": "L-104", "order": "00311", "factory": "Shenzhen Elec Ltd", "goods": "Электрика",
         "cn_ship": d(3), "cn_arrive": None, "ext_ship": None, "eta": None, "fact": None, "ext_cost": 0.0},
        {**base, "shipment": "L-100", "order": "00298", "factory": "Guangzhou Meili Textile", "goods": "Текстиль",
         "cn_ship": d(70), "cn_arrive": d(66), "ext_ship": d(62), "eta": d(35), "fact": d(30), "ext_paid": 2600.0},
    ]


def demo_orders() -> list[dict]:
    return [{"code": "50003", "qty": 500, "eta": None}, {"code": "50010", "qty": 500, "eta": None},
            {"code": "50141", "qty": 1000, "eta": None}, {"code": "99999", "qty": 200, "eta": None}]
