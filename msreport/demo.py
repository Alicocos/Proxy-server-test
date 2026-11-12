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

    names = ["Основной склад 1", "Основной склад 2"] + [f"Базар, ангар {i}" for i in range(1, 9)]
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

    cny = {"currency": {"isoCode": "CNY", "name": "юань"}, "value": 72.5}
    usd = {"currency": {"isoCode": "USD", "name": "доллар"}, "value": 510.0}
    factories = [{"id": uid(), "name": n} for n in (
        "Yiwu Hongda Trading", "Guangzhou Meili Textile", "Ningbo Tools Co", "Foshan Ceramics", "Shenzhen Elec Ltd")]
    orders, payments = [], []
    for n in range(22):
        fac, rate = rnd.choice(factories), rnd.choice([cny, cny, usd])
        total = rnd.randint(20, 400) * 100000
        age = rnd.randint(3, 150)
        paid = round(total * rnd.choice([0, 0.3, 0.3, 1.0, 1.0]))
        recv = round(total * (rnd.choice([1, 1, 0.6]) if age > 70 and paid == total else 0))
        orders.append({
            "name": f"{300 + n:05d}", "moment": f(age), "sum": total, "payedSum": paid, "shippedSum": recv,
            "applicable": True, "agent": {**_meta("counterparty", fac["id"]), "name": fac["name"]},
            "state": {"name": rnd.choice(["Новый", "В производстве", "Отгружен", "В пути"])}, "rate": rate,
            "deliveryPlannedMoment": f(age - 55), "description": "",
        })
        if paid and age <= 90:
            payments.append({
                "meta": {"type": rnd.choice(["paymentout", "paymentout", "cashout"])}, "name": f"{900 + n:05d}",
                "moment": f(max(age - 2, 0)), "sum": paid, "rate": rate, "applicable": True,
                "agent": {**_meta("counterparty", fac["id"]), "name": fac["name"]},
                "operations": [{"name": f"{300 + n:05d}"}], "paymentPurpose": "Оплата по заказу",
            })
    payments.append({"meta": {"type": "paymentout"}, "name": "00999", "moment": f(5), "sum": 5000000,
                     "agent": {**_meta("counterparty", uid()), "name": "Аренда ангара"}, "operations": []})

    return {"stores": stores, "stock_all": stock_all, "stock_bystore": bystore, "sales30": s30, "sales90": s90,
            "sales90_by_store": by_store, "orders": orders, "payments": payments}


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
