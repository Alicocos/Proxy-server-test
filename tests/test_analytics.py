import tomllib
from datetime import datetime, timedelta
from pathlib import Path

from msreport import analytics as a
from msreport import demo, excel, manual

ROOT = Path(__file__).resolve().parents[1]
CFG = tomllib.loads((ROOT / "config.toml").read_text(encoding="utf-8"))
TODAY = datetime(2026, 11, 12, 10, 0, 0)
H = "https://api.moysklad.ru/api/remap/1.2/entity"
PID = "11111111-1111-1111-1111-111111111111"
S1, S2 = "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa", "bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb"


def raw(stock=0, transit=0, q30=0, q90=0, stock_days=200, cost=1000):
    assort = {"meta": {"href": f"{H}/product/{PID}"}, "name": "Товар", "code": "1", "article": "A"}
    sale = lambda q: [{"assortment": assort, "sellQuantity": q, "sellSum": q * cost * 150,  # noqa: E731
                       "sellCostSum": q * cost * 100, "profit": q * cost * 50}] if q else []
    return {
        "stores": [], "stock_bystore": [], "sales90_by_store": {}, "orders": [], "payments": [],
        "stock_all": [{"meta": {"href": f"{H}/product/{PID}?expand=supplier"}, "name": "Товар", "stock": stock,
                       "inTransit": transit, "price": cost * 100, "stockDays": stock_days}],
        "sales30": sale(q30), "sales90": sale(q90),
    }


def one(**kw):
    return a.build_products(raw(**kw), CFG)[0]


def test_statuses():
    # 1 шт/день; срок поставки 45, цель 45+14+30=89 дней
    assert one(stock=20, q30=30, q90=90)["status"] == a.URGENT
    p = one(stock=60, q30=30, q90=90)
    assert p["status"] == a.ORDER and p["rec"] == 29 and p["rec_value"] == 29000
    assert one(stock=60, transit=40, q30=30, q90=90)["status"] == a.OK
    p = one(stock=280, q30=30, q90=90)
    assert p["status"] == a.OVER and p["frozen"] == 100 * 1000  # сверх 180 дней
    p = one(stock=50)
    assert p["status"] == a.DEAD and p["frozen"] == 50000
    assert one(stock=50, stock_days=10)["status"] == a.NEW
    assert a.build_products(raw(), CFG) == []  # пустая карточка в отчёт не идёт


def test_orders_and_payments():
    agent = {"meta": {"href": f"{H}/counterparty/{S1}"}}
    r = raw()
    r["agents"] = {S1: "Завод"}
    r["balances"] = {S1: -21000000}
    r["orders"] = [
        {"name": "1", "moment": "2026-10-01 00:00:00.000", "sum": 1000000, "payedSum": 300000, "shippedSum": 0,
         "agent": agent, "rate": {"currency": {"isoCode": "CNY"}, "value": 70},
         "deliveryPlannedMoment": "2026-11-02 00:00:00.000"},
        {"name": "2", "moment": "2026-09-01 00:00:00.000", "sum": 500000, "payedSum": 0, "shippedSum": 0,
         "agent": agent, "state": {"name": "Полный келди"}},
        {"name": "3", "moment": "2026-09-01 00:00:00.000", "sum": 500000, "payedSum": 0, "shippedSum": 0, "agent": agent},
    ]
    r["supplies"] = [{"sum": 500000, "agent": agent}, {"sum": 100000, "agent": agent, "purchaseOrder": {}}]
    r["payments"] = [
        {"meta": {"type": "paymentout"}, "moment": "2026-10-02 00:00:00.000", "sum": 300000, "agent": agent,
         "rate": {"currency": {"isoCode": "CNY"}, "value": 70}, "operations": [{"name": "1"}]},
        {"meta": {"type": "paymentout"}, "moment": "2026-10-03 00:00:00.000", "sum": 999,
         "agent": {"meta": {"href": f"{H}/counterparty/{S2}"}}},
    ]
    orders, sup, pays, q = a.build_orders(r, CFG, TODAY)
    assert [o["name"] for o in orders] == ["1", "3"]  # заказ в закрытом статусе не показываем
    o = orders[0]
    assert (o["to_pay"], o["overdue"], o["to_pay_base"], o["flags"]) == (7000, 10, 490000, "нет статуса")
    assert "нет плановой даты" in orders[1]["flags"] and "нет привязанных оплат" in orders[1]["flags"]
    assert len(pays) == 1 and pays[0]["linked"] == "1" and pays[0]["agent"] == "Завод"  # платёж не заводу отфильтрован
    assert (sup[0]["deposit"], sup[0]["debt"], sup[0]["payments_base"], sup[0]["supplied_base"]) == (210000, 0, 210000, 6000)
    assert q == {"orders": 2, "orders_no_pay": 1, "orders_no_plan": 1, "supplies": 2, "supplies_unlinked": 1,
                 "payments": 1, "payments_unlinked": 0}


def test_sales_from_docs_and_excluded_store():
    from msreport.fetch import sales_from_docs
    assort = {"meta": {"href": f"{H}/product/{PID}", "type": "product"}, "name": "Товар"}
    doc = lambda m, q, sid: {"moment": m, "store": {"meta": {"href": f"{H}/store/{sid}"}},  # noqa: E731
                             "positions": {"rows": [{"quantity": q, "price": 150000, "discount": 10, "assortment": assort}]}}
    dem = [doc("2026-11-01 10:00:00.000", 10, S1), doc("2026-09-01 10:00:00.000", 5, S2)]
    ret = [doc("2026-11-05 10:00:00.000", 2, S1)]
    rows = sales_from_docs(dem, ret, {PID: 100000}, "2026-08-14 00:00:00")
    assert (rows[0]["sellQuantity"], rows[0]["returnQuantity"], rows[0]["sellSum"]) == (15, 2, 15 * 135000)
    assert rows[0]["profit"] == 13 * 35000
    assert sales_from_docs(dem, ret, {}, "2026-10-13 00:00:00")[0]["sellQuantity"] == 10
    assert sales_from_docs(dem, ret, {}, "2026-08-14 00:00:00", S2)[0]["sellQuantity"] == 5

    r = raw(stock=20, q30=30, q90=90)  # без заказа — «срочно», с открытым заказом на 80 шт — нет
    pos = [{"assortment": {"meta": {"href": f"{H}/product/{PID}"}}, "quantity": 100, "shipped": 20}]
    r["orders"] = [{"sum": 100, "shippedSum": 20, "positions_rows": pos}]
    p = a.build_products(r, CFG)[0]
    assert (p["transit"], p["status"]) == (80, a.OK)
    r["orders"][0]["state"] = {"name": "полный келди"}
    assert a.build_products(r, CFG)[0]["status"] == a.URGENT

    r = raw(stock=100, q30=30, q90=90)
    r["stock_bystore"] = [{"meta": {"href": f"{H}/product/{PID}"}, "stockByStore": [
        {"meta": {"href": f"{H}/store/{S1}"}, "name": "Основной", "stock": 60},
        {"meta": {"href": f"{H}/store/{S2}"}, "name": "Брак склад", "stock": 40}]}]
    assert a.build_products(r, CFG)[0]["stock"] == 60


def test_transfers():
    r = raw(stock=200, q30=30, q90=90)
    store = lambda sid, n, q: {"meta": {"href": f"{H}/store/{sid}"}, "name": n, "stock": q}  # noqa: E731
    r["stores"] = [{"id": S1, "name": "Основной"}, {"id": S2, "name": "Ангар"}]
    r["stock_bystore"] = [{"meta": {"href": f"{H}/product/{PID}"},
                           "stockByStore": [store(S1, "Основной", 195), store(S2, "Ангар", 5)]}]
    r["sales90_by_store"] = {S1: [], S2: r["sales90"]}
    stores, tr = a.build_stores(r, a.build_products(r, CFG), CFG)
    assert len(tr) == 1 and tr[0]["from"] == "Основной" and tr[0]["to"] == "Ангар" and tr[0]["qty"] == 25
    main = next(s for s in stores if s["store"] == "Основной")
    assert main["no_sales_value"] == 195 * 1000


def test_logistics_statuses():
    d = lambda n: TODAY - timedelta(days=n)  # noqa: E731
    rows = a.build_logistics([
        {"shipment": "a", "cn_ship": d(30), "ext_ship": d(20), "eta": d(5), "ext_cost": 100.0, "ext_paid": 40.0, "volume": 2.0},
        {"shipment": "b", "cn_ship": d(3)},
        {"shipment": "c", "cn_ship": d(30), "ext_ship": d(20), "eta": d(5), "fact": d(4)},
    ], TODAY)
    by = {r["shipment"]: r for r in rows}
    assert (by["a"]["status"], by["a"]["overdue"], by["a"]["ext_unpaid"], by["a"]["per_m3"]) == ("В пути (внешняя)", 5, 60, 50)
    assert by["b"]["status"] == "В пути по Китаю"
    assert by["c"]["status"] == "Прибыл" and by["c"]["overdue"] == 0 and by["c"]["days"] == 26


def test_demo_end_to_end(tmp_path):
    tpl = tmp_path / "in" / "logistics.xlsx"
    assert manual.read_logistics(tpl) == [] and tpl.exists()  # шаблон создаётся и читается
    rep = a.build_report(demo.demo_raw(TODAY), demo.demo_logistics(TODAY), CFG, TODAY)
    k = rep["kpi"]
    assert k["stock_value"] > 0 and 0 < k["frozen_share"] < 1 and k["orders_open"] > 0 and k["ship_overdue"] == 1
    assert abs(sum(s["stock_value"] for s in rep["stores"]) - k["stock_value"] - k["excluded_value"]) < 1
    out = excel.write_report(rep, tmp_path / "r.xlsx")
    assert out.stat().st_size > 10000
