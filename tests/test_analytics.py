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


def test_money():
    ag = lambda i: {"meta": {"href": f"{H}/counterparty/{i}"}}  # noqa: E731
    r = raw()
    r["counterparties"] = [{"id": S1, "name": "Завод", "tags": ["Поставщики Китай"]},
                           {"id": S2, "name": "Строи сити склад", "tags": ["поставщики китай"]},
                           {"id": PID, "name": "Аренда", "tags": []}]
    r["bazaar_ids"] = [S2]
    r["balances"] = {S1: -21000000, S2: 99900}
    pay = lambda m, s, a, item: {"meta": {"type": "paymentout"}, "moment": m, "sum": s, "agent": ag(a),  # noqa: E731
                                 "expenseItem": {"name": item}}
    r["payments"] = [pay("2026-11-02 00:00:00.000", 300000, S1, "Закупка товаров"),
                     pay("2026-06-02 00:00:00.000", 700000, S1, "Закупка товаров"),
                     pay("2026-11-03 00:00:00.000", 5000, PID, "Доставка"),
                     pay("2026-11-03 00:00:00.000", 999, PID, "Аренда"),
                     pay("2026-11-03 00:00:00.000", 999, S2, "Закупка товаров")]
    r["supplies"] = [{"moment": "2026-06-10 00:00:00.000", "sum": 500000, "agent": ag(S1)}]
    sup, reg, monthly = a.build_money(r, CFG, TODAY)
    assert len(sup) == 1  # базар и аренда заводами не считаются
    s = sup[0]
    assert (s["deposit"], s["debt"], s["pay30"], s["pay365"], s["sup365"]) == (210000, 0, 3000, 10000, 5000)
    assert s["days_no_sup"] == 155 and s["flag"] == "Предоплата без движения"
    assert [x["sum"] for x in reg] == [3000]  # в реестре только последние 90 дней
    nov = next(m for m in monthly if m["month"] == "2026-11")
    assert (nov["paid"], nov["logistics"]) == (3000, 50)


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

    assert sales_from_docs(dem, ret, {}, "2026-08-14 00:00:00", agent_ids={S1}) == []

    r = raw(stock=100, q30=30, q90=90)
    r["stock_bystore"] = [{"meta": {"href": f"{H}/product/{PID}"}, "stockByStore": [
        {"meta": {"href": f"{H}/store/{S1}"}, "name": "Основной", "stock": 60},
        {"meta": {"href": f"{H}/store/{S2}"}, "name": "Брак склад", "stock": 40}]}]
    assert a.build_products(r, CFG)[0]["stock"] == 60


def test_manual_orders(tmp_path):
    path = tmp_path / "orders.xlsx"
    assert manual.read_orders(path) == [] and path.exists()
    from openpyxl import load_workbook
    wb = load_workbook(path)
    wb["Заказы"].append(["Z1", datetime(2026, 10, 1), "Завод", 1, None, None, 100, 20, 5, "CNY", datetime(2026, 11, 2)])
    wb["Заказы"].append(["Z1", None, "Завод", None, "нет такого", None, 50])
    wb.save(path)
    rows = manual.read_orders(path)
    r = raw(stock=20, q30=30, q90=90)
    r["stock_all"][0]["code"] = "1"
    transit, out = a.match_orders(r, rows, TODAY)
    assert transit == {PID: 80}
    assert [(o["status"], o["left"]) for o in out] == [("Товар не найден в МойСклад", 50), ("Просрочен", 80)]
    assert out[1]["overdue"] == 10 and out[1]["left_sum"] == 400
    assert a.build_products(r, CFG)[0]["status"] == a.URGENT
    p = a.build_products(r, CFG, transit)[0]
    assert (p["transit"], p["status"]) == (80, a.OK)


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
    rep = a.build_report(demo.demo_raw(TODAY), demo.demo_logistics(TODAY), CFG, TODAY, demo.demo_orders(TODAY))
    assert (rep["kpi"]["ord_lines"], rep["kpi"]["ord_unmatched"], rep["kpi"]["ord_overdue"]) == (4, 1, 1)
    k = rep["kpi"]
    assert k["stock_value"] > 0 and 0 < k["frozen_share"] < 1 and k["deposits"] > 0 and k["baz_sum"] > 0 and len(rep["monthly"]) > 3 and k["ship_overdue"] == 1
    assert abs(sum(s["stock_value"] for s in rep["stores"]) - k["stock_value"] - k["excluded_value"]) < 1
    out = excel.write_report(rep, tmp_path / "r.xlsx")
    assert out.stat().st_size > 10000
