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


def test_transit_from_purchase_orders():
    r = raw(stock=20, q30=30, q90=90)
    r["counterparties"] = [{"id": S1, "name": "Завод", "tags": []}]
    pos = lambda q, sh: {"assortment": {"meta": {"href": f"{H}/product/{PID}"}}, "quantity": q, "shipped": sh,  # noqa: E731
                         "price": 1000, "discount": 0}
    r["orders"] = [
        {"name": "1", "moment": "2026-11-01 00:00:00.000", "deliveryPlannedMoment": "2026-11-10 00:00:00.000",
         "sum": 10000, "agent": {"meta": {"href": f"{H}/counterparty/{S1}"}}, "rate": {"value": 70},
         "positions_rows": [pos(100, 30)]},
        {"name": "2", "moment": "2026-11-05 00:00:00.000", "positions_rows": [pos(50, 50)]},          # принят полностью
        {"name": "3", "moment": "2026-11-05 00:00:00.000", "state": {"name": "ПОЛНЫЙ КЕЛДИ"},
         "positions_rows": [pos(40, 0)]},                                                              # закрыт статусом
        {"name": "4", "moment": "2026-11-06 00:00:00.000", "applicable": False, "positions_rows": [pos(40, 0)]},
    ]
    # по умолчанию статус не важен: проведённый заказ 3 тоже в пути, непроведённый 4 — нет
    transit, _, rows = a.build_orders(r, CFG, TODAY)
    assert transit == {PID: 110} and [o["name"] for o in rows] == ["1", "3"]
    CLOSED = {**CFG, "orders": {**CFG["orders"], "closed_states": ["полный келди"]}}
    transit, etas, rows = a.build_orders(r, CLOSED, TODAY)
    assert transit == {PID: 70} and etas[PID] == datetime(2026, 11, 10)
    assert len(rows) == 1
    o = rows[0]
    assert (o["agent"], o["left_qty"], o["left_sum"], o["overdue"], o["status"]) == ("Завод", 70, 70 * 10 * 70, 2, "Просрочен")
    rep = a.build_report({**r, "supplies": []}, [], CLOSED, TODAY)
    p = rep["products"][0]
    assert (p["transit"], p["status"], p["eta"]) == (70, a.OK, datetime(2026, 11, 10))
    assert (rep["kpi"]["orders_open"], rep["kpi"]["orders_overdue"], rep["kpi"]["transit_sku"]) == (1, 1, 1)


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
    assert (rep["kpi"]["orders_open"], rep["kpi"]["orders_overdue"], rep["kpi"]["orders_noplan"]) == (4, 1, 2)
    k = rep["kpi"]
    assert k["stock_value"] > 0 and 0 < k["frozen_share"] < 1 and k["deposits"] > 0 and k["baz_sum"] > 0 and len(rep["monthly"]) > 3 and k["ship_overdue"] == 1
    assert abs(sum(s["stock_value"] for s in rep["stores"]) - k["stock_value"] - k["excluded_value"]) < 1
    out = excel.write_report(rep, tmp_path / "r.xlsx")
    assert out.stat().st_size > 10000


def test_mail_payload(tmp_path):
    from msreport import mail
    rep = a.build_report(demo.demo_raw(TODAY), demo.demo_logistics(TODAY), CFG, TODAY)
    f = excel.write_report(rep, tmp_path / "r.xlsx")
    sent = {}

    class Resp:
        status_code = 201
        def json(self): return {"messageId": "x"}

    class S:
        def post(self, url, json, timeout, headers):
            sent.update(url=url, json=json, headers=headers)
            return Resp()

    mail.send(f, rep, ["a@x.kz", "b@x.kz"], "a@x.kz", session=S())
    j = sent["json"]
    assert sent["url"] == mail.URL and sent["headers"]["api-key"] == "proxy-injected"
    assert [t["email"] for t in j["to"]] == ["a@x.kz", "b@x.kz"] and j["sender"]["email"] == "a@x.kz"
    assert j["attachment"][0]["name"] == "r.xlsx" and len(j["attachment"][0]["content"]) > 1000
    assert "Срочно заказать" in j["htmlContent"] and "00001" in j["htmlContent"]  # просроченный демо-заказ
    assert "Больше всего заморожено" in j["htmlContent"]


def test_sales_by_groups(tmp_path):
    from msreport import sales
    ag = lambda i: {"agent": {"meta": {"href": f"{H}/counterparty/{i}"}}}  # noqa: E731
    r = {"counterparties": [{"id": S1, "name": "А", "tags": ["клиенты", "Сержан клиент"]},
                            {"id": S2, "name": "Б", "tags": ["клиенты"]}, {"id": PID, "name": "В", "tags": ["сотрудники"]}],
         "sales_docs": {"demand": [
             {"moment": "2026-11-12 09:00:00.000", "sum": 100000, **ag(S1)},   # сегодня
             {"moment": "2026-11-11 09:00:00.000", "sum": 50000, **ag(S1)},    # вчера
             {"moment": "2026-11-02 09:00:00.000", "sum": 30000, **ag(S2)},    # месяц с начала
             {"moment": "2026-10-05 09:00:00.000", "sum": 70000, **ag(S2)},    # прошлый месяц за те же дни
             {"moment": "2026-10-20 09:00:00.000", "sum": 900000, **ag(S2)},   # прошлый месяц, позже 12-го
             {"moment": "2026-11-05 09:00:00.000", "sum": 1000, **ag(PID)},
             {"moment": "2025-10-31 09:00:00.000", "sum": 999999, **ag(S1)},   # раньше окна
             {"moment": "2026-11-10 09:00:00.000", "sum": 555, "applicable": False, **ag(S1)}],
          "salesreturn": [{"moment": "2026-11-12 10:00:00.000", "sum": 20000, **ag(S1)}]}}
    rep = sales.build_sales(r, CFG, TODAY)
    assert rep["months"][0] == "2025-11" and rep["cur"] == "2026-11" and len(rep["full"]) == 12
    s, o = rep["G"]["Сержан клиенты"], rep["G"]["Общие клиенты"]
    assert (s["periods"]["today"]["net"], s["periods"]["yesterday"]["net"], s["periods"]["mtd"]["net"]) == (800, 500, 1300)
    assert (s["periods"]["mtd"]["n"], s["periods"]["mtd"]["clients"], s["total"]) == (2, 1, 0)
    assert (o["periods"]["mtd"]["net"], o["periods"]["prev_mtd"]["net"], o["total"]) == (300, 700, 9700)
    assert rep["G"]["Прочие"]["periods"]["mtd"]["net"] == 10 and "Прочие" in rep["groups"]
    assert "Данияр клиенты" in rep["groups"] and rep["G"]["Данияр клиенты"]["clients"] == 0
    html = sales.html_section(rep)
    assert "Сержан клиенты" in html and "-57%" in html  # 300 к 700 у общих клиентов
    out = sales.write_sales(rep, tmp_path / "s.xlsx")
    from openpyxl import load_workbook
    wb = load_workbook(out)
    assert {"Сводка", "Отгрузки", "Возвраты", "Сержан клиенты", "Общие клиенты", "Как читать"} <= set(wb.sheetnames)
    full = sales.build_sales(demo.demo_raw(TODAY), CFG, TODAY)
    assert full["total"] > 0 and len(full["groups"]) == 4 and "Строи сити (базар)" not in full["groups"]


def test_debts(tmp_path):
    from msreport import sales
    ag = lambda i: {"agent": {"meta": {"href": f"{H}/counterparty/{i}"}}}  # noqa: E731
    r = {"counterparties": [{"id": S1, "name": "Должник", "tags": ["клиенты", "данияр клиент"]},
                            {"id": S2, "name": "Завод", "tags": ["поставщики китай"]},
                            {"id": PID, "name": "Переплатил", "tags": ["клиенты"]}],
         "balances": {S1: -50000000, S2: -9900000, PID: 7700000},
         "sales_docs": {"demand": [{"moment": "2026-11-02 09:00:00.000", "sum": 20000000, **ag(S1)},   # 10 дней
                                   {"moment": "2026-09-23 09:00:00.000", "sum": 20000000, **ag(S1)},   # 50 дней
                                   {"moment": "2026-07-01 09:00:00.000", "sum": 30000000, **ag(S1)}]},  # 134 дня
         "pay_in": [{"moment": "2026-09-01 09:00:00.000", "sum": 5000000, **ag(S1)},
                    {"moment": "2026-08-01 09:00:00.000", "sum": 1000000, **ag(S1)}]}
    d = sales.build_debts(r, CFG, TODAY)
    assert d["n"] == 1 and d["total"] == 500000  # завод и переплативший клиент — не должники
    x = d["rows"][0]
    assert (x["group"], x["b0"], x["b1"], x["b2"], x["b3"]) == ("Данияр клиенты", 200000, 200000, 0, 100000)
    assert (x["age"], x["no_pay_days"], x["last_pay_sum"], x["flag"]) == (134, 72, 50000, "не платит больше 60 дней")
    assert d["old"] == 100000 and d["by_group"][0]["group"] == "Данияр клиенты"
    r["balances"][S1] = -90000000  # долг больше отгрузок за год
    x = sales.build_debts(r, CFG, TODAY)["rows"][0]
    assert x["age"] is None and x["age_text"] == "старше года" and x["b3"] == 500000
    assert sales.build_debts({**r, "balances": None}, CFG, TODAY) is None
    assert "Должник" in sales.debts_html(d)
    out = sales.write_debts_book(d, tmp_path / "d.xlsx")
    from openpyxl import load_workbook
    assert load_workbook(out).sheetnames == ["Сводка", "Дебиторка", "Дебиторка по группам", "Как читать"]
    r["counterparties"].append({"id": "c" * 36, "name": "Учредитель", "tags": ["учредитель"]})
    r["balances"]["c" * 36] = -70000000
    d = sales.build_debts(r, CFG, TODAY)
    assert (d["n"], d["other_total"], d["other"][0]["name"]) == (1, 700000, "Учредитель")  # не клиент — отдельно
    out = sales.write_debts_book(d, tmp_path / "d2.xlsx")
    assert "Прочие долги" in load_workbook(out).sheetnames and "700 000" in sales.debts_html(d)


def test_yoy_stockout_deposit_date(tmp_path):
    from msreport import sales
    ag = lambda i: {"agent": {"meta": {"href": f"{H}/counterparty/{i}"}}}  # noqa: E731
    r = {"counterparties": [{"id": S1, "name": "А", "tags": ["клиенты", "сержан клиент"]}],
         "sales_docs": {"demand": [
             {"moment": "2026-11-05 09:00:00.000", "sum": 300000, **ag(S1)},
             {"moment": "2025-11-03 09:00:00.000", "sum": 200000, **ag(S1)},   # те же дни прошлого года
             {"moment": "2025-11-20 09:00:00.000", "sum": 900000, **ag(S1)},   # позже 12-го — не сравниваем
             {"moment": "2026-10-10 09:00:00.000", "sum": 500000, **ag(S1)},
             {"moment": "2025-10-10 09:00:00.000", "sum": 250000, **ag(S1)},
             {"moment": "2025-01-10 09:00:00.000", "sum": 100000, **ag(S1)}]}}
    y = sales.build_yoy(r, CFG, TODAY)
    cur, octo = y["rows"][-1], y["rows"][-2]
    assert cur["partial"] and cur["ИТОГО"] == (3000, 2000) and cur["Сержан клиенты"] == (3000, 2000)
    assert octo["month"] == "2026-10" and octo["ИТОГО"] == (5000, 2500)
    assert y["rows"][0]["month"] == "2025-11" and y["rows"][0]["has_prev"] is False  # данных за ноябрь 2024 нет
    assert y["year"]["ИТОГО"][1] is None
    assert "+50%" in sales.yoy_html(y) and "+100%" in sales.yoy_html(y)
    from openpyxl import load_workbook
    out = sales.write_sales(sales.build_sales(r, CFG, TODAY), tmp_path / "s.xlsx", y)
    assert load_workbook(out).sheetnames[1] == "Сверка по месяцам"

    # стокаут: остаток 0 при продажах, плюс почти кончившийся товар
    rr = raw(stock=0, q30=30, q90=90)
    rr["last_sale"] = {PID: "2026-11-02 10:00:00.000"}
    p = a.build_products(rr, CFG)
    so = a.build_stockout(p, rr, CFG, TODAY)
    assert so[0]["so_status"] == "Нет в наличии" and so[0]["days_out"] == 10 and so[0]["lost_day"] == 1500
    assert a.build_stockout(a.build_products(raw(stock=10, q30=30, q90=90), CFG), {}, CFG, TODAY)[0]["so_status"] == "Скоро кончится"
    assert a.build_stockout(a.build_products(raw(stock=500, q30=30, q90=90), CFG), {}, CFG, TODAY) == []

    # дата предоплаты: депозит 150 покрывается двумя последними платежами
    m = raw()
    m["counterparties"] = [{"id": S1, "name": "Завод", "tags": ["поставщики китай"]}]
    m["balances"] = {S1: -15000}
    pay = lambda d, s: {"meta": {"type": "paymentout"}, "moment": d, "sum": s, **ag(S1)}  # noqa: E731
    m["payments"] = [pay("2026-11-01 00:00:00.000", 10000), pay("2026-10-01 00:00:00.000", 10000), pay("2026-08-01 00:00:00.000", 10000)]
    m["supplies"] = []
    s = a.build_money(m, CFG, TODAY)[0][0]
    assert (s["dep_since"], s["dep_days"]) == (datetime(2026, 10, 1), 42)
    m["balances"] = {S1: -99000}
    assert a.build_money(m, CFG, TODAY)[0][0]["dep_text"] == "больше года"


def test_daily_by_rep(tmp_path):
    from msreport import sales
    ag = lambda i: {"agent": {"meta": {"href": f"{H}/counterparty/{i}"}}}  # noqa: E731
    F = "f" * 8 + "-ffff-ffff-ffff-" + "f" * 12
    r = {"counterparties": [{"id": S1, "name": "Клиент", "tags": ["клиенты", "сержан клиент"]},
                            {"id": F, "name": "Фараби роз.клиенты", "tags": ["клиенты"]},
                            {"id": S2, "name": "Без менеджера", "tags": ["клиенты"]}],
         "sales_docs": {"demand": [{"moment": "2026-11-10 09:00:00.000", "sum": 100000, **ag(S1)},
                                   {"moment": "2026-11-10 10:00:00.000", "sum": 50000, **ag(F)},
                                   {"moment": "2026-11-12 10:00:00.000", "sum": 70000, **ag(S2)},
                                   {"moment": "2026-10-30 10:00:00.000", "sum": 999999, **ag(S1)}],
                        "salesreturn": [{"moment": "2026-11-12 11:00:00.000", "sum": 20000, **ag(S1)}]}}
    rep = sales.build_sales(r, CFG, TODAY)
    from datetime import date
    assert rep["reps"] == ["Алдияр", "Сержан", "Данияр", "Фараби", "Остальные"] and len(rep["days"]) == 12
    assert rep["daily"][date(2026, 11, 10)]["Сержан"] == 1000 and rep["daily"][date(2026, 11, 10)]["Фараби"] == 500
    assert rep["daily"][date(2026, 11, 12)]["Сержан"] == -200 and rep["daily"][date(2026, 11, 12)]["Остальные"] == 700
    html = sales.html_section(rep)
    assert "по торговым представителям" in html and "Фараби" in html
    from openpyxl import load_workbook
    ws = load_workbook(sales.write_sales(rep, tmp_path / "s.xlsx"))["Сводка"]
    texts = [c.value for row in ws.iter_rows() for c in row if isinstance(c.value, str)]
    assert any("Продажи по дням" in x for x in texts) and "10.11 вт" in texts
