"""Сборка Excel-отчёта."""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from . import analytics as a

HEAD_FILL = PatternFill("solid", fgColor="1F4E78")
HEAD_FONT = Font(bold=True, color="FFFFFF")
FILLS = {k: PatternFill("solid", fgColor=v) for k, v in {
    "red": "F8CBAD", "orange": "FFE699", "green": "C6E0B4", "grey": "D9D9D9", "blue": "BDD7EE"}.items()}
STATUS_FILL = {
    a.URGENT: "red", a.ORDER: "orange", a.OK: "green", a.OVER: "orange", a.DEAD: "red", a.NEW: "blue",
    a.TRANSIT: "blue", "Прибыл": "green", "В пути (внешняя)": "blue", "На складе карго": "orange",
    "В пути по Китаю": "orange", "Ждёт отгрузки с завода": "grey",
}
MONEY, QTY, DAYS, PCT, DATE = "#,##0", "#,##0.##", "0", "0%", "DD.MM.YYYY"


PROD_COLS = [
    ("Статус", "status", 17, None), ("ABC", "abc", 6, None), ("Товар", "name", 40, None),
    ("Артикул", "article", 14, None), ("Код", "code", 12, None), ("Группа", "folder", 24, None),
    ("Остаток", "stock", 10, QTY), ("В пути, шт (впишите)", "transit_in", 11, QTY), ("План прибытия (впишите)", "eta", 13, DATE), ("Резерв", "reserve", 9, QTY),
    ("Продажи 30 дн", "q30", 10, QTY), ("Продажи 90 дн", "q90", 10, QTY),
    ("в т.ч. на базар 90 дн", "baz90", 10, QTY), ("Доля базара", "baz_share", 8, PCT),
    ("Продаж в день", "daily", 9, "0.00"), ("Запас, дней", "cover", 9, DAYS),
    ("Запас с путём, дней", "cover_transit", 10, DAYS), ("Себест. ед.", "cost", 11, MONEY),
    ("Запас, сумма", "stock_value", 13, MONEY), ("Заказать, шт", "rec", 10, QTY),
    ("Заказать, сумма", "rec_value", 13, MONEY), ("Заморожено", "frozen", 13, MONEY),
    ("Выручка 90 дн", "rev90", 13, MONEY), ("Прибыль 90 дн", "profit90", 13, MONEY),
    ("Дней на складе", "stock_days", 9, DAYS),
]
COL = {key: get_column_letter(i) for i, (_, key, _, _) in enumerate(PROD_COLS, 1)}
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")


def _live_purchase(ws, rep):
    """Лист «Закуп»: жёлтые столбцы для ввода и формулы, которые пересчитывают заказ сразу в Excel."""
    k = rep["kpi"]
    c = COL
    for r, p in enumerate(rep["products"], 2):
        for key in ("transit_in", "eta"):
            ws[f"{c[key]}{r}"].fill = INPUT_FILL
        if not p["daily"]:
            continue
        have = f"({c['stock']}{r}+N({c['transit_in']}{r}))"
        ws[f"{c['cover_transit']}{r}"] = f"={have}/{c['daily']}{r}"
        ws[f"{c['rec']}{r}"] = f"=MAX(0,ROUNDUP({c['daily']}{r}*{k['target']}-{have},0))"
        ws[f"{c['rec_value']}{r}"] = f"={c['rec']}{r}*{c['cost']}{r}"
        ws[f"{c['status']}{r}"] = (
            f'=IF({have}/{c["daily"]}{r}<{k["lead"]},"{a.URGENT}",IF({c["rec"]}{r}>0,"{a.ORDER}",'
            f'IF({c["stock"]}{r}/{c["daily"]}{r}>{k["overstock"]},"{a.OVER}","{a.OK}")))')
        ws[f"{c['status']}{r}"].fill = PatternFill(fill_type=None)
    n = len(rep["products"]) + 1
    for text, color in ((a.URGENT, "red"), (a.ORDER, "orange"), (a.OVER, "orange"), (a.OK, "green")):
        ws.conditional_formatting.add(f"A2:A{n}", CellIsRule(operator="equal", formula=[f'"{text}"'], fill=FILLS[color]))
    ws.freeze_panes = "D2"


def table(ws, cols, rows, top=1, red_if_positive=()):
    """cols: (заголовок, ключ, ширина, формат)."""
    for i, (title, _, width, _) in enumerate(cols, 1):
        c = ws.cell(row=top, column=i, value=title)
        c.font, c.fill = HEAD_FONT, HEAD_FILL
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.row_dimensions[top].height = 42
    for r, row in enumerate(rows, top + 1):
        for i, (_, key, _, fmt) in enumerate(cols, 1):
            v = row.get(key)
            c = ws.cell(row=r, column=i, value=v)
            if fmt:
                c.number_format = fmt
            if key == "status" and v in STATUS_FILL:
                c.fill = FILLS[STATUS_FILL[v]]
            if key in red_if_positive and isinstance(v, (int, float)) and v > 0:
                c.fill = FILLS["red"]
    if top == 1:
        ws.freeze_panes = "A2"
        if rows:
            ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{len(rows) + 1}"


def _summary(ws, rep):
    k = rep["kpi"]
    n = max(k["n_products"] + 1, 2)
    ws.column_dimensions["A"].width = 46
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 60
    ws["A1"] = "Сводка по закупу и логистике"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = (f"Дата отчёта: {rep['today']:%d.%m.%Y %H:%M}. Суммы — в валюте учёта МойСклад, товар — по себестоимости. "
                f"Продажи: {k['sales_source']}.")
    blocks = [
        ("Товарный запас", [
            ("Запас на складах", k["stock_value"], MONEY, ""),
            ("Позиций в наличии", k["sku_in_stock"], MONEY, ""),
            ("Оборачиваемость, дней", k["turnover_days"], DAYS, "за сколько дней продаётся текущий запас"),
            ("Заморожено в товаре", k["frozen"], MONEY, "неликвид + излишек сверх нормы"),
            ("Доля заморозки", k["frozen_share"], PCT, ""),
            ("  неликвид (нет продаж 90 дней)", k["dead_value"], MONEY, f"{k['dead_sku']} позиций"),
            ("  излишек сверх нормы", k["over_value"], MONEY, f"{k['over_sku']} позиций"),
            ("На исключённых складах (брак)", k["excluded_value"], MONEY, "в анализ закупа не входит"),
            ("Можно перераспределить между складами", k["transfers_value"], MONEY, f"{k['transfers']} перемещений вместо закупа"),
        ]),
        ("Закуп", [
            ("Позиций «срочно заказать»", f"=COUNTIF('Закуп'!{COL['status']}:{COL['status']},\"{a.URGENT}\")", MONEY, "запас вместе с товаром в пути кончится раньше срока поставки"),
            ("Рекомендовано к заказу", f"=SUM('Закуп'!{COL['rec_value']}:{COL['rec_value']})", MONEY, "по себестоимости"),
            ("Товар в пути", f"=SUMPRODUCT('Закуп'!{COL['transit_in']}2:{COL['transit_in']}{n},'Закуп'!{COL['cost']}2:{COL['cost']}{n})", MONEY, "то, что вы вписали в жёлтый столбец листа «Закуп», по себестоимости"),
            ("Позиций в пути", k["transit_sku"], MONEY, f"с прошедшей плановой датой: {k['eta_overdue']}"),
        ]),
        ("Деньги у заводов", [
            ("Депозиты у заводов (предоплата без товара)", k["deposits"], MONEY, "по балансу взаиморасчётов: оплачено больше, чем принято"),
            ("  из них без движения", k["stale"], MONEY, f"{k['stale_n']} заводов: предоплата есть, приёмок давно нет"),
            ("Наш долг заводам", k["debts"], MONEY, "принято больше, чем оплачено"),
            ("Оплачено заводам за 30 дней", k["pay30"], MONEY, ""),
            ("Оплачено заводам за 90 дней", k["pay90"], MONEY, ""),
            ("Принято от заводов за 90 дней", k["sup90"], MONEY, ""),
        ]),
        ("Базар (ангары)", [
            ("Отгружено на базар за 90 дней", k["baz_sum"], MONEY, "по ценам отгрузки"),
            ("Доля базара в отгрузках", k["baz_share"], PCT, ""),
            ("Получено денег от базара за 90 дней", k["baz_in"], MONEY, ""),
            ("Баланс взаиморасчётов с базаром", k["baz_balance"], MONEY, "плюс — денег получено больше, чем отгружено"),
            ("Позиций отгружалось на базар", k["baz_sku"], MONEY, "остатки и продажи самих ангаров в МойСклад не видны"),
        ]),
        ("Логистика", [
            ("Расходы на логистику за 3 последних месяца", k["logi3"], MONEY, "статьи расходов МойСклад из config.toml"),
            ("Логистика в % от принятого товара", k["logi3_share"], "0.0%", ""),
            ("Отправок в работе", k["ship_active"], MONEY, "из input/logistics.xlsx"),
            ("Из них просрочено", k["ship_overdue"], MONEY, ""),
            ("Не оплачено за внешнюю логистику", k["ship_unpaid"], None, ""),
        ]),
    ]
    r = 4
    for title, lines in blocks:
        c = ws.cell(row=r, column=1, value=title)
        c.font, c.fill = HEAD_FONT, HEAD_FILL
        ws.cell(row=r, column=2).fill = ws.cell(row=r, column=3).fill = HEAD_FILL
        r += 1
        for name, val, fmt, note in lines:
            ws.cell(row=r, column=1, value=name)
            v = ws.cell(row=r, column=2, value=val if val is not None else "—")
            v.alignment = Alignment(horizontal="right")
            if fmt:
                v.number_format = fmt
            ws.cell(row=r, column=3, value=note).font = Font(italic=True, color="7F7F7F")
            r += 1
        r += 1
    return r


def write_report(rep: dict, path: Path) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Сводка"
    _summary(ws, rep)

    table(wb.create_sheet("Закуп"), PROD_COLS, rep["products"])
    _live_purchase(wb["Закуп"], rep)
    frozen_cols = [("В пути, шт", "transit", 10, QTY) if c[1] == "transit_in" else c for c in PROD_COLS if c[1] != "eta"]
    frozen = sorted((p for p in rep["products"] if p["frozen"] > 0), key=lambda p: -p["frozen"])
    table(wb.create_sheet("Заморозка"), frozen_cols, frozen)

    table(wb.create_sheet("Склады"), [
        ("Склад", "store", 30, None), ("Запас, сумма", "stock_value", 15, MONEY), ("Позиций", "sku", 10, MONEY),
        ("Себест. продаж 90 дн", "cogs90", 16, MONEY), ("Оборачиваемость, дней", "turnover", 14, DAYS),
        ("Товар без продаж 90 дн на этом складе", "no_sales_value", 18, MONEY),
        ("Доля без продаж", "no_sales_share", 11, PCT),
    ], rep["stores"])

    table(wb.create_sheet("Перемещения"), [
        ("Товар", "name", 40, None), ("Артикул", "article", 14, None),
        ("КУДА (дефицит)", "to", 22, None), ("Остаток", "to_stock", 9, QTY), ("Продажи 90 дн", "to_q90", 10, QTY),
        ("Запас, дней", "to_cover", 9, DAYS),
        ("ОТКУДА (излишек)", "from", 22, None), ("Остаток", "from_stock", 9, QTY),
        ("Продажи 90 дн", "from_q90", 10, QTY), ("Запас, дней", "from_cover", 9, DAYS),
        ("Переместить, шт", "qty", 11, QTY), ("Сумма", "value", 13, MONEY),
    ], rep["transfers"])

    table(wb.create_sheet("Базар"), [
        ("Товар", "name", 40, None), ("Артикул", "article", 14, None), ("Код", "code", 12, None),
        ("На базар 30 дн, шт", "baz30", 11, QTY), ("На базар 90 дн, шт", "baz90", 11, QTY),
        ("На базар 90 дн, сумма", "baz_sum90", 14, MONEY), ("Доля базара в продажах товара", "baz_share", 12, PCT),
        ("Все продажи 90 дн, шт", "q90", 11, QTY), ("Остаток на наших складах", "stock", 11, QTY),
        ("Статус закупа", "status", 17, None),
    ], rep["bazaar"])

    table(wb.create_sheet("Заводы"), [
        ("Завод", "agent", 34, None), ("Депозит у завода (предоплата)", "deposit", 16, MONEY),
        ("Наш долг заводу", "debt", 15, MONEY), ("Внимание", "flag", 24, None),
        ("Оплачено 30 дн", "pay30", 14, MONEY), ("Оплачено 90 дн", "pay90", 14, MONEY),
        ("Оплачено за год", "pay365", 15, MONEY), ("Принято 90 дн", "sup90", 14, MONEY),
        ("Принято за год", "sup365", 15, MONEY), ("Последний платёж", "last_pay", 12, DATE),
        ("Последняя приёмка", "last_sup", 12, DATE), ("Дней без приёмок", "days_no_sup", 9, DAYS),
        ("Платежей за год", "payments", 9, MONEY), ("Приёмок за год", "supplies", 9, MONEY),
    ], rep["suppliers"])
    for row in wb["Заводы"].iter_rows(min_row=2, min_col=4, max_col=4):
        if row[0].value:
            row[0].fill = FILLS["red"]

    table(wb.create_sheet("Платежи"), [
        ("Дата", "moment", 12, DATE), ("Документ №", "name", 12, None), ("Завод", "agent", 34, None),
        ("Тип", "kind", 10, None), ("Сумма", "sum", 14, MONEY), ("Статья расходов", "item", 20, None),
        ("Назначение / комментарий", "purpose", 50, None),
    ], rep["payments"])

    table(wb.create_sheet("Помесячно"), [
        ("Месяц", "month", 10, None), ("Оплачено заводам", "paid", 16, MONEY),
        ("Принято от заводов", "received", 16, MONEY), ("Оплачено минус принято", "diff", 16, MONEY),
        ("Расходы на логистику", "logistics", 16, MONEY), ("Логистика, % от принятого", "logi_share", 13, "0.0%"),
    ], rep["monthly"])

    table(wb.create_sheet("Логистика"), [
        ("Статус", "status", 22, None), ("Отправка №", "shipment", 12, None), ("Заказ №", "order", 12, None),
        ("Завод", "factory", 24, None), ("Товар", "goods", 26, None), ("Мест", "places", 7, QTY),
        ("Вес, кг", "weight", 9, QTY), ("Объём, м3", "volume", 9, QTY),
        ("Китай: перевозчик", "cn_carrier", 16, None), ("Китай: стоимость", "cn_cost", 12, MONEY),
        ("Вал.", "cn_cur", 6, None), ("Отправка с завода", "cn_ship", 12, DATE),
        ("На складе карго", "cn_arrive", 12, DATE), ("Карго", "ext_carrier", 16, None),
        ("Маршрут", "ext_route", 18, None), ("Внешняя: стоимость", "ext_cost", 12, MONEY),
        ("Вал.", "ext_cur", 6, None), ("Не оплачено", "ext_unpaid", 12, MONEY),
        ("Ставка за м3", "per_m3", 10, MONEY), ("Ставка за кг", "per_kg", 9, "0.00"),
        ("Внешняя отправка", "ext_ship", 12, DATE), ("План прибытия", "eta", 12, DATE),
        ("Факт прибытия", "fact", 12, DATE), ("Дней в пути", "days", 8, DAYS),
        ("Просрочка, дней", "overdue", 10, DAYS), ("Комментарий", "comment", 30, None),
    ], rep["logistics"], red_if_positive=("overdue", "ext_unpaid"))

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path
