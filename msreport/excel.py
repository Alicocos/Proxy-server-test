"""Сборка Excel-отчёта."""
from pathlib import Path

from openpyxl import Workbook
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
    k, ql = rep["kpi"], rep["quality"]
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
            ("Позиций «срочно заказать»", k["urgent_sku"], MONEY, "запас с учётом пути кончится раньше срока поставки"),
            ("Рекомендовано к заказу", k["rec_value"], MONEY, f"{k['rec_sku']} позиций, по себестоимости"),
            ("Товар в пути", k["transit_value"], MONEY, "позиции открытых заказов заводам, ещё не принятые"),
        ]),
        ("Деньги у заводов", [
            ("Депозиты у заводов (предоплата без товара)", k["deposits"], MONEY, "по балансу взаиморасчётов: оплачено больше, чем принято"),
            ("Наш долг заводам", k["debts"], MONEY, "принято больше, чем оплачено"),
            ("Оплачено заводам за 30 дней", k["payments_30"], MONEY, ""),
            ("Открытых заказов", k["orders_open"], MONEY, ""),
            ("Сумма открытых заказов", k["orders_sum"], MONEY, ""),
            ("Осталось доплатить по открытым заказам", k["orders_to_pay"], MONEY, "верно, только если платежи привязаны к заказам"),
            ("Заказов с просроченной поставкой", k["orders_overdue"], MONEY, "считается только там, где указана плановая дата"),
        ]),
        ("Качество учёта в МойСклад", [
            ("Открытые заказы без привязанных оплат", f"{ql['orders_no_pay']} из {ql['orders']}", None, "по ним не видно, сколько уже оплачено"),
            ("Открытые заказы без плановой даты", f"{ql['orders_no_plan']} из {ql['orders']}", None, "по ним нельзя посчитать просрочку"),
            ("Приёмки без привязки к заказу", f"{ql['supplies_unlinked']} из {ql['supplies']}", None, "заказ не закрывается, товар в пути не считается"),
            ("Платежи заводам без привязки к документу", f"{ql['payments_unlinked']} из {ql['payments']}", None, "депозит виден только суммой по заводу"),
        ]),
        ("Логистика", [
            ("Отправок в работе", k["ship_active"], MONEY, ""),
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

    prod_cols = [
        ("Статус", "status", 17, None), ("ABC", "abc", 6, None), ("Товар", "name", 40, None),
        ("Артикул", "article", 14, None), ("Код", "code", 12, None), ("Группа", "folder", 24, None),
        ("Остаток", "stock", 10, QTY), ("В пути", "transit", 10, QTY), ("Резерв", "reserve", 9, QTY),
        ("Продажи 30 дн", "q30", 10, QTY), ("Продажи 90 дн", "q90", 10, QTY),
        ("Продаж в день", "daily", 9, "0.00"), ("Запас, дней", "cover", 9, DAYS),
        ("Запас с путём, дней", "cover_transit", 10, DAYS), ("Себест. ед.", "cost", 11, MONEY),
        ("Запас, сумма", "stock_value", 13, MONEY), ("Заказать, шт", "rec", 10, QTY),
        ("Заказать, сумма", "rec_value", 13, MONEY), ("Заморожено", "frozen", 13, MONEY),
        ("Выручка 90 дн", "rev90", 13, MONEY), ("Прибыль 90 дн", "profit90", 13, MONEY),
        ("Дней на складе", "stock_days", 9, DAYS),
    ]
    table(wb.create_sheet("Закуп"), prod_cols, rep["products"])
    frozen = sorted((p for p in rep["products"] if p["frozen"] > 0), key=lambda p: -p["frozen"])
    table(wb.create_sheet("Заморозка"), prod_cols, frozen)

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

    table(wb.create_sheet("Заводы"), [
        ("Завод", "agent", 34, None), ("Депозит у завода (предоплата)", "deposit", 17, MONEY),
        ("Наш долг заводу", "debt", 15, MONEY), ("Платежи за период", "payments_base", 15, MONEY),
        ("Платежей, шт", "payments", 9, MONEY), ("Приёмки за год", "supplied_base", 15, MONEY),
        ("Приёмок, шт", "supplies", 9, MONEY), ("Открытых заказов", "orders", 10, MONEY),
        ("Сумма открытых заказов", "sum_base", 15, MONEY), ("Просроченных заказов", "overdue_orders", 12, MONEY),
    ], rep["suppliers"], red_if_positive=("overdue_orders",))

    table(wb.create_sheet("Заказы заводам"), [
        ("Заказ №", "name", 12, None), ("Дата", "moment", 12, DATE), ("Завод", "agent", 30, None),
        ("Статус", "state", 18, None), ("Валюта", "currency", 8, None), ("Сумма", "sum", 14, MONEY),
        ("Оплачено (привязано)", "paid", 14, MONEY), ("% оплаты", "paid_share", 9, PCT),
        ("Принято (привязано)", "received", 14, MONEY), ("Доплатить", "to_pay", 14, MONEY),
        ("План поставки", "planned", 12, DATE), ("Заказу дней", "age", 9, DAYS),
        ("Просрочка, дней", "overdue", 10, DAYS), ("Сумма, вал. учёта", "sum_base", 14, MONEY),
        ("Чего не хватает в учёте", "flags", 38, None), ("Комментарий", "description", 30, None),
    ], rep["orders"], red_if_positive=("overdue",))

    table(wb.create_sheet("Платежи"), [
        ("Дата", "moment", 12, DATE), ("Документ №", "name", 12, None), ("Завод", "agent", 28, None),
        ("Тип", "kind", 10, None), ("Валюта", "currency", 8, None), ("Сумма", "sum", 13, MONEY),
        ("Сумма, вал. учёта", "sum_base", 14, MONEY), ("Привязан к", "linked", 18, None),
        ("Назначение", "purpose", 40, None),
    ], rep["payments"])

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
