"""Ручной ввод логистики: шаблон input/logistics.xlsx и его чтение."""
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

# (ключ, заголовок, тип, ширина)
COLUMNS = [
    ("shipment", "Отправка №", "s", 14), ("order", "Заказ поставщику №", "s", 18), ("factory", "Завод", "s", 24),
    ("goods", "Товар", "s", 28), ("places", "Мест", "n", 8), ("weight", "Вес, кг", "n", 10),
    ("volume", "Объём, м3", "n", 10),
    ("cn_carrier", "Китай: перевозчик", "s", 18), ("cn_track", "Китай: трек", "s", 18),
    ("cn_cost", "Китай: стоимость", "n", 14), ("cn_cur", "Китай: валюта", "s", 10),
    ("cn_ship", "Китай: дата отправки с завода", "d", 16), ("cn_arrive", "Китай: прибыл на склад карго", "d", 16),
    ("ext_carrier", "Внешняя: карго", "s", 18), ("ext_route", "Внешняя: маршрут", "s", 20),
    ("ext_cost", "Внешняя: стоимость", "n", 14), ("ext_cur", "Внешняя: валюта", "s", 10),
    ("ext_paid", "Внешняя: оплачено", "n", 14), ("ext_ship", "Внешняя: дата отправки", "d", 16),
    ("eta", "План прибытия", "d", 14), ("fact", "Факт прибытия", "d", 14), ("comment", "Комментарий", "s", 30),
]
SHEET = "Логистика"


LOGISTICS_HELP = [
    "Одна строка = одна отправка (партия груза).",
    "Даты вводите как даты Excel (например 05.11.2026).",
    "Статус считается сам по заполненным датам: отправка с завода → склад карго → внешняя отправка → факт прибытия.",
    "Названия столбцов не меняйте, порядок менять можно.",
]


def create_template(path: Path, columns=None, sheet=None, help_lines=None) -> None:
    columns, sheet, help_lines = columns or COLUMNS, sheet or SHEET, help_lines or LOGISTICS_HELP
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = sheet
    for i, (_, title, _, width) in enumerate(columns, 1):
        c = ws.cell(row=1, column=i, value=title)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78" if not title.startswith(("Китай", "Внешняя")) else
                             "7F6000" if title.startswith("Китай") else "375623")
        c.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[c.column_letter].width = width
    ws.row_dimensions[1].height = 45
    ws.freeze_panes = "B2"
    info = wb.create_sheet("Инструкция")
    for line in help_lines:
        info.append([line])
    info.column_dimensions["A"].width = 110
    wb.save(path)


def _date(v):
    if isinstance(v, datetime):
        return v
    if hasattr(v, "year"):
        return datetime(v.year, v.month, v.day)
    if isinstance(v, str) and v.strip():
        for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y", "%d.%m.%y"):
            try:
                return datetime.strptime(v.strip(), fmt)
            except ValueError:
                pass
    return None


def _num(v):
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(" ", "").replace(",", "."))
    except (TypeError, ValueError):
        return None


def _read(path: Path, columns, sheet, help_lines) -> list[dict]:
    if not path.exists():
        create_template(path, columns, sheet, help_lines)
        return []
    ws = load_workbook(path, data_only=True)[sheet]
    by_title = {title: (key, kind) for key, title, kind, _ in columns}
    header = {i: by_title[str(c.value).strip()] for i, c in enumerate(ws[1]) if c.value and str(c.value).strip() in by_title}
    rows = []
    for raw in ws.iter_rows(min_row=2, values_only=True):
        if not any(v not in (None, "") for v in raw):
            continue
        row = {}
        for i, (key, kind) in header.items():
            v = raw[i] if i < len(raw) else None
            if kind == "s" and isinstance(v, float) and v.is_integer():
                v = int(v)  # код 50012, который Excel сохранил числом
            row[key] = _date(v) if kind == "d" else _num(v) if kind == "n" else ("" if v is None else str(v).strip())
        rows.append(row)
    return rows


def read_logistics(path: Path) -> list[dict]:
    return _read(path, COLUMNS, SHEET, LOGISTICS_HELP)


TRANSIT_TITLE = "В пути, шт (впишите)"
ETA_TITLE = "План прибытия (впишите)"


def read_transit(report: Path) -> list[dict]:
    """То, что вы вписали в жёлтые столбцы листа «Закуп» прошлого отчёта: код товара → в пути и план прибытия."""
    wb = load_workbook(report)
    if "Закуп" not in wb.sheetnames:
        return []
    ws = wb["Закуп"]
    col = {str(c.value).strip(): i for i, c in enumerate(ws[1]) if c.value}
    if not {"Код", TRANSIT_TITLE} <= set(col):
        return []
    out = []
    for raw in ws.iter_rows(min_row=2, values_only=True):
        code, qty = raw[col["Код"]], _num(raw[col[TRANSIT_TITLE]])
        if code not in (None, "") and qty and qty > 0:
            eta = raw[col[ETA_TITLE]] if ETA_TITLE in col else None
            out.append({"code": str(code).strip(), "qty": qty, "eta": _date(eta)})
    return out
