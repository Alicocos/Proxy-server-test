"""Лист «Продажи по менеджерам» по шаблону гендиректора: менеджер × бренд, лимит, долг, касса, продажи, возвраты, план."""
import re
from collections import defaultdict
from datetime import datetime

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

UUID = re.compile(r"[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}")
M = "#,##0"
COLS = ["Счет", "Бренд", "Менеджер", "Лимит", "Долг", "Остаток лимита", "Касса за 1 д.", "Касса за месяц",
        "Общая продажа за 1 день", "Продажа", "Возврат за 1 день", "Общий возврат", "План", "До плана", "План %"]
HEAD = PatternFill("solid", fgColor="1F4E78")
INPUT = PatternFill("solid", fgColor="FFF2CC")
NODATA = PatternFill("solid", fgColor="EDEDED")
TOTAL = PatternFill("solid", fgColor="DDEBF7")
thin = Side(style="thin", color="BFBFBF")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)


def _aid(d):
    m = UUID.search((d.get("agent") or {}).get("meta", {}).get("href", ""))
    return m.group(0) if m else None


def figures(raw: dict, today: datetime, who) -> dict:
    """Суммы по ключу who(id контрагента): продажи и возвраты за день и месяц, касса за день и месяц, долг."""
    day, month = today.strftime("%Y-%m-%d"), today.strftime("%Y-%m")
    f = defaultdict(lambda: defaultdict(float))
    docs = raw.get("sales_docs", {})
    for kind, key in (("demand", "sale"), ("retaildemand", "sale"), ("salesreturn", "ret"), ("retailsalesreturn", "ret")):
        for d in docs.get(kind, []):
            if d.get("applicable") is False or d["moment"][:7] != month:
                continue
            k = who(_aid(d))
            v = (d.get("sum") or 0) / 100
            f[k][key + "_m"] += v
            if d["moment"][:10] == day:
                f[k][key + "_d"] += v
    for d in raw.get("pay_in", []):
        if d.get("applicable") is False or d["moment"][:7] != month:
            continue
        k = who(_aid(d))
        v = (d.get("sum") or 0) / 100
        f[k]["cash_m"] += v
        if d["moment"][:10] == day:
            f[k]["cash_d"] += v
    if raw.get("balances") is not None:
        for c in raw.get("counterparties", []):
            debt = -(raw["balances"].get(c["id"]) or 0) / 100
            if debt > 0:
                f[who(c["id"])]["debt"] += debt
    return f


def build(raw: dict, cfg: dict, today: datetime, sc_raw: dict | None = None) -> dict:
    g = cfg.get("gd", {})
    reps = cfg["sales"].get("reps", [])
    cp = {c["id"]: c for c in raw.get("counterparties", [])}
    skip = [w.lower() for w in cfg["sales"].get("exclude_agents", [])]

    def who(aid):
        c = cp.get(aid, {})
        name, tags = c.get("name", "").lower(), {t.lower() for t in c.get("tags", [])}
        if any(w in name for w in skip):
            return None
        for r in reps:
            if r.get("agent") and r["agent"].lower() in name:
                return r["name"]
        for r in reps:
            if r.get("tag") and r["tag"].lower() in tags:
                return r["name"]
        return "клиенты" if "клиенты" in tags else None  # не клиенты (сотрудники, инвесторы и т.п.) в таблицу не идут

    base = figures(raw, today, who)
    sc = None
    if sc_raw:
        sc_tags = {t.lower() for t in g.get("sc_client_tags", ["клиент"])}
        sc_cp = {c["id"]: c for c in sc_raw.get("counterparties", [])}
        sc = figures(sc_raw, today, lambda a: "sc" if sc_tags & {t.lower() for t in sc_cp.get(a, {}).get("tags", [])} else None)["sc"]
    return {"today": today, "reps": reps, "base": base, "sc": sc, "cfg": g}


def write(wb, gd: dict, index: int = 1) -> None:
    today, g, base = gd["today"], gd["cfg"], gd["base"]
    ws = wb.create_sheet("Продажи по менеджерам", index)
    ws["A1"] = f"Продажи по менеджерам — {today:%d.%m.%Y %H:%M} (Алматы)"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = (f"Период: с 01.{today:%m.%Y} по {today:%d.%m.%Y}. «За 1 день» — {today:%d.%m.%Y}. "
                "Жёлтые ячейки — вводятся вручную (в МойСклад этих данных нет), серые — данных нет, см. пояснения внизу.")
    for i, (t, w) in enumerate(zip(COLS, [16, 12, 16, 14, 14, 14, 13, 14, 14, 14, 13, 13, 15, 15, 9]), 1):
        c = ws.cell(row=4, column=i, value=t)
        c.font, c.fill, c.border = Font(bold=True, color="FFFFFF"), HEAD, BOX
        c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[4].height = 36

    def put(r, col, v, fmt=M, fill=None, bold=False):
        c = ws.cell(row=r, column=col, value=v)
        c.border = BOX
        if fmt:
            c.number_format = fmt
        if fill:
            c.fill = fill
        if bold:
            c.font = Font(bold=True)
        return c

    limits = g.get("limits", {})
    r = 5
    block_start = r
    rows = [(rp.get("region", ""), rp["name"]) for rp in gd["reps"]]
    for region, name in rows:
        f = base.get(name, {})
        has = bool(f)
        for b, brand in enumerate(("ART HOME", "ARTDEC")):
            put(r, 2, brand, None)
            if b == 0:
                put(r, 1, region, None)
                put(r, 3, name.upper(), None)
                put(r, 4, limits.get(name), fill=INPUT)
                put(r, 5, f.get("debt", 0) if has else None, fill=None if has else NODATA)
                put(r, 6, f"=IF(D{r}=\"\",\"\",D{r}-E{r})")
                put(r, 7, f.get("cash_d") or None, fill=None if has else NODATA)
                put(r, 8, f.get("cash_m") or None, fill=None if has else NODATA)
                for col, k in ((9, "sale_d"), (10, "sale_m"), (11, "ret_d"), (12, "ret_m")):
                    put(r, col, f.get(k) or None, fill=None if has else NODATA)
                for col in (1, 3, 4, 5, 6, 7, 8):
                    ws.merge_cells(start_row=r, start_column=col, end_row=r + 1, end_column=col)
            else:
                for col in range(9, 13):
                    put(r, col, None, fill=NODATA)
            r += 1
    # клиенты без менеджера — чтобы итог сходился с МойСклад
    f = base.get("клиенты", {})
    put(r, 1, "БАЗА", None); put(r, 2, "ART HOME", None); put(r, 3, "БЕЗ МЕНЕДЖЕРА", None)
    put(r, 5, f.get("debt", 0)); put(r, 7, f.get("cash_d") or None); put(r, 8, f.get("cash_m") or None)
    for col, k in ((9, "sale_d"), (10, "sale_m"), (11, "ret_d"), (12, "ret_m")):
        put(r, col, f.get(k) or None)
    r += 1
    end = r - 1
    put(block_start, 13, g.get("plan_base"), fill=INPUT)
    ws.merge_cells(start_row=block_start, start_column=13, end_row=end, end_column=13)
    total_base = r
    put(r, 1, "Total BAZA", None, TOTAL, True)
    for col in range(2, 13):
        L = get_column_letter(col)
        put(r, col, f"=SUM({L}{block_start}:{L}{end})" if col >= 4 else None, fill=TOTAL, bold=True)
    put(r, 13, f"=M{block_start}", fill=TOTAL, bold=True)
    put(r, 14, f"=IF(M{r}=\"\",\"\",M{r}-J{r})", fill=TOTAL, bold=True)
    put(r, 15, f"=IF(N(M{r})=0,\"\",J{r}/M{r})", "0%", TOTAL, True)
    r += 2

    sc_start = r
    for name, brand in g.get("sc_managers", [["ЕРКАНАТ", "ART HOME"], ["НУРМУХАММЕТ", "ARTDEC"]]):
        put(r, 1, "СТРОЙ СИТИ", None); put(r, 2, brand, None); put(r, 3, name.upper(), None)
        for col in range(4, 13):
            put(r, col, None, fill=NODATA)
        r += 1
    sc = gd["sc"]
    put(r, 1, "СТРОЙ СИТИ", None); put(r, 2, "все бренды", None); put(r, 3, "ИТОГО ПО АККАУНТУ", None)
    if sc is not None:
        put(r, 5, sc.get("debt", 0)); put(r, 7, sc.get("cash_d") or None); put(r, 8, sc.get("cash_m") or None)
        for col, k in ((9, "sale_d"), (10, "sale_m"), (11, "ret_d"), (12, "ret_m")):
            put(r, col, sc.get(k) or None)
    else:
        for col in range(4, 13):
            put(r, col, None, fill=NODATA)
    sc_end = r
    put(sc_start, 13, g.get("plan_sc"), fill=INPUT)
    ws.merge_cells(start_row=sc_start, start_column=13, end_row=sc_end, end_column=13)
    r += 1
    total_sc = r
    put(r, 1, "Total STROY CITY", None, TOTAL, True)
    for col in range(2, 13):
        L = get_column_letter(col)
        put(r, col, f"=SUM({L}{sc_start}:{L}{sc_end})" if col >= 4 else None, fill=TOTAL, bold=True)
    put(r, 13, f"=M{sc_start}", fill=TOTAL, bold=True)
    put(r, 14, f"=IF(M{r}=\"\",\"\",M{r}-J{r})", fill=TOTAL, bold=True)
    put(r, 15, f"=IF(N(M{r})=0,\"\",J{r}/M{r})", "0%", TOTAL, True)
    r += 2
    put(r, 1, "Total General", None, TOTAL, True)
    for col in range(2, 16):
        L = get_column_letter(col)
        if col == 14:
            put(r, col, f"=M{r}-J{r}", fill=TOTAL, bold=True)
        elif col == 15:
            put(r, col, f"=IF(N(M{r})=0,\"\",J{r}/M{r})", "0%", TOTAL, True)
        else:
            put(r, col, f"={L}{total_base}+{L}{total_sc}" if col >= 4 else None, fill=TOTAL, bold=True)
    ws.freeze_panes = "D5"

    notes = [
        ("Что заполнено из МойСклад", "Долг (баланс взаиморасчётов клиентов менеджера), касса (входящие платежи клиентов менеджера), "
         "продажи (отгрузки) и возвраты за день и с начала месяца. Менеджер определяется по метке клиента "
         "(«алдияр клиенты», «данияр клиент», «сержан клиент»); Фараби — покупатель «Фараби роз.клиенты»."),
        ("Лимит, План — жёлтые", "В МойСклад их нет. Подставлены значения из таблицы гендиректора; меняются в config.toml, раздел [gd]. "
         "«Остаток лимита» и «До плана», «План %» считаются формулами."),
        ("ARTDEC у менеджеров базы — серые", "В первом аккаунте бренд ARTDEC ничем не отмечен: проект «Артдек базар» не ставится "
         "на отгрузки. Все продажи менеджера стоят в строке ART HOME. Чтобы разделить — ставить проект «Артдек базар» на отгрузки ARTDEC."),
        ("Касса по бренду", "Платежи в МойСклад без проекта, поэтому касса показана на менеджера целиком (объединённая ячейка)."),
        ("Нурлан Кали — серый", "Клиентов с его меткой нет (он есть только как сотрудник). Чтобы заполнялось — поставить его клиентам "
         "метку «нурлан клиент»."),
        ("Без менеджера", "Клиенты с меткой «клиенты», но без менеджера. Добавлено, чтобы итог по базе сходился с МойСклад."),
        ("Строй Сити", "Ерканат и Нурмухаммет в МойСклад Строи сити не отмечены (там метки «ернар клиент», «асхат клиент»), "
         "поэтому заполнен только итог по аккаунту, клиенты с меткой «клиент». Строй Сити сюда попадает, только когда в тот же "
         "запуск собран его отчёт."),
        ("За 1 день", f"Считается за {today:%d.%m.%Y}, то есть за день отчёта; в утреннем отчёте он почти пустой. "
         "Продажа — сумма отгрузок без вычета возвратов, возвраты — отдельно, как в шаблоне."),
    ]
    r += 3
    ws.cell(row=r, column=1, value="Пояснения").font = Font(bold=True, size=12)
    for i, (a, b) in enumerate(notes, r + 1):
        ws.cell(row=i, column=1, value=a).font = Font(bold=True)
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=15)
        c = ws.cell(row=i, column=2, value=b)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[i].height = 32
