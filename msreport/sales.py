"""Продажи по группам клиентов: помесячно за год и за короткие периоды (вчера, 7 дней, месяц)."""
import re
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

UUID = re.compile(r"[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}")
RU = ["янв", "фев", "мар", "апр", "май", "июн", "июл", "авг", "сен", "окт", "ноя", "дек"]
OTHER = "Прочие"


def first_month(today: datetime) -> datetime:
    """Начало окна: 12 полных месяцев до текущего."""
    return datetime(today.year - 1, today.month, 1)


def _months(today: datetime) -> list[str]:
    out, y, m = [], today.year - 1, today.month
    for _ in range(13):
        out.append(f"{y}-{m:02d}")
        y, m = (y + 1, 1) if m == 12 else (y, m + 1)
    return out


def build_sales(raw: dict, cfg: dict, today: datetime) -> dict:
    groups = [(g["name"], g["tag"].lower()) for g in cfg["sales"]["groups"]]
    names = [g for g, _ in groups] + [OTHER]
    cp = {c["id"]: c for c in raw.get("counterparties", [])}

    def group(aid):
        tags = {t.lower() for t in cp.get(aid, {}).get("tags", [])}
        return next((g for g, tag in groups if tag in tags), OTHER)

    months = _months(today)
    full, cur = months[:12], months[-1]
    day0 = today.replace(hour=0, minute=0, second=0, microsecond=0)
    prev_m = months[-2]
    periods = {  # ключ: (с, по) включительно по дате
        "today": (day0, day0), "yesterday": (day0 - timedelta(days=1), day0 - timedelta(days=1)),
        "d7": (day0 - timedelta(days=6), day0), "mtd": (day0.replace(day=1), day0),
        "prev_mtd": (datetime(int(prev_m[:4]), int(prev_m[5:]), 1),
                     datetime(int(prev_m[:4]), int(prev_m[5:]), 1) + timedelta(days=day0.day - 1)),
    }
    z = lambda: defaultdict(float)  # noqa: E731
    cl = defaultdict(lambda: {"ship": z(), "ret": z(), "n": z(), "last": None})
    per = {g: {k: {"net": 0.0, "n": 0, "clients": set()} for k in periods} for g in names}
    docs = raw.get("sales_docs", {})
    for kind, sign in (("demand", 1), ("retaildemand", 1), ("salesreturn", -1), ("retailsalesreturn", -1)):
        for d in docs.get(kind, []):
            if d.get("applicable") is False:
                continue
            m = UUID.search((d.get("agent") or {}).get("meta", {}).get("href", ""))
            when = datetime.strptime(d["moment"][:19], "%Y-%m-%d %H:%M:%S")
            ym = d["moment"][:7]
            if not m or ym not in months:
                continue
            aid, v = m.group(0), (d.get("sum") or 0) / 100
            c = cl[aid]
            if sign > 0:
                c["ship"][ym] += v
                c["n"][ym] += 1
                c["last"] = max(c["last"] or when, when)
            else:
                c["ret"][ym] += v
            g = group(aid)
            for k, (a, b) in periods.items():
                if a <= when < b + timedelta(days=1):
                    per[g][k]["net"] += sign * v
                    if sign > 0:
                        per[g][k]["n"] += 1
                        per[g][k]["clients"].add(aid)
    for g in names:
        for k in periods:
            per[g][k]["clients"] = len(per[g][k]["clients"])

    clients = []
    for aid, c in cl.items():
        net = {m: c["ship"][m] - c["ret"][m] for m in months}
        clients.append({"id": aid, "name": cp.get(aid, {}).get("name", aid), "group": group(aid), "net": net,
                        "ship": dict(c["ship"]), "ret": dict(c["ret"]), "n": dict(c["n"]), "last": c["last"],
                        "total": sum(net[m] for m in full), "docs": sum(c["n"][m] for m in full),
                        "active_months": sum(1 for m in full if c["ship"][m] > 0)})
    clients.sort(key=lambda c: -c["total"])
    G = {}
    for g in names:
        mem = [c for c in clients if c["group"] == g]
        G[g] = {
            "net": {m: sum(c["net"][m] for c in mem) for m in months},
            "ship": {m: sum(c["ship"].get(m, 0) for c in mem) for m in months},
            "ret": {m: sum(c["ret"].get(m, 0) for c in mem) for m in months},
            "n": {m: sum(c["n"].get(m, 0) for c in mem) for m in months},
            "active": {m: sum(1 for c in mem if c["ship"].get(m, 0) > 0) for m in months},
            "clients": len(mem), "total": sum(c["total"] for c in mem), "periods": per[g],
            "silent": [c["name"] for c in mem[:10] if c["last"] and (today - c["last"]).days > 45],
        }
    shown = [g for g in names if g != OTHER or G[g]["clients"]]
    return {"today": today, "months": months, "full": full, "cur": cur, "groups": shown, "G": G, "clients": clients,
            "total": sum(G[g]["total"] for g in names), "period_keys": list(periods)}


BUCKETS = ["0–30 дней", "31–60 дней", "61–90 дней", "старше 90 дней"]


def build_debts(raw: dict, cfg: dict, today: datetime) -> dict | None:
    """Дебиторка: клиенты с долгом по балансу взаиморасчётов МойСклад, возраст долга и последние платежи.

    Возраст считается так: долг относим на самые свежие отгрузки (старые считаем оплаченными первыми).
    """
    balances = raw.get("balances")
    if balances is None:
        return None
    groups = [(g["name"], g["tag"].lower()) for g in cfg["sales"]["groups"]]
    sup_tags = {t.lower() for t in cfg["suppliers"]["tags"]}
    aid = lambda d: (UUID.search((d.get("agent") or {}).get("meta", {}).get("href", "")) or [None])[0]  # noqa: E731
    when = lambda d: datetime.strptime(d["moment"][:19], "%Y-%m-%d %H:%M:%S")  # noqa: E731
    ships, pays = defaultdict(list), defaultdict(list)
    for k in ("demand", "retaildemand"):
        for d in raw.get("sales_docs", {}).get(k, []):
            if d.get("applicable") is not False and aid(d):
                ships[aid(d)].append((when(d), (d.get("sum") or 0) / 100))
    for d in raw.get("pay_in", []):
        if d.get("applicable") is not False and aid(d):
            pays[aid(d)].append((when(d), (d.get("sum") or 0) / 100))
    rows = []
    for c in raw.get("counterparties", []):
        debt = -(balances.get(c["id"]) or 0) / 100
        tags = {t.lower() for t in c.get("tags", [])}
        if debt < 1 or tags & sup_tags:
            continue
        sh, pa = sorted(ships[c["id"]], reverse=True), sorted(pays[c["id"]], reverse=True)
        left, buckets, since = debt, [0.0] * 4, None
        for d, v in sh:
            if left <= 0:
                break
            part = min(left, v)
            age = (today - d).days
            buckets[0 if age <= 30 else 1 if age <= 60 else 2 if age <= 90 else 3] += part
            left -= part
            since = d
        if left > 0.5:  # долг больше всех отгрузок за год — остаток старше окна выгрузки
            buckets[3] += left
            since = None
        d90 = lambda lst: sum(v for d, v in lst if (today - d).days <= 90)  # noqa: E731
        rows.append({
            "name": c["name"], "group": next((g for g, tag in groups if tag in tags), OTHER), "debt": debt,
            "b0": buckets[0], "b1": buckets[1], "b2": buckets[2], "b3": buckets[3],
            "since": since, "age": (today - since).days if since else None,
            "age_text": "" if since else "старше года",
            "last_pay": pa[0][0] if pa else None, "last_pay_sum": pa[0][1] if pa else None,
            "no_pay_days": (today - pa[0][0]).days if pa else None,
            "last_ship": sh[0][0] if sh else None, "no_ship_days": (today - sh[0][0]).days if sh else None,
            "ship90": d90(sh), "pay90": d90(pa),
            "flag": ("нет платежей за год" if not pa else "не платит больше 60 дней" if (today - pa[0][0]).days > 60
                     else "не платит больше 30 дней" if (today - pa[0][0]).days > 30 else ""),
        })
    rows.sort(key=lambda r: -r["debt"])
    names = [g for g, _ in groups] + [OTHER]
    by_group = [{"group": g, "n": sum(1 for r in rows if r["group"] == g),
                 **{k: sum(r[k] for r in rows if r["group"] == g) for k in ("debt", "b0", "b1", "b2", "b3")}} for g in names]
    by_group = [g for g in by_group if g["n"]]
    total = sum(r["debt"] for r in rows)
    return {"today": today, "rows": rows, "by_group": by_group, "total": total,
            "old": sum(r["b2"] + r["b3"] for r in rows), "n": len(rows),
            "top5_share": sum(r["debt"] for r in rows[:5]) / total if total else 0}


def write_debts(wb, debts: dict) -> None:
    D = "DD.MM.YYYY"
    cols = [("Клиент", "name", 38, None), ("Группа", "group", 20, None), ("Долг", "debt", 14, M),
            ("Внимание", "flag", 24, None), ("Возраст долга, дней", "age", 10, "0"), ("", "age_text", 11, None),
            ("Долг тянется с", "since", 12, D), (BUCKETS[0], "b0", 13, M), (BUCKETS[1], "b1", 13, M),
            (BUCKETS[2], "b2", 13, M), (BUCKETS[3], "b3", 13, M), ("Последний платёж", "last_pay", 12, D),
            ("Сумма последнего платежа", "last_pay_sum", 13, M), ("Дней без платежей", "no_pay_days", 10, "0"),
            ("Последняя отгрузка", "last_ship", 12, D), ("Дней без отгрузок", "no_ship_days", 10, "0"),
            ("Отгружено за 90 дн", "ship90", 14, M), ("Оплачено за 90 дн", "pay90", 14, M)]
    ws = wb.create_sheet("Дебиторка", 1)
    _head(ws, [c[0] for c in cols], [c[2] for c in cols])
    red = PatternFill("solid", fgColor="F8CBAD")
    for r, row in enumerate(debts["rows"], 2):
        for i, (_, k, _, f) in enumerate(cols, 1):
            c = ws.cell(row=r, column=i, value=row[k] or None if k.startswith("b") else row[k])
            if f:
                c.number_format = f
            if k == "flag" and row[k]:
                c.fill = red
    n = len(debts["rows"]) + 2
    ws.cell(row=n, column=1, value="ИТОГО").font = BOLD
    for i, (_, k, _, f) in enumerate(cols, 1):
        if k in ("debt", "b0", "b1", "b2", "b3", "ship90", "pay90"):
            col = get_column_letter(i)
            c = ws.cell(row=n, column=i, value=f"=SUM({col}2:{col}{n - 1})")
            c.font, c.number_format = BOLD, M
    ws.freeze_panes = "B2"
    if debts["rows"]:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{len(debts['rows']) + 1}"
    g = wb.create_sheet("Дебиторка по группам", 2)
    _head(g, ["Группа", "Должников", "Долг"] + BUCKETS + ["Доля старше 60 дней"], [28, 11, 15, 14, 14, 14, 14, 12])
    for r, row in enumerate(debts["by_group"], 2):
        vals = [row["group"], row["n"], row["debt"], row["b0"], row["b1"], row["b2"], row["b3"],
                (row["b2"] + row["b3"]) / row["debt"] if row["debt"] else None]
        for i, v in enumerate(vals, 1):
            g.cell(row=r, column=i, value=v).number_format = "0%" if i == 8 else M
    n = len(debts["by_group"]) + 2
    g.cell(row=n, column=1, value="ИТОГО").font = BOLD
    for i in range(2, 8):
        col = get_column_letter(i)
        c = g.cell(row=n, column=i, value=f"=SUM({col}2:{col}{n - 1})")
        c.font, c.number_format = BOLD, M


def debts_html(debts: dict) -> str:
    n = lambda v: "—" if v is None else f"{v:,.0f}".replace(",", " ")  # noqa: E731
    d = lambda v: f"{v:%d.%m.%Y}" if v else "—"  # noqa: E731
    th = '<tr style="background:#1F4E78;color:#fff">'
    tb = '<table border="1" cellpadding="4" style="border-collapse:collapse;font-size:13px;border-color:#ddd">'
    grp = "".join(f'<tr><td>{g["group"]}</td><td align="right">{g["n"]}</td><td align="right">{n(g["debt"])}</td>'
                  f'<td align="right">{n(g["b2"] + g["b3"])}</td></tr>' for g in debts["by_group"])
    top = "".join(f'<tr><td>{r["name"]}</td><td>{r["group"]}</td><td align="right">{n(r["debt"])}</td>'
                  f'<td align="right">{n(r["age"]) if r["age"] is not None else r["age_text"]}</td>'
                  f'<td align="right">{d(r["last_pay"])}</td><td align="right">{n(r["no_pay_days"])}</td></tr>'
                  for r in debts["rows"][:10])
    return (f'<h3>Дебиторка: клиенты должны {n(debts["total"])}</h3>'
            f'<p>Должников: {debts["n"]}. Из долга старше 60 дней: {n(debts["old"])}. '
            f'На пять крупнейших должников приходится {debts["top5_share"]:.0%} долга.</p>'
            f'{tb}{th}<th align="left">Группа</th><th>Должников</th><th>Долг</th><th>Из него старше 60 дней</th></tr>{grp}</table>'
            f'<p><b>Топ-10 должников</b></p>'
            f'{tb}{th}<th align="left">Клиент</th><th>Группа</th><th>Долг</th><th>Возраст долга, дней</th>'
            f'<th>Последний платёж</th><th>Дней без платежей</th></tr>{top}</table>')


def label(m: str, rep: dict) -> str:
    return f"{RU[int(m[5:]) - 1]} {m[:4]}" + (f" (по {rep['today']:%d.%m})" if m == rep["cur"] else "")


HF, HB, BOLD = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="1F4E78"), Font(bold=True)
M, P = "#,##0", "0.0%"


def _head(ws, titles, widths, row=1):
    for i, (t, w) in enumerate(zip(titles, widths), 1):
        c = ws.cell(row=row, column=i, value=t)
        c.font, c.fill = HF, HB
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[row].height = 36


def _matrix(ws, rep, first, rows, top, fmt=M, share=True, width0=28):
    months, full = rep["months"], rep["full"]
    _head(ws, [first] + [label(m, rep) for m in months] + ["Итого 12 мес"] + (["Доля"] if share else []),
          [width0] + [12] * len(months) + [15, 8], top)
    grand = sum(sum(v.get(m, 0) for m in full) for _, v in rows) or 1
    r = top
    for r, (name, v) in enumerate(rows, top + 1):
        ws.cell(row=r, column=1, value=name)
        for j, m in enumerate(months, 2):
            ws.cell(row=r, column=j, value=round(v.get(m, 0), 2) or None).number_format = fmt
        t = sum(v.get(m, 0) for m in full)
        ws.cell(row=r, column=len(months) + 2, value=t).number_format = fmt
        if share:
            ws.cell(row=r, column=len(months) + 3, value=t / grand).number_format = P
    r += 1
    ws.cell(row=r, column=1, value="ИТОГО").font = BOLD
    for j in range(2, len(months) + 3):
        col = get_column_letter(j)
        c = ws.cell(row=r, column=j, value=f"=SUM({col}{top + 1}:{col}{r - 1})")
        c.font, c.number_format = BOLD, fmt
    ws.freeze_panes = ws.cell(row=top + 1, column=2)
    return r


def write_sales(rep: dict, path: Path, debts: dict | None = None) -> Path:
    G, groups, today = rep["G"], rep["groups"], rep["today"]
    wb = Workbook()
    ws = wb.active
    ws.title = "Сводка"
    ws["A1"] = f"Продажи по группам клиентов на {today:%d.%m.%Y %H:%M} (Алматы)"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = ("Чистые продажи = отгрузки минус возвраты покупателей, по ценам отгрузки, тенге. "
                "«Итого 12 мес» — 12 полных месяцев, текущий месяц в итог не входит.")
    heads = ["Группа", "Сегодня", "Вчера", "7 дней", "Месяц с начала", "Прошлый месяц за те же дни", "Изменение",
             "Отгрузок за месяц", "Покупало клиентов за месяц", "Средний чек за месяц"]
    _head(ws, heads, [28, 13, 13, 14, 15, 16, 11, 11, 13, 14], 4)
    for r, g in enumerate(groups, 5):
        p = G[g]["periods"]
        vals = [g, p["today"]["net"], p["yesterday"]["net"], p["d7"]["net"], p["mtd"]["net"], p["prev_mtd"]["net"],
                (p["mtd"]["net"] / p["prev_mtd"]["net"] - 1) if p["prev_mtd"]["net"] > 0 else None,
                p["mtd"]["n"], p["mtd"]["clients"], p["mtd"]["net"] / p["mtd"]["n"] if p["mtd"]["n"] else None]
        for i, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=i, value=v)
            c.number_format = "+0%;-0%" if i == 7 else M
    r = 5 + len(groups)
    ws.cell(row=r, column=1, value="ИТОГО").font = BOLD
    for i in (2, 3, 4, 5, 6, 8):
        col = get_column_letter(i)
        c = ws.cell(row=r, column=i, value=f"=SUM({col}5:{col}{r - 1})")
        c.font, c.number_format = BOLD, M
    top = r + 3
    ws.cell(row=top - 1, column=1, value="Чистые продажи по месяцам").font = BOLD
    end = _matrix(ws, rep, "Группа", [(g, G[g]["net"]) for g in groups], top)
    ws.freeze_panes = None
    ch = LineChart()
    ch.title, ch.height, ch.width = "Чистые продажи по группам", 9, 26
    ch.y_axis.numFmt = '#,##0,,"млн"'
    ch.add_data(Reference(ws, min_col=1, max_col=13, min_row=top + 1, max_row=top + len(groups)),
                from_rows=True, titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=2, max_col=13, min_row=top))
    ws.add_chart(ch, f"A{end + 3}")

    for title, key, share in (("Отгрузки", "ship", True), ("Возвраты", "ret", True), ("Число отгрузок", "n", True),
                              ("Активные клиенты", "active", False)):
        _matrix(wb.create_sheet(title), rep, "Группа", [(g, G[g][key]) for g in groups], 1, share=share)
    w = wb.create_sheet("Средний чек")
    _head(w, ["Группа"] + [label(m, rep) for m in rep["months"]] + ["За 12 мес"], [28] + [12] * 13 + [15])
    for r, g in enumerate(groups, 2):
        w.cell(row=r, column=1, value=g)
        for j, m in enumerate(rep["months"], 2):
            w.cell(row=r, column=j, value=G[g]["ship"][m] / G[g]["n"][m] if G[g]["n"][m] else None).number_format = M
        s, n = sum(G[g]["ship"][m] for m in rep["full"]), sum(G[g]["n"][m] for m in rep["full"])
        w.cell(row=r, column=15, value=s / n if n else None).number_format = M
    w.freeze_panes = "B2"

    for g in groups:
        mem = [c for c in rep["clients"] if c["group"] == g]
        if not mem:
            continue
        w = wb.create_sheet(g[:31])
        _matrix(w, rep, "Клиент", [(c["name"], c["net"]) for c in mem], 1, width0=40)
        x = len(rep["months"]) + 4
        for k, t in enumerate(("Отгрузок за 12 мес", "Месяцев с покупками", "Последняя отгрузка", "Дней без покупок")):
            c = w.cell(row=1, column=x + k, value=t)
            c.font, c.fill = HF, HB
            c.alignment = Alignment(wrap_text=True, horizontal="center")
            w.column_dimensions[get_column_letter(x + k)].width = 13
        for r, c in enumerate(mem, 2):
            w.cell(row=r, column=x, value=c["docs"])
            w.cell(row=r, column=x + 1, value=c["active_months"])
            w.cell(row=r, column=x + 2, value=c["last"]).number_format = "DD.MM.YYYY"
            w.cell(row=r, column=x + 3, value=(today - c["last"]).days if c["last"] else None)

    if debts:
        write_debts(wb, debts)
    h = wb.create_sheet("Как читать")
    h.column_dimensions["A"].width = 30
    h.column_dimensions["B"].width = 100
    for r, (a, b) in enumerate([
        ("Чистые продажи", "Сумма отгрузок минус сумма возвратов покупателей, по ценам отгрузки. Возврат вычитается в день оформления."),
        ("Группы", "По группам (тегам) контрагента в МойСклад, список — в config.toml, раздел [sales]. «Прочие» — покупатели без этих тегов."),
        ("Строи сити", "Отгрузка на базар считается продажей, хотя товар может ещё лежать в ангарах."),
        ("Сегодня / Вчера / 7 дней", "По дате документа в МойСклад. Если отгрузки вносят с опозданием, свежие дни будут занижены."),
        ("Месяц с начала", "С 1-го числа текущего месяца по сегодня."),
        ("Прошлый месяц за те же дни", "С 1-го числа прошлого месяца по то же число — честное сравнение с неполным текущим месяцем."),
        ("Изменение", "Месяц с начала к прошлому месяцу за те же дни."),
        ("Итого 12 мес", "12 полных месяцев до текущего; текущий месяц показан отдельной колонкой и в итог не входит."),
        ("Активные клиенты", "Сколько разных клиентов купили в месяце хотя бы раз."),
        ("Средний чек", "Отгрузки ÷ число документов отгрузки."),
        ("Дебиторка", "Клиенты, которые должны нам по балансу взаиморасчётов МойСклад (отгружено больше, чем оплачено). Поставщики не входят."),
        ("Возраст долга", "Оценка: долг относится на самые свежие отгрузки, более старые считаются оплаченными. «Долг тянется с» — дата самой старой отгрузки, которая ещё не покрыта оплатой. «Старше года» — долг больше всех отгрузок за год."),
        ("0–30 / 31–60 / 61–90 / старше 90", "Какая часть долга приходится на отгрузки такого возраста."),
        ("Дней без платежей", "Сколько дней прошло с последнего входящего платежа клиента. Пусто — платежей за год не было."),
        ("Листы по группам", "Каждый клиент по месяцам, сортировка по сумме за 12 месяцев; справа — число отгрузок, месяцев с покупками, последняя отгрузка и дней без покупок."),
    ], 1):
        h.cell(row=r, column=1, value=a).font = BOLD
        h.cell(row=r, column=2, value=b).alignment = Alignment(wrap_text=True, vertical="top")
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def html_section(rep: dict) -> str:
    n = lambda v: "—" if v is None else f"{v:,.0f}".replace(",", " ")  # noqa: E731
    rows, tot = "", defaultdict(float)
    for g in rep["groups"]:
        p = rep["G"][g]["periods"]
        ch = f"{p['mtd']['net'] / p['prev_mtd']['net'] - 1:+.0%}" if p["prev_mtd"]["net"] > 0 else "—"
        for k in ("yesterday", "d7", "mtd", "prev_mtd"):
            tot[k] += p[k]["net"]
        rows += (f"<tr><td>{g}</td>" + "".join(f'<td align="right">{n(p[k]["net"])}</td>' for k in ("yesterday", "d7", "mtd", "prev_mtd"))
                 + f'<td align="right">{ch}</td><td align="right">{p["mtd"]["n"]}</td><td align="right">{p["mtd"]["clients"]}</td></tr>')
    ch = f"{tot['mtd'] / tot['prev_mtd'] - 1:+.0%}" if tot["prev_mtd"] > 0 else "—"
    rows += ("<tr><td><b>Итого</b></td>" + "".join(f'<td align="right"><b>{n(tot[k])}</b></td>' for k in ("yesterday", "d7", "mtd", "prev_mtd"))
             + f'<td align="right"><b>{ch}</b></td><td></td><td></td></tr>')
    silent = [(g, rep["G"][g]["silent"]) for g in rep["groups"] if rep["G"][g]["silent"]]
    extra = ""
    if silent:
        extra = "<p><b>Крупные клиенты без покупок больше 45 дней:</b> " + "; ".join(
            f"{g} — {', '.join(s)}" for g, s in silent) + "</p>"
    return ("<h3>Продажи по группам клиентов</h3>"
            '<table border="1" cellpadding="4" style="border-collapse:collapse;font-size:13px;border-color:#ddd">'
            '<tr style="background:#1F4E78;color:#fff"><th align="left">Группа</th><th>Вчера</th><th>7 дней</th><th>Месяц с начала</th>'
            "<th>Прошлый месяц за те же дни</th><th>Изменение</th><th>Отгрузок за месяц</th><th>Клиентов за месяц</th></tr>"
            f"{rows}</table>{extra}")
