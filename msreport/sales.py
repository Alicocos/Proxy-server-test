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


def yoy_start(today: datetime) -> datetime:
    """Начало окна для сверки год к году: тот же месяц два года назад."""
    return datetime(today.year - 2, today.month, 1)


def _excluded_ids(raw: dict, cfg: dict) -> set:
    words = [w.lower() for w in cfg["sales"].get("exclude_agents", [])]
    return {c["id"] for c in raw.get("counterparties", []) if any(w in c.get("name", "").lower() for w in words)}


def build_yoy(raw: dict, cfg: dict, today: datetime) -> dict:
    """Сверка по месяцам: каждый из последних 13 месяцев против того же месяца годом раньше.

    Текущий месяц сравнивается с теми же днями прошлого года, чтобы неполный месяц не выглядел провалом.
    """
    groups = [(g["name"], g["tag"].lower()) for g in cfg["sales"]["groups"]]
    names = [g for g, _ in groups] + [OTHER]
    cp = {c["id"]: c for c in raw.get("counterparties", [])}

    def group(aid):
        tags = {t.lower() for t in cp.get(aid, {}).get("tags", [])}
        return next((g for g, tag in groups if tag in tags), OTHER)

    cur = f"{today.year}-{today.month:02d}"
    months = _months(today)  # 13 месяцев: с того же месяца прошлого года по текущий
    prev = lambda m: f"{int(m[:4]) - 1}{m[4:]}"  # noqa: E731
    net = defaultdict(float)  # (группа, месяц) -> сумма
    first_doc = None
    skip = _excluded_ids(raw, cfg)
    for kind, sign in (("demand", 1), ("retaildemand", 1), ("salesreturn", -1), ("retailsalesreturn", -1)):
        for d in raw.get("sales_docs", {}).get(kind, []):
            if d.get("applicable") is False:
                continue
            m = UUID.search((d.get("agent") or {}).get("meta", {}).get("href", ""))
            if not m or m.group(0) in skip:
                continue
            first_doc = min(first_doc or d["moment"], d["moment"])
            ym, day = d["moment"][:7], int(d["moment"][8:10])
            # для прошлогоднего «текущего месяца» берём только те же дни
            if ym == prev(cur) and day > today.day:
                ym = ym + "+"
            net[(group(m.group(0)), ym)] += sign * (d.get("sum") or 0) / 100
    # месяц сравниваем, только если учёт в МойСклад уже вёлся с его первого числа
    has = lambda m: bool(first_doc) and (first_doc[:7] < m or (first_doc[:7] == m and first_doc[8:10] == "01"))  # noqa: E731
    rows = []
    for m in months:
        pm = prev(m)
        row = {"month": m, "prev_month": pm, "partial": m == cur, "has_prev": has(pm)}
        for g in names + ["ИТОГО"]:
            gs = names if g == "ИТОГО" else [g]
            a = sum(net[(x, m)] for x in gs)
            b = sum(net[(x, pm)] for x in gs)
            row[g] = (a, b if row["has_prev"] else None)
        rows.append(row)
    full = [r for r in rows if not r["partial"]][-12:]
    year = {g: (sum(r[g][0] for r in full), sum(r[g][1] or 0 for r in full) if all(r["has_prev"] for r in full) else None)
            for g in names + ["ИТОГО"]}
    shown = [g for g in names if g != OTHER or any(r[g][0] or (r[g][1] or 0) for r in rows)]
    return {"today": today, "rows": rows, "groups": shown, "year": year, "first_doc": first_doc,
            "compared": sum(1 for r in rows if r["has_prev"])}


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

    # торговые представители: по метке клиента или по самому покупателю (например, розница Фараби)
    reps = cfg["sales"].get("reps", [])

    def rep_of(aid):
        c = cp.get(aid, {})
        tags, name = {t.lower() for t in c.get("tags", [])}, c.get("name", "").lower()
        for r in reps:
            if r.get("agent") and r["agent"].lower() in name:
                return r["name"]
        for r in reps:
            if r.get("tag") and r["tag"].lower() in tags:
                return r["name"]
        return "Остальные"

    rep_names = [r["name"] for r in reps] + ["Остальные"]
    month_start = today.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    days = [(month_start + timedelta(days=i)).date() for i in range((today - month_start).days + 1)]
    daily = {d: defaultdict(float) for d in days}  # день -> представитель -> чистые продажи
    daily_n = {d: defaultdict(int) for d in days}

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
    skip = _excluded_ids(raw, cfg)
    for kind, sign in (("demand", 1), ("retaildemand", 1), ("salesreturn", -1), ("retailsalesreturn", -1)):
        for d in docs.get(kind, []):
            if d.get("applicable") is False:
                continue
            m = UUID.search((d.get("agent") or {}).get("meta", {}).get("href", ""))
            when = datetime.strptime(d["moment"][:19], "%Y-%m-%d %H:%M:%S")
            ym = d["moment"][:7]
            if not m or ym not in months or m.group(0) in skip:
                continue
            aid, v = m.group(0), (d.get("sum") or 0) / 100
            c = cl[aid]
            if sign > 0:
                c["ship"][ym] += v
                c["n"][ym] += 1
                c["last"] = max(c["last"] or when, when)
            else:
                c["ret"][ym] += v
            if reps and when.date() in daily:
                rp = rep_of(aid)
                daily[when.date()][rp] += sign * v
                if sign > 0:
                    daily_n[when.date()][rp] += 1
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
    rep_shown = [r for r in rep_names if any(daily[d][r] for d in days)]
    return {"daily": daily, "daily_n": daily_n, "days": days, "reps": rep_shown if reps else [],
            "today": today, "months": months, "full": full, "cur": cur, "groups": shown, "G": G, "clients": clients,
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
    skip = _excluded_ids(raw, cfg)
    for c in raw.get("counterparties", []):
        if c["id"] in skip:
            continue
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
    other = [r for r in rows if r["group"] == OTHER]  # не клиенты: учредители, инвесторы, сотрудники и т.п.
    rows = [r for r in rows if r["group"] != OTHER]
    names = [g for g, _ in groups]
    by_group = [{"group": g, "n": sum(1 for r in rows if r["group"] == g),
                 **{k: sum(r[k] for r in rows if r["group"] == g) for k in ("debt", "b0", "b1", "b2", "b3")}} for g in names]
    by_group = [g for g in by_group if g["n"]]
    total = sum(r["debt"] for r in rows)
    return {"today": today, "rows": rows, "by_group": by_group, "total": total,
            "other": other, "other_total": sum(r["debt"] for r in other),
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
    red = PatternFill("solid", fgColor="F8CBAD")
    for pos, (title, data) in enumerate((("Дебиторка", debts["rows"]), ("Прочие долги", debts["other"])), 1):
        if title == "Прочие долги" and not data:
            continue
        ws = wb.create_sheet(title, pos if pos == 1 else 3)
        _debt_sheet(ws, cols, data, red)
    _debt_groups(wb, debts)


def _debt_sheet(ws, cols, data, red):
    _head(ws, [c[0] for c in cols], [c[2] for c in cols])
    for r, row in enumerate(data, 2):
        for i, (_, k, _, f) in enumerate(cols, 1):
            c = ws.cell(row=r, column=i, value=row[k] or None if k.startswith("b") else row[k])
            if f:
                c.number_format = f
            if k == "flag" and row[k]:
                c.fill = red
    n = len(data) + 2
    ws.cell(row=n, column=1, value="ИТОГО").font = BOLD
    for i, (_, k, _, f) in enumerate(cols, 1):
        if k in ("debt", "b0", "b1", "b2", "b3", "ship90", "pay90"):
            col = get_column_letter(i)
            c = ws.cell(row=n, column=i, value=f"=SUM({col}2:{col}{n - 1})")
            c.font, c.number_format = BOLD, M
    ws.freeze_panes = "B2"
    if data:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{len(data) + 1}"


def _debt_groups(wb, debts):
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


def write_debts_book(debts: dict, path: Path) -> Path:
    """Отдельный файл по дебиторке: сводка, список должников, группы, прочие долги."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Сводка"
    ws["A1"] = f"Дебиторка на {debts['today']:%d.%m.%Y %H:%M} (Алматы)"
    ws["A1"].font = Font(bold=True, size=14)
    lines = [("Клиенты должны всего", debts["total"], M), ("Должников", debts["n"], M),
             ("Из долга старше 60 дней", debts["old"], M),
             ("Доля пяти крупнейших должников", debts["top5_share"], "0%"),
             ("Не клиенты должны (лист «Прочие долги»)", debts["other_total"], M)]
    for r, (t, v, f) in enumerate(lines, 3):
        ws.cell(row=r, column=1, value=t)
        ws.cell(row=r, column=2, value=v).number_format = f
    ws.column_dimensions["A"].width = 44
    ws.column_dimensions["B"].width = 18
    write_debts(wb, debts)
    h = wb.create_sheet("Как читать")
    h.column_dimensions["A"].width = 30
    h.column_dimensions["B"].width = 100
    for r, (a, b) in enumerate(DEBT_HELP, 1):
        h.cell(row=r, column=1, value=a).font = BOLD
        h.cell(row=r, column=2, value=b).alignment = Alignment(wrap_text=True, vertical="top")
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


DEBT_HELP = [
    ("Дебиторка", "Клиенты, которые должны нам по балансу взаиморасчётов МойСклад (отгружено больше, чем оплачено). Поставщики не входят."),
    ("Прочие долги", "Контрагенты с долгом перед нами, которые не клиенты: учредители, инвесторы, сотрудники и другие без клиентских тегов. В дебиторку и её итоги не входят."),
    ("Возраст долга", "Оценка: оплата гасит самые старые отгрузки первыми, долг относится на самые свежие. «Долг тянется с» — дата самой старой отгрузки, ещё не покрытой оплатой. «Старше года» — долг больше всех отгрузок за год."),
    ("0–30 / 31–60 / 61–90 / старше 90", "Какая часть долга приходится на отгрузки такого возраста."),
    ("Внимание", "«Не платит больше 30/60 дней» — столько прошло с последнего платежа. «Нет платежей за год» — ни одного входящего платежа за 12 месяцев."),
    ("Дней без платежей", "Сколько дней прошло с последнего входящего платежа клиента."),
    ("Отгружено / оплачено за 90 дн", "Если отгружено заметно больше, чем оплачено, долг растёт."),
]


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
            + (f'<p style="color:#777;font-size:12px">Не клиенты (учредители, инвесторы, сотрудники и т.п.) должны ещё '
               f'{n(debts["other_total"])} — они на листе «Прочие долги» файла «Дебиторка» и в эту сумму не входят.</p>' if debts["other_total"] else "") +
            f'<p><b>Топ-10 должников</b></p>'
            f'{tb}{th}<th align="left">Клиент</th><th>Группа</th><th>Долг</th><th>Возраст долга, дней</th>'
            f'<th>Последний платёж</th><th>Дней без платежей</th></tr>{top}</table>')


def _pct(a, b):
    return (a / b - 1) if b and b > 0 else None


def write_yoy(wb, yoy: dict, index: int = 1) -> None:
    ws = wb.create_sheet("Сверка по месяцам", index)
    today = yoy["today"]
    ws["A1"] = "Сверка по месяцам: этот год против того же месяца прошлого года"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = (f"Чистые продажи (отгрузки минус возвраты), тенге. Текущий месяц сравнивается с 1–{today.day} числом "
                "того же месяца прошлого года. «нет данных» — учёт в МойСклад за тот месяц ещё не вёлся или вёлся не с начала месяца"
                + (f" (первая отгрузка в МойСклад — {yoy['first_doc'][8:10]}.{yoy['first_doc'][5:7]}.{yoy['first_doc'][:4]})." if yoy.get("first_doc") else "."))
    cols = ["ИТОГО"] + yoy["groups"]
    heads = ["Месяц"]
    for g in cols:
        heads += [f"{g}: этот год", "прошлый год", "изменение"]
    _head(ws, heads, [22] + [14, 14, 10] * len(cols), 4)
    rows = yoy["rows"] + [None]
    for r, row in enumerate(rows, 5):
        if row is None:
            title = "12 полных месяцев"
            vals = [yoy["year"][g] for g in cols]
        else:
            m = row["month"]
            title = f"{RU[int(m[5:]) - 1]} {m[:4]} vs {m[:4] and int(m[:4]) - 1}" + (f" (1–{today.day})" if row["partial"] else "")
            vals = [row[g] for g in cols]
        c = ws.cell(row=r, column=1, value=title)
        if row is None:
            c.font = BOLD
        for j, (a, b) in enumerate(vals):
            col = 2 + j * 3
            ws.cell(row=r, column=col, value=a).number_format = M
            ws.cell(row=r, column=col + 1, value=b if b is not None else "нет данных").number_format = M
            p = _pct(a, b)
            cc = ws.cell(row=r, column=col + 2, value=p)
            cc.number_format = "+0%;-0%;0%"
            if p is not None:
                cc.fill = PatternFill("solid", fgColor="C6E0B4" if p >= 0 else "F8CBAD")
            if row is None:
                for k in range(3):
                    ws.cell(row=r, column=col + k).font = BOLD
    ws.freeze_panes = "B5"


def yoy_html(yoy: dict) -> str:
    n = lambda v: "нет данных" if v is None else f"{v:,.0f}".replace(",", " ")  # noqa: E731
    pc = lambda a, b: f"{_pct(a, b):+.0%}" if _pct(a, b) is not None else "—"  # noqa: E731
    cur, last = yoy["rows"][-1], [r for r in yoy["rows"] if not r["partial"]][-1]
    lab = lambda m: f"{RU[int(m[5:]) - 1]} {m[:4]}"  # noqa: E731
    hdr = (f"<th align='left'>Группа</th><th>{lab(cur['month'])} (1–{yoy['today'].day})</th>"
           f"<th>{lab(cur['prev_month'])} (1–{yoy['today'].day})</th><th>Изм.</th>"
           f"<th>{lab(last['month'])}</th><th>{lab(last['prev_month'])}</th><th>Изм.</th>")
    body = ""
    for g in ["ИТОГО"] + yoy["groups"]:
        a, b = cur[g]
        c, d = last[g]
        b_ = "<b>" if g == "ИТОГО" else ""
        e_ = "</b>" if g == "ИТОГО" else ""
        body += (f"<tr><td>{b_}{g}{e_}</td><td align='right'>{b_}{n(a)}{e_}</td><td align='right'>{n(b)}</td>"
                 f"<td align='right'>{pc(a, b)}</td><td align='right'>{b_}{n(c)}{e_}</td><td align='right'>{n(d)}</td>"
                 f"<td align='right'>{pc(c, d)}</td></tr>")
    return ("<h3>Сверка с прошлым годом</h3>"
            '<table border="1" cellpadding="4" style="border-collapse:collapse;font-size:13px;border-color:#ddd">'
            f'<tr style="background:#1F4E78;color:#fff">{hdr}</tr>{body}</table>'
            "<p style='color:#777;font-size:12px'>Все 13 месяцев с прошлым годом — на листе «Сверка по месяцам».</p>")


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


WD = ["пн", "вт", "ср", "чт", "пт", "сб", "вс"]


def _daily_block(ws, rep: dict, top: int) -> int:
    """Ежедневная динамика текущего месяца по торговым представителям. Возвращает последнюю занятую строку."""
    reps, days = rep["reps"], rep["days"]
    ws.cell(row=top - 1, column=1, value=f"Продажи по дням: {RU[rep['today'].month - 1]} {rep['today'].year}, "
                                         "по торговым представителям").font = BOLD
    _head(ws, ["День"] + reps + ["Итого за день", "Отгрузок"], [28] + [14] * len(reps) + [15, 10], top)
    r = top
    for r, d in enumerate(days, top + 1):
        ws.cell(row=r, column=1, value=f"{d:%d.%m} {WD[d.weekday()]}")
        vals = [rep["daily"][d][x] for x in reps]
        for j, v in enumerate(vals, 2):
            ws.cell(row=r, column=j, value=round(v, 2) or None).number_format = M
        ws.cell(row=r, column=len(reps) + 2, value=sum(vals) or None).number_format = M
        ws.cell(row=r, column=len(reps) + 3, value=sum(rep["daily_n"][d].values()) or None)
        if d.weekday() == 0:  # понедельник — выходной
            for j in range(1, len(reps) + 4):
                ws.cell(row=r, column=j).font = Font(color="999999")
    r += 1
    ws.cell(row=r, column=1, value="ИТОГО с начала месяца").font = BOLD
    for j in range(2, len(reps) + 4):
        col = get_column_letter(j)
        c = ws.cell(row=r, column=j, value=f"=SUM({col}{top + 1}:{col}{r - 1})")
        c.font, c.number_format = BOLD, M
    r += 1
    work = [d for d in days if d.weekday() != 0 and d < rep["today"].date()] or days
    ws.cell(row=r, column=1, value="В среднем за рабочий день").font = Font(italic=True)
    for j, x in enumerate(reps, 2):
        ws.cell(row=r, column=j, value=sum(rep["daily"][d][x] for d in work) / len(work)).number_format = M
    ws.cell(row=r, column=len(reps) + 2, value=sum(sum(rep["daily"][d][x] for x in reps) for d in work) / len(work)).number_format = M
    ch = LineChart()
    ch.title, ch.height, ch.width = "Продажи по дням", 8, 26
    ch.y_axis.numFmt = '#,##0,,"млн"'
    ch.add_data(Reference(ws, min_col=2, max_col=len(reps) + 1, min_row=top, max_row=top + len(days)), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=top + 1, max_row=top + len(days)))
    ws.add_chart(ch, f"{get_column_letter(len(reps) + 5)}{top}")
    return r


def write_sales(rep: dict, path: Path, yoy: dict | None = None, gd: dict | None = None) -> Path:
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
    if rep["reps"]:
        top = _daily_block(ws, rep, top) + 3
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

    if yoy:
        write_yoy(wb, yoy)
    if gd:
        from . import gd as gdmod
        gdmod.write(wb, gd, 1)
    h = wb.create_sheet("Как читать")
    h.column_dimensions["A"].width = 30
    h.column_dimensions["B"].width = 100
    for r, (a, b) in enumerate([
        ("Сверка по месяцам", "Каждый месяц против того же месяца прошлого года, по группам и итого. Текущий месяц — против тех же дней прошлого года. Внизу — 12 полных месяцев против предыдущих 12."),
        ("Чистые продажи", "Сумма отгрузок минус сумма возвратов покупателей, по ценам отгрузки. Возврат вычитается в день оформления."),
        ("Группы", "По группам (тегам) контрагента в МойСклад, список — в config.toml, раздел [sales]. «Прочие» — покупатели без этих тегов."),
        ("Строи сити", "В этот отчёт не входит: по Строи сити — отдельный отчёт из его собственного аккаунта МойСклад."),
        ("Сегодня / Вчера / 7 дней", "По дате документа в МойСклад. Если отгрузки вносят с опозданием, свежие дни будут занижены."),
        ("Месяц с начала", "С 1-го числа текущего месяца по сегодня."),
        ("Прошлый месяц за те же дни", "С 1-го числа прошлого месяца по то же число — честное сравнение с неполным текущим месяцем."),
        ("Изменение", "Месяц с начала к прошлому месяцу за те же дни."),
        ("Итого 12 мес", "12 полных месяцев до текущего; текущий месяц показан отдельной колонкой и в итог не входит."),
        ("Активные клиенты", "Сколько разных клиентов купили в месяце хотя бы раз."),
        ("Средний чек", "Отгрузки ÷ число документов отгрузки."),
        ("Листы по группам", "Каждый клиент по месяцам, сортировка по сумме за 12 месяцев; справа — число отгрузок, месяцев с покупками, последняя отгрузка и дней без покупок."),
    ], 1):
        h.cell(row=r, column=1, value=a).font = BOLD
        h.cell(row=r, column=2, value=b).alignment = Alignment(wrap_text=True, vertical="top")
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def daily_html(rep: dict, last: int = 7) -> str:
    if not rep.get("reps"):
        return ""
    n = lambda v: f"{v:,.0f}".replace(",", " ") if v else "—"  # noqa: E731
    days = rep["days"][-last:]
    head = "".join(f"<th>{x}</th>" for x in rep["reps"])
    body = ""
    for d in days:
        vals = [rep["daily"][d][x] for x in rep["reps"]]
        grey = ' style="color:#999"' if d.weekday() == 0 else ""
        body += (f"<tr{grey}><td>{d:%d.%m} {WD[d.weekday()]}</td>" + "".join(f'<td align="right">{n(v)}</td>' for v in vals)
                 + f'<td align="right"><b>{n(sum(vals))}</b></td></tr>')
    tot = [sum(rep["daily"][d][x] for d in rep["days"]) for x in rep["reps"]]
    body += ("<tr><td><b>С начала месяца</b></td>" + "".join(f'<td align="right"><b>{n(v)}</b></td>' for v in tot)
             + f'<td align="right"><b>{n(sum(tot))}</b></td></tr>')
    return ("<h3>Продажи по дням по торговым представителям</h3>"
            '<table border="1" cellpadding="4" style="border-collapse:collapse;font-size:13px;border-color:#ddd">'
            f'<tr style="background:#1F4E78;color:#fff"><th align="left">День</th>{head}<th>Итого</th></tr>{body}</table>'
            "<p style='color:#777;font-size:12px'>Последние 7 дней; весь месяц по дням — на первой странице файла.</p>")


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
    silent = [(g, rep["G"][g]["silent"]) for g in rep["groups"] if rep["G"][g]["silent"] and g != OTHER]
    extra = ""
    if silent:
        extra = "<p><b>Крупные клиенты без покупок больше 45 дней:</b> " + "; ".join(
            f"{g} — {', '.join(s)}" for g, s in silent) + "</p>"
    return (daily_html(rep) + "<h3>Продажи по группам клиентов</h3>"
            '<table border="1" cellpadding="4" style="border-collapse:collapse;font-size:13px;border-color:#ddd">'
            '<tr style="background:#1F4E78;color:#fff"><th align="left">Группа</th><th>Вчера</th><th>7 дней</th><th>Месяц с начала</th>'
            "<th>Прошлый месяц за те же дни</th><th>Изменение</th><th>Отгрузок за месяц</th><th>Клиентов за месяц</th></tr>"
            f"{rows}</table>{extra}")
