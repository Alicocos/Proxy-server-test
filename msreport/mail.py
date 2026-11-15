"""Текст письма с отчётом и отправка на почту через API Brevo (бывший Sendinblue).

Ключ API подставляет прокси окружения в заголовок api-key, поэтому в коде его нет.
Адреса передаются при запуске и в репозитории не хранятся.
"""
import base64
from pathlib import Path

import requests

URL = "https://api.brevo.com/v3/smtp/email"


def _fmt(v, pct=False):
    if v is None:
        return "—"
    return f"{v:.0%}" if pct else f"{v:,.0f}".replace(",", " ")


def summary_html(rep: dict) -> str:
    """Текст письма: главные цифры и три короткие таблицы. Без вложения — Excel остаётся в чате."""
    from . import analytics as a
    k = rep["kpi"]
    td = lambda v, right=True: f"<td{' align=\"right\"' if right else ''}>{v}</td>"  # noqa: E731
    def table(head, rows):
        h = "".join(f"<th{' align=\"left\"' if i == 0 else ''}>{x}</th>" for i, x in enumerate(head))
        b = "".join("<tr>" + "".join(td(v, i > 0) for i, v in enumerate(r)) + "</tr>" for r in rows)
        return ('<table border="1" cellpadding="4" style="border-collapse:collapse;font-size:13px;border-color:#ddd">'
                f'<tr style="background:#1F4E78;color:#fff">{h}</tr>{b}</table>')
    kpi = [
        ("Запас на складах", _fmt(k["stock_value"])),
        ("Заморожено в товаре", f"{_fmt(k['frozen'])} ({_fmt(k['frozen_share'], True)})"),
        ("Срочно заказать, позиций", _fmt(k["urgent_sku"])),
        ("Рекомендовано к заказу", _fmt(k["rec_value"])),
        ("Товар в пути", _fmt(k["transit_value"])),
        ("Открытых заказов / просрочено", f"{k['orders_open']} / {k['orders_overdue']}"),
        ("Депозиты у заводов", f"{_fmt(k['deposits'])} (без движения {_fmt(k['stale'])})"),
        ("Наш долг заводам", _fmt(k["debts"])),
        ("Оплачено заводам за 30 дней", _fmt(k["pay30"])),
        ("Отгружено на базар за 90 дней", f"{_fmt(k['baz_sum'])} ({_fmt(k['baz_share'], True)} отгрузок)"),
    ]
    head = "".join(f"<tr><td>{x}</td><td align=\"right\"><b>{v}</b></td></tr>" for x, v in kpi)
    urgent = sorted((p for p in rep["products"] if p["status"] == a.URGENT), key=lambda p: -p["rec_value"])[:15]
    frozen = sorted((p for p in rep["products"] if p["frozen"] > 0), key=lambda p: -p["frozen"])[:10]
    stale = [s for s in rep["suppliers"] if s["flag"] and (s["deposit"] or 0) >= 100000]
    late = [o for o in rep.get("orders", []) if o["status"] in ("Просрочен", "Нет плановой даты")]
    parts = [
        f'<div style="font-family:Arial,sans-serif;font-size:14px"><p>Отчёт по закупу и логистике на '
        f'{rep["today"]:%d.%m.%Y %H:%M} (Алматы). Суммы в тенге, товар по себестоимости МойСклад.</p>',
        f'<table cellpadding="4" style="border-collapse:collapse;font-size:13px">{head}</table>',
        "<h3>Срочно заказать — топ-15 по сумме</h3>",
        table(["Товар", "Остаток", "Хватит, дней", "Заказать, шт", "Сумма"],
              [[p["name"], _fmt(p["stock"]), _fmt(p["cover_transit"]), _fmt(p["rec"]), _fmt(p["rec_value"])] for p in urgent]),
        "<h3>Больше всего заморожено — топ-10</h3>",
        table(["Товар", "Статус", "Запас, дней", "Заморожено"],
              [[p["name"], p["status"], _fmt(p["cover"]), _fmt(p["frozen"])] for p in frozen]),
    ]
    if late:
        parts += ["<h3>Заказы поставщикам, требующие внимания</h3>",
                  table(["Заказ", "Завод", "Статус", "Ещё не принято"],
                        [[o["name"], o["agent"], o["status"], _fmt(o["left_sum"])] for o in late])]
    if stale:
        parts += ["<h3>Предоплата без движения</h3>",
                  table(["Завод", "Депозит", "Дней без приёмок"],
                        [[s["agent"], _fmt(s["deposit"]),
                          _fmt(s["days_no_sup"]) if s["days_no_sup"] is not None else "приёмок за год не было"] for s in stale])]
    parts.append('<p style="color:#777;font-size:12px">«Хватит, дней» — на сколько дней хватит остатка вместе с товаром в пути. '
                 "«Заморожено» — сумма сверх нормы запаса или весь остаток без продаж за 90 дней. "
                 "Полный отчёт в Excel — в чате Claude «Отчетность логистики и закупок».</p></div>")
    return "".join(parts)


def send(path: Path, rep: dict, to: list[str], sender: str, sender_name: str = "Отчёт закуп и логистика",
         api_key: str = "proxy-injected", session=None) -> dict:
    if not to:
        raise ValueError("Не указаны адреса получателей")
    payload = {
        "sender": {"name": sender_name, "email": sender},
        "to": [{"email": a} for a in to],
        "subject": f"Отчёт закуп и логистика — {rep['today']:%d.%m.%Y %H:%M}",
        "htmlContent": summary_html(rep),
        "attachment": [{"name": Path(path).name, "content": base64.b64encode(Path(path).read_bytes()).decode()}],
    }
    r = (session or requests).post(URL, json=payload, timeout=60,
                                   headers={"api-key": api_key, "accept": "application/json"})
    if r.status_code >= 300:
        raise RuntimeError(f"Brevo вернул {r.status_code}: {r.text[:300]}")
    return r.json()
