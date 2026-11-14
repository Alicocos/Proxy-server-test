"""Отправка отчёта на почту через API Brevo (бывший Sendinblue).

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
    k = rep["kpi"]
    rows = [
        ("Запас на складах", _fmt(k["stock_value"])),
        ("Заморожено в товаре", f"{_fmt(k['frozen'])} ({_fmt(k['frozen_share'], True)})"),
        ("Срочно заказать, позиций", _fmt(k["urgent_sku"])),
        ("Рекомендовано к заказу", _fmt(k["rec_value"])),
        ("Товар в пути", _fmt(k["transit_value"])),
        ("Открытых заказов / просрочено", f"{k['orders_open']} / {k['orders_overdue']}"),
        ("Депозиты у заводов", _fmt(k["deposits"])),
        ("Наш долг заводам", _fmt(k["debts"])),
    ]
    trs = "".join(f"<tr><td style='padding:4px 12px 4px 0'>{a}</td><td style='padding:4px 0;text-align:right'><b>{b}</b></td></tr>"
                  for a, b in rows)
    late = [o for o in rep.get("orders", []) if o["status"] in ("Просрочен", "Нет плановой даты")]
    extra = ""
    if late:
        items = "".join(f"<li>№{o['name']} — {o['agent']}: {o['status'].lower()}</li>" for o in late)
        extra = f"<p><b>Заказы, требующие внимания:</b></p><ul>{items}</ul>"
    return (f"<p>Отчёт по закупу и логистике на {rep['today']:%d.%m.%Y %H:%M} (Алматы). "
            f"Суммы в тенге, товар по себестоимости. Полный отчёт — во вложении.</p>"
            f"<table style='border-collapse:collapse;font-family:Arial,sans-serif;font-size:14px'>{trs}</table>{extra}")


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
