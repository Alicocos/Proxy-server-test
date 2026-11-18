#!/usr/bin/env python3
"""Отчёт по закупу и логистике из МойСклад.

    python run.py            # боевой отчёт (нужен MOYSKLAD_TOKEN в окружении или в файле .env)
    python run.py --demo     # отчёт на выдуманных данных, без токена
"""
import argparse
import os
import sys
import tomllib
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

from msreport import analytics, demo, excel, manual, sales
from msreport.api import ApiError, MoySklad
from msreport.fetch import fetch_all

ROOT = Path(__file__).resolve().parent


def token() -> str | None:
    if os.environ.get("MOYSKLAD_TOKEN"):
        return os.environ["MOYSKLAD_TOKEN"].strip()
    env = ROOT / ".env"
    if env.exists():
        for line in env.read_text(encoding="utf-8").splitlines():
            if line.strip().startswith("MOYSKLAD_TOKEN="):
                return line.split("=", 1)[1].strip().strip("\"'")
    return None


def run_account(args, cfg: dict, today: datetime) -> int:
    """Отчёты «Продажи по клиентам» и «Дебиторка» для второго аккаунта МойСклад."""
    import copy
    from msreport.fetch import fetch_sales_only
    acc = cfg["accounts"][args.account]
    acfg = copy.deepcopy(cfg)
    acfg["sales"]["groups"] = acc["groups"]
    acfg["sales"]["exclude_agents"] = acc.get("exclude_agents", [])
    title, tag = acc["title"], acc["title"].replace(" ", "_")
    if args.raw:
        import pickle
        saved = pickle.loads(Path(args.raw).read_bytes())
        raw, today = saved["raw"], saved["today"]
    else:
        tok = token()
        if not tok:
            print("Не найден MOYSKLAD_TOKEN", file=sys.stderr)
            return 2
        try:
            raw = fetch_sales_only(MoySklad(tok, acc.get("base_url", "https://api.moysklad.ru/api/remap/1.2")), acfg, today)
        except ApiError as e:
            print(f"Ошибка: {e}", file=sys.stderr)
            return 1
        if args.save_raw:
            import pickle
            Path(args.save_raw).write_bytes(pickle.dumps({"raw": raw, "today": today}))
    stamp = f"{today:%Y-%m-%d_%H%M}"
    folder = ROOT / cfg["files"]["reports_dir"]
    srep, yoy = sales.build_sales(raw, acfg, today), sales.build_yoy(raw, acfg, today)
    sout = sales.write_sales(srep, folder / f"Продажи_по_клиентам_{tag}_{stamp}.xlsx", yoy)
    print(f"Продажи ({title}): {sout}")
    if raw.get("forbidden"):
        print(f"Нет прав в МойСклад на: {', '.join(raw['forbidden'])} — эти документы в отчёт не вошли")
    debts = sales.build_debts(raw, acfg, today)
    if debts:
        dout = sales.write_debts_book(debts, folder / f"Дебиторка_{tag}_{stamp}.xlsx")
        print(f"Дебиторка ({title}): {dout}")
    if args.html:
        wrap = lambda t, body: (f'<div style="font-family:Arial,sans-serif;font-size:14px"><p>{t} — {title}, на '  # noqa: E731
                                f'{today:%d.%m.%Y %H:%M} (Алматы). Полный отчёт — в файле Excel во вложении.</p>{body}</div>')
        base = Path(args.html)
        base.with_name(base.stem + "_sales.html").write_text(wrap("Продажи по клиентам", sales.html_section(srep) + sales.yoy_html(yoy)), encoding="utf-8")
        if debts:
            base.with_name(base.stem + "_debts.html").write_text(wrap("Дебиторка", sales.debts_html(debts)), encoding="utf-8")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--account", help="второй аккаунт из [accounts] в config.toml: только продажи по клиентам и дебиторка")
    ap.add_argument("--demo", action="store_true", help="выдуманные данные вместо МойСклад")
    ap.add_argument("--out", help="путь к файлу отчёта")
    ap.add_argument("--save-raw", help="сохранить выгрузку из МойСклад в файл (для отладки без повторной выгрузки)")
    ap.add_argument("--raw", help="взять выгрузку из файла, сохранённого --save-raw, вместо обращения к МойСклад")
    ap.add_argument("--html", help="сохранить текст письма (HTML) в этот файл")
    ap.add_argument("--email", help="отправить отчёт на эти адреса (через запятую); можно задать MAIL_TO")
    ap.add_argument("--sender", help="адрес отправителя, подтверждённый в Brevo; можно задать MAIL_FROM")
    args = ap.parse_args()

    cfg = tomllib.loads((ROOT / "config.toml").read_text(encoding="utf-8"))
    today = datetime.now(ZoneInfo(cfg["files"].get("timezone", "Asia/Almaty"))).replace(microsecond=0, tzinfo=None)

    if args.account:
        return run_account(args, cfg, today)

    if args.demo:
        raw, logistics = demo.demo_raw(today), demo.demo_logistics(today)
    else:
        if args.raw:
            import pickle
            saved = pickle.loads(Path(args.raw).read_bytes())
            raw, today = saved["raw"], saved["today"]
        else:
            tok = token()
            if not tok:
                print("Не найден MOYSKLAD_TOKEN. Задайте переменную окружения или создайте файл .env "
                      "со строкой MOYSKLAD_TOKEN=ваш_токен. Для пробы: python run.py --demo", file=sys.stderr)
                return 2
            try:
                raw = fetch_all(MoySklad(tok), cfg, today)
            except ApiError as e:
                print(f"Ошибка: {e}", file=sys.stderr)
                return 1
            if args.save_raw:
                import pickle
                Path(args.save_raw).write_bytes(pickle.dumps({"raw": raw, "today": today}))
        logistics = manual.read_logistics(ROOT / cfg["files"]["logistics"])

    rep = analytics.build_report(raw, logistics, cfg, today)
    name = f"Отчет_закуп_логистика_{today:%Y-%m-%d_%H%M}{'_ДЕМО' if args.demo else ''}.xlsx"
    out = Path(args.out) if args.out else ROOT / cfg["files"]["reports_dir"] / name
    excel.write_report(rep, out)
    srep = sales.build_sales(raw, cfg, today)
    sout = out.with_name(out.name.replace("Отчет_закуп_логистика", "Продажи_по_клиентам"))
    if sout == out:
        sout = out.with_name("Продажи_" + out.name)
    debts = sales.build_debts(raw, cfg, today)
    yoy = sales.build_yoy(raw, cfg, today)
    sales.write_sales(srep, sout, yoy)
    dout = None
    if debts:
        dout = sout.with_name(sout.name.replace("Продажи_по_клиентам", "Дебиторка"))
        if dout == sout:
            dout = sout.with_name("Дебиторка_" + sout.name)
        sales.write_debts_book(debts, dout)

    if args.html:
        from msreport import mail
        # три письма — по одному на файл, чтобы каждый отчёт приходил отдельно и не терялся
        wrap = lambda title, body: (f'<div style="font-family:Arial,sans-serif;font-size:14px"><p>{title} на '  # noqa: E731
                                    f'{today:%d.%m.%Y %H:%M} (Алматы). Полный отчёт — в файле Excel во вложении.</p>{body}</div>')
        base = Path(args.html)
        base.write_text(mail.summary_html(rep), encoding="utf-8")
        base.with_name(base.stem + "_sales.html").write_text(wrap("Продажи по группам клиентов", sales.html_section(srep) + sales.yoy_html(yoy)), encoding="utf-8")
        if debts:
            base.with_name(base.stem + "_debts.html").write_text(wrap("Дебиторка", sales.debts_html(debts)), encoding="utf-8")
        print(f"Тексты писем: {base}, {base.stem}_sales.html, {base.stem}_debts.html")

    to = [a.strip() for a in (args.email or os.environ.get("MAIL_TO", "")).split(",") if a.strip()]
    if to:
        from msreport import mail
        sender = args.sender or os.environ.get("MAIL_FROM") or to[0]
        try:
            mail.send(out, rep, to, sender, api_key=os.environ.get("BREVO_API_KEY", "proxy-injected"))
            print(f"Отправлено на: {', '.join(to)}")
        except Exception as e:  # отчёт уже собран — сообщаем об ошибке почты, но не теряем файл
            print(f"Почта не отправлена: {e}", file=sys.stderr)
            return 3

    k = rep["kpi"]
    print(f"\nГотово: {out}")
    print(f"Продажи: {sout}")
    if dout:
        print(f"Дебиторка: {dout}")
    print(f"Запас {k['stock_value']:,.0f} | заморожено {k['frozen']:,.0f} ({k['frozen_share']:.0%}) | "
          f"депозиты у заводов {k['deposits'] or 0:,.0f} | наш долг заводам {k['debts'] or 0:,.0f} | "
          f"срочно заказать {k['urgent_sku']} поз.".replace(",", " "))
    return 0


if __name__ == "__main__":
    sys.exit(main())
