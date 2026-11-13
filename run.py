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

from msreport import analytics, demo, excel, manual
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


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--demo", action="store_true", help="выдуманные данные вместо МойСклад")
    ap.add_argument("--out", help="путь к файлу отчёта")
    args = ap.parse_args()

    cfg = tomllib.loads((ROOT / "config.toml").read_text(encoding="utf-8"))
    today = datetime.now(ZoneInfo(cfg["files"].get("timezone", "Asia/Almaty"))).replace(microsecond=0, tzinfo=None)

    if args.demo:
        raw, logistics = demo.demo_raw(today), demo.demo_logistics(today)
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
        logistics = manual.read_logistics(ROOT / cfg["files"]["logistics"])

    rep = analytics.build_report(raw, logistics, cfg, today)
    name = f"Отчет_закуп_логистика_{today:%Y-%m-%d_%H%M}{'_ДЕМО' if args.demo else ''}.xlsx"
    out = Path(args.out) if args.out else ROOT / cfg["files"]["reports_dir"] / name
    excel.write_report(rep, out)

    k = rep["kpi"]
    print(f"\nГотово: {out}")
    print(f"Запас {k['stock_value']:,.0f} | заморожено {k['frozen']:,.0f} ({k['frozen_share']:.0%}) | "
          f"депозиты у заводов {k['deposits'] or 0:,.0f} | наш долг заводам {k['debts'] or 0:,.0f} | "
          f"срочно заказать {k['urgent_sku']} поз.".replace(",", " "))
    return 0


if __name__ == "__main__":
    sys.exit(main())
