#!/usr/bin/env python3
"""Передача отчётов второго аккаунта между окружениями через закрытую страницу Claude.

    python scripts/handoff.py pack   <папка_выхода> <продажи.xlsx> <дебиторка.xlsx> <письмо_продажи.html> <письмо_дебиторка.html>
    python scripts/handoff.py unpack <папка_со_скачанными_файлами> <папка_выхода>

Страница принимает только текстовые файлы, поэтому Excel кладётся в base64 (.b64.txt) и раскладывается обратно байт в байт.
"""
import base64
import json
import shutil
import sys
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent


def pack(out: Path, sales: Path, debts: Path, sales_mail: Path, debts_mail: Path) -> None:
    out.mkdir(parents=True, exist_ok=True)
    (out / "sales.xlsx.b64.txt").write_text(base64.b64encode(sales.read_bytes()).decode())
    (out / "debts.xlsx.b64.txt").write_text(base64.b64encode(debts.read_bytes()).decode())
    shutil.copy(sales_mail, out / "sales_mail.txt")
    shutil.copy(debts_mail, out / "debts_mail.txt")
    shutil.copy(HERE / "handoff_index.html", out / "index.html")
    (out / "meta.json").write_text(json.dumps({"built_at": datetime.now().astimezone().isoformat(timespec="seconds"),
                                               "sales_file": sales.name, "debts_file": debts.name}, ensure_ascii=False))
    print(f"Готово к публикации: {out}")


def unpack(src: Path, out: Path) -> None:
    meta = json.loads((src / "meta.json").read_text())
    out.mkdir(parents=True, exist_ok=True)
    for key, name in (("sales", meta["sales_file"]), ("debts", meta["debts_file"])):
        (out / name).write_bytes(base64.b64decode((src / f"{key}.xlsx.b64.txt").read_text()))
        print(out / name)
    print("built_at", meta["built_at"])


if __name__ == "__main__":
    cmd, *args = sys.argv[1:]
    if cmd == "pack":
        pack(*map(Path, args))
    elif cmd == "unpack":
        unpack(*map(Path, args))
    else:
        sys.exit(__doc__)
