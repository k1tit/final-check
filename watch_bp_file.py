# -*- coding: utf-8 -*-
"""
Проверка: меняется ли 3804 BP.xlsx на диске (без запуска полного отчёта).
Запуск: python watch_bp_file.py
Потом в ДРУГОМ окне: python new_access_pf_checks.py
"""
from __future__ import annotations

import sys
import time
import zipfile
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from build_checks import load_runtime_paths_dict  # noqa: E402

SO = "3804"
NAME = "3804 BP.xlsx"


def ok(path: Path) -> bool:
    try:
        with zipfile.ZipFile(path):
            return True
    except Exception:
        return False


def main() -> None:
    base = load_runtime_paths_dict()["base_dir"]
    path = base / SO / NAME
    print(f"Слежу за: {path.resolve()}")
    print("Ctrl+C — выход\n")
    if not path.is_file():
        print("Файл не найден. Проверьте runtime_paths.json (data_dir).")
        return
    prev = None
    n = 0
    while True:
        st = path.stat()
        snap = (st.st_size, st.st_mtime, ok(path))
        if snap != prev:
            n += 1
            readable = "читается" if snap[2] else "НЕ xlsx / битый"
            print(f"#{n}  размер={snap[0]}  mtime={snap[1]:.0f}  {readable}")
            prev = snap
        time.sleep(1)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nСтоп.")
