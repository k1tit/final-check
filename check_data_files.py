# -*- coding: utf-8 -*-
"""Проверка xlsx в папках выгрузок — какой файл не читается (BadZipFile)."""
from __future__ import annotations

import sys
import zipfile
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from build_checks import BASE_DIR, load_runtime_paths_dict  # noqa: E402

PATTERNS = [
    ("Base", "*Base*.xlsx"),
    ("BP", "*BP*.xlsx"),
    ("PY", "*PY*.xlsx"),
    ("ZY", "*ZY*.xlsx"),
]

SO_LIST = ["3801", "3802", "3803", "3804", "3805", "3806"]


def _is_zip_xlsx(path: Path) -> tuple[bool, str]:
    if not path.is_file():
        return False, "не файл"
    if path.stat().st_size < 100:
        return False, f"слишком маленький ({path.stat().st_size} байт)"
    if path.name.startswith("~$"):
        return False, "lock-файл Excel (~$) — закройте книгу в Excel"
    try:
        with zipfile.ZipFile(path) as zf:
            names = zf.namelist()
            if not any(n.startswith("xl/") for n in names):
                return False, "не похож на xlsx (нет xl/ внутри)"
        return True, "OK"
    except zipfile.BadZipFile:
        return False, "BadZipFile — не настоящий/битый xlsx"
    except Exception as exc:
        return False, str(exc)


def main() -> int:
    paths = load_runtime_paths_dict()
    print(f"data_dir:  {paths['data_dir']}")
    print(f"base_dir:  {paths['base_dir']}")
    print(f"существует: {paths['base_dir'].exists()}\n")

    if not paths["base_dir"].exists():
        print("Каталог выгрузок не найден. Проверьте runtime_paths.json (data_dir).")
        return 1

    bad = 0
    for so in SO_LIST:
        folder = paths["base_dir"] / so
        if not folder.is_dir():
            print(f"--- SO {so}: папки нет ---")
            continue
        print(f"=== SO {so} ===")
        for label, pat in PATTERNS:
            files = sorted(
                [
                    p
                    for p in folder.glob(pat)
                    if not p.name.startswith("~$") and not p.name.startswith(".")
                ],
                key=lambda p: p.stat().st_mtime,
                reverse=True,
            )
            if not files:
                print(f"  {label}: не найден ({pat})")
                continue
            if len(files) > 1:
                print(f"  {label}: найдено {len(files)} файлов:")
            for f in files:
                ok, msg = _is_zip_xlsx(f)
                mark = "OK" if ok else "ОШИБКА"
                size_kb = f.stat().st_size / 1024
                print(f"    [{mark}] {f.name} ({size_kb:.1f} KB) — {msg}")
                if not ok:
                    bad += 1
                    print(f"           полный путь: {f.resolve()}")
        print()

    if bad:
        print(f"Итого проблемных файлов: {bad}")
        print("Исправление: перевыгрузить макросом, удалить лишние/битые xlsx, закрыть Excel.")
        return 1
    print("Все проверенные файлы читаются как xlsx.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
