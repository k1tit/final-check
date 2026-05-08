# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import traceback
from datetime import datetime
from pathlib import Path
from typing import Optional

import numpy as np
import pandas as pd

# ── Paths ─────────────────────────────────────────────────────────────────────

def _project_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent

_ROOT = _project_root()
RUNTIME_PATHS_JSON = _ROOT / "runtime_paths.json"
BUILD_PROGRESS_JSON = _ROOT / "build_progress.json"
_D_BASE = Path(r"C:\Users\kitit\check\check_PF_BP-PY-ZY\data\1 Нулевые файлы выгрузки макроса + файл исключений")
_D_OUT  = Path(r"C:\Users\kitit\check\check_PF_BP-PY-ZY\data\result")
_D_EXC  = _ROOT / "data" / "Exception.xlsx"


def _load_runtime_paths() -> tuple[Path, Path, Path]:
    base, out, exc = _D_BASE, _D_OUT, _D_EXC
    try:
        if RUNTIME_PATHS_JSON.is_file():
            cfg = json.loads(RUNTIME_PATHS_JSON.read_text(encoding="utf-8"))
            if cfg.get("base_dir"):       base = Path(cfg["base_dir"])
            if cfg.get("output_dir"):     out  = Path(cfg["output_dir"])
            if cfg.get("exception_file"): exc  = Path(cfg["exception_file"])
    except Exception:
        pass
    if os.environ.get("REPORTS_BASE_DIR"):       base = Path(os.environ["REPORTS_BASE_DIR"])
    if os.environ.get("REPORTS_OUTPUT_DIR"):     out  = Path(os.environ["REPORTS_OUTPUT_DIR"])
    if os.environ.get("REPORTS_EXCEPTION_FILE"): exc  = Path(os.environ["REPORTS_EXCEPTION_FILE"])
    return base, out, exc


BASE_DIR, OUTPUT_DIR, WEB_EXCEPTION_FILE = _load_runtime_paths()

# ── Config ────────────────────────────────────────────────────────────────────

PAIRS = {
    "3801_3803": {"folders": ["3801", "3803"]},
    "3802_3804": {"folders": ["3802", "3804"]},
    "3805_3806": {"folders": ["3805", "3806"]},
}

NAME_BLACKLIST = (
    "sampling", "сэмплинг", "самплинг", "семплинг",
    "dummy", "дамми", "внутреннее", "внп", "bf_", "test", "тест",
)
BLOCKED_ORBLK = {"S", "PR"}
FAKE_INN      = ("1111111111", "2222222222")

# Columns for "Привязка Bill-to" sheet — q_3804_Joined_3_BillToByINN
BILL_TO_COLS = [
    "SO", "Customer", "Name", "Tax Number 1", "CGrp", "Grp4",
    "Search Term 2", "SDst", "A7", "OrBlk1", "OrBlk2",
    "Cust OrBlk Bill-to", "BP now", "Check Cust&BP",
    "BP SO", "BP Customer", "BP Name", "BP Tax Number 1", "BP CGrp",
    "BP OrBlk1", "BP OrBlk2", "BP status",
]

# ── Progress (web UI) ─────────────────────────────────────────────────────────

def report_build_progress(percent: int, message: str, *, job_index: int = 1, job_total: int = 1) -> None:
    payload = {"percent": max(0, min(100, int(percent))), "message": message,
               "job_index": job_index, "job_total": job_total}
    try:
        tmp = BUILD_PROGRESS_JSON.with_suffix(".json.tmp")
        tmp.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
        tmp.replace(BUILD_PROGRESS_JSON)
    except OSError:
        pass

# ── Normalisation helpers ─────────────────────────────────────────────────────

def _norm(v: object) -> str:
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return ""
    s = str(v).strip()
    return s[:-2] if s.endswith(".0") else s


def _norm_cust(v: object) -> str:
    s = _norm(v)
    if s.isdigit():
        return s.lstrip("0") or "0"
    return s


def _ns(s: pd.Series) -> pd.Series:
    """Normalise a Series: strip + remove trailing .0"""
    return s.fillna("").astype(str).str.strip().apply(lambda x: x[:-2] if x.endswith(".0") else x)


def _nc(s: pd.Series) -> pd.Series:
    """Normalise customer key Series: strip leading zeros from digit-only."""
    base = _ns(s)
    return base.where(~base.str.match(r"^\d+$"), base.str.lstrip("0").replace("", "0"))


def _col(df: pd.DataFrame, *candidates: str) -> str | None:
    lo = {c.casefold(): c for c in df.columns}
    for c in candidates:
        found = lo.get(c.casefold())
        if found is not None:
            return found
    return None


def _so_col(df: pd.DataFrame) -> str | None:
    return _col(df, "SOrg#", "SOrg.", "VKORG", "SO", "Sales Org.")


def _get_file(folder: Path, pattern: str) -> Path | None:
    if not folder.exists():
        return None
    files = list(folder.glob(pattern))
    return files[0] if files else None

# ── 1. Load & normalise Base ──────────────────────────────────────────────────

_CYR = re.compile(r"[А-Яа-яЁё]")


def _read_base(path: Path, folder: str) -> pd.DataFrame:
    df = pd.read_excel(path, dtype=str)
    # OrBlk.1 → OrBlk1 at load time so downstream code uses stable names
    df = df.rename(columns={"OrBlk.1": "OrBlk1"})
    so = _so_col(df)
    if so and so != "SO":
        df = df.rename(columns={so: "SO"})
    # Normalise Customer
    cust = _col(df, "Customer", "KUNNR")
    if cust and cust != "Customer":
        df = df.rename(columns={cust: "Customer"})
    if "Customer" in df.columns:
        df["Customer"] = _nc(df["Customer"])
    df["_folder"] = folder
    return df


def dedupe_base(df: pd.DataFrame) -> pd.DataFrame:
    """q_3804_Base_DelDup: GROUP BY Customer, prefer Cyrillic name (min alpha)."""
    if df.empty:
        return df
    rows = []
    for _, grp in df.groupby("Customer", sort=False, dropna=False):
        row = grp.iloc[0].to_dict()
        if "Name" in grp.columns:
            names = [s.strip() for s in grp["Name"].fillna("").astype(str) if s.strip()]
            cyr = [n for n in names if _CYR.search(n)]
            row["Name"] = min(cyr) if cyr else (min(names) if names else "")
        rows.append(row)
    result = pd.DataFrame(rows)
    for c in df.columns:
        if c not in result.columns:
            result[c] = ""
    return result[df.columns].reset_index(drop=True)

# ── 2. Load partner file (BP / PY / ZY) ──────────────────────────────────────

def _read_partner(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, dtype=str)
    kunnr = _col(df, "KUNNR", "Customer", "Sold-to")
    ktonr = _col(df, "KTONR", "BP", "PY", "ZY")
    if not kunnr or not ktonr:
        return pd.DataFrame(columns=["KUNNR", "KTONR"])
    out = df[[kunnr, ktonr]].rename(columns={kunnr: "KUNNR", ktonr: "KTONR"})
    out["KUNNR"] = _nc(out["KUNNR"])
    out["KTONR"] = _nc(out["KTONR"])
    return out.dropna(subset=["KUNNR"]).drop_duplicates(subset=["KUNNR"], keep="first")

# ── 3. Merge one partner (BP / PY / ZY) ──────────────────────────────────────

def merge_partner(base_lookup: pd.DataFrame, partner_df: pd.DataFrame | None, prefix: str) -> pd.DataFrame:
    """
    SQL equivalent:
      SELECT pf.KUNNR AS Customer, pf.KTONR AS {prefix},
             d.Name, d.[Tax Number 1], d.OrBlk AS {prefix} OrBlk1, d.OrBlk1 AS {prefix} OrBlk2
      FROM {prefix}_table pf LEFT JOIN q_Base_DelDup d ON pf.KTONR = d.Customer
    Returns lookup: Customer → {prefix}, {prefix} Name, Tax, OrBlk1, OrBlk2
    """
    # Build enrichment from base
    keep = {"Customer": prefix}
    extra = {}
    for src, dst in [("Name", f"{prefix} Name"), ("Tax Number 1", f"{prefix} Tax Number 1"),
                     ("OrBlk", f"{prefix} OrBlk1"), ("OrBlk1", f"{prefix} OrBlk2")]:
        if src in base_lookup.columns:
            extra[src] = dst
    bk = base_lookup[list(keep) + list(extra)].rename(columns={**keep, **extra})
    bk = bk.drop_duplicates(subset=[prefix], keep="first")

    if partner_df is None or partner_df.empty:
        return pd.DataFrame(columns=["Customer", prefix] + list(extra.values()))

    pf = partner_df[["KUNNR", "KTONR"]].rename(columns={"KUNNR": "Customer", "KTONR": prefix})
    merged = pf.merge(bk, on=prefix, how="left")
    return merged.drop_duplicates(subset=["Customer"], keep="first")


def merge_all_partners(base: pd.DataFrame, bp_df, py_df, zy_df) -> pd.DataFrame:
    """Enrich base with BP, PY, ZY columns."""
    result = base.copy()
    for prefix, pf_df in [("BP", bp_df), ("PY", py_df), ("ZY", zy_df)]:
        lk = merge_partner(base, pf_df, prefix)
        result = result.merge(lk, on="Customer", how="left")
        # Если клиент не в partner file → prefix = NaN → дефолт = Customer
        if prefix not in result.columns:
            result[prefix] = result["Customer"]
        else:
            no_partner = result[prefix].fillna("").astype(str).str.strip() == ""
            result[prefix] = result[prefix].where(~no_partner, result["Customer"])

        # is_self: строки где partner = Customer (явно или по дефолту)
        is_self = result[prefix].fillna("").astype(str).str.strip() == \
                  result["Customer"].fillna("").astype(str).str.strip()

        # Подтягиваем атрибуты партнёра из base. Если partner = Customer — берём из base напрямую.
        for src, dst in [("Name", f"{prefix} Name"), ("Tax Number 1", f"{prefix} Tax Number 1"),
                         ("OrBlk", f"{prefix} OrBlk1"), ("OrBlk1", f"{prefix} OrBlk2")]:
            if dst not in result.columns:
                result[dst] = ""
            result[dst] = result[dst].fillna("")
            if src in base.columns:
                # Для self-rows берём значение из base (это данные самого клиента = его партнёра)
                result.loc[is_self & (result[dst] == ""), dst] = \
                    result.loc[is_self & (result[dst] == ""), src].fillna("")
    return result



# ── 4. Load all data for a pair ───────────────────────────────────────────────

def load_data(folders: list[str]) -> tuple[pd.DataFrame, pd.DataFrame | None, pd.DataFrame | None, pd.DataFrame | None]:
    bases, bps, pys, zys = [], [], [], []
    for folder in folders:
        fp = BASE_DIR / folder
        f = _get_file(fp, "*Base*.xlsx")
        if f:
            bases.append(_read_base(f, folder))
        for lst, pat in [(bps, "*BP*.xlsx"), (pys, "*PY*.xlsx"), (zys, "*ZY*.xlsx")]:
            f = _get_file(fp, pat)
            if f:
                lst.append(_read_partner(f))
    if not bases:
        return pd.DataFrame(), None, None, None
    base = dedupe_base(pd.concat(bases, ignore_index=True))

    def _cat(parts):
        if not parts:
            return None
        df = pd.concat(parts, ignore_index=True)
        return df.drop_duplicates(subset=["KUNNR"], keep="first")

    return base, _cat(bps), _cat(pys), _cat(zys)

# ── 5. Exceptions ─────────────────────────────────────────────────────────────

def load_exception(base_dir: Path) -> pd.DataFrame:
    parts = []
    for f in base_dir.rglob("Exception.xlsx"):
        try:
            parts.append(pd.read_excel(f, dtype=str))
        except Exception:
            pass
    if WEB_EXCEPTION_FILE.exists():
        try:
            parts.append(pd.read_excel(WEB_EXCEPTION_FILE, dtype=str))
        except Exception:
            pass
    if not parts:
        return pd.DataFrame(columns=["SO", "Customer", "Comment OM"])
    df = pd.concat(parts, ignore_index=True)
    so_c = _col(df, "SO", "SOrg#", "SOrg.", "VKORG")
    cu_c = _col(df, "Customer", "KUNNR")
    if not so_c or not cu_c:
        return pd.DataFrame(columns=["SO", "Customer", "Comment OM"])
    co_c = _col(df, "Comment OM", "Comment", "Комментарий")
    cols: dict[str, str] = {so_c: "SO", cu_c: "Customer"}
    if co_c:
        cols[co_c] = "Comment OM"
    out = df.rename(columns=cols)[list(cols.values())].copy()
    if "Comment OM" not in out.columns:
        out["Comment OM"] = ""
    out["SO"] = _ns(out["SO"])
    out["Customer"] = _nc(out["Customer"])
    out = out[(out["SO"] != "") & (out["Customer"] != "")]
    return out.drop_duplicates(subset=["SO", "Customer"], keep="last").reset_index(drop=True)


def exception_keys(exc_df: pd.DataFrame) -> Set[Tuple[str, str]]:
    if exc_df.empty:
        return set()
    return {(_norm(r["SO"]), _norm_cust(r["Customer"])) for _, r in exc_df.iterrows()}

# ── 6. Filters (q_3804_Joined_1_Checks WHERE clause) ─────────────────────────

def apply_filters(df: pd.DataFrame, exc_keys: Set[Tuple[str, str]]) -> pd.DataFrame:
    if df.empty:
        return df
    # Exception exclusion
    if exc_keys:
        so_s = _ns(df.get("SO", pd.Series("", index=df.index)))
        cu_s = _nc(df["Customer"])
        keep = pd.Series([(_norm(s), _norm_cust(c)) not in exc_keys
                          for s, c in zip(so_s, cu_s)], index=df.index)
        df = df[keep].copy()
    if df.empty:
        return df

    name_s = df["Name"].fillna("").astype(str).str.lower() if "Name" in df.columns else pd.Series("", index=df.index)
    mask_name = ~name_s.apply(lambda n: any(b in n for b in NAME_BLACKLIST))

    cgrp_s = df["CGrp"].fillna("").astype(str).str.upper().str.strip() if "CGrp" in df.columns else pd.Series("", index=df.index)
    mask_cgrp = cgrp_s != "Z"

    orblk_s  = df["OrBlk"].fillna("").astype(str).str.upper().str.strip() if "OrBlk" in df.columns else pd.Series("", index=df.index)
    orblk1_s = df["OrBlk1"].fillna("").astype(str).str.upper().str.strip() if "OrBlk1" in df.columns else pd.Series("", index=df.index)
    mask_orblk = ~orblk_s.isin(BLOCKED_ORBLK) & ~orblk1_s.isin(BLOCKED_ORBLK)

    tax_s = df["Tax Number 1"].fillna("").astype(str).str.strip() if "Tax Number 1" in df.columns else pd.Series("", index=df.index)
    mask_inn = ~tax_s.apply(lambda t: any(f in t for f in FAKE_INN))

    return df[mask_name & mask_cgrp & mask_orblk & mask_inn].copy()

# ── 7. Checks (q_3804_Joined_1_Checks computed columns) ──────────────────────

def add_checks(df: pd.DataFrame) -> pd.DataFrame:
    """
    Adds: OrBlk1/OrBlk2 aliases, Cust OrBlk Bill-to, BP OrBlk Bill-to,
          Check BP-PY-ZY, Check Tax Number Cust&BP, Check Cust&BP,
          Comment MD Analyst
    In base: OrBlk = first OB (b.[OrBlk] AS OrBlk1 in SQL output)
             OrBlk1 = second OB (b.[OrBlk1] AS OrBlk2 in SQL output)
    """
    out = df.copy()

    def _s(col: str) -> pd.Series:
        return out[col].fillna("").astype(str).str.strip() if col in out.columns else pd.Series("", index=out.index)

    # Output column aliases: OrBlk→OrBlk1, OrBlk1→OrBlk2
    if "OrBlk" in out.columns and "OrBlk1" not in out.columns:
        out = out.rename(columns={"OrBlk": "OrBlk1"})
    elif "OrBlk" in out.columns and "OrBlk1" in out.columns:
        out = out.rename(columns={"OrBlk1": "OrBlk2", "OrBlk": "OrBlk1"})

    orblk1    = _s("OrBlk1").str.upper()
    bp_orblk1 = _s("BP OrBlk1").str.upper()
    bp        = _s("BP");   py   = _s("PY");   zy  = _s("ZY")
    cust      = _s("Customer")
    tax       = _s("Tax Number 1"); bp_tax = _s("BP Tax Number 1")
    bp_name   = _s("BP Name")

    out["Cust OrBlk Bill-to"] = np.where(orblk1 == "M", "Bill-to", "Ship-to")
    out["BP OrBlk Bill-to"]   = np.where(bp_orblk1 == "M", "Bill-to", "Ship-to")
    out["Check BP-PY-ZY"]              = np.where((bp == py) & (bp == zy), "TRUE", "FALSE")
    out["Check Tax Number Cust&BP"]    = np.where(tax == bp_tax, "TRUE", "FALSE")
    out["Check Cust&BP"]               = np.where(cust == bp,    "TRUE", "FALSE")

    # Comment MD Analyst — nested IIf from q_3804_Joined_1_Checks, strictly in order
    bp_mismatch = (bp != py) | (bp != zy)
    tax_mismatch = tax != bp_tax
    cust_ne_bp   = cust != bp

    comment = pd.Series("", index=out.index)
    comment = comment.mask(bp_mismatch & ~tax_mismatch, "Несоответствие BP-PY-ZY")
    comment = comment.mask(bp_mismatch & tax_mismatch,  "Несоответствие BP-PY-ZY; несоответствие ИНН SP и BP")
    comment = comment.mask(~bp_mismatch & tax_mismatch, "Несоответствие ИНН Cust и BP")
    comment = comment.mask(~bp_mismatch & ~tax_mismatch & (bp_name == ""), "BP помечен на удаление")
    comment = comment.mask(
        ~bp_mismatch & ~tax_mismatch & (bp_name != "") & (orblk1 == "M") & cust_ne_bp,
        "Bill-to прикреплён к другому Bill-to"
    )
    comment = comment.mask(
        ~bp_mismatch & ~tax_mismatch & (bp_name != "") & (orblk1 != "M") & (bp_orblk1 != "M") & cust_ne_bp,
        "Ship-to прикреплён к BP без OB M"
    )
    out["Comment MD Analyst"] = comment
    return out

# ── 8. Bill-to by INN (q_3804_Joined_3_BillToByINN) ─────────────────────────

def build_bill_to(df_checks: pd.DataFrame, base: pd.DataFrame) -> pd.DataFrame:
    """
    pool1: Ship-to WHERE Check Cust&BP=TRUE AND Comment=""
    pool2: base WHERE OrBlk="M"  (Bill-to candidates)
    JOIN on Tax Number 1
    BP status = "active"  (pool2 already filtered to valid Bill-to)
    """
    if df_checks.empty or base.empty:
        return pd.DataFrame()

    pool1 = df_checks[
        (df_checks["Cust OrBlk Bill-to"] == "Ship-to") &
        (df_checks["Check Cust&BP"] == "TRUE") &
        (df_checks["Comment MD Analyst"] == "")
    ].copy()
    if pool1.empty:
        return pd.DataFrame()

    # pool2: Bill-to records — OrBlk = M
    orblk_col = "OrBlk"  # base still has pre-rename "OrBlk" at this point
    if orblk_col not in base.columns:
        return pd.DataFrame()
    pool2 = base[base[orblk_col].fillna("").astype(str).str.upper().str.strip() == "M"].copy()
    if pool2.empty:
        return pd.DataFrame()

    so_c = _so_col(pool2)
    rename_p2: dict[str, str] = {"Customer": "BP Customer", "OrBlk": "BP OrBlk1"}
    if so_c:
        rename_p2[so_c] = "BP SO"
    for src, dst in [("Name", "BP Name"), ("Tax Number 1", "BP Tax Number 1"),
                     ("CGrp", "BP CGrp"), ("OrBlk1", "BP OrBlk2")]:
        if src in pool2.columns:
            rename_p2[src] = dst
    pool2 = pool2.rename(columns=rename_p2)

    pool1["_tax"] = _ns(pool1["Tax Number 1"]) if "Tax Number 1" in pool1.columns else ""
    pool2["_tax"] = _ns(pool2["BP Tax Number 1"]) if "BP Tax Number 1" in pool2.columns else ""

    p2_cols = ["_tax"] + [c for c in ["BP SO", "BP Customer", "BP Name", "BP Tax Number 1",
                                        "BP CGrp", "BP OrBlk1", "BP OrBlk2"] if c in pool2.columns]
    joined = pool1.merge(pool2[p2_cols], on="_tax", how="inner")
    joined = joined.drop(columns=["_tax"], errors="ignore")
    if "BP" in joined.columns:
        joined = joined.rename(columns={"BP": "BP now"})

    joined["BP status"] = "active"

    sort_cols = [c for c in ["Tax Number 1", "Customer", "BP Customer"] if c in joined.columns]
    if sort_cols:
        joined = joined.sort_values(sort_cols, kind="stable")
    dedup = [c for c in ["SO", "Customer", "Tax Number 1", "BP Customer"] if c in joined.columns]
    return joined.drop_duplicates(subset=dedup, keep="first").reset_index(drop=True)

# ── 9. Exception persistence ──────────────────────────────────────────────────

def collect_and_persist_global_exception(base_dir: Path, output_dir: Path) -> pd.DataFrame:
    exc = load_exception(base_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    if exc.empty:
        try:
            (output_dir / "00_Exception_собрано_нет_данных.txt").write_text(
                "Исключения не загружены\n", encoding="utf-8")
        except Exception:
            pass
    else:
        try:
            exc.to_excel(output_dir / "00_Exception_собрано_перед_отчётами.xlsx",
                         index=False, sheet_name="Exception")
        except Exception:
            pass
    return exc

# ── 10. Excel export ──────────────────────────────────────────────────────────

def _prep_bill_to_sheet(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for col in BILL_TO_COLS:
        if col not in out.columns:
            out[col] = ""
    return out[BILL_TO_COLS]


def save_excel(pair_name: str, errors_df: pd.DataFrame,
               bill_to_df: pd.DataFrame, exc_df: pd.DataFrame) -> None:
    pair_dir = OUTPUT_DIR / pair_name
    pair_dir.mkdir(parents=True, exist_ok=True)
    date_str  = datetime.now().strftime("%d.%m.%Y")
    out_file  = pair_dir / f"Check PF BP-PY-ZY {pair_name} {date_str} - Необходимый итоговый файл.xlsx"

    def _write(path: Path) -> None:
        with pd.ExcelWriter(path, engine="openpyxl") as w:
            # Sheet 1: Несоответствия
            if not errors_df.empty:
                errors_df.to_excel(w, sheet_name="Несоответствия", index=False)
            else:
                pd.DataFrame({"Сообщение": ["Нет данных"]}).to_excel(
                    w, sheet_name="Несоответствия", index=False)

            # Sheet 2: Привязка Bill-to (one sheet per SO)
            if not bill_to_df.empty:
                bt = _prep_bill_to_sheet(bill_to_df)
                if "SO" in bt.columns:
                    for so_val, grp in bt.groupby("SO", dropna=False):
                        sname = f"Привязка Bill-to {_norm(so_val) or 'ALL'}"[:31]
                        grp.to_excel(w, sheet_name=sname, index=False)
                else:
                    bt.to_excel(w, sheet_name="Привязка Bill-to по ИНН", index=False)
            else:
                pd.DataFrame({"Сообщение": ["Нет данных"]}).to_excel(
                    w, sheet_name="Привязка Bill-to по ИНН", index=False)

            # Sheet 3: Exception
            if not exc_df.empty:
                exc_df.to_excel(w, sheet_name="Exception", index=False)
            else:
                pd.DataFrame({"Сообщение": ["Нет записей-исключений"]}).to_excel(
                    w, sheet_name="Exception", index=False)

    try:
        _write(out_file)
    except PermissionError:
        ts = datetime.now().strftime("%H%M%S")
        _write(pair_dir / f"Check PF BP-PY-ZY {pair_name} {date_str} - Необходимый итоговый файл_{ts}.xlsx")

    if not errors_df.empty:
        errors_df.to_excel(pair_dir / f"{pair_name}_ErrorsOnly.xlsx", index=False)
    if not bill_to_df.empty:
        _prep_bill_to_sheet(bill_to_df).to_excel(pair_dir / f"{pair_name}_BillToByINN.xlsx", index=False)

# ── 11. Process one pair ──────────────────────────────────────────────────────

def process_pair(pair_name: str, folders: list[str], exc_df: pd.DataFrame,
                 *, progress_lo: int = 5, progress_hi: int = 95,
                 job_index: int = 1, job_total: int = 1) -> bool:

    def prog(step: float, msg: str) -> None:
        span = max(1, progress_hi - progress_lo)
        pct  = int(progress_lo + span * max(0.0, min(1.0, step)))
        report_build_progress(pct, f"[{job_index}/{job_total}] {msg}",
                              job_index=job_index, job_total=job_total)

    prog(0.02, f"«{pair_name}»: чтение файлов…")
    base, bp_df, py_df, zy_df = load_data(folders)
    if base.empty:
        print(f"[build_checks] Пропуск «{pair_name}»: нет *Base*.xlsx", flush=True)
        return False

    print(f"[build_checks] «{pair_name}»: base={len(base)} строк, "
          f"bp={len(bp_df) if bp_df is not None else 0}, "
          f"py={len(py_df) if py_df is not None else 0}, "
          f"zy={len(zy_df) if zy_df is not None else 0}", flush=True)

    prog(0.20, f"«{pair_name}»: merge BP/PY/ZY…")
    merged = merge_all_partners(base, bp_df, py_df, zy_df)

    prog(0.38, f"«{pair_name}»: фильтры…")
    exc_keys_set = exception_keys(exc_df)
    filtered = apply_filters(merged, exc_keys_set)
    print(f"[build_checks] «{pair_name}»: после фильтров {len(filtered)} строк", flush=True)

    if filtered.empty:
        save_excel(pair_name, pd.DataFrame(), pd.DataFrame(), exc_df)
        return True

    prog(0.55, f"«{pair_name}»: проверки MD…")
    checked = add_checks(filtered)

    errors_only = checked[checked["Comment MD Analyst"] != ""].copy()
    print(f"[build_checks] «{pair_name}»: несоответствий {len(errors_only)}", flush=True)

    prog(0.72, f"«{pair_name}»: Bill-to по ИНН…")
    bill_to = build_bill_to(checked, base)
    print(f"[build_checks] «{pair_name}»: Bill-to строк {len(bill_to)}", flush=True)

    prog(0.88, f"«{pair_name}»: сохранение Excel…")
    save_excel(pair_name, errors_only, bill_to, exc_df)
    prog(1.0,  f"«{pair_name}»: готово")
    return True

# ── 12. Job builder ───────────────────────────────────────────────────────────

def _all_so_folders() -> list[str]:
    known = ["3801", "3802", "3803", "3804", "3805", "3806"]
    return [f for f in known if (BASE_DIR / f).is_dir()] or known


def _parse_folders(raw: str) -> list[str]:
    if not raw:
        return []
    seen: set[str] = set()
    out = []
    for token in raw.split(","):
        for f in re.findall(r"\d{4}", token.strip()):
            if f not in seen:
                out.append(f); seen.add(f)
    return out


def build_jobs(mode: str, folders_csv: str = "") -> list[tuple[str, list[str]]]:
    if mode == "pairs":
        return [(name, cfg["folders"]) for name, cfg in PAIRS.items()]
    all_f = _all_so_folders()
    if mode == "single":
        return [(f, [f]) for f in all_f]
    sel = _parse_folders(folders_csv) or all_f
    if mode == "custom_single":
        return [(f, [f]) for f in sel]
    if mode == "custom_group":
        name = "_".join(sel)
        return [(f"custom_{name}", sel)]
    raise ValueError(f"Неизвестный режим: {mode}")

# ── 13. Main ──────────────────────────────────────────────────────────────────

def main() -> int:
    parser = argparse.ArgumentParser(description="Анализ PF BP-PY-ZY")
    parser.add_argument("--mode", choices=["pairs", "single", "custom_single", "custom_group"], default="pairs")
    parser.add_argument("--folders", default="")
    parser.add_argument("--skip-manual-exceptions", action="store_true")  # ignored, kept for web UI compat
    args = parser.parse_args()

    if not BASE_DIR.exists():
        print(f"[build_checks] Нет каталога: {BASE_DIR}", flush=True)
        return 1

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    report_build_progress(2, "Загрузка исключений…", job_index=0, job_total=1)
    exc_df = collect_and_persist_global_exception(BASE_DIR, OUTPUT_DIR)

    try:
        jobs = build_jobs(args.mode, args.folders)
    except ValueError as e:
        print(f"[build_checks] {e}", flush=True)
        return 1

    if not jobs:
        return 0

    nj = len(jobs)
    report_build_progress(5, f"Задач: {nj}", job_index=0, job_total=nj)
    produced = False
    try:
        for i, (job_name, job_folders) in enumerate(jobs):
            lo = 6 + int(88 * i / nj)
            hi = 6 + int(88 * (i + 1) / nj)
            report_build_progress(lo, f"Задача «{job_name}» ({i+1}/{nj})…",
                                  job_index=i+1, job_total=nj)
            if process_pair(job_name, job_folders, exc_df,
                            progress_lo=lo, progress_hi=hi,
                            job_index=i+1, job_total=nj):
                produced = True
    except Exception:
        traceback.print_exc()
        return 1

    if not produced:
        report_build_progress(95, "Ни один отчёт не сохранён", job_index=nj, job_total=max(1, nj))
        print("[build_checks] Ни один отчёт не сохранён — нет файлов Base.", flush=True)

    report_build_progress(100, "Готово", job_index=nj, job_total=max(1, nj))
    return 0


if __name__ == "__main__":
    sys.exit(main())

