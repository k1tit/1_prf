"""
Точечная диагностика JOIN KNB1 <-> KNA1 (правила RCCOMP_113.1 / 115.1).

Запуск из корня проекта:
  python debug_knb1_kna1_join.py
  python debug_knb1_kna1_join.py --db path/to/db_april.db
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from utils.column_map_resolver import apply_column_headers_for_rules, load_column_map
from utils.sap_account_keys import norm_sap_account_group
from utils.sqlite_safe import connect_sqlite, resolve_database_path


def norm_customer_partner_key(v) -> str:
    """Копия core/checker.py::_norm_customer_partner_key."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).replace("\ufeff", "").replace("\u00a0", " ").strip().strip("'").strip('"').strip()
    if s.lower() in {"", "none", "null", "nan", "<na>", "nat", "-", ".", "n/a", "na"}:
        return ""
    if re.fullmatch(r"\d+\.0+", s):
        s = s.split(".")[0]
    digits = re.sub(r"\D", "", s)
    if not digits:
        return ""
    return digits.zfill(10)


def non_empty(series: pd.Series) -> int:
    return int(series.apply(norm_customer_partner_key).ne("").sum())


def pick_best_kunnr(df: pd.DataFrame, table: str) -> str | None:
    tn = table.upper()
    order = (
        ("CUSTOMER", "KUNNR", "CL_", "CLIENT", "CUSTOMER_CODE", "KUNNR_KNB1")
        if tn == "KNB1"
        else ("KUNNR", "CUSTOMER", "CL_", "CLIENT", "CUSTOMER_CODE", "PARTNER")
    )
    upper = {str(c).strip().upper(): c for c in df.columns}
    candidates = []
    for name in order:
        if name in upper:
            candidates.append(upper[name])
    for c in df.columns:
        cu = str(c).strip().upper().replace(" ", "")
        if cu in ("CUSTOMER", "KUNNR", "CL_", "CLIENT") and c not in candidates:
            candidates.append(c)
    if not candidates:
        return None
    return max(candidates, key=lambda col: non_empty(df[col]))


def pick_ktokd_col(df: pd.DataFrame) -> str | None:
    best = None
    best_n = -1
    for c in df.columns:
        cl = str(c).strip().lower()
        if cl in ("group_1", "ktokd", "account_group_code"):
            n = non_empty(df[c]) if cl != "ktokd" else int(df[c].apply(norm_sap_account_group).ne("").sum())
            if cl == "ktokd":
                n = int(df[c].apply(norm_sap_account_group).ne("").sum())
            else:
                n = int(df[c].astype(str).str.strip().ne("").sum())
            if n > best_n:
                best, best_n = c, n
    return best


def sample_values(series: pd.Series, n: int = 8) -> list:
    raw = series.dropna().astype(str).str.strip()
    raw = raw[raw.ne("") & ~raw.str.lower().isin({"nan", "none", "null"})]
    return raw.head(n).tolist()


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--db", default=None, help="SQLite (по умолчанию config/database.json)")
    ap.add_argument("--sample", type=int, default=15, help="Сколько примеров несовпадений показать")
    args = ap.parse_args()

    db_path, _src = resolve_database_path(ROOT, args.db, must_exist=True)
    col_map = load_column_map(str(ROOT))

    print(f"DB: {db_path}\n")

    conn = connect_sqlite(db_path)
    knb1_raw = pd.read_sql_query('SELECT * FROM "KNB1"', conn)
    kna1_raw = pd.read_sql_query('SELECT * FROM "KNA1"', conn)
    conn.close()

    print(f"Сырые строки: KNB1={len(knb1_raw):,}, KNA1={len(kna1_raw):,}")
    print(f"KNB1 колонки (клиент): {[c for c in knb1_raw.columns if 'cust' in c.lower() or 'kunn' in c.lower() or c in ('Customer', 'Cl_')]}")
    print(f"KNA1 колонки (клиент): {[c for c in kna1_raw.columns if 'cust' in c.lower() or 'kunn' in c.lower() or c in ('Customer', 'Cl_', 'Group_1')]}\n")

    knb1 = apply_column_headers_for_rules(knb1_raw.copy(), "KNB1", col_map, str(ROOT), log_renames=True)
    kna1 = apply_column_headers_for_rules(kna1_raw.copy(), "KNA1", col_map, str(ROOT), log_renames=True)

    knb1_key_col = pick_best_kunnr(knb1, "KNB1")
    kna1_key_col = pick_best_kunnr(kna1, "KNA1")
    kna1_ktokd_col = pick_ktokd_col(kna1)

    print(f"\nВыбранные колонки (как в checker):")
    print(f"  KNB1 ключ JOIN: [{knb1_key_col}]  непустых ключей: {non_empty(knb1[knb1_key_col]):,}")
    print(f"  KNA1 ключ lookup: [{kna1_key_col}]  непустых: {non_empty(kna1[kna1_key_col]):,}")
    print(f"  KNA1 KTOKD: [{kna1_ktokd_col}]")

    for label, df, col in (
        ("KNB1", knb1, knb1_key_col),
        ("KNA1", kna1, kna1_key_col),
    ):
        if col and col in df.columns:
            print(f"\n  {label} примеры RAW [{col}]: {sample_values(df[col])}")
            norm = df[col].apply(norm_customer_partner_key)
            print(f"  {label} примеры NORM: {norm[norm.ne('')].head(8).tolist()}")

    knb1_jk = knb1[knb1_key_col].apply(norm_customer_partner_key)
    lookup = kna1[[kna1_key_col, kna1_ktokd_col]].copy()
    lookup["_join_key"] = lookup[kna1_key_col].apply(norm_customer_partner_key)
    lookup["KTOKD_norm"] = lookup[kna1_ktokd_col].apply(norm_sap_account_group)
    lookup = lookup.drop_duplicates(subset=["_join_key"], keep="first")

    knb1_nonempty = knb1_jk.ne("")
    kna1_keys = set(lookup["_join_key"].tolist())
    knb1_keys = set(knb1_jk[knb1_nonempty].tolist())

    matched = knb1_jk.isin(lookup["_join_key"])
    n_match = int(matched.sum())
    n_knb1_keys = len(knb1_keys)
    n_kna1_keys = len(kna1_keys)
    overlap = len(knb1_keys & kna1_keys)

    print("\n" + "=" * 60)
    print("РЕЗУЛЬТАТ JOIN (нормализованный ключ, zfill 10)")
    print("=" * 60)
    print(f"  Уникальных ключей KNB1 (непустых): {n_knb1_keys:,}")
    print(f"  Уникальных ключей KNA1 lookup:      {n_kna1_keys:,}")
    print(f"  Пересечение множеств:              {overlap:,}")
    print(f"  Строк KNB1 с совпадением в lookup: {n_match:,} / {len(knb1):,}")

    merged = knb1.copy()
    merged["_join_key"] = knb1_jk
    merged = merged.merge(
        lookup[["_join_key", "KTOKD_norm"]].rename(columns={"KTOKD_norm": "KTOKD"}),
        on="_join_key",
        how="left",
    )
    filled = int(merged["KTOKD"].fillna("").astype(str).str.strip().ne("").sum())
    n9038 = int((merged["KTOKD"] == "9038").sum())
    print(f"  KTOKD заполнен после merge:        {filled:,}")
    print(f"  KTOKD = 9038:                      {n9038:,}")

    # Альтернативные колонки на KNB1
    print("\n--- Сравнение всех кандидатов ключа на KNB1 ---")
    for c in knb1.columns:
        cu = str(c).strip().upper()
        if cu in ("CUSTOMER", "KUNNR", "CL_", "CLIENT", "CUSTOMER_CODE") or "KUNNR" in cu:
            jk = knb1[c].apply(norm_customer_partner_key)
            m = int(jk.isin(lookup["_join_key"]).sum())
            print(f"  [{c:20}] непустых={non_empty(knb1[c]):>10,}  совпало строк={m:>10,}")

    # Несовпадения: примеры
    miss_mask = knb1_nonempty & ~matched
    n_miss = int(miss_mask.sum())
    print(f"\n--- Примеры KNB1 без совпадения в KNA1 (первые {args.sample}) ---")
    if n_miss == 0 and n_match == 0:
        print("  (все ключи KNB1 пустые после нормализации?)")
    elif n_miss == 0:
        print("  Все непустые ключи KNB1 нашлись в KNA1.")
    else:
        sub = knb1.loc[miss_mask, [knb1_key_col]].head(args.sample).copy()
        sub["norm_key"] = sub[knb1_key_col].apply(norm_customer_partner_key)
        print(sub.to_string(index=False))
        print(f"\n  ... всего строк без match: {n_miss:,}")

    # Проверка: может KNB1.Customer — это не KUNNR, а внутренний ID?
    if "Customer" in knb1.columns and "KUNNR" in knb1.columns and knb1_key_col == "Customer":
        same = (knb1["Customer"].astype(str) == knb1["KUNNR"].astype(str)).sum()
        print(f"\n  Customer == KUNNR (сырые строки): {same:,} / {len(knb1):,}")

    # Длины ключей до zfill
    print("\n--- Распределение длины цифр в RAW (до zfill) ---")
    for label, col, df in (("KNB1", knb1_key_col, knb1), ("KNA1", kna1_key_col, kna1)):
        raw_digits = df[col].apply(
            lambda v: len(re.sub(r"\D", "", str(v).split(".")[0])) if pd.notna(v) and str(v).strip() else 0
        )
        print(f"  {label} [{col}]: {raw_digits.value_counts().head(8).to_dict()}")

    # KTOKD 9038 в справочнике KNA1
    n9038_kna1 = int((lookup["KTOKD_norm"] == "9038").sum())
    print(f"\n  В KNA1 lookup клиентов с KTOKD=9038: {n9038_kna1:,}")


if __name__ == "__main__":
    main()
