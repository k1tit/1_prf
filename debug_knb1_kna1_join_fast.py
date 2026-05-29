"""Быстрая диагностика JOIN — только колонки клиента и Group_1.

Типичная причина «ключ совпал 0»:
  column_map: Cl_ -> KUNNR на KNA1, но Cl_ в выгрузке = '400' (не номер клиента).
  Customer на KNA1 = 3800000001 — правильный ключ.
  Старый код брал KUNNR (из Cl_) при равной заполненности с Customer.
"""
import re
import sqlite3
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))
from utils.sqlite_safe import resolve_database_path


def norm_key(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).replace("\ufeff", "").strip().strip("'").strip('"')
    if s.lower() in {"", "nan", "none", "null"}:
        return ""
    if re.fullmatch(r"\d+\.0+", s):
        s = s.split(".")[0]
    d = re.sub(r"\D", "", s)
    return d.zfill(10) if d else ""


def ne(s):
    return int(s.apply(norm_key).ne("").sum())


db, _ = resolve_database_path(ROOT, must_exist=True)
conn = sqlite3.connect(db)

# какие колонки есть
for tbl in ("KNB1", "KNA1"):
    cols = pd.read_sql_query(f'PRAGMA table_info("{tbl}")', conn)["name"].tolist()
    pick = [c for c in cols if any(x in c.lower() for x in ("customer", "kunnr", "cl_", "group_1", "group"))]
    print(f"{tbl} relevant cols: {pick}")

knb1 = pd.read_sql_query('SELECT "Customer", "KUNNR" FROM "KNB1" LIMIT 500000', conn)
# full count
n_knb1 = conn.execute('SELECT COUNT(*) FROM "KNB1"').fetchone()[0]
kna1 = pd.read_sql_query('SELECT "Customer", "Cl_", "Group_1", "KUNNR" FROM "KNA1"', conn)
conn.close()

print(f"\nKNB1 total rows: {n_knb1:,} (sample loaded {len(knb1):,})")
print(f"KNA1 rows: {len(kna1):,}")

for c in knb1.columns:
    print(f"  KNB1[{c}] ne={ne(knb1[c]):,} samples={knb1[c].dropna().head(3).tolist()}")

for c in kna1.columns:
    print(f"  KNA1[{c}] ne={ne(kna1[c]):,} samples={kna1[c].dropna().head(3).tolist()}")

# lookup from KNA1: prefer Cl_ then Customer then KUNNR
for key_col in ("Cl_", "Customer", "KUNNR"):
    if key_col not in kna1.columns:
        continue
    lk = kna1[[key_col, "Group_1"]].copy()
    lk["jk"] = lk[key_col].apply(norm_key)
    lk = lk[lk["jk"].ne("")].drop_duplicates("jk", keep="first")
    keys = set(lk["jk"])
    for knb_col in ("Customer", "KUNNR"):
        if knb_col not in knb1.columns:
            continue
        jk = knb1[knb_col].apply(norm_key)
        m = int(jk.isin(keys).sum())
        print(f"  JOIN KNB1.{knb_col} -> KNA1.{key_col}: matched {m:,}/{len(knb1):,}  lookup keys={len(keys):,}")

# full KNB1 if sample was partial - use SQL only keys
print("\n--- Full table via chunked overlap estimate ---")
conn = sqlite3.connect(db)
kna1_keys = pd.read_sql_query('SELECT DISTINCT "Cl_" as k FROM "KNA1" WHERE "Cl_" IS NOT NULL AND TRIM("Cl_") != ""', conn)
kna1_keys["jk"] = kna1_keys["k"].apply(norm_key)
keyset = set(kna1_keys["jk"]) - {""}
print(f"KNA1 distinct Cl_ keys (norm): {len(keyset):,}")

# count matches in KNB1 Customer
cur = conn.execute('SELECT "Customer" FROM "KNB1"')
match = 0
total = 0
empty = 0
examples_miss = []
while True:
    rows = cur.fetchmany(50000)
    if not rows:
        break
    for (cust,) in rows:
        total += 1
        jk = norm_key(cust)
        if not jk:
            empty += 1
            continue
        if jk in keyset:
            match += 1
        elif len(examples_miss) < 10:
            examples_miss.append((cust, jk))
print(f"KNB1 Customer: total={total:,} match Cl_={match:,} empty_norm={empty:,}")
print(f"  miss examples (raw, norm): {examples_miss[:5]}")

# try KNA1 Customer as lookup key
kna1_cust = pd.read_sql_query(
    'SELECT DISTINCT "Customer" as k FROM "KNA1" WHERE "Customer" IS NOT NULL', conn
)
kna1_cust["jk"] = kna1_cust["k"].apply(norm_key)
keyset2 = set(kna1_cust["jk"]) - {""}
print(f"KNA1 distinct Customer keys (norm): {len(keyset2):,}")

cur = conn.execute('SELECT "Customer" FROM "KNB1"')
match2 = 0
total2 = 0
while True:
    rows = cur.fetchmany(50000)
    if not rows:
        break
    for (cust,) in rows:
        total2 += 1
        jk = norm_key(cust)
        if jk and jk in keyset2:
            match2 += 1
print(f"KNB1 Customer vs KNA1.Customer: match={match2:,}/{total2:,}")

conn.close()
