import re
import sqlite3
import pandas as pd
from utils.column_map_resolver import load_column_map, apply_column_headers_for_rules, resolve_column_in_df

conn = sqlite3.connect("db_april.db")
cols = pd.read_sql_query('PRAGMA table_info("ADR2")', conn)["name"].tolist()
print("RAW cols (tel-related):", [c for c in cols if "tel" in c.lower() or "phone" in c.lower() or "r3" in c.lower() or "pers" in c.lower()])

df = pd.read_sql_query('SELECT * FROM "ADR2" LIMIT 20000', conn)
print("all cols:", cols)
conn.close()
cm = load_column_map(".")
mapped = apply_column_headers_for_rules(df, "ADR2", cm, ".", log_renames=False)
tel_col = resolve_column_in_df(mapped, "TEL_NUMBER", "ADR2", cm, ".")
print("TEL col:", tel_col)
if not tel_col:
    print("mapped cols:", list(mapped.columns)[:40])
else:
    s = mapped[tel_col].astype(str).str.strip()
    nonempty = s[~s.isin(["", "nan", "None", "null"])]
    only_digits = nonempty.str.fullmatch(r"[0-9]+", na=False)
    has_nondigit = nonempty.apply(lambda v: bool(re.search(r"[^0-9]", str(v))))
    print("nonempty", len(nonempty))
    print("only_digits", int(only_digits.sum()))
    print("has_non_digit", int(has_nondigit.sum()))
    print("samples fail:", nonempty[has_nondigit].head(15).tolist())
    print("samples ok:", nonempty[~has_nondigit].head(10).tolist())
