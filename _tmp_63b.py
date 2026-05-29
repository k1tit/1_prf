import sqlite3
import pandas as pd
import utils.column_map_resolver as cmr
from utils.sqlite_safe import resolve_database_path
from utils.column_map_resolver import load_column_map, apply_column_headers_for_rules
from table_scripts.taxnum_handler import TaxNumHandler


class MC:
    rules_file = "json files/rules.json"
    column_map = load_column_map(".")


class MM:
    def get_table(self, n):
        return None


cmr._cache.clear()
conn = sqlite3.connect(resolve_database_path(".")[0])
df = pd.read_sql(
    "SELECT * FROM DFKKBPTAXNUM WHERE Tax_Number_Category = 'RU5'",
    conn,
)
conn.close()

df2 = apply_column_headers_for_rules(df, "DFKKBPTAXNUM5", MC.column_map, ".", log_renames=False)
print("cols", [c for c in df2.columns if "TAX" in c.upper() or "Tax" in c])
h = TaxNumHandler("DFKKBPTAXNUM5", df2, MM(), MC())
r = h.validate_rule(
    {
        "rule_code": "RCCONF_63.1",
        "rule_description": "x",
        "quality_category": "Conformity",
        "column_name_checked": "TAXNUM5",
    }
)
print("total", r["total_records"], "failed", r["failed"], "passed", r["passed"])
