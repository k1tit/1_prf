"""
Полностью пустые строки таблицы (как незаполненные строки в Excel ниже данных).

Используется при проверке качества: такие строки не должны входить в total_rows и в валидаторы,
так как проверять нечего.
"""

from __future__ import annotations

import pandas as pd

# После strip / lower — как «нет данных» (совместимо с dedupe_and_clean_db)
_EMPTY_CELL_NORMALIZED = frozenset({"", "none", "nan", "null", "<na>"})


def cell_is_empty_value(v) -> bool:
    """True, если значение ячейки считаем пустым (NaN, пробелы, none/nan/null)."""
    if pd.isna(v):
        return True
    s = str(v).strip()
    if s == "":
        return True
    return s.lower() in _EMPTY_CELL_NORMALIZED


def is_row_empty(row: pd.Series) -> bool:
    """True, если во всех колонках строки — пустые значения."""
    for v in row:
        if not cell_is_empty_value(v):
            return False
    return True


def fully_empty_rows_mask(df: pd.DataFrame) -> pd.Series:
    """
    Векторно: True там, где строка полностью пустая (все ячейки пустые).
    Без df.apply — пригодно для больших таблиц.
    """
    if df.empty:
        return pd.Series(False, index=df.index)
    if df.shape[1] == 0:
        return pd.Series(False, index=df.index)

    row_has_non_empty = pd.Series(False, index=df.index)
    for col in df.columns:
        s = df[col]
        is_na = s.isna()
        st = s.astype(str).str.strip()
        sl = st.str.lower()
        cell_empty = is_na | sl.isin(_EMPTY_CELL_NORMALIZED)
        row_has_non_empty |= ~cell_empty
    return ~row_has_non_empty
