"""
Нормализация ключей SAP для сравнения с conf_*.json и JOIN.

AKONT / reconciliation_account в conf_recon_accounts.json — с ведущими нулями (10 знаков),
в SQLite/Excel часто без них (число или строка «178490000»).
"""
from __future__ import annotations

import re
from typing import Any

import pandas as pd

_EMPTY = {"", "none", "null", "nan", "<na>", "nat", "-", ".", "n/a", "na"}


def norm_sap_recon_account(value: Any, *, length: int = 10) -> str:
    """
    Счёт сверки (AKONT): только цифры, дополнение слева нулями до `length`.
    Пусто / 0 / только нули → ''.
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, (int, float)):
        if value == 0:
            return ""
        if float(value) == int(value):
            digits = str(int(value))
            if not digits or set(digits) == {"0"}:
                return ""
            return digits.zfill(length) if len(digits) <= length else digits

    s = str(value).replace("\ufeff", "").replace("\u00a0", " ").strip()
    s = s.strip("'").strip('"').strip()
    if s.lower() in _EMPTY or s in {"0", "0.0", "0.00"}:
        return ""
    if re.fullmatch(r"0+(\.0+)?", s):
        return ""
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    if re.fullmatch(r"\d+\.0+", s):
        s = s.split(".")[0]
    digits = re.sub(r"\D", "", s)
    if not digits or set(digits) == {"0"}:
        return ""
    if len(digits) <= length:
        return digits.zfill(length)
    return digits


def norm_sap_account_group(value: Any) -> str:
    """KTOKD / account_group_code: без ведущих нулей, только цифры/текст группы."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    s = str(value).replace("\ufeff", "").replace("\u00a0", " ").strip().strip("'").strip('"')
    if s.lower() in _EMPTY:
        return ""
    if re.fullmatch(r"\d+\.0+", s):
        s = s.split(".")[0]
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return s
