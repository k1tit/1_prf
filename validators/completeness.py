# validators/completeness.py
"""Валидатор полноты данных"""

import pandas as pd
from .base_validator import BaseValidator


def _find_col(columns, hints):
    col_map = {str(c).strip().upper(): c for c in columns}
    for h in hints:
        hu = str(h).strip().upper()
        if hu in col_map:
            return col_map[hu]
    for c in columns:
        cu = str(c).strip().upper()
        for h in hints:
            hu = str(h).strip().upper()
            if hu in cu or (len(hu) >= 4 and hu in cu.replace("_", "").replace(" ", "")):
                return c
    return None


def _norm_join_key(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    s = str(v).strip()
    if not s or s.lower() in ("none", "null", "nan", "na"):
        return None
    if "." in s and s.replace(".", "", 1).isdigit():
        s = s.split(".", 1)[0]
    if s.isdigit() or (s.startswith("-") and s[1:].isdigit()):
        s = s.lstrip("0") or "0"
    return s


def _value_filled_mask(series):
    s = series.astype(str).str.strip()
    s_for_zero = s.str.replace(",", ".", regex=False)
    zeroish = s_for_zero.str.match(r"^-?0+(?:[.][0]+)?$", na=False)
    return (
        series.notna()
        & (s != "")
        & (~s.str.lower().isin(["none", "null", "nan", "na"]))
        & (~zeroish)
    )


def _adr2_rows_with_adr6_email(df_adr2, adr6_df):
    """True для строк ADR2, у которых есть непустой e-mail в ADR6 (по PERSNUMBER и/или ADDRNUMBER)."""
    if df_adr2 is None or df_adr2.empty or adr6_df is None or adr6_df.empty:
        return pd.Series(False, index=df_adr2.index if df_adr2 is not None else [])

    email_col = _find_col(
        adr6_df.columns,
        ["SMTP_ADDR", "SMTP_ADDRESS", "E-Mail Address", "E-Mail_Address", "email", "Email"],
    )
    if not email_col:
        return pd.Series(False, index=df_adr2.index)

    pers_adr2 = _find_col(df_adr2.columns, ["PERSNUMBER", "PERS_NUMBER", "Person"])
    addr_adr2 = _find_col(df_adr2.columns, ["ADDRNUMBER", "ADDR_NUMBER", "Addr. No."])
    pers_adr6 = _find_col(adr6_df.columns, ["PERSNUMBER", "PERS_NUMBER", "Person"])
    addr_adr6 = _find_col(adr6_df.columns, ["ADDRNUMBER", "ADDR_NUMBER", "Addr. No."])

    filled = _value_filled_mask(adr6_df[email_col])
    adr6_ok = adr6_df.loc[filled].copy()
    if adr6_ok.empty:
        return pd.Series(False, index=df_adr2.index)

    use_pers = bool(pers_adr2 and pers_adr6)
    use_addr = bool(addr_adr2 and addr_adr6)

    if use_pers and use_addr:
        adr6_ok["_pk"] = adr6_ok[pers_adr6].map(_norm_join_key)
        adr6_ok["_ak"] = adr6_ok[addr_adr6].map(_norm_join_key)
        valid = adr6_ok["_pk"].notna() & adr6_ok["_ak"].notna()
        keys = set(zip(adr6_ok.loc[valid, "_pk"], adr6_ok.loc[valid, "_ak"]))
        pk = df_adr2[pers_adr2].map(_norm_join_key)
        ak = df_adr2[addr_adr2].map(_norm_join_key)
        return pd.Series(
            [(p, a) in keys if pd.notna(p) and pd.notna(a) else False for p, a in zip(pk, ak)],
            index=df_adr2.index,
        )

    if use_pers:
        keys = set(adr6_ok[pers_adr6].map(_norm_join_key).dropna())
        return df_adr2[pers_adr2].map(_norm_join_key).isin(keys)

    if use_addr:
        keys = set(adr6_ok[addr_adr6].map(_norm_join_key).dropna())
        return df_adr2[addr_adr2].map(_norm_join_key).isin(keys)

    return pd.Series(False, index=df_adr2.index)


class CompletenessValidator(BaseValidator):
    """Проверка полноты данных"""

    def validate(self, df, column_name, **kwargs):
        """Проверка на отсутствие значений"""
        if column_name not in df.columns:
            return 0, 0, None  # правило не применимо — 0 оценённых строк

        rule_code = str(self.rule_info.get("rule_code", "")).strip().upper()
        treat_zero_as_missing_for = {
            "RCCOMP_375.1",
            "RCCOMP_375.1.2",
            "RCCOMP_372.1",
        }
        treat_zero_as_missing = rule_code in treat_zero_as_missing_for
        invert_filled_is_error = False

        s = df[column_name].astype(str).str.strip()
        s_for_zero = s.str.replace(",", ".", regex=False)
        zeroish = s_for_zero.str.match(r"^-?0+(?:[.][0]+)?$", na=False)

        empty_mask = (
            df[column_name].isna()
            | (s == "")
            | (s.str.lower().isin(["none", "null", "nan", "na"]))
        )
        if treat_zero_as_missing:
            empty_mask = empty_mask | zeroish
        if invert_filled_is_error:
            empty_mask = empty_mask | zeroish

        err_type = "COMPLETENESS"
        err_desc = f"Missing value in column {column_name}"

        if rule_code == "RCCOMP_375.1.2":
            adr6_df = kwargs.get("adr6_df")
            if adr6_df is not None and not getattr(adr6_df, "empty", True):
                has_tel = ~empty_mask
                has_email = _adr2_rows_with_adr6_email(df, adr6_df)
                error_mask = ~has_tel & ~has_email
                err_desc = (
                    f"Missing {column_name} and no E-mail Address in ADR6 "
                    f"(TEL_NUMBER or ADR6 e-mail required)"
                )
            else:
                error_mask = empty_mask
        elif invert_filled_is_error:
            error_mask = ~empty_mask
            err_type = "CONFORMITY"
            err_desc = (
                f"Account group 9038: {column_name} must be empty (NULL); "
                f"filled reconciliation account is not allowed"
            )
        else:
            error_mask = empty_mask

        error_count = int(error_mask.sum())
        total_rows = len(df)

        error_df = self._prepare_error_dataframe(df, error_mask, err_type, err_desc)
        if error_df is not None and rule_code == "RCCOMP_113.1":
            error_df = error_df.copy()
            if column_name in df.columns and "AKONT" not in error_df.columns:
                error_df["AKONT"] = df.loc[error_mask, column_name].values

            col_lower = {str(c).strip().lower(): c for c in df.columns}
            if "ktokd" in col_lower:
                error_df["KTOKD"] = df.loc[error_mask, col_lower["ktokd"]].values
            else:
                for name in (
                    "kna.ktokd",
                    "account_group_code",
                    "b.account_group_code",
                    "b.ktokd",
                    "group_1",
                ):
                    if name in col_lower:
                        error_df["KTOKD"] = df.loc[error_mask, col_lower[name]].values
                        break
            if "KTOKD_SOURCE" in col_lower:
                error_df["KTOKD_SOURCE"] = df.loc[error_mask, col_lower["ktokd_source"]].values
            elif "KTOKD" in error_df.columns:
                error_df["KTOKD_SOURCE"] = "KNA1"
            if "rule_scope" in col_lower:
                error_df["RULE_SCOPE"] = df.loc[error_mask, col_lower["rule_scope"]].values

        if error_df is not None and self.error_saver:
            self._save_errors_if_needed(error_df)

        return total_rows, error_count, error_df
