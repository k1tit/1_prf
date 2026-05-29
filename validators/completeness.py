# validators/completeness.py
"""Валидатор полноты данных"""

import pandas as pd
from .base_validator import BaseValidator

class CompletenessValidator(BaseValidator):
    """Проверка полноты данных"""
    
    def validate(self, df, column_name, **kwargs):
        """Проверка на отсутствие значений"""
        if column_name not in df.columns:
            return 0, 0, None  # правило не применимо — 0 оценённых строк

        # Для некоторых правил значение "0" в данных является техническим "пусто"
        # и должно считаться ошибкой полноты.
        rule_code = str(self.rule_info.get("rule_code", "")).strip().upper()
        # Для некоторых правил "0" в данных технически означает пусто/не заполнено.
        # Важно: тут держим все rule_code, которые используют полноту (CompletenessValidator),
        # иначе часть нулей не попадёт в ошибки.
        treat_zero_as_missing_for = {
            "RCCOMP_375.1",
            "RCCOMP_375.1.2",
            "RCCOMP_372.1",
        }
        treat_zero_as_missing = rule_code in treat_zero_as_missing_for
        # RCCOMP_113.1: стандартная полнота (Missing Reconciliation account):
        # пустой AKONT = ошибка.
        invert_filled_is_error = False

        s = df[column_name].astype(str).str.strip()
        # "битые нули" из Excel иногда приходят как 0 / 0.0 / 00 / -0.000
        # Делаем детект устойчивым к формату десятичного разделителя '.' vs ','
        # Примеры: '0', '00', '-0', '0.0', '0,0', '000,000', '0.0000'
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
            # 0 / 0.0 для этих правил = «пусто» (нет счёта сверки) → не ошибка
            empty_mask = empty_mask | zeroish

        if invert_filled_is_error:
            error_mask = ~empty_mask
            err_type = "CONFORMITY"
            err_desc = (
                f"Account group 9038: {column_name} must be empty (NULL); "
                f"filled reconciliation account is not allowed"
            )
        else:
            error_mask = empty_mask
            err_type = "COMPLETENESS"
            err_desc = f"Missing value in column {column_name}"

        error_count = int(error_mask.sum())
        total_rows = len(df)

        error_df = self._prepare_error_dataframe(df, error_mask, err_type, err_desc)
        if error_df is not None and rule_code == "RCCOMP_113.1":
            error_df = error_df.copy()
            # Для отчёта по RCCOMP_113.1 всегда даём явную колонку AKONT,
            # даже если физическая колонка в выгрузке называется иначе (например Recon.acct).
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

        # Сохраняем ошибки через error_saver если он есть
        if error_df is not None and self.error_saver:
            self._save_errors_if_needed(error_df)
        
        return total_rows, error_count, error_df