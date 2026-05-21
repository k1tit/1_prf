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
        treat_zero_as_missing_for = {"RCCOMP_375.1", "RCCOMP_375.1.2", "RCCOMP_372.1"}
        treat_zero_as_missing = rule_code in treat_zero_as_missing_for
        
        s = df[column_name].astype(str).str.strip()
        # "битые нули" из Excel иногда приходят как 0 / 0.0 / 00 / -0.000
        # Делаем детект устойчивым к формату десятичного разделителя '.' vs ','
        # Примеры: '0', '00', '-0', '0.0', '0,0', '000,000', '0.0000'
        s_for_zero = s.str.replace(",", ".", regex=False)
        zeroish = s_for_zero.str.match(r"^-?0+(?:[.][0]+)?$", na=False)

        # Маска для пустых значений
        mask = (
            df[column_name].isna() |
            (s == "") |
            (s.str.lower().isin(["none", "null", "nan", "na"]))
        )

        # Для указанных правил "0"/"0.0" и т.п. должны считаться пустыми
        if treat_zero_as_missing:
            mask = mask | zeroish
        
        error_count = mask.sum()
        total_rows = len(df)
        
        # Сохраняем ошибки
        error_df = self._prepare_error_dataframe(
            df, mask, 
            'COMPLETENESS', 
            f'Missing value in column {column_name}'
        )
        
        # Сохраняем ошибки через error_saver если он есть
        if error_df is not None and self.error_saver:
            self._save_errors_if_needed(error_df)
        
        return total_rows, error_count, error_df