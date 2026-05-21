"""Валидатор для проверки длины строк - для RCCONF_15.1"""

import pandas as pd
from .base_validator import BaseValidator

class UppercaseValidator(BaseValidator):
    """Проверка верхнего регистра - ИСПРАВЛЕННАЯ ВЕРСИЯ для RCCONF_12.4"""
    
    def validate(self, df, column_to_check, country_column="COUNTRY", excluded_countries=None, **kwargs):
        """Проверяет, что значение в верхнем регистре, с исключениями для определенных стран"""
        rule_code = self.rule_info.get('rule_code', '')
        
        if rule_code != "RCCONF_12.4":
            # Для других правил используем обычную проверку
            return self._simple_uppercase_check(df, column_to_check)
        
        print(f"Выполняем проверку для правила {rule_code}")
        
        if excluded_countries is None:
            excluded_countries = ['AM', 'BY', 'CZ', 'SK']
        
        # Проверяем существование колонок
        if column_to_check not in df.columns:
            return 0, 0, None
        
        if country_column not in df.columns:
            print(f"Колонка страны {country_column} не найдена, проверяем только верхний регистр")
            return self._simple_uppercase_check(df, column_to_check)
        
        print(f"Проверяем верхний регистр с исключениями для стран: {excluded_countries}")
        
        # Преобразуем в строки; код страны приводим к верхнему регистру для сравнения
        df_clean = df.copy()
        df_clean[column_to_check] = df_clean[column_to_check].astype(str).str.strip()
        df_clean[country_column] = df_clean[country_column].astype(str).str.strip().str.upper()
        excluded_upper = [c.upper() for c in excluded_countries]
        
        # Логика из technical_definition_RU:
        # 1. Если country_code IN ('AM','BY','CZ','SK') - пропускаем (не проверяем)
        # 2. Если organization_1_name IS NULL - пропускаем
        # 3. Иначе проверяем верхний регистр
        
        # Маска для строк, которые нужно проверять
        check_mask = (
            (df_clean[column_to_check] != "") &  # Не пустое
            (df_clean[column_to_check] != "NULL") &  # Не null
            (df_clean[column_to_check] != "null") &  # Не null
            (~df_clean[country_column].isin(excluded_upper))  # Страна не в исключениях
        )
        
        # Среди тех, что нужно проверять, ищем те, что не в верхнем регистре
        error_mask = check_mask & (df_clean[column_to_check] != df_clean[column_to_check].str.upper())
        
        error_count = error_mask.sum()
        # total_rows = количество строк с результатом '1' или '0' (только строки, которые были проверены)
        total_rows = check_mask.sum()
        
        if error_count > 0:
            print(f"Найдено {error_count} строк не в верхнем регистре")
        
        error_df = self._prepare_error_dataframe(
            df, error_mask,
            'CONFORMITY',
            f'Значение должно быть в верхнем регистре (кроме стран {excluded_countries})'
        )
        
        # Сохраняем ошибки через error_saver если он есть
        if error_df is not None and self.error_saver:
            self._save_errors_if_needed(error_df)
        
        return total_rows, error_count, error_df
    
    def _simple_uppercase_check(self, df, column_to_check):
        """Простая проверка верхнего регистра (для других правил)"""
        if column_to_check not in df.columns:
            return 0, 0, None
        
        df_clean = df.copy()
        df_clean[column_to_check] = df_clean[column_to_check].astype(str).str.strip()
        
        # Проверяем только непустые значения
        check_mask = (
            (df_clean[column_to_check] != "") &
            (df_clean[column_to_check] != "NULL")
        )
        error_mask = check_mask & (df_clean[column_to_check] != df_clean[column_to_check].str.upper())
        
        error_count = error_mask.sum()
        # total_rows = количество строк с результатом '1' или '0' (только строки, которые были проверены)
        total_rows = check_mask.sum()
        
        error_df = self._prepare_error_dataframe(
            df, error_mask,
            'CONFORMITY',
            f'Значение должно быть в верхнем регистре'
        )
        
        # Сохраняем ошибки через error_saver если он есть
        if error_df is not None and self.error_saver:
            self._save_errors_if_needed(error_df)
        
        return total_rows, error_count, error_df