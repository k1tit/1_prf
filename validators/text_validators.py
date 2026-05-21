# validators/text_validators.py
"""Текстовые валидаторы"""

import re
import pandas as pd
from .base_validator import BaseValidator

class SpecialCharactersValidator(BaseValidator):
    """Проверка на специальные символы"""
    
    def validate(self, df, column_name, special_characters_ref=None, **kwargs):
        if column_name not in df.columns:
            return 0, 0, None
        
        if special_characters_ref is None:
            special_characters_ref = [
                '!', '@', '#', '$', '%', '^', '&', '*', '(', ')', '_', '+', '=', '{', '}', 
                '[', ']', '|', '\\', ':', ';', '"', "'", '<', '>', '?', '/', '~', '`'
            ]
        
        def has_special_chars(text):
            if pd.isna(text) or text == "":
                return False
            text_str = str(text)
            return any(char in text_str for char in special_characters_ref)
        
        # Проверяем только непустые значения
        check_mask = df[column_name].notna() & (df[column_name].astype(str).str.strip() != "")
        error_mask = check_mask & df[column_name].apply(has_special_chars)
        
        error_count = error_mask.sum()
        # total_rows = количество строк с результатом '1' или '0' (только строки, которые были проверены)
        total_rows = check_mask.sum()
        
        error_df = self._prepare_error_dataframe(
            df, error_mask,
            'CONFORMITY',
            f'Special characters found in column {column_name}'
        )
        
        # Сохраняем ошибки через error_saver если он есть
        if error_df is not None and self.error_saver:
            self._save_errors_if_needed(error_df)
        
        return total_rows, error_count, error_df


class ConsecutiveSpacesValidator(BaseValidator):
    """Проверка на множественные пробелы"""
    
    def validate(self, df, column_name, **kwargs):
        if column_name not in df.columns:
            return 0, 0, None
        
        # Проверяем только непустые значения
        check_mask = df[column_name].notna() & (df[column_name].astype(str).str.strip() != "")
        error_mask = check_mask & df[column_name].astype(str).str.contains(r'\s{2,}', regex=True)
        
        error_count = error_mask.sum()
        # total_rows = количество строк с результатом '1' или '0' (только строки, которые были проверены)
        total_rows = check_mask.sum()
        
        error_df = self._prepare_error_dataframe(
            df, error_mask,
            'CONFORMITY',
            f'Consecutive spaces found in column {column_name}'
        )
        
        # Сохраняем ошибки через error_saver если он есть
        if error_df is not None and self.error_saver:
            self._save_errors_if_needed(error_df)
        
        return total_rows, error_count, error_df


class UppercaseValidator(BaseValidator):
    """Проверка на заглавные буквы"""
    
    def validate(self, df, column_name, country_column=None, excluded_countries=None, **kwargs):
        if column_name not in df.columns:
            return 0, 0, None
        
        # Проверяем только непустые значения
        check_mask = df[column_name].notna() & (df[column_name].astype(str).str.strip() != "")
        
        if country_column and country_column in df.columns and excluded_countries:
            excluded_mask = df[country_column].astype(str).isin(excluded_countries)
            check_mask = check_mask & ~excluded_mask
        
        error_mask = check_mask & (df[column_name].astype(str).str.strip() != df[column_name].astype(str).str.strip().str.upper())
        
        error_count = error_mask.sum()
        # total_rows = количество строк с результатом '1' или '0' (только строки, которые были проверены)
        total_rows = check_mask.sum()
        
        error_df = self._prepare_error_dataframe(
            df, error_mask,
            'CONFORMITY',
            f'Text should be in uppercase in column {column_name}'
        )
        
        # Сохраняем ошибки через error_saver если он есть
        if error_df is not None and self.error_saver:
            self._save_errors_if_needed(error_df)
        
        return total_rows, error_count, error_df