"""validators/dependency_validator.py"""

import pandas as pd
from .base_validator import BaseValidator

class DependencyValidator(BaseValidator):
    """Проверка зависимостей между колонками (например, NAME_ORG4 можно использовать только если NAME_ORG3 заполнен)"""
    
    def validate(self, df, column_to_check, dependent_column=None, **kwargs):
        """Проверяет зависимость между колонками"""
        if column_to_check not in df.columns:
            print(f"[!] Колонка {column_to_check} не найдена!")
            return 0, 0, None
        
        if dependent_column is None or dependent_column not in df.columns:
            print(f"[!] Зависимая колонка '{dependent_column}' не найдена!")
            return 0, 0, None
        
        print(f"🔍 Проверка зависимости: {column_to_check} зависит от {dependent_column}")
        
        # Преобразуем данные
        df_check = df[[column_to_check, dependent_column]].copy()
        
        # Очищаем значения
        df_check[column_to_check] = df_check[column_to_check].astype(str).str.strip()
        df_check[dependent_column] = df_check[dependent_column].astype(str).str.strip()
        
        # Заменяем пустые значения
        empty_values = ['', 'nan', 'NaN', 'NAN', 'null', 'NULL', 'None', 'none']
        df_check[column_to_check] = df_check[column_to_check].replace(empty_values, '')
        df_check[dependent_column] = df_check[dependent_column].replace(empty_values, '')
        
        # Определяем ошибки:
        # ОШИБКА: column_to_check заполнен, а dependent_column НЕ заполнен
        # Проверяем только строки, где column_to_check заполнен (они получают результат '1' или '0')
        check_mask = (df_check[column_to_check] != '')
        error_mask = (
            check_mask &  # Колонка заполнена
            (df_check[dependent_column] == '')   # А зависимая колонка пуста
        )
        
        # total_rows = количество строк с результатом '1' или '0' (только строки, где column_to_check заполнен)
        total_rows = check_mask.sum()
        error_count = error_mask.sum()
        
        print(f"[*] Статистика:")
        print(f"   Всего строк: {total_rows:,}")
        print(f"   {column_to_check} заполнено: {(df_check[column_to_check] != '').sum():,}")
        print(f"   {dependent_column} заполнено: {(df_check[dependent_column] != '').sum():,}")
        print(f"   Нарушений зависимости: {error_count:,}")
        
        if error_count > 0:
            print(f"   Примеры нарушений:")
            samples = df_check[error_mask].head(3)
            for idx, row in samples.iterrows():
                print(f"     - {column_to_check}: '{row[column_to_check][:30]}'")
                print(f"       {dependent_column}: '{row[dependent_column][:30]}'")
        
        # Сохраняем ошибки
        error_df = self._prepare_error_dataframe(
            df, error_mask,
            'CONFORMITY',
            f'{column_to_check} should only be used when {dependent_column} is filled'
        )
        
        # Сохраняем ошибки через error_saver если он есть
        if error_df is not None and self.error_saver:
            self._save_errors_if_needed(error_df)
        
        return total_rows, error_count, error_df