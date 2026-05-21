"""cross_column.py - ИСПРАВЛЕННАЯ ВЕРСИЯ"""

import pandas as pd
from .base_validator import BaseValidator


class CrossColumnEqualityValidator(BaseValidator):
    def validate(self, df, column1, second_column, flag_inequality_as_error=False, **kwargs):
        """Проверка равенства/неравенства двух колонок.

        По умолчанию (flag_inequality_as_error=False): две колонки НЕ ДОЛЖНЫ совпадать:
        - если ОБА заполнены и РАВНЫ → ОШИБКА;
        - если ОБА заполнены и РАЗНЫЕ → OK.

        Если flag_inequality_as_error=True (правило «должны быть равны»):
        - если ОБА заполнены и РАВНЫ → OK;
        - если ОБА заполнены и РАЗНЫЕ → ОШИБКА.
        """
        if flag_inequality_as_error:
            print(f"Проверка: {column1} == {second_column} (должны БЫТЬ ОДИНАКОВЫМИ)")
        else:
            print(f"Проверка: {column1} != {second_column} (должны БЫТЬ РАЗНЫМИ)")
        
        # Проверяем наличие колонок
        if column1 not in df.columns:
            print(f"Колонка '{column1}' не найдена в таблице")
            print(f"Доступные колонки: {list(df.columns)[:10]}...")
            return 0, 0, None
            
        if second_column not in df.columns:
            print(f"Колонка '{second_column}' не найдена в таблице")
            print(f"Доступные колонки: {list(df.columns)[:10]}...")
            return 0, 0, None
        
        # Берем значения - ОСОБЕННОСТЬ: для Categorical данных
        values1_raw = df[column1]
        values2_raw = df[second_column]
        
        # ПРЕОБРАЗОВАНИЕ ДЛЯ CATEGORICAL ДАННЫХ
        # Если колонка categorical, сначала преобразуем в строку, потом fillna
        if pd.api.types.is_categorical_dtype(values1_raw):
            values1_clean = values1_raw.astype(str).fillna('').str.strip()
        else:
            values1_clean = values1_raw.fillna('').astype(str).str.strip()
            
        if pd.api.types.is_categorical_dtype(values2_raw):
            values2_clean = values2_raw.astype(str).fillna('').str.strip()
        else:
            values2_clean = values2_raw.fillna('').astype(str).str.strip()
        
        # СТРОГАЯ ПРОВЕРКА ПУСТЫХ ЗНАЧЕНИЙ
        # Пустыми считаем: '', 'nan', 'none', 'null', '-', '.', а также pd.NA, None
        def is_empty(val):
            if pd.isna(val):
                return True
            val_str = str(val).strip().lower()
            return val_str in ('', 'nan', 'none', 'null', '-', '.')
        
        empty1_mask = values1_clean.apply(is_empty)
        empty2_mask = values2_clean.apply(is_empty)
        both_filled = ~empty1_mask & ~empty2_mask
        both_filled_count = both_filled.sum()
        
        # Статистика
        empty1_count = empty1_mask.sum()
        empty2_count = empty2_mask.sum()
        
        print(f"Статистика:")
        print(f"  Пустых в '{column1}': {empty1_count:,}")
        print(f"  Пустых в '{second_column}': {empty2_count:,}")
        print(f"  Обе колонки заполнены: {both_filled_count:,}")
        
        # ВАЖНО: Проверяем только строки, где ОБЕ колонки заполнены
        if both_filled_count == 0:
            print(f"  Нет строк с заполненными обеими колонками - правило не применимо")
            return 0, 0, None  # 0 оценённых строк (все скипнуты — пусто по правилу)
        
        # ДОПОЛНИТЕЛЬНАЯ ПРОВЕРКА: смотрим первые несколько значений для отладки
        if both_filled_count > 0:
            sample_df = df[both_filled].head(5)
            print(f"  [DEBUG] Примеры строк для проверки (первые 5):")
            for idx in sample_df.index:
                v1 = values1_clean.loc[idx]
                v2 = values2_clean.loc[idx]
                eq = (v1 == v2)
                print(f"    Строка {idx}: '{column1}'='{v1}' vs '{second_column}'='{v2}' → {'РАВНЫ (ОШИБКА!)' if eq else 'РАЗНЫЕ (OK)'}")
        
        # ОШИБКА: среди заполненных строк, где значения СОВПАДАЮТ
        # По technical_definition: IF X = Y THEN '0' (ошибка) ELSE '1' (нет ошибки)
        equal_values = (values1_clean == values2_clean)
        errors_mask = both_filled & equal_values  # РАВНЫ = ошибка!
        error_count = errors_mask.sum()
        
        # Также считаем, сколько строк НЕ равны (это правильно)
        not_equal_count = (both_filled & ~equal_values).sum()
        
        if both_filled_count > 0:
            error_percent = (error_count / both_filled_count) * 100
            print(f"  Строк с одинаковыми значениями (ОШИБКА): {error_count:,} ({error_percent:.1f}%)")
            print(f"  Строк с разными значениями (OK): {not_equal_count:,}")
        
        error_df = None
        if error_count > 0:
            error_df = df[errors_mask].copy()
            error_df['error_type'] = 'DUPLICATE_VALUES'
            error_df['error_message'] = (
                f"Значения в колонках '{column1}' и '{second_column}' совпадают, хотя \"cannot be the same\""
            )
            
            # Примеры ошибок
            print(f"\nПримеры ошибок (первые 3):")
            sample_errors = error_df.head(3)
            for idx, row in sample_errors.iterrows():
                print(f"  Строка {idx}:")
                print(f"    {column1}: '{row[column1]}'")
                print(f"    {second_column}: '{row[second_column]}'")
                print()
            
            if self.error_saver:
                self._save_errors_if_needed(error_df)
        else:
            print(f"Ошибок не найдено!")
        
        # total_rows = количество строк с результатом '1' или '0' (только строки, где обе колонки заполнены)
        total_rows = both_filled_count
        return total_rows, error_count, error_df


# ДОПОЛНИТЕЛЬНО: версия с маппингом для твоих правил
class CrossColumnEqualityValidatorWithMapping(BaseValidator):
    
    def __init__(self, error_saver=None):
        super().__init__(error_saver)
        # Простой маппинг для часто используемых полей
        self.column_mapping = {
            # Для KNB1
            'AKONT': 'Recon.acct',  # Reconciliation account
            'ZTERM': 'PayT',        # Terms of Payment
            'FDGRV': 'GrP',         # Cash Management Group
            
            # Для KNB5  
            'MAHNA': 'Cl.',         # Dunning procedure
            'MAHNS': 'Procedure',   # Dunning procedure (alternate)
            
            # Для BUT000 - уже правильные имена
            'NAME_ORG1': 'NAME_ORG1',
            'NAME_ORG2': 'NAME_ORG2',
            'NAME_ORG3': 'NAME_ORG3',
            'NAME_ORG4': 'NAME_ORG4',
            'MC_NAME2': 'MC_NAME2'
        }
    
    def _map_column(self, column_name):
        """Простой маппинг для часто используемых полей"""
        return self.column_mapping.get(column_name, column_name)
    
    def validate(self, df, column1, second_column, **kwargs):
        """Версия с маппингом"""
        
        # Применяем маппинг
        mapped_col1 = self._map_column(column1)
        mapped_col2 = self._map_column(second_column)
        
        print(f"Проверка: {mapped_col1} != {mapped_col2} (было: {column1} != {second_column})")
        
        # Используем базовую версию
        validator = CrossColumnEqualityValidator(self.error_saver)
        return validator.validate(df, mapped_col1, mapped_col2, **kwargs)


# Простая утилитарная функция
def check_columns_not_equal(df, col1, col2):
    """Проверяет, что две колонки не равны"""
    if col1 not in df.columns or col2 not in df.columns:
        return 0, None
    
    # Берем как есть, без сложной обработки
    mask_both_filled = df[col1].notna() & df[col2].notna()
    mask_equal = df[col1].astype(str) == df[col2].astype(str)
    
    errors = df[mask_both_filled & mask_equal]
    return len(errors), errors if len(errors) > 0 else None