"""Валидатор для проверки равенства колонок (не должны быть равны)"""
import pandas as pd
from .base_validator import BaseValidator

class CrossColumnEqualityCheckValidator(BaseValidator):
    """
    Валидатор для правил типа "Check if Name 1 equals Name 2" или "не должны быть равны".
    equality_required=True (RCCONF_12.6.x): проверка РАВЕНСТВА — равные OK, ошибка когда РАЗНЫЕ.
    equality_required=False: колонки НЕ ДОЛЖНЫ быть равны — равные = ошибка.
    other_columns=(list): если задан — проверка «значение column1 не должно совпадать ни с одним из other_columns» (для правил TAXNUM: tax_1 = tax_0 OR tax_1 = tax_2 OR ...).
    """
    
    def validate(self, df, column1, second_column=None, equality_required=False, other_columns=None, **kwargs):
        """equality_required: True = поля должны быть равны (ошибка при различии); False = не должны быть равны (ошибка при равенстве).
        other_columns: список колонок — ошибка, если column1 равен любой из них (для правил «same value as other VAT/Tax number»)."""
        
        other_cols = (list(other_columns) if other_columns else None) or []
        if other_cols:
            # Режим «column1 не должен равняться ни одной из other_columns» (все tax_* кроме проверяемого)
            return self._validate_against_many(df, column1, other_cols, equality_required)
        
        if second_column is None:
            print(f"Колонка для сравнения не задана (second_column или other_columns)")
            return 0, 0, None
        
        if equality_required:
            print(f"Проверка равенства: {column1} == {second_column} (должны совпадать)")
        else:
            print(f"Проверка неравенства: {column1} != {second_column} (не должны быть равны)")
        
        # Проверяем наличие колонок
        if column1 not in df.columns:
            print(f"Колонка '{column1}' не найдена в таблице")
            print(f"Доступные колонки: {list(df.columns)[:10]}...")
            return 0, 0, None
            
        if second_column not in df.columns:
            print(f"Колонка '{second_column}' не найдена в таблице")
            print(f"Доступные колонки: {list(df.columns)[:10]}...")
            return 0, 0, None
        
        # Берем значения
        values1_raw = df[column1]
        values2_raw = df[second_column]
        
        # Преобразование для Categorical данных
        if pd.api.types.is_categorical_dtype(values1_raw):
            values1_clean = values1_raw.astype(str).fillna('').str.strip()
        else:
            values1_clean = values1_raw.fillna('').astype(str).str.strip()
            
        if pd.api.types.is_categorical_dtype(values2_raw):
            values2_clean = values2_raw.astype(str).fillna('').str.strip()
        else:
            values2_clean = values2_raw.fillna('').astype(str).str.strip()
        
        # Пустые значения считаем так же, как и в CrossColumnEqualityValidator:
        # '', 'nan', 'none', 'null', '-', '.' (и любые pd.NA/NaN).
        empty_tokens = {'', 'nan', 'none', 'null', '-', '.'}
        v1_lower = values1_clean.astype(str).str.strip().str.lower()
        v2_lower = values2_clean.astype(str).str.strip().str.lower()

        empty1_mask = v1_lower.isin(empty_tokens)
        empty2_mask = v2_lower.isin(empty_tokens)

        # Статистика
        empty1_count = int(empty1_mask.sum())
        empty2_count = int(empty2_mask.sum())
        both_filled = ~empty1_mask & ~empty2_mask
        both_filled_count = both_filled.sum()
        
        print(f"Статистика:")
        print(f"  Пустых в '{column1}': {empty1_count:,}")
        print(f"  Пустых в '{second_column}': {empty2_count:,}")
        print(f"  Обе колонки заполнены: {both_filled_count:,}")
        
        # ВАЖНО: Проверяем только строки, где ОБЕ колонки заполнены
        # Правило: по technical_definition IF X = Y THEN '0' ELSE '1'
        # Это означает: если они РАВНЫ → '0' = ОШИБКА
        #               если они НЕ РАВНЫ → '1' = НЕТ ОШИБКИ
        # Колонки НЕ ДОЛЖНЫ быть равны. Если они РАВНЫ - это ошибка.
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
                if equality_required:
                    print(f"    Строка {idx}: '{column1}'='{v1}' vs '{second_column}'='{v2}' → {'РАВНЫ (OK)' if eq else 'РАЗНЫЕ (ОШИБКА!)'}")
                else:
                    print(f"    Строка {idx}: '{column1}'='{v1}' vs '{second_column}'='{v2}' → {'РАВНЫ (ОШИБКА!)' if eq else 'РАЗНЫЕ (OK)'}")
        
        equal_values = values1_clean == values2_clean
        if equality_required:
            # RCCONF_12.6.x: проверка на равенство — ошибка когда РАЗНЫЕ
            errors_mask = both_filled & (~equal_values)
            unequal_count = errors_mask.sum()
            equal_count = (both_filled & equal_values).sum()
        else:
            # По умолчанию: ошибка когда РАВНЫ (не должны совпадать)
            errors_mask = both_filled & equal_values
            unequal_count = (both_filled & ~equal_values).sum()
            equal_count = errors_mask.sum()
        error_count = errors_mask.sum()
        
        if both_filled_count > 0:
            error_percent = (error_count / both_filled_count) * 100
            if equality_required:
                print(f"  Строк с равными значениями (OK): {equal_count:,}")
                print(f"  Строк с разными значениями (ОШИБКА): {error_count:,} ({error_percent:.1f}%)")
            else:
                print(f"  Строк с разными значениями (OK): {unequal_count:,}")
                print(f"  Строк с равными значениями (ОШИБКА): {error_count:,} ({error_percent:.1f}%)")
        
        error_df = None
        if error_count > 0:
            error_df = df[errors_mask].copy()

            # Чтобы в выгрузке было видно, какие именно колонки сравниваются
            error_df["DQ_COLUMN_CHECKED_1"] = column1
            error_df["DQ_COLUMN_CHECKED_2"] = second_column

            if equality_required:
                error_df['error_type'] = 'UNEQUAL_VALUES'
                error_df['error_message'] = (
                    f"Значения в колонках '{column1}' и '{second_column}' должны совпадать, но они различаются"
                )
            else:
                error_df['error_type'] = 'EQUAL_VALUES'
                error_df['error_message'] = (
                    f"Значения в колонках '{column1}' и '{second_column}' не должны быть равны, но они совпадают"
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
        
        total_rows = both_filled_count
        return total_rows, error_count, error_df

    def _validate_against_many(self, df, column1, other_cols, equality_required=False):
        """Проверка: значение в column1 не должно совпадать ни с одной из other_cols (правила TAXNUM: tax_X = tax_0 OR tax_X = tax_1 OR ...).
        Оценённые строки: где column1 и все other_cols заполнены (по правилу IF any NULL THEN '')."""
        if column1 not in df.columns:
            print(f"Колонка '{column1}' не найдена в таблице")
            return 0, 0, None
        missing = [c for c in other_cols if c not in df.columns]
        if missing:
            print(f"Колонки для сравнения не найдены: {missing[:5]}{'...' if len(missing) > 5 else ''}")
            return 0, 0, None
        
        values1 = df[column1].fillna('').astype(str).str.strip()
        filled_main = values1 != ''
        filled_others = pd.Series(True, index=df.index)
        for c in other_cols:
            filled_others = filled_others & (df[c].fillna('').astype(str).str.strip() != '')
        evaluated = filled_main & filled_others
        evaluated_count = evaluated.sum()
        total_rows = len(df)
        excluded = total_rows - evaluated_count
        print(f"Проверка: '{column1}' не должен совпадать ни с одной из {len(other_cols)} колонок")
        print(f"  Оценённых строк (все поля tax_* заполнены): {evaluated_count:,} из {total_rows:,}" + (f", исключено из-за NULL: {excluded:,}" if excluded else ""))
        
        if evaluated_count == 0:
            return 0, 0, None
        
        errors_mask = pd.Series(False, index=df.index)
        for c in other_cols:
            other_vals = df[c].fillna('').astype(str).str.strip()
            errors_mask = errors_mask | (evaluated & (values1 == other_vals))
        error_count = errors_mask.sum()
        
        if evaluated_count > 0:
            pct = (error_count / evaluated_count) * 100
            print(f"  Строк, где значение совпадает с другим tax (ОШИБКА): {error_count:,} ({pct:.1f}%)")
        
        error_df = None
        if error_count > 0:
            error_df = df[errors_mask].copy()
            error_df['error_type'] = 'EQUAL_TO_OTHER_TAX'
            error_df['error_message'] = (
                f"Значение в '{column1}' совпадает с одним из других полей tax (не должно повторяться)"
            )
            if self.error_saver:
                self._save_errors_if_needed(error_df)
        
        return int(evaluated_count), int(error_count), error_df


