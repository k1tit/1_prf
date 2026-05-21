import pandas as pd
import numpy as np
from utils.file_manager import ErrorFileManager


def _rcconf_15_1_effective_len(val):
    """
    Длина NAME_ORG3 как в сырых данных: ведущий пробел/таб НЕ отбрасываются
    (иначе строка из 35 символов «пробел + 34 знака» превращается в 34 и даёт ложные ошибки).
    Убирается только BOM U+FEFF в начале строки — артефакт экспорта.
    """
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return 0
    s = str(val)
    if s.startswith("\ufeff"):
        s = s[1:]
    return len(s)


class LogicalValidator:
    """Валидатор для логических бизнес-правил"""
    
    def __init__(self, rule_info, error_manager: ErrorFileManager):
        self.rule_info = rule_info
        self.error_manager = error_manager
        
    def validate(self, df, column_to_check, **kwargs):
        """
        Проверка логических правил
        """
        if df.empty:
            return 0, 0, None
        
        rule_code = self.rule_info.get('rule_code', '')
        total_rows = len(df)
        
        try:
            # RCCONF_15.1 - NAME_ORG4 можно использовать только если Name 3 заполнен до максимума
            if rule_code == 'RCCONF_15.1':
                return self._validate_rcconf_15_1(
                    df,
                    column_to_check,
                    org3_column_hint=kwargs.get('org3_column_resolved'),
                )
            
            # Другие логические правила можно добавить здесь
            
            # По умолчанию - возвращаем пустой результат
            return total_rows, 0, None
            
        except Exception as e:
            print(f"Ошибка в LogicalValidator для правила {rule_code}: {e}")
            return total_rows, 0, None
    
    def _validate_rcconf_15_1(self, df, org4_column, org3_column_hint=None):
        """
        RCCONF_15.1:
        IF organization_4_name IS NULL OR organization_3_name IS NULL THEN ''
        ELSE IF len(organization_3_name без BOM) < 34 AND organization_4_name IS NOT NULL THEN '0'
        (ведущий пробел в длину входит)
        ELSE '1'
        """
        total_rows = len(df)
        
        org3_column = None
        if org3_column_hint and org3_column_hint in df.columns:
            org3_column = org3_column_hint
        want_upper = 'NAME_ORG3'
        if org3_column is None:
            for col in df.columns:
                if col.upper() == want_upper or col.upper().replace(' ', '') == want_upper.replace('_', ''):
                    org3_column = col
                    break
        if org3_column is None:
            for col in df.columns:
                col_lower = col.lower()
                if any(name in col_lower for name in ['organization_3_name', 'name_org3', 'name3', 'name_3', 'mc_name3']):
                    org3_column = col
                    break
        
        if org3_column is None:
            print(f"      [!] Для правила RCCONF_15.1 не найдена колонка NAME_ORG3. Доступные: {list(df.columns)[:15]}...")
            return 0, 0, None  # правило не применимо — 0 оценённых строк
        
        org4_series = df[org4_column].astype(str).replace(['nan', 'None', 'null'], '', regex=True)
        org3_series = df[org3_column].astype(str).replace(['nan', 'None', 'null'], '', regex=True)
        
        def is_empty(value):
            if pd.isna(value):
                return True
            return str(value).strip() == ''
        
        error_real_indices = []
        evaluated_indices = []  # Индексы строк, которые получили результат '1' или '0'
        for pos, (org4_val, org3_val) in enumerate(zip(org4_series, org3_series)):
            if is_empty(org4_val) or is_empty(org3_val):
                continue  # По правилу это '' (skip), не учитываем в total
            evaluated_indices.append(df.index[pos])  # Эта строка получила результат '1' или '0'
            org3_len = _rcconf_15_1_effective_len(org3_val)
            if org3_len < 34:
                real_idx = df.index[pos]
                error_real_indices.append(real_idx)
        
        error_count = len(error_real_indices)
        total_rows = len(evaluated_indices)  # Количество строк с результатом '1' или '0'
        
        if error_count > 0:
            error_df = df.loc[error_real_indices].copy()
            # Страховка: в ошибки не должны попадать строки, где NAME_ORG3 уже >= 34,
            # или одна из колонок фактически пустая после strip.
            org3_len_series = error_df[org3_column].apply(
                lambda v: _rcconf_15_1_effective_len(v)
                if str(v).strip().lower() not in ('', 'none', 'null', 'nan')
                else 0
            )
            org4_non_empty = error_df[org4_column].astype(str).apply(lambda v: str(v).strip().lower() not in ('', 'none', 'null', 'nan'))
            org3_non_empty = error_df[org3_column].astype(str).apply(lambda v: str(v).strip().lower() not in ('', 'none', 'null', 'nan'))
            valid_error_mask = org4_non_empty & org3_non_empty & (org3_len_series < 34)
            if (~valid_error_mask).any():
                before_cnt = len(error_df)
                error_df = error_df[valid_error_mask].copy()
                error_count = len(error_df)
                print(f"      [FILTER] RCCONF_15.1: исключены строки, где NAME_ORG3 >= 34 или одно из полей пустое: {before_cnt} -> {error_count}")
            if error_df.empty:
                print(f"      [RCCONF_15.1] После финальной фильтрации ошибок не осталось")
                return total_rows, 0, None
            error_df['error_type'] = 'RCCONF_15.1'
            error_df['error_message'] = (
                f'{org4_column} заполнено, но длина {org3_column} < 34 символов '
                f'(по полному значению, ведущий пробел считается)'
            )
            print(
                f"      [RCCONF_15.1] Найдено ошибок: {error_count:,} "
                f"(NAME_ORG4 заполнено, длина NAME_ORG3 по сырому полю < 34)"
            )
        else:
            error_df = None
            print(f"      [RCCONF_15.1] Ошибок не найдено")
        
        return total_rows, error_count, error_df