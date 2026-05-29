# validators/conformity.py
"""Валидатор соответствия данных"""

import pandas as pd
import re
import os
import json
from .base_validator import BaseValidator

_FULLWIDTH_DIGITS = str.maketrans("０１２３４５６７８９", "0123456789")


def _strict_digits_only_tel(val) -> bool:
    """RCCONF_38.3 / 39.3 / 39.3.2: только 0-9, без +, пробелов и прочих символов (REGEXP '[^0-9]')."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return True
    try:
        if isinstance(val, (int, float)) and not pd.isna(val) and float(val) == int(float(val)):
            s = str(int(float(val)))
            return bool(s) and s.isdigit()
    except (ValueError, TypeError, OverflowError):
        pass
    s = str(val).strip().replace("\ufeff", "").strip().translate(_FULLWIDTH_DIGITS)
    if not s or s.lower() in ("none", "null", "nan", "na"):
        return True
    if re.match(r"^\d+\.0+$", s):
        s = str(int(float(s)))
    return re.search(r"[^0-9]", s) is None


class ConformityValidator(BaseValidator):
    """Проверка соответствия данных справочнику"""
    
    def validate(self, df, column_name, allowed_values=None, technical_definition=None, rule_code=None, **kwargs):
        """Проверка на соответствие справочнику"""
        if column_name not in df.columns:
            return 0, 0, None  # правило не применимо — 0 оценённых строк
        
        # Fail-safe для RCCONF_113.1:
        # даже если правило попало в generic ConformityValidator, считаем по reference-логике
        # IF AKONT IS NULL OR account_group_code IS NULL THEN '' ELSE check pair in conf_recon_accounts.json
        effective_rule_code = str(rule_code or self.rule_info.get("rule_code", "")).strip().upper()

        # RCCONF_63.1: TAXNUM5 — только цифры, длина 8/9/10/12 (conf_tax_number_format.json)
        if effective_rule_code == "RCCONF_63.1":
            import os
            current_dir = os.path.dirname(os.path.abspath(__file__))
            project_root = os.path.dirname(current_dir)
            conf_path = os.path.join(project_root, "json files", "conf_tax_number_format.json")
            allowed_lengths = {8, 9, 10, 12}
            try:
                with open(conf_path, "r", encoding="utf-8") as f:
                    conf_list = json.load(f)
                for item in conf_list if isinstance(conf_list, list) else []:
                    if str(item.get("country_code", "")).upper() == "RU":
                        ln = item.get("length")
                        if ln is not None:
                            allowed_lengths.add(int(ln))
            except Exception as e:
                print(f"      [WARN] RCCONF_63.1: conf_tax_number_format: {e}")

            def _norm_tax(v):
                if v is None or (isinstance(v, float) and pd.isna(v)):
                    return ""
                if isinstance(v, (int, float)) and v == int(v):
                    return str(int(v))
                s = str(v).strip()
                if s.endswith(".0") and s[:-2].isdigit():
                    s = s[:-2]
                return s

            short_c = next((c for c in ("Tax_Number", "TAXNUM", column_name) if c in df.columns), None)
            long_c = next((c for c in ("Tax_Number_Long", "TAXNUM_LONG") if c in df.columns), None)
            if not short_c and not long_c:
                return 0, 0, None
            short_s = df[short_c].map(_norm_tax) if short_c else pd.Series("", index=df.index)
            long_s = df[long_c].map(_norm_tax) if long_c else pd.Series("", index=df.index)
            ser = short_s.where(short_s != "", long_s)
            non_empty = ser != ""
            total_rows = int(non_empty.sum())
            if total_rows == 0:
                return 0, 0, None
            ok = non_empty & ser.str.match(r"^\d+$", na=False) & ser.str.len().isin(allowed_lengths)
            error_mask = non_empty & ~ok
            error_count = int(error_mask.sum())
            if error_count == 0:
                return total_rows, 0, None
            error_df = self._prepare_error_dataframe(
                df,
                error_mask,
                "CONFORMITY",
                f"Invalid TAXNUM5 format (Tax_Number|Tax_Number_Long). RU lengths: {sorted(allowed_lengths)}",
            )
            return total_rows, error_count, error_df

        if effective_rule_code == "RCCONF_113.1":
            print("      [DEBUG] RCCONF_113.1 fail-safe in ConformityValidator is ACTIVE")
            from utils.sap_account_keys import norm_sap_account_group, norm_sap_recon_account

            # account_group_code может быть уже в df после JOIN
            account_group_col = None
            for c in df.columns:
                cu = str(c).strip().lower()
                if cu in ("account_group_code", "b.account_group_code", "ktokd"):
                    account_group_col = c
                    break
            if not account_group_col:
                # Без account_group_code правило невозможно оценить корректно -> SKIP, не ложные ошибки
                print("      [WARN] RCCONF_113.1 in ConformityValidator: account_group_code отсутствует, пропускаем правило")
                return 0, 0, None

            # Загружаем reference table conf_recon_accounts.json
            current_dir = os.path.dirname(os.path.abspath(__file__))
            project_root = os.path.dirname(current_dir)
            conf_path = os.path.join(project_root, "json files", "conf_recon_accounts.json")
            allowed_pairs = set()
            try:
                with open(conf_path, "r", encoding="utf-8") as f:
                    conf = json.load(f)
                rows = conf.get("conf_recon_accounts", []) if isinstance(conf, dict) else []
                for row in rows:
                    g = norm_sap_account_group(row.get("account_group_code"))
                    a = norm_sap_recon_account(row.get("reconciliation_account"))
                    if g and a:
                        allowed_pairs.add((g, a))
            except Exception as e:
                print(f"      [WARN] RCCONF_113.1 in ConformityValidator: не удалось загрузить conf_recon_accounts.json: {e}")
                return 0, 0, None

            if not allowed_pairs:
                return 0, 0, None

            recon_norm = df[column_name].apply(norm_sap_recon_account)
            group_norm = df[account_group_col].apply(norm_sap_account_group)
            evaluated_mask = (recon_norm != "") & (group_norm != "")
            total_rows = int(evaluated_mask.sum())
            if total_rows == 0:
                return 0, 0, None

            eval_idx = df.index[evaluated_mask]
            pair_keys = pd.Series(
                list(zip(group_norm.loc[eval_idx], recon_norm.loc[eval_idx])),
                index=eval_idx,
            )
            exists_mask = pd.Series(False, index=df.index)
            exists_mask.loc[eval_idx] = pair_keys.isin(allowed_pairs)

            error_mask = evaluated_mask & (~exists_mask)
            error_count = int(error_mask.sum())
            if error_count == 0:
                return total_rows, 0, None

            error_df = df[error_mask].copy()
            error_df["DQ_ERROR_TYPE"] = "INVALID_COMBINATION"
            error_df["DQ_RULE_CODE"] = "RCCONF_113.1"
            error_df["DQ_RULE_DESCRIPTION"] = self.rule_info.get("rule_description", "")
            error_df["DQ_COLUMN_CHECKED"] = column_name
            error_df["DQ_ERROR_DESCRIPTION"] = "Invalid combination of account_group_code and reconciliation_account"
            error_df["DQ_TIMESTAMP"] = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
            return total_rows, error_count, error_df

        # RCCONF_383.1 / RCCONF_384.1:
        # Формат геокоординат должен быть строго (-)x.xxxxxx / (-)xx.xxxxxx / (-)xxx.xxxxxx
        # (десятичный разделитель — точка, ровно 6 знаков после точки).
        # Скип:
        #   - NULL/пусто
        #   - TO_INTEGER(value) = 0
        #   - account_group_code LIKE '7%' (если колонка account_group_code доступна в df)
        if effective_rule_code in {"RCCONF_383.1", "RCCONF_384.1"}:
            print(f"      [DEBUG] {effective_rule_code}: strict geo-format validator active (dot + 6 decimals)")
            # Находим account_group_code, если он есть в срезе данных
            account_group_col = None
            for c in df.columns:
                cu = str(c).strip().lower()
                if cu in ("account_group_code", "b.account_group_code", "ktokd", "b.ktokd", "kna.ktokd", "kna.KTOKD".lower()):
                    account_group_col = c
                    break

            s = df[column_name].astype(str).str.strip()
            is_null_like = (
                df[column_name].isna()
                | (s == "")
                | (s.str.lower().isin(["none", "null", "nan", "na"]))
            )
            # TO_INTEGER(value)=0: берём целую часть до точки/запятой, убираем знак
            integer_part = (
                s.str.replace(",", ".", regex=False)
                 .str.extract(r"^\s*([+-]?\d+)", expand=False)
                 .fillna("")
            )
            integer_zero = integer_part.str.lstrip("+-").str.lstrip("0").eq("")
            zero_skip = integer_zero & (~is_null_like)

            account_group_skip = pd.Series(False, index=df.index)
            if account_group_col is not None:
                ag = df[account_group_col].astype(str).str.strip()
                account_group_skip = ag.str.startswith("7")

            skip_mask = is_null_like | zero_skip | account_group_skip
            evaluated_mask = ~skip_mask
            total_rows = int(evaluated_mask.sum())
            if total_rows == 0:
                return 0, 0, None

            # Строгий формат: только точка как разделитель, ровно 6 цифр после точки.
            # Запятая невалидна.
            fmt_ok = s.str.match(r"^-?\d{1,3}\.\d{6}$", na=False)
            error_mask = evaluated_mask & (~fmt_ok)
            error_count = int(error_mask.sum())
            print(f"      [DEBUG] {effective_rule_code}: evaluated={total_rows:,}, errors={error_count:,}")
            if error_count == 0:
                return total_rows, 0, None

            error_df = self._prepare_error_dataframe(
                df,
                error_mask,
                "CONFORMITY",
                f"Invalid coordinate format in {column_name}. Expected (-)x.xxxxxx / (-)xx.xxxxxx / (-)xxx.xxxxxx with dot as decimal separator."
            )
            return total_rows, error_count, error_df
        
        total_rows = len(df)
        
        # RCCONF_38.3 / 39.3 / 39.3.2: только цифры 0-9 (REGEXP '[^0-9]').
        # Scope (PERSNUMBER, R3_USER) задаётся в checker до вызова валидатора.
        if technical_definition and rule_code in ["RCCONF_38.3", "RCCONF_39.3", "RCCONF_39.3.2"]:
            if rule_code == "RCCONF_38.3":
                r3_user_col = None
                for col in df.columns:
                    col_lower = col.lower()
                    if col_lower == 'r3_user' or col_lower == 'r3user' or 'r3_user' in col_lower or 'r3user' in col_lower:
                        r3_user_col = col
                        break
                if r3_user_col:
                    base_mask = df[r3_user_col].astype(str).str.strip() == '1'
                    if not base_mask.any():
                        return 0, 0, None
                    df_filtered = df[base_mask].copy()
                else:
                    print(f"      [WARN] R3_USER не найден в валидаторе для правила {rule_code}")
                    df_filtered = df.copy()
            else:
                df_filtered = df.copy()
            
            # Создаем маску для ошибок (изначально все False)
            error_mask = pd.Series([False] * len(df_filtered), index=df_filtered.index)
            
            # Пропускаем NULL/пустые значения (они не являются ошибками согласно правилу)
            null_mask = df_filtered[column_name].isna() | (df_filtered[column_name].astype(str).str.strip() == '')
            
            # Для не-NULL значений проверяем, что только цифры
            non_null_mask = ~null_mask
            
            if non_null_mask.any():
                for idx in df_filtered[non_null_mask].index:
                    tel_value = df_filtered.loc[idx, column_name]
                    if not _strict_digits_only_tel(tel_value):
                        error_mask.loc[idx] = True
            
            error_count = error_mask.sum()
            if rule_code == "RCCONF_39.3.2":
                error_description = (
                    f"Invalid telephone number format in {column_name}. "
                    "Must contain only digits 0-9 (no +, spaces, or other characters). "
                    "Empty values are allowed. (Only rows with filled PERSNUMBER.)"
                )
            elif rule_code == "RCCONF_39.3":
                error_description = (
                    f"Invalid telephone number format in {column_name}. "
                    "Must contain only digits 0-9 (no +, spaces, or other characters). "
                    "Empty values are allowed. (Only rows with empty PERSNUMBER.)"
                )
            else:
                error_description = (
                    f"Invalid telephone number format in {column_name}. "
                    "Must contain only digits 0-9 (no +, spaces, or other characters). "
                    "Empty values are allowed. Only fixed phones (R3_USER=1) are checked."
                )
            
            if error_count > 0:
                error_df = self._prepare_error_dataframe(df_filtered, error_mask, 'CONFORMITY', error_description)
            else:
                error_df = None
            
            # total_rows = количество строк с результатом '1' или '0' (не-NULL значения, которые были проверены)
            total_rows = non_null_mask.sum() if non_null_mask.any() else 0
            return total_rows, error_count, error_df
        
        # Если есть technical_definition с проверкой формата телефона (REGEXP '[^0-9]')
        elif technical_definition and rule_code == "RCCONF_38.5":
            # Проверка формата телефона: только цифры
            # Согласно technical_definition_RU: "IF contact_medium_value REGEXP '[^0-9]' THEN '0'"
            # Это означает, что если есть не-цифры, то ошибка
            
            # Создаем маску для ошибок (изначально все False)
            error_mask = pd.Series([False] * len(df), index=df.index)
            
            # Пропускаем NULL значения (они не являются ошибками)
            null_mask = df[column_name].isna()
            
            # Для не-NULL значений проверяем формат
            non_null_mask = ~null_mask
            
            if non_null_mask.any():
                # Проходим по всем не-NULL значениям
                for idx in df[non_null_mask].index:
                    phone_value = str(df.loc[idx, column_name]).strip()
                    
                    # Пропускаем пустые значения (они не являются ошибками для этого правила)
                    if not phone_value or phone_value.lower() in ['none', 'null', 'nan', '']:
                        continue
                    
                    # Проверяем, что только цифры (REGEXP '[^0-9]' означает наличие не-цифр)
                    if re.search(r'[^0-9]', phone_value):
                        error_mask.loc[idx] = True  # Ошибка: есть не-цифры
                        continue
                    
                    # Дополнительная проверка формата согласно technical_definition_RU:
                    # - Если начинается с '9' и длина = 9, то OK
                    # - Если начинается с '9' и длина = 10, то OK (добавлено по требованию)
                    # - Если начинается с '8', длина = 11 и второй символ == '9' (формат "89...."), то OK
                    # - Иначе ошибка (если не пустое)
                    
                    is_valid_format = False
                    
                    # Проверяем формат: начинается с '9' и длина = 9
                    if phone_value.startswith('9') and len(phone_value) == 9:
                        is_valid_format = True
                    
                    # Проверяем формат: начинается с '9' и длина = 10 (добавлено)
                    elif phone_value.startswith('9') and len(phone_value) == 10:
                        is_valid_format = True
                    
                    # Проверяем формат: начинается с '8', длина = 11 (ровно 11, не больше и не меньше) и второй символ == '9' (формат "89....")
                    elif phone_value.startswith('8') and len(phone_value) == 11 and len(phone_value) > 1 and phone_value[1] == '9':
                        is_valid_format = True
                    
                    # Если не соответствует ни одному формату - ошибка
                    if not is_valid_format:
                        error_mask.loc[idx] = True
            
            error_count = error_mask.sum()
            error_description = f'Invalid telephone format in {column_name}. Must contain only digits and match format (9 or 10 digits starting with 9, or 11 digits starting with 89)'
            
            # Создаем error_df для правила 38.5
            if error_count > 0:
                error_df = self._prepare_error_dataframe(df, error_mask, 'CONFORMITY', error_description)
            else:
                error_df = None
            
            # Сохраняем ошибки через error_saver если он есть
            if error_df is not None and self.error_saver:
                self._save_errors_if_needed(error_df)
            
            # total_rows = количество строк с результатом '1' или '0' (не-NULL значения, которые были проверены)
            total_rows = non_null_mask.sum() if non_null_mask.any() else 0
            return total_rows, error_count, error_df
        
        # Специальная обработка для правила RCCONF_39.5 / RCCONF_39.5.2: проверка формата телефона
        # Изоляция от RCCOMP_375.1: проверяем ТОЛЬКО заполненные TEL_NUMBER. Пустые не считаются ошибкой (их ловит RCCOMP_375.1).
        # В technical_definition_RU "IF TEL_NUMBER IS NULL THEN ''" — в коде это трактуется как SKIP (запись не проверяется, не входит в total_rows).
        # В checker перед вызовом валидатора для 39.5/39.5.2 строки с пустым TEL_NUMBER уже отфильтрованы, сюда приходят только заполненные.
        # Нормализация: из значения берём только цифры, затем проверяем формат.
        # RCCONF_39.5: 10 цифр с 9; 11 цифр 89…; 11 цифр 79… (российский +7). Допускаем скобки, дефисы, пробелы, ведущий +.
        # RCCONF_39.5.2: 10 цифр с 9; 11 цифр 8x (вторая ≠ 9); 11 цифр 79….
        elif technical_definition and rule_code in ["RCCONF_39.5", "RCCONF_39.5.2"]:
            # Считаем только строки с непустым TEL_NUMBER; пустые не входят ни в total_rows, ни в ошибки
            null_mask = df[column_name].isna() | (df[column_name].astype(str).str.strip() == '')
            empty_vals = df[column_name].astype(str).str.strip().str.lower().isin(['none', 'null', 'nan', 'na'])
            null_mask = null_mask | empty_vals
            non_null_mask = ~null_mask
            _fullwidth = str.maketrans('０１２３４５６７８９', '0123456789')

            def _normalize_raw(val):
                """Убираем BOM, обрезаем пробелы по краям, убираем ВСЕ пробельные символы (в т.ч. \\t, \\n, \\xa0)."""
                if val is None or (isinstance(val, float) and pd.isna(val)):
                    return ''
                s = str(val).strip().replace('\ufeff', '').strip()
                s = re.sub(r'\s+', '', s)  # все пробелы (space, tab, newline, nbsp и т.д.)
                return s.translate(_fullwidth)

            def _to_digits(val):
                if val is None or (isinstance(val, float) and pd.isna(val)):
                    return ''
                s = str(val).strip().replace('\ufeff', '').strip()
                s = re.sub(r'\s+', '', s)  # убираем все пробельные символы до извлечения цифр
                s = s.translate(_fullwidth)
                if not s or s.lower() in ('none', 'null', 'nan'):
                    return ''
                try:
                    if isinstance(val, (int, float)) and val == int(val):
                        return str(int(val))
                except (ValueError, TypeError):
                    pass
                if re.match(r'^\d+\.0+$', s):
                    return str(int(float(s)))
                try:
                    n = float(s)
                    if n == int(n):
                        return str(int(n))
                except (ValueError, TypeError):
                    pass
                return re.sub(r'\D', '', s)

            def _raw_has_only_digits_or_plus_separators(raw, d):
                """После удаления пробелов, дефисов, скобок, точек и ведущего '+' остаётся только цифровая строка d."""
                if not raw or not d:
                    return raw == d or not raw
                r = str(raw).translate(_fullwidth)
                r = re.sub(r'[\s\-\(\)\.]', '', r)
                r = r.lstrip('+').strip()
                return r == d

            def _is_valid_format_39_5(d):
                """RCCONF_39.5: 10 цифр с 9; 11 цифр 89…; 11 цифр 79… (российский +7)."""
                if len(d) == 10 and d[0] == '9':
                    return True
                if len(d) == 11 and d.startswith('89'):
                    return True
                if len(d) == 11 and d.startswith('79'):
                    return True
                return False

            def _is_valid_format_39_5_2(d):
                """RCCONF_39.5.2: 10 цифр с 9; 11 цифр 89… (моб. 8-9xx); 11 цифр 79…; 11 цифр 8x (вторая ≠ 9)."""
                if len(d) == 10 and d[0] == '9':
                    return True
                if len(d) == 11 and d.startswith('89'):
                    return True
                if len(d) == 11 and d.startswith('79'):
                    return True
                if len(d) == 11 and d[0] == '8' and d[1] != '9':
                    return True
                return False

            def _is_error(val, rc):
                if val is None or (isinstance(val, float) and pd.isna(val)):
                    return False
                raw = _normalize_raw(val)
                if not raw or raw.lower() in ('none', 'null', 'nan'):
                    return False
                d = _to_digits(val)
                if not d:
                    return False
                # Допустимы только цифры и допустимые разделители (пробелы, дефисы, скобки, точки, ведущий +)
                if not _raw_has_only_digits_or_plus_separators(raw, d):
                    return True  # есть недопустимые символы
                is_valid = _is_valid_format_39_5(d) if rc == "RCCONF_39.5" else _is_valid_format_39_5_2(d)
                return not is_valid

            error_mask = df[column_name].apply(lambda v: _is_error(v, rule_code))
            error_count = int(error_mask.sum())
            if rule_code == "RCCONF_39.5.2":
                error_description = f'Invalid telephone number format in {column_name}. RCCONF_39.5.2: 10 digits starting with 9, or 11 digits 89…/79…, or 11 digits 8x (second digit not 9). All digits only (spaces, dashes, brackets, leading + allowed).'
            else:
                error_description = f'Invalid telephone number format in {column_name}. RCCONF_39.5: 10 digits starting with 9, or 11 digits 89…/79…. All digits only (spaces, dashes, brackets, leading + allowed).'
            
            if error_count > 0:
                error_df = self._prepare_error_dataframe(df, error_mask, 'CONFORMITY', error_description)
            else:
                error_df = None
            
            # Сохраняем ошибки через error_saver если он есть
            if error_df is not None and self.error_saver:
                self._save_errors_if_needed(error_df)
            
            # total_rows = количество строк с результатом '1' или '0' (не-NULL значения, которые были проверены)
            total_rows = non_null_mask.sum() if non_null_mask.any() else 0
            return total_rows, error_count, error_df
        else:
            # Базовая маска для непустых значений
            mask = df[column_name].notna() & (~df[column_name].astype(str).isin(["", "None", "null"]))
            
            # Проверяем соответствие справочнику если он есть
            if allowed_values:
                mask = mask & df[column_name].astype(str).isin(allowed_values)
            
            error_mask = ~mask
            error_count = error_mask.sum()
            
            # Формируем описание ошибки
            if allowed_values:
                sample_values = list(set(allowed_values))[:5]
                error_description = f'Invalid value in column {column_name}. Allowed: {sample_values}'
            else:
                error_description = f'Invalid value in column {column_name}'
        
        # Сохраняем ошибки
        error_df = self._prepare_error_dataframe(
            df, error_mask, 
            'CONFORMITY', 
            error_description
        )
        
        # Сохраняем ошибки через error_saver если он есть
        if error_df is not None and self.error_saver:
            self._save_errors_if_needed(error_df)
        
        # total_rows = количество строк с результатом '1' или '0' (непустые значения, которые были проверены)
        # mask содержит строки, которые были проверены (непустые значения)
        total_rows = mask.sum() if 'mask' in locals() else len(df)
        return total_rows, error_count, error_df