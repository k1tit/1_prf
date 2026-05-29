"""Специальный обработчик для таблиц DFKKBPTAXNUM. Использует conf_tax_number_format (таблица в БД или файл conf_tax_number_format.json)."""

import os
import json
import pandas as pd
import re
import numpy as np
from datetime import datetime
from typing import Dict, List, Tuple, Optional, Any, Iterable
from collections import Counter

class TaxNumHandler:
    """Обработчик проверок для таблиц с налоговыми номерами"""
    
    def __init__(self, table_name: str, df: pd.DataFrame, memory_manager, checker):
        self.table_name = table_name
        self.df = df
        self.memory_manager = memory_manager
        self.checker = checker
        
        self.results = []
        self.errors = {}
        self.current_result = None
        self.current_errors = []
        
        # Загружаем conf_tax_number_format.json (и при наличии — таблицу в БД)
        try:
            self.tax_formats, self.valid_lengths_by_country = self._load_tax_formats()
            self._format_regex_by_country = self._build_format_regex_cache()
        except Exception as e:
            print(f"      Ошибка загрузки конфигурации налоговых номеров: {e}")
            self.tax_formats = {'RU': ['[0-9]' * 8, '[0-9]' * 9, '[0-9]' * 10, '[0-9]' * 12]}
            self.valid_lengths_by_country = {'RU': {8, 9, 10, 12}}
            self._format_regex_by_country = self._build_format_regex_cache()
        
        self._apply_export_column_map()
        
        # Определяем какой это TAXNUM (1, 2, 3, 4, 5, 6)
        self.taxnum_type = self._get_taxnum_type(table_name)
        
        initial_len = len(self.df) if self.df is not None and not self.df.empty else 0
        # Срез по типу: DFKKBPTAXNUM5 / DFKKBPTAXNUM_RU5 — только taxtype=5 (RU5)
        needs_type_slice = bool(
            self.taxnum_type
            and (
                re.search(r'_RU\d$', table_name, re.I)
                or re.search(r'DFKKBPTAXNUM\d$', table_name, re.I)
            )
        )
        if needs_type_slice:
            type_col = self._find_taxtype_column()
            if type_col is not None:
                # Есть колонка типа (TAXTYPE и т.д.) — оставляем только строки с этим типом
                try:
                    type_ser = self.df[type_col].astype(str).str.strip()
                    numeric_type = pd.to_numeric(type_ser, errors='coerce')
                    # Допускаем: число N, строка "N", строка "RUN" (RU1, RU2, …)
                    mask = (
                        (numeric_type == self.taxnum_type)
                        | (type_ser == str(self.taxnum_type))
                        | (type_ser.str.upper() == f'RU{self.taxnum_type}')
                    )
                    before = len(self.df)
                    self.df = self.df.loc[mask].copy()
                    if before != len(self.df):
                        print(f"      [{table_name}] В отчёт «всего записей» только тип TAXNUM{self.taxnum_type}: {len(self.df):,} из {before:,}")
                    else:
                        # Фильтр не отсек строк — данные уже точечно по алиасу (срез из memory_manager) или колонка типа не та; оставляем как есть
                        pass
                except Exception as e:
                    print(f"      [WARN] [{table_name}] Ошибка фильтра по типу: {e}")
            else:
                # Нет колонки типа — фильтруем только если есть отдельная колонка TAXNUM1/TAXNUM2/… и она реально отсекает строки
                col = self._find_taxnum_column_for_type(self.taxnum_type)
                if col and col in self.df.columns:
                    mask = self.df[col].notna() & (self.df[col].astype(str).str.strip() != '')
                    before = len(self.df)
                    self.df = self.df.loc[mask].copy()
                    if before != len(self.df):
                        print(f"      [{table_name}] В отчёт «всего записей» только строки с заполненным TAXNUM{self.taxnum_type}: {len(self.df):,} из {before:,}")
                    else:
                        # Колонка заполнена у всех — по ней нельзя разделить типы, в отчёт не пишем весь массив
                        col_names = list(self.df.columns)
                        self.df = self.df.iloc[0:0].copy()
                        print(f"      [WARN] [{table_name}] TAXNUM{self.taxnum_type} заполнен у всех строк — в «всего записей» будет 0. Укажите колонку с типом (1,2,3,5) в conf_dfkkbptaxnum.json: taxtype_column. Колонки: {col_names[:25]}")
                else:
                    # Одна общая колонка TAXNUM без колонки типа — нельзя разделить по типам; в отчёт пишем 0, не весь массив
                    col_names = list(self.df.columns)
                    before = len(self.df)
                    self.df = self.df.iloc[0:0].copy()
                    print(f"      [WARN] [{table_name}] Колонка типа налога не найдена — в «всего записей» записано 0 (не весь массив). Укажите имя колонки с типом (1,2,3,5) в conf_dfkkbptaxnum.json (ключ taxtype_column). Доступные колонки: {col_names[:20]}")
                    if before > 0:
                        print(f"      [WARN] Пример: {{\"taxtype_column\": \"имя_вашей_колонки\"}}")
        
            # Защита: если НЕ отфильтровали и при этом df большой (похоже на полную таблицу) — предупреждение (данные могли прийти точечно по алиасу из memory_manager, тогда initial_len уже по алиасу)
            if re.search(r'_RU\d$', table_name, re.I) and self.taxnum_type and initial_len > 0 and len(self.df) == initial_len and initial_len > 500_000:
                print(f"      [INFO] [{table_name}] В отчёт «всего записей»: {initial_len:,} строк (если данные загружены точечно по алиасу — это корректно)")
        
        print(f"      Загружено {len(self.valid_lengths_by_country)} стран с конфигурацией налоговых номеров")
        for country, lengths in self.valid_lengths_by_country.items():
            print(f"        {country}: допустимые длины {sorted(lengths)}")
    
    def _conf_tax_format_json_paths(self) -> List[str]:
        paths: List[str] = []
        if hasattr(self.checker, 'rules_file') and self.checker.rules_file:
            paths.append(
                os.path.join(
                    os.path.dirname(self.checker.rules_file),
                    'conf_tax_number_format.json',
                )
            )
        script_dir = os.path.dirname(os.path.abspath(__file__))
        paths.extend([
            os.path.join(script_dir, '..', 'json files', 'conf_tax_number_format.json'),
            os.path.join(script_dir, '..', 'config', 'conf_tax_number_format.json'),
        ])
        try:
            paths.append(os.path.join(os.getcwd(), 'json files', 'conf_tax_number_format.json'))
        except Exception:
            pass
        seen: set[str] = set()
        out: List[str] = []
        for p in paths:
            ap = os.path.abspath(p)
            if ap not in seen:
                seen.add(ap)
                out.append(ap)
        return out

    def _load_tax_formats(self) -> Tuple[Dict[str, List[str]], Dict[str, set]]:
        """Загружает форматы из conf_tax_number_format.json (и при наличии — из БД).
        RCCONF_63.1: IF format not defined for country THEN '' (строка не оценивается)."""
        formats: Dict[str, List[str]] = {}
        valid_lengths: Dict[str, set] = {}

        def add_entry(country: str, tax_format: str, length: Optional[int] = None):
            country = country.upper()
            if not country:
                return
            if country not in formats:
                formats[country] = []
                valid_lengths[country] = set()
            if tax_format and tax_format not in formats[country]:
                formats[country].append(tax_format)
            if length is not None:
                valid_lengths[country].add(length)
            elif tax_format:
                n = tax_format.count('[0-9]')
                if n > 0:
                    valid_lengths[country].add(n)

        def load_from_records(records: Iterable[Any], source: str) -> int:
            n = 0
            for item in records:
                if isinstance(item, dict):
                    country = str(item.get('country_code', ''))
                    tax_format = str(item.get('tax_format', ''))
                    length = item.get('length')
                    if length is not None:
                        try:
                            length = int(length)
                        except (TypeError, ValueError):
                            length = None
                    add_entry(country, tax_format, length)
                    n += 1
                elif hasattr(item, 'get'):
                    pass
            if n:
                print(f"      Загрузка conf_tax_number_format из {source}: {n} записей")
            return n

        # 1) JSON в проекте (основной источник для RCCONF_63.1)
        for conf_path in self._conf_tax_format_json_paths():
            if not os.path.isfile(conf_path):
                continue
            try:
                with open(conf_path, 'r', encoding='utf-8') as f:
                    conf_list = json.load(f)
                load_from_records(
                    conf_list if isinstance(conf_list, list) else [conf_list],
                    conf_path,
                )
            except Exception as e:
                print(f"      Ошибка загрузки {conf_path}: {e}")

        # 2) Таблица в БД (дополняет JSON)
        try:
            formats_df = self.memory_manager.get_table('conf_tax_number_format')
            if formats_df is not None and not formats_df.empty:
                for _, row in formats_df.iterrows():
                    country = str(row.get('country_code', ''))
                    tax_format = str(row.get('tax_format', ''))
                    length = row.get('length')
                    if length is not None and not pd.isna(length):
                        try:
                            length = int(length)
                        except (TypeError, ValueError):
                            length = None
                    add_entry(country, tax_format, length)
                print(f"      Дополнение из БД conf_tax_number_format: {len(formats_df)} строк")
        except Exception as e:
            print(f"      conf_tax_number_format в БД недоступна: {e}")

        # 3) Дефолт для России, если конфиг не найден
        if not formats:
            default_ru = [
                ('[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]', 8),
                ('[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]', 9),
                ('[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]', 10),
                ('[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]', 12),
            ]
            formats['RU'] = [f for f, _ in default_ru]
            valid_lengths['RU'] = {l for _, l in default_ru}
            print(f"      Использована конфигурация по умолчанию для RU: длины {sorted(valid_lengths['RU'])}")
        
        return formats, valid_lengths

    def _build_format_regex_cache(self) -> Dict[str, List[re.Pattern]]:
        cache: Dict[str, List[re.Pattern]] = {}
        for country, fmt_list in (self.tax_formats or {}).items():
            patterns: List[re.Pattern] = []
            for fmt in fmt_list:
                try:
                    patterns.append(self._convert_format_to_regex(fmt))
                except Exception:
                    continue
            if patterns:
                cache[country.upper()] = patterns
        return cache

    def _apply_export_column_map(self) -> None:
        """Шапка выгрузки (Tax_Number, Tax_Number_Category) -> TAXNUM/TAXTYPE для правил."""
        if self.df is None or self.df.empty:
            return
        try:
            from utils.column_map_resolver import load_column_map, apply_column_headers_for_rules
            root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            cm = getattr(self.checker, 'column_map', None) or load_column_map(root)
            if cm:
                self.df = apply_column_headers_for_rules(
                    self.df, self.table_name, cm, root, log_renames=False
                )
        except Exception as e:
            print(f"      [WARN] column_map для {self.table_name}: {e}")

    def _country_has_reference_format(self, country: str) -> bool:
        c = (country or '').strip().upper() or 'RU'
        return bool(
            (self._format_regex_by_country or {}).get(c)
            or (self.valid_lengths_by_country or {}).get(c)
        )

    def _tax_value_matches_country(self, value: str, country: str) -> bool:
        c = (country or '').strip().upper() or 'RU'
        patterns = (self._format_regex_by_country or {}).get(c)
        if patterns:
            return any(p.match(value) for p in patterns)
        lengths = (self.valid_lengths_by_country or {}).get(c)
        if lengths and value.isdigit():
            return len(value) in lengths
        return False

    def _get_taxnum_type(self, table_name: str) -> int:
        """Определяет тип TAXNUM из названия таблицы (DFKKBPTAXNUM1 → 1, DFKKBPTAXNUM_RU2 → 2 и т.д.)."""
        m = re.search(r'DFKKBPTAXNUM(\d)$', table_name, re.I)
        if m:
            return int(m.group(1))
        m = re.search(r'_RU(\d)$', table_name, re.I)
        if m:
            return int(m.group(1))
        if 'TAXNUM5' in table_name.upper():
            return 5
        elif 'TAXNUM4' in table_name.upper():
            return 4
        elif 'TAXNUM3' in table_name.upper():
            return 3
        elif 'TAXNUM2' in table_name.upper():
            return 2
        elif 'TAXNUM1' in table_name.upper():
            return 1
        elif 'TAXNUM6' in table_name.upper():
            return 6
        else:
            # По умолчанию ищем в колонках
            for col in self.df.columns:
                if 'TAXNUM' in col.upper():
                    num_part = col.upper().replace('TAXNUM', '')
                    if num_part.isdigit():
                        return int(num_part)
            return 0
    
    def _get_country_for_row(self, row) -> str:
        """Определяет страну для строки данных"""
        # Ищем колонки с информацией о стране
        country_columns = ['COUNTRY', 'LAND1', 'LAND', 'COUNTRY_CODE', 'CTRY', 'LANDX', 'NATION']
        
        for col in country_columns:
            if col in row and not pd.isna(row[col]):
                country = str(row[col]).strip().upper()
                if country:
                    return country
        
        # Если не нашли - проверяем есть ли колонки с названием страны
        for col in self.df.columns:
            col_upper = col.upper()
            if any(keyword in col_upper for keyword in ['COUNTRY', 'LAND', 'NATION', 'CTRY']):
                if col in row and not pd.isna(row[col]):
                    country = str(row[col]).strip().upper()
                    if country:
                        return country
        
        # Если не нашли - используем Россию по умолчанию
        return 'RU'
    
    def _load_dfkkbptaxnum_config(self) -> Optional[str]:
        """Читает conf_dfkkbptaxnum.json и возвращает имя колонки типа (taxtype_column), если задано."""
        candidates = self._load_dfkkbptaxnum_config_candidates()
        return candidates[0] if candidates else None
    
    def _load_dfkkbptaxnum_config_candidates(self) -> List[str]:
        """Возвращает список имён колонок типа для перебора: taxtype_column + taxtype_column_alternatives."""
        result = []
        if __file__:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            candidates = [
                os.path.join(script_dir, '..', 'json files', 'conf_dfkkbptaxnum.json'),
                os.path.join(script_dir, '..', '..', 'json files', 'conf_dfkkbptaxnum.json'),
            ]
        else:
            candidates = []
        try:
            cwd = os.getcwd()
            candidates.append(os.path.join(cwd, 'json files', 'conf_dfkkbptaxnum.json'))
            candidates.append(os.path.join(cwd, 'config', 'conf_dfkkbptaxnum.json'))
        except Exception:
            pass
        for path in candidates:
            path = os.path.abspath(path)
            if os.path.isfile(path):
                try:
                    with open(path, 'r', encoding='utf-8') as f:
                        cfg = json.load(f)
                    col = cfg.get('taxtype_column') or cfg.get('taxtype_column_name')
                    if col and isinstance(col, str) and col.strip():
                        result.append(col.strip())
                    alts = cfg.get('taxtype_column_alternatives') or cfg.get('taxtype_columns') or []
                    if isinstance(alts, list):
                        for c in alts:
                            if c and isinstance(c, str) and c.strip() and c.strip() not in result:
                                result.append(c.strip())
                    if result:
                        return result
                except Exception:
                    pass
        return result
    
    def _find_taxtype_column(self) -> Optional[str]:
        """Ищет колонку с типом налога (1, 2, 3, 5) для фильтра «только записи этого типа»."""
        if self.df is None or self.df.empty:
            return None
        # 1) Из конфига (основное имя + альтернативы)
        config_candidates = self._load_dfkkbptaxnum_config_candidates()
        for config_col in config_candidates:
            if config_col and config_col in self.df.columns:
                print(f"      [DFKKBPTAXNUM] Колонка типа из конфига: {config_col}")
                return config_col
        if config_candidates:
            print(f"      [WARN] conf_dfkkbptaxnum.json задаёт колонку(и) {config_candidates}, ни одна не найдена в таблице. Доступные: {list(self.df.columns)[:15]}...")
        # 2) Эвристика: колонка, у которой все значения из {1,2,3,4,5,6}
        allowed = {1, 2, 3, 4, 5, 6}
        for col in self.df.columns:
            try:
                ser = self.df[col].dropna().astype(str).str.strip()
                nums = pd.to_numeric(ser, errors='coerce').dropna()
                if len(nums) == 0:
                    continue
                uniq = set(int(x) for x in nums.unique() if x == int(x))
                if uniq and uniq <= allowed:
                    print(f"      [DFKKBPTAXNUM] Колонка типа по значениям: {col}")
                    return col
            except Exception:
                continue
        # 3) По имени
        candidates = [
            'Tax_Number_Category', 'TAXTYPE', 'taxtype', 'TAX_TYPE', 'tax_type',
            'TAXNUMTYPE', 'type', 'TYPE',
        ]
        for c in candidates:
            if c in self.df.columns:
                return c
        for col in self.df.columns:
            cu = str(col).upper().replace(' ', '').replace('_', '')
            if cu in ('TAXTYPE', 'TAXTYP', 'TAXNUMTYPE', 'TYPE'):
                return col
            if 'TAXTYPE' in cu or 'TAXTYP' in cu:
                return col
        return None
    
    def _find_taxnum_column_for_type(self, taxnum_type: int) -> Optional[str]:
        """Находит колонку с данными для TAXNUM N (для фильтра «только строки с заполненным TAXNUMx»).
        Пробует: TAXNUM1, tax_1_value, Tax Number 1, TAXNUM 1 и варианты по регистру."""
        if not taxnum_type or self.df is None or self.df.empty:
            return None
        candidates = [
            f'TAXNUM{taxnum_type}',
            f'TAXNUM {taxnum_type}',
            f'tax_{taxnum_type}_value',
            f'Tax Number {taxnum_type}',
            f'TaxNum{taxnum_type}',
        ]
        for c in candidates:
            if c in self.df.columns:
                return c
        for col in self.df.columns:
            cu = str(col).upper()
            if cu == f'TAXNUM{taxnum_type}' or cu.replace(' ', '') == f'TAXNUM{taxnum_type}':
                return col
            if f'TAXNUM{taxnum_type}' in cu or (f'TAX_{taxnum_type}' in cu and 'VALUE' in cu):
                return col
        return None
    
    def find_column(self, column_name: str) -> Optional[str]:
        """Находит реальное имя колонки (TAXNUM1 в правиле → TAXNUM в DFKKBPTAXNUM1)."""
        if not column_name or self.df is None or self.df.empty:
            return None
        if column_name in self.df.columns:
            return column_name

        try:
            from utils.column_map_resolver import resolve_column_in_df, load_column_map
            root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            cm = getattr(self.checker, 'column_map', None) or load_column_map(root)
            found = resolve_column_in_df(
                self.df, column_name, self.table_name, cm, root
            )
            if found:
                return found
        except ImportError:
            pass

        tname = str(self.table_name or "").upper()
        cname = str(column_name).upper()
        if tname.startswith("DFKKBPTAXNUM"):
            if re.match(r"TAXNUM\d$", cname):
                for prefer in ("Tax_Number", "Tax_Number_Long", "TAXNUM", "TAXNUM_LONG"):
                    if prefer in self.df.columns:
                        return prefer
                for col in self.df.columns:
                    cu = str(col).strip().upper().replace(" ", "").replace("_", "")
                    if cu in ("TAXNUMBER", "TAXNUM") and "CATEGORY" not in cu and "LONG" not in cu:
                        return col
                    if cu == "TAXNUMBERLONG":
                        return col
            if re.match(r"TAXTYPE\d$", cname):
                for col in self.df.columns:
                    cu = str(col).strip().upper()
                    if cu in ("TAXTYPE", "TAXTYP", "TYPE"):
                        return col

        column_upper = column_name.upper().replace(" ", "").replace("_", "")
        for col in self.df.columns:
            cu = col.upper().replace(" ", "").replace("_", "")
            if column_upper == cu:
                return col
            if column_upper == "TAXNUM" and cu in ("TAXNUMBER", "TAXNUM"):
                return col
            if column_upper == "TAXNUM" and cu == "TAXNUMBERLONG":
                continue
            if len(column_upper) >= 5 and column_upper == cu:
                return col

        return None
    
    def validate_rule(self, rule: dict):
        """Выполняет проверку правила для налоговых номеров. Возвращает dict для core checker (сводка)."""
        rule_code = rule.get("rule_code", "UNKNOWN")
        rule_description = rule.get("rule_description", "")
        quality_category = rule.get("quality_category", "")
        column_to_check = rule.get("column_name_checked", "")
        
        self.current_result = None
        self.current_errors = []
        
        # Находим реальную колонку (Tax_Number, не Tax_Number_Long)
        real_column = self.find_column(column_to_check)
        if real_column and real_column not in self.df.columns:
            real_column = None
        if not real_column:
            print(f"      Колонка '{column_to_check}' не найдена в {self.table_name}")
            self._save_empty_result(rule_code, rule_description, column_to_check, rule)
            return self._build_result_for_core(rule_code, rule_description, quality_category, column_to_check, 0, 0, 0, 0.0, 0.0, "ОШИБКА ВЫПОЛНЕНИЯ", pd.DataFrame())
        
        print(f"      Используем колонку: {real_column}")
        
        start_time = datetime.now()
        
        try:
            # Всего записей = только по этой таблице-алиасу (DFKKBPTAXNUM_RU1, _RU2, …); df уже срез по типу из memory_manager
            total = len(self.df)
            error_count = 0
            error_df = pd.DataFrame()
            
            # Исходный вариант: только правила 63.1 (формат) и 63.7 (уникальность)
            if rule_code == 'RCCONF_63.1':
                error_count, error_df, total = self._validate_taxnum_format(real_column)
            elif rule_code == 'RCCONF_63.7':
                error_count, error_df = self._validate_taxnum_uniqueness(real_column)
            else:
                print(f"      Правило {rule_code} не поддерживается для налоговых номеров")
                error_count = 0
                error_df = pd.DataFrame()

            error_df = self._annotate_error_df(error_df, rule_code, rule_description, real_column)
            
            execution_time = (datetime.now() - start_time).total_seconds()
            
            # Проверяем на массовые ошибки
            is_suspicious = self._check_if_suspicious(rule_code, error_count, total)
            
            # Сохраняем результат
            self._save_result(rule_code, rule_description, rule, real_column, 
                            total, error_count, execution_time, is_suspicious, error_df)
            
            # Выводим результат
            self._print_result(rule_code, error_count, total, execution_time, is_suspicious)
            
            success = total - error_count
            success_rate = (success / total * 100) if total > 0 else 0
            status = "УСПЕШНО" if error_count == 0 else ("МАССОВЫЕ ОШИБКИ" if is_suspicious else "ОШИБКИ")
            return self._build_result_for_core(rule_code, rule_description, quality_category, real_column, total, success, error_count, success_rate, execution_time, status, error_df)
            
        except Exception as e:
            execution_time = (datetime.now() - start_time).total_seconds()
            print(f"      Ошибка выполнения: {str(e)}")
            import traceback
            traceback.print_exc()
            self._save_empty_result(rule_code, rule_description, column_to_check, rule)
            return self._build_result_for_core(rule_code, rule_description, quality_category, column_to_check, 0, 0, 0, 0.0, execution_time, "ОШИБКА ВЫПОЛНЕНИЯ", pd.DataFrame())
    
    def _build_full_error_row(self, idx, dq_extras: Optional[dict] = None) -> dict:
        """Все колонки строки из DFKKBPTAXNUM* + служебные поля DQ_* (для полноценной выгрузки в CSV)."""
        row = self.df.loc[idx]
        out = {str(c): row.loc[c] for c in self.df.columns}
        out["DQ_SOURCE_ROW_INDEX"] = idx
        if dq_extras:
            for k, v in dq_extras.items():
                key = k if str(k).startswith("DQ_") else f"DQ_{k}"
                out[key] = v
        return out

    def _annotate_error_df(
        self,
        error_df: pd.DataFrame,
        rule_code: str,
        rule_description: str,
        real_column: str,
    ) -> pd.DataFrame:
        """Единые метаданные качества, как у стандартных валидаторов."""
        if error_df is None or error_df.empty:
            return error_df
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        out = error_df.copy()
        out["DQ_RULE_CODE"] = rule_code
        out["DQ_RULE_DESCRIPTION"] = rule_description
        out["DQ_COLUMN_CHECKED"] = real_column
        out["DQ_ERROR_TYPE"] = "CONFORMITY"
        out["DQ_TIMESTAMP"] = ts
        return out

    def _coalesce_tax_number_series(self) -> Tuple[str, pd.Series]:
        """
        Номер в выгрузке SAP: Tax_Number и/или Tax_Number_Long (часто заполнено только одно).
        Не Tax_Number_Category (там RU5 = тип, не значение номера).
        """
        short_col = None
        for name in ("Tax_Number", "TAXNUM", f"tax_{self.taxnum_type}_value" if self.taxnum_type else ""):
            if name and name in self.df.columns:
                short_col = name
                break
        long_col = None
        for name in ("Tax_Number_Long", "TAXNUM_LONG"):
            if name in self.df.columns:
                long_col = name
                break
        if not short_col and not long_col:
            return "", pd.Series("", index=self.df.index)

        short_s = (
            self.df[short_col].map(self._normalize_taxnum_to_str)
            if short_col
            else pd.Series("", index=self.df.index)
        )
        long_s = (
            self.df[long_col].map(self._normalize_taxnum_to_str)
            if long_col
            else pd.Series("", index=self.df.index)
        )
        combined = short_s.where(short_s != "", long_s)
        if short_col and long_col:
            label = f"{short_col}|{long_col}"
        else:
            label = short_col or long_col or ""
        return label, combined

    def _normalize_taxnum_to_str(self, value) -> str:
        """Приводит значение к строке из цифр: число без .0 и научной записи, кавычки с краёв убираем."""
        if pd.isna(value):
            return ""
        if isinstance(value, (int, float)):
            if value == int(value):
                return str(int(value))
            return str(value).strip()
        s = str(value).strip()
        for q in ("'", '"', "`"):
            if len(s) >= 2 and s[0] == q and s[-1] == q:
                s = s[1:-1].strip()
                break
        if s.endswith(".0") and s[:-2].isdigit():
            s = s[:-2]
        return s

    def _validate_taxnum_format(self, column_name: str) -> Tuple[int, pd.DataFrame, int]:
        """
        RCCONF_63.1 (TAXNUM5 / tax_5_value):
        IF tax_5_value IS NULL or reference format not defined THEN ''
        ELSE IF tax_5_value has correct format THEN '1' ELSE '0'
        Значение: Tax_Number, если пусто — Tax_Number_Long (не Tax_Number_Category/RU5).
        """
        value_label, ser = self._coalesce_tax_number_series()
        if not value_label or ser.empty:
            print(f"      [WARN] RCCONF_63.1: нет Tax_Number / Tax_Number_Long в {self.table_name}")
            return 0, pd.DataFrame(), 0

        ru_lengths = sorted((self.valid_lengths_by_country or {}).get("RU", set()))
        if not ru_lengths:
            ru_lengths = [8, 9, 10, 12]
        allowed_lengths = set(ru_lengths)

        print(
            f"      RCCONF_63.1: значение из {value_label} (не Tax_Number_Category) — "
            f"только цифры, длина in {ru_lengths}"
        )
        non_empty = ser != ""
        evaluated = int(non_empty.sum())
        if evaluated == 0:
            print(f"      [INFO] RCCONF_63.1: нет заполненных значений в {column_name}")
            return 0, pd.DataFrame(), 0

        digits_only = ser.str.match(r"^\d+$", na=False)
        length_ok = ser.str.len().isin(allowed_lengths)
        ok_mask = non_empty & digits_only & length_ok
        error_mask = non_empty & ~ok_mask
        error_count = int(error_mask.sum())
        ok_count = int(ok_mask.sum())

        error_df = pd.DataFrame()
        if error_count > 0:
            # В файл ошибок — полная строка из DFKKBPTAXNUM* (все колонки выгрузки), без обрезки
            error_df = self.df.loc[error_mask].copy()
            error_df["DQ_SOURCE_ROW_INDEX"] = error_df.index
            norm_ser = ser.loc[error_mask]
            error_df["DQ_TAX_NUMBER_NORMALIZED"] = norm_ser
            error_df["DQ_TAXNUM_LENGTH"] = norm_ser.str.len()
            error_df["DQ_COUNTRY_USED_FOR_RULE"] = "RU"
            error_df["DQ_CHECKED_COLUMN"] = value_label
            error_df["DQ_EXPECTED_FORMAT"] = f"conf_tax_number_format RU: lengths {ru_lengths}"
            error_df["DQ_ERROR_DESCRIPTION"] = error_df["DQ_TAX_NUMBER_NORMALIZED"].map(
                lambda s: (
                    f"Недопустимый TAXNUM{self.taxnum_type}: "
                    f"длина {len(str(s))}, значение {str(s)[:40]}. "
                    f"Допустимо: только цифры, длина {ru_lengths}"
                )
            )
            print(f"      [INFO] В файл ошибок: {len(error_df):,} строк, {len(error_df.columns)} колонок (полная строка таблицы)")

        skipped_null = len(self.df) - evaluated
        print(f"      Статистика RCCONF_63.1:")
        print(f"        Строк в срезе TAXNUM{self.taxnum_type}: {len(self.df):,}")
        print(f"        Пропущено (пусто, IF NULL THEN ''): {skipped_null:,}")
        print(f"        Оценено: {evaluated:,} (успешно: {ok_count:,}, ошибок: {error_count:,})")

        return error_count, error_df, evaluated
    
    def _convert_format_to_regex(self, tax_format: str) -> re.Pattern:
        """Конвертирует формат из конфигурации в regex"""
        # Пример: [0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]
        # Преобразуем в ^\d{10}$
        
        # Убираем квадратные скобки
        cleaned = tax_format.replace('[0-9]', '\\d')
        
        # Считаем количество \d
        count = cleaned.count('\\d')
        
        if count > 0:
            return re.compile(f'^\\d{{{count}}}$')
        else:
            # Если формат не распознан, создаем regex из исходной строки
            return re.compile(f'^{re.escape(tax_format)}$')
    
    def _validate_taxnum_uniqueness(self, column_name: str) -> Tuple[int, pd.DataFrame]:
        """Проверяет только TAXNUM5 — не должен совпадать с TAXNUM 1,2,3,4,6. При одной таблице (алиасы RU1..RU6) — пропуск (O(1))."""
        errors = []
        
        if column_name not in self.df.columns:
            return 0, pd.DataFrame()
        
        if self.taxnum_type != 5:
            return 0, pd.DataFrame()
        
        other_types = [1, 2, 3, 4, 6]
        print(f"      Проверка уникальности TAXNUM5 (не должен совпадать с TAXNUM {other_types})")
        taxnum_tables = self._get_all_taxnum_tables()
        
        # Быстрый путь: если все «другие» таблицы — тот же DataFrame (алиасы одной DFKKBPTAXNUM), не делаем O(n²)
        self_id = id(self.df)
        all_same = all(
            taxnum_tables.get(ot) is not None and id(taxnum_tables[ot]) == self_id
            for ot in other_types
        )
        if all_same:
            print(f"      [INFO] Одна таблица DFKKBPTAXNUM (срезы DFKKBPTAXNUM1..6), сравнение TAXNUM5 с другими типами пропущено (не применимо).")
            return 0, pd.DataFrame()
        
        # Таблицы разные: используем множества вместо вложенных циклов O(n*m)
        for other_type in other_types:
            other_table = taxnum_tables.get(other_type)
            if other_table is not None and not other_table.empty and id(other_table) != self_id:
                errors.extend(self._compare_with_other_taxnum_fast(column_name, other_table, other_type))
        
        error_df = pd.DataFrame(errors) if errors else pd.DataFrame()
        return len(errors), error_df
    
    def _get_all_taxnum_tables(self) -> Dict[int, pd.DataFrame]:
        """Ищем таблицы по имени: DFKKBPTAXNUM1 → 1, DFKKBPTAXNUM2 → 2 или DFKKBPTAXNUM_RU1 → 1 и т.д."""
        tables = {}
        for table_name in self.memory_manager.data_cache.keys():
            s = str(table_name)
            m = re.search(r'DFKKBPTAXNUM(\d)$', s, re.I)
            if not m:
                m = re.search(r'_RU(\d)$', s, re.I)
            if m:
                num = int(m.group(1))
                tables[num] = self.memory_manager.get_table(table_name)
        if not tables and self.df is not None and not self.df.empty:
            tables[self.taxnum_type] = self.df
        return tables
    
    def _compare_with_other_taxnum(self, column_name: str, other_df: pd.DataFrame, other_type: int) -> List[dict]:
        """Сравнивает текущий TAXNUM с другим (ищем колонку TAXNUM other_type в other_df). Только для маленьких таблиц; для больших используйте _compare_with_other_taxnum_fast."""
        errors = []
        target = f'TAXNUM{other_type}'
        other_column = None
        for col in other_df.columns:
            cu = str(col).upper().replace(' ', '')
            if cu == target or target in cu or (f'TAX_{other_type}' in cu and 'VALUE' in cu):
                other_column = col
                break
        if not other_column:
            for col in other_df.columns:
                if 'TAXNUM' in col.upper():
                    other_column = col
                    break
        if not other_column:
            return errors
        
        # Проходим по всем строкам текущей таблицы
        for idx, row in self.df.iterrows():
            current_value = row.get(column_name)
            
            # Пропускаем NULL
            if pd.isna(current_value):
                continue
            
            current_str = str(current_value).strip()
            
            # Пропускаем пустые
            if not current_str:
                continue
            
            # Ищем такой же номер в другой таблице
            for other_idx, other_row in other_df.iterrows():
                other_value = other_row.get(other_column)
                
                if pd.isna(other_value):
                    continue
                
                other_str = str(other_value).strip()
                
                if not other_str:
                    continue
                
                # Если значения совпадают - ошибка
                if current_str == other_str:
                    desc = f'TAXNUM{self.taxnum_type} совпадает с TAXNUM{other_type}'
                    errors.append(
                        self._build_full_error_row(
                            idx,
                            {
                                "DQ_ERROR_DESCRIPTION": desc,
                                "DQ_OTHER_TAXNUM_TYPE": other_type,
                                "DQ_CURRENT_TAXNUM": current_value,
                                "DQ_OTHER_TAXNUM": other_value,
                            },
                        )
                    )
                    break  # Одна ошибка на строку достаточно
        
        return errors
    
    def _compare_with_other_taxnum_fast(self, column_name: str, other_df: pd.DataFrame, other_type: int) -> List[dict]:
        """Быстрое сравнение через множество значений O(n+m). Лимит ошибок для отчёта."""
        max_errors = 100_000
        errors = []
        target = f'TAXNUM{other_type}'
        other_column = None
        for col in other_df.columns:
            cu = str(col).upper().replace(' ', '')
            if cu == target or target in cu or (f'TAX_{other_type}' in cu and 'VALUE' in cu):
                other_column = col
                break
        if not other_column:
            for col in other_df.columns:
                if 'TAXNUM' in col.upper():
                    other_column = col
                    break
        if not other_column:
            return errors
        
        # Множество значений из другой таблицы (нормализованные строки)
        other_ser = other_df[other_column].astype(str).str.strip()
        other_set = set(other_ser[other_ser != ''].dropna().unique())
        
        # Проверяем каждую строку текущей таблицы: значение в множестве?
        cur_ser = self.df[column_name].astype(str).str.strip()
        mask = cur_ser.isin(other_set) & (cur_ser != '') & cur_ser.notna()
        hit_indices = self.df.index[mask].tolist()
        
        for i, idx in enumerate(hit_indices):
            if len(errors) >= max_errors:
                break
            current_value = self.df.at[idx, column_name]
            desc = f'TAXNUM{self.taxnum_type} совпадает с TAXNUM{other_type}'
            errors.append(
                self._build_full_error_row(
                    idx,
                    {
                        "DQ_ERROR_DESCRIPTION": desc,
                        "DQ_OTHER_TAXNUM_TYPE": other_type,
                        "DQ_CURRENT_TAXNUM": current_value,
                    },
                )
            )
        
        if mask.sum() > max_errors:
            print(f"      [INFO] Ограничение отчёта: показано {max_errors} из {mask.sum()} совпадений с TAXNUM{other_type}")
        return errors
    
    def _check_if_suspicious(self, rule_code: str, error_count: int, total_rows: int) -> bool:
        """Проверяет подозрительные правила"""
        if error_count > 1000000:
            return True
        if total_rows > 0 and (error_count / total_rows) > 0.8:
            return True
        return False
    
    def _save_result(self, rule_code: str, rule_description: str, rule: dict, 
                    matched_column: str, total_rows: int, error_count: int, 
                    execution_time: float, is_suspicious: bool, error_df: pd.DataFrame):
        """Сохраняет результат проверки"""
        success = total_rows - error_count
        success_rate = (success / total_rows * 100) if total_rows > 0 else 0
        
        if error_count == 0:
            status = "УСПЕШНО"
            status_color = "green"
        elif is_suspicious:
            status = "МАССОВЫЕ ОШИБКИ"
            status_color = "orange"
        else:
            status = "ОШИБКИ"
            status_color = "red"
        
        result = {
            'rule_code': rule_code,
            'total_records': total_rows,
            'passed': success,
            'failed': error_count,
            'success_rate_%': round(success_rate, 2),
            'execution_time_sec': round(execution_time, 2),
            'status': status,
            'status_color': status_color,
            'matched_column': matched_column
        }
        
        self.current_result = result
        
        # Сохраняем ошибки
        if error_count > 0 and not error_df.empty:
            self.errors[rule_code] = {
                'error_df': error_df,
                'error_count': error_count,
                'is_suspicious': is_suspicious,
                'total_rows': total_rows
            }
    
    def _save_empty_result(self, rule_code: str, rule_description: str, 
                          column_checked: str, rule: dict):
        """Сохраняет пустой результат при ошибке"""
        result = {
            'rule_code': rule_code,
            'total_records': 0,
            'passed': 0,
            'failed': 0,
            'success_rate_%': 0,
            'execution_time_sec': 0,
            'status': "ОШИБКА ВЫПОЛНЕНИЯ",
            'status_color': "dark_red",
            'matched_column': column_checked
        }
        
        self.current_result = result
    
    def _build_result_for_core(self, rule_code: str, rule_description: str, quality_category: str,
                              column_checked: str, total_records: int, passed: int, failed: int,
                              success_rate: float, execution_time_sec: float, status: str,
                              error_df: pd.DataFrame) -> dict:
        """Формирует dict результата в формате, ожидаемом core checker (сводка и отчёт)."""
        status_color = "green" if status == "УСПЕШНО" else ("orange" if "МАССОВЫЕ" in status or "ПОДОЗРИТЕЛЬНО" in status else "red")
        error_file = "Есть" if failed > 0 and error_df is not None and not error_df.empty else "Нет"
        comments = ""
        if failed > 0 and total_records > 0:
            pct = (failed / total_records * 100)
            if pct > 50:
                comments = f"ПОДОЗРИТЕЛЬНО: {pct:.1f}% ДАННЫХ С ОШИБКАМИ - ПРОВЕРИТЬ ЛОГИКУ ПРАВИЛА"
        return {
            "rule_code": rule_code,
            "rule_description": rule_description,
            "quality_category": quality_category,
            "table_name": self.table_name,
            "column_checked": column_checked,
            "total_records": total_records,
            "passed": passed,
            "failed": failed,
            "error_count": failed,
            "success_rate_%": round(success_rate, 2),
            "execution_time_sec": round(execution_time_sec, 2),
            "status": status,
            "status_color": status_color,
            "error_df": error_df if error_df is not None else pd.DataFrame(),
            "error_file": error_file,
            "comments": comments
        }
    
    def _print_result(self, rule_code: str, error_count: int, total_rows: int, 
                     execution_time: float, is_suspicious: bool):
        """Выводит результат проверки"""
        if error_count == 0:
            print(f"      УСПЕХ: 0 ошибок ({execution_time:.2f}с)")
        elif is_suspicious:
            error_percent = (error_count / total_rows * 100) if total_rows > 0 else 0
            print(f"      МАССОВЫЕ ОШИБКИ: {error_count:,} из {total_rows:,} ({error_percent:.1f}%) ({execution_time:.2f}с)")
        else:
            success_rate = ((total_rows - error_count) / total_rows * 100) if total_rows > 0 else 0
            print(f"      ОШИБКИ: {error_count:,} ({success_rate:.1f}% успеха, {execution_time:.2f}с)")
    
    def get_results(self) -> List[dict]:
        """Возвращает результаты"""
        if self.current_result:
            return [self.current_result]
        return []
    
    def get_errors(self) -> Dict[str, Any]:
        """Возвращает ошибки"""
        return self.errors