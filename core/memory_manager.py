# core/memory_manager.py
"""Менеджер памяти и кэширования данных"""

import gc
import os
import json
import re
import sqlite3
from datetime import datetime
import pandas as pd
from concurrent.futures import ThreadPoolExecutor

from utils.empty_rows import fully_empty_rows_mask
from utils.sqlite_safe import connect_sqlite

# Опционально: async загрузка таблиц (pip install aiosqlite)
try:
    import asyncio
    import aiosqlite
    _HAS_AIOSQLITE = True
except ImportError:
    _HAS_AIOSQLITE = False

TERMINAL_SYMBOLS_LOCAL = {
    'SUCCESS': '[OK]',
    'ERROR': '[ERROR]',
    'WARNING': '[WARN]',
    'INFO': '[INFO]',
    'SKIP': '[SKIP]',
    'ROCKET': '[START]',
    'GEAR': '[PROC]',
    'MAGNIFYING_GLASS': '[CHECK]',
    'CHECKMARK': '[DONE]',
    'CHART_UP': '[STAT+]',
    'CHART_DOWN': '[STAT-]',
    'CLIPBOARD': '[CLIP]',
    'FILE_FOLDER': '[DIR]',
    'PAGE': '[FILE]',
    'TABLE': '[TABLE]',
    'COLUMN': '[COL]',
    'BOOKS': '[DATA]',
    'SAVE': '[SAVE]',
    'TARGET': '[TARGET]',
    'PALETTE': '[STYLE]',
    'CELEBRATION': '[DONE]',
    'BAR_CHART': '[CHART]',
    'MEMO': '[NOTE]',
}

class MemoryManager:
    """Управляет загрузкой и оптимизацией данных в памяти"""

    # Производные таблицы AUSP по значению ATINN
    AUSP_DERIVED_NAMES = ("AUSP_143", "AUSP_604", "AUSP_148", "AUSP_151")
    AUSP_ATINN_TO_COLUMN = {"143": "CCAF", "604": "RED_OUTLET", "148": "ZGLOBAL_CUSTOMER", "151": "ZTRADE_NAME"}

    # DFKKBPTAXNUM: одна таблица в БД — загружаем, фильтруем по taxtype и создаём постоянные таблицы DFKKBPTAXNUM1..6
    DFKKBPTAXNUM_TABLES = ("DFKKBPTAXNUM1", "DFKKBPTAXNUM2", "DFKKBPTAXNUM3", "DFKKBPTAXNUM4", "DFKKBPTAXNUM5", "DFKKBPTAXNUM6")

    # Таблицы, которые при загрузке приводятся к «чистому списку контрагентов»: одна строка на PARTNER (дубли по торговым точкам убираются)
    TABLES_UNIQUE_PARTNER = ("ZBUT0000P3VVI9", "ZBUT0000P", "ZBUT0000P3VV19")

    # Алиасы имён таблиц: логическое имя (в правилах, например /LOT/GC_ADR) -> физическое имя в БД (LOTGC_ADR). Одна таблица — два имени.
    TABLE_NAME_ALIASES = {
        "/LOT/GC_ADR": "LOTGC_ADR",
    }

    def __init__(self, db_path):
        self.db_path = db_path
        self.data_cache = {}
        self.reference_cache = {}
        # Количество уникальных партнёров для таблиц ZBUT0000P3VVI9 и аналогов (для колонки «Всего записей»)
        self._unique_partner_count = {}
        self._ensure_db_integrity()

    def _ensure_db_integrity(self):
        """
        Защита от битой/невалидной sqlite БД.
        Если PRAGMA quick_check вернул ошибку — текущий файл переименовываем в *.corrupted_* и создаём пустой новый.
        """
        try:
            if not self.db_path:
                return
            if not os.path.exists(self.db_path):
                # Файл не существует — создастся при первом connect (оставляем так)
                return

            conn = connect_sqlite(self.db_path)
            try:
                row = conn.execute("PRAGMA quick_check(1);").fetchone()
                res = row[0] if row else None
            finally:
                conn.close()

            if res and str(res).strip().lower() != "ok":
                stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                corrupted_path = f"{self.db_path}.corrupted_{stamp}"
                os.rename(self.db_path, corrupted_path)
                # WAL/SHM
                for suf in ("-wal", "-shm"):
                    p = self.db_path + suf
                    if os.path.exists(p):
                        try:
                            os.rename(p, corrupted_path + suf)
                        except OSError:
                            pass
                connect_sqlite(self.db_path).close()
                print(f"   [WARN] БД {self.db_path} битая ({res}); переименована в {os.path.basename(corrupted_path)} и создана пустая.")
        except Exception:
            # Не ломаем пайплайн из-за проверки целостности — просто пробуем продолжить.
            return

    def _normalize_atinn_value(self, value):
        """Нормализация ATINN для сравнения: 143, 143.0, ' 143 ' -> '143'."""
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return ""
        s = str(value).strip()
        try:
            return str(int(float(s)))
        except (ValueError, TypeError):
            return re.sub(r"\.0+$", "", s) if s else ""

    def _find_ausp_columns(self, columns):
        """Находит колонки ATINN и ATWRT в AUSP по имени или по позиции (2-я и 3-я)."""
        cols = [c for c in columns if c is not None]
        atinn_col = atwrt_col = None
        for c in cols:
            raw = re.sub(r"[^A-Za-z0-9]", "", str(c).upper())
            if raw == "ATINN":
                atinn_col = c
            if raw == "ATWRT":
                atwrt_col = c
        if not atinn_col:
            for c in cols:
                if "ATINN" in str(c).upper():
                    atinn_col = c
                    break
        if not atwrt_col:
            for c in cols:
                if "ATWRT" in str(c).upper():
                    atwrt_col = c
                    break
        if not atinn_col and len(cols) >= 2:
            atinn_col = cols[1]
        if not atwrt_col and len(cols) >= 3:
            atwrt_col = cols[2]
        return atinn_col, atwrt_col

    def _get_ausp_cache_key(self):
        """Возвращает ключ таблицы AUSP в data_cache (без учёта регистра) или None."""
        for k in self.data_cache:
            if str(k).strip().upper() == "AUSP":
                return k
        return None

    def _build_ausp_derived_tables(self):
        """Строит производные таблицы AUSP_143, AUSP_604, AUSP_148, AUSP_151. Не создаёт заново, если уже есть в кэше."""
        ausp_key = self._get_ausp_cache_key()
        if not ausp_key:
            return
        # Если все производные уже в кэше — не пересоздаём
        if all(d in self.data_cache for d in self.AUSP_DERIVED_NAMES):
            return
        df = self.data_cache[ausp_key]
        if df is None or df.empty:
            return
        atinn_col, atwrt_col = self._find_ausp_columns(df.columns)
        if not atinn_col or not atwrt_col:
            print(f"   {TERMINAL_SYMBOLS_LOCAL['WARNING']} AUSP: колонки ATINN/ATWRT не найдены, производные таблицы не созданы")
            return
        total_ausp = len(df)
        # Какие значения ATINN реально есть в данных (нормализованные)
        atinn_normalized = df[atinn_col].apply(self._normalize_atinn_value)
        unique_atinn = sorted(set(atinn_normalized) - {""}, key=lambda x: (len(str(x)), str(x)))
        if unique_atinn:
            sample = ", ".join(str(v) for v in list(unique_atinn)[:15])
            if len(unique_atinn) > 15:
                sample += f", ... (всего {len(unique_atinn)} значений)"
            print(f"   {TERMINAL_SYMBOLS_LOCAL['INFO']} AUSP: всего {total_ausp:,} строк, в колонке ATINN найдены значения: {sample}")
        for atinn_val, col_name in self.AUSP_ATINN_TO_COLUMN.items():
            derived_name = f"AUSP_{atinn_val}"
            if derived_name in self.data_cache:
                continue
            mask = atinn_normalized == atinn_val
            slice_df = df.loc[mask].copy()
            slice_df = slice_df.rename(columns={atwrt_col: col_name})
            self.data_cache[derived_name] = slice_df
            n = len(slice_df)
            if n > 0:
                print(f"   {TERMINAL_SYMBOLS_LOCAL['INFO']} AUSP: создана таблица {derived_name} (ATINN={atinn_val}, колонка {col_name}): {n:,} строк")
            else:
                print(f"   {TERMINAL_SYMBOLS_LOCAL['WARNING']} AUSP: таблица {derived_name} пуста (нет строк с ATINN={atinn_val}); ожидаемые ATINN: 143, 604, 148, 151")

    def _get_dfkkbptaxnum_taxtype_column(self, df):
        """Возвращает имя колонки с типом налога (1,2,3,5) для DFKKBPTAXNUM или None. Сначала conf_dfkkbptaxnum.json, затем эвристика по значениям/имени."""
        if df is None or df.empty or not hasattr(df, 'columns'):
            return None
        cols = list(df.columns)
        # Конфиг
        for root in [os.path.dirname(os.path.dirname(os.path.abspath(__file__))) if __file__ else None, os.getcwd()]:
            if not root:
                continue
            for sub in ('json files', 'config'):
                path = os.path.join(root, sub, 'conf_dfkkbptaxnum.json')
                path = os.path.abspath(path)
                if os.path.isfile(path):
                    try:
                        with open(path, 'r', encoding='utf-8') as f:
                            cfg = json.load(f)
                        for key in ('taxtype_column', 'taxtype_column_name'):
                            col = cfg.get(key)
                            if col and isinstance(col, str) and col.strip():
                                col = col.strip()
                                if col in cols:
                                    return col
                                for c in cols:
                                    if str(c).strip().upper() == col.upper():
                                        return c
                        for c in cfg.get('taxtype_column_alternatives') or []:
                            if c and isinstance(c, str) and c.strip() and c.strip() in cols:
                                return c.strip()
                    except Exception:
                        pass
        # Эвристика: колонка со значениями из {1,2,3,4,5,6}
        allowed = {1, 2, 3, 4, 5, 6}
        for col in cols:
            try:
                ser = df[col].dropna().astype(str).str.strip()
                nums = pd.to_numeric(ser, errors='coerce').dropna()
                if len(nums) == 0:
                    continue
                uniq = set(int(x) for x in nums.unique() if x == int(x))
                if uniq and uniq <= allowed:
                    return col
            except Exception:
                continue
        for col in cols:
            cu = str(col).upper().replace(' ', '').replace('_', '')
            if cu in ('TAXTYPE', 'TAXTYP', 'TAXNUMTYPE', 'TYPE') or 'TAXTYPE' in cu or 'TAXTYP' in cu:
                return col
        return None

    def _get_dfkkbptaxnum_cache_key(self):
        """Возвращает ключ таблицы DFKKBPTAXNUM в data_cache (без учёта регистра) или None."""
        for k in self.data_cache:
            if str(k).strip().upper() == "DFKKBPTAXNUM":
                return k
        return None

    def _build_dfkkbptaxnum_alias_tables(self):
        """Загружает DFKKBPTAXNUM, фильтрует по колонке taxtype (1..6) и создаёт постоянные таблицы в БД:
        DFKKBPTAXNUM1, DFKKBPTAXNUM2, ... DFKKBPTAXNUM6 — каждая с отфильтрованным содержимым по TAXNUM. Колонка «Всего записей» — по каждой таблице отдельно."""
        cache_key = self._get_dfkkbptaxnum_cache_key()
        if not cache_key:
            return
        df = self.data_cache[cache_key]
        if df is None or df.empty:
            for name in self.DFKKBPTAXNUM_TABLES:
                self.data_cache[name] = df
            return
        type_col = self._get_dfkkbptaxnum_taxtype_column(df)
        if not type_col or type_col not in df.columns:
            print(f"   {TERMINAL_SYMBOLS_LOCAL['WARNING']} DFKKBPTAXNUM: колонка типа не найдена — постоянные таблицы не созданы. Укажите taxtype_column в conf_dfkkbptaxnum.json. Колонки: {list(df.columns)[:20]}")
            for name in self.DFKKBPTAXNUM_TABLES:
                self.data_cache[name] = df
            return
        total = len(df)
        try:
            conn = connect_sqlite(self.db_path)
            for table_name in self.DFKKBPTAXNUM_TABLES:
                m = re.match(r'DFKKBPTAXNUM(\d)$', table_name, re.I)
                if not m:
                    self.data_cache[table_name] = df
                    continue
                typ = int(m.group(1))
                type_ser = df[type_col].astype(str).str.strip()
                numeric_type = pd.to_numeric(type_ser, errors='coerce')
                mask = (
                    (numeric_type == typ)
                    | (type_ser == str(typ))
                    | (type_ser.str.upper() == f'RU{typ}')
                )
                slice_df = df.loc[mask].copy()
                n = len(slice_df)
                slice_df.to_sql(table_name, conn, if_exists='replace', index=False)
                self.data_cache[table_name] = slice_df
                if n != total and n > 0:
                    print(f"   {TERMINAL_SYMBOLS_LOCAL['INFO']} DFKKBPTAXNUM: постоянная таблица {table_name} — {n:,} строк (taxtype={typ}), всего в исходной: {total:,}")
            conn.close()
            print(f"   {TERMINAL_SYMBOLS_LOCAL['INFO']} DFKKBPTAXNUM: созданы постоянные таблицы DFKKBPTAXNUM1..6 по полю taxtype, «Всего записей» — по каждой отдельно")
        except Exception as e:
            print(f"   {TERMINAL_SYMBOLS_LOCAL['WARNING']} DFKKBPTAXNUM: не удалось записать постоянные таблицы в БД: {e}")
            for table_name in self.DFKKBPTAXNUM_TABLES:
                m = re.match(r'DFKKBPTAXNUM(\d)$', table_name, re.I)
                if m:
                    typ = int(m.group(1))
                    type_ser = df[type_col].astype(str).str.strip()
                    numeric_type = pd.to_numeric(type_ser, errors='coerce')
                    mask = (numeric_type == typ) | (type_ser == str(typ)) | (type_ser.str.upper() == f'RU{typ}')
                    self.data_cache[table_name] = df.loc[mask].copy()
                else:
                    self.data_cache[table_name] = df

    def _load_partner_column_config(self, table_name):
        """Читает conf_zbut0000p_partner.json и возвращает имя колонки партнёра для дедупликации."""
        if table_name not in self.TABLES_UNIQUE_PARTNER:
            return None
        root = os.path.dirname(os.path.dirname(os.path.abspath(__file__))) if __file__ else os.getcwd()
        for path in [
            os.path.join(root, 'json files', 'conf_zbut0000p_partner.json'),
            os.path.join(os.getcwd(), 'json files', 'conf_zbut0000p_partner.json'),
            os.path.join(root, 'config', 'conf_zbut0000p_partner.json'),
        ]:
            path = os.path.abspath(path)
            if os.path.isfile(path):
                try:
                    with open(path, 'r', encoding='utf-8') as f:
                        cfg = json.load(f)
                    col = cfg.get('partner_column') or cfg.get('partner_column_name')
                    if col and isinstance(col, str) and col.strip():
                        return col.strip()
                except Exception:
                    pass
        return None

    def _find_partner_column_for_dedup(self, table_name, df):
        """Возвращает имя колонки PARTNER для таблиц из TABLES_UNIQUE_PARTNER (для дедупликации в чистый список контрагентов)."""
        if df is None or df.empty or table_name not in self.TABLES_UNIQUE_PARTNER:
            return None
        cols = list(df.columns)
        config_col = self._load_partner_column_config(table_name)
        if config_col:
            for c in cols:
                if str(c).strip().upper() == config_col.strip().upper():
                    return c
            if config_col in cols:
                return config_col
        col_upper = {str(c).strip().upper(): c for c in cols}
        for name in ('PARTNER', 'PARTNERS', 'PARTNER_ID', 'KUNNR', 'BP', 'CUSTOMER', 'CLIENT'):
            if name in col_upper:
                return col_upper[name]
        for cu, col in col_upper.items():
            if 'PARTNER' in cu or 'KUNNR' in cu or cu == 'BP':
                return col
        return None

    def _apply_unique_partner_dedup(self, table_name):
        """Для ZBUT0000P3VVI9 и аналогов: проверки идут по сырым данным (в кэше остаётся сырая таблица).
        Создаётся отдельная постоянная таблица {table_name}_partners с уникальными партнёрами (без дублей) —
        только для подстановки в колонку «Всего записей» количества кастомеров (бизнес-логика)."""
        if table_name not in self.TABLES_UNIQUE_PARTNER or table_name not in self.data_cache:
            return
        df = self.data_cache[table_name]
        if df is None or df.empty:
            return
        partner_col = self._find_partner_column_for_dedup(table_name, df)
        if not partner_col:
            print(f"   {TERMINAL_SYMBOLS_LOCAL['WARNING']} {table_name}: колонка PARTNER не найдена — таблица партнёров не создана. Укажите partner_column в conf_zbut0000p_partner.json. Колонки: {list(df.columns)[:25]}")
            return
        before = len(df)
        df_unique = df.drop_duplicates(subset=[partner_col], keep='first').copy()
        n_after = len(df_unique)
        # В кэше оставляем сырые данные — по ним идут проверки
        # self.data_cache[table_name] не меняем

        # Отдельная постоянная таблица с партнёрами (без дублей) — только для колонки «Всего записей»
        partners_table_name = f"{table_name}_partners"
        try:
            conn = connect_sqlite(self.db_path)
            df_unique.to_sql(partners_table_name, conn, if_exists='replace', index=False)
            conn.close()
            self._unique_partner_count[table_name] = n_after
            if before > n_after:
                print(f"   {TERMINAL_SYMBOLS_LOCAL['INFO']} {table_name}: создана таблица {partners_table_name} для колонки «Всего записей»: {n_after:,} кастомеров (в сырых данных {before:,} строк, дубли по торговым точкам не считаются)")
            else:
                print(f"   {TERMINAL_SYMBOLS_LOCAL['INFO']} {table_name}: создана таблица {partners_table_name}: {n_after:,} кастомеров (дублей не было)")
        except Exception as e:
            print(f"   {TERMINAL_SYMBOLS_LOCAL['WARNING']} {table_name}: не удалось создать таблицу {partners_table_name}: {e}")

    def get_unique_partner_count(self, table_name):
        """Возвращает количество уникальных партнёров (кастомеров) для таблицы ZBUT0000P3VVI9 и аналогов — для колонки «Всего записей». None если таблица не из списка или счёт не вычислялся."""
        return self._unique_partner_count.get(table_name)

    def _get_all_table_names(self):
        """Возвращает список имён всех таблиц в БД"""
        conn = connect_sqlite(self.db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'")
        names = [row[0] for row in cursor.fetchall()]
        conn.close()
        return names
    
    def load_selected_tables_to_ram(self, table_names: list, add_reference_tables: bool = True):
        """Загружает в ОЗУ только указанные таблицы (и при необходимости справочники для правил)."""
        if not table_names:
            self.load_all_data_to_ram()
            return
        all_in_db = self._get_all_table_names()
        to_load = set()
        for t in table_names:
            exact = t if t in all_in_db else None
            if exact is not None:
                to_load.add(exact)
            else:
                match = next((x for x in all_in_db if str(x).strip().upper() == str(t).strip().upper()), None)
                if match:
                    to_load.add(match)
        # Алиасы: при запросе логического имени (например /LOT/GC_ADR) подгружаем физическую таблицу (LOTGC_ADR)
        for logical, physical in self.TABLE_NAME_ALIASES.items():
            if logical in table_names:
                match = next((t for t in all_in_db if str(t).strip().upper() == str(physical).strip().upper()), None)
                if match:
                    to_load.add(match)
        # DFKKBPTAXNUM: при запросе DFKKBPTAXNUM1..6 или DFKKBPTAXNUM — подгружаем исходную таблицу из БД (по имени без учёта регистра)
        if set(table_names).intersection(set(self.DFKKBPTAXNUM_TABLES)) or any(str(t).strip().upper() == "DFKKBPTAXNUM" for t in table_names):
            dfkk_key = next((t for t in all_in_db if str(t).strip().upper() == "DFKKBPTAXNUM"), None)
            if dfkk_key:
                to_load.add(dfkk_key)
        if add_reference_tables:
            for ref in ("T005", "ZW2_CMDEMAND", "BUT020", "KNVV"):
                match = next((t for t in all_in_db if str(t).strip().upper() == ref), None)
                if match:
                    to_load.add(match)
        # RCCOMP_180.1 и др.: BUT0BK/BUT051 джойнятся к KNA1.KUNNR (в GUI — Customer).
        kna1_dependent = {"BUT0BK", "BUT051", "KNB1", "KNVV", "KNVP", "KNVH", "ADR2", "ADRC", "BUT050"}
        kna1_requested = any(str(t).strip().upper() == "KNA1" for t in table_names)
        if kna1_dependent.intersection({str(t).strip().upper() for t in table_names}) or kna1_requested:
            match = next((t for t in all_in_db if str(t).strip().upper() == "KNA1"), None)
            if match:
                to_load.add(match)
        # RCCONF_173.1: дата назначения блока из CDHDR/CDPOS
        if kna1_requested:
            for ref in ("CDHDR", "CDPOS"):
                match = next((t for t in all_in_db if str(t).strip().upper() == ref), None)
                if match:
                    to_load.add(match)
        # ADRC: подгружаем по имени без учёта регистра (в БД может быть "adrc")
        for t in table_names:
            if str(t).strip().upper() == "ADRC" and not any(str(x).strip().upper() == "ADRC" for x in to_load):
                match = next((x for x in all_in_db if str(x).strip().upper() == "ADRC"), None)
                if match:
                    to_load.add(match)
        to_load = sorted(to_load)
        if not to_load:
            print(f"{TERMINAL_SYMBOLS_LOCAL['WARNING']} Нет таблиц для загрузки (проверьте имена). Запрошено: {table_names[:10]}{'...' if len(table_names) > 10 else ''}")
            return
        print(f"{TERMINAL_SYMBOLS_LOCAL['ROCKET']} ЗАГРУЗКА В ОЗУ: {len(to_load)} таблиц")
        loaded_count = 0
        with ThreadPoolExecutor(max_workers=8) as executor:
            futures = {t: executor.submit(self._load_single_table, t) for t in to_load}
            for table_name in to_load:
                try:
                    print(f"   [Ожидание] {table_name}...", end=" ", flush=True)
                    df = futures[table_name].result()
                    print(f"готово", flush=True)
                    loaded_count += 1
                    if df is not None:
                        if str(table_name).strip().upper() == "ADRC" and len(df) > 0:
                            df = self._apply_adrc_reserved_filter(df, table_name)
                        self.data_cache[table_name] = df
                        # Кэшируем ту же таблицу под логическим именем (например /LOT/GC_ADR для LOTGC_ADR)
                        for logical, physical in self.TABLE_NAME_ALIASES.items():
                            if str(physical).strip().upper() == str(table_name).strip().upper():
                                self.data_cache[logical] = df
                                break
                        mem = df.memory_usage(deep=True).sum() / 1024 / 1024
                        n = len(df)
                        sym = TERMINAL_SYMBOLS_LOCAL['SUCCESS'] if n > 0 else TERMINAL_SYMBOLS_LOCAL['WARNING']
                        print(f"   [{loaded_count}/{len(to_load)}] {sym} {table_name}: {n:,} строк, {mem:.1f} MB")
                    else:
                        print(f"   [{loaded_count}/{len(to_load)}] {TERMINAL_SYMBOLS_LOCAL['ERROR']} {table_name}: ошибка загрузки")
                except Exception as e:
                    loaded_count += 1
                    print(f"ошибка", flush=True)
                    print(f"   [{loaded_count}/{len(to_load)}] {TERMINAL_SYMBOLS_LOCAL['ERROR']} {table_name}: {e}")
        print(f"   {TERMINAL_SYMBOLS_LOCAL['INFO']} Постобработка (DFKKBPTAXNUM, дедупликация партнёров)...")
        # DFKKBPTAXNUM: при наличии в кэше (любой регистр) создаём постоянные таблицы DFKKBPTAXNUM1..6
        if self._get_dfkkbptaxnum_cache_key():
            self._build_dfkkbptaxnum_alias_tables()
        # Чистый список контрагентов: таблицы ZBUT0000P3VVI9 и аналоги — одна строка на PARTNER
        for t in self.TABLES_UNIQUE_PARTNER:
            if t in self.data_cache:
                self._apply_unique_partner_dedup(t)
        self._optimize_dataframes()
        total_rows = sum(len(df) for df in self.data_cache.values())
        total_memory = self.get_memory_usage()
        print(f"{TERMINAL_SYMBOLS_LOCAL['CHECKMARK']} ЗАГРУЗКА ЗАВЕРШЕНА!")
        print(f"   {TERMINAL_SYMBOLS_LOCAL['CHART_UP']} Строк в памяти: {total_rows:,}, память: {total_memory:.1f} MB")
    
    def load_all_data_to_ram(self):
        """Загружает все данные из базы в оперативную память"""
        print(f"{TERMINAL_SYMBOLS_LOCAL['ROCKET']} ЗАГРУЗКА ВСЕХ ДАННЫХ В ОЗУ...")
        
        all_tables = self._get_all_table_names()
        print(f"{TERMINAL_SYMBOLS_LOCAL['INFO']} Найдено таблиц: {len(all_tables)}")
        
        # Загружаем параллельно с прогрессом
        loaded_count = 0
        with ThreadPoolExecutor(max_workers=8) as executor:
            futures = {}
            for table_name in all_tables:
                future = executor.submit(self._load_single_table, table_name)
                futures[table_name] = future
            
            for table_name, future in futures.items():
                try:
                    print(f"   [Ожидание] {table_name}...", end=" ", flush=True)
                    df = future.result()
                    print(f"готово", flush=True)
                    loaded_count += 1
                    if df is not None:
                        if str(table_name).strip().upper() == "ADRC" and len(df) > 0:
                            df = self._apply_adrc_reserved_filter(df, table_name)
                        self.data_cache[table_name] = df
                        # Кэшируем ту же таблицу под логическим именем (например /LOT/GC_ADR для LOTGC_ADR)
                        for logical, physical in self.TABLE_NAME_ALIASES.items():
                            if str(physical).strip().upper() == str(table_name).strip().upper():
                                self.data_cache[logical] = df
                                break
                        memory_usage = df.memory_usage(deep=True).sum() / 1024 / 1024
                        row_count = len(df)
                        if row_count > 0:
                            print(f"   [{loaded_count}/{len(all_tables)}] {TERMINAL_SYMBOLS_LOCAL['SUCCESS']} {table_name}: {row_count:,} строк, {memory_usage:.1f} MB")
                        else:
                            print(f"   [{loaded_count}/{len(all_tables)}] {TERMINAL_SYMBOLS_LOCAL['WARNING']} {table_name}: таблица пустая")
                except Exception as e:
                    loaded_count += 1
                    print(f"ошибка", flush=True)
                    print(f"   [{loaded_count}/{len(all_tables)}] {TERMINAL_SYMBOLS_LOCAL['ERROR']} {table_name}: ошибка загрузки - {e}")
        
        print(f"   {TERMINAL_SYMBOLS_LOCAL['INFO']} Все таблицы загружены. Постобработка...")
        # DFKKBPTAXNUM: при наличии в кэше (любой регистр) создаём постоянные таблицы DFKKBPTAXNUM1..6
        if self._get_dfkkbptaxnum_cache_key():
            print(f"   {TERMINAL_SYMBOLS_LOCAL['GEAR']} Создание алиасов DFKKBPTAXNUM...")
            self._build_dfkkbptaxnum_alias_tables()
        # Чистый список контрагентов: таблицы ZBUT0000P3VVI9 и аналоги — одна строка на PARTNER
        if self.TABLES_UNIQUE_PARTNER:
            print(f"   {TERMINAL_SYMBOLS_LOCAL['GEAR']} Дедупликация по партнёрам для: {', '.join(t for t in self.TABLES_UNIQUE_PARTNER if t in self.data_cache)}...")
        for t in self.TABLES_UNIQUE_PARTNER:
            if t in self.data_cache:
                self._apply_unique_partner_dedup(t)
        self._optimize_dataframes()
        
        # Статистика
        total_rows = sum(len(df) for df in self.data_cache.values())
        total_memory = self.get_memory_usage()
        
        print(f"{TERMINAL_SYMBOLS_LOCAL['CHECKMARK']} ЗАГРУЗКА ЗАВЕРШЕНА!")
        print(f"   {TERMINAL_SYMBOLS_LOCAL['CHART_UP']} Всего строк: {total_rows:,}")
        print(f"   {TERMINAL_SYMBOLS_LOCAL['CHART_UP']} Использовано памяти: {total_memory:.1f} MB")
    
    def _resolve_table_name(self, conn, table_name):
        """Возвращает реальное имя таблицы в БД (с учётом регистра и алиасов)."""
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'")
        rows = cursor.fetchall()
        # Алиас: логическое имя из правил (например /LOT/GC_ADR) -> физическое в БД (LOTGC_ADR)
        name_to_resolve = self.TABLE_NAME_ALIASES.get(table_name, table_name)
        want_upper = str(name_to_resolve).strip().upper()
        for (name,) in rows:
            if name and str(name).strip().upper() == want_upper:
                return name
        return table_name

    def _load_single_table(self, table_name):
        """Загружает одну таблицу с оптимизацией. Имя таблицы ищется в БД без учёта регистра (KNVV найдёт knvv)."""
        try:
            conn = connect_sqlite(self.db_path)
            # Реальное имя таблицы в БД (регистр как в БД), иначе запрос к "KNVV" не найдёт таблицу "knvv"
            actual_name = self._resolve_table_name(conn, table_name)
            escaped_table_name = f'"{actual_name}"'
            
            # Сначала проверяем размер таблицы
            cursor = conn.cursor()
            cursor.execute(f'SELECT COUNT(*) FROM {escaped_table_name}')
            row_count = cursor.fetchone()[0]
            
            # Для больших таблиц (>500k строк) используем chunking
            if row_count > 500000:
                print(f"   {TERMINAL_SYMBOLS_LOCAL['INFO']} {table_name}: большая таблица ({row_count:,} строк), загрузка по частям...")
                chunks = []
                chunk_size = 200000
                offset = 0
                while offset < row_count:
                    chunk_df = pd.read_sql_query(
                        f'SELECT * FROM {escaped_table_name} LIMIT {chunk_size} OFFSET {offset}',
                        conn
                    )
                    if len(chunk_df) == 0:
                        break
                    # Оптимизируем типы сразу для каждого chunk
                    chunk_df = self._optimize_dataframe(chunk_df)
                    chunks.append(chunk_df)
                    offset += chunk_size
                df = pd.concat(chunks, ignore_index=True) if chunks else pd.DataFrame()
                del chunks
            else:
                df = pd.read_sql_query(f"SELECT * FROM {escaped_table_name}", conn)
                # Оптимизируем типы сразу при загрузке
                if len(df) > 0:
                    df = self._optimize_dataframe(df)
            
            # Очистка от полных дублей: только строки, совпадающие по всем колонкам (не по одному полю)
            if len(df) > 0:
                before = len(df)
                df = df.drop_duplicates()  # без subset = по всей строке
                if len(df) < before:
                    print(f"   {TERMINAL_SYMBOLS_LOCAL['INFO']} {table_name}: удалено дубликатов {before - len(df):,} (осталось {len(df):,})", flush=True)
            
            # ADRC: при загрузке исключаем строки с NAME1=RESERVED (не участвуют в проверках и подсчётах)
            if df is not None and len(df) > 0 and str(table_name).strip().upper() == "ADRC":
                df = self._apply_adrc_reserved_filter(df, table_name)
            if df is not None and len(df) > 0:
                df = self._drop_fully_empty_rows(df, log_label=actual_name)
            
            conn.close()
            return df
        except Exception as e:
            print(f"   {TERMINAL_SYMBOLS_LOCAL['ERROR']} Ошибка загрузки {table_name}: {e}")
            return None
    
    def _apply_adrc_reserved_filter(self, df, table_name="ADRC"):
        """Исключает из ADRC строки с NAME1=RESERVED. Возвращает отфильтрованный DataFrame (копию)."""
        name1_col = None
        for c in df.columns:
            if str(c).strip().upper() == "NAME1":
                name1_col = c
                break
        if name1_col is None:
            for c in df.columns:
                if "NAME" in str(c).upper() and "1" in str(c):
                    name1_col = c
                    break
        if name1_col is None:
            # Поиск по содержимому: колонка, где чаще всего встречается значение RESERVED
            best_col, best_count = None, 0
            for c in df.columns:
                try:
                    cnt = (df[c].astype(str).str.strip().str.upper() == "RESERVED").sum()
                    if cnt > best_count:
                        best_count = cnt
                        best_col = c
                except Exception:
                    pass
            if best_col and best_count > 0:
                name1_col = best_col
        if name1_col is None or name1_col not in df.columns:
            return df
        before = len(df)
        val_str = df[name1_col].astype(str).str.strip().str.upper()
        df = df[val_str != "RESERVED"].copy()
        dropped = before - len(df)
        if dropped > 0:
            print(f"   {TERMINAL_SYMBOLS_LOCAL['INFO']} ADRC: исключены строки NAME1=RESERVED: {dropped:,} (в кэше {len(df):,})", flush=True)
        return df

    def _drop_fully_empty_rows(self, df, log_label: str | None = None):
        """Убирает строки без единого значения (как пустой хвост в Excel); единая логика с checker/dedupe."""
        if df is None or len(df) == 0 or df.shape[1] == 0:
            return df
        mask = fully_empty_rows_mask(df)
        n = int(mask.sum())
        if n <= 0:
            return df
        label = log_label or "таблица"
        print(
            f"   {TERMINAL_SYMBOLS_LOCAL['INFO']} {label}: исключено полностью пустых строк: {n:,} (осталось {len(df) - n:,})",
            flush=True,
        )
        return df.loc[~mask].copy()

    @staticmethod
    def _downcast_numeric(series, *, kind: str = "float"):
        """downcast без errors='ignore' (deprecated в pandas 2.x). При сбое — исходная серия."""
        try:
            if kind == "integer":
                return pd.to_numeric(series, downcast="integer")
            return pd.to_numeric(series, downcast="float")
        except (ValueError, TypeError):
            return series

    def _optimize_dataframe(self, df):
        """Оптимизирует DataFrame: типы данных, категории, downcast чисел. Подавляем RuntimeWarning при cast NaN→int."""
        if len(df) == 0:
            return df
        
        import gc
        import warnings
        
        with warnings.catch_warnings():
            warnings.filterwarnings("ignore", "invalid value encountered in cast", RuntimeWarning)
            
            for col in df.columns:
                col_type = df[col].dtype
                
                # Оптимизация числовых типов (не downcast в int при наличии NaN)
                if col_type in ['int64', 'int32']:
                    ser = df[col]
                    if ser.notna().all():
                        df[col] = self._downcast_numeric(ser, kind="integer")
                elif col_type in ['float64', 'float32']:
                    df[col] = self._downcast_numeric(df[col], kind="float")
                
                # Оптимизация строковых типов
                elif col_type == 'object':
                    unique_ratio = df[col].nunique() / len(df)
                    if unique_ratio < 0.5:
                        try:
                            df[col] = df[col].astype('category')
                        except Exception:
                            pass
                    elif df[col].dtype == 'object':
                        try:
                            numeric = pd.to_numeric(df[col], errors='coerce', downcast='integer')
                            filled = numeric.notna().sum() / len(df)
                            if filled > 0.9:
                                if numeric.isna().any():
                                    df[col] = pd.to_numeric(df[col], errors='coerce', downcast='float')
                                else:
                                    df[col] = numeric
                            elif filled > 0.5:
                                df[col] = pd.to_numeric(df[col], errors='coerce', downcast='float')
                        except Exception:
                            pass
        
        gc.collect()
        return df
    
    def _optimize_dataframes(self):
        """Оптимизирует все DataFrame в кэше (уже оптимизированы при загрузке, но можно дооптимизировать)"""
        print(f"{TERMINAL_SYMBOLS_LOCAL['GEAR']} Финальная оптимизация памяти...")
        
        for table_name, df in self.data_cache.items():
            if len(df) == 0:
                continue
            
            # Принудительная сборка мусора для больших таблиц
            if len(df) > 100000:
                gc.collect()
    
    def get_memory_usage(self):
        """Возвращает общее использование памяти"""
        total_memory = 0
        for df in self.data_cache.values():
            total_memory += df.memory_usage(deep=True).sum() / 1024 / 1024
        return total_memory
    
    def get_table(self, table_name):
        """Получить таблицу из кэша. Учитывает алиасы имён (например /LOT/GC_ADR -> LOTGC_ADR).
        Для ADRC: при первом обращении удаляем из кэша строки с NAME1=RESERVED и храним только «чистую» таблицу."""
        requested_upper = str(table_name or "").strip().upper()
        df = None
        cache_key = None
        if table_name in self.data_cache:
            df = self.data_cache[table_name]
            cache_key = table_name
        else:
            physical = self.TABLE_NAME_ALIASES.get(table_name)
            if physical and physical in self.data_cache:
                df = self.data_cache[physical]
                cache_key = physical
        if df is None:
            for k in self.data_cache:
                if str(k).strip().upper() == requested_upper:
                    df = self.data_cache[k]
                    cache_key = k
                    break
        if df is not None and not df.empty and requested_upper == "ADRC":
            df = self._apply_adrc_reserved_filter(df, "ADRC")
            if cache_key is not None:
                self.data_cache[cache_key] = df
        if df is not None:
            return df
        return None
    
    def table_exists(self, table_name):
        """Проверить наличие таблицы в кэше (включая алиасы)."""
        if table_name in self.data_cache:
            return True
        physical = self.TABLE_NAME_ALIASES.get(table_name)
        return physical is not None and physical in self.data_cache

    # --- Async загрузка (ускорение I/O при многих таблицах) ---
    
    async def _load_single_table_async(self, table_name):
        """Асинхронно загружает одну таблицу. Возвращает (table_name, df или None)."""
        if not _HAS_AIOSQLITE:
            return table_name, None
        try:
            async with aiosqlite.connect(self.db_path) as conn:
                escaped = f'"{table_name}"'
                async with conn.execute(f'SELECT COUNT(*) FROM {escaped}') as cur:
                    row = await cur.fetchone()
                    row_count = row[0] if row else 0
                if row_count > 500000:
                    chunks = []
                    chunk_size = 200000
                    offset = 0
                    while offset < row_count:
                        async with conn.execute(
                            f'SELECT * FROM {escaped} LIMIT {chunk_size} OFFSET {offset}'
                        ) as cur:
                            rows = await cur.fetchall()
                            if not rows:
                                break
                            col_names = [d[0] for d in cur.description]
                            chunk_df = pd.DataFrame(rows, columns=col_names)
                        chunk_df = self._optimize_dataframe(chunk_df)
                        chunks.append(chunk_df)
                        offset += chunk_size
                    df = pd.concat(chunks, ignore_index=True) if chunks else pd.DataFrame()
                else:
                    async with conn.execute(f'SELECT * FROM {escaped}') as cur:
                        rows = await cur.fetchall()
                        col_names = [d[0] for d in cur.description]
                        df = pd.DataFrame(rows, columns=col_names)
                    if len(df) > 0:
                        df = self._optimize_dataframe(df)
                if len(df) > 0:
                    before = len(df)
                    df = df.drop_duplicates()
                    if len(df) < before and (before - len(df)) > 0:
                        print(f"   {TERMINAL_SYMBOLS_LOCAL['INFO']} {table_name}: удалено дубликатов {before - len(df):,}")
                if df is not None and len(df) > 0:
                    df = self._drop_fully_empty_rows(df, log_label=table_name)
                return table_name, df
        except Exception as e:
            print(f"   {TERMINAL_SYMBOLS_LOCAL['ERROR']} {table_name}: {e}")
            return table_name, None

    async def load_selected_tables_to_ram_async(self, table_names: list, add_reference_tables: bool = True):
        """Асинхронно загружает указанные таблицы в ОЗУ. Исходный вариант DFKKBPTAXNUM: одна таблица, алиасы RU1..RU6."""
        if not _HAS_AIOSQLITE:
            print(f"   {TERMINAL_SYMBOLS_LOCAL['WARNING']} aiosqlite не установлен. Используется синхронная загрузка.")
            self.load_selected_tables_to_ram(table_names, add_reference_tables)
            return
        all_in_db = self._get_all_table_names()
        to_load = set(table_names) & set(all_in_db)
        if set(table_names).intersection(set(self.DFKKBPTAXNUM_TABLES)) or "DFKKBPTAXNUM" in set(table_names):
            if "DFKKBPTAXNUM" in all_in_db:
                to_load.add("DFKKBPTAXNUM")
        for t in table_names:
            if str(t).strip().upper() == "ADRC" and not any(str(x).strip().upper() == "ADRC" for x in to_load):
                match = next((x for x in all_in_db if str(x).strip().upper() == "ADRC"), None)
                if match:
                    to_load.add(match)
        if add_reference_tables:
            for ref in ("T005", "ZW2_CMDEMAND", "BUT020", "KNVV"):
                if ref in all_in_db:
                    to_load.add(ref)
        kna1_requested = any(str(t).strip().upper() == "KNA1" for t in table_names)
        if kna1_requested:
            for ref in ("KNA1", "CDHDR", "CDPOS"):
                match = next((t for t in all_in_db if str(t).strip().upper() == ref), None)
                if match:
                    to_load.add(match)
        to_load = sorted(to_load)
        if not to_load:
            print(f"   {TERMINAL_SYMBOLS_LOCAL['WARNING']} Нет таблиц для загрузки")
            return
        print(f"   {TERMINAL_SYMBOLS_LOCAL['INFO']} Async загрузка: {len(to_load)} таблиц")
        tasks = [self._load_single_table_async(t) for t in to_load]
        results = await asyncio.gather(*tasks, return_exceptions=True)
        for i, res in enumerate(results):
            if isinstance(res, Exception):
                print(f"   {TERMINAL_SYMBOLS_LOCAL['ERROR']} {to_load[i]}: {res}")
                continue
            tname, df = res
            if df is not None:
                if str(tname).strip().upper() == "ADRC" and len(df) > 0:
                    df = self._apply_adrc_reserved_filter(df, tname)
                self.data_cache[tname] = df
                row_count = len(df)
                mem = df.memory_usage(deep=True).sum() / 1024 / 1024
                if row_count > 0:
                    print(f"   {TERMINAL_SYMBOLS_LOCAL['SUCCESS']} {tname}: {row_count:,} строк, {mem:.1f} MB")
        if "DFKKBPTAXNUM" in self.data_cache:
            self._build_dfkkbptaxnum_alias_tables()
        for t in self.TABLES_UNIQUE_PARTNER:
            if t in self.data_cache:
                self._apply_unique_partner_dedup(t)
        self._optimize_dataframes()
        total_rows = sum(len(d) for d in self.data_cache.values())
        total_mem = self.get_memory_usage()
        print(f"{TERMINAL_SYMBOLS_LOCAL['CHECKMARK']} Async загрузка завершена: {total_rows:,} строк, {total_mem:.1f} MB")

    def load_selected_tables_to_ram_async_sync(self, table_names: list, add_reference_tables: bool = True):
        """Синхронная обёртка для async загрузки (для вызова из main)."""
        if not _HAS_AIOSQLITE:
            self.load_selected_tables_to_ram(table_names, add_reference_tables)
            return
        asyncio.run(self.load_selected_tables_to_ram_async(table_names, add_reference_tables))