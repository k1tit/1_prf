"""
Обработчик правил KNA1 — переписан с нуля.
Реализует ровно логику из rules.json без лишней фильтрации.
"""

import json
import os
import logging
from datetime import datetime
from typing import Dict, Any, Optional, List, Tuple

import pandas as pd


# Имена колонок в KNA1 (SAP / маппинг)
COL_KTOKD = "KTOKD"           # account group (9038 = оцениваем по правилам 103.1, 108.1, 109.1)
COL_KUNNR = "KUNNR"
COL_KATR1 = "KATR1"           # customer_activity_cluster_code
COL_KATR6 = "KATR6"           # trade_channel_code
COL_KATR7 = "KATR7"           # sub_trade_channel_code
COL_BRAN1 = "BRAN1"           # industry_code1
COL_HZUOR = "HZUOR"           # assignment_hierarchy_level
COL_KATR4 = "KATR4"           # distribution_type_code
COL_KUKLA = "KUKLA"           # customer_classification_code
COL_AUFSD = "AUFSD"           # central_order_block_code (F, TS, R)


def _find_col(df: pd.DataFrame, *candidates: str) -> Optional[str]:
    """
    Возвращает первую колонку из df, имя которой совпадает с candidates.

    Важно: эта функция намеренно НЕ делает "размытых" (substring) совпадений,
    потому что это приводило к выбору неправильных колонок (например PSTLZ вместо KUNNR).
    """
    if df is None or df.empty:
        return None

    # 1) Точное совпадение (с учётом регистра)
    cols_set = set(df.columns)
    for cand in candidates:
        if cand in cols_set:
            return cand

    # 2) Точное совпадение без учёта регистра
    upper = {str(c).strip().upper(): c for c in df.columns}
    for cand in candidates:
        cu = cand.strip().upper()
        if cu in upper:
            return upper[cu]
    return None


def _empty_series(ser: pd.Series) -> pd.Series:
    """True где значение считается пустым (NULL, пустая строка, 'NULL', 'NONE')."""
    s = ser.astype(str).str.strip()
    return ser.isna() | (s == "") | (s.str.upper().isin(["NULL", "NONE", "NAN", "NA"]))


class KNA1Handler:
    """Единственная точка обработки правил KNA1. Логика строго по technical_definition_RU."""

    def __init__(self, table_name: str, df: pd.DataFrame, memory_manager, checker):
        self.table_name = table_name
        self.df = df.copy() if df is not None and not df.empty else pd.DataFrame()
        self.memory_manager = memory_manager
        self.checker = checker
        self.logger = logging.getLogger("KNA1Handler")
        self._column_map = self._load_column_map()
        self._conf_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        # KUNNR — критичный ключ для join'ов: берём строго KUNNR (если есть),
        # иначе пытаемся найти по точным alias'ам. Фаззи-матч здесь запрещён.
        if "KUNNR" in self.df.columns:
            self._kunnr_col = "KUNNR"
        else:
            self._kunnr_col = _find_col(self.df, "KUNNR_KNA1", "Customer") or COL_KUNNR

    def _load_column_map(self) -> Dict[str, str]:
        """Маппинг логическое_имя -> физическая колонка для KNA1."""
        for root in (os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "config"),
                    os.path.dirname(os.path.dirname(os.path.abspath(__file__)))):
            path = os.path.join(root, "column_map.json") if "config" in root else os.path.join(root, "json files", "column_map.json")
            if not os.path.exists(path):
                path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "json files", "column_map.json")
            if os.path.exists(path):
                try:
                    with open(path, "r", encoding="utf-8") as f:
                        data = json.load(f)
                    if isinstance(data, dict) and "KNA1" in data:
                        return data["KNA1"]
                except Exception:
                    pass
        return {}

    def _col(self, logical: str, fallback_physical: str) -> Optional[str]:
        """Имя колонки в self.df: из маппинга или fallback (например KATR1)."""
        if self._column_map and logical in self._column_map:
            c = self._column_map[logical]
            if c in self.df.columns:
                return c
        return _find_col(self.df, fallback_physical, logical)

    def _result(self, rule: Dict, total: int, failed: int, error_df: Optional[pd.DataFrame],
                column_checked: str, actual_column: str) -> Dict[str, Any]:
        """Единый формат ответа для checker с защитой от некорректных типов и деления на ноль."""
        # Базовые расчёты с приведением типов
        try:
            total_int = int(total) if total is not None else 0
            failed_int = int(failed) if failed is not None else 0
        except (TypeError, ValueError) as e:
            # Логируем и обнуляем счётчики, чтобы не падать
            self.logger.error(
                "ERROR in _result calculation: %s, total=%r, failed=%r",
                e,
                total,
                failed,
            )
            total_int = 0
            failed_int = 0

        passed = max(total_int - failed_int, 0)
        success_rate = round(passed / total_int * 100, 2) if total_int > 0 else 0

        status = (
            "УСПЕШНО"
            if failed_int == 0
            else (
                "ПОДОЗРИТЕЛЬНО"
                if self.checker._check_if_suspicious(
                    rule.get("rule_code", ""),
                    failed_int,
                    total_int,
                )
                else "ОШИБКИ"
            )
        )

        # Отладочный лог по расчётам
        self.logger.info(
            "DEBUG _result: rule=%s, total=%s, failed=%s, passed=%s, success_rate=%s%%, status=%s",
            rule.get("rule_code", "UNKNOWN"),
            total_int,
            failed_int,
            passed,
            success_rate,
            status,
        )

        return {
            "rule_code": rule.get("rule_code", "UNKNOWN"),
            "rule_description": rule.get("rule_description", ""),
            "quality_category": rule.get("quality_category", ""),
            "table_name": self.table_name,
            "column_checked": column_checked,
            "matched_column": actual_column,
            "actual_column": actual_column,
            "standard_column": rule.get("column_name_checked", ""),
            "total_records": total_int,
            "passed": passed,
            "failed": failed_int,
            "success_rate_%": success_rate,
            "execution_time_sec": 0,
            "check_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "status": status,
            "status_color": "green" if failed_int == 0 else ("orange" if "ПОДОЗРИТЕЛЬНО" in status else "red"),
            "error_file": "Есть" if failed_int > 0 else "Нет",
            "comments": "",
            "error_count": failed_int,
            "error_df": error_df,
        }

    def _error_result(self, rule: Dict, message: str) -> Dict[str, Any]:
        return {
            "rule_code": rule.get("rule_code", "UNKNOWN"),
            "rule_description": rule.get("rule_description", ""),
            "quality_category": rule.get("quality_category", ""),
            "table_name": self.table_name,
            "column_checked": rule.get("column_name_checked", ""),
            "matched_column": "",
            "actual_column": "",
            "standard_column": "",
            "total_records": 0,
            "passed": 0,
            "failed": 0,
            "success_rate_%": 0,
            "execution_time_sec": 0,
            "check_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "status": "ОШИБКА ВЫПОЛНЕНИЯ",
            "status_color": "dark_red",
            "error_file": "Нет",
            "comments": message,
            "error_count": 0,
            "error_df": None,
        }

    def _error_df_sample(self, df: pd.DataFrame, mask: pd.Series, columns: List[str], msg: str,
                         max_rows: int = 500) -> Optional[pd.DataFrame]:
        """Фрагмент df по mask с колонками columns и сообщением."""
        if mask.sum() == 0:
            return None
        cols = [c for c in columns if c in df.columns]
        if not cols:
            cols = list(df.columns)[:5]
        out = df.loc[mask, cols].head(max_rows).copy()
        out["error_message"] = msg
        out["row_id"] = out.index.astype(int) + 1
        return out

    def validate_rule(self, rule: Dict[str, Any]) -> Dict[str, Any]:
        code = (rule.get("rule_code") or "").strip()
        if code == "RCCOMP_103.1":
            return self._rule_103_1(rule)
        if code == "RCCOMP_108.1":
            return self._rule_108_1(rule)
        if code == "RCCOMP_109.1":
            return self._rule_109_1(rule)
        if code == "RCCONF_70.1":
            return self._rule_rcconf_70_1(rule)
        if code == "RCCOMP_187.1":
            return self._rule_187_1(rule)
        if code == "RCCONF_103.4":
            return self._rule_rcconf_103_4(rule)
        if code == "RCCOMP_70.1":
            return self._rule_70_1(rule)
        if code == "RCCOMP_106.1":
            return self._rule_106_1(rule)
        if code == "RCCOMP_68.1":
            return self._rule_68_1(rule)
        if code == "RCCONF_173.1":
            return self._rule_173_1(rule)
        return self._error_result(rule, f"Неизвестное правило: {code}")

    # ---- RCCOMP_103.1, 108.1, 109.1: IF account_group_code != '9038' THEN '' ELSE IF column IS NULL THEN '0' ELSE '1' ----
    def _rule_103_1(self, rule: Dict) -> Dict:
        original_df = self.df.copy()
        try:
            return self._rule_9038_completeness(rule, "customer_activity_cluster_code", COL_KATR1)
        finally:
            self.df = original_df

    def _rule_108_1(self, rule: Dict) -> Dict:
        original_df = self.df.copy()
        try:
            return self._rule_9038_completeness(rule, "trade_channel_code", COL_KATR6)
        finally:
            self.df = original_df

    def _rule_109_1(self, rule: Dict) -> Dict:
        original_df = self.df.copy()
        try:
            return self._rule_9038_completeness(rule, "sub_trade_channel_code", COL_KATR7)
        finally:
            self.df = original_df

    def _rule_9038_completeness(self, rule: Dict, logical_name: str, fallback_col: str) -> Dict:
        """Оцениваем только строки с account_group_code == '9038'; среди них ошибка = пустое значение колонки."""
        ktokd_col = _find_col(self.df, COL_KTOKD, "KTOKD", "account_group_code")
        if not ktokd_col or ktokd_col not in self.df.columns:
            return self._error_result(rule, "В KNA1 не найдена колонка группы счетов (KTOKD)")
        col = self._col(logical_name, fallback_col)
        if not col or col not in self.df.columns:
            return self._error_result(rule, f"Колонка для проверки ({fallback_col}) не найдена")
        ktokd_str = self.df[ktokd_col].astype(str).str.strip()
        mask_9038 = ktokd_str == "9038"
        df_eval = self.df[mask_9038]
        total = len(df_eval)
        if total == 0:
            return self._result(rule, 0, 0, None, rule.get("column_name_checked", fallback_col), col)
        empty = _empty_series(df_eval[col])
        failed = int(empty.sum())
        err_df = None
        if failed > 0:
            idx_err = df_eval.index[empty]
            cols_err = [c for c in [self._kunnr_col, ktokd_col, col] if c in self.df.columns]
            err_df = self.df.loc[idx_err, cols_err].head(getattr(self.checker, "MAX_ERRORS_TO_SAVE", 500)).copy()
            err_df["error_message"] = "Пустое значение (оценка только для группы 9038)"
            err_df["row_id"] = err_df.index.astype(int) + 1
        return self._result(rule, total, failed, err_df, rule.get("column_name_checked", fallback_col), col)

    # ---- RCCONF_70.1: IF industry_code1 IS NULL THEN '' ELSE IF industry_code1_map_flag = '1' THEN '1' ELSE '0' ----
    def _rule_rcconf_70_1(self, rule: Dict) -> Dict:
        """Проверка по справочнику ZW2_CMDEMAND: только непустой BRAN1 должен быть в справочнике."""
        original_df = self.df.copy()
        try:
            col = self._col("industry_code1", COL_BRAN1)
            if not col or col not in self.df.columns:
                return self._error_result(rule, "Колонка BRAN1 не найдена")
            ref_df = self.memory_manager.get_table("ZW2_CMDEMAND")
            if ref_df is None or ref_df.empty:
                return self._error_result(rule, "Справочник ZW2_CMDEMAND не загружен")
            ref_df = ref_df.copy(deep=True)
            ref_col = _find_col(ref_df, "BRAN1", "INDUSTRY_CODE1", "SUBDEMAND")
            if ref_col is None and len(ref_df.columns):
                ref_col = ref_df.columns[0]
            if ref_col is None:
                return self._error_result(rule, "В ZW2_CMDEMAND не найдена колонка кодов")
            valid = set(ref_df[ref_col].dropna().astype(str).str.strip().str.upper())
            non_null = ~_empty_series(self.df[col])
            df_eval = self.df[non_null]
            total = len(df_eval)
            if total == 0:
                return self._result(rule, 0, 0, None, rule.get("column_name_checked", ""), col)
            wrong = ~df_eval[col].astype(str).str.strip().str.upper().isin(valid)
            failed = int(wrong.sum())
            err_df = None
            if failed:
                mask_err = non_null & ~self.df[col].astype(str).str.strip().str.upper().isin(valid)
                err_df = self._error_df_sample(self.df, mask_err, [self._kunnr_col, col], "Код не найден в справочнике ZW2_CMDEMAND", getattr(self.checker, "MAX_ERRORS_TO_SAVE", 500))
            return self._result(rule, total, failed, err_df, rule.get("column_name_checked", ""), col)
        finally:
            self.df = original_df

    # ---- RCCOMP_187.1: IF assignment_hierrarchy_level IS NULL THEN '0' ELSE '1' ----
    def _rule_187_1(self, rule: Dict) -> Dict:
        original_df = self.df.copy()
        try:
            col = self._col("assignment_hierrarchy_level", COL_HZUOR)
            if not col or col not in self.df.columns:
                return self._error_result(rule, "Колонка HZUOR не найдена")
            total = len(self.df)
            if total == 0:
                return self._result(rule, 0, 0, None, rule.get("column_name_checked", ""), col)
            empty = _empty_series(self.df[col])
            failed = int(empty.sum())
            err_df = self._error_df_sample(self.df, empty, [self._kunnr_col, col], "Пустое значение", getattr(self.checker, "MAX_ERRORS_TO_SAVE", 500))
            return self._result(rule, total, failed, err_df, rule.get("column_name_checked", ""), col)
        finally:
            self.df = original_df

    # ---- RCCONF_103.4: все три KATR1,KATR6,KATR7 не пусты -> комбинация должна быть в conf_cac_tc_stc_mapping ----
    def _rule_rcconf_103_4(self, rule: Dict) -> Dict:
        original_df = self.df.copy()
        try:
            c1 = self._col("customer_activity_cluster_code", COL_KATR1)
            c6 = self._col("trade_channel_code", COL_KATR6)
            c7 = self._col("sub_trade_channel_code", COL_KATR7)
            for name, cx in (("KATR1", c1), ("KATR6", c6), ("KATR7", c7)):
                if not cx or cx not in self.df.columns:
                    return self._error_result(rule, f"Колонка {name} не найдена")
            path = os.path.join(self._conf_dir, "json files", "conf_cac_tc_stc_mapping.json")
            if not os.path.exists(path):
                return self._error_result(rule, "Файл conf_cac_tc_stc_mapping.json не найден")
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            mapping_list = data.get("conf_cac_tc_stc_mapping", data) if isinstance(data, dict) else data
            if not isinstance(mapping_list, list):
                return self._error_result(rule, "Неверный формат conf_cac_tc_stc_mapping.json")
            keys_logical = ("customer_activity_cluster_code", "trade_channel_code", "sub_trade_channel_code")
            valid_triples = set()
            for row in mapping_list:
                if isinstance(row, dict):
                    t = tuple(str(row.get(k, "")).strip() for k in keys_logical)
                    if all(t):
                        valid_triples.add(t)
            non_null = ~(_empty_series(self.df[c1]) | _empty_series(self.df[c6]) | _empty_series(self.df[c7]))
            df_eval = self.df[non_null]
            total = len(df_eval)
            if total == 0:
                return self._result(rule, 0, 0, None, rule.get("column_name_checked", ""), c1)
            triples = list(zip(
                df_eval[c1].astype(str).str.strip(),
                df_eval[c6].astype(str).str.strip(),
                df_eval[c7].astype(str).str.strip(),
            ))
            wrong = [i for i, t in enumerate(triples) if t not in valid_triples]
            failed = len(wrong)
            err_df = None
            if failed:
                idx_err = df_eval.iloc[wrong].index
                err_df = self.df.loc[idx_err, [self._kunnr_col, c1, c6, c7]].head(getattr(self.checker, "MAX_ERRORS_TO_SAVE", 500)).copy()
                err_df["error_message"] = "Комбинация CAC+TC+STC не найдена в conf_cac_tc_stc_mapping"
                err_df["row_id"] = err_df.index.astype(int) + 1
            return self._result(rule, total, failed, err_df, rule.get("column_name_checked", ""), c1)
        finally:
            self.df = original_df

    # ---- RCCOMP_70.1: IF industry_code1 IS NULL THEN '0' ELSE '1' ----
    def _rule_70_1(self, rule: Dict) -> Dict:
        original_df = self.df.copy()
        try:
            col = self._col("industry_code1", COL_BRAN1)
            if not col or col not in self.df.columns:
                return self._error_result(rule, "Колонка BRAN1 не найдена")
            total = len(self.df)
            if total == 0:
                return self._result(rule, 0, 0, None, rule.get("column_name_checked", ""), col)
            empty = _empty_series(self.df[col])
            failed = int(empty.sum())
            err_df = self._error_df_sample(self.df, empty, [self._kunnr_col, col], "Пустое значение", getattr(self.checker, "MAX_ERRORS_TO_SAVE", 500))
            return self._result(rule, total, failed, err_df, rule.get("column_name_checked", ""), col)
        finally:
            self.df = original_df

    # ---- RCCOMP_106.1: IF distribution_type_code IS NULL THEN '0' ELSE '1' ----
    def _rule_106_1(self, rule: Dict) -> Dict:
        original_df = self.df.copy()
        try:
            col = self._col("distribution_type_code", COL_KATR4)
            if not col or col not in self.df.columns:
                return self._error_result(rule, "Колонка KATR4 не найдена")
            total = len(self.df)
            if total == 0:
                return self._result(rule, 0, 0, None, rule.get("column_name_checked", ""), col)
            empty = _empty_series(self.df[col])
            failed = int(empty.sum())
            err_df = self._error_df_sample(self.df, empty, [self._kunnr_col, col], "Пустое значение", getattr(self.checker, "MAX_ERRORS_TO_SAVE", 500))
            return self._result(rule, total, failed, err_df, rule.get("column_name_checked", ""), col)
        finally:
            self.df = original_df

    # ---- RCCOMP_68.1: IF customer_classification_code IS NULL THEN '0' ELSE '1' ----
    def _rule_68_1(self, rule: Dict) -> Dict:
        original_df = self.df.copy()
        try:
            col = self._col("customer_classification_code", COL_KUKLA)
            if not col or col not in self.df.columns:
                return self._error_result(rule, "Колонка KUKLA не найдена")
            total = len(self.df)
            if total == 0:
                return self._result(rule, 0, 0, None, rule.get("column_name_checked", ""), col)
            empty = _empty_series(self.df[col])
            failed = int(empty.sum())
            err_df = self._error_df_sample(self.df, empty, [self._kunnr_col, col], "Пустое значение", getattr(self.checker, "MAX_ERRORS_TO_SAVE", 500))
            return self._result(rule, total, failed, err_df, rule.get("column_name_checked", ""), col)
        finally:
            self.df = original_df

    # ---- RCCONF_173.1: Order Block Assignment Date ----
    # Источник блока: KNA1.central_order_block_code (AUFSD)
    # Источник даты назначения блока: CDHDR+CDPOS (max UDATE+UTIME по клиенту для изменений KNA1.AUFSD)
    def _rule_173_1(self, rule: Dict) -> Dict:
        """
        RCCONF_173.1: Order Block Assignment Date
        IF b.central_order_block_code IS NULL OR a.central_order_block_assignment_date IS NULL THEN ''
        ELSE IF months_since_date <= b.max_months THEN '1' ELSE '0'
        """
        if self._kunnr_col not in self.df.columns:
            return self._error_result(rule, "В KNA1 не найдена колонка KUNNR")

        # --- 1. Подготовка KNA1 ---
        df_kna1 = self.df.copy(deep=True)
        df_kna1['KUNNR_clean'] = (
            df_kna1[self._kunnr_col]
            .astype(str)
            .str.strip()
            .str.replace(r"\.0$", "", regex=True)
            .str.replace(r"\D+", "", regex=True)
            .str.zfill(10)
        )
        block_col = self._col("central_order_block_code", COL_AUFSD) or _find_col(df_kna1, COL_AUFSD, "AUFSD")
        if not block_col or block_col not in df_kna1.columns:
            return self._error_result(rule, "В KNA1 отсутствует central_order_block_code (AUFSD)")
        df_kna1["BLOCK_CODE"] = df_kna1[block_col].astype(str).str.strip().str.upper()

        # --- 2. Загружаем конфиг с лимитами ---
        path = os.path.join(self._conf_dir, "json files", "conf_order_block_time.json")
        if not os.path.exists(path):
            return self._error_result(rule, "Файл conf_order_block_time.json не найден")

        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)

        block_conf = data.get("conf_order_block_time", data) if isinstance(data, dict) else data
        max_months = {}
        if isinstance(block_conf, list):
            for item in block_conf:
                if isinstance(item, dict):
                    code = str(item.get("central_order_block_code", "")).strip().upper()
                    try:
                        months = int(item.get("max_months", 0))
                    except (ValueError, TypeError):
                        months = 0
                    if code and months > 0:
                        max_months[code] = months
        if not max_months:
            return self._error_result(rule, "conf_order_block_time.json пустой или некорректный")

        # --- 3. Получаем дату назначения блока ---
        # Приоритет: CDHDR+CDPOS (как в техническом описании), fallback: колонка даты в KNA1.
        assignment_dt_by_customer = None
        assignment_date_source = None  # 'CDHDR' | 'KNA1' — для подписей в файле ошибок
        cdhdr_df = self.memory_manager.get_table("CDHDR")
        cdpos_df = self.memory_manager.get_table("CDPOS")
        if (cdhdr_df is not None and not cdhdr_df.empty) and (cdpos_df is not None and not cdpos_df.empty):
            try:
                cp = cdpos_df.copy(deep=True)
                ch = cdhdr_df.copy(deep=True)

                cp_objectclas = _find_col(cp, "OBJECTCLAS")
                cp_tabname = _find_col(cp, "TABNAME")
                cp_fname = _find_col(cp, "FNAME")
                cp_chngind = _find_col(cp, "CHNGIND")
                cp_obj = _find_col(cp, "OBJECTID", "TABKEY", "OBJECT_ID")
                cp_changenr = _find_col(cp, "CHANGENR", "CHANGENR_CDHDR", "CHANGE_NR")

                ch_changenr = _find_col(ch, "CHANGENR", "CHANGE_NR")
                ch_udate = _find_col(ch, "UDATE", "CHANGE_DATE")
                ch_utime = _find_col(ch, "UTIME", "CHANGE_TIME")

                if cp_obj and cp_changenr and ch_changenr and ch_udate and ch_utime:
                    cp_fil = cp.copy()
                    if cp_objectclas:
                        cp_fil = cp_fil[cp_fil[cp_objectclas].astype(str).str.strip().str.upper() == "DEBI"]
                    if cp_tabname:
                        cp_fil = cp_fil[cp_fil[cp_tabname].astype(str).str.strip().str.upper() == "KNA1"]
                    if cp_fname:
                        cp_fil = cp_fil[cp_fil[cp_fname].astype(str).str.strip().str.upper() == "AUFSD"]
                    if cp_chngind:
                        cp_fil = cp_fil[cp_fil[cp_chngind].astype(str).str.strip().str.upper().isin(["U", "I"])]

                    if not cp_fil.empty:
                        cp_fil["KUNNR_clean"] = (
                            cp_fil[cp_obj]
                            .astype(str)
                            .str.strip()
                            .str.replace(r"\.0$", "", regex=True)
                            .str.replace(r"\D+", "", regex=True)
                            .str[-10:]
                            .str.zfill(10)
                        )
                        cp_fil = cp_fil[cp_fil["KUNNR_clean"].str.match(r"^\d{10}$", na=False)]
                        cp_fil = cp_fil[[cp_changenr, "KUNNR_clean"]].dropna().drop_duplicates()

                        ch_key = ch[[ch_changenr, ch_udate, ch_utime]].dropna(subset=[ch_changenr]).copy()
                        ch_key["ASSIGN_DT"] = pd.to_datetime(
                            ch_key[ch_udate].astype(str).str.strip() + " " + ch_key[ch_utime].astype(str).str.zfill(6),
                            format="%Y%m%d %H%M%S",
                            errors="coerce",
                        )
                        merged_hist = cp_fil.merge(
                            ch_key[[ch_changenr, "ASSIGN_DT"]],
                            left_on=cp_changenr,
                            right_on=ch_changenr,
                            how="left",
                        )
                        merged_hist = merged_hist.dropna(subset=["ASSIGN_DT"])
                        if not merged_hist.empty:
                            assignment_dt_by_customer = (
                                merged_hist.groupby("KUNNR_clean", as_index=False)["ASSIGN_DT"].max()
                            )
                            assignment_date_source = "CDHDR"
            except Exception as e:
                self.logger.warning("RCCONF_173.1: не удалось построить дату по CDHDR/CDPOS: %s", e)

        if assignment_dt_by_customer is None:
            date_col = _find_col(
                df_kna1,
                "central_order_block_assignment_date",
                "AUFSD_DATE",
                "BLOCK_DATE",
                "DATUB",
                "ERDAT",
            )
            if date_col and date_col in df_kna1.columns:
                assignment_dt_by_customer = df_kna1[["KUNNR_clean", date_col]].copy()
                assignment_dt_by_customer["ASSIGN_DT"] = pd.to_datetime(
                    assignment_dt_by_customer[date_col], errors="coerce"
                )
                assignment_dt_by_customer = (
                    assignment_dt_by_customer[["KUNNR_clean", "ASSIGN_DT"]]
                    .dropna(subset=["ASSIGN_DT"])
                    .drop_duplicates(subset=["KUNNR_clean"], keep="last")
                )
                assignment_date_source = "KNA1"
            else:
                return self._error_result(rule, "Не найдена дата назначения блока (CDHDR/CDPOS или date-column в KNA1)")

        # --- 4. JOIN по customer_code ---
        df_merged = df_kna1.merge(
            assignment_dt_by_customer[["KUNNR_clean", "ASSIGN_DT"]],
            on="KUNNR_clean",
            how="left",
        )

        # Лог для отладки источника даты
        self.logger.info(
            "RCCONF_173.1: источники: block_col=%s, dates_rows=%s, kna1_rows=%s",
            block_col,
            len(assignment_dt_by_customer),
            len(df_kna1),
        )

        # --- 5. Оцениваем только строки по technical_definition ---
        code_series = df_merged["BLOCK_CODE"]
        has_conf = code_series.isin(max_months.keys())
        has_date = df_merged["ASSIGN_DT"].notna()
        evaluable_mask = has_conf & has_date
        evaluable_count = int(evaluable_mask.sum())

        if getattr(self.checker, "debug", False):
            try:
                kna1_u = int(df_kna1['KUNNR_clean'].nunique(dropna=True))
                dates_u = int(assignment_dt_by_customer['KUNNR_clean'].nunique(dropna=True))
                kna1_s = set(df_kna1['KUNNR_clean'].dropna().head(200))
                date_s = set(assignment_dt_by_customer['KUNNR_clean'].dropna().head(200))
                self.logger.info(
                    "RCCONF_173.1 DEBUG: unique KUNNR_clean KNA1=%s, dates=%s, sample_common=%s",
                    kna1_u,
                    dates_u,
                    len(kna1_s & date_s),
                )
            except Exception as e:
                self.logger.warning("RCCONF_173.1 DEBUG: diagnostics failed: %s", e)

        self.logger.info(
            "RCCONF_173.1: total_kna1=%s, in_conf=%s, has_date=%s, evaluable=%s",
            len(df_kna1),
            int(has_conf.sum()),
            int(has_date.sum()),
            evaluable_count,
        )

        if evaluable_count == 0:
            # В отчёте по 173.1 — физические имена полей: UDATE (дата), AUFSD (код блока в KNA1)
            return self._result(rule, 0, 0, None, "UDATE", "UDATE")

        # --- 6. Расчет просрочки ---
        failed = 0
        error_indices = []
        now = getattr(self.checker, "reference_datetime", None)
        if now is None:
            now = datetime.now()
        self.logger.info(
            "RCCONF_173.1: опорная дата для расчёта (now − дата назначения): %s; "
            "колонка UDATE в отчёте — дата назначения блока из данных, не эта дата",
            now,
        )
        evaluable_dates = df_merged.loc[evaluable_mask, "ASSIGN_DT"]
        evaluable_codes = df_merged.loc[evaluable_mask, "BLOCK_CODE"]
        delta_months = (now - evaluable_dates).dt.days / 30.44
        months_allowed = evaluable_codes.map(max_months)
        over_limit = delta_months > months_allowed
        failed = int(over_limit.sum())
        error_indices = df_merged.loc[evaluable_mask].index[over_limit].tolist()

        self.logger.info("RCCONF_173.1: Превысили лимит: %s", failed)

        # --- 7. Создание error_df для выгрузки ---
        err_df = None
        if failed > 0 and error_indices:
            error_mask = df_merged.index.isin(error_indices)
            cols = [self._kunnr_col, "BLOCK_CODE", "ASSIGN_DT"]
            if assignment_date_source == "CDHDR":
                msg = (
                    "Превышение лимита: дата из CDHDR (UDATE + UTIME) старше допустимого срока "
                    "(conf_order_block_time.max_months) для кода блока в KNA1.AUFSD"
                )
            else:
                msg = (
                    "Превышение лимита: дата назначения блока старше допустимого срока для AUFSD "
                    "(источник даты — KNA1, не CDHDR)"
                )
            err_df = self._error_df_sample(
                df_merged,
                error_mask,
                cols,
                msg,
                getattr(self.checker, "MAX_ERRORS_TO_SAVE", 500),
            )
            if err_df is not None:
                rename_map = {}
                if "BLOCK_CODE" in err_df.columns:
                    rename_map["BLOCK_CODE"] = "AUFSD"
                if "ASSIGN_DT" in err_df.columns:
                    rename_map["ASSIGN_DT"] = "UDATE"
                if rename_map:
                    err_df = err_df.rename(columns=rename_map)
                err_df["max_months_allowed"] = df_merged.loc[error_mask, "BLOCK_CODE"].map(max_months).values
                # Явно: какая «сегодняшняя» дата использовалась в формуле (не путать с UDATE из CDHDR)
                err_df["reference_as_of"] = now.strftime("%Y-%m-%d %H:%M:%S")
                # Убираем дубли имён колонок (если в исходном df уже была AUFSD и т.п.)
                err_df = err_df.loc[:, ~err_df.columns.duplicated(keep="first")]

        # --- 8. Возвращаем результат ---
        self.logger.info("RCCONF_173.1: ИТОГ: total=%s, failed=%s, passed=%s", evaluable_count, failed, evaluable_count - failed)
        return self._result(
            rule,
            evaluable_count,
            failed,
            err_df,
            "UDATE",
            "UDATE",
        )