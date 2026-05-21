import pandas as pd
import sqlite3

from utils.sqlite_safe import connect_sqlite

from .base_validator import BaseValidator


class PaymentTermsConsistencyValidator(BaseValidator):
    """
    RCCONF_119.2
    Проверка согласованности Terms of Payment (ZTERM) между KNB1 и KNVV.

    Оценка:
    - если ZTERM на стороне KNB1 пустой -> запись не оценивается (0, без ошибок)
    - если для KUNNR на стороне KNVV нет записей -> запись не оценивается (0, без ошибок)
    - иначе: ошибка, если ZTERM KNB1 не найден среди ZTERM KNVV для того же KUNNR

    В ошибках обязательно выводим две колонки:
    - KNB1_ZTERM
    - KNVV_ZTERM (выбранный "первый" ZTERM для KUNNR, чтобы показать расхождение)
    - KNVV_ZTERM_ALL (все уникальные ZTERM из KNVV для KUNNR)
    """

    def _is_empty_token(self, v) -> bool:
        if v is None:
            return True
        if isinstance(v, float) and pd.isna(v):
            return True
        s = str(v).strip().lower()
        return s in ("", "nan", "none", "null", "-", ".")

    def _norm_kunnr(self, v) -> str:
        if self._is_empty_token(v):
            return ""
        # В SAP KUNNR часто хранится со срезанными/ведущими нулями — нормализуем
        return str(v).strip().lstrip("0") or "0"

    def _norm_zterm(self, v) -> str:
        if self._is_empty_token(v):
            return ""
        return str(v).strip()

    def validate(
        self,
        df: pd.DataFrame,
        column_name: str,
        knb1_kunnr_col: str = None,
        db_path: str = None,
        knvv_df: pd.DataFrame = None,
        knvv_kunnr_col: str = None,
        knvv_zterm_col: str = None,
        **kwargs,
    ):
        print("[PaymentTermsConsistencyValidator] start")
        print(f"[PaymentTermsConsistencyValidator] KNB1 rows: {0 if df is None else len(df)}")
        print(f"[PaymentTermsConsistencyValidator] KNB1 KUNNR col: {knb1_kunnr_col}")
        print(f"[PaymentTermsConsistencyValidator] KNB1 ZTERM col: {column_name}")

        if df is None or df.empty:
            return 0, 0, None

        if not knb1_kunnr_col or knb1_kunnr_col not in df.columns:
            return 0, 0, None

        if not column_name or column_name not in df.columns:
            return 0, 0, None

        # Источник KNVV:
        # 1) если передан knvv_df — используем его
        # 2) иначе делаем простой SQL и получаем агрегат по KUNNR
        knvv_map = {}
        if knvv_df is not None and not knvv_df.empty and knvv_kunnr_col in knvv_df.columns and knvv_zterm_col in knvv_df.columns:
            print(f"[PaymentTermsConsistencyValidator] KNVV source: memory, rows={len(knvv_df)}")
            print(f"[PaymentTermsConsistencyValidator] KNVV KUNNR col: {knvv_kunnr_col}")
            print(f"[PaymentTermsConsistencyValidator] KNVV ZTERM col: {knvv_zterm_col}")
            knvv_pairs = knvv_df[[knvv_kunnr_col, knvv_zterm_col]].copy()
            knvv_pairs.columns = ["_kunnr", "_zterm"]
            knvv_pairs["_kunnr"] = knvv_pairs["_kunnr"].apply(self._norm_kunnr)
            knvv_pairs["_zterm"] = knvv_pairs["_zterm"].apply(self._norm_zterm)
            knvv_pairs = knvv_pairs[(knvv_pairs["_kunnr"] != "") & (knvv_pairs["_zterm"] != "")]
            if knvv_pairs.empty:
                return 0, 0, None
            knvv_pairs = knvv_pairs.drop_duplicates()
            knvv_map = (
                knvv_pairs.groupby("_kunnr")["_zterm"]
                .apply(lambda s: sorted(pd.unique(s).tolist()))
                .to_dict()
            )
        else:
            if not db_path:
                return 0, 0, None
            try:
                print(f"[PaymentTermsConsistencyValidator] KNVV source: sqlite ({db_path})")
                conn = connect_sqlite(db_path)
                sql = """
                    SELECT
                        CASE
                            WHEN LTRIM(COALESCE(KUNNR, ''), '0') = '' THEN '0'
                            ELSE LTRIM(COALESCE(KUNNR, ''), '0')
                        END AS kunnr_norm,
                        GROUP_CONCAT(DISTINCT TRIM(ZTERM)) AS zterm_all
                    FROM KNVV
                    WHERE ZTERM IS NOT NULL
                      AND TRIM(ZTERM) <> ''
                      AND UPPER(TRIM(ZTERM)) NOT IN ('NULL', 'NONE', 'NAN', '-', '.')
                    GROUP BY 1
                """
                agg = pd.read_sql_query(sql, conn)
                conn.close()
                if agg is None or agg.empty:
                    return 0, 0, None
                for _, row in agg.iterrows():
                    k = str(row["kunnr_norm"]).strip()
                    z_raw = str(row["zterm_all"]).strip() if row["zterm_all"] is not None else ""
                    vals = [self._norm_zterm(x) for x in z_raw.split(",")] if z_raw else []
                    vals = sorted([v for v in vals if v])
                    if k and vals:
                        knvv_map[k] = vals
            except Exception:
                return 0, 0, None

        # Нормализуем KNB1
        knb1_zterm_norm = df[column_name].apply(self._norm_zterm)
        knb1_kunnr_norm = df[knb1_kunnr_col].apply(self._norm_kunnr)
        if not knvv_map:
            return 0, 0, None

        # Evaluated: b not null AND c not null
        evaluated_mask = (knb1_zterm_norm != "") & (knb1_kunnr_norm != "") & (knb1_kunnr_norm.isin(set(knvv_map.keys())))
        total_rows = int(evaluated_mask.sum())

        # Ошибка: оцениваемые строки и ZTERM не найден среди KNVV_ZTERM по KUNNR
        exists_pair = pd.Series(
            [z in set(knvv_map.get(k, [])) if (k and z) else False for k, z in zip(knb1_kunnr_norm, knb1_zterm_norm)],
            index=df.index,
        )
        error_mask = evaluated_mask & (~exists_pair)
        error_count = int(error_mask.sum())
        print(f"[PaymentTermsConsistencyValidator] evaluated_rows={total_rows}, error_count={error_count}")

        if error_count == 0:
            return total_rows, 0, None

        # Формируем error_df
        error_df = df.loc[error_mask].copy()

        # Выводим сравниваемые колонки (именно то, что просили)
        error_df["KNB1_ZTERM"] = knb1_zterm_norm.loc[error_mask].values

        # Предпосчитаем строку "все ZTERM" на KUNNR (без повторов) — это сильно ускоряет заполнение ошибок
        knvv_all_by_kunnr = {k: "|".join(vs) for k, vs in knvv_map.items()}
        error_df["KNVV_ZTERM_ALL"] = knb1_kunnr_norm.loc[error_mask].map(
            lambda k: knvv_all_by_kunnr.get(k, "")
        ).values

        # Для наглядности: выбираем "первый" ZTERM из KNVV для KUNNR
        knvv_first = {k: (vs[0] if vs else "") for k, vs in knvv_map.items()}
        error_df["KNVV_ZTERM"] = knb1_kunnr_norm.loc[error_mask].map(lambda k: knvv_first.get(k, "")).values

        # Метаданные об ошибке
        error_df["DQ_ERROR_TYPE"] = "UNEQUAL_VALUES"
        error_df["DQ_RULE_CODE"] = self.rule_info["rule_code"]
        error_df["DQ_RULE_DESCRIPTION"] = self.rule_info.get("rule_description", "")
        error_df["DQ_COLUMN_CHECKED"] = "KNB1_ZTERM"
        error_df["DQ_ERROR_DESCRIPTION"] = (
            "Consistency between Payment Terms on KNB1 and KNVV: "
            "KNB1_ZTERM not found in KNVV ZTERM set for the same KUNNR"
        )
        error_df["DQ_TIMESTAMP"] = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")

        if self.error_saver:
            self._save_errors_if_needed(error_df)

        return total_rows, error_count, error_df

