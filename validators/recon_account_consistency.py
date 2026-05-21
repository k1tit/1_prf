import json
import os
import re

import pandas as pd

from .base_validator import BaseValidator


class ReconAccountConsistencyValidator(BaseValidator):
    """
    RCCONF_113.1
    Consistency between Recon. Account and Account Group.

    Rule logic:
    - IF reconciliation_account IS NULL OR account_group_code IS NULL THEN '' (skip)
    - ELSE IF (account_group_code, reconciliation_account) exists in conf_recon_accounts THEN '1'
    - ELSE '0'
    """

    EMPTY_TOKENS = {"", "none", "null", "nan", "<na>", "nat", "-", ".", "n/a", "na"}

    def _norm(self, v) -> str:
        if v is None:
            return ""
        if isinstance(v, float) and pd.isna(v):
            return ""
        # Явный быстрый SKIP для числового нуля
        if isinstance(v, (int, float)) and v == 0:
            return ""
        s = str(v)
        # Нормализация "невидимых" пробелов/служебных символов из Excel/CSV
        s = s.replace("\u00a0", " ").replace("\ufeff", "")
        s = re.sub(r"\s+", " ", s).strip()
        s = s.strip("'").strip('"').strip()
        if s in {"''", '""'}:
            return ""
        # Частый артефакт загрузки: пустой AKONT превращается в 0 / 0.0
        if s in {"0", "0.0", "0.00"}:
            return ""
        if s.replace("0", "") == "":
            return ""
        if re.fullmatch(r"0+(\.0+)?", s):
            return ""
        if s.endswith(".0") and s[:-2].isdigit():
            s = s[:-2]
        return "" if s.lower() in self.EMPTY_TOKENS else s

    def _load_allowed_pairs(self, reference_path: str = None):
        candidate_paths = []
        if reference_path:
            candidate_paths.append(reference_path)

        current_dir = os.path.dirname(os.path.abspath(__file__))
        project_root = os.path.dirname(current_dir)
        candidate_paths.extend(
            [
                os.path.join(project_root, "json files", "conf_recon_accounts.json"),
                os.path.join(os.getcwd(), "json files", "conf_recon_accounts.json"),
            ]
        )

        for path in candidate_paths:
            try:
                if not path or not os.path.exists(path):
                    continue
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                rows = data.get("conf_recon_accounts", []) if isinstance(data, dict) else []
                allowed = set()
                for row in rows:
                    ag = self._norm(row.get("account_group_code"))
                    ra = self._norm(row.get("reconciliation_account"))
                    if ag and ra:
                        allowed.add((ag, ra))
                if allowed:
                    return allowed, path
            except Exception:
                continue
        return set(), None

    def validate(
        self,
        df: pd.DataFrame,
        column_name: str,
        account_group_col: str = None,
        reference_path: str = None,
        **kwargs,
    ):
        if df is None or df.empty:
            return 0, 0, None

        if not column_name or column_name not in df.columns:
            return 0, 0, None

        # account_group_code может прийти как account_group_code или b.account_group_code
        if not account_group_col or account_group_col not in df.columns:
            for c in df.columns:
                cu = str(c).strip().lower()
                if cu in ("account_group_code", "b.account_group_code", "ktokd", "b.ktokd"):
                    account_group_col = c
                    break
        if not account_group_col or account_group_col not in df.columns:
            return 0, 0, None

        allowed_pairs, used_ref = self._load_allowed_pairs(reference_path=reference_path)
        if not allowed_pairs:
            return 0, 0, None

        recon_norm = df[column_name].apply(self._norm)
        group_norm = df[account_group_col].apply(self._norm)

        # IF a.reconciliation_account IS NULL OR b.account_group_code IS NULL THEN '' (skip)
        has_recon = recon_norm != ""
        has_group = group_norm != ""
        evaluated_mask = has_recon & has_group
        total_rows = int(evaluated_mask.sum())
        skip_rows = int((~evaluated_mask).sum())
        print(
            f"[ReconAccountConsistencyValidator] rows={len(df)}, "
            f"evaluated={total_rows}, skipped_empty={skip_rows}"
        )
        if total_rows == 0:
            return 0, 0, None

        # Проверяем пары только для оценённых строк
        exists_mask = pd.Series(False, index=df.index)
        for idx in df.index[evaluated_mask]:
            g = group_norm.loc[idx]
            r = recon_norm.loc[idx]
            exists_mask.loc[idx] = (g, r) in allowed_pairs

        error_mask = evaluated_mask & (~exists_mask)
        error_count = int(error_mask.sum())
        empty_recon_in_errors = int(((recon_norm == "") & error_mask).sum())
        if empty_recon_in_errors > 0:
            print(
                f"[ReconAccountConsistencyValidator][WARN] empty_recon_in_errors={empty_recon_in_errors} "
                "(это не должно происходить)"
            )
            # Диагностика сырого значения AKONT для проблемных строк
            bad_idx = df.index[(recon_norm == "") & error_mask]
            sample_idx = list(bad_idx[:5])
            sample_raw = [repr(df.loc[i, column_name]) for i in sample_idx]
            sample_group_raw = [repr(df.loc[i, account_group_col]) for i in sample_idx]
            print(
                "[ReconAccountConsistencyValidator][WARN] sample raw AKONT/group: "
                f"{list(zip(sample_raw, sample_group_raw))}"
            )
        print(f"[ReconAccountConsistencyValidator] error_count={error_count}")
        if error_count == 0:
            return total_rows, 0, None

        error_df = df.loc[error_mask].copy()
        error_df["ACCOUNT_GROUP_CODE"] = group_norm.loc[error_mask].values
        error_df["RECONCILIATION_ACCOUNT"] = recon_norm.loc[error_mask].values
        error_df["DQ_ERROR_TYPE"] = "INVALID_COMBINATION"
        error_df["DQ_RULE_CODE"] = self.rule_info["rule_code"]
        error_df["DQ_RULE_DESCRIPTION"] = self.rule_info.get("rule_description", "")
        error_df["DQ_COLUMN_CHECKED"] = column_name
        error_df["DQ_ERROR_DESCRIPTION"] = (
            "Invalid combination of account_group_code and reconciliation_account "
            f"(reference: {os.path.basename(used_ref) if used_ref else 'conf_recon_accounts'})"
        )
        error_df["DQ_TIMESTAMP"] = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")

        if self.error_saver:
            self._save_errors_if_needed(error_df)

        return total_rows, error_count, error_df

