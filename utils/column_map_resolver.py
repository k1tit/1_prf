"""
Поиск физической колонки в DataFrame по rules / column_map.json.
Логическое имя (account_group_code) и SAP (KTOKD) — оба варианта.
"""
from __future__ import annotations

import json
import os
import re
from typing import Any, Dict, Iterable, List, Optional

import pandas as pd

from utils.column_matcher import ColumnMatcher

_matcher = ColumnMatcher()
_cache: Dict[str, dict] = {}


def _norm(name: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(name or "").upper())


def load_column_map(project_root: str) -> dict:
    if project_root in _cache:
        return _cache[project_root]
    for rel in (os.path.join("json files", "column_map.json"), os.path.join("config", "column_map.json")):
        path = os.path.join(project_root, rel)
        if os.path.isfile(path):
            try:
                with open(path, encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, dict):
                    _cache[project_root] = data
                    return data
            except (json.JSONDecodeError, OSError):
                pass
    _cache[project_root] = {}
    return {}


def _is_flat_table_section(m: dict) -> bool:
    """Фрагмент одной таблицы: {"customer_code": "KUNNR", "_aliases": {...}} без ключа KNA1."""
    if not isinstance(m, dict) or not m:
        return False
    for k, v in m.items():
        if str(k).startswith("_"):
            continue
        if isinstance(v, dict):
            return False
    return True


def _table_mapping(column_map: dict, table_name: str) -> Optional[dict]:
    if not column_map or not table_name:
        if column_map and _is_flat_table_section(column_map):
            return column_map
        return None
    if table_name in column_map:
        m = column_map[table_name]
        return m if isinstance(m, dict) else None
    tn = str(table_name).strip().upper()
    for k, v in column_map.items():
        if str(k).strip().upper() == tn and isinstance(v, dict):
            return v
    if _is_flat_table_section(column_map):
        return column_map
    return None


def _collect_equivalent_names(table_mapping: Optional[dict], target: str) -> List[str]:
    """
    Все синонимы для target: SAP, логические имена, ключи/значения column_map, _aliases.
    Позволяет искать колонку и по KATR1, и по «Attribute 1», и по customer_activity_cluster_code.
    """
    if not table_mapping or not target:
        return [target] if target else []
    out: List[str] = []
    seen: set[str] = set()

    def add(x: str) -> None:
        s = str(x or "").strip()
        if not s or s.startswith("_"):
            return
        k = s.upper()
        if k not in seen:
            seen.add(k)
            out.append(s)

    add(target)
    tu = target.upper()

    for logical, physical in table_mapping.items():
        if str(logical).startswith("_"):
            continue
        log = str(logical).strip()
        phys = str(physical).strip()
        if tu in (log.upper(), phys.upper()):
            add(log)
            add(phys)
            # ключ часто = исходный заголовок выгрузки (Customer, Attribute 1)
            if log.upper() != phys.upper():
                add(log)

    aliases = table_mapping.get("_aliases")
    if isinstance(aliases, dict):
        for sap_name, alias_list in aliases.items():
            sap = str(sap_name).strip()
            if tu == sap.upper() or tu == str(table_mapping.get(sap, "")).upper():
                add(sap)
                if isinstance(alias_list, list):
                    for a in alias_list:
                        add(str(a))
            elif isinstance(alias_list, list):
                for a in alias_list:
                    if str(a).strip().upper() == tu:
                        add(sap)
                        for x in alias_list:
                            add(str(x))

    return out


def _global_equivalents_for_sap(
    column_map: dict,
    sap_or_target: str,
    table_name: str = "",
) -> List[str]:
    """
    Синонимы SAP-поля из всех таблиц column_map (valid_from/DATAB из KNVV помогает KNVH и т.д.).
    """
    if not column_map or not sap_or_target:
        return []
    tm = _table_mapping(column_map, table_name)
    sap = str(sap_or_target).strip()
    tu = sap.upper()
    if tm:
        for logical, physical in tm.items():
            if str(logical).startswith("_"):
                continue
            phys = str(physical).strip()
            log = str(logical).strip()
            if tu in (log.upper(), phys.upper()):
                sap = phys
                break
    su = sap.upper()
    out: List[str] = []
    seen: set[str] = set()

    def add(x: str) -> None:
        s = str(x or "").strip()
        if not s or s.startswith("_"):
            return
        k = s.upper()
        if k not in seen:
            seen.add(k)
            out.append(s)

    add(sap)
    for _tname, tmap in column_map.items():
        if not isinstance(tmap, dict):
            continue
        for logical, physical in tmap.items():
            if str(logical).startswith("_"):
                continue
            if str(physical).strip().upper() == su:
                add(str(logical))
        aliases = tmap.get("_aliases")
        if not isinstance(aliases, dict):
            continue
        for sap_key, alias_list in aliases.items():
            if str(sap_key).strip().upper() != su:
                continue
            add(str(sap_key))
            if isinstance(alias_list, list):
                for a in alias_list:
                    add(str(a))
    return out


# Типовые сокращения SAP GUI / выгрузки проекта (дополняют column_map.json, не дублируют логические имена).
_SAP_GUI_EXPORT: Dict[str, List[str]] = {
    "VKORG": ["SOrg_", "SOrg"],
    "VTWEG": ["DChl"],
    "SPART": ["Dv"],
    "HSPART": ["HLDiv"],
    "HKUNNR": ["HgLvCust_"],
    "HVKORG": ["HLSOr"],
    "HVTWEG": ["HLDCh"],
    "VKGRP": ["SGrp"],
    "VSBED": ["SC"],
    "VWERK": ["Plnt"],
    "KTGRD": ["AAGC"],
    "KVGR4": ["Grp4"],
    "KONDA": ["CPG"],
    "KDGRP": ["CGrp"],
    "ZTERM": ["PayT"],
    "VKBUR": ["SOff_"],
    "PARVW": ["Funct", "Par.func.", "Partner Function", "Partner function"],
    "TERRID": ["Territory", "Territory_ID", "Terr_ID"],
    "DATE_FROM": ["Valid_from", "Date_from", "Date From"],
    "DATE_TO": ["Valid_to", "Date_to", "Date To"],
    "CREDIT_SGMNT": ["Credit_Segment"],
    "PARTNER": ["Business_Partner", "PARTNER"],
    "PARTNER1": ["Business_Partner"],
    "PARTNER2": ["Business_Partner_1"],
    "CLIENT": ["Cl_"],
    "RELNR": ["BP_Relation__No_"],
    "DATE_TO": ["Valid_To"],
    "RELTYP": ["RelCat"],
    "XRF": ["RD"],
    "FNCTN": ["Function_name"],
    "PAFKT": ["Fct"],
    "DPRTMNT": ["Company_department"],
    "ABTNR": ["Dept"],
    "PAAUTH": ["Authority"],
    "PAVIP": ["VIP"],
    "PAREM": ["Note"],
    "TEL_NUMBER": ["Telephone"],
    "TEL_EXTENS": ["Extension"],
    "FAX_NUMBER": ["Fax"],
    "FAX_EXTENS": ["Extension_1"],
    "SMTP_ADDRESS": ["E-Mail_Address", "E-Mail Address"],
    "REL_PER": ["%"],
    "REL_AMO": ["Amount"],
    "REL_CUR": ["Curr_"],
    "CALL_RULEID": ["Rule_ID/Call"],
    "VISIT_RULEID": ["Rule_ID_/_Visit"],
    "CALL_GUID": ["Calendar_Schema_GUID"],
    "VISIT_GUID": ["Calendar_Schema_GUID_1"],
    "BP_EEW_BUT051": ["Dummy_function_in_length_1"],
    "BP_EEW_BUT051_SP": ["Dummy_function_in_length_1_1"],
    "TYPE": ["TYPE"],
    "POST_CODE1": ["Postl_Code", "Postal Code", "Postl Code"],
    "BANKS": ["C/R", "Bank Ctry", "Bank country"],
}


def _gui_export_headers_for_sap(sap: str) -> List[str]:
    su = str(sap or "").strip().upper()
    out: List[str] = []
    for key, headers in _SAP_GUI_EXPORT.items():
        if key.upper() == su:
            out.extend(headers)
    return out


def _export_header_candidates(target: str, table_mapping: Optional[dict] = None) -> List[str]:
    """
    Типовые заголовки выгрузки для SAP-имени (без дублирования в column_map.json).
    NAME_ORG1 -> Name_1, NAME_1, MC_NAME1; organization_1_name — из маппинга таблицы.
    """
    out: List[str] = []
    seen: set[str] = set()

    def add(x: str) -> None:
        s = str(x or "").strip()
        if not s:
            return
        k = s.upper()
        if k not in seen:
            seen.add(k)
            out.append(s)

    t = str(target or "").strip()
    add(t)
    m = re.match(r"^NAME_ORG(\d+)$", t.upper())
    if m:
        n = m.group(1)
        add(f"Name_{n}")
        add(f"NAME_{n}")
        add(f"Name {n}")
        add(f"NAME{n}")
        add(f"MC_NAME{n}")
    m_org = re.match(r"^organization_(\d+)_name$", t.lower())
    if m_org:
        n = m_org.group(1)
        add(f"NAME_ORG{n}")
        add(f"Name_{n}")
        add(f"NAME_{n}")
        add(f"Name {n}")
    # Типовые SAP-выгрузки: valid_from -> DATAB, SOrg_ -> VKORG (через нормализацию)
    tl = t.lower()
    if tl in ("datab", "valid_from", "validfrom"):
        add("Valid_from")
        add("Valid from")
        add("Valid From")
        add("DATAB")
    if tl in ("datbi", "valid_to", "validto"):
        add("Valid_to")
        add("Valid to")
        add("Valid To")
        add("DATBI")
    for gui in _gui_export_headers_for_sap(t):
        add(gui)
    if table_mapping:
        tu = t.upper()
        for logical, physical in table_mapping.items():
            if str(logical).startswith("_"):
                continue
            phys = str(physical).strip()
            log = str(logical).strip()
            if tu in (phys.upper(), log.upper()):
                add(log)
                add(phys)
    return out


def _column_matches_equivalent(col: str, equivalent: str) -> bool:
    """Сопоставление физического заголовка выгрузки с синонимом из column_map."""
    if not col or not equivalent:
        return False
    if col == equivalent:
        return True
    if str(col).strip().upper() == str(equivalent).strip().upper():
        return True
    cn = _norm(col)
    en = _norm(equivalent)
    if not cn or not en:
        return False
    if cn == en:
        return True
    if len(en) >= 4 and en in cn:
        # TAXNUM внутри Tax_Number_Category / Tax_Number_Long — не номер, а тип/длинное поле
        if en == "TAXNUM" and cn in ("TAXNUMBERCATEGORY", "TAXNUMBERLONG"):
            return False
        return True
    # Не матчить короткое имя колонки (DATE) внутри длинного логического (…assignment_date)
    if len(cn) >= 4 and cn in en and len(cn) >= max(6, int(len(en) * 0.65)):
        return True
    return False


def _reverse_scan_columns(
    columns: Iterable[str],
    target: str,
    table_mapping: Optional[dict],
    column_map: Optional[dict],
    table_name: str = "",
) -> Optional[str]:
    """
    Обратный поиск: физическая колонка выгрузки (Valid_from, Name_1, SOrg_) -> SAP-имя из rules.
    """
    cols = list(columns)
    if not cols or not target:
        return None
    tu = str(target).strip().upper()
    sap_guess = str(target).strip()
    if table_mapping:
        for logical, physical in table_mapping.items():
            if str(logical).startswith("_"):
                continue
            log = str(logical).strip()
            phys = str(physical).strip()
            if tu in (log.upper(), phys.upper()):
                sap_guess = phys
                break

    equivalents: List[str] = []
    seen: set[str] = set()

    def add_eq(x: str) -> None:
        s = str(x or "").strip()
        if not s:
            return
        k = s.upper()
        if k not in seen:
            seen.add(k)
            equivalents.append(s)

    for eq in _expand_candidates(table_mapping, target):
        add_eq(eq)
    for eq in _export_header_candidates(target, table_mapping):
        add_eq(eq)
    if column_map:
        for eq in _global_equivalents_for_sap(column_map, sap_guess, table_name):
            add_eq(eq)

    best: Optional[str] = None
    best_score = -1
    for col in cols:
        col_s = str(col).strip()
        if col_s.upper() == tu or _norm(col_s) == _norm(target):
            return col_s
        sap = _canonical_sap_name_for_column(col_s, table_mapping or {})
        if sap and sap.upper() == sap_guess.upper():
            score = 100 + len(_norm(col_s))
            if score > best_score:
                best_score = score
                best = col_s
            continue
        for eq in equivalents:
            if _column_matches_equivalent(col_s, eq):
                score = 50 + len(_norm(eq))
                if score > best_score:
                    best_score = score
                    best = col_s
                break
        if _matcher.find_column_match([col_s], target):
            score = 40 + len(_norm(col_s))
            if score > best_score:
                best_score = score
                best = col_s
    return best


def map_logical_to_sap(
    table_name: str,
    name: str,
    column_map: Optional[dict] = None,
    project_root: str = "",
) -> str:
    """Логическое/datamart-имя или SAP -> каноническое SAP-имя по column_map.json."""
    if not name:
        return name
    if column_map is None and project_root:
        column_map = load_column_map(project_root)
    tm = _table_mapping(column_map or {}, table_name)
    if not tm:
        return name
    n = str(name).strip()
    nu = n.upper()
    if n in tm:
        return str(tm[n]).strip()
    for logical, physical in tm.items():
        if str(logical).startswith("_"):
            continue
        if str(logical).upper() == nu:
            return str(physical).strip()
    for logical, physical in tm.items():
        if str(logical).startswith("_"):
            continue
        if str(physical).upper() == nu:
            return str(physical).strip()
    return name


def _expand_candidates(table_mapping: Optional[dict], *names: str) -> List[str]:
    """Все варианты имён: SAP, логические, исходные заголовки, пары из column_map."""
    out: List[str] = []
    seen: set[str] = set()

    def add(x: str) -> None:
        s = str(x or "").strip()
        if not s or s.startswith("_"):
            return
        key = s.upper()
        if key not in seen:
            seen.add(key)
            out.append(s)

    for n in names:
        for eq in _collect_equivalent_names(table_mapping, n):
            add(eq)
    return out


def resolve_column(
    columns: Iterable[str],
    target: str,
    table_name: str = "",
    column_map: Optional[dict] = None,
    project_root: str = "",
) -> Optional[str]:
    """
    Находит колонку в columns для target (KTOKD, account_group_code, KATR1, ...).
    """
    cols = list(columns)
    if not cols or not target:
        return None

    if column_map is None and project_root:
        column_map = load_column_map(project_root)
    tm = _table_mapping(column_map or {}, table_name)
    sap_target = map_logical_to_sap(table_name, target, column_map or {}, "")

    global_eq = _global_equivalents_for_sap(column_map or {}, sap_target, table_name) if column_map else []
    candidates: List[str] = []
    seen_c: set[str] = set()
    for c in (
        _expand_candidates(tm, target)
        + _expand_candidates(tm, sap_target)
        + _export_header_candidates(target, tm)
        + _export_header_candidates(sap_target, tm)
        + global_eq
    ):
        k = str(c).upper()
        if k not in seen_c:
            seen_c.add(k)
            candidates.append(c)

    for cand in candidates:
        for col in cols:
            if col == cand:
                return col
        cu = cand.upper()
        for col in cols:
            if str(col).strip().upper() == cu:
                return col
        cn = _norm(cand)
        if cn:
            for col in cols:
                if _norm(col) == cn:
                    return col

    # column_map.json: "_aliases": { "KATR1": ["...", "..."] }
    if tm and isinstance(tm.get("_aliases"), dict):
        aliases = tm.get("_aliases", {})
        tu = target.upper()
        sap_key: Optional[str] = None
        if target in aliases or tu in {str(k).upper() for k in aliases}:
            for k in aliases:
                if str(k).upper() == tu or k == target:
                    sap_key = str(k)
                    break
        if sap_key is None:
            for k, alias_list in aliases.items():
                if not isinstance(alias_list, list):
                    continue
                if any(str(a).strip().upper() == tu for a in alias_list):
                    sap_key = str(k)
                    break
        if sap_key:
            names: List[str] = [sap_key]
            extra = aliases.get(sap_key)
            if isinstance(extra, list):
                names.extend(str(a) for a in extra)
            for cand in sorted(names, key=lambda x: -len(str(x))):
                for col in cols:
                    if col == cand:
                        return col
                cu = str(cand).strip().upper()
                for col in cols:
                    if str(col).strip().upper() == cu:
                        return col
                cn = _norm(cand)
                if cn:
                    for col in cols:
                        if _norm(col) == cn:
                            return col

    matched = _matcher.find_column_match(cols, target)
    if matched:
        return matched

    tn = _norm(target)
    if not tn:
        return None
    for col in cols:
        cn = _norm(col)
        if cn == tn:
            return col
        if len(tn) >= 4 and tn in cn:
            return col

    rev = _reverse_scan_columns(cols, target, tm, column_map, table_name)
    if rev:
        return rev
    if sap_target != target:
        rev = _reverse_scan_columns(cols, sap_target, tm, column_map, table_name)
        if rev:
            return rev
    return None


def resolve_column_in_df(
    df: pd.DataFrame,
    target: str,
    table_name: str = "",
    column_map: Optional[dict] = None,
    project_root: str = "",
) -> Optional[str]:
    if df is None or df.empty:
        return None
    return resolve_column(df.columns, target, table_name, column_map, project_root)


def _canonical_sap_name_for_column(col: str, table_mapping: dict) -> Optional[str]:
    """
    К какому SAP-имени (KTOKD, KATR1) привести колонку из выгрузки.
    None — оставить имя как в БД.
    """
    if not col or not table_mapping:
        return None
    col_s = str(col).strip()
    cn = _norm(col_s)
    if not cn:
        return None

    for logical, physical in table_mapping.items():
        if str(logical).startswith("_"):
            continue
        phys = str(physical).strip()
        log = str(logical).strip()
        if col_s == phys or col_s == log:
            return phys
        if cn in (_norm(phys), _norm(log)):
            return phys

    for sap, headers in _SAP_GUI_EXPORT.items():
        for h in headers:
            hs = str(h).strip()
            if col_s == hs or cn == _norm(hs):
                return sap

    aliases = table_mapping.get("_aliases")
    if isinstance(aliases, dict):
        # Сначала длинные alias (чтобы "Customer Group 4" не схлопнулся в "Customer Group")
        items: List[tuple] = []
        for sap_name, alias_list in aliases.items():
            sap = str(sap_name).strip()
            if col_s == sap or cn == _norm(sap):
                return sap
            if not isinstance(alias_list, list):
                continue
            for alias in alias_list:
                a = str(alias).strip()
                if a:
                    items.append((len(a), sap, a))
        for _ln, sap, a in sorted(items, key=lambda x: (-x[0], x[1])):
            if col_s == a or cn == _norm(a):
                return sap

    # Частичное совпадение: длинный alias внутри заголовка (Sales org. -> VKORG).
    # Не используем cn in an: однобуквенная колонка "R" ложно матчилась на KATR1/KATR2/...
    best_sap: Optional[str] = None
    best_len = 0
    for sap_name, alias_list in (aliases or {}).items():
        if not isinstance(alias_list, list):
            continue
        sap = str(sap_name).strip()
        for alias in [sap_name, sap] + alias_list:
            an = _norm(str(alias))
            if len(an) < 4:
                continue
            if an == cn:
                return sap
            if an in cn and len(an) > best_len:
                if an == "TAXNUM" and cn in ("TAXNUMBERCATEGORY", "TAXNUMBERLONG"):
                    continue
                if an == "TELEPHONE" and cn != "TELEPHONE":
                    continue
                best_len = len(an)
                best_sap = sap
    if best_sap:
        return best_sap

    return None


def _non_empty_count(series: pd.Series) -> int:
    if series is None or series.empty:
        return 0
    s = series.astype(str).str.strip()
    return int((~series.isna() & ~s.isin(["", "nan", "None", "null", "NaN", "NA", "<NA>"])).sum())


def _parse_zbut_composite_date_to(v) -> tuple:
    """
    Выгрузка ZBUT0000P*: иногда TERRID + DATE_FROM + DATE_TO в одной колонке DATE_TO
    (формат SAP: 99.991.231.235.959 -> to=9999-12-31, from=1000-01-01, terrid=231235959).
    """
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "", "", ""
    s = str(v).strip().strip("'").strip('"')
    if not s or s.lower() in ("nan", "none", "null", "na", "<na>"):
        return "", "", ""
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", s):
        return "", s, s
    if re.fullmatch(r"\d+\.\d+\.\d+\.\d+\.\d+", s):
        parts = s.split(".")
        terrid = "".join(parts[2:5])
        if len(parts) >= 2 and parts[0] == "99" and parts[1] == "991":
            return terrid, "1000-01-01", "9999-12-31"
        dfrom = ".".join(parts[:3]) if len(parts) >= 3 else s
        return terrid, dfrom, "9999-12-31"
    return s, "", s


def _expand_zbut_territory_from_composite_date_to(
    df: pd.DataFrame,
    table_mapping: Optional[dict],
    *,
    log_renames: bool = True,
) -> pd.DataFrame:
    """Разбор SAP-композита в DATE_TO -> TERRID, DATE_FROM, DATE_TO (ISO)."""
    if df is None or df.empty:
        return df
    expand = table_mapping.get("_expand_composite_date_to") if table_mapping else False
    if not expand:
        return df

    out = df.copy()
    has_terrid = "TERRID" in out.columns and _non_empty_count(out["TERRID"]) > 0
    has_dfrom = "DATE_FROM" in out.columns and _non_empty_count(out["DATE_FROM"]) > 0
    if has_terrid and has_dfrom:
        return out

    src = None
    for c in out.columns:
        if _norm(c) == "DATETO":
            src = c
            break
    if not src:
        return out

    parsed = out[src].apply(_parse_zbut_composite_date_to)
    terrid_s = parsed.apply(lambda x: x[0])
    dfrom_s = parsed.apply(lambda x: x[1])
    dto_s = parsed.apply(lambda x: x[2])

    if not has_terrid:
        out["TERRID"] = terrid_s
    if not has_dfrom:
        out["DATE_FROM"] = dfrom_s
    out["DATE_TO"] = dto_s

    if log_renames:
        n_ter = int((terrid_s.astype(str).str.strip() != "").sum())
        print(
            f"   [MAP] ZBUT*: из [{src}] разобраны TERRID/DATE_FROM/DATE_TO "
            f"(SAP-композит), заполнено TERRID: {n_ter:,}/{len(out):,}"
        )
    return out


def _sap_copy_source_rank(col_name: str, sap: str) -> int:
    """
    При нескольких выгрузочных колонках на один SAP (KNA1: Customer и Cl_ -> KUNNR)
    предпочитаем номер клиента (Customer), не Cl_ (в SAP это другое поле).
    """
    cn = _norm(col_name)
    su = _norm(sap)
    if su == "KUNNR":
        if cn in ("CUSTOMER", "CUSTOMERCODE", "KUNNR", "KUNN"):
            return 0
        if "CUSTOMER" in cn:
            return 1
        if cn in ("CL", "CL_"):
            return 99
        return 50
    if su == "TELNUMBER":
        # ADR2: Telephone = SAP TEL_NUMBER (цифры); Telephone_number = форматированный (+7…)
        if cn == "TELEPHONE":
            return 0
        if cn == "TELEPHONENUMBER":
            return 1
        return 50
    return 50


def apply_column_headers_for_rules(
    df: pd.DataFrame,
    table_name: str,
    column_map: Optional[dict] = None,
    project_root: str = "",
    *,
    log_renames: bool = True,
) -> pd.DataFrame:
    """
    Подготовка шапки для прогона правил (копия DataFrame; SQLite не меняется).

    Гибкий режим:
    - исходные имена колонок из выгрузки **сохраняются**;
    - по column_map добавляются SAP/логические **копии** (Customer + KUNNR, Attribute 1 + KATR1);
    - правила и resolve_column находят колонку по любому из вариантов.
    """
    if df is None or df.empty:
        return df
    if column_map is None and project_root:
        column_map = load_column_map(project_root)
    tm = _table_mapping(column_map or {}, table_name)
    if not tm:
        return df.copy()

    out = df.copy()
    added: List[tuple[str, str]] = []
    taken = {_norm(c) for c in out.columns}

    # Несколько заголовков могут мапиться на один SAP (Group и Group_1 -> KTOKD).
    # Берём колонку с максимумом непустых значений.
    sap_sources: Dict[str, tuple[str, int, bool]] = {}
    for col in df.columns:
        sap = _canonical_sap_name_for_column(col, tm)
        if not sap:
            continue
        col_s = str(col).strip()
        if _norm(col_s) == _norm(sap):
            continue
        filled = _non_empty_count(df[col])
        exact = _norm(col_s) == _norm(sap)
        prev = sap_sources.get(sap)
        if prev is None:
            sap_sources[sap] = (col_s, filled, exact)
        else:
            _pc, prev_filled, prev_exact = prev
            new_rank = _sap_copy_source_rank(col_s, sap)
            prev_rank = _sap_copy_source_rank(_pc, sap)
            if exact and not prev_exact:
                sap_sources[sap] = (col_s, filled, exact)
            elif exact == prev_exact and filled > prev_filled:
                sap_sources[sap] = (col_s, filled, exact)
            elif filled == prev_filled and new_rank < prev_rank:
                sap_sources[sap] = (col_s, filled, exact)

    for sap, (col_s, filled, _exact) in sap_sources.items():
        if filled <= 0 and sap in out.columns:
            continue
        if sap in out.columns:
            if _norm(col_s) == _norm(sap):
                continue
            if _non_empty_count(out[sap]) >= filled:
                continue
            out[sap] = df[col_s]
            added.append((col_s, sap))
            continue
        final = sap
        suffix = 1
        while _norm(final) in taken:
            final = f"{sap}_{suffix}"
            suffix += 1
        out[final] = df[col_s]
        taken.add(_norm(final))
        added.append((col_s, final if final != sap else sap))

    # DFKKBPTAXNUM: TAXNUM = Tax_Number, если пусто — Tax_Number_Long (не RU5 из Category)
    tn_u = str(table_name or "").strip().upper()
    if tn_u.startswith("DFKKBPTAXNUM"):
        short_c = next((c for c in ("Tax_Number", "TAXNUM") if c in out.columns), None)
        long_c = next((c for c in ("Tax_Number_Long", "TAXNUM_LONG") if c in out.columns), None)
        if short_c or long_c:

            def _nz(series: pd.Series) -> pd.Series:
                return series.map(
                    lambda v: ""
                    if pd.isna(v)
                    else str(int(v))
                    if isinstance(v, (int, float)) and v == int(v)
                    else str(v).strip()
                )

            s = _nz(out[short_c]) if short_c else pd.Series("", index=out.index)
            lng = _nz(out[long_c]) if long_c else pd.Series("", index=out.index)
            bad = {"", "nan", "None", "null", "NaN", "NA", "<NA>"}
            s = s.where(~s.str.lower().isin(bad), "")
            lng = lng.where(~lng.str.lower().isin(bad), "")
            out["TAXNUM"] = s.where(s != "", lng)
            taken.add(_norm("TAXNUM"))
            if log_renames:
                print(
                    f"   [MAP] {table_name}: TAXNUM = coalesce({short_c or '-'}, {long_c or '-'})"
                )

    # SAP-имя в rules (NAME_ORG1): подтянуть физическую колонку выгрузки (Name_1) по column_map + шаблону
    sap_targets: set[str] = set()
    for logical, physical in tm.items():
        if str(logical).startswith("_"):
            continue
        phys = str(physical).strip()
        if phys:
            sap_targets.add(phys)
    for phys in sorted(sap_targets):
        if phys in out.columns:
            continue
        src = resolve_column(out.columns, phys, table_name, column_map, project_root)
        if not src:
            src = _reverse_scan_columns(out.columns, phys, tm, column_map, table_name)
        if src and src != phys:
            out[phys] = out[src]
            added.append((src, phys))
            taken.add(_norm(phys))

    if tn_u in ("ZBUT0000P3VVI9", "ZBUT0000P", "ZBUT0000P3VV19"):
        out = _expand_zbut_territory_from_composite_date_to(out, tm, log_renames=log_renames)

    if added and log_renames:
        shown = added[:12]
        extra = f" и ещё {len(added) - len(shown)}" if len(added) > len(shown) else ""
        print(
            f"   [MAP] {table_name}: для правил добавлены SAP-колонки ({len(added)}), "
            f"исходные заголовки сохранены: "
            + ", ".join(f"{a!r}->{b!r}" for a, b in shown)
            + extra
        )
    return out
