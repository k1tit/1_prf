#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sqlite3
import os
import csv
import time
import sys
from datetime import datetime

# Настройка кодировки вывода для Windows (чтобы не было кракозябр)
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')

_PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

from utils.sqlite_safe import resolve_database_path

# Та же БД, что и у main.py (config/database.json → DQ_DATABASE → --db при обёртке)
DB_PATH, DB_SOURCE = resolve_database_path(_PROJECT_ROOT)
RULE_CODE = "RCCOMP_103.1"
RULE_DESC = "Missing Customer Activity Cluster"
RULE_CATEGORY = "Completeness"
CHECK_COLUMN = "KATR1"
CONDITION_COLUMN = "account_group_code"
CONDITION_VALUE = "9038"

# Папки
CHECK_RULES_DIR = "check_rules"
ERR_DIR = "err"
os.makedirs(CHECK_RULES_DIR, exist_ok=True)
os.makedirs(ERR_DIR, exist_ok=True)


def find_table_with_columns(conn, required_columns):
    """Ищет таблицу, содержащую все указанные колонки."""
    cursor = conn.cursor()
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
    tables = cursor.fetchall()
    for (tbl,) in tables:
        cursor.execute(f"PRAGMA table_info({tbl})")
        existing_cols = {row[1] for row in cursor.fetchall()}
        if all(col in existing_cols for col in required_columns):
            return tbl
    return None


def run_check():
    start_time = time.time()
    result = {
        "total": 0, "passed": 0, "failed": 0,
        "all_data": [], "failed_data": [],
        "comment": "", "status": ""
    }

    conn = None
    try:
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        # Определяем реальное имя таблицы
        required_cols = [CHECK_COLUMN, CONDITION_COLUMN]
        table_name = find_table_with_columns(conn, required_cols)

        if not table_name:
            # Выводим список всех таблиц и их колонок для диагностики
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
            all_tables = cursor.fetchall()
            error_msg = f"Не найдена таблица, содержащая колонки {required_cols}. Доступные таблицы:\n"
            for (tbl,) in all_tables:
                cursor.execute(f"PRAGMA table_info({tbl})")
                cols = [row[1] for row in cursor.fetchall()]
                error_msg += f"  - {tbl}: {', '.join(cols)}\n"
            raise Exception(error_msg)

        print(f"✅ Найдена таблица: {table_name}")

        # Проверяем условие: только строки с account_group_code = '9038'
        query = f"""
            SELECT * FROM {table_name}
            WHERE {CONDITION_COLUMN} = ?
        """
        cursor.execute(query, (CONDITION_VALUE,))
        rows = cursor.fetchall()
        total = len(rows)
        result["total"] = total

        if total == 0:
            result["comment"] = f"Нет записей с {CONDITION_COLUMN} = '{CONDITION_VALUE}'"
            result["status"] = "Нет данных"
            return result

        passed = 0
        failed = 0
        all_data = []
        failed_data = []

        for row in rows:
            row_dict = dict(row)
            val = row_dict.get(CHECK_COLUMN)
            # Успех: значение не NULL и не пустая строка
            is_ok = val is not None and (not isinstance(val, str) or val.strip() != "")
            if is_ok:
                passed += 1
                row_dict['_check_status'] = 1
            else:
                failed += 1
                row_dict['_check_status'] = 0
                failed_data.append(row_dict)
            all_data.append(row_dict)

        result["passed"] = passed
        result["failed"] = failed
        result["all_data"] = all_data
        result["failed_data"] = failed_data
        result["status"] = "Успех" if failed == 0 else "Ошибка"

    except Exception as e:
        result["comment"] = str(e)
        result["status"] = "Ошибка выполнения"
        # Сохраняем лог ошибки в папку err
        with open(os.path.join(ERR_DIR, f"{RULE_CODE}_error.log"), "w", encoding="utf-8") as f:
            f.write(f"[{datetime.now()}] {e}\n")
    finally:
        if conn:
            conn.close()

    result["elapsed"] = round(time.time() - start_time, 2)
    return result, table_name if 'table_name' in locals() else None


def save_csv(data, filepath):
    if not data:
        return
    with open(filepath, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=data[0].keys())
        writer.writeheader()
        writer.writerows(data)
    print(f"💾 Сохранено: {filepath} ({len(data)} записей)")


def main():
    print("🔍 Проверка правила RCCOMP_103.1")
    print(f"📁 База: {DB_PATH}")
    print(f"   ({DB_SOURCE})")

    result, used_table = run_check()

    # Сохраняем все проверенные строки
    if result["all_data"]:
        all_path = os.path.join(CHECK_RULES_DIR, f"{RULE_CODE}_all_checked.csv")
        save_csv(result["all_data"], all_path)

    # Сохраняем только ошибочные
    if result["failed_data"]:
        err_path = os.path.join(ERR_DIR, f"{RULE_CODE}_failed.csv")
        save_csv(result["failed_data"], err_path)

    # Формируем отчёт
    percent = (result["passed"] / result["total"] * 100) if result["total"] > 0 else 0.0
    report = {
        "Код правила": RULE_CODE,
        "Описание": RULE_DESC,
        "Категория": RULE_CATEGORY,
        "Таблица": used_table if used_table else "Не найдена",
        "Тип TAXNUM": "",
        "Колонка": CHECK_COLUMN,
        "Всего записей": result["total"],
        "Успешно": result["passed"],
        "Ошибок": result["failed"],
        "% успеха": f"{percent:.2f}",
        "Статус": result["status"],
        "Время (сек)": result["elapsed"],
        "Файл ошибок": f"{ERR_DIR}/{RULE_CODE}_failed.csv" if result["failed"] > 0 else "",
        "Комментарии": result["comment"],
        "Список записей (файл)": f"{CHECK_RULES_DIR}/{RULE_CODE}_all_checked.csv" if result["total"] > 0 else "",
    }

    # Вывод в консоль (таблица)
    print("\n" + "=" * 120)
    print("ОТЧЁТ ПО ПРОВЕРКЕ")
    print("=" * 120)
    for key, val in report.items():
        print(f"{key:25} : {val}")
    print("=" * 120)

    # Сохраняем отчёт как CSV
    report_path = os.path.join(CHECK_RULES_DIR, f"{RULE_CODE}_report.csv")
    with open(report_path, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=report.keys())
        writer.writeheader()
        writer.writerow(report)
    print(f"\n📄 Отчёт сохранён: {report_path}")

    if result["status"] == "Ошибка выполнения":
        sys.exit(1)


if __name__ == "__main__":
    main()