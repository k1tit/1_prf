#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Просмотр SQLite БД проекта проверки качества данных.
Использует ту же БД, что main.py / add_table.py: config/database.json в job_project_clean.
"""

from __future__ import annotations

import argparse
import os
import sys

import pandas as pd

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass


def _bootstrap_sys_path():
    """Добавить в sys.path каталог проекта с utils/ (до импорта sqlite_safe)."""
    here = os.path.dirname(os.path.abspath(__file__))
    parent = os.path.dirname(here)
    for candidate in (
        here,
        parent,
        os.path.join(parent, "data_quality_checker"),
        os.path.join(parent, "job_project_clean"),
        os.path.join(here, "job_project_clean"),
    ):
        if os.path.isfile(os.path.join(candidate, "utils", "sqlite_safe.py")):
            if candidate not in sys.path:
                sys.path.insert(0, candidate)
            return candidate
    if here not in sys.path:
        sys.path.insert(0, here)
    return here


_bootstrap_sys_path()

from utils.sqlite_safe import connect_sqlite, resolve_database_path  # noqa: E402

try:
    from utils.sqlite_safe import find_dq_project_root  # noqa: E402
except ImportError:
    from utils.sqlite_safe import find_project_root as find_dq_project_root  # noqa: E402

_DQ_ROOT = find_dq_project_root(__file__)
if _DQ_ROOT not in sys.path:
    sys.path.insert(0, _DQ_ROOT)


def resolve_db(cli_path: str | None = None) -> tuple[str, str]:
    path, source = resolve_database_path(_DQ_ROOT, cli_path, must_exist=True)
    return path, source


def list_tables(conn) -> list[str]:
    df = pd.read_sql_query(
        "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name",
        conn,
    )
    return df["name"].tolist() if not df.empty else []


def table_row_count(conn, table_name: str) -> int:
    q = f'SELECT COUNT(*) AS c FROM "{table_name.replace(chr(34), "")}"'
    return int(pd.read_sql_query(q, conn)["c"].iloc[0])


def show_table_info(conn, table_name: str) -> None:
    cols = pd.read_sql_query(f'PRAGMA table_info("{table_name}")', conn)
    print(f"\nТаблица: {table_name}")
    try:
        n = table_row_count(conn, table_name)
        print(f"Строк: {n:,}")
    except Exception as e:
        print(f"Строк: не удалось посчитать ({e})")
    if not cols.empty:
        print("Колонки:", ", ".join(cols["name"].astype(str).tolist()))


def preview_table(conn, table_name: str, limit: int = 20) -> None:
    show_table_info(conn, table_name)
    df = pd.read_sql_query(f'SELECT * FROM "{table_name}" LIMIT {int(limit)}', conn)
    print(f"\nПервые {limit} строк:")
    if df.empty:
        print("(пусто)")
    else:
        pd.set_option("display.max_columns", 20)
        pd.set_option("display.width", 200)
        print(df.to_string(index=False))


def search_in_table(conn, table_name: str, column: str, value: str, limit: int = 10) -> None:
    cols = pd.read_sql_query(f'PRAGMA table_info("{table_name}")', conn)
    names = cols["name"].astype(str).tolist()
    if column not in names:
        print(f"Колонка '{column}' не найдена. Доступно: {', '.join(names)}")
        return
    safe_val = value.replace("'", "''")
    q = (
        f'SELECT * FROM "{table_name}" '
        f'WHERE CAST("{column}" AS TEXT) LIKE \'%{safe_val}%\' LIMIT {int(limit)}'
    )
    df = pd.read_sql_query(q, conn)
    print(f"\nПоиск {column} LIKE '%{value}%': найдено {len(df)} (лимит {limit})")
    if not df.empty:
        print(df.to_string(index=False))
    else:
        print("(ничего)")


def interactive_loop(db_path: str, default_limit: int = 20) -> None:
    print(f"Подключение: {db_path}")
    conn = connect_sqlite(db_path)
    try:
        while True:
            tables = list_tables(conn)
            if not tables:
                print("В базе нет пользовательских таблиц.")
                return

            print("\n" + "=" * 60)
            print("Таблицы в БД:")
            for i, t in enumerate(tables, 1):
                try:
                    n = table_row_count(conn, t)
                    print(f"  {i:3d}. {t} ({n:,} строк)")
                except Exception:
                    print(f"  {i:3d}. {t}")

            cmd = input(
                "\nНомер таблицы / 'q' выход / 'r' обновить список: "
            ).strip().lower()
            if cmd in ("q", "quit", "exit"):
                break
            if cmd in ("r", "refresh", ""):
                continue
            if not cmd.isdigit():
                print("Введите номер таблицы.")
                continue
            idx = int(cmd)
            if idx < 1 or idx > len(tables):
                print("Неверный номер.")
                continue
            table_name = tables[idx - 1]

            while True:
                preview_table(conn, table_name, default_limit)
                print("\n  [1] Другой лимит строк")
                print("  [2] Поиск по колонке")
                print("  [3] Назад к списку таблиц")
                sub = input("Действие [3]: ").strip() or "3"
                if sub == "1":
                    try:
                        lim = int(input("Сколько строк показать: ").strip() or str(default_limit))
                        preview_table(conn, table_name, lim)
                    except ValueError:
                        print("Введите число.")
                elif sub == "2":
                    col = input("Колонка: ").strip()
                    val = input("Значение (подстрока): ").strip()
                    if col and val:
                        search_in_table(conn, table_name, col, val)
                elif sub == "3":
                    break
                else:
                    print("Неизвестная команда.")
    finally:
        conn.close()
    print("\nГотово.")


def parse_args():
    p = argparse.ArgumentParser(description="Просмотр SQLite БД (config/database.json)")
    p.add_argument("--db", metavar="PATH", help="Путь к .db (переопределяет config)")
    p.add_argument("--table", metavar="NAME", help="Сразу показать таблицу и выйти")
    p.add_argument("--limit", type=int, default=20, help="Строк в превью (по умолчанию 20)")
    p.add_argument("-i", "--interactive", action="store_true", help="Интерактивный режим")
    return p.parse_args()


def main():
    args = parse_args()
    try:
        db_path, source = resolve_db(args.db)
    except FileNotFoundError as e:
        print(f"ОШИБКА: {e}")
        print(f"Корень проекта DQ: {_DQ_ROOT}")
        sys.exit(2)

    print(f"Корень проекта: {_DQ_ROOT}")
    print(f"БД ({source}): {db_path}")

    if args.table:
        conn = connect_sqlite(db_path)
        try:
            tables = list_tables(conn)
            if args.table not in tables:
                print(f"Таблица '{args.table}' не найдена. Есть: {', '.join(tables[:20])}")
                sys.exit(1)
            preview_table(conn, args.table, args.limit)
        finally:
            conn.close()
        return

    if args.interactive or len(sys.argv) == 1:
        interactive_loop(db_path, args.limit)
        return

    conn = connect_sqlite(db_path)
    try:
        tables = list_tables(conn)
        print("\nТаблицы:", ", ".join(tables) if tables else "(нет)")
        print("Для интерактива: python view_bd.py -i")
    finally:
        conn.close()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nПрервано.")
