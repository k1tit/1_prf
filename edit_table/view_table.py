#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Просмотр SQLite БД (то же, что view_bd.py).
Запуск: python edit_table/view_table.py
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
    here = os.path.dirname(os.path.abspath(__file__))
    parent = os.path.dirname(here)
    for candidate in (here, parent):
        if os.path.isfile(os.path.join(candidate, "utils", "sqlite_safe.py")):
            if candidate not in sys.path:
                sys.path.insert(0, candidate)
            return candidate
    if parent not in sys.path:
        sys.path.insert(0, parent)
    return parent


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
    return resolve_database_path(_DQ_ROOT, cli_path, must_exist=True)


def list_tables(conn) -> list[str]:
    df = pd.read_sql_query(
        "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name",
        conn,
    )
    return df["name"].tolist() if not df.empty else []


def table_row_count(conn, table_name: str) -> int:
    safe = table_name.replace('"', "")
    return int(pd.read_sql_query(f'SELECT COUNT(*) AS c FROM "{safe}"', conn)["c"].iloc[0])


def show_table_info(conn, table_name: str) -> None:
    cols = pd.read_sql_query(f'PRAGMA table_info("{table_name}")', conn)
    print(f"\nТаблица: {table_name}")
    try:
        print(f"Строк: {table_row_count(conn, table_name):,}")
    except Exception as e:
        print(f"Строк: не удалось посчитать ({e})")
    if not cols.empty:
        print("Колонки:", ", ".join(cols["name"].astype(str).tolist()))


def preview_table(conn, table_name: str, limit: int = 20) -> None:
    show_table_info(conn, table_name)
    df = pd.read_sql_query(f'SELECT * FROM "{table_name}" LIMIT {int(limit)}', conn)
    print(f"\nПервые {limit} строк:")
    print(df.to_string(index=False) if not df.empty else "(пусто)")


def interactive_loop(db_path: str, default_limit: int = 20) -> None:
    print(f"Подключение: {db_path}")
    conn = connect_sqlite(db_path)
    try:
        while True:
            tables = list_tables(conn)
            if not tables:
                print("В базе нет таблиц.")
                return
            print("\n" + "=" * 60)
            for i, t in enumerate(tables, 1):
                try:
                    print(f"  {i:3d}. {t} ({table_row_count(conn, t):,} строк)")
                except Exception:
                    print(f"  {i:3d}. {t}")
            cmd = input("\nНомер / q выход: ").strip().lower()
            if cmd in ("q", "quit"):
                break
            if not cmd.isdigit() or int(cmd) < 1 or int(cmd) > len(tables):
                continue
            preview_table(conn, tables[int(cmd) - 1], default_limit)
    finally:
        conn.close()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--db")
    ap.add_argument("--table")
    ap.add_argument("--limit", type=int, default=20)
    ap.add_argument("-i", "--interactive", action="store_true")
    args = ap.parse_args()
    try:
        db_path, source = resolve_db(args.db)
    except FileNotFoundError as e:
        print(f"ОШИБКА: {e}\nКорень: {_DQ_ROOT}")
        sys.exit(2)
    print(f"Корень: {_DQ_ROOT}\nБД ({source}): {db_path}")
    if args.table:
        conn = connect_sqlite(db_path)
        try:
            preview_table(conn, args.table, args.limit)
        finally:
            conn.close()
        return
    if args.interactive or len(sys.argv) == 1:
        interactive_loop(db_path, args.limit)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nПрервано.")
