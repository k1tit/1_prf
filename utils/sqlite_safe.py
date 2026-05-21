"""
Единые настройки SQLite: таймаут ожидания блокировки и busy_timeout.
Снижает OperationalError: database is locked при конкурирующих процессах/забытых соединениях.
"""
from __future__ import annotations

import os
import sqlite3
import time
from typing import Optional, Tuple

# Имя файла БД по умолчанию в корне проекта (менять здесь и при необходимости в скриптах с хардкодом)
DEFAULT_DB_FILENAME = "db_april.db"

# Секунды: сколько ждать освобождения файла при connect()
SQLITE_CONNECT_TIMEOUT_SEC = 120.0

# Миллисекунды: PRAGMA busy_timeout (ожидание внутри API SQLite)
SQLITE_BUSY_TIMEOUT_MS = 120_000


def find_project_root(start_file: str) -> str:
    """
    Каталог проекта: ищем вверх по дереву каталог с main.py (корень data_quality_checker).
    """
    cur = os.path.dirname(os.path.abspath(start_file))
    for _ in range(8):
        if os.path.isfile(os.path.join(cur, "main.py")):
            return cur
        parent = os.path.dirname(cur)
        if parent == cur:
            break
        cur = parent
    return os.path.dirname(os.path.abspath(start_file))


def default_db_path(project_file: str) -> str:
    """Путь к БД в корне проекта относительно любого файла внутри репо (обычно __file__)."""
    root = find_project_root(project_file)
    return os.path.join(root, DEFAULT_DB_FILENAME)


def connect_sqlite(
    db_path: str,
    *,
    timeout: Optional[float] = None,
    busy_timeout_ms: Optional[int] = None,
    **kwargs,
) -> sqlite3.Connection:
    """
    Подключение с ожиданием блокировки и PRAGMA busy_timeout.
    kwargs передаются в sqlite3.connect (например detect_types, isolation_level).
    """
    t = SQLITE_CONNECT_TIMEOUT_SEC if timeout is None else timeout
    bt = SQLITE_BUSY_TIMEOUT_MS if busy_timeout_ms is None else busy_timeout_ms
    conn = sqlite3.connect(db_path, timeout=t, **kwargs)
    try:
        conn.execute(f"PRAGMA busy_timeout = {int(bt)}")
    except Exception:
        pass
    return conn


def probe_db_writable(
    db_path: str,
    *,
    retries: int = 8,
    sleep_sec: float = 1.5,
) -> Tuple[bool, Optional[BaseException]]:
    """
    Проверка, что БД доступна для записи (BEGIN IMMEDIATE).
    При database is locked — несколько попыток с паузой (другой процесс может отпустить файл).
    """
    last: Optional[BaseException] = None
    for _ in range(max(1, retries)):
        conn = None
        try:
            conn = connect_sqlite(db_path)
            conn.execute("BEGIN IMMEDIATE")
            conn.execute("ROLLBACK")
            return True, None
        except sqlite3.OperationalError as e:
            last = e
            msg = str(e).lower()
            if "locked" in msg or "busy" in msg:
                time.sleep(sleep_sec)
                continue
            return False, e
        except Exception as e:
            return False, e
        finally:
            if conn is not None:
                try:
                    conn.close()
                except Exception:
                    pass
    return False, last


def is_lock_error(exc: BaseException) -> bool:
    s = str(exc).lower()
    return "locked" in s or "busy" in s
