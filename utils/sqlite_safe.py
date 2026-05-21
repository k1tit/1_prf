"""
Единые настройки SQLite: таймаут ожидания блокировки и busy_timeout.
Снижает OperationalError: database is locked при конкурирующих процессах/забытых соединениях.
"""
from __future__ import annotations

import json
import os
import sqlite3
import time
from typing import Optional, Tuple

# Запасной файл БД, если нет config/database.json и переменной DQ_DATABASE
DEFAULT_DB_FILENAME = "db_april.db"

# Относительный путь к конфигу смены БД (одно место на месяц)
DB_CONFIG_REL = os.path.join("config", "database.json")

# Переменная окружения: имя файла или полный путь к SQLite
ENV_DB_VAR = "DQ_DATABASE"

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
    path, _ = resolve_database_path(root)
    return path


def _normalize_db_spec(project_root: str, spec: str) -> str:
    """Имя файла или путь → абсолютный путь к SQLite."""
    spec = (spec or "").strip()
    if not spec:
        raise ValueError("пустое имя/путь к базе данных")
    if os.path.isabs(spec):
        return os.path.normpath(spec)
    return os.path.normpath(os.path.join(project_root, spec))


def load_database_config(project_root: str) -> dict:
    """Читает config/database.json; при отсутствии файла — пустой dict."""
    path = os.path.join(project_root, DB_CONFIG_REL)
    if not os.path.isfile(path):
        return {}
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (json.JSONDecodeError, OSError):
        return {}


def resolve_database_path(
    project_root: str,
    cli_path: Optional[str] = None,
    *,
    must_exist: bool = False,
) -> Tuple[str, str]:
    """
    Единая точка выбора файла SQLite на месяц.

    Приоритет:
      1) аргумент CLI (--db)
      2) переменная окружения DQ_DATABASE
      3) поле database в config/database.json
      4) DEFAULT_DB_FILENAME в корне проекта

    Возвращает (абсолютный_путь, описание_источника).
    """
    if cli_path and str(cli_path).strip():
        path = _normalize_db_spec(project_root, str(cli_path))
        source = "аргумент --db"
    elif os.environ.get(ENV_DB_VAR, "").strip():
        path = _normalize_db_spec(project_root, os.environ[ENV_DB_VAR])
        source = f"переменная окружения {ENV_DB_VAR}"
    else:
        cfg = load_database_config(project_root)
        db_spec = (cfg.get("database") or "").strip()
        if db_spec:
            path = _normalize_db_spec(project_root, db_spec)
            period = (cfg.get("period") or "").strip()
            source = f"config/{DB_CONFIG_REL.replace(os.sep, '/')}"
            if period:
                source += f" (period={period})"
        else:
            path = _normalize_db_spec(project_root, DEFAULT_DB_FILENAME)
            source = f"запасной DEFAULT_DB_FILENAME ({DEFAULT_DB_FILENAME})"

    if must_exist and not os.path.isfile(path):
        raise FileNotFoundError(
            f"Файл базы данных не найден: {path}\n"
            f"Источник: {source}. Положите .db в корень проекта и обновите "
            f"{DB_CONFIG_REL} (поле database) или задайте {ENV_DB_VAR} / --db."
        )
    return path, source


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
