# main.py
import os
import sys
import argparse
import logging
from datetime import datetime, time

# Базовый каталог проекта (где лежит main.py)
_PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))

DB_PATH = os.path.join(_PROJECT_ROOT, "db_mrt.db")
RULES_FILE = os.path.join(_PROJECT_ROOT, "json files", "rules.json")
OUTPUT_DIR = os.path.join(_PROJECT_ROOT, "quality_reports")

def print_project_info():
    """Выводит информацию о проекте"""
    print(f"{'='*80}")
    print("СИСТЕМА ПРОВЕРКИ КАЧЕСТВА ДАННЫХ")
    print(f"{'='*80}")
    print(f"Версия: 2.0")
    print(f"Дата: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"{'-'*80}")

def setup_environment():
    """Настраивает окружение Python"""
    current_dir = os.path.dirname(os.path.abspath(__file__))
    print(f"Рабочая директория: {current_dir}")
    
    if current_dir not in sys.path:
        sys.path.insert(0, current_dir)
    
    # Проверка структуры проекта (config — опционально)
    required_dirs = ["core", "utils", "validators", "table_scripts"]
    for dir_name in required_dirs:
        dir_path = os.path.join(current_dir, dir_name)
        if os.path.exists(dir_path):
            files = [f for f in os.listdir(dir_path) if f.endswith('.py')]
            print(f"{dir_name}/: {len(files)} файлов")
        else:
            print(f"{dir_name}/: НЕ НАЙДЕН!")
    if os.path.isdir(os.path.join(current_dir, "config")):
        print(f"config/: найден")
    # config/ не обязателен — маппинг может лежать в json files/ (column_map.json)
    
    print(f"{'-'*80}")
    
    # Проверка необходимых файлов
    required_files = [
        (DB_PATH, "База данных SQLite"),
        (RULES_FILE, "Файл правил JSON"),
    ]
    
    for file_path, description in required_files:
        full_path = os.path.join(current_dir, file_path) if not os.path.isabs(file_path) else file_path
        if os.path.exists(full_path):
            size_kb = os.path.getsize(full_path) / 1024
            print(f"{description}: {file_path} ({size_kb:.1f} KB)")
        else:
            print(f"{description}: {file_path} - НЕ НАЙДЕН!")
    
    # Проверка дополнительных файлов (маппинг: config/ или json files/)
    column_map_candidates = [
        os.path.join(current_dir, "config", "column_map.json"),
        os.path.join(current_dir, "json files", "column_map.json"),
    ]
    column_map_found = None
    for p in column_map_candidates:
        if os.path.exists(p):
            column_map_found = p
            break
    if column_map_found:
        size_kb = os.path.getsize(column_map_found) / 1024
        rel = os.path.relpath(column_map_found, current_dir)
        print(f"Файл маппинга колонок: {rel} ({size_kb:.1f} KB)")
    else:
        print(f"Файл маппинга колонок: не найден (искали config/column_map.json и json files/column_map.json), будет использоваться стандартный маппинг")
    
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    print(f"Выходная директория: {OUTPUT_DIR}")
    
    return current_dir

def load_checker_module():
    """Загружает модуль checker из core.checker (нормальный import, без exec)."""
    print(f"\n{'='*80}")
    print("ИМПОРТ МОДУЛЕЙ")
    print(f"{'='*80}")
    
    current_dir = os.path.dirname(os.path.abspath(__file__))
    checker_path = os.path.join(current_dir, "core", "checker.py")
    
    if not os.path.exists(checker_path):
        print(f"ФАТАЛЬНАЯ ОШИБКА: Файл {checker_path} не найден!")
        sys.exit(1)
    
    # Чтобы подхватить изменения в core/checker.py: сбрасываем кэш .pyc и кэш модулей
    cache_dir = os.path.join(current_dir, "core", "__pycache__")
    if os.path.isdir(cache_dir):
        for name in os.listdir(cache_dir):
            if name.startswith("checker.") and name.endswith(".pyc"):
                try:
                    os.remove(os.path.join(cache_dir, name))
                except OSError:
                    pass
    for mod in ("core.checker", "core"):
        sys.modules.pop(mod, None)
    
    try:
        print("Загружаем core.checker...")
        if current_dir not in sys.path:
            sys.path.insert(0, current_dir)
        from core.checker import FastDataQualityChecker
        print("FastDataQualityChecker загружен успешно")
        return FastDataQualityChecker
    except Exception as e:
        print(f"ОШИБКА ЗАГРУЗКИ МОДУЛЯ: {type(e).__name__}: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

def list_tables(checker):
    """Выводит список таблиц доступных для проверки"""
    print(f"\n{'='*80}")
    print("ДОСТУПНЫЕ ТАБЛИЦЫ ДЛЯ ПРОВЕРКИ")
    print(f"{'='*80}")
    
    tables = checker.list_available_tables()
    
    if tables:
        print(f"Всего таблиц: {len(tables)}")
        print(f"{'-'*80}")
        
        for i, table in enumerate(tables, 1):
            rules = checker.get_table_rules(table)
            print(f"{i:3d}. {table:25} - {len(rules):3d} правил")
        
        print(f"{'='*80}")
    else:
        print("[!] Нет доступных таблиц для проверки")
    
    return tables

def _refresh_handlers_before_run(checker):
    """В интерактивном режиме подхватываем правки в table_scripts/*_handler без перезапуска Python."""
    if hasattr(checker, "reload_table_handlers"):
        checker.reload_table_handlers()


def run_full_check(checker):
    """Запускает полную проверку всех таблиц"""
    print(f"\n{'='*80}")
    print("ЗАПУСК ПОЛНОЙ ПРОВЕРКИ ВСЕХ ТАБЛИЦ")
    print(f"{'='*80}")
    
    start_time = datetime.now()
    print(f"Начало: {start_time.strftime('%H:%M:%S')}")
    
    try:
        _refresh_handlers_before_run(checker)
        checker.run()  # Без аргументов - проверяем все таблицы
    except Exception as e:
        print(f"ОШИБКА ВЫПОЛНЕНИЯ: {type(e).__name__}: {e}")
        import traceback
        traceback.print_exc()
        return False
    
    end_time = datetime.now()
    elapsed = end_time - start_time
    
    print(f"\n{'='*80}")
    print("ПРОВЕРКА ЗАВЕРШЕНА")
    print(f"{'='*80}")
    print(f"Начало: {start_time.strftime('%H:%M:%S')}")
    print(f"Конец:  {end_time.strftime('%H:%M:%S')}")
    print(f"Длительность: {elapsed}")
    
    return True

def _parse_only_rule_codes(only_rules_arg):
    """Список кодов правил из строки --only-rules или None."""
    if not only_rules_arg:
        return None
    codes = {s.strip() for s in only_rules_arg.split(',') if s.strip()}
    return codes if codes else None


def parse_reference_date_string(s):
    """
    Дата снимка данных для правил «на дату» (например RCCONF_173.1).
    Возвращает datetime (конец указанного календарного дня 23:59:59) или None.
    Форматы: YYYY-MM-DD, DD.MM.YYYY
    """
    if s is None or not str(s).strip():
        return None
    s = str(s).strip()
    for fmt in ("%Y-%m-%d", "%d.%m.%Y"):
        try:
            d = datetime.strptime(s, fmt).date()
            return datetime.combine(d, time(23, 59, 59))
        except ValueError:
            continue
    raise ValueError(f"Неверный формат даты: {s!r} (ожидается YYYY-MM-DD или DD.MM.YYYY)")


def prompt_reference_datetime(checker):
    """Интерактивный выбор: текущее время компьютера или введённая дата."""
    print("\nОпорная дата для расчётов «на дату» (например RCCONF_173.1 — срок с даты назначения блока):")
    print("  [Enter] — текущие дата и время компьютера")
    print("  или введите дату снимка данных: YYYY-MM-DD или DD.MM.YYYY (учёт до конца этого дня)")
    try:
        s = input("> ").strip()
    except EOFError:
        s = ""
    if not s:
        checker.reference_datetime = None
        print("[INFO] Опорная дата: текущее время системы")
        return
    try:
        checker.reference_datetime = parse_reference_date_string(s)
        print(f"[INFO] Опорная дата (конец дня): {checker.reference_datetime}")
    except ValueError as e:
        print(f"[WARN] {e}. Используется текущее время системы.")
        checker.reference_datetime = None


def run_table_check(checker, table_name, only_rule_codes=None):
    """Запускает проверку конкретной таблицы. only_rule_codes — опционально ограничить правила."""
    print(f"\n{'='*80}")
    print(f"ЗАПУСК ПРОВЕРКИ ТАБЛИЦЫ: {table_name}")
    print(f"{'='*80}")
    
    start_time = datetime.now()
    print(f"Начало: {start_time.strftime('%H:%M:%S')}")
    
    try:
        # Проверяем что таблица существует в правилах
        rules = checker.get_table_rules(table_name)
        if not rules:
            print(f"ОШИБКА: Для таблицы '{table_name}' нет правил!")
            return False
        
        print(f"Найдено правил: {len(rules)}")
        print(f"{'-'*80}")
        
        for i, rule in enumerate(rules, 1):
            rule_desc = rule.get('rule_description', 'Без описания')
            if len(rule_desc) > 50:
                rule_desc = rule_desc[:47] + "..."
            print(f"{i:3d}. {rule.get('rule_code', 'N/A'):15} - {rule_desc}")
        
        print(f"{'-'*80}")
        if only_rule_codes:
            print(f"Только правила: {sorted(only_rule_codes)}")
        print("Запускаем проверку...")
        
        _refresh_handlers_before_run(checker)
        checker.run(specific_table=table_name, only_rule_codes=only_rule_codes)
        
    except Exception as e:
        print(f"[!] ОШИБКА ВЫПОЛНЕНИЯ: {type(e).__name__}: {e}")
        import traceback
        traceback.print_exc()
        return False
    
    end_time = datetime.now()
    elapsed = end_time - start_time
    
    print(f"\n{'='*80}")
    print(f"ПРОВЕРКА ТАБЛИЦЫ {table_name} ЗАВЕРШЕНА")
    print(f"{'='*80}")
    print(f"Начало: {start_time.strftime('%H:%M:%S')}")
    print(f"Конец:  {end_time.strftime('%H:%M:%S')}")
    print(f"Длительность: {elapsed}")
    
    return True

def run_selected_tables_check(checker, table_names, only_rule_codes=None):
    """Запускает проверку для выбранных таблиц. only_rule_codes — опционально."""
    print(f"\n{'='*80}")
    print(f"ЗАПУСК ПРОВЕРКИ ДЛЯ ВЫБРАННЫХ ТАБЛИЦ")
    print(f"{'='*80}")
    
    print(f"Таблицы для проверки: {', '.join(table_names)}")
    print(f"Количество таблиц: {len(table_names)}")
    print(f"{'-'*80}")
    
    start_time = datetime.now()
    print(f"Начало: {start_time.strftime('%H:%M:%S')}")
    
    try:
        _refresh_handlers_before_run(checker)
        import inspect
        sig = inspect.signature(checker.run)
        if "table_list" in sig.parameters:
            checker.run(table_list=table_names, only_rule_codes=only_rule_codes)
        else:
            for name in table_names:
                checker.run(specific_table=name, only_rule_codes=only_rule_codes)
        
    except Exception as e:
        print(f"ОШИБКА ВЫПОЛНЕНИЯ: {type(e).__name__}: {e}")
        import traceback
        traceback.print_exc()
        return False
    
    end_time = datetime.now()
    elapsed = end_time - start_time
    
    print(f"\n{'='*80}")
    print(f"ПРОВЕРКА ВЫБРАННЫХ ТАБЛИЦ ЗАВЕРШЕНА")
    print(f"{'='*80}")
    print(f"Начало: {start_time.strftime('%H:%M:%S')}")
    print(f"Конец:  {end_time.strftime('%H:%M:%S')}")
    print(f"Длительность: {elapsed}")
    
    return True

def interactive_mode(checker):
    """Интерактивный режим с выбором таблицы"""
    print(f"\n{'='*80}")
    print("ИНТЕРАКТИВНЫЙ РЕЖИМ")
    print(f"{'='*80}")
    
    while True:
        print("\nДоступные команды:")
        print("  [L] - Список таблиц")
        print("  [F] - Полная проверка всех таблиц")
        print("  [1] - Проверить таблицу по номеру")
        print("  [N] - Проверить таблицу по имени")
        print("  [M] - Проверить несколько таблиц")
        print("  [Q] - Выход")
        print("  Перед проверкой F/1/N/M можно задать опорную дату для правил «на дату» (например RCCONF_173.1)")
        print("  (перед каждой проверкой обработчики KNA1/BUT000/... перезагружаются с диска)")
        print(f"{'-'*80}")
        
        choice = input("Выберите действие: ").strip().upper()
        
        if choice == 'Q':
            print("Выход из программы...")
            break
        
        elif choice == 'L':
            list_tables(checker)
        
        elif choice == 'F':
            confirm = input("Запустить полную проверку всех таблиц? (y/N): ").strip().upper()
            if confirm == 'Y':
                prompt_reference_datetime(checker)
                run_full_check(checker)
        
        elif choice == '1':
            tables = list_tables(checker)
            if tables:
                try:
                    table_num = int(input("Введите номер таблицы: ").strip())
                    if 1 <= table_num <= len(tables):
                        table_name = tables[table_num - 1]
                        prompt_reference_datetime(checker)
                        run_table_check(checker, table_name)
                    else:
                        print(f"Неверный номер. Допустимый диапазон: 1-{len(tables)}")
                except ValueError:
                    print("Введите число!")
        
        elif choice == 'N':
            table_name = input("Введите имя таблицы: ").strip()
            if table_name:
                prompt_reference_datetime(checker)
                run_table_check(checker, table_name)
            else:
                print("Имя таблицы не может быть пустым!")
        
        elif choice == 'M':
            tables = list_tables(checker)
            if tables:
                print(f"\nВведите номера таблиц через пробел (например: 1 3 5):")
                try:
                    input_str = input("Номера таблиц: ").strip()
                    if input_str:
                        numbers = [int(n) for n in input_str.split()]
                        selected_tables = []
                        for num in numbers:
                            if 1 <= num <= len(tables):
                                selected_tables.append(tables[num - 1])
                            else:
                                print(f"Неверный номер {num}. Допустимый диапазон: 1-{len(tables)}")
                        
                        if selected_tables:
                            prompt_reference_datetime(checker)
                            run_selected_tables_check(checker, selected_tables)
                        else:
                            print("Не выбрано ни одной таблицы!")
                except ValueError:
                    print("Введите числа через пробел!")
        
        else:
            print("Неизвестная команда. Попробуйте еще раз.")

def parse_arguments():
    """Парсит аргументы командной строки"""
    parser = argparse.ArgumentParser(
        description='Система проверки качества данных',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Примеры использования:
  python main.py                    # Запуск в интерактивном режиме
  python main.py --all              # Проверка всех таблиц
  python main.py --all --async-load # Загрузка таблиц асинхронно (быстрее при многих таблицах)
  python main.py --all --parallel-tables 4   # Параллельная обработка 4 таблиц
  python main.py --table KNA1       # Проверка только таблицы KNA1
  python main.py --table BUT000 --only-rules RCCONF_15.1  # Одна таблица + только эти правила
  python main.py --table KNA1 --log-file kna1.log   # Логи KNA1 в файл
  python main.py --table KNA1 --debug               # Подробные логи (DEBUG)
  python main.py --tables KNA1 BUT000  # Проверка нескольких таблиц
  python main.py --only-rules RCCOMP_375.1,RCCONF_39.5  # Только указанные правила (по всем таблицам)
  python main.py --reference-date 2026-04-01  # Опорная дата для правил «на дату» (RCCONF_173.1 и др.)
  python main.py --list             # Показать список таблиц
  python main.py --help             # Показать эту справку
        """
    )
    
    parser.add_argument(
        '--all', 
        action='store_true',
        help='Запустить проверку всех таблиц'
    )
    
    parser.add_argument(
        '--table', 
        type=str,
        metavar='TABLE_NAME',
        help='Проверить конкретную таблицу'
    )
    
    parser.add_argument(
        '--tables', 
        type=str,
        nargs='+',
        metavar='TABLE',
        help='Проверить указанные таблицы (через пробел)'
    )
    
    parser.add_argument(
        '--list', 
        action='store_true',
        help='Показать список доступных таблиц'
    )
    
    parser.add_argument(
        '--output', 
        type=str,
        metavar='DIR',
        default=OUTPUT_DIR,
        help=f'Директория для отчетов (по умолчанию: {OUTPUT_DIR})'
    )
    
    parser.add_argument(
        '--db', 
        type=str,
        metavar='PATH',
        default=DB_PATH,
        help=f'Путь к базе данных (по умолчанию: {DB_PATH})'
    )
    
    parser.add_argument(
        '--rules', 
        type=str,
        metavar='PATH',
        default=RULES_FILE,
        help=f'Путь к файлу правил (по умолчанию: {RULES_FILE})'
    )
    
    parser.add_argument(
        '--only-rules',
        type=str,
        metavar='RULE1,RULE2,...',
        default=None,
        help='Выполнить только указанные правила (изолированный запуск). Пример: --only-rules RCCOMP_375.1,RCCONF_39.5'
    )
    
    parser.add_argument(
        '--debug',
        action='store_true',
        help='Подробное логирование (DEBUG): checker, KNA1Handler и др.'
    )
    parser.add_argument(
        '--log-file',
        type=str,
        metavar='PATH',
        default=None,
        help='Дополнительно писать логи в файл (например, kna1.log при проверке KNA1)'
    )

    parser.add_argument(
        '--reference-date',
        type=str,
        metavar='YYYY-MM-DD',
        default=None,
        help='Опорная дата снимка данных для правил «на дату» (например RCCONF_173.1). '
             'Формат: YYYY-MM-DD или DD.MM.YYYY; конец этого календарного дня. '
             'Для архивных выгрузок обязательно укажите дату актуальности данных, иначе расчёт от «сегодня» исказит результат. '
             'Без параметра — текущее время компьютера.'
    )

    return parser.parse_args()

def main():
    """Основная функция"""
    args = parse_arguments()

    # Логи при проверке таблиц (KNA1 и др.): консоль + опционально файл
    log_level = logging.DEBUG if args.debug else logging.INFO
    log_fmt = "%(levelname)s:%(name)s: %(message)s"
    logging.basicConfig(level=log_level, format=log_fmt, force=True)
    if getattr(args, 'log_file', None):
        log_path = args.log_file
        # Относительный путь → всегда в корне проекта (легко найти)
        if not os.path.isabs(log_path):
            log_path = os.path.join(_PROJECT_ROOT, log_path)
        fh = logging.FileHandler(log_path, encoding='utf-8')
        fh.setFormatter(logging.Formatter(log_fmt))
        logging.getLogger().addHandler(fh)
        print(f"[LOG] Логи пишутся в файл: {os.path.abspath(log_path)}")

    # Глобальные переменные могут быть переопределены аргументами
    global DB_PATH, RULES_FILE, OUTPUT_DIR
    if args.db:
        DB_PATH = args.db
    if args.rules:
        RULES_FILE = args.rules
    if args.output:
        OUTPUT_DIR = args.output

    reference_dt = None
    if getattr(args, "reference_date", None):
        try:
            reference_dt = parse_reference_date_string(args.reference_date)
        except ValueError as e:
            print(f"ОШИБКА: {e}")
            sys.exit(2)
    
    # Выводим информацию о проекте
    print_project_info()
    
    # Настраиваем окружение
    current_dir = setup_environment()
    
    # Загружаем модуль checker
    FastDataQualityChecker = load_checker_module()
    
    # Создаем экземпляр checker
    try:
        checker = FastDataQualityChecker(
            DB_PATH, RULES_FILE, OUTPUT_DIR,
            parallel_tables=getattr(args, 'parallel_tables', 0),
            use_async_load=getattr(args, 'async_load', False),
            debug=getattr(args, 'debug', False),
            reference_datetime=reference_dt,
        )
    except Exception as e:
        print(f"ОШИБКА СОЗДАНИЯ CHECKER: {type(e).__name__}: {e}")
        sys.exit(1)

    if reference_dt:
        print(f"[INFO] Опорная дата для правил «на дату»: {reference_dt} (конец календарного дня)")
    else:
        print("[INFO] Опорная дата: не задана — для правил «на дату» используется текущее время системы")
        print(
            "[WARN] Проверка старых выгрузок: без опорной даты возраст считается от «сегодня», а не от даты снимка. "
            "Итоги по RCCONF_173.1 и аналогам будут искажены. Задайте дату актуальности данных "
            "(--reference-date или ввод при запуске), совпадающую с датой/регламентом выгрузки."
        )
        # При запуске из консоли без --reference-date можно задать дату здесь (как в интерактивном режиме F/1/N/M)
        cli_runs_check = bool(
            args.all or args.table or args.tables or getattr(args, "only_rules", None)
        )
        if cli_runs_check and sys.stdin.isatty():
            print(
                "[INFO] Задать опорную дату сейчас (RCCONF_173.1 и др.)? Enter — оставить время системы; "
                "или введите дату ниже."
            )
            prompt_reference_datetime(checker)

    # Обрабатываем аргументы командной строки
    if args.list:
        list_tables(checker)
    
    elif args.all:
        run_full_check(checker)
    
    elif args.table:
        run_table_check(checker, args.table, only_rule_codes=_parse_only_rule_codes(args.only_rules))
    
    elif args.tables:
        run_selected_tables_check(
            checker, args.tables, only_rule_codes=_parse_only_rule_codes(args.only_rules)
        )
    
    elif getattr(args, 'only_rules', None):
        only_rules = [s.strip() for s in args.only_rules.split(',') if s.strip()]
        if only_rules:
            checker.run(only_rule_codes=set(only_rules))
        else:
            print("Укажите хотя бы одно правило для --only-rules (через запятую).")
    
    else:
        # Если нет аргументов - запускаем интерактивный режим
        interactive_mode(checker)
    
    print(f"\n{'='*80}")
    print("РАБОТА ПРОГРАММЫ ЗАВЕРШЕНА")
    print(f"{'='*80}")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print(f"\n\nПрограмма прервана пользователем.")
        sys.exit(0)
    except Exception as e:
        print(f"\n НЕОБРАБОТАННАЯ ОШИБКА: {type(e).__name__}: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)