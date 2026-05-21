# # utils/file_manager.py
# """Управление файлами и директориями"""

# import os
# from datetime import datetime
# import pandas as pd
# from .symbols import Symbols

# class ErrorFileManager:
#     """Менеджер для сохранения ошибок в файлы"""
    
#     def __init__(self, output_dir):
#         self.output_dir = output_dir
#         self.rule_errors_dir = None
#         self.symbols = Symbols()
    
#     # def save_errors(self, table_name, rule_code, quality_category, error_df):
#     #     """Сохраняет ошибки в CSV файл"""
#     #     if self.rule_errors_dir is None:
#     #         timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
#     #         self.rule_errors_dir = os.path.join(self.output_dir, f"rule_errors_{timestamp}")
#     #         os.makedirs(self.rule_errors_dir, exist_ok=True)
#     #         self.symbols.print_with_symbol('FILE_FOLDER', f'Создана папка для ошибок: {self.rule_errors_dir}')
        
#     #     filename = f"{rule_code}.csv"
#     #     filepath = os.path.join(self.rule_errors_dir, filename)
        
#     #     error_df.to_csv(filepath, index=False, encoding='utf-8-sig')
        
#     #     return filepath

#     def save_errors(self, table_name, rule_code, quality_category, error_df):
#         """Сохраняет ошибки в XLSX файл"""
#         if self.rule_errors_dir is None:
#             timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
#             self.rule_errors_dir = os.path.join(self.output_dir, f"rule_errors_{timestamp}")
#             os.makedirs(self.rule_errors_dir, exist_ok=True)
#             self.symbols.print_with_symbol('FILE_FOLDER', f'Создана папка для ошибок: {self.rule_errors_dir}')
        
#         filename = f"{rule_code}.xlsx"
#         filepath = os.path.join(self.rule_errors_dir, filename)
        
#         error_df.to_excel(filepath, index=False)
        
#         return filepath
    
#     def get_errors_directory(self):
#         """Получить директорию с ошибками"""
#         return self.rule_errors_dir

# utils/file_manager.py
"""Управление файлами и директориями"""

import os
from datetime import datetime
import pandas as pd
from .symbols import Symbols

class ErrorFileManager:
    """Менеджер для сохранения ошибок в файлы"""
    
    def __init__(self, output_dir):
        self.output_dir = output_dir
        self.rule_errors_dir = None
        self.symbols = Symbols()
    
    def save_errors(self, table_name, rule_code, quality_category, error_df):
        """Сохраняет ошибки в XLSX файл с ограничением по количеству строк"""
        if self.rule_errors_dir is None:
            timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            self.rule_errors_dir = os.path.join(self.output_dir, f"rule_errors_{timestamp}")
            os.makedirs(self.rule_errors_dir, exist_ok=True)
            self.symbols.print_with_symbol('FILE_FOLDER', f'Создана папка для ошибок: {self.rule_errors_dir}')
        
        # Проверяем количество строк
        total_rows = len(error_df)
        
        # Максимальное количество строк для Excel файла (чтобы не загружать систему)
        MAX_ROWS_FOR_FULL_EXPORT = 100000  # 100k строк максимум для полного экспорта
        SAMPLE_SIZE = 10000  # для предпросмотра ошибок при большом объеме
        
        if total_rows > MAX_ROWS_FOR_FULL_EXPORT:
            print(f"Ошибок слишком много: {total_rows:,} строк")
            print(f"Сохраняем только первые {SAMPLE_SIZE:,} строк для анализа")
            
            # Сохраняем только выборку в Excel
            filename = f"{rule_code}_SAMPLE_{SAMPLE_SIZE}.xlsx"
            filepath = os.path.join(self.rule_errors_dir, filename)
            
            # Берем выборку ошибок для анализа
            sample_df = error_df.head(SAMPLE_SIZE)
            
            # Добавляем информационные листы
            with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                # Лист с ошибками
                sample_df.to_excel(writer, sheet_name='Пример ошибок', index=False)
                
                # Лист со статистикой
                stats_data = {
                    'Параметр': [
                        'Код правила',
                        'Описание таблицы', 
                        'Категория качества',
                        'Всего записей в проверке',
                        'Найдено ошибок',
                        'Процент ошибок',
                        'Сохранено строк в файле',
                        'Дата сохранения'
                    ],
                    'Значение': [
                        rule_code,
                        table_name,
                        quality_category,
                        'N/A', 
                        f'{total_rows:,}',
                        'N/A', 
                        f'{SAMPLE_SIZE:,} (из {total_rows:,})',
                        datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    ]
                }
                stats_df = pd.DataFrame(stats_data)
                stats_df.to_excel(writer, sheet_name='Статистика', index=False)
            
            print(f"Пример ошибок в Excel: {filename} ({SAMPLE_SIZE:,} строк из {total_rows:,})")
            
            return filepath
            
        elif total_rows > 50000:
            # Для 50k-100k строк сохраняем полностью, но с предупреждением
            print(f"Много ошибок: {total_rows:,} строк")
            print(f"Сохраняем все ошибки в Excel (может занять время)...")
            
            filename = f"{rule_code}.xlsx"
            filepath = os.path.join(self.rule_errors_dir, filename)
            
            # Сохраняем с информационными листами
            with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                error_df.to_excel(writer, sheet_name='Ошибки', index=False)
                
                # Лист со статистикой
                stats_data = {
                    'Параметр': ['Код правила', 'Таблица', 'Категория', 'Всего ошибок', 'Дата сохранения'],
                    'Значение': [rule_code, table_name, quality_category, f'{total_rows:,}', 
                                datetime.now().strftime("%Y-%m-%d %H:%M:%S")]
                }
                stats_df = pd.DataFrame(stats_data)
                stats_df.to_excel(writer, sheet_name='Информация', index=False)
            
            print(f"Сохранено в Excel: {filename}")
            
            return filepath
            
        else:
            # Для небольших объемов сохраняем как обычно
            if total_rows > 10000:
                print(f"Сохраняем {total_rows:,} ошибок в Excel...")
            
            filename = f"{rule_code}.xlsx"
            filepath = os.path.join(self.rule_errors_dir, filename)
            
            # Сохраняем ошибки
            error_df.to_excel(filepath, index=False, engine='openpyxl')
            
            if total_rows > 0:
                print(f"      [OK] Сохранено в Excel: {filename}")
            
            return filepath
    
    def get_errors_directory(self):
        """Получить директорию с ошибками"""
        return self.rule_errors_dir