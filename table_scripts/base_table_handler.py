# table_scripts/base_table_handler.py
"""Базовый класс для обработки таблиц"""

from abc import ABC, abstractmethod
import pandas as pd
import logging

class BaseTableHandler(ABC):
    """Базовый класс для обработки проверок таблиц"""
    
    def __init__(self, table_name: str, df: pd.DataFrame, memory_manager, error_manager, config: dict = None):
        self.table_name = table_name
        self.df = df.copy() if df is not None else None
        self.memory_manager = memory_manager
        self.error_manager = error_manager
        self.config = config or {}
        self.logger = logging.getLogger(f"TableHandler.{table_name}")
        
        # Результаты проверок
        self.results = []
        self.errors = {}
    
    @abstractmethod
    def get_table_name(self):
        """Возвращает имя таблицы"""
        return self.table_name
    
    @abstractmethod
    def validate_rule(self, rule: dict):
        """Выполняет проверку одного правила"""
        pass
    
    def get_available_tables(self):
        """Возвращает список доступных таблиц в memory_manager"""
        if hasattr(self.memory_manager, 'data_cache'):
            return list(self.memory_manager.data_cache.keys())
        return []
    
    def get_table(self, table_name: str) -> pd.DataFrame:
        """Получает таблицу из memory_manager"""
        return self.memory_manager.get_table(table_name)
    
    def add_result(self, rule_code: str, total_rows: int, error_count: int, 
                   execution_time: float, is_suspicious: bool = False, 
                   matched_column: str = "", error_df: pd.DataFrame = None):
        """Добавляет результат проверки"""
        
        success = total_rows - error_count
        success_rate = (success / total_rows * 100) if total_rows > 0 else 0
        
        if error_count == 0:
            status = "УСПЕШНО"
            status_color = "green"
        elif is_suspicious:
            if error_count > self.error_manager.MAX_ERRORS_TO_SAVE:
                status = "МАССОВЫЕ ОШИБКИ"
            else:
                status = "ПОДОЗРИТЕЛЬНО"
            status_color = "orange"
        else:
            status = "ОШИБКИ"
            status_color = "red"
        
        result = {
            "rule_code": rule_code,
            "total_records": total_rows,
            "passed": success,
            "failed": error_count,
            "success_rate_%": round(success_rate, 2),
            "execution_time_sec": round(execution_time, 2),
            "status": status,
            "status_color": status_color,
            "matched_column": matched_column
        }
        
        self.results.append(result)
        
        # Сохраняем ошибки если есть
        if error_df is not None and not error_df.empty:
            self.errors[rule_code] = {
                'error_df': error_df,
                'error_count': error_count,
                'total_rows': total_rows,
                'is_suspicious': is_suspicious
            }
    
    def get_results(self):
        """Возвращает все результаты"""
        return self.results
    
    def get_errors(self):
        """Возвращает все ошибки"""
        return self.errors