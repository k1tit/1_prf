"""
Обработчик правил для таблицы ADRC
Использует стандартную обработку через checker
"""

import pandas as pd
import logging
from typing import Dict, Any

class ADRCHandler:
    """Обработчик для таблицы ADRC - использует стандартную обработку"""
    
    def __init__(self, table_name: str, df: pd.DataFrame, memory_manager, checker):
        self.table_name = table_name
        self.df = df
        self.memory_manager = memory_manager
        self.checker = checker
        self.logger = logging.getLogger("ADRCHandler")
        
    def validate_rule(self, rule: Dict[str, Any]) -> Dict[str, Any]:
        """
        Валидирует одно правило для таблицы ADRC
        Использует стандартный метод обработки из checker, но БЕЗ сохранения результата
        (сохранение происходит в _process_with_table_handler)
        """
        from datetime import datetime
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # Вызываем _process_single_rule БЕЗ сохранения результата
        # Для этого используем внутренний метод, который не сохраняет результат
        error_count, total_rows = self.checker._process_single_rule_without_save(
            rule, 
            self.table_name, 
            self.df, 
            timestamp
        )
        
        # Формируем результат в формате, который ожидает _process_with_table_handler
        rule_code = rule.get('rule_code', 'UNKNOWN')
        rule_description = rule.get('rule_description', 'Unknown rule')
        quality_category = rule.get('quality_category', 'Unknown')
        column_checked = rule.get('column_name_checked', '')
        
        is_suspicious = self.checker._check_if_suspicious(rule_code, error_count, total_rows)
        
        # Получаем информацию о сохраненных ошибках
        key = f"{rule_code}_{self.table_name}"
        error_file_status = 'Нет'
        error_df = None
        if key in self.checker.rule_errors:
            error_file_status = 'Да'
            error_df = self.checker.rule_errors[key].get('error_df', None)
        
        return {
            'rule_code': rule_code,
            'rule_description': rule_description,
            'quality_category': quality_category,
            'table_name': self.table_name,
            'column_checked': column_checked,
            'total_records': total_rows,
            'passed': total_rows - error_count,
            'failed': error_count,
            'error_count': error_count,  # Дублируем для совместимости с _process_with_table_handler
            'success_rate_%': round((total_rows - error_count) / total_rows * 100, 2) if total_rows > 0 else 0,
            'execution_time_sec': 0,  # Будет установлено в _process_with_table_handler
            'status': 'УСПЕШНО' if error_count == 0 else ('ПОДОЗРИТЕЛЬНО' if is_suspicious else 'ОШИБКИ'),
            'status_color': 'green' if error_count == 0 else ('orange' if is_suspicious else 'red'),
            'error_file': error_file_status,
            'comments': '',
            'error_df': error_df  # Передаем error_df для сохранения в _process_with_table_handler
        }

