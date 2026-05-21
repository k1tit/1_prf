# utils/column_matcher.py
import pandas as pd
"""Утилиты для поиска соответствия колонок"""

"""Маппер для поиска колонок в таблицах - ИСПРАВЛЕННАЯ ВЕРСИЯ"""

class ColumnMatcher:
    """Находит соответствие между названиями колонок"""
    
    def __init__(self):
        self.cache = {}
    
    def find_column_match(self, available_columns, target_column):
        """Находит совпадение колонки в списке доступных"""
        # Проверяем кэш
        cache_key = f"{','.join(sorted(available_columns))}:{target_column}"
        if cache_key in self.cache:
            return self.cache[cache_key]
        
        # Преобразуем target_column к нижнему регистру для сравнения
        target_lower = target_column.lower().strip()
        
        # Варианты написания target_column
        target_variants = [
            target_lower,
            target_lower.replace('_', ''),
            target_lower.replace('_', ' '),
            target_lower.replace(' ', '_'),
        ]
        
        # Ищем точное совпадение
        for col in available_columns:
            col_lower = col.lower()
            
            # 1. Точное совпадение (без учета регистра)
            if col_lower == target_lower:
                self.cache[cache_key] = col
                return col
            
            # 2. Совпадение без специальных символов
            col_simple = col_lower.replace('_', '').replace(' ', '')
            target_simple = target_lower.replace('_', '').replace(' ', '')
            
            if col_simple == target_simple:
                self.cache[cache_key] = col
                return col
            
            # 3. Частичное совпадение (если target является частью column)
            if target_simple in col_simple and len(target_simple) > 3:
                self.cache[cache_key] = col
                return col
        
        # Если не нашли, возвращаем None
        self.cache[cache_key] = None
        return None
    
    # ДОБАВЛЯЕМ АЛЬТЕРНАТИВНОЕ НАЗВАНИЕ МЕТОДА (если используется другое имя)
    def find_match(self, available_columns, target_column):
        """Альтернативное название метода"""
        return self.find_column_match(available_columns, target_column)
    
    def match_column(self, available_columns, target_column):
        """Еще одно альтернативное название"""
        return self.find_column_match(available_columns, target_column)