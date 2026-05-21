import pandas as pd
from datetime import datetime
from .base_validator import BaseValidator

class DateValidator(BaseValidator):
    """Проверка дат"""
    
    def validate(self, df, column_name, technical_definition=None, **kwargs):
        if column_name not in df.columns:
            return 0, 0, None  # правило не применимо — 0 оценённых строк
        
        total_rows = len(df)
        
        # Простая проверка: не-NULL значения
        error_mask = df[column_name].isna()
        error_count = error_mask.sum()
        
        # Для более сложных проверок можно использовать technical_definition
        
        error_df = self._prepare_error_dataframe(
            df, error_mask,
            'CONFORMITY',
            f'Invalid date in {column_name}'
        )
        
        # Сохраняем ошибки через error_saver если он есть
        if error_df is not None and self.error_saver:
            self._save_errors_if_needed(error_df)
        
        return total_rows, error_count, error_df