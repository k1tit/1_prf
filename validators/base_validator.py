# validators/base_validator.py
"""Базовый класс для валидаторов"""

from abc import ABC, abstractmethod
import pandas as pd
from datetime import datetime

# ВМЕСТО: from ..dq_util.symbols import Symbols
# Используем абсолютный импорт:

try:
    # Пробуем импорт из utils
    from utils.symbols import Symbols
except ImportError:
    # Если не работает, создаем локальные символы
    TERMINAL_SYMBOLS = {
        'SUCCESS': '[OK]',
        'ERROR': '[ERROR]',
        'WARNING': '[WARN]',
        'INFO': '[INFO]',
        'SKIP': '[SKIP]',
        'ROCKET': '[START]',
        'GEAR': '[PROC]',
        'MAGNIFYING_GLASS': '[CHECK]',
        'CHECKMARK': '[DONE]',
        'CHART_UP': '[STAT+]',
        'CHART_DOWN': '[STAT-]',
        'CLIPBOARD': '[CLIP]',
        'FILE_FOLDER': '[DIR]',
        'PAGE': '[FILE]',
        'TABLE': '[TABLE]',
        'COLUMN': '[COL]',
        'BOOKS': '[DATA]',
        'SAVE': '[SAVE]',
        'TARGET': '[TARGET]',
        'PALETTE': '[STYLE]',
        'CELEBRATION': '[DONE]',
        'BAR_CHART': '[CHART]',
        'MEMO': '[NOTE]',
    }
    
    class Symbols:
        @staticmethod
        def get_terminal(symbol_name):
            return TERMINAL_SYMBOLS.get(symbol_name, '')

class BaseValidator(ABC):
    """Абстрактный базовый класс валидатора"""
    
    def __init__(self, rule_info, error_saver=None):
        self.rule_info = rule_info
        self.error_saver = error_saver
        self.symbols = Symbols()
    
    @abstractmethod
    def validate(self, df, column_name, **kwargs):
        """Основной метод валидации"""
        pass
    
    def _prepare_error_dataframe(self, df, error_mask, error_type, error_description):
        """Подготовить DataFrame с ошибками"""
        if not error_mask.any():
            return None
        
        error_df = df[error_mask].copy()
        
        # Добавляем метаданные об ошибке
        error_df['DQ_ERROR_TYPE'] = error_type
        error_df['DQ_RULE_CODE'] = self.rule_info['rule_code']
        error_df['DQ_RULE_DESCRIPTION'] = self.rule_info['rule_description']
        error_df['DQ_COLUMN_CHECKED'] = self.rule_info.get('matched_column', '')
        error_df['DQ_ERROR_DESCRIPTION'] = error_description
        error_df['DQ_TIMESTAMP'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        return error_df
    
    def _save_errors_if_needed(self, error_df):
        """Сохранить ошибки если есть и настроен saver"""
        if error_df is not None and self.error_saver:
            return self.error_saver.save_errors(
                self.rule_info['table_name'],
                self.rule_info['rule_code'],
                self.rule_info.get('quality_category', 'UNKNOWN'),
                error_df
            )
        return None