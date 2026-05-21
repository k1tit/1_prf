# utils/symbols.py
"""Символы и иконки для отчетов"""

# Unicode коды для Excel отчета
EXCEL_SYMBOLS = {
    'SUCCESS': chr(0x2705), 
    'ERROR': chr(0x274C),    
    'WARNING': chr(0x26A0),  
    'INFO': chr(0x2139),  
    'SKIP': chr(0x23ED),   
    'ROCKET': chr(0x1F680), 
    'GEAR': chr(0x2699),     
    'MAGNIFYING_GLASS': chr(0x1F50D), 
    'CHECKMARK': chr(0x2713), 
    'CHART_UP': chr(0x1F4C8),   
    'CHART_DOWN': chr(0x1F4C9), 
    'CLIPBOARD': chr(0x1F4CB),   
    'FILE_FOLDER': chr(0x1F4C1), 
    'PAGE': chr(0x1F4C4),      
    'TABLE': chr(0x1F4BC),      
    'COLUMN': chr(0x1F4D1),     
    'BOOKS': chr(0x1F4DA),     
    'SAVE': chr(0x1F4BE),   
    'TARGET': chr(0x1F3AF), 
    'PALETTE': chr(0x1F3A8),
    'CELEBRATION': chr(0x1F389), 
    'BAR_CHART': chr(0x1F4CA),   
    'MEMO': chr(0x1F4DD),       
}

# Текстовые обозначения для терминала
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
    """Класс для работы с символами"""
    
    @staticmethod
    def get_excel(symbol_name):
        """Получить символ для Excel"""
        return EXCEL_SYMBOLS.get(symbol_name, '')
    
    @staticmethod
    def get_terminal(symbol_name):
        """Получить символ для терминала"""
        return TERMINAL_SYMBOLS.get(symbol_name, '')
    
    @staticmethod
    def print_with_symbol(symbol_name, message, end='\n'):
        """Вывести сообщение с символом"""
        print(f"{TERMINAL_SYMBOLS.get(symbol_name, '')} {message}", end=end)