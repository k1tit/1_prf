# validators/advanced_special_characters.py
"""Продвинутый валидатор специальных символов для правил RCCONF_18.2 и RCCONF_22.2"""

import re
import pandas as pd
import os
from .base_validator import BaseValidator

class AdvancedSpecialCharactersValidator(BaseValidator):
    """Проверка специальных символов с учетом:
    1. Специальных символов из конфигурационного файла
    2. Последовательных повторяющихся специальных символов
    3. Четности кавычек (должны быть парами)
    4. Правильности скобок (должны быть парами и в правильном порядке)
    """
    
    def __init__(self, rule_info):
        super().__init__(rule_info)
        self.rule_code = rule_info.get('rule_code', '')
        self.rule_config = self._load_rule_config()
        self.special_chars_config = self._load_special_chars_config()
    
    def _load_rule_config(self):
        """Загружает конфигурацию правила из JSON файла"""
        import json
        config_paths = [
            os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'json files', 'conf_special_characters.json'),
            os.path.join('json files', 'conf_special_characters.json'),
            os.path.join('data', 'conf_special_characters.json')
        ]
        
        for config_path in config_paths:
            try:
                if os.path.exists(config_path):
                    with open(config_path, 'r', encoding='utf-8') as f:
                        config = json.load(f)
                    
                    if self.rule_code in config:
                        rule_config = config[self.rule_code]
                        print(f"      [INFO] Загружена конфигурация для {self.rule_code} из {config_path}")
                        return rule_config
            except Exception as e:
                continue
        
        print(f"      [WARN] Конфигурация для {self.rule_code} не найдена, используются дефолтные значения")
        return None
    
    def _load_special_chars_config(self):
        """Загружает конфигурацию специальных символов из Excel файла"""
        # Пробуем разные пути
        possible_paths = [
            os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data', 'conf_special_characters.xlsx'),
            os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'conf_special_characters.xlsx'),
            'data/conf_special_characters.xlsx',
            'conf_special_characters.xlsx'
        ]
        
        special_chars = set()
        
        for config_path in possible_paths:
            try:
                if os.path.exists(config_path):
                    df = pd.read_excel(config_path)
                    print(f"      [INFO] Загружен conf_special_characters.xlsx из {config_path}")
                    
                    # Ищем колонку с символами и колонку с флагом '1'
                    char_col = None
                    flag_col = None
                    
                    for col in df.columns:
                        col_lower = str(col).lower()
                        if 'char' in col_lower or 'символ' in col_lower or 'symbol' in col_lower:
                            char_col = col
                        if 'flag' in col_lower or 'mark' in col_lower or 'значение' in col_lower:
                            flag_col = col
                    
                    # Если нашли колонки, читаем символы с флагом '1'
                    if char_col and flag_col:
                        for idx, row in df.iterrows():
                            flag = str(row.get(flag_col, '')).strip()
                            if flag == '1':
                                char = str(row.get(char_col, '')).strip()
                                if char and len(char) == 1:
                                    special_chars.add(char)
                    elif len(df.columns) > 0:
                        # Альтернативный способ: ищем все символы в первой колонке
                        for idx, row in df.iterrows():
                            char = str(row.iloc[0]).strip()
                            if len(char) == 1:  # Один символ
                                special_chars.add(char)
                    
                    if special_chars:
                        break  # Если загрузили, выходим
            except Exception as e:
                continue  # Пробуем следующий путь
        
        if not special_chars:
            print(f"      [WARN] Не удалось загрузить conf_special_characters.xlsx, используются дефолтные правила")
        
        return special_chars
    
    def _check_consecutive_chars(self, text, chars_to_check):
        """Проверяет, есть ли последовательные повторяющиеся символы из списка"""
        if not text or len(text) < 2:
            return False
        
        # Создаем паттерн для поиска последовательных повторений
        for char in chars_to_check:
            # Экранируем специальные символы для regex
            escaped_char = re.escape(char)
            # Ищем 2 или более подряд идущих символа
            pattern = rf'{escaped_char}{{2,}}'
            if re.search(pattern, text):
                return True
        return False
    
    def _check_quotes_pairs(self, text, quote_char):
        """Проверяет, что кавычки используются парами (четное количество)"""
        count = text.count(quote_char)
        return count % 2 != 0  # Если нечетное - ошибка
    
    def _check_brackets_pairs(self, text, bracket_pairs):
        """Проверяет, что скобки используются парами и в правильном порядке"""
        # bracket_pairs - список кортежей: [('[', ']'), ('{', '}'), ('(', ')')]
        # Используем стек для проверки правильности вложенности и порядка
        
        # Создаем словарь для быстрого поиска пар скобок
        bracket_map = {}
        for open_bracket, close_bracket in bracket_pairs:
            bracket_map[open_bracket] = close_bracket
            bracket_map[close_bracket] = open_bracket
        
        # Используем стек для проверки правильности скобок
        stack = []
        
        for i, char in enumerate(text):
            if char in bracket_map:
                if char in [pair[0] for pair in bracket_pairs]:  # Открывающая скобка
                    stack.append((char, i))
                else:  # Закрывающая скобка
                    if not stack:
                        return True  # Ошибка: закрывающая скобка без открывающей
                    
                    last_open, last_pos = stack[-1]
                    expected_close = bracket_map[last_open]
                    
                    if char != expected_close:
                        return True  # Ошибка: неправильная закрывающая скобка
                    
                    # Проверяем, что открывающая скобка идет перед закрывающей
                    if last_pos >= i:
                        return True  # Ошибка: неправильный порядок
                    
                    stack.pop()
        
        # Если остались незакрытые скобки - ошибка
        if stack:
            return True  # Ошибка: остались незакрытые скобки
        
        return False  # Все правильно
    
    def validate(self, df, column_name, technical_definition=None, rule_code=None, **kwargs):
        """Валидация с учетом всех правил - ОПТИМИЗИРОВАННАЯ ВЕРСИЯ"""
        if column_name not in df.columns:
            return 0, 0, None
        
        total_rows = len(df)
        
        # Загружаем параметры из JSON конфига, если есть
        if self.rule_config:
            consecutive_chars = self.rule_config.get('consecutive_chars', '.,/-_&[]{}()"\'`:;`|~+')
            quote_char = self.rule_config.get('quote_char', '"')
            bracket_pairs_raw = self.rule_config.get('bracket_pairs', [['[', ']'], ['{', '}'], ['(', ')']])
            bracket_pairs = [tuple(pair) for pair in bracket_pairs_raw]  # Преобразуем в кортежи
            check_space = self.rule_config.get('check_space', False)
            forbidden_chars_from_json = set(self.rule_config.get('forbidden_chars', []))
            use_config_file = self.rule_config.get('forbidden_chars_from_config', True)
        else:
            # Дефолтные значения, если конфиг не найден
            if rule_code == 'RCCONF_18.2':
                consecutive_chars = '.,/-_&[]{}()"\'`:;`|~+'
                quote_char = '"'
                bracket_pairs = [('[', ']'), ('{', '}'), ('(', ')')]
                check_space = True
            elif rule_code == 'RCCONF_22.2':
                consecutive_chars = '-&()./:[]_`+,\\'
                quote_char = "'"
                bracket_pairs = [('[', ']'), ('(', ')')]
                check_space = False
            else:
                consecutive_chars = '.,/-_&[]{}()"\'`:;`|~+'
                quote_char = '"'
                bracket_pairs = [('[', ']'), ('{', '}'), ('(', ')')]
                check_space = False
            forbidden_chars_from_json = set()
            use_config_file = True
        
        # Загружаем специальные символы из Excel конфига, если нужно
        forbidden_chars_from_excel = set()
        if use_config_file:
            forbidden_chars_from_excel = self.special_chars_config if self.special_chars_config else set()
        
        # Объединяем запрещенные символы из JSON и Excel.
        # Важно: символы из блока consecutive_chars / quotes / brackets
        # не должны запрещаться "сами по себе", если правило требует
        # проверять только их повтор подряд или парность.
        forbidden_chars = forbidden_chars_from_json | forbidden_chars_from_excel
        structural_chars = set(consecutive_chars or "")
        if quote_char:
            structural_chars.add(quote_char)
        for open_bracket, close_bracket in bracket_pairs:
            structural_chars.add(open_bracket)
            structural_chars.add(close_bracket)

        # Явно запрещённые символы из JSON всегда сохраняем.
        forbidden_chars = (forbidden_chars - structural_chars) | forbidden_chars_from_json

        rule_code_u = str(rule_code or "").strip().upper()
        # Для 18.2 и 22.2:
        #   - '.' и ';' как одиночные символы допустимы (не forbidden),
        #   - но повтор подряд ('..' / ';;') должен ловиться через consecutive-проверку.
        if rule_code_u in {"RCCONF_18.2", "RCCONF_22.2"}:
            forbidden_chars.discard(".")
            forbidden_chars.discard(";")
            base_consecutive = consecutive_chars or ""
            # Гарантируем, что '.' и ';' участвуют в проверке повторов подряд.
            if "." not in base_consecutive:
                base_consecutive += "."
            if ";" not in base_consecutive:
                base_consecutive += ";"
            # Удаляем дубликаты с сохранением порядка.
            consecutive_chars = "".join(dict.fromkeys(base_consecutive))

        # Дополнительно для RCCONF_18.2: ',' должен считаться ошибкой всегда.
        # Слэш '/' в улицах (напр. «вл.1/5») допустим: не forbidden и не ловим '//'.
        if rule_code_u == "RCCONF_18.2":
            forbidden_chars.add(",")
            forbidden_chars.discard("/")
            consecutive_chars = "".join(
                dict.fromkeys(c for c in consecutive_chars if c != "/")
            )
        
        # Для RCCONF_18.2 также проверяем пробел, если указано в конфиге
        if check_space and ' ' not in forbidden_chars:
            # Если check_space=True, но пробел не в списке, добавляем его
            forbidden_chars.add(' ')
        
        # Фильтруем только непустые значения для проверки
        mask_not_empty = df[column_name].notna() & (df[column_name].astype(str).str.strip() != '')
        
        if not mask_not_empty.any():
            return 0, 0, None  # 0 оценённых (все скипнуты — пусто по правилу)
        
        # Работаем только с непустыми значениями
        df_to_check = df[mask_not_empty].copy()
        texts = df_to_check[column_name].astype(str).str.strip()
        
        # Создаем маску ошибок
        error_mask = pd.Series(False, index=df_to_check.index)
        
        # 1. Проверка на действительно запрещённые символы из конфига.
        # Для RCCONF_18.2 это НЕ все знаки препинания, а только символы,
        # явно отмеченные как forbidden (например ';'). Обычные адресные
        # разделители вроде '.', ',', '-' должны проходить и отдельно
        # проверяются только на повтор подряд.
        if forbidden_chars:
            for char in forbidden_chars:
                char_mask = texts.str.contains(re.escape(char), regex=True, na=False)
                error_mask = error_mask | char_mask
        
        # 2. Проверка на последовательные повторяющиеся специальные символы.
        # Одиночные '.', ',', '-', '/', '"' и т.п. допустимы; ошибка только
        # если символ повторяется подряд два и более раз.
        if consecutive_chars:
            # Создаем паттерн для поиска последовательных символов
            consecutive_pattern = '|'.join([re.escape(char) + '{2,}' for char in consecutive_chars])
            if consecutive_pattern:
                consecutive_mask = texts.str.contains(consecutive_pattern, regex=True, na=False)
                error_mask = error_mask | consecutive_mask
        
        # 3. Проверка на четность кавычек (векторизованная)
        if quote_char:
            quote_counts = texts.str.count(re.escape(quote_char))
            quote_mask = (quote_counts % 2) != 0
            error_mask = error_mask | quote_mask
        
        # 4. Проверка на правильность скобок (для каждой строки, но только для строк с ошибками)
        # Оптимизация: проверяем скобки только для строк, которые еще не помечены как ошибки
        # или для строк, где есть скобки
        if bracket_pairs:
            # Создаем паттерн для поиска любых скобок
            all_brackets = ''.join([pair[0] + pair[1] for pair in bracket_pairs])
            bracket_pattern = '[' + re.escape(all_brackets) + ']'
            has_brackets = texts.str.contains(bracket_pattern, regex=True, na=False)
            
            # Проверяем скобки только для строк, где они есть
            for idx in df_to_check[has_brackets].index:
                if not error_mask.loc[idx]:  # Если еще не помечено как ошибка
                    text = texts.loc[idx]
                    if self._check_brackets_pairs(text, bracket_pairs):
                        error_mask.loc[idx] = True
        
        # Получаем индексы строк с ошибками
        error_indices = df_to_check[error_mask].index
        
        evaluated_count = int(mask_not_empty.sum())
        if len(error_indices) == 0:
            return evaluated_count, 0, None
        
        # Формируем DataFrame с ошибками
        error_df = self._prepare_error_dataframe(
            df, 
            pd.Series([idx in error_indices for idx in df.index], index=df.index),
            'CONFORMITY',
            f'Invalid special characters format in {column_name}'
        )
        
        error_count = len(error_indices)
        return evaluated_count, error_count, error_df

