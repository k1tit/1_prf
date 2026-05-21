# validators/__init__.py
from .completeness import CompletenessValidator
from .conformity import ConformityValidator
from .cross_column import CrossColumnEqualityValidator
from .text_validators import (
    SpecialCharactersValidator, 
    ConsecutiveSpacesValidator,
    UppercaseValidator
)

# Добавляем новый валидатор
try:
    from .length_validator import LengthValidator
except ImportError:
    LengthValidator = None

__all__ = [
    'CompletenessValidator',
    'ConformityValidator', 
    'CrossColumnEqualityValidator',
    'SpecialCharactersValidator',
    'ConsecutiveSpacesValidator',
    'UppercaseValidator',
    'LengthValidator'
]