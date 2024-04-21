import re
from typing import Any

from expyvalidations.exceptions import ValidateException

def validate_int(value: Any) -> int:
    """
    Verificações padrão do tipo INT
    """
    _value = re.sub(r'\.0$', '', str(value))
    nums = re.findall(r"[0-9]+", _value)
    if len(nums) == 1:
        try: 
            return int(nums[0])
        except ValueError:
            raise ValidateException(f"Value '{value}' is not a valid number")
    raise ValueError(f"Value '{value}' is not a valid number")
