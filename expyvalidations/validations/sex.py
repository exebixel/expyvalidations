from typing import Any
from expyvalidations.exceptions import ValidateException


def validate_sex(value: Any) -> str:
    value = str(value).upper().strip()
    if value == "M" or value == "F":
        return value
    else:
        raise ValidateException("Value must be M or F")
