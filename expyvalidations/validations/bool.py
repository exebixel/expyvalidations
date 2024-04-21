from typing import Any

from expyvalidations.exceptions import ValidateException


def validate_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value

    if str(value).upper().strip() in ['Y', 'S', '1', '1.0', 'TRUE', 'T', 'YES']:
        return True
    elif str(value).upper().strip() in ['N', '0', '0.0', 'FALSE', 'F', 'NO']:
        return False

    raise ValidateException(f"Value '{value}' is not a valid boolean format (Y/N, S/N, 1/0, TRUE/FALSE, T/F, YES/NO)")