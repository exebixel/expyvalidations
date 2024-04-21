import re
from typing import Any

from expyvalidations.exceptions import ValidateException


def validate_float(value: Any) -> float:
    """
    Verificações de tipo float
    """
    strnum = str(value).strip()
    nums = re.findall(r"[0-9?.?,]+", strnum)
    if len(nums) == 1:
        if "," in nums[0]:
            nums[0] = re.sub(r",", ".", nums[0])
        try:
            return float(nums[0])
        except ValueError:
            raise ValidateException("Value '{value}' is not a valid number")

    if len(nums) == 0:
        raise ValidateException("Number not found")
    raise ValidateException("Theres more than one number in the value")
