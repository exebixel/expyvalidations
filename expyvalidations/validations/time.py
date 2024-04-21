import re
from typing import Any
from datetime import time

from expyvalidations.exceptions import ValidateException


def validate_time(value: Any) -> time:
    value = str(value).strip()
    # Define the regex pattern for matching times with HH:MM and HH:MM:SS formats
    time_pattern = r"\b(\d{1,2}):(\d{2})(?::(\d{2}))?(?:\s?[AP]M)?\b"

    # Search for the first occurrence of a time format in the input string
    match = re.search(time_pattern, value, re.IGNORECASE)

    # If a match is found, convert it into a time object
    if match:
        hours, minutes, seconds = (
            int(match.group(1)),
            int(match.group(2)),
            int(match.group(3) or 0),
        )
        return time(hour=hours, minute=minutes, second=seconds)
    else:
        raise ValidateException(f"Value '{value}' is not a valid time format HH:MM:SS")
