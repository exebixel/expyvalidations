from datetime import datetime

from expyvalidations.exceptions import ValidateException


def validate_date(date_string) -> datetime.date:
    date_string = str(date_string).strip()

    try:
        # Try parsing as ISO format (YYYY-MM-DD)
        date = datetime.strptime(date_string, "%Y-%m-%d")
        return date.date()
    except ValueError:
        pass

    try:
        # Try parsing as Brazilian format (DD/MM/YYYY)
        date = datetime.strptime(date_string, "%d/%m/%Y")
        return date.date()
    except ValueError:
        pass

    # If none of the formats match, raise an exception
    raise ValidateException(
        f"Value '{date_string}' is not a valid date format (YYYY-MM-DD or DD/MM/YYYY)"
    )
