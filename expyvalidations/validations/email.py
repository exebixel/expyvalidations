import re
from expyvalidations.exceptions import ValidateException

def validate_email(email: str) -> str:
    """ Verifica se um email é valido
    """
    if email is None:
        return None
    regex = re.compile(r'^[\w\.-]+@(?:[a-zA-Z-1-9-_]+\.)+[a-zA-Z]{2,}$',
                       flags=re.ASCII)
    email =  re.sub(r' ', '', email).lower()

    if regex.match(email):
        return email
    raise ValidateException(f"Value '{email}' is not a valid email")
