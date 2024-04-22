from enum import Enum
from typing import Any, Callable
from pydantic import BaseModel, Field


class Types(Enum):
    STRING = "string"
    INTEGER = "int"
    FLOAT = "float"
    TIME = "time"
    BOOL = "bool"
    DATE = "date"
    EMAIL = "email"
    PHONE = "cel"
    CPF = "cpf"
    SEX = "sex"


class TypeError(Enum):
    WARNING = "Warning"
    ERROR = "ERROR"
    DUPLICATED = "Duplication"
    CRITICAL = "CRITICAL"


class Error(BaseModel):
    row: int
    column: str | None
    type: TypeError = TypeError.ERROR
    message: str


class ColumnDefinition(BaseModel):
    key: str
    default: Any = None
    function_validation: Callable
    custom_function_before: Callable | None = Field(default=None)
    custom_function_after: Callable | None = Field(default=None)
