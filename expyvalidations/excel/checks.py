""" Modulo com as verificações e tratamentos de padrões
"""

import re
from datetime import datetime, time
from typing import Any, Callable

from pandas import DataFrame, notnull, isnull
from alive_progress import alive_bar
from expyvalidations import config
from expyvalidations.config import CheckException, config_bar_excel
from expyvalidations.utils import (
    check_email,
    get_bool,
    get_cel,
    get_cpf,
    get_date,
    get_float,
    get_int,
    get_sex,
    get_time,
    print_error,
    print_warning,
)


class CheckTypes:
    """Nessa Classe contem os netodos de verificações padrão \n

    Todos os metodos de verificação devem possuir como parametros \n
    value: valor a ser verificado, \n
    key: chave do valor que sera verificado (usado para log), \n
    row: linha da planilha em que o dado (value) esta (usado para log), \n
    default: valor padrão a ser usado caso as verificações não passem
    (usado apenas com RESOLVE_ERROS = True)

    E irão retornar o value ao final das verificações e tratamentos,
    caso alguma verificação não passe devera retornar uma Exception
    """

    def get_type_function(self, types: str) -> Callable:
        """Retorna a função referente ao tipo de verificação"""
        return getattr(self, f"check_{types}")

    def check_string(self, value: Any) -> str:
        """
        Verificações de tipo string
        """
        return str(value).strip()

    def check_float(self, value: Any) -> float:
        """
        Verificações de tipo float
        """
        try:
            value = get_float(value)
        except ValueError as exp:
            raise CheckException(f"Value '{value}' is not a valid number")

        return value

    def check_int(self, value: Any) -> int:
        """
        Verificações padrão do tipo INT
        """
        try:
            value = get_int(value)
        except ValueError as exp:
            raise CheckException(f"Value '{value}' is not a valid number")

        return value

    def check_time(self, value: Any) -> str:
        """Verificações padrão do tipo TIME"""
        try:
            index_value = get_time(value)
            value = str(time(*index_value[:3]))
        except TypeError as exp:
            raise CheckException(f"Value '{value}' is not a valid time format HH:MM:SS")

        return value

    def check_date(self, value: Any) -> str:
        """Verificações padrão do tipo DATE"""
        try:
            date = get_date(value)
            value = datetime.strftime(date, "%Y-%m-%dT00:00:00")
        except ValueError as exp:
            raise CheckException(f"Value '{value}' is not a valid date format")

        return value

    def check_bool(self, value: Any) -> bool:
        """
        Verifica se o valor pode ser convertido em booleano
        retorna os mesmo valores com as devidas alterações
        """
        value = str(value).strip()
        value = get_bool(value)
        if value is None:
            raise CheckException(f"Value '{value}' is not a valid boolean")
        return value

    def check_cel(self, value: Any) -> str:
        """
        Verificações de telefone,
        retira caracteres especiais deixando apenas números
        caso não for valido, retorna None no campo
        """
        try:
            value = get_cel(value)
        except ValueError as exp:
            raise CheckException(f"Value '{value}' is not a valid phone number")

        return value

    def check_email(self, value: Any) -> str:
        """Verificações padrão do tipo EMAIL"""
        email = check_email(value)
        if email is None:
            CheckException(f"Value '{value}' is not a valid email")

        return email

    def check_cpf(self, value: Any) -> str:
        """Verificações padrão do tipo CPF"""
        value = re.sub(r"\.0$", "", str(value))
        cpf = get_cpf(value)
        if cpf is None:
            raise CheckException(f"Value '{value}' is invalid CPF")
        return cpf

    def check_sex(self, value: Any) -> str:
        sex = get_sex(value)
        if sex is None:
            CheckException(f"Value '{value}' is not a valid sex")
        return sex

    def check_duplications(self, data: DataFrame, keys: list[str]) -> DataFrame:
        """
        Verifica se há duplicatas no dataframe
        """
        erros = []

        config_bar_excel()
        with alive_bar(len(keys), title="Checking duplications...") as pbar:
            for col in keys:
                # Lista duplicidades todos os registros duplicados (com exceção da primeira ocorrência)
                duplicated = data[data.duplicated(subset=col)][[col]]
                # Exclui items nulos
                duplicated = duplicated.dropna(subset=[col])
                if not duplicated.empty:
                    for index, row in duplicated.iterrows():
                        erros.append(
                            {
                                "error": f"value '{row[col]}' is duplicated",
                                "column": col,
                                "line": index,
                            }
                        )
                pbar()

        if not erros:
            return data
        raise CheckException(erros)
