import re
from typing import Any, Callable, Union

import pandas as pd
from alive_progress import alive_bar

from expyvalidations import config
from expyvalidations.excel.models import ColumnDefinition, Error, TypeError, Types
from expyvalidations.exceptions import ValidateException
from expyvalidations.utils import string_normalize
from expyvalidations.validations.bool import validate_bool
from expyvalidations.validations.cpf import validate_cpf
from expyvalidations.validations.date import validate_date
from expyvalidations.validations.duplications import validate_duplications
from expyvalidations.validations.email import validate_email
from expyvalidations.validations.float import validate_float
from expyvalidations.validations.int import validate_int
from expyvalidations.validations.sex import validate_sex
from expyvalidations.validations.string import validate_string
from expyvalidations.validations.time import validate_time


class ExpyValidations:

    def __init__(
        self,
        path_file: str,
        sheet_name: str = "sheet",
        header_row: int = 1,
    ):

        self.column_details: list[ColumnDefinition] = []
        self.__book = pd.ExcelFile(path_file)

        self.__errors_list: list[Error] = []
        """
        indica se ouve algum erro nas validações da planilha
        se tiver erro o método data_all não deve
        retornar dados
        """

        self.excel: pd.DataFrame
        try:
            excel = pd.read_excel(
                self.__book, self.__sheet_name(sheet_name), header=header_row
            )
            # retirando linhas e colunas em brando do Data Frame
            excel = excel.dropna(how="all")
            excel.columns = excel.columns.astype("string")
            excel = excel.loc[:, ~excel.columns.str.contains("^Unnamed")]
            excel = excel.astype(object)
            excel = excel.where(pd.notnull(excel), None)
            self.excel = excel

        except ValueError as exp:
            raise ValueError(exp)

        self.__header_row = header_row

        self.__validations_types = {
            "string": validate_string,
            "float": validate_float,
            "int": validate_int,
            "date": validate_date,
            "time": validate_time,
            "bool": validate_bool,
            "cpf": validate_cpf,
            "email": validate_email,
            "sex": validate_sex,
        }

    def __sheet_name(self, search: str) -> str:
        """
        Função responsável por pesquisa a string do parâmetro 'search'
        nas planilhas (sheets) do 'book' especificado no __init__
        e retornar o 1º nome de planilha que encontrar na pesquisa

        Caso tenha apenas 1 planilha no arquivo ela é retornada
        """
        if len(self.__book.sheet_names) == 1:
            return self.__book.sheet_names[0]

        for names in self.__book.sheet_names:
            name = string_normalize(names)
            if re.search(search, name, re.IGNORECASE):
                return names
        raise ValueError(f"ERROR! Sheet '{search}' not found! Rename your sheet!")

    def __column_name(self, column_name: Union[str, list[str]]) -> str:
        """
        Resquias e retorna o nome da coluna da planilha,
        se não encontrar, retorna ValueError
        """
        excel = self.excel
        if isinstance(column_name, str):
            column_name = [column_name]

        for header in excel.keys():
            header_name = string_normalize(header)
            count = 0
            for name in column_name:
                if re.search(name, header_name, re.IGNORECASE):
                    count += 1
            if count == len(column_name):
                return header

        column_formated = " ".join(column_name)
        raise ValueError(f"Column '{column_formated}' not found!")

    def add_column(
        self,
        key: str,
        name: Union[str, list[str]],
        required: bool = True,
        default: Any = None,
        types: Types = "string",
        custom_function_before: Callable = None,
        custom_function_after: Callable = None,
    ):
        """
        Função responsável por adicionar as colunas que serão lidas
        da planilha \n
        Parâmetros: \n
        key: nome da chave do dicionario com os dados da coluna \n
        name: nome da coluna da planilha, não é necessário informar o
        nome completo da coluna, apenas uma palavra para busca, se o nome da
        coluna não foi encontrado o programa fechará \n
        default: se a coluna não for encontrada ou o valor não foi informado
        então será considerado o valor default \n
        types: tipo de dado que deve ser retirado da coluna \n
        required: define se a coluna é obrigatória na planilha \n
        length: Número máximo de caracteres que o dado pode ter,
        padrão 1 ou seja ilimitado \n

        custom_function: recebe a referencia de uma função que sera executada
            apos as verificações padrão, essa função deve conter os parametros:
            value: (valor que sera verificado),
            key: (Chave do valor que sera verificado, para fins de log),
            row: (Linha da planilha que esta o valor que sera verificado,
                para fins de log),
            default: (Valor padrão que deve ser usado caso caso ocorra algum
                erro na verificação, para resolução de problemas).
            Essa custom_funcition deverá retornar o valor (value) verificado
            em caso de sucesso na verficação/tratamento, caso contratio,
            deve retornar uma Exception
        """
        excel = self.excel

        try:
            default_validation = self.__validations_types[types]
        except KeyError:
            raise ValueError(f"Type '{types}' not found!")

        try:
            column_name = self.__column_name(name)
        except ValueError as exp:
            if required:
                print(f"ERROR! Required {exp}")
                self.__errors_list.append(
                    Error(
                        type=TypeError.CRITICAL,
                        row=self.__header_row,
                        column=None,
                        message=f"Required {exp}",
                    )
                )
            else:
                excel[key] = pd.Series(default, dtype="object")
        else:
            excel.rename({column_name: key}, axis="columns", inplace=True)

            self.column_details.append(
                ColumnDefinition(
                    key=key,
                    default=default,
                    function_validation=default_validation,
                    custom_function_before=custom_function_before,
                    custom_function_after=custom_function_after,
                )
            )

    def check_all(
        self,
        check_row: Callable = None,
        check_duplicated_keys: list[str] = None,
        checks_final: list[Callable] = None,
    ) -> bool:
        """
        Função responsável por verificar todas as colunas
        da planilha

        Parâmetros:
        check_row: função que será executada para cada linha da planilha
        checks_final: lista de funções que serão executadas
            considerando todos os dados da planilha

        Retorno:
        False se NÃO ouve erros na verificação
        True se ouve erros na verificação
        """
        excel = self.excel

        # configuração padrão da barra de progresso
        config.config_bar_excel()

        # verificando todas as colunas
        with alive_bar(
            len(excel.index) * len(self.column_details), title="Checking for columns..."
        ) as pbar:
            # Verificações por coluna
            for column in self.column_details:
                for index in excel.index:
                    value = excel.at[index, column.key]
                    self.__check_value(
                        value=value, index=index, column_definition=column
                    )
                    pbar()

        # Verificações por linha
        if check_row is not None:
            with alive_bar(len(excel.index), title="Checking for rows...") as pbar:
                list_colums = list(map(lambda col: col.key, self.column_details))
                for row in excel.index:
                    try:
                        data = excel[list_colums].loc[row].to_dict()
                        data = check_row(data)
                        for key, value in data.items():
                            excel.at[row, key] = value

                    except ValidateException as exp:
                        self.__errors_list.append(
                            Error(
                                row=self.__row(row),
                                column=None,
                                message=str(exp),
                            )
                        )
                    pbar()

        # Verificações totais (duplicação de dados)
        if check_duplicated_keys is not None:
            try:
                excel = validate_duplications(data=excel, keys=check_duplicated_keys)
            except ValidateException as exp:
                for error in exp.args[0]:
                    error.row = self.__row(error.row)
                    self.__errors_list.append(error)

        # if checks_final is not None:
        #     for check in checks_final:
        #         try:
        #             excel = check(excel)
        #         except CheckException:
        #             self.erros = True
        #         pbar()

        self.excel = excel
        return True if self.__errors_list else False

    def __check_value(
        self,
        value: Any,
        index: int,
        column_definition: ColumnDefinition,
    ) -> None:
        """Executa todas as verificações em um valor especifico,
        retorna True um False para caso as verificações passarem ou não
        """
        key = column_definition.key
        function_validation = column_definition.function_validation
        custom_function_before = column_definition.custom_function_before
        custom_function_after = column_definition.custom_function_after

        functions = []
        if custom_function_before is not None:
            functions.append(custom_function_before)
        functions.append(function_validation)
        if custom_function_after is not None:
            functions.append(custom_function_after)

        for func in functions:
            try:
                value = func(value)
            except ValidateException as exp:
                self.__errors_list.append(
                    Error(
                        row=self.__row(index),
                        column=key,
                        message=str(exp),
                    )
                )
                break

        self.excel.at[index, key] = value

    def __row(self, index: int) -> int:
        """
        Retorna a linha do respectivo index passado
        """
        return index + self.__header_row + 2

    def get_result(self, force: bool = False) -> dict:
        excel = self.excel
        if not force and self.has_errors():
            raise ValidateException("Errors found in the validations")

        if excel.empty:
            return {}

        list_colums = list(map(lambda col: col.key, self.column_details))

        excel = excel.where(pd.notnull(excel), None)
        return excel[list_colums].to_dict("records")

    def has_errors(self) -> bool:
        return True if self.__errors_list else False

    def print_errors(self):
        for error in self.__errors_list:
            if error.column is None:
                print(f"{error.type.value}! in line {error.row}: {error.message}")
            else:
                print(
                    f"{error.type.value}! in line {error.row}, Column {error.column}: {error.message}"
                )

    def get_errors(self) -> list[dict]:
        errors = []
        for error in self.__errors_list:
            errors.append(
                {
                    "type": error.type.value,
                    "row": error.row,
                    "column": error.column,
                    "message": error.message,
                }
            )
        return errors
