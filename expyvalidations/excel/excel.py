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
    """
    A class that provides functionality for validating Excel files.

    Parameters:
    - path_file (str): The path to the Excel file.
    - sheet_name (str): The name of the sheet in the Excel file. Default is "sheet".
    - header_row (int): The row number where the header is located. Default is 1.

    Attributes:
    - __column_details (list[ColumnDefinition]): A list of column definitions.
    - __book (pd.ExcelFile): The Excel file object.
    - __errors_list (list[Error]): A list of errors encountered during validation.
    - excel (pd.DataFrame): The Excel data as a DataFrame.

    Methods:
    - __init__(self, path_file: str, sheet_name: str = "sheet", header_row: int = 1): Initializes the ExpyValidations object.
    - __sheet_name(self, search: str) -> str: Searches for a sheet name in the Excel file.
    - __column_name(self, column_name: Union[str, list[str]]) -> str: Returns the column name in the Excel file.
    - add_column(self, key: str, name: Union[str, list[str]], required: bool = True, default: Any = None, types: Types = "string", custom_function_before: Callable = None, custom_function_after: Callable = None): Adds a column to be validated.
    - check_all(self, check_row: Callable = None, check_duplicated_keys: list[str] = None, checks_final: list[Callable] = None) -> bool: Performs all the validations on the Excel data.
    - __check_value(self, value: Any, index: int, column_definition: ColumnDefinition) -> None: Checks a specific value against the column definition.

    """

    def __init__(
        self,
        path_file: str,
        sheet_name: str = "sheet",
        header_row: int = 1,
    ):
        """
        Initializes the ExpyValidations object.

        Parameters:
        - path_file (str): The path to the Excel file.
        - sheet_name (str): The name of the sheet in the Excel file. Default is "sheet".
        - header_row (int): The row number where the header is located. Default is 1.
        """
        self.__column_details: list[ColumnDefinition] = []
        self.__book = pd.ExcelFile(path_file)

        self.__errors_list: list[Error] = []

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
        Returns the name of the sheet that matches the given search pattern, if file has only one sheet, return the name of the sheet.

        Args:
            search (str): The search pattern to match against the sheet names.

        Returns:
            str: The name of the matching sheet.

        Raises:
            ValueError: If no sheet matching the search pattern is found.
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
        Returns the name of the column in the Excel file that matches the given column name(s).

        Args:
            column_name (Union[str, list[str]]): The name(s) of the column(s) to search for.

        Returns:
            str: The name of the column that matches the given column name(s).

        Raises:
            ValueError: If the column name(s) are not found in the Excel file.
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
        Add a column to validate in the Excel file.

        Args:
            key (str): The key or name of the column.
            name (Union[str, list[str]]): The name or names of the column in the Excel file.
            required (bool, optional): Whether the column is required. Defaults to True.
            default (Any, optional): The default value for the column. Defaults to None.
            types (Types, optional): The type or types of data allowed in the column. Defaults to "string".
            custom_function_before (Callable, optional): A custom function to be executed before validating the column. Defaults to None.
            custom_function_after (Callable, optional): A custom function to be executed after validating the column. Defaults to None.
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

            self.__column_details.append(
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
    ) -> bool:
        """
        Perform various checks on the Excel data.

        Args:
            check_row (Callable, optional): A function that performs additional checks on each row of the Excel data. Defaults to None.
            check_duplicated_keys (list[str], optional): A list of column keys to check for duplicated values. Defaults to None.

        Returns:
            bool: True if there are errors, False otherwise.
        """
        excel = self.excel

        # configuração padrão da barra de progresso
        config.config_bar_excel()

        # verificando todas as colunas
        with alive_bar(
            len(excel.index) * len(self.__column_details),
            title="Checking for columns...",
        ) as pbar:
            # Verificações por coluna
            for column in self.__column_details:
                for index in excel.index:
                    value = excel.at[index, column.key]
                    self.__check_value(
                        value=value, index=index, column_definition=column
                    )
                    pbar()

        # Verificações por linha
        if check_row is not None:
            with alive_bar(len(excel.index), title="Checking for rows...") as pbar:
                list_colums = list(map(lambda col: col.key, self.__column_details))
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

        if check_duplicated_keys is not None:
            try:
                excel = validate_duplications(data=excel, keys=check_duplicated_keys)
            except ValidateException as exp:
                for error in exp.args[0]:
                    error.row = self.__row(error.row)
                    self.__errors_list.append(error)

        self.excel = excel
        return True if self.__errors_list else False

    def __check_value(
        self,
        value: Any,
        index: int,
        column_definition: ColumnDefinition,
    ) -> None:
        """
        Check the value against the provided column definition and perform validations.

        Args:
            value (Any): The value to be checked.
            index (int): The index of the value in the Excel sheet.
            column_definition (ColumnDefinition): The definition of the column.

        Returns:
            None

        Raises:
            ValidateException: If any validation fails.

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
        Calculates the row number based on the given index.

        Args:
            index (int): The index of the row.

        Returns:
            int: The calculated row number.

        """
        return index + self.__header_row + 2

    def get_result(self, force: bool = False) -> dict:
        """
        Retrieves the result of the Excel validations.

        Args:
            force (bool, optional): If True, the validations will be performed even if there are errors.
                                    Defaults to False.

        Returns:
            dict: A dictionary containing the result of the validations. The keys are the column names and
                    the values are the corresponding values from the Excel sheet.

        Raises:
            ValidateException: If there are errors in the validations and force is set to False.
        """
        excel = self.excel
        if not force and self.has_errors():
            raise ValidateException("Errors found in the validations")

        if excel.empty:
            return {}

        list_colums = list(map(lambda col: col.key, self.__column_details))

        excel = excel.where(pd.notnull(excel), None)
        return excel[list_colums].to_dict("records")

    def has_errors(self) -> bool:
        """
        Check if there are any errors in the Excel object.

        Returns:
            bool: True if there are errors, False otherwise.
        """
        return True if self.__errors_list else False

    def print_errors(self):
        """
        Prints the errors in the error list.

        This method iterates over the list of errors and prints each error message along with its type, line number, and column number (if available).

        Example usage:
        >>> excel_obj = Excel()
        >>> excel_obj.print_errors()
        Error! in line 1: Invalid value
        Error! in line 2, Column A: Missing data
        """
        for error in self.__errors_list:
            if error.column is None:
                print(f"{error.type.value}! in line {error.row}: {error.message}")
            else:
                print(
                    f"{error.type.value}! in line {error.row}, Column {error.column}: {error.message}"
                )

    def get_errors(self) -> list[dict]:
        """
        Returns a list of dictionaries representing the errors.

        Each dictionary in the list contains the following keys:
        - 'type': The type of the error.
        - 'row': The row number where the error occurred.
        - 'column': The column number where the error occurred.
        - 'message': The error message.

        Returns:
        - A list of dictionaries representing the errors.
        """
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
