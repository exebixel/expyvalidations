import json
from typing import Any

from pandas import ExcelFile

from expyvalidations.excel.excel import ExpyValidations
from expyvalidations.exceptions import ValidateException


def servico(book: ExcelFile, header: int = 1, reset: bool = False):
    """
    Book: planilha com todos os dados \n
    reset: True se todos os serviços do sistema
    serão excluídos para cadastrar os serviços da planilha \n

    Nessa função tem toda a descrição do json que vai ser enviado
    para as rotas de cadastro do sistema

    Return None
    """
    print("analyzing spreadsheet")

    ex = ExpyValidations(path_file=book, sheet_name="Servico", header_row=header)

    ex.add_column(key="flagAtivo", name="ativo", required=True, default=True)
    ex.add_column(key="descricao", name="nome", custom_function_before=check_descricao)
    ex.add_column(key="gservId", name="grupo", custom_function_after=check_descricao)
    ex.add_column(key="preco", name="valor", default=1, types="float")
    ex.add_column(
        key="comissao",
        name="comissao",
        default=0,
        types="float",
        custom_function_after=check_comissao,
    )
    ex.add_column(
        key="tempoExecucao", name="execucao", default="00:30:00", types="time"
    )
    ex.add_column(
        key="custosGerais", name="custo", required=False, default=0, types="float"
    )
    ex.add_column(
        key="intervaloMarcacao",
        name="intervalo",
        required=False,
        default="00:10:00",
        types="time",
    )
    ex.add_column(
        key="permiteEncaixe", name="encaixe", required=False, default=True, types="bool"
    )
    ex.add_column(
        key="permiteSimultaneidade",
        name="simultaneidade",
        required=False,
        default=True,
        types="bool",
    )
    ex.add_column(
        key="valPercComissao", name="tipo comissao", required=False, default="P"
    )
    ex.add_column(key="valPercCustos", name="tipo custo", required=False, default="P")
    ex.add_column(
        key="flagMobilidade",
        name="mobilidade",
        required=False,
        default=True,
        types="bool",
    )

    ex.check_all(check_row=row_test, check_duplicated_keys=["descricao"])
    ex.print_errors()
    list_errors = ex.get_errors()
    for error in list_errors:
        print(error)

    data = ex.get_result(force=True)
    for item in data:
        print("{")
        for key, value in item.items():
            print(f"    {key}: {value}")
        print("}")


def check_descricao(value: Any) -> str:
    if value is None:
        raise ValidateException("Empty value")
    return value


def check_comissao(value: Any) -> float:
    if value <= 1:
        value = value * 100
    return value


def row_test(data: dict) -> dict:
    return data


if __name__ == "__main__":
    servico("Parametrização Sistema One Beleza.xlsx")
