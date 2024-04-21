from pandas import DataFrame
from alive_progress import alive_bar
from expyvalidations.config import config_bar_excel


from expyvalidations.exceptions import ValidateException


def validate_duplications(data: DataFrame, keys: list[str]) -> DataFrame:
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
    raise ValidateException(erros)