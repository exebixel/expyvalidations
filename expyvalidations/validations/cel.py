import re

from expyvalidations.exceptions import ValidateException


def validate_cel(word: str) -> str:
    """
    Retorna os números de uma string,
    verificando se existem 11 ou 9 números na string final,
    Se não tiver o 11 ou 9 números o programa é encerrado
    """
    word = re.sub(r"\.0$", "", str(word))
    cel = str("".join(ele for ele in word if ele.isdigit()))
    if len(cel) <= 11 and len(cel) >= 8:
        return cel

    raise ValidateException(f"Invalid phone {word}")
