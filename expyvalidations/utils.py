import re
import unicodedata


def deemphasize(word: str) -> str:
    """
    Retorna os caracteres da string em seu equivalente em Latin,
    em outras palavras, tira os acentos da string
    """
    if word is None:
        return word
    word = str(word)
    nfkd = unicodedata.normalize("NFKD", word)
    word = "".join([c for c in nfkd if not unicodedata.combining(c)])
    return word.lower()


def string_normalize(word: str) -> str:
    """
    Retorna todos as letras e números de uma string
    """
    word = deemphasize(word)
    # Retornar a palavra apenas com números, letras e espaço
    return re.sub(r"[^a-zA-Z0-9 \\]", "", word)
