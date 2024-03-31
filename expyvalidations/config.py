""" Informações globais de configuração
"""
from alive_progress import config_handler

RESOLVE_ERROS = False
NO_WARNING = False


def config_bar_excel():
    """ Configuração padrão da barra de progresso os módulos de excel
    """
    config_handler.set_global(bar=None,
                              spinner=False,
                              receipt=False,
                              enrich_print=False,
                              stats=False,
                              elapsed=False)


class CheckException(Exception):
    """ Exception especifica para algum erro de verificação do modulo de excel
    """
