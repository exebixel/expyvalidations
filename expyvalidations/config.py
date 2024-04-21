""" Informações globais de configuração
"""

from alive_progress import config_handler


def config_bar_excel():
    """Configuração padrão da barra de progresso os módulos de excel"""
    config_handler.set_global(
        bar=None,
        spinner=False,
        receipt=True,
        enrich_print=False,
        stats=False,
        elapsed=False,
    )
