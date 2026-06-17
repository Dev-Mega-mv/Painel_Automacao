"""Pacote das automacoes centralizadas.

As classes ficam nos submodulos para evitar carregar Selenium e dependencias de
automacao logo ao importar o pacote.
"""

__all__ = [
    'WebAppAtivador',
    'WebAppAlteradorConciliacao',
    'WebAppAlteradorTaxa',
    'WebAppExtrator',
    'automation_status',
    'conciliacao_status',
    'taxa_status',
    'taxa_status2',
    'extraction_status',
]


def __getattr__(name):
    if name in ('WebAppAtivador', 'automation_status'):
        from . import cielo_cardse
        if name == 'WebAppAtivador':
            return cielo_cardse.WebAppAtivador
        return cielo_cardse.automation_status

    if name in ('WebAppAlteradorConciliacao', 'conciliacao_status'):
        from . import alt_conciliacao
        if name == 'WebAppAlteradorConciliacao':
            return alt_conciliacao.WebAppAlteradorConciliacao
        return alt_conciliacao.automation_status

    if name in ('WebAppAlteradorTaxa', 'taxa_status', 'taxa_status2'):
        from . import alt_taxa
        if name == 'WebAppAlteradorTaxa':
            return alt_taxa.WebAppAlteradorTaxa
        if name == 'taxa_status':
            return alt_taxa.automation_status
        return alt_taxa.automation_status2

    if name in ('WebAppExtrator', 'extraction_status'):
        from . import extrator
        if name == 'WebAppExtrator':
            return extrator.WebAppExtrator
        return extrator.extraction_status

    raise AttributeError(name)
