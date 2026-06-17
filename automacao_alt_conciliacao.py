"""Compatibilidade com o import legado.

O codigo principal foi movido para automacoes.alt_conciliacao.
"""

from automacoes.alt_conciliacao import WebAppAlteradorConciliacao, automation_status

__all__ = ['WebAppAlteradorConciliacao', 'automation_status']
