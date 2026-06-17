"""Compatibilidade com o import legado.

O codigo principal foi movido para automacoes.alt_taxa.
"""

from automacoes.alt_taxa import (
    WebAppAlteradorTaxa,
    automation_status,
    automation_status2,
    iniciar_duplo,
)

__all__ = [
    'WebAppAlteradorTaxa',
    'automation_status',
    'automation_status2',
    'iniciar_duplo',
]
