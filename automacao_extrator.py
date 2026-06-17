"""Compatibilidade com o import legado.

O codigo principal foi movido para automacoes.extrator.
"""

from automacoes.extrator import WebAppExtrator, extraction_status

__all__ = ['WebAppExtrator', 'extraction_status']
