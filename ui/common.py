# -*- coding: utf-8 -*-
"""Itens auxiliares de tabela compartilhados entre as telas."""
from PySide6.QtWidgets import QTableWidgetItem


def numeric_or_text_key(value):
    try:
        return (0, int(str(value).strip()))
    except (TypeError, ValueError):
        return (1, str(value or "").lower())


class SortItem(QTableWidgetItem):
    """QTableWidgetItem que ordena por uma chave Python arbitrária, em vez
    do texto exibido (necessário para códigos numéricos, datas e o ranking
    de status)."""

    def __init__(self, text, sort_key=None):
        super().__init__(text)
        self.sort_key = sort_key if sort_key is not None else text

    def __lt__(self, other):
        if isinstance(other, SortItem):
            try:
                return self.sort_key < other.sort_key
            except TypeError:
                return str(self.sort_key) < str(other.sort_key)
        return super().__lt__(other)
