"""Tests para la función normalize_folder_name."""

import pytest
from src.pst_to_mbox2 import normalize_folder_name


class TestNormalizeFolderName:
    """Tests para normalize_folder_name()."""

    def test_inbox_returns_entrada(self):
        """Test que 'inbox' se traduce a 'Entrada'."""
        assert normalize_folder_name("inbox") == "Entrada"

    def test_sent_items_returns_enviados(self):
        """Test que 'sent items' se traduce a 'Enviados'."""
        assert normalize_folder_name("sent items") == "Enviados"

    def test_deleted_items_returns_eliminados(self):
        """Test que 'deleted items' se traduce a 'Eliminados'."""
        assert normalize_folder_name("deleted items") == "Eliminados"

    def test_unknown_folder_unchanged(self):
        """Test que carpetas desconocidas se mantienen igual."""
        assert normalize_folder_name("Mi Carpeta Personal") == "Mi Carpeta Personal"

    def test_special_characters_replaced(self):
        """Test que caracteres especiales se reemplazan por _."""
        assert normalize_folder_name("Carpeta/Test") == "Carpeta_Test"
        assert normalize_folder_name('Carpeta: especial') == "Carpeta_especial"

    def test_strips_whitespace(self):
        """Test que espacios al inicio y final se eliminan."""
        assert normalize_folder_name("  inbox  ") == "Entrada"