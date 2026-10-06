"""
Тесты логики автоматического именования нового набора при вынесении ("Вынести в новый набор"):
- Если в 'до' выбран ровно 1 материал: имя берется из названия материала без показа диалога ввода,
  при этом количество выбранных материалов в 'после' не влияет на присвоение имени.
- Если в 'до' выбрано 0 или >1 материалов: запрашивается имя через диалог ввода.
"""
import sys
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

sys.path.insert(0, str(Path(__file__).parent))

from PySide6.QtWidgets import QApplication

# Инициализируем оффлайн QApplication для headless тестов
_app = QApplication.instance() or QApplication(["-platform", "offscreen"])

from app.gui.replacement_sets_editor_widget import ReplacementSetsEditorWidget
from app.models import CatalogEntry


class SplitCopyNamingTest(unittest.TestCase):
    """Тесты метода _determine_split_name и логики _split_copy."""

    def setUp(self):
        self.mock_loader = MagicMock()
        self.widget = ReplacementSetsEditorWidget(catalog_loader=self.mock_loader)
        self.entry_a = CatalogEntry(
            part="ДЕТАЛЬ_1",
            workshop="Цех 1",
            role="",
            before_name="Сталь 45 ГОСТ 1050-2013",
            unit="кг",
            norm=2.5,
        )
        self.entry_b = CatalogEntry(
            part="ДЕТАЛЬ_1",
            workshop="Цех 1",
            role="",
            before_name="Круг 20 ГОСТ 2590-2006",
            unit="кг",
            norm=1.0,
        )
        self.entry_to1 = CatalogEntry(
            part="ДЕТАЛЬ_1",
            workshop="Цех 1",
            role="",
            before_name="Заменитель 1",
            unit="кг",
            norm=2.0,
        )
        self.entry_to2 = CatalogEntry(
            part="ДЕТАЛЬ_1",
            workshop="Цех 1",
            role="",
            before_name="Заменитель 2",
            unit="кг",
            norm=1.0,
        )

        self.widget._from_materials = [self.entry_a, self.entry_b]
        self.widget._to_materials = [self.entry_to1, self.entry_to2]
        self.widget._render_table(self.widget.from_table, self.widget._from_materials, is_from=True)
        self.widget._render_table(self.widget.to_table, self.widget._to_materials, is_from=False)

    def test_single_from_material_auto_named_without_dialog(self):
        """Если в 'до' выбран 1 материал — имя присваивается по материалу без диалога."""
        with patch("app.gui.replacement_sets_editor_widget.QInputDialog.getText") as mock_input:
            name, ok = self.widget._determine_split_name([0])
            self.assertTrue(ok)
            self.assertEqual(name, "Сталь 45 ГОСТ 1050-2013")
            mock_input.assert_not_called()

    def test_single_from_material_second_row_auto_named(self):
        """Проверка выбора второй строки из 'до'."""
        with patch("app.gui.replacement_sets_editor_widget.QInputDialog.getText") as mock_input:
            name, ok = self.widget._determine_split_name([1])
            self.assertTrue(ok)
            self.assertEqual(name, "Круг 20 ГОСТ 2590-2006")
            mock_input.assert_not_called()

    def test_zero_from_materials_prompts_dialog(self):
        """Если в 'до' выбрано 0 материалов — запрашивается ввод имени."""
        with patch("app.gui.replacement_sets_editor_widget.QInputDialog.getText") as mock_input:
            mock_input.return_value = ("Набор только после", True)
            name, ok = self.widget._determine_split_name([])
            self.assertTrue(ok)
            self.assertEqual(name, "Набор только после")
            mock_input.assert_called_once()

    def test_multiple_from_materials_prompts_dialog(self):
        """Если в 'до' выбрано более 1 материала — запрашивается ввод имени."""
        with patch("app.gui.replacement_sets_editor_widget.QInputDialog.getText") as mock_input:
            mock_input.return_value = ("Групповой набор", True)
            name, ok = self.widget._determine_split_name([0, 1])
            self.assertTrue(ok)
            self.assertEqual(name, "Групповой набор")
            mock_input.assert_called_once()

    def test_multiple_from_materials_dialog_cancelled(self):
        """Если пользователь отменил диалог ввода имени — возвращается ok=False."""
        with patch("app.gui.replacement_sets_editor_widget.QInputDialog.getText") as mock_input:
            mock_input.return_value = ("", False)
            name, ok = self.widget._determine_split_name([0, 1])
            self.assertFalse(ok)
            self.assertIsNone(name)

    def test_single_from_material_blank_name_falls_back_to_dialog(self):
        """Если у выбранного материала пустое наименование — запрашивается ввод через диалог."""
        self.widget._from_materials[0].before_name = ""
        self.widget._render_table(self.widget.from_table, self.widget._from_materials, is_from=True)

        with patch("app.gui.replacement_sets_editor_widget.QInputDialog.getText") as mock_input:
            mock_input.return_value = ("Имя взамен пустого", True)
            name, ok = self.widget._determine_split_name([0])
            self.assertTrue(ok)
            self.assertEqual(name, "Имя взамен пустого")
            mock_input.assert_called_once()

    def test_split_copy_full_workflow_single_from_multiple_to(self):
        """Проверка полного вызова _split_copy: 1 материал в 'до' и 2 в 'после'."""
        from PySide6.QtCore import QItemSelectionModel
        self.widget._current_to_set_id = 99
        self.widget.from_table.selectRow(0)
        self.widget.to_table.selectionModel().select(
            self.widget.to_table.model().index(0, 0),
            QItemSelectionModel.Select | QItemSelectionModel.Rows
        )
        self.widget.to_table.selectionModel().select(
            self.widget.to_table.model().index(1, 0),
            QItemSelectionModel.Select | QItemSelectionModel.Rows
        )

        self.mock_loader.split_replacement_pair_copy.return_value = 101

        with patch("app.gui.replacement_sets_editor_widget.QInputDialog.getText") as mock_input:
            self.widget._split_copy()
            mock_input.assert_not_called()

        self.mock_loader.split_replacement_pair_copy.assert_called_once_with(
            99,
            [0],
            [0, 1],
            "Сталь 45 ГОСТ 1050-2013",
        )


if __name__ == "__main__":
    unittest.main()
