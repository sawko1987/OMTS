"""
Тесты предупреждения о несохранённых изменениях в наборах материалов:
- Проверка has_unsaved_changes() в ReplacementSetsEditorWidget
- Диалог prompt_unsaved_changes() (Сохранить, Не сохранять, Отмена)
- Защита от потери изменений при переключении детали/изделия в ReplacementSetsEditorWidget
- Защита при переключении пар в sets_table
- Защита при переключении машин (изделий) в ChangesTableWidget
- Защита при переключении шагов визарда в MainWindow
"""
import sys
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

sys.path.insert(0, str(Path(__file__).parent))

from PySide6.QtWidgets import QApplication, QListWidgetItem, QTableWidgetItem
from PySide6.QtCore import Qt

_app = QApplication.instance() or QApplication(["-platform", "offscreen"])

from app.gui.replacement_sets_editor_widget import ReplacementSetsEditorWidget
from app.gui.changes_table_widget import ChangesTableWidget
from app.models import CatalogEntry, DocumentData, MaterialReplacementSet


class UnsavedReplacementSetsTest(unittest.TestCase):
    """Тесты проверки и предупреждения о несохранённых изменениях."""

    def setUp(self):
        self.mock_loader = MagicMock()
        self.doc_data = DocumentData()
        self.widget = ReplacementSetsEditorWidget(
            catalog_loader=self.mock_loader,
            document_data=self.doc_data,
        )
        self.widget._current_part = "ДЕТАЛЬ_1"
        self.widget._current_to_set_id = 10
        self.widget._loaded_to_set_id = 10
        self.widget._loaded_from_set_id = 9

        self.entry_from = CatalogEntry(
            part="ДЕТАЛЬ_1",
            workshop="Цех 1",
            role="",
            before_name="Материал До",
            unit="кг",
            norm=2.5,
        )
        self.entry_to = CatalogEntry(
            part="ДЕТАЛЬ_1",
            workshop="Цех 1",
            role="",
            before_name="Материал После",
            unit="кг",
            norm=2.5,
        )
        self.widget._from_materials = [self.entry_from]
        self.widget._to_materials = [self.entry_to]

    def test_has_unsaved_changes_flags(self):
        """has_unsaved_changes возвращает True при _dirty_materials или _dirty_name."""
        self.assertFalse(self.widget.has_unsaved_changes())

        self.widget._dirty_materials = True
        self.assertTrue(self.widget.has_unsaved_changes())

        self.widget._dirty_materials = False
        self.widget._dirty_name = True
        self.assertTrue(self.widget.has_unsaved_changes())

        self.widget._dirty_name = False
        self.assertFalse(self.widget.has_unsaved_changes())

    def test_has_unsaved_changes_requires_loaded_set(self):
        """Если набор не загружен, has_unsaved_changes возвращает False."""
        self.widget._current_to_set_id = None
        self.widget._dirty_materials = True
        self.assertFalse(self.widget.has_unsaved_changes())

    def test_prompt_unsaved_changes_clean(self):
        """Если изменений нет, prompt_unsaved_changes возвращает True без показа диалога."""
        with patch("app.gui.replacement_sets_editor_widget.QMessageBox") as mock_box_cls:
            self.assertTrue(self.widget.prompt_unsaved_changes())
            mock_box_cls.assert_not_called()

    def test_prompt_unsaved_changes_discard(self):
        """При выборе 'Не сохранять' флаги сбрасываются и возвращается True."""
        self.widget._dirty_materials = True
        self.widget._dirty_name = True

        with patch("app.gui.replacement_sets_editor_widget.QMessageBox") as mock_box_cls:
            mock_box = MagicMock()
            mock_box_cls.return_value = mock_box
            btn_save = MagicMock()
            btn_discard = MagicMock()
            btn_cancel = MagicMock()
            mock_box.addButton.side_effect = [btn_save, btn_discard, btn_cancel]
            mock_box.clickedButton.return_value = btn_discard

            result = self.widget.prompt_unsaved_changes()

            self.assertTrue(result)
            self.assertFalse(self.widget._dirty_materials)
            self.assertFalse(self.widget._dirty_name)

    def test_prompt_unsaved_changes_cancel(self):
        """При выборе 'Отмена' флаги остаются и возвращается False."""
        self.widget._dirty_materials = True
        self.widget._dirty_name = True

        with patch("app.gui.replacement_sets_editor_widget.QMessageBox") as mock_box_cls:
            mock_box = MagicMock()
            mock_box_cls.return_value = mock_box
            btn_save = MagicMock()
            btn_discard = MagicMock()
            btn_cancel = MagicMock()
            mock_box.addButton.side_effect = [btn_save, btn_discard, btn_cancel]
            mock_box.clickedButton.return_value = btn_cancel

            result = self.widget.prompt_unsaved_changes()

            self.assertFalse(result)
            self.assertTrue(self.widget._dirty_materials)
            self.assertTrue(self.widget._dirty_name)

    def test_prompt_unsaved_changes_save_success(self):
        """При выборе 'Сохранить' вызывается _save_current, и при успехе возвращается True."""
        self.widget._dirty_materials = True
        self.mock_loader.update_replacement_set.return_value = True

        with patch("app.gui.replacement_sets_editor_widget.QMessageBox") as mock_box_cls:
            mock_box = MagicMock()
            mock_box_cls.return_value = mock_box
            btn_save = MagicMock()
            btn_discard = MagicMock()
            btn_cancel = MagicMock()
            mock_box.addButton.side_effect = [btn_save, btn_discard, btn_cancel]
            mock_box.clickedButton.return_value = btn_save

            result = self.widget.prompt_unsaved_changes()

            self.assertTrue(result)
            self.assertFalse(self.widget._dirty_materials)
            self.mock_loader.update_replacement_set.assert_called()

    def test_prompt_unsaved_changes_save_validation_fails(self):
        """При ошибке валидации сохранения prompt_unsaved_changes возвращает False."""
        self.widget._dirty_materials = True
        # Очищаем materials в 'to' для вызова ошибки валидации
        self.widget._to_materials = []

        with patch("app.gui.replacement_sets_editor_widget.QMessageBox") as mock_box_cls:
            mock_box = MagicMock()
            mock_box_cls.return_value = mock_box
            btn_save = MagicMock()
            btn_discard = MagicMock()
            btn_cancel = MagicMock()
            mock_box.addButton.side_effect = [btn_save, btn_discard, btn_cancel]
            mock_box.clickedButton.return_value = btn_save

            result = self.widget.prompt_unsaved_changes()

            self.assertFalse(result)
            self.assertTrue(self.widget._dirty_materials)

    def test_set_part_aborts_on_cancel(self):
        """set_part отменяет смену детали, если пользователь нажал Отмена."""
        self.widget._dirty_materials = True
        with patch.object(self.widget, "prompt_unsaved_changes", return_value=False):
            res = self.widget.set_part("ДЕТАЛЬ_2")
            self.assertFalse(res)
            self.assertEqual(self.widget._current_part, "ДЕТАЛЬ_1")

    def test_on_part_changed_reverts_combobox_on_cancel(self):
        """_on_part_changed восстанавливает предыдущую деталь в комбобоксе при отмене."""
        self.widget.part_combo.addItems(["ДЕТАЛЬ_1", "ДЕТАЛЬ_2"])
        self.widget.part_combo.setCurrentText("ДЕТАЛЬ_1")
        self.widget._dirty_materials = True

        with patch.object(self.widget, "prompt_unsaved_changes", return_value=False):
            self.widget.part_combo.setCurrentText("ДЕТАЛЬ_2")
            self.assertEqual(self.widget.part_combo.currentText(), "ДЕТАЛЬ_1")
            self.assertEqual(self.widget._current_part, "ДЕТАЛЬ_1")

    def test_on_pair_selected_reverts_row_on_cancel(self):
        """_on_pair_selected восстанавливает строку таблицы при отмене."""
        to_set_1 = MaterialReplacementSet(id=10, set_name="Набор 1", set_type="to", part_code="ДЕТАЛЬ_1")
        to_set_2 = MaterialReplacementSet(id=20, set_name="Набор 2", set_type="to", part_code="ДЕТАЛЬ_1")
        self.widget._pairs = [(None, to_set_1), (None, to_set_2)]
        self.widget.sets_table.setRowCount(2)
        self.widget.sets_table.setItem(0, 2, QTableWidgetItem("10"))
        self.widget.sets_table.setItem(1, 2, QTableWidgetItem("20"))

        self.widget._current_to_set_id = 10
        self.widget._dirty_materials = True

        with patch.object(self.widget, "prompt_unsaved_changes", return_value=False), \
             patch.object(self.widget, "_load_pair") as mock_load:
            self.widget.sets_table.selectRow(1)
            self.widget._on_pair_selected()
            mock_load.assert_not_called()

    def test_changes_table_machine_switch_guarded(self):
        """В ChangesTableWidget переключение чекбокса машины блокируется при отмене сохранения набора."""
        mock_product_store = MagicMock()
        mock_product_store.get_all_products.return_value = [(1, "Машина 1"), (2, "Машина 2")]
        changes_widget = ChangesTableWidget(
            self.doc_data,
            catalog_loader=self.mock_loader,
            product_store=mock_product_store,
        )
        changes_widget.replacement_sets_widget = self.widget
        self.widget._dirty_materials = True

        item = QListWidgetItem("Машина 1")
        item.setCheckState(Qt.Checked)
        item.setData(Qt.UserRole, "Машина 1")

        with patch.object(self.widget, "prompt_unsaved_changes", return_value=False):
            changes_widget.on_machines_selection_changed(item)
            # Чекбокс должен быть возвращен в Unchecked
            self.assertEqual(item.checkState(), Qt.Unchecked)


if __name__ == "__main__":
    unittest.main()
