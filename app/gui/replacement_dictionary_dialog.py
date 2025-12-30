"""
Диалог выбора варианта замены из глобального словаря замен (до -> варианты после).
"""

from __future__ import annotations

from typing import Optional, List

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QLabel,
    QTableWidget,
    QTableWidgetItem,
    QHeaderView,
    QMessageBox,
)

from app.catalog_loader import CatalogLoader
from app.models import CatalogEntry
from app.gui.material_selection_dialog import MaterialSelectionDialog


class ReplacementDictionaryDialog(QDialog):
    """
    Показать варианты "после" для выбранного материала "до" и применить один вариант.
    Также позволяет добавить вариант в словарь через MaterialSelectionDialog(set_type='to').
    """

    def __init__(
        self,
        catalog_loader: CatalogLoader,
        from_entry: CatalogEntry,
        part_code_for_picker: str,
        parent=None,
    ):
        super().__init__(parent)
        self.catalog_loader = catalog_loader
        self.from_entry = from_entry
        self.part_code_for_picker = part_code_for_picker or ""

        self._selected_to_entry: Optional[CatalogEntry] = None
        self._options: List[CatalogEntry] = []

        self.setWindowTitle("Выбор замены из словаря")
        self.setMinimumSize(900, 550)

        self._init_ui()
        self._reload_options()

    def _init_ui(self) -> None:
        layout = QVBoxLayout(self)

        header = QLabel(
            "Материал 'до': "
            f"{(self.from_entry.workshop or '').strip()} | "
            f"{(self.from_entry.before_name or '').strip()} | "
            f"{(self.from_entry.unit or '').strip()}"
        )
        header.setStyleSheet("font-weight: bold; font-size: 12pt;")
        layout.addWidget(header)

        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(["Цех", "Наименование", "Ед. изм.", "Норма", "Примечание"])
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)

        header_view = self.table.horizontalHeader()
        header_view.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header_view.setSectionResizeMode(1, QHeaderView.Stretch)
        header_view.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header_view.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        header_view.setSectionResizeMode(4, QHeaderView.Stretch)

        self.table.itemDoubleClicked.connect(lambda _item: self._apply_selected())
        layout.addWidget(self.table, 1)

        btns = QHBoxLayout()
        self.btn_add = QPushButton("Добавить вариант...")
        self.btn_add.clicked.connect(self._add_variant)
        btns.addWidget(self.btn_add)

        self.btn_delete = QPushButton("Удалить вариант")
        self.btn_delete.clicked.connect(self._delete_selected_variant)
        btns.addWidget(self.btn_delete)

        btns.addStretch()

        self.btn_cancel = QPushButton("Отмена")
        self.btn_cancel.clicked.connect(self.reject)
        btns.addWidget(self.btn_cancel)

        self.btn_apply = QPushButton("Применить")
        self.btn_apply.setDefault(True)
        self.btn_apply.clicked.connect(self._apply_selected)
        btns.addWidget(self.btn_apply)

        layout.addLayout(btns)

    def _reload_options(self) -> None:
        self._options = self.catalog_loader.get_replacement_dictionary_options(self.from_entry)
        self.table.setRowCount(len(self._options))
        for row, entry in enumerate(self._options):
            w = QTableWidgetItem(entry.workshop or "")
            n = QTableWidgetItem(entry.before_name or "")
            u = QTableWidgetItem(entry.unit or "")
            norm_text = "" if (entry.norm or 0) == 0 else f"{float(entry.norm or 0.0):.4f}"
            norm = QTableWidgetItem(norm_text)
            norm.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            c = QTableWidgetItem(entry.comment or "")

            # keep entry for later
            for it in (w, n, u, norm, c):
                it.setData(Qt.UserRole, entry)

            self.table.setItem(row, 0, w)
            self.table.setItem(row, 1, n)
            self.table.setItem(row, 2, u)
            self.table.setItem(row, 3, norm)
            self.table.setItem(row, 4, c)

        if self._options:
            self.table.selectRow(0)

    def _get_selected_option(self) -> Optional[CatalogEntry]:
        row = self.table.currentRow()
        if row < 0 or row >= len(self._options):
            return None
        return self._options[row]

    def _add_variant(self) -> None:
        picker = MaterialSelectionDialog(
            self.part_code_for_picker,
            self.catalog_loader,
            self,
            set_type="to",
        )
        if not picker.exec():
            return

        selected = picker.get_selected_entries()
        if not selected:
            return

        added_any = False
        for to_entry in selected:
            added = self.catalog_loader.add_replacement_dictionary_link(self.from_entry, to_entry)
            added_any = added_any or added

        if not added_any:
            QMessageBox.information(self, "Информация", "Выбранные варианты уже есть в словаре.")

        self._reload_options()

    def _delete_selected_variant(self) -> None:
        opt = self._get_selected_option()
        if not opt:
            return

        reply = QMessageBox.question(
            self,
            "Подтверждение удаления",
            "Удалить выбранный вариант из словаря замен?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            return

        ok = self.catalog_loader.delete_replacement_dictionary_link(self.from_entry, opt)
        if not ok:
            QMessageBox.warning(self, "Ошибка", "Не удалось удалить вариант (возможно, уже удалён).")
        self._reload_options()

    def _apply_selected(self) -> None:
        opt = self._get_selected_option()
        if not opt:
            QMessageBox.warning(self, "Ошибка", "Выберите вариант замены")
            return
        self._selected_to_entry = opt
        self.accept()

    def get_selected_to_entry(self) -> Optional[CatalogEntry]:
        return self._selected_to_entry


