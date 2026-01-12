"""
Вкладка для управления наборами материалов (парами from+to) по выбранной детали.
"""

from __future__ import annotations

from typing import Optional
import logging

from PySide6.QtCore import Qt, QEvent
from PySide6.QtGui import QKeyEvent
from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QComboBox,
    QPushButton,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QHeaderView,
    QMessageBox,
    QLineEdit,
    QGroupBox,
    QInputDialog,
)

from app.catalog_loader import CatalogLoader
from app.models import (
    CatalogEntry,
    DocumentData,
    PartChanges,
    MaterialChange,
    MaterialReplacementSet,
)
from app.gui.part_creation_dialog import MaterialEntryDialog
from app.gui.material_selection_dialog import MaterialSelectionDialog
from app.gui.replacement_dictionary_dialog import ReplacementDictionaryDialog
from app.gui.set_selection_dialog import ReplacementSetSelectionDialog
from app.history_store import HistoryStore
from app.product_store import ProductStore

logger = logging.getLogger(__name__)


class ReplacementSetsEditorWidget(QWidget):
    """
    Вкладка-редактор наборов для детали.

    На первом этапе — каркас UI + выбор детали. Логику операций (copy/delete/split/save)
    добавим дальше по плану.
    """

    def __init__(
        self,
        catalog_loader: CatalogLoader,
        parent: Optional[QWidget] = None,
        document_data: Optional[DocumentData] = None,
        history_store: Optional[HistoryStore] = None,
        product_store: Optional[ProductStore] = None,
        get_current_additional_page: Optional[callable] = None,
        changes_widget: Optional[QWidget] = None,
    ):
        super().__init__(parent)
        self.catalog_loader = catalog_loader
        self.document_data = document_data
        self.history_store = history_store
        self.product_store = product_store
        self.get_current_additional_page = get_current_additional_page
        self.changes_widget = changes_widget
        self._current_part: Optional[str] = None
        self._pairs: list[tuple[Optional[object], object]] = []
        self._current_to_set_id: Optional[int] = None
        self._loaded_from_set_id: Optional[int] = None
        self._loaded_to_set_id: Optional[int] = None
        self._from_materials: list[CatalogEntry] = []
        self._to_materials: list[CatalogEntry] = []
        self._loaded_set_name: Optional[str] = None
        self._dirty_name: bool = False
        self._dirty_materials: bool = False

        self._init_ui()

    def _init_ui(self) -> None:
        layout = QVBoxLayout(self)

        title = QLabel("Наборы материалов по детали")
        title.setStyleSheet("font-size: 16px; font-weight: bold;")
        layout.addWidget(title)

        # Выбор детали (fallback — даже если не синхронизированы с вкладкой изменений)
        top = QHBoxLayout()
        top.addWidget(QLabel("Деталь:"))
        self.part_combo = QComboBox()
        self.part_combo.setEditable(True)
        self.part_combo.setMinimumWidth(220)
        self.part_combo.currentTextChanged.connect(self._on_part_changed)
        top.addWidget(self.part_combo)

        self.btn_refresh = QPushButton("Обновить")
        self.btn_refresh.clicked.connect(self.refresh)
        top.addWidget(self.btn_refresh)
        top.addStretch()
        layout.addLayout(top)

        splitter = QSplitter(Qt.Horizontal)
        layout.addWidget(splitter, 1)

        # Левая часть: список пар наборов
        left = QWidget()
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(0, 0, 0, 0)

        left_buttons = QHBoxLayout()
        self.btn_new = QPushButton("Новый")
        self.btn_new.clicked.connect(self._create_new_pair)
        left_buttons.addWidget(self.btn_new)

        self.btn_copy = QPushButton("Копировать пару")
        self.btn_copy.clicked.connect(self._copy_pair)
        left_buttons.addWidget(self.btn_copy)

        self.btn_delete = QPushButton("Удалить пару")
        self.btn_delete.clicked.connect(self._delete_pair)
        left_buttons.addWidget(self.btn_delete)

        left_buttons.addStretch()
        left_layout.addLayout(left_buttons)

        self.sets_table = QTableWidget()
        self.sets_table.setColumnCount(6)
        self.sets_table.setHorizontalHeaderLabels(
            ["Дата", "Название", "ID to", "ID from", "Кол-во from", "Кол-во to"]
        )
        header = self.sets_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        left_layout.addWidget(self.sets_table)
        self.sets_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.sets_table.setSelectionMode(QTableWidget.SingleSelection)
        self.sets_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.sets_table.itemSelectionChanged.connect(self._on_pair_selected)
        splitter.addWidget(left)

        # Правая часть: редактор выбранной пары
        right = QWidget()
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(0, 0, 0, 0)

        meta = QHBoxLayout()
        meta.addWidget(QLabel("Название набора:"))
        self.name_edit = QLineEdit()
        self.name_edit.setPlaceholderText("Например: ПЗУ (часть 1)")
        self.name_edit.textEdited.connect(self._on_name_edited)
        meta.addWidget(self.name_edit, 0.5)

        self.btn_revert = QPushButton("Отменить изменения")
        self.btn_revert.clicked.connect(self._revert_loaded)
        meta.addWidget(self.btn_revert)

        self.btn_save = QPushButton("Сохранить")
        self.btn_save.clicked.connect(self._save_current)
        meta.addWidget(self.btn_save)

        self.btn_save_and_add = QPushButton("Сохранить и добавить")
        self.btn_save_and_add.clicked.connect(self._save_and_add_to_document)
        meta.addWidget(self.btn_save_and_add)
        right_layout.addLayout(meta)

        # Таблица "до"
        from_group = QGroupBox("Набор материалов 'до' (from)")
        from_layout = QVBoxLayout(from_group)
        from_btns = QHBoxLayout()
        self.btn_from_add_new = QPushButton("Добавить новый")
        self.btn_from_add_new.clicked.connect(lambda: self._add_material(is_from=True, from_catalog=False))
        from_btns.addWidget(self.btn_from_add_new)
        self.btn_from_add_catalog = QPushButton("Выбрать из каталога")
        self.btn_from_add_catalog.clicked.connect(lambda: self._add_material(is_from=True, from_catalog=True))
        from_btns.addWidget(self.btn_from_add_catalog)
        self.btn_from_remove = QPushButton("Удалить выбранный")
        self.btn_from_remove.setText("Удалить выбранные")
        self.btn_from_remove.clicked.connect(lambda: self._remove_selected_multi(is_from=True))
        from_btns.addWidget(self.btn_from_remove)

        self.btn_from_pick_replacement = QPushButton("Выбор замены")
        self.btn_from_pick_replacement.clicked.connect(self._pick_replacement_for_selected_from)
        from_btns.addWidget(self.btn_from_pick_replacement)
        from_btns.addStretch()
        from_layout.addLayout(from_btns)

        self.from_table = QTableWidget()
        self.from_table.setColumnCount(4)
        self.from_table.setHorizontalHeaderLabels(["Цех", "Наименование", "Ед. изм.", "Норма"])
        self.from_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.from_table.setSelectionMode(QTableWidget.ExtendedSelection)
        self.from_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.from_table.itemChanged.connect(lambda item: self._on_material_item_changed(item, is_from=True))
        self.from_table.installEventFilter(self)
        from_layout.addWidget(self.from_table)
        right_layout.addWidget(from_group)

        # Таблица "после"
        to_group = QGroupBox("Набор материалов 'после' (to)")
        to_layout = QVBoxLayout(to_group)
        to_btns = QHBoxLayout()
        self.btn_to_add_new = QPushButton("Добавить новый")
        self.btn_to_add_new.clicked.connect(lambda: self._add_material(is_from=False, from_catalog=False))
        to_btns.addWidget(self.btn_to_add_new)
        self.btn_to_add_catalog = QPushButton("Выбрать из каталога")
        self.btn_to_add_catalog.clicked.connect(lambda: self._add_material(is_from=False, from_catalog=True))
        to_btns.addWidget(self.btn_to_add_catalog)
        self.btn_to_remove = QPushButton("Удалить выбранный")
        self.btn_to_remove.setText("Удалить выбранные")
        self.btn_to_remove.clicked.connect(lambda: self._remove_selected_multi(is_from=False))
        to_btns.addWidget(self.btn_to_remove)
        to_btns.addStretch()
        to_layout.addLayout(to_btns)

        self.to_table = QTableWidget()
        self.to_table.setColumnCount(4)
        self.to_table.setHorizontalHeaderLabels(["Цех", "Наименование", "Ед. изм.", "Норма"])
        self.to_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.to_table.setSelectionMode(QTableWidget.ExtendedSelection)
        self.to_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.to_table.itemChanged.connect(lambda item: self._on_material_item_changed(item, is_from=False))
        self.to_table.installEventFilter(self)
        to_layout.addWidget(self.to_table)
        right_layout.addWidget(to_group)

        # Split button and Add to document button
        split_row = QHBoxLayout()
        self.btn_add_to_document = QPushButton("Добавить деталь")
        self.btn_add_to_document.clicked.connect(self._add_part_to_document)
        split_row.addWidget(self.btn_add_to_document)
        self.btn_split = QPushButton("Вынести в новый набор")
        self.btn_split.clicked.connect(self._split_copy)
        split_row.addWidget(self.btn_split)
        split_row.addStretch()
        right_layout.addLayout(split_row)

        self._set_editor_enabled(False)
        splitter.addWidget(right)

        splitter.setStretchFactor(0, 2)
        splitter.setStretchFactor(1, 3)

    def load_catalog(self) -> None:
        """Перезагрузить список деталей (для комбобокса)"""
        parts = self.catalog_loader.get_all_parts()
        current = self.part_combo.currentText()
        self.part_combo.blockSignals(True)
        self.part_combo.clear()
        self.part_combo.addItems(parts)
        if current:
            idx = self.part_combo.findText(current)
            if idx >= 0:
                self.part_combo.setCurrentIndex(idx)
            else:
                self.part_combo.setEditText(current)
        self.part_combo.blockSignals(False)

    def set_part(self, part_code: str) -> None:
        part_code = (part_code or "").strip()
        if not part_code:
            return
        self.part_combo.setEditText(part_code)
        self._current_part = part_code
        self.refresh()

    def _on_part_changed(self, text: str) -> None:
        part = (text or "").strip()
        self._current_part = part if part else None
        self.refresh()

    def refresh(self) -> None:
        """Обновить список наборов для текущей детали"""
        self.sets_table.setRowCount(0)
        self._pairs = []
        self._current_to_set_id = None
        self._from_materials = []
        self._to_materials = []
        self.name_edit.setText("")
        self._loaded_set_name = None
        self._dirty_name = False
        self._dirty_materials = False
        self._set_editor_enabled(False)

        if not self._current_part:
            return

        all_sets = self.catalog_loader.get_replacement_sets_by_part(self._current_part)
        to_sets = [s for s in all_sets if s.set_type == "to"]
        from_sets = [s for s in all_sets if s.set_type == "from"]

        # Для каждого to подбираем from
        for to_set in to_sets:
            from_set = self.catalog_loader.find_matching_from_set(from_sets, to_set)
            self._pairs.append((from_set, to_set))

        self.sets_table.setRowCount(len(self._pairs))
        for row, (from_set, to_set) in enumerate(self._pairs):
            created = ""
            if to_set.created_at:
                created = str(to_set.created_at)
            name = to_set.set_name or ""

            id_to = str(to_set.id or "")
            id_from = str(from_set.id if from_set and from_set.id else "")
            cnt_from = str(len(from_set.materials) if from_set else 0)
            cnt_to = str(len(to_set.materials) if to_set.materials else 0)

            item0 = QTableWidgetItem(created)
            item1 = QTableWidgetItem(name)
            item2 = QTableWidgetItem(id_to)
            item3 = QTableWidgetItem(id_from)
            item4 = QTableWidgetItem(cnt_from)
            item5 = QTableWidgetItem(cnt_to)
            item2.setData(Qt.UserRole, int(to_set.id or 0))

            self.sets_table.setItem(row, 0, item0)
            self.sets_table.setItem(row, 1, item1)
            self.sets_table.setItem(row, 2, item2)
            self.sets_table.setItem(row, 3, item3)
            self.sets_table.setItem(row, 4, item4)
            self.sets_table.setItem(row, 5, item5)

        if self._pairs:
            self.sets_table.selectRow(0)
        
        # Обновляем состояние кнопки "Добавить деталь"
        if hasattr(self, 'btn_add_to_document'):
            self.btn_add_to_document.setEnabled(
                self.document_data is not None and self._current_part is not None
            )

    # -----------------------
    # Pair selection / editor
    # -----------------------

    def _set_editor_enabled(self, enabled: bool) -> None:
        for w in [
            self.name_edit,
            self.btn_revert,
            self.btn_save,
            self.btn_save_and_add,
            self.btn_from_add_new,
            self.btn_from_add_catalog,
            self.btn_from_remove,
            self.btn_to_add_new,
            self.btn_to_add_catalog,
            self.btn_to_remove,
            self.btn_split,
            self.from_table,
            self.to_table,
        ]:
            w.setEnabled(enabled)
        # Кнопка "Добавить деталь" доступна только если есть document_data и выбрана деталь
        if hasattr(self, 'btn_add_to_document'):
            self.btn_add_to_document.setEnabled(enabled and self.document_data is not None and self._current_part is not None)
        # Кнопка "Сохранить и добавить" доступна только если есть document_data и выбрана деталь
        if hasattr(self, 'btn_save_and_add'):
            self.btn_save_and_add.setEnabled(enabled and self.document_data is not None and self._current_part is not None)

    def _on_pair_selected(self) -> None:
        row = self.sets_table.currentRow()
        if row < 0 or row >= len(self._pairs):
            self._set_editor_enabled(False)
            return
        _, to_set = self._pairs[row]
        if not to_set.id:
            self._set_editor_enabled(False)
            return
        self._load_pair(to_set.id)

    def _load_pair(self, to_set_id: int) -> None:
        from_set, to_set = self.catalog_loader.get_replacement_pair_by_to_id(to_set_id)
        if not to_set or not to_set.id:
            self._set_editor_enabled(False)
            return

        self._current_to_set_id = to_set.id
        self._loaded_from_set_id = from_set.id if from_set and from_set.id else None
        self._loaded_to_set_id = to_set.id

        self._loaded_set_name = to_set.set_name
        self._dirty_name = False
        self._dirty_materials = False
        self.name_edit.setText(to_set.set_name or "")

        self._from_materials = list(from_set.materials) if from_set else []
        self._to_materials = list(to_set.materials) if to_set.materials else []

        self._render_table(self.from_table, self._from_materials, is_from=True)
        self._render_table(self.to_table, self._to_materials, is_from=False)

        self._set_editor_enabled(True)

    def _on_name_edited(self, _text: str) -> None:
        self._dirty_name = True

    def eventFilter(self, watched, event) -> bool:
        # Массовое удаление по Delete в таблицах
        if event.type() == QEvent.KeyPress:
            is_from_table = hasattr(self, "from_table") and watched is getattr(self, "from_table", None)
            is_to_table = hasattr(self, "to_table") and watched is getattr(self, "to_table", None)
            if is_from_table or is_to_table:
                key_event = event  # type: ignore[assignment]
                if isinstance(key_event, QKeyEvent) and key_event.key() == Qt.Key_Delete:
                    self._remove_selected_multi(is_from=is_from_table)
                    return True
        return super().eventFilter(watched, event)

    def _revert_loaded(self) -> None:
        if not self._current_to_set_id:
            return
        self._load_pair(self._current_to_set_id)

    def _render_table(self, table: QTableWidget, materials: list[CatalogEntry], is_from: bool) -> None:
        table.blockSignals(True)
        table.setRowCount(len(materials))
        for row, entry in enumerate(materials):
            w = QTableWidgetItem(entry.workshop or "")
            w.setData(Qt.UserRole, entry)
            n = QTableWidgetItem(entry.before_name or "")
            n.setData(Qt.UserRole, entry)
            u = QTableWidgetItem(entry.unit or "")
            u.setData(Qt.UserRole, entry)

            if is_from:
                norm_text = f"{entry.norm:.4f}" if entry.norm is not None else ""
            else:
                norm_text = f"{entry.norm:.4f}" if entry.norm and entry.norm > 0 else ""
            norm = QTableWidgetItem(norm_text)
            norm.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            norm.setData(Qt.UserRole, entry)

            table.setItem(row, 0, w)
            table.setItem(row, 1, n)
            table.setItem(row, 2, u)
            table.setItem(row, 3, norm)
        table.blockSignals(False)

    def _on_material_item_changed(self, item: QTableWidgetItem, is_from: bool) -> None:
        entry = item.data(Qt.UserRole)
        if not entry:
            return
        col = item.column()
        val = item.text().strip()
        if col == 0:
            entry.workshop = val
        elif col == 1:
            entry.before_name = val
        elif col == 2:
            entry.unit = val
        elif col == 3:
            if not val and not is_from:
                entry.norm = 0.0
                return
            try:
                new_norm = float(val.replace(",", "."))
                if is_from and new_norm <= 0:
                    return
                if new_norm < 0:
                    return
                entry.norm = new_norm
                # форматируем
                item.setText("" if (not is_from and new_norm == 0) else f"{new_norm:.4f}")
            except ValueError:
                return
        self._dirty_materials = True

    def _remove_selected(self, is_from: bool) -> None:
        table = self.from_table if is_from else self.to_table
        mats = self._from_materials if is_from else self._to_materials
        row = table.currentRow()
        if row < 0 or row >= len(mats):
            return
        mats.pop(row)
        self._dirty_materials = True
        self._render_table(table, mats, is_from=is_from)

    def _remove_selected_multi(self, is_from: bool) -> None:
        table = self.from_table if is_from else self.to_table
        mats = self._from_materials if is_from else self._to_materials
        rows = sorted({idx.row() for idx in table.selectionModel().selectedRows()})
        if not rows:
            return

        reply = QMessageBox.question(
            self,
            "Подтверждение удаления",
            f"Удалить выбранные строки ({len(rows)} шт.)?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            return

        for r in sorted(rows, reverse=True):
            if 0 <= r < len(mats):
                mats.pop(r)

        self._dirty_materials = True
        self._render_table(table, mats, is_from=is_from)

    def _add_material(self, is_from: bool, from_catalog: bool) -> None:
        if not self._current_part:
            QMessageBox.warning(self, "Ошибка", "Сначала выберите деталь")
            return
        mats = self._from_materials if is_from else self._to_materials
        table = self.from_table if is_from else self.to_table

        if from_catalog:
            dialog = MaterialSelectionDialog(
                self._current_part,
                self.catalog_loader,
                self,
                set_type="from" if is_from else "to",
            )
            if dialog.exec():
                mats.extend(dialog.get_selected_entries())
        else:
            dlg = MaterialEntryDialog(self, norm_required=is_from)
            if dlg.exec():
                entry = dlg.get_entry()
                if entry:
                    mats.append(entry)

        self._dirty_materials = True
        self._render_table(table, mats, is_from=is_from)

    def _pick_replacement_for_selected_from(self) -> None:
        if not self._current_part:
            QMessageBox.warning(self, "Ошибка", "Сначала выберите деталь")
            return
        if not self._from_materials:
            QMessageBox.warning(self, "Ошибка", "Сначала добавьте материалы в набор 'до'")
            return

        row = self.from_table.currentRow()
        if row < 0 or row >= len(self._from_materials):
            QMessageBox.warning(self, "Ошибка", "Выберите строку в наборе 'до'")
            return

        from_entry = self._from_materials[row]
        if not (from_entry.workshop or "").strip() or not (from_entry.before_name or "").strip() or not (from_entry.unit or "").strip():
            QMessageBox.warning(self, "Ошибка", "В выбранной строке 'до' должны быть заполнены Цех, Наименование и Ед. изм.")
            return

        dlg = ReplacementDictionaryDialog(
            self.catalog_loader,
            from_entry=from_entry,
            part_code_for_picker=self._current_part,
            parent=self,
        )
        if not dlg.exec():
            return
        to_entry = dlg.get_selected_to_entry()
        if not to_entry:
            return

        # Применяем по индексу строки "до" (модель сопоставления по индексу)
        while len(self._to_materials) < row:
            self._to_materials.append(
                CatalogEntry(
                    part="",
                    workshop="",
                    role="",
                    before_name="",
                    unit="",
                    norm=0.0,
                    comment="",
                    id=None,
                    is_part_of_set=False,
                    replacement_set_id=None,
                )
            )

        if len(self._to_materials) == row:
            self._to_materials.append(to_entry)
        else:
            self._to_materials[row] = to_entry

        self._dirty_materials = True
        self._render_table(self.to_table, self._to_materials, is_from=False)

    # -----------------------
    # Operations (new/copy/delete/save/split)
    # -----------------------

    def _create_new_pair(self) -> None:
        if not self._current_part:
            QMessageBox.warning(self, "Ошибка", "Сначала выберите деталь")
            return

        default_name = "Новый набор"
        name, ok = QInputDialog.getText(self, "Новый набор", "Название набора:", text=default_name)
        if not ok:
            return

        new_from_id = self.catalog_loader.add_replacement_set(
            self._current_part,
            [],
            [],
            set_name=name.strip() or None,
        )
        if not new_from_id:
            QMessageBox.warning(self, "Ошибка", "Не удалось создать новый набор")
            return

        new_to_id = new_from_id + 1
        self.refresh()
        for r in range(self.sets_table.rowCount()):
            it = self.sets_table.item(r, 2)
            if it and it.text().strip() == str(new_to_id):
                self.sets_table.selectRow(r)
                break

    def _copy_pair(self) -> None:
        if not self._current_to_set_id:
            return
        base = self.name_edit.text().strip()
        default_name = f"{base} (копия)" if base else "Новый набор (копия)"
        name, ok = QInputDialog.getText(self, "Копировать пару", "Название нового набора:", text=default_name)
        if not ok:
            return
        new_to_id = self.catalog_loader.clone_replacement_pair(self._current_to_set_id, name.strip() or None)
        if not new_to_id:
            QMessageBox.warning(self, "Ошибка", "Не удалось создать копию пары")
            return
        self.refresh()
        # пытаемся выбрать новый
        for r in range(self.sets_table.rowCount()):
            it = self.sets_table.item(r, 2)
            if it and it.text().strip() == str(new_to_id):
                self.sets_table.selectRow(r)
                break

    def _delete_pair(self) -> None:
        if not self._current_to_set_id:
            return
        reply = QMessageBox.question(
            self,
            "Подтверждение удаления",
            "Удалить выбранную пару наборов (from+to)?\nЭто действие нельзя отменить.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            return
        ok = self.catalog_loader.delete_replacement_pair(self._current_to_set_id)
        if not ok:
            QMessageBox.warning(self, "Ошибка", "Не удалось удалить пару")
            return
        self.refresh()

    def _save_current(self, silent: bool = False) -> bool:
        """Сохранить текущий набор материалов.
        
        Args:
            silent: Если True, не показывать сообщения об успехе (только об ошибках)
            
        Returns:
            True если сохранение успешно, False в противном случае
        """
        if not self._current_to_set_id or not self._loaded_to_set_id:
            return False

        # 1) Сохраняем имя пары всегда (если меняли)
        name_text = self.name_edit.text().strip()
        name = name_text or None

        if self._dirty_name:
            if not name and not silent:
                QMessageBox.information(self, "Подсказка", "Название очищено — набор будет без имени.")
            self.catalog_loader.update_set_name_for_pair(self._current_to_set_id, name)

        # 2) Если материалы не меняли — на этом всё (переименование не должно требовать валидации материалов)
        if not self._dirty_materials:
            if not silent:
                QMessageBox.information(self, "Успех", "Изменения сохранены")
            self.refresh()
            return True

        # 3) Валидация материалов (только если редактировали материалы)
        if not self._from_materials:
            QMessageBox.warning(self, "Ошибка", "Добавьте хотя бы один материал в набор 'до'")
            return False
        if not self._to_materials:
            QMessageBox.warning(self, "Ошибка", "Добавьте хотя бы один материал в набор 'после'")
            return False

        for i, m in enumerate(self._from_materials, start=1):
            if not (m.before_name or "").strip() or not (m.unit or "").strip():
                QMessageBox.warning(self, "Ошибка", f"В наборе 'до' материал #{i}: заполните Наименование и Ед. изм.")
                return False
            if (m.norm or 0) <= 0:
                QMessageBox.warning(self, "Ошибка", f"В наборе 'до' материал #{i}: Норма должна быть больше 0")
                return False

        for i, m in enumerate(self._to_materials, start=1):
            if not (m.before_name or "").strip() or not (m.unit or "").strip():
                QMessageBox.warning(self, "Ошибка", f"В наборе 'после' материал #{i}: заполните Наименование и Ед. изм.")
                return False

        # 4) Сохраняем материалы
        ok = True
        if self._loaded_from_set_id:
            ok = ok and self.catalog_loader.update_replacement_set(self._loaded_from_set_id, self._from_materials)
        ok = ok and self.catalog_loader.update_replacement_set(self._loaded_to_set_id, self._to_materials)

        if not ok:
            QMessageBox.warning(self, "Ошибка", "Не удалось сохранить изменения")
            return False
        if not silent:
            QMessageBox.information(self, "Успех", "Набор сохранён")
        self.refresh()
        return True

    def _save_and_add_to_document(self) -> None:
        """Сохранить набор материалов и сразу добавить деталь в документ"""
        # Сохраняем набор (без показа сообщения об успехе, так как покажем общее сообщение)
        if not self._save_current(silent=True):
            # Если сохранение не удалось, сообщение об ошибке уже показано в _save_current
            return

        # После успешного сохранения добавляем деталь в документ
        # Используем существующий метод _add_part_to_document, но без показа его сообщения об успехе
        # Вместо этого покажем общее сообщение
        if not self.document_data:
            QMessageBox.warning(self, "Ошибка", "Документ не инициализирован")
            return

        part = self._current_part
        if not part:
            QMessageBox.warning(self, "Ошибка", "Выберите деталь")
            return

        # Проверяем, не добавлена ли уже эта деталь
        for part_change in self.document_data.part_changes:
            if part_change.part == part:
                materials_count = len(part_change.materials)
                QMessageBox.information(
                    self,
                    "Информация",
                    f"Набор сохранён, но деталь '{part}' уже добавлена в документ ({materials_count} материалов)",
                )
                return

        # Вызываем логику добавления из _add_part_to_document, но без показа сообщения
        # Создаем временный флаг для подавления сообщения
        replacement_sets = self.catalog_loader.get_replacement_sets_by_part(part)

        if replacement_sets:
            # Есть наборы замены - используем их
            from_sets = [s for s in replacement_sets if s.set_type == "from"]
            to_sets = [s for s in replacement_sets if s.set_type == "to"]

            # Выбираем набор "to"
            selected_to_set: Optional[MaterialReplacementSet] = None
            if len(to_sets) == 1:
                # Один набор - используем автоматически
                selected_to_set = to_sets[0]
            elif len(to_sets) > 1:
                # Несколько наборов - показываем диалог выбора
                dialog = ReplacementSetSelectionDialog(
                    to_sets, part, self, catalog_loader=self.catalog_loader
                )
                if dialog.exec():
                    selected_to_set = dialog.get_selected_set()
                else:
                    # Пользователь отменил выбор
                    return

            # Выбираем ОДИН набор "from", соответствующий выбранному "to"
            selected_from_set = self._pick_matching_from_set(from_sets, selected_to_set)
            if selected_from_set is None:
                QMessageBox.warning(
                    self, "Ошибка", f"Для детали '{part}' не найден набор материалов 'до'"
                )
                return

            from_materials = self._deduplicate_materials(selected_from_set.materials or [])

            # Создаём изменения для детали с наборами
            additional_page = None
            if (
                self.get_current_additional_page
                and callable(self.get_current_additional_page)
            ):
                additional_page = self.get_current_additional_page()
            part_changes = PartChanges(part=part, additional_page_number=additional_page)

            to_materials = selected_to_set.materials if selected_to_set else []
            added_materials_keys = set()

            for idx, from_entry in enumerate(from_materials):
                material_key = (from_entry.workshop, from_entry.role, from_entry.before_name)
                if material_key in added_materials_keys:
                    continue

                added_materials_keys.add(material_key)
                material_change = MaterialChange(catalog_entry=from_entry, is_changed=True)

                if idx < len(to_materials):
                    to_entry = to_materials[idx]
                    material_change.after_name = to_entry.before_name
                    if to_entry.unit != from_entry.unit:
                        material_change.after_unit = to_entry.unit
                    if to_entry.norm > 0 and to_entry.norm != from_entry.norm:
                        material_change.after_norm = to_entry.norm

                part_changes.materials.append(material_change)

            if len(to_materials) > len(from_materials):
                if from_materials:
                    template_entry = from_materials[-1]
                else:
                    template_entry = CatalogEntry(
                        part=part,
                        workshop="",
                        role="",
                        before_name="",
                        unit="",
                        norm=0.0,
                        comment="",
                    )

                for idx in range(len(from_materials), len(to_materials)):
                    to_entry = to_materials[idx]
                    dummy_entry = CatalogEntry(
                        part=template_entry.part,
                        workshop=template_entry.workshop,
                        role=template_entry.role,
                        before_name="",
                        unit=template_entry.unit,
                        norm=0.0,
                        comment="",
                    )
                    material_change = MaterialChange(
                        catalog_entry=dummy_entry,
                        is_changed=False,
                        after_name=to_entry.before_name,
                        after_unit=to_entry.unit if to_entry.unit else template_entry.unit,
                        after_norm=to_entry.norm if to_entry.norm > 0 else None,
                    )
                    part_changes.materials.append(material_change)

            self.document_data.part_changes.append(part_changes)
        else:
            # Нет наборов - используем старую логику
            entries = self.catalog_loader.get_entries_by_part(part)
            if not entries:
                QMessageBox.warning(self, "Ошибка", f"Деталь '{part}' не найдена в справочнике")
                return

            entries = [e for e in entries if not e.is_part_of_set]
            entries = self._deduplicate_materials(entries)

            additional_page = None
            if (
                self.get_current_additional_page
                and callable(self.get_current_additional_page)
            ):
                additional_page = self.get_current_additional_page()
            part_changes = PartChanges(part=part, additional_page_number=additional_page)
            for entry in entries:
                material_change = MaterialChange(catalog_entry=entry)
                part_changes.materials.append(material_change)

            self.document_data.part_changes.append(part_changes)

        # Автоматически привязываем деталь к изделию, если изделие выбрано
        self._link_part_to_product(part)

        # Обновляем таблицу во второй вкладке
        if self.changes_widget:
            self.changes_widget.refresh()

        QMessageBox.information(
            self, "Успех", f"Набор сохранён и деталь '{part}' добавлена в документ"
        )

    def _split_copy(self) -> None:
        if not self._current_to_set_id:
            return

        from_rows = sorted({idx.row() for idx in self.from_table.selectionModel().selectedRows()})
        to_rows = sorted({idx.row() for idx in self.to_table.selectionModel().selectedRows()})
        if not from_rows and not to_rows:
            QMessageBox.warning(self, "Ошибка", "Выделите строки в 'до' и/или 'после' для вынесения")
            return

        base = self.name_edit.text().strip()
        default_name = f"{base} (часть)" if base else "Новый набор (часть)"
        name, ok = QInputDialog.getText(self, "Вынести в новый набор", "Название нового набора:", text=default_name)
        if not ok:
            return

        new_to_id = self.catalog_loader.split_replacement_pair_copy(
            self._current_to_set_id,
            from_rows,
            to_rows,
            name.strip() or None,
        )
        if not new_to_id:
            QMessageBox.warning(self, "Ошибка", "Не удалось создать новый набор из выделенного")
            return
        self.refresh()
        for r in range(self.sets_table.rowCount()):
            it = self.sets_table.item(r, 2)
            if it and it.text().strip() == str(new_to_id):
                self.sets_table.selectRow(r)
                break

    def _deduplicate_materials(self, materials: list[CatalogEntry]) -> list[CatalogEntry]:
        """Дедупликация материалов по уникальному ключу (workshop, role, before_name, unit)"""
        seen = set()
        unique_materials = []
        for material in materials:
            # Создаем уникальный ключ из workshop, role, before_name и unit
            # Материалы с одинаковым названием, но разными единицами измерения считаются разными
            key = (material.workshop, material.role, material.before_name, material.unit)
            if key not in seen:
                seen.add(key)
                unique_materials.append(material)
        return unique_materials

    def _pick_matching_from_set(
        self,
        from_sets: list[MaterialReplacementSet],
        selected_to_set: Optional[MaterialReplacementSet],
    ) -> Optional[MaterialReplacementSet]:
        """
        Подобрать один набор 'from', соответствующий выбранному набору 'to'.
        """
        if not from_sets:
            return None
        if selected_to_set is None:
            # Берём самый свежий (список уже отсортирован по created_at desc)
            return from_sets[0]

        to_name = selected_to_set.set_name or ""

        # 1) Пытаемся сопоставить по set_name
        same_name = [s for s in from_sets if (s.set_name or "") == to_name]
        if same_name:
            return same_name[0]

        # 2) Пытаемся сопоставить по парному ID (при создании from_set_id, затем to_set_id)
        if selected_to_set.id is not None:
            id_match = [s for s in from_sets if s.id == (selected_to_set.id - 1)]
            if id_match:
                return id_match[0]

            # 3) Фоллбек: ближайший по ID
            def _dist(s: MaterialReplacementSet) -> int:
                return abs((s.id or 0) - selected_to_set.id)  # type: ignore[arg-type]

            return sorted(from_sets, key=_dist)[0]

        return from_sets[0]

    def _link_part_to_product(self, part: str) -> None:
        """Привязать деталь к выбранным изделиям"""
        if not self.document_data or not self.product_store:
            return
        selected_products = self.document_data.products
        if selected_products:
            for product_name in selected_products:
                self.product_store.link_part_to_product_by_name(product_name, part)

    def _add_part_to_document(self) -> None:
        """Добавить текущую деталь в документ"""
        if not self.document_data:
            QMessageBox.warning(self, "Ошибка", "Документ не инициализирован")
            return

        part = self._current_part
        if not part:
            QMessageBox.warning(self, "Ошибка", "Выберите деталь")
            return

        # Проверяем, не добавлена ли уже эта деталь
        for part_change in self.document_data.part_changes:
            if part_change.part == part:
                materials_count = len(part_change.materials)
                QMessageBox.information(
                    self,
                    "Информация",
                    f"Деталь '{part}' уже добавлена ({materials_count} материалов)",
                )
                return

        # Проверяем наличие наборов замены для детали
        replacement_sets = self.catalog_loader.get_replacement_sets_by_part(part)

        if replacement_sets:
            # Есть наборы замены - используем их
            from_sets = [s for s in replacement_sets if s.set_type == "from"]
            to_sets = [s for s in replacement_sets if s.set_type == "to"]

            # Выбираем набор "to"
            selected_to_set: Optional[MaterialReplacementSet] = None
            if len(to_sets) == 1:
                # Один набор - используем автоматически
                selected_to_set = to_sets[0]
            elif len(to_sets) > 1:
                # Несколько наборов - показываем диалог выбора
                dialog = ReplacementSetSelectionDialog(
                    to_sets, part, self, catalog_loader=self.catalog_loader
                )
                if dialog.exec():
                    selected_to_set = dialog.get_selected_set()
                else:
                    # Пользователь отменил выбор
                    return

            # Выбираем ОДИН набор "from", соответствующий выбранному "to"
            selected_from_set = self._pick_matching_from_set(from_sets, selected_to_set)
            if selected_from_set is None:
                QMessageBox.warning(
                    self, "Ошибка", f"Для детали '{part}' не найден набор материалов 'до'"
                )
                return

            from_materials = self._deduplicate_materials(selected_from_set.materials or [])

            # Создаём изменения для детали с наборами
            # Получаем текущий номер доп. страницы, если установлен
            additional_page = None
            if (
                self.get_current_additional_page
                and callable(self.get_current_additional_page)
            ):
                additional_page = self.get_current_additional_page()
                logger.info(f"[_add_part_to_document] Деталь '{part}': получен номер страницы из get_current_additional_page() = {additional_page}")
            else:
                logger.info(f"[_add_part_to_document] Деталь '{part}': get_current_additional_page не доступен, additional_page = None")
            
            # ВАЖНО: Если номер страницы установлен, убеждаемся, что он больше максимального существующего
            # Это гарантирует, что деталь всегда попадает на новую страницу
            if additional_page is not None:
                max_existing_page = 0
                existing_pages = []
                for existing_part_change in self.document_data.part_changes:
                    if existing_part_change.additional_page_number is not None:
                        max_existing_page = max(max_existing_page, existing_part_change.additional_page_number)
                        existing_pages.append((existing_part_change.part, existing_part_change.additional_page_number))
                logger.info(f"[_add_part_to_document] Деталь '{part}': максимальный существующий номер страницы = {max_existing_page}")
                logger.info(f"[_add_part_to_document] Деталь '{part}': существующие детали по страницам: {existing_pages}")
                
                # Если установленный номер страницы не больше максимального, увеличиваем его
                if additional_page <= max_existing_page:
                    old_page = additional_page
                    additional_page = max_existing_page + 1
                    logger.warning(f"[_add_part_to_document] Деталь '{part}': номер страницы увеличен с {old_page} до {additional_page} (был <= максимального {max_existing_page})")
                else:
                    logger.info(f"[_add_part_to_document] Деталь '{part}': номер страницы {additional_page} корректен (больше максимального {max_existing_page})")
            
            logger.info(f"[_add_part_to_document] Деталь '{part}': финальный номер страницы = {additional_page}")
            part_changes = PartChanges(part=part, additional_page_number=additional_page)

            # Получаем материалы из набора "to" для сопоставления
            to_materials = selected_to_set.materials if selected_to_set else []

            # Добавляем материалы из набора "from" в левую колонку
            # и сопоставляем с материалами из набора "to" по порядку
            # Используем set для отслеживания уже добавленных материалов (по ключу workshop+role+before_name)
            added_materials_keys = set()

            # Сначала добавляем все материалы "from" - все они должны иметь is_changed=True
            # чтобы попасть в левую колонку документа
            for idx, from_entry in enumerate(from_materials):
                # Проверяем, не добавлен ли уже такой материал
                material_key = (from_entry.workshop, from_entry.role, from_entry.before_name)
                if material_key in added_materials_keys:
                    continue

                added_materials_keys.add(material_key)

                # Создаём изменение для материала "from"
                # ВСЕ материалы "from" должны иметь is_changed=True, чтобы попасть в левую колонку
                material_change = MaterialChange(catalog_entry=from_entry, is_changed=True)

                # Если есть соответствующий материал "to" по индексу, заполняем колонку "после"
                if idx < len(to_materials):
                    to_entry = to_materials[idx]
                    material_change.after_name = to_entry.before_name
                    # Также можно обновить единицу измерения и норму, если они отличаются
                    if to_entry.unit != from_entry.unit:
                        material_change.after_unit = to_entry.unit
                    # Устанавливаем норму только если она указана (больше 0)
                    if to_entry.norm > 0 and to_entry.norm != from_entry.norm:
                        material_change.after_norm = to_entry.norm
                    # Если норма не указана (0), after_norm остается None (пустое значение)

                part_changes.materials.append(material_change)

            # Теперь добавляем лишние материалы "to", которых больше чем "from"
            # Эти материалы должны попасть только в правую колонку (after_name заполнен, is_changed=False)
            if len(to_materials) > len(from_materials):
                # Используем последний from_entry как шаблон для создания фиктивного catalog_entry
                # или создаём фиктивный entry, если from_materials пуст
                if from_materials:
                    template_entry = from_materials[-1]
                else:
                    # Создаём фиктивный entry, если нет материалов "from"
                    template_entry = CatalogEntry(
                        part=part,
                        workshop="",
                        role="",
                        before_name="",
                        unit="",
                        norm=0.0,
                        comment="",
                    )

                # Добавляем лишние материалы "to"
                for idx in range(len(from_materials), len(to_materials)):
                    to_entry = to_materials[idx]
                    # Создаём фиктивный catalog_entry на основе шаблона
                    dummy_entry = CatalogEntry(
                        part=template_entry.part,
                        workshop=template_entry.workshop,
                        role=template_entry.role,
                        before_name="",  # Пустое, так как это только для правой колонки
                        unit=template_entry.unit,
                        norm=0.0,
                        comment="",
                    )
                    # Создаём MaterialChange только для правой колонки (is_changed=False, но after_name заполнен)
                    material_change = MaterialChange(
                        catalog_entry=dummy_entry,
                        is_changed=False,  # Не попадает в левую колонку
                        after_name=to_entry.before_name,  # Попадает в правую колонку
                        after_unit=to_entry.unit if to_entry.unit else template_entry.unit,
                        after_norm=to_entry.norm if to_entry.norm > 0 else None,
                    )
                    part_changes.materials.append(material_change)

            self.document_data.part_changes.append(part_changes)
        else:
            # Нет наборов - используем старую логику
            entries = self.catalog_loader.get_entries_by_part(part)
            if not entries:
                QMessageBox.warning(self, "Ошибка", f"Деталь '{part}' не найдена в справочнике")
                return

            # Фильтруем материалы, которые входят в наборы (чтобы избежать дублирования)
            # Материалы из наборов должны использоваться только через наборы
            entries = [e for e in entries if not e.is_part_of_set]

            # Дедупликация записей из каталога (на случай дубликатов в БД)
            entries = self._deduplicate_materials(entries)

            # Создаём изменения для детали
            # Получаем текущий номер доп. страницы, если установлен
            additional_page = None
            if (
                self.get_current_additional_page
                and callable(self.get_current_additional_page)
            ):
                additional_page = self.get_current_additional_page()
            part_changes = PartChanges(part=part, additional_page_number=additional_page)
            for entry in entries:
                material_change = MaterialChange(catalog_entry=entry)
                part_changes.materials.append(material_change)

            self.document_data.part_changes.append(part_changes)

        # Автоматически привязываем деталь к изделию, если изделие выбрано
        self._link_part_to_product(part)

        # Обновляем таблицу во второй вкладке
        if self.changes_widget:
            self.changes_widget.refresh()

        QMessageBox.information(
            self, "Успех", f"Деталь '{part}' успешно добавлена в документ"
        )


