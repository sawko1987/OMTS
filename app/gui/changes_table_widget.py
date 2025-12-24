"""
Виджет для редактирования таблицы изменений материалов
"""
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QComboBox,
    QTableWidget, QTableWidgetItem, QCheckBox, QHeaderView,
    QMessageBox, QLabel, QLineEdit
)
from PySide6.QtCore import Qt
from typing import List, Optional

from app.models import DocumentData, PartChanges, MaterialChange, CatalogEntry
from app.catalog_loader import CatalogLoader
from app.history_store import HistoryStore
from app.product_store import ProductStore
from app.gui.part_creation_dialog import PartCreationDialog


class ChangesTableWidget(QWidget):
    """Виджет таблицы изменений"""
    
    def __init__(self, document_data: DocumentData, 
                 catalog_loader: CatalogLoader = None,
                 history_store: HistoryStore = None,
                 product_store: ProductStore = None):
        super().__init__()
        self.document_data = document_data
        self.catalog_loader = catalog_loader or CatalogLoader()
        self.history_store = history_store or HistoryStore()
        self.product_store = product_store
        self.init_ui()
    
    def init_ui(self):
        """Инициализация интерфейса"""
        layout = QVBoxLayout(self)
        
        # Заголовок
        title = QLabel("Изменения материалов по деталям")
        title.setStyleSheet("font-size: 16px; font-weight: bold;")
        layout.addWidget(title)
        
        # Информация о выбранном изделии
        self.product_info_label = QLabel()
        self.product_info_label.setStyleSheet("color: #666; font-style: italic;")
        layout.addWidget(self.product_info_label)
        
        # Панель управления
        control_layout = QVBoxLayout()
        
        # Поиск деталей
        search_layout = QHBoxLayout()
        search_layout.addWidget(QLabel("Поиск детали:"))
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Введите код детали для поиска...")
        self.search_edit.textChanged.connect(self.on_search_changed)
        search_layout.addWidget(self.search_edit)
        control_layout.addLayout(search_layout)
        
        # Выбор и добавление детали
        add_layout = QHBoxLayout()
        add_layout.addWidget(QLabel("Добавить деталь:"))
        self.part_combo = QComboBox()
        self.part_combo.setEditable(True)
        self.part_combo.setMinimumWidth(200)
        self.part_combo.currentTextChanged.connect(self.on_part_combo_changed)
        add_layout.addWidget(self.part_combo)
        
        self.btn_add_part = QPushButton("Добавить")
        self.btn_add_part.clicked.connect(self.add_part)
        add_layout.addWidget(self.btn_add_part)
        
        self.btn_create_part = QPushButton("Создать новую деталь")
        self.btn_create_part.clicked.connect(self.create_part)
        add_layout.addWidget(self.btn_create_part)
        
        add_layout.addStretch()
        
        self.btn_remove_part = QPushButton("Удалить деталь")
        self.btn_remove_part.clicked.connect(self.remove_selected_part)
        add_layout.addWidget(self.btn_remove_part)
        
        control_layout.addLayout(add_layout)
        layout.addLayout(control_layout)
        
        # Таблица
        self.table = QTableWidget()
        self.table.setColumnCount(8)
        self.table.setHorizontalHeaderLabels([
            "Деталь",
            "Цех",
            "Тип",
            "Материал 'до'",
            "Ед.изм.",
            "Норма",
            "Меняем",
            "Материал 'после'"
        ])
        
        # Настройка таблицы
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.Stretch)
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(6, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(7, QHeaderView.Stretch)
        
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.itemChanged.connect(self.on_item_changed)
        
        layout.addWidget(self.table)
    
    def update_product_filter(self):
        """Обновить фильтр деталей по выбранному изделию"""
        self.load_parts_list()
    
    def load_parts_list(self):
        """Загрузить список деталей с учетом фильтра по изделию"""
        try:
            product_name = self.document_data.product.strip()
            
            if product_name and self.product_store:
                # Загружаем только детали, привязанные к изделию
                parts = self.product_store.get_parts_by_product_name(product_name)
                self.product_info_label.setText(
                    f"Показаны детали для изделия: {product_name} ({len(parts)} деталей)"
                )
            else:
                # Загружаем все детали
                parts = self.catalog_loader.get_all_parts()
                if product_name:
                    self.product_info_label.setText(
                        f"Изделие '{product_name}' не найдено. Показаны все детали."
                    )
                else:
                    self.product_info_label.setText("Выберите изделие на первой вкладке")
            
            # Обновляем комбобокс
            current_text = self.part_combo.currentText()
            self.part_combo.clear()
            self.part_combo.addItems(parts)
            
            # Восстанавливаем текст поиска, если был
            if current_text:
                index = self.part_combo.findText(current_text)
                if index >= 0:
                    self.part_combo.setCurrentIndex(index)
                else:
                    self.part_combo.setEditText(current_text)
            
        except Exception as e:
            QMessageBox.warning(self, "Ошибка", f"Не удалось загрузить список деталей:\n{e}")
    
    def on_search_changed(self, text: str):
        """Обработка изменения текста поиска"""
        if not text.strip():
            self.load_parts_list()
            return
        
        # Выполняем поиск
        try:
            product_name = self.document_data.product.strip()
            
            if product_name and self.product_store:
                # Ищем среди деталей изделия
                all_parts = self.product_store.get_parts_by_product_name(product_name)
                # Фильтруем по поисковому запросу
                search_results = self.catalog_loader.search_parts(text)
                # Пересечение: только те, что есть и в изделии, и в результатах поиска
                parts = [p for p in search_results if p in all_parts]
            else:
                # Ищем среди всех деталей
                parts = self.catalog_loader.search_parts(text)
            
            # Обновляем комбобокс
            current_text = self.part_combo.currentText()
            self.part_combo.clear()
            self.part_combo.addItems(parts)
            
            if current_text and current_text in parts:
                index = self.part_combo.findText(current_text)
                if index >= 0:
                    self.part_combo.setCurrentIndex(index)
            
        except Exception as e:
            QMessageBox.warning(self, "Ошибка", f"Ошибка поиска:\n{e}")
    
    def on_part_combo_changed(self, text: str):
        """Обработка изменения выбора детали в комбобоксе"""
        pass
    
    def load_catalog(self):
        """Загрузить каталог и обновить список деталей"""
        try:
            self.catalog_loader.load()
            self.load_parts_list()
        except Exception as e:
            QMessageBox.warning(self, "Ошибка", f"Не удалось загрузить справочник:\n{e}")
    
    def create_part(self):
        """Создать новую деталь"""
        dialog = PartCreationDialog(self.catalog_loader, self)
        if dialog.exec():
            # Проверяем режим работы
            if not dialog.is_replacement_set_mode:
                # Одиночный материал
                entry = dialog.get_entry()
                if entry:
                    # Сохраняем в БД
                    entry_id = self.catalog_loader.add_entry(entry)
                    if entry_id:
                        QMessageBox.information(self, "Успех", f"Деталь '{entry.part}' успешно создана")
                        
                        # Предлагаем привязать к текущему изделию
                        product_name = self.document_data.product.strip()
                        if product_name and self.product_store:
                            reply = QMessageBox.question(
                                self, "Привязать к изделию?",
                                f"Привязать деталь '{entry.part}' к изделию '{product_name}'?",
                                QMessageBox.Yes | QMessageBox.No
                            )
                            if reply == QMessageBox.Yes:
                                self.product_store.link_part_to_product_by_name(product_name, entry.part)
                        
                        # Обновляем список
                        self.load_parts_list()
                        # Устанавливаем новую деталь в комбобоксе
                        index = self.part_combo.findText(entry.part)
                        if index >= 0:
                            self.part_combo.setCurrentIndex(index)
                    else:
                        QMessageBox.warning(self, "Ошибка", "Не удалось сохранить деталь")
            else:
                # Набор материалов
                part_code, from_materials, to_materials = dialog.get_replacement_set_data()
                set_id = self.catalog_loader.add_replacement_set(part_code, from_materials, to_materials)
                if set_id:
                    QMessageBox.information(self, "Успех", 
                                          f"Набор материалов для детали '{part_code}' успешно создан")
                    
                    # Предлагаем привязать к текущему изделию
                    product_name = self.document_data.product.strip()
                    if product_name and self.product_store:
                        reply = QMessageBox.question(
                            self, "Привязать к изделию?",
                            f"Привязать деталь '{part_code}' к изделию '{product_name}'?",
                            QMessageBox.Yes | QMessageBox.No
                        )
                        if reply == QMessageBox.Yes:
                            self.product_store.link_part_to_product_by_name(product_name, part_code)
                    
                    # Обновляем список
                    self.load_parts_list()
                    # Устанавливаем новую деталь в комбобоксе
                    index = self.part_combo.findText(part_code)
                    if index >= 0:
                        self.part_combo.setCurrentIndex(index)
                else:
                    QMessageBox.warning(self, "Ошибка", "Не удалось сохранить набор материалов")(self, "Ошибка", "Не удалось сохранить деталь")
    
    def add_part(self):
        """Добавить деталь в таблицу"""
        part = self.part_combo.currentText().strip()
        if not part:
            QMessageBox.warning(self, "Ошибка", "Укажите деталь")
            return
        
        # Проверяем, не добавлена ли уже эта деталь
        for part_change in self.document_data.part_changes:
            if part_change.part == part:
                QMessageBox.information(self, "Информация", f"Деталь '{part}' уже добавлена")
                return
        
        # Загружаем записи из каталога
        entries = self.catalog_loader.get_entries_by_part(part)
        if not entries:
            QMessageBox.warning(self, "Ошибка", f"Деталь '{part}' не найдена в справочнике")
            return
        
        # Создаём изменения для детали
        part_changes = PartChanges(part=part)
        for entry in entries:
            material_change = MaterialChange(catalog_entry=entry)
            part_changes.materials.append(material_change)
        
        self.document_data.part_changes.append(part_changes)
        
        # Автоматически привязываем деталь к изделию, если изделие выбрано
        product_name = self.document_data.product.strip()
        if product_name and self.product_store:
            self.product_store.link_part_to_product_by_name(product_name, part)
        
        self.refresh()
    
    def remove_selected_part(self):
        """Удалить выбранную деталь"""
        current_row = self.table.currentRow()
        if current_row < 0:
            QMessageBox.warning(self, "Ошибка", "Выберите строку для удаления")
            return
        
        # Находим деталь по строке
        part_item = self.table.item(current_row, 0)
        if not part_item:
            return
        
        part = part_item.text()
        
        # Удаляем из данных
        self.document_data.part_changes = [
            pc for pc in self.document_data.part_changes
            if pc.part != part
        ]
        
        self.refresh()
    
    def refresh(self):
        """Обновить таблицу"""
        self.table.setRowCount(0)
        
        row = 0
        for part_change in self.document_data.part_changes:
            for material in part_change.materials:
                self.table.insertRow(row)
                
                # Деталь
                part_item = QTableWidgetItem(part_change.part)
                part_item.setFlags(part_item.flags() & ~Qt.ItemIsEditable)
                self.table.setItem(row, 0, part_item)
                
                # Цех
                workshop_item = QTableWidgetItem(material.catalog_entry.workshop)
                workshop_item.setFlags(workshop_item.flags() & ~Qt.ItemIsEditable)
                self.table.setItem(row, 1, workshop_item)
                
                # Тип позиции
                role_item = QTableWidgetItem(material.catalog_entry.role)
                role_item.setFlags(role_item.flags() & ~Qt.ItemIsEditable)
                self.table.setItem(row, 2, role_item)
                
                # Материал 'до'
                before_item = QTableWidgetItem(material.catalog_entry.before_name)
                before_item.setFlags(before_item.flags() & ~Qt.ItemIsEditable)
                self.table.setItem(row, 3, before_item)
                
                # Ед. изм.
                unit_item = QTableWidgetItem(material.catalog_entry.unit)
                unit_item.setFlags(unit_item.flags() & ~Qt.ItemIsEditable)
                self.table.setItem(row, 4, unit_item)
                
                # Норма
                norm_item = QTableWidgetItem(str(material.catalog_entry.norm))
                norm_item.setFlags(norm_item.flags() & ~Qt.ItemIsEditable)
                self.table.setItem(row, 5, norm_item)
                
                # Чекбокс "Меняем"
                checkbox = QCheckBox()
                checkbox.setChecked(material.is_changed)
                checkbox.stateChanged.connect(
                    lambda state, r=row: self.on_checkbox_changed(r, state == Qt.Checked)
                )
                self.table.setCellWidget(row, 6, checkbox)
                
                # Материал 'после' (с автодополнением)
                # Проверяем, есть ли предложения из истории
                suggestions = self.history_store.get_suggestions(material.catalog_entry)
                
                if suggestions:
                    # Используем комбобокс
                    combo = QComboBox()
                    combo.setEditable(True)
                    combo.addItems(suggestions)
                    combo.setCurrentText(material.after_name or "")
                    
                    def on_text_changed(text):
                        material.after_name = text.strip() or None
                        if material.after_name:
                            self.history_store.add_replacement(material.catalog_entry, material.after_name)
                    
                    combo.currentTextChanged.connect(on_text_changed)
                    self.table.setCellWidget(row, 7, combo)
                else:
                    # Используем обычную ячейку
                    after_item = QTableWidgetItem(material.after_name or "")
                    after_item.setData(Qt.UserRole, material)
                    self.table.setItem(row, 7, after_item)
                
                row += 1
    
    def on_checkbox_changed(self, row: int, checked: bool):
        """Обработка изменения чекбокса"""
        part_item = self.table.item(row, 0)
        if not part_item:
            return
        
        part = part_item.text()
        
        # Находим соответствующий material
        for part_change in self.document_data.part_changes:
            if part_change.part == part:
                # Находим материал по строке (нужно сопоставить по данным)
                before_item = self.table.item(row, 3)
                if before_item:
                    before_name = before_item.text()
                    for material in part_change.materials:
                        if material.catalog_entry.before_name == before_name:
                            material.is_changed = checked
                            break
                break
    
    def on_item_changed(self, item: QTableWidgetItem):
        """Обработка изменения ячейки"""
        if item.column() == 7:  # Материал 'после'
            material = item.data(Qt.UserRole)
            if material:
                material.after_name = item.text().strip() or None
                
                # Сохраняем в историю, если материал указан
                if material.after_name:
                    self.history_store.add_replacement(
                        material.catalog_entry,
                        material.after_name
                    )
    
    def update_document_data(self):
        """Обновить данные документа из таблицы"""
        # Собираем данные из всех виджетов таблицы
        for row in range(self.table.rowCount()):
            # Получаем чекбокс
            checkbox_widget = self.table.cellWidget(row, 6)
            if isinstance(checkbox_widget, QCheckBox):
                is_checked = checkbox_widget.isChecked()
            else:
                is_checked = False
            
            # Получаем материал 'после'
            after_widget = self.table.cellWidget(row, 7)
            after_name = None
            
            if isinstance(after_widget, QComboBox):
                # Если это комбобокс
                after_name = after_widget.currentText().strip() or None
            else:
                # Если это обычная ячейка
                after_item = self.table.item(row, 7)
                if after_item:
                    after_name = after_item.text().strip() or None
            
            # Обновляем данные материала
            part_item = self.table.item(row, 0)
            before_item = self.table.item(row, 3)
            
            if part_item and before_item:
                part = part_item.text()
                before_name = before_item.text()
                
                # Находим соответствующий material
                for part_change in self.document_data.part_changes:
                    if part_change.part == part:
                        for material in part_change.materials:
                            if material.catalog_entry.before_name == before_name:
                                material.is_changed = is_checked
                                material.after_name = after_name
                                
                                # Сохраняем в историю, если материал указан
                                if after_name:
                                    self.history_store.add_replacement(
                                        material.catalog_entry,
                                        after_name
                                    )
                                break
                        break
