"""
Виджет для редактирования таблицы изменений материалов
"""
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QComboBox,
    QTableWidget, QTableWidgetItem, QCheckBox, QHeaderView,
    QMessageBox, QLabel, QLineEdit, QCompleter
)
from PySide6.QtCore import Qt, QStringListModel, QTimer
from PySide6.QtGui import QColor
from typing import List, Optional
import logging

from app.models import DocumentData, PartChanges, MaterialChange, CatalogEntry, MaterialReplacementSet
from app.catalog_loader import CatalogLoader
from app.history_store import HistoryStore
from app.product_store import ProductStore
from app.gui.part_creation_dialog import PartCreationDialog
from app.gui.set_selection_dialog import ReplacementSetSelectionDialog

# Настройка логирования
logger = logging.getLogger(__name__)


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
        
        # Таймер для debounce поиска (задержка 300мс)
        self.search_timer = QTimer()
        self.search_timer.setSingleShot(True)
        self.search_timer.timeout.connect(self._perform_search)
        self._pending_search_text = ""
        
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
        
        # Настраиваем автодополнение для комбобокса
        self.part_completer_model = QStringListModel()
        self.part_completer = QCompleter(self.part_completer_model, self)
        self.part_completer.setCaseSensitivity(Qt.CaseInsensitive)
        self.part_completer.setCompletionMode(QCompleter.PopupCompletion)
        self.part_completer.setFilterMode(Qt.MatchContains)
        self.part_completer.setMaxVisibleItems(10)
        self.part_combo.setCompleter(self.part_completer)
        
        # Обработчик для фильтрации при вводе в комбобокс
        line_edit = self.part_combo.lineEdit()
        if line_edit:
            line_edit.textEdited.connect(self.on_part_combo_text_edited)
        
        # Сохраняем полный список деталей для фильтрации
        self._all_parts_list: List[str] = []
        
        add_layout.addWidget(self.part_combo)
        
        self.btn_add_part = QPushButton("Добавить")
        self.btn_add_part.clicked.connect(self.add_part)
        add_layout.addWidget(self.btn_add_part)
        
        self.btn_create_part = QPushButton("Создать новую деталь")
        self.btn_create_part.clicked.connect(self.create_part)
        add_layout.addWidget(self.btn_create_part)
        
        self.btn_copy_and_edit = QPushButton("Копировать и редактировать")
        self.btn_copy_and_edit.clicked.connect(self.copy_and_edit_part)
        add_layout.addWidget(self.btn_copy_and_edit)
        
        self.btn_edit_set = QPushButton("Редактировать набор материалов")
        self.btn_edit_set.clicked.connect(self.edit_replacement_set)
        add_layout.addWidget(self.btn_edit_set)
        
        add_layout.addStretch()
        
        self.btn_remove_part = QPushButton("Удалить деталь")
        self.btn_remove_part.clicked.connect(self.remove_selected_part)
        add_layout.addWidget(self.btn_remove_part)
        
        control_layout.addLayout(add_layout)
        layout.addLayout(control_layout)
        
        # Таблица
        self.table = QTableWidget()
        self.table.setColumnCount(9)
        self.table.setHorizontalHeaderLabels([
            "Деталь",
            "Цех",
            "Тип",
            "Материал 'до'",
            "Ед.изм.",
            "Норма",
            "Меняем",
            "Материал 'после'",
            "Страница"
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
        header.setSectionResizeMode(8, QHeaderView.ResizeToContents)
        
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.itemChanged.connect(self.on_item_changed)
        
        layout.addWidget(self.table)
        
        # Загружаем список деталей при инициализации
        self.load_parts_list()
    
    def update_product_filter(self):
        """Обновить фильтр деталей по выбранному изделию"""
        self.load_parts_list()
    
    def load_parts_list(self):
        """Загрузить список деталей с учетом фильтра по изделию"""
        try:
            product_name = self.document_data.product.strip()
            logger.debug(f"Загрузка списка деталей для изделия: '{product_name}'")
            
            if product_name and self.product_store:
                # Загружаем только детали, привязанные к изделию
                parts = self.product_store.get_parts_by_product_name(product_name)
                logger.info(f"Загружено {len(parts)} деталей для изделия '{product_name}'")
                self.product_info_label.setText(
                    f"Показаны детали для изделия: {product_name} ({len(parts)} деталей)"
                )
            else:
                # Загружаем все детали
                parts = self.catalog_loader.get_all_parts()
                logger.info(f"Загружено всех деталей: {len(parts)}")
                if product_name:
                    logger.warning(f"Изделие '{product_name}' не найдено, показаны все детали")
                    self.product_info_label.setText(
                        f"Изделие '{product_name}' не найдено. Показаны все детали."
                    )
                else:
                    self.product_info_label.setText("Выберите изделие на первой вкладке")
            
            # Сохраняем полный список для фильтрации
            self._all_parts_list = parts.copy() if parts else []
            
            # Обновляем модель автодополнения
            self.part_completer_model.setStringList(parts)
            
            # Обновляем комбобокс
            current_text = self.part_combo.currentText()
            self.part_combo.clear()
            self.part_combo.addItems(parts)
            logger.debug(f"Обновлен комбобокс с {len(parts)} деталями")
            
            # Восстанавливаем текст поиска, если был
            if current_text:
                index = self.part_combo.findText(current_text)
                if index >= 0:
                    self.part_combo.setCurrentIndex(index)
                else:
                    self.part_combo.setEditText(current_text)
            
        except Exception as e:
            logger.error(f"Ошибка при загрузке списка деталей: {e}", exc_info=True)
            QMessageBox.warning(self, "Ошибка", f"Не удалось загрузить список деталей:\n{e}")
    
    def on_search_changed(self, text: str):
        """Обработка изменения текста поиска (с debounce)"""
        search_text = text.strip()
        
        # Останавливаем предыдущий таймер
        self.search_timer.stop()
        
        if not search_text:
            logger.debug("Пустой текст поиска, загружаем полный список деталей")
            self.load_parts_list()
            return
        
        # Сохраняем текст для поиска и запускаем таймер
        self._pending_search_text = search_text
        # Запускаем таймер с задержкой 300мс для debounce
        self.search_timer.start(300)
    
    def _perform_search(self):
        """Выполнить поиск (вызывается после debounce)"""
        search_text = self._pending_search_text
        if not search_text:
            return
        
        logger.info(f"Выполнение поиска: запрос='{search_text}'")
        
        # Выполняем поиск при вводе от 1 символа
        if len(search_text) >= 1:
            try:
                product_name = self.document_data.product.strip()
                logger.debug(f"Поиск для изделия: '{product_name}'")
                
                if product_name and self.product_store:
                    # Ищем среди деталей изделия
                    all_parts = self.product_store.get_parts_by_product_name(product_name)
                    logger.info(f"Деталей в изделии '{product_name}': {len(all_parts)}")
                    
                    # Фильтруем по поисковому запросу
                    search_results = self.catalog_loader.search_parts(search_text)
                    logger.info(f"Результатов поиска по запросу '{search_text}': {len(search_results)}")
                    
                    # Пересечение: только те, что есть и в изделии, и в результатах поиска
                    parts = [p for p in search_results if p in all_parts]
                    logger.info(f"Результатов после фильтрации по изделию: {len(parts)}")
                    
                    if len(search_results) > 0 and len(parts) == 0:
                        logger.warning(f"Поиск нашел {len(search_results)} деталей, но ни одна не принадлежит изделию '{product_name}'")
                else:
                    # Ищем среди всех деталей
                    logger.debug("Поиск среди всех деталей (изделие не указано или не найдено)")
                    parts = self.catalog_loader.search_parts(search_text)
                    logger.info(f"Результатов поиска: {len(parts)}")
                
                # Сохраняем полный список для фильтрации
                self._all_parts_list = parts.copy() if parts else []
                
                # Обновляем модель автодополнения с результатами поиска
                self.part_completer_model.setStringList(parts)
                logger.debug(f"Обновлена модель автодополнения с {len(parts)} деталями")
                
                # Обновляем комбобокс, сохраняя текущий текст, если он есть
                current_text = self.part_combo.currentText()
                self.part_combo.blockSignals(True)  # Блокируем сигналы, чтобы не вызывать обработчики
                self.part_combo.clear()
                if parts:
                    self.part_combo.addItems(parts)
                    logger.debug(f"Обновлен комбобокс с {len(parts)} деталями")
                else:
                    logger.warning(f"Не найдено деталей по запросу '{search_text}'")
                
                # Восстанавливаем текст, если он был, но только если он не из поля поиска
                # Не устанавливаем текст из поля поиска в комбобокс, чтобы не мешать вводу
                self.part_combo.blockSignals(False)
                
            except Exception as e:
                logger.error(f"Ошибка при поиске деталей (запрос='{search_text}'): {e}", exc_info=True)
                QMessageBox.warning(self, "Ошибка", f"Ошибка поиска:\n{e}")
    
    def on_part_combo_changed(self, text: str):
        """Обработка изменения выбора детали в комбобоксе"""
        pass
    
    def on_part_combo_text_edited(self, text: str):
        """Обработка ввода текста в комбобокс - фильтрация вариантов"""
        search_text = text.strip()
        logger.debug(f"Редактирование текста в комбобоксе: '{search_text}'")
        
        # Если текст пустой, загружаем все детали
        if len(search_text) == 0:
            logger.debug("Пустой текст в комбобоксе, загружаем полный список")
            self.load_parts_list()
            return
        
        # Если есть сохраненный список, фильтруем его
        if self._all_parts_list:
            try:
                logger.debug(f"Фильтрация сохраненного списка ({len(self._all_parts_list)} элементов)")
                # Фильтруем список по введенному тексту
                filtered = [
                    p for p in self._all_parts_list 
                    if search_text.lower() in p.lower()
                ]
                logger.debug(f"Результатов фильтрации: {len(filtered)}")
                # Обновляем модель completer
                self.part_completer_model.setStringList(filtered)
            except Exception as e:
                logger.error(f"Ошибка при фильтрации списка деталей: {e}", exc_info=True)
        else:
            # Если списка нет, выполняем поиск через каталог
            try:
                logger.debug("Сохраненного списка нет, выполняем поиск через каталог")
                product_name = self.document_data.product.strip()
                
                if product_name and self.product_store:
                    all_parts = self.product_store.get_parts_by_product_name(product_name)
                    search_results = self.catalog_loader.search_parts(search_text)
                    parts = [p for p in search_results if p in all_parts]
                    logger.info(f"Поиск в комбобоксе: найдено {len(parts)} деталей для изделия '{product_name}'")
                else:
                    parts = self.catalog_loader.search_parts(search_text)
                    logger.info(f"Поиск в комбобоксе: найдено {len(parts)} деталей")
                
                if parts:
                    self._all_parts_list = parts.copy()
                    self.part_completer_model.setStringList(parts)
            except Exception as e:
                logger.error(f"Ошибка при поиске деталей в комбобоксе: {e}", exc_info=True)
    
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
                
                # Проверяем режим редактирования
                if dialog.is_edit_mode:
                    from_set_id, to_set_id = dialog.get_edit_set_ids()
                    success = False
                    
                    if from_set_id:
                        success = self.catalog_loader.update_replacement_set(from_set_id, from_materials)
                    if to_set_id and success:
                        success = self.catalog_loader.update_replacement_set(to_set_id, to_materials)
                    
                    if success:
                        QMessageBox.information(self, "Успех", 
                                              f"Набор материалов для детали '{part_code}' успешно обновлён")
                        self.load_parts_list()
                    else:
                        QMessageBox.warning(self, "Ошибка", "Не удалось обновить набор материалов")
                else:
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
                        QMessageBox.warning(self, "Ошибка", "Не удалось сохранить набор материалов")
    
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
        
        # Проверяем наличие наборов замены для детали
        replacement_sets = self.catalog_loader.get_replacement_sets_by_part(part)
        
        if replacement_sets:
            # Есть наборы замены - используем их
            from_sets = [s for s in replacement_sets if s.set_type == 'from']
            to_sets = [s for s in replacement_sets if s.set_type == 'to']
            
            # Выбираем набор "from" (обычно один, если несколько - объединяем материалы)
            from_materials = []
            for from_set in from_sets:
                from_materials.extend(from_set.materials)
            
            # Выбираем набор "to"
            selected_to_set: Optional[MaterialReplacementSet] = None
            if len(to_sets) == 1:
                # Один набор - используем автоматически
                selected_to_set = to_sets[0]
            elif len(to_sets) > 1:
                # Несколько наборов - показываем диалог выбора
                dialog = ReplacementSetSelectionDialog(to_sets, part, self)
                if dialog.exec():
                    selected_to_set = dialog.get_selected_set()
                else:
                    # Пользователь отменил выбор
                    return
            
            # Создаём изменения для детали с наборами
            # Получаем текущий номер доп. страницы, если установлен
            additional_page = None
            if hasattr(self, 'get_current_additional_page') and callable(self.get_current_additional_page):
                additional_page = self.get_current_additional_page()
            part_changes = PartChanges(part=part, additional_page_number=additional_page)
            
            # Получаем материалы из набора "to" для сопоставления
            to_materials = selected_to_set.materials if selected_to_set else []
            
            # Добавляем материалы из набора "from" в левую колонку
            # и сопоставляем с материалами из набора "to" по порядку
            for idx, from_entry in enumerate(from_materials):
                # Создаём изменение для материала "from"
                material_change = MaterialChange(catalog_entry=from_entry, is_changed=False)
                
                # Если есть соответствующий материал "to" по индексу, заполняем колонку "после"
                if idx < len(to_materials):
                    to_entry = to_materials[idx]
                    material_change.is_changed = True
                    material_change.after_name = to_entry.before_name
                    # Также можно обновить единицу измерения и норму, если они отличаются
                    if to_entry.unit != from_entry.unit:
                        material_change.after_unit = to_entry.unit
                    if to_entry.norm != from_entry.norm:
                        material_change.after_norm = to_entry.norm
                
                part_changes.materials.append(material_change)
            
            self.document_data.part_changes.append(part_changes)
        else:
            # Нет наборов - используем старую логику
            entries = self.catalog_loader.get_entries_by_part(part)
            if not entries:
                QMessageBox.warning(self, "Ошибка", f"Деталь '{part}' не найдена в справочнике")
                return
            
            # Создаём изменения для детали
            # Получаем текущий номер доп. страницы, если установлен
            additional_page = None
            if hasattr(self, 'get_current_additional_page') and callable(self.get_current_additional_page):
                additional_page = self.get_current_additional_page()
            part_changes = PartChanges(part=part, additional_page_number=additional_page)
            for entry in entries:
                material_change = MaterialChange(catalog_entry=entry)
                part_changes.materials.append(material_change)
            
            self.document_data.part_changes.append(part_changes)
        
        # Автоматически привязываем деталь к изделию, если изделие выбрано
        self._link_part_to_product(part)
        
        self.refresh()
    
    def _link_part_to_product(self, part: str):
        """Привязать деталь к изделию"""
        product_name = self.document_data.product.strip()
        if product_name and self.product_store:
            self.product_store.link_part_to_product_by_name(product_name, part)
    
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
                
                # Страница (визуальная индикация)
                page_label = ""
                if part_change.additional_page_number is not None:
                    page_label = f"{part_change.additional_page_number}+"
                else:
                    page_label = "1"
                page_item = QTableWidgetItem(page_label)
                page_item.setFlags(page_item.flags() & ~Qt.ItemIsEditable)
                page_item.setTextAlignment(Qt.AlignCenter)
                self.table.setItem(row, 8, page_item)
                
                # Визуальное выделение строк для доп. страниц (светло-голубой фон)
                if part_change.additional_page_number is not None:
                    light_blue = QColor(173, 216, 230)  # Светло-голубой цвет
                    for col in range(9):
                        item = self.table.item(row, col)
                        if item:
                            item.setBackground(light_blue)
                        else:
                            # Если ячейка содержит виджет (чекбокс, комбобокс), устанавливаем цвет через таблицу
                            widget = self.table.cellWidget(row, col)
                            if widget:
                                widget.setStyleSheet(f"background-color: rgb({light_blue.red()}, {light_blue.green()}, {light_blue.blue()});")
                
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
    
    def copy_and_edit_part(self):
        """Копировать деталь с теми же материалами и открыть диалог редактирования"""
        # Определяем выбранную деталь
        part = None
        
        # Приоритет: выбранная строка в таблице
        current_row = self.table.currentRow()
        if current_row >= 0:
            part_item = self.table.item(current_row, 0)
            if part_item:
                part = part_item.text()
        
        # Если строка не выбрана, используем деталь из комбобокса
        if not part:
            part = self.part_combo.currentText().strip()
        
        if not part:
            QMessageBox.warning(self, "Ошибка", "Выберите деталь для копирования")
            return
        
        # Получаем материалы детали
        replacement_sets = self.catalog_loader.get_replacement_sets_by_part(part)
        
        if replacement_sets:
            # Деталь имеет наборы замены
            from_sets = [s for s in replacement_sets if s.set_type == 'from']
            to_sets = [s for s in replacement_sets if s.set_type == 'to']
            
            # Объединяем материалы из всех наборов "from"
            from_materials = []
            for from_set in from_sets:
                # Создаём копии материалов с очищенными ID
                for entry in from_set.materials:
                    copied_entry = CatalogEntry(
                        part="",  # Будет установлен пользователем
                        workshop=entry.workshop,
                        role=entry.role,
                        before_name=entry.before_name,
                        unit=entry.unit,
                        norm=entry.norm,
                        comment=entry.comment,
                        id=None,
                        is_part_of_set=False,
                        replacement_set_id=None
                    )
                    from_materials.append(copied_entry)
            
            # Выбираем набор "to" (если несколько - берём первый)
            to_materials = []
            if to_sets:
                selected_to_set = to_sets[0]
                for entry in selected_to_set.materials:
                    copied_entry = CatalogEntry(
                        part="",  # Будет установлен пользователем
                        workshop=entry.workshop,
                        role=entry.role,
                        before_name=entry.before_name,
                        unit=entry.unit,
                        norm=entry.norm,
                        comment=entry.comment,
                        id=None,
                        is_part_of_set=False,
                        replacement_set_id=None
                    )
                    to_materials.append(copied_entry)
            
            # Открываем диалог создания с предзаполненными материалами
            dialog = PartCreationDialog(
                self.catalog_loader, 
                self,
                copy_from_materials=(from_materials, to_materials)
            )
            if dialog.exec():
                # Получаем данные
                part_code, new_from_materials, new_to_materials = dialog.get_replacement_set_data()
                
                # Создаём набор материалов
                set_id = self.catalog_loader.add_replacement_set(part_code, new_from_materials, new_to_materials)
                if set_id:
                    QMessageBox.information(self, "Успех", 
                                          f"Деталь '{part_code}' успешно создана")
                    
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
                    
                    # Открываем диалог редактирования для изменения норм
                    replacement_sets_new = self.catalog_loader.get_replacement_sets_by_part(part_code)
                    to_sets_new = [s for s in replacement_sets_new if s.set_type == 'to']
                    if to_sets_new:
                        edit_dialog = PartCreationDialog(
                            self.catalog_loader, 
                            self, 
                            edit_set_id=to_sets_new[0].id
                        )
                        if edit_dialog.exec():
                            # Сохраняем изменения
                            _, edited_from_materials, edited_to_materials = edit_dialog.get_replacement_set_data()
                            from_set_id, to_set_id = edit_dialog.get_edit_set_ids()
                            
                            success = True
                            if from_set_id:
                                success = self.catalog_loader.update_replacement_set(from_set_id, edited_from_materials)
                            if to_set_id and success:
                                success = self.catalog_loader.update_replacement_set(to_set_id, edited_to_materials)
                            
                            if success:
                                QMessageBox.information(self, "Успех", 
                                                      f"Нормы для детали '{part_code}' успешно обновлены")
                                self.load_parts_list()
                else:
                    QMessageBox.warning(self, "Ошибка", "Не удалось сохранить набор материалов")
        else:
            # Деталь без наборов - копируем обычные записи каталога
            entries = self.catalog_loader.get_entries_by_part(part)
            if not entries:
                QMessageBox.warning(self, "Ошибка", f"Деталь '{part}' не найдена в справочнике")
                return
            
            # Создаём копии материалов
            from_materials = []
            for entry in entries:
                copied_entry = CatalogEntry(
                    part="",  # Будет установлен пользователем
                    workshop=entry.workshop,
                    role=entry.role,
                    before_name=entry.before_name,
                    unit=entry.unit,
                    norm=entry.norm,
                    comment=entry.comment,
                    id=None,
                    is_part_of_set=False,
                    replacement_set_id=None
                )
                from_materials.append(copied_entry)
            
            # Для деталей без наборов создаём набор с одинаковыми материалами в "from" и "to"
            to_materials = []
            for entry in entries:
                copied_entry = CatalogEntry(
                    part="",  # Будет установлен пользователем
                    workshop=entry.workshop,
                    role=entry.role,
                    before_name=entry.before_name,
                    unit=entry.unit,
                    norm=entry.norm,
                    comment=entry.comment,
                    id=None,
                    is_part_of_set=False,
                    replacement_set_id=None
                )
                to_materials.append(copied_entry)
            
            # Открываем диалог создания с предзаполненными материалами
            dialog = PartCreationDialog(
                self.catalog_loader, 
                self,
                copy_from_materials=(from_materials, to_materials)
            )
            if dialog.exec():
                # Получаем данные
                part_code, new_from_materials, new_to_materials = dialog.get_replacement_set_data()
                
                # Создаём набор материалов
                set_id = self.catalog_loader.add_replacement_set(part_code, new_from_materials, new_to_materials)
                if set_id:
                    QMessageBox.information(self, "Успех", 
                                          f"Деталь '{part_code}' успешно создана")
                    
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
                    
                    # Открываем диалог редактирования для изменения норм
                    replacement_sets_new = self.catalog_loader.get_replacement_sets_by_part(part_code)
                    to_sets_new = [s for s in replacement_sets_new if s.set_type == 'to']
                    if to_sets_new:
                        edit_dialog = PartCreationDialog(
                            self.catalog_loader, 
                            self, 
                            edit_set_id=to_sets_new[0].id
                        )
                        if edit_dialog.exec():
                            # Сохраняем изменения
                            _, edited_from_materials, edited_to_materials = edit_dialog.get_replacement_set_data()
                            from_set_id, to_set_id = edit_dialog.get_edit_set_ids()
                            
                            success = True
                            if from_set_id:
                                success = self.catalog_loader.update_replacement_set(from_set_id, edited_from_materials)
                            if to_set_id and success:
                                success = self.catalog_loader.update_replacement_set(to_set_id, edited_to_materials)
                            
                            if success:
                                QMessageBox.information(self, "Успех", 
                                                      f"Нормы для детали '{part_code}' успешно обновлены")
                                self.load_parts_list()
                else:
                    QMessageBox.warning(self, "Ошибка", "Не удалось сохранить набор материалов")
    
    def edit_replacement_set(self):
        """Редактировать набор материалов для выбранной детали"""
        part = self.part_combo.currentText().strip()
        if not part:
            QMessageBox.warning(self, "Ошибка", "Выберите деталь для редактирования набора")
            return
        
        # Получаем наборы для детали
        replacement_sets = self.catalog_loader.get_replacement_sets_by_part(part)
        if not replacement_sets:
            QMessageBox.information(self, "Информация", 
                                  f"Для детали '{part}' не найдено наборов материалов")
            return
        
        # Фильтруем наборы "to" (редактируем набор "после")
        to_sets = [s for s in replacement_sets if s.set_type == 'to']
        
        if not to_sets:
            QMessageBox.information(self, "Информация", 
                                  f"Для детали '{part}' не найдено наборов 'после'")
            return
        
        # Если наборов несколько, показываем диалог выбора
        selected_to_set: Optional[MaterialReplacementSet] = None
        if len(to_sets) == 1:
            selected_to_set = to_sets[0]
        else:
            dialog = ReplacementSetSelectionDialog(to_sets, part, self)
            if dialog.exec():
                selected_to_set = dialog.get_selected_set()
            else:
                return
        
        if not selected_to_set:
            return
        
        # Открываем диалог редактирования
        edit_dialog = PartCreationDialog(self.catalog_loader, self, edit_set_id=selected_to_set.id)
        if edit_dialog.exec():
            # Сохраняем изменения
            part_code, from_materials, to_materials = edit_dialog.get_replacement_set_data()
            from_set_id, to_set_id = edit_dialog.get_edit_set_ids()
            
            success = True
            if from_set_id:
                success = self.catalog_loader.update_replacement_set(from_set_id, from_materials)
            if to_set_id and success:
                success = self.catalog_loader.update_replacement_set(to_set_id, to_materials)
            
            if success:
                QMessageBox.information(self, "Успех", 
                                      f"Набор материалов для детали '{part_code}' успешно обновлён")
                self.load_parts_list()
            else:
                QMessageBox.warning(self, "Ошибка", "Не удалось обновить набор материалов")
