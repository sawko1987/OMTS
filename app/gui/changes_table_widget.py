"""
Виджет для редактирования таблицы изменений материалов
"""
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QComboBox,
    QTableWidget, QTableWidgetItem, QCheckBox, QHeaderView,
    QMessageBox, QLabel, QLineEdit, QCompleter, QListWidget, QListWidgetItem
)
from PySide6.QtCore import Qt, QStringListModel, QTimer, Signal
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

    # Сигнал: выбранная/активная деталь (для синхронизации с другими вкладками)
    part_selected = Signal(str)
    
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
        
        # Выбор машин (изделий)
        machines_label = QLabel("Выберите машины (изделия):")
        machines_label.setStyleSheet("font-weight: bold;")
        layout.addWidget(machines_label)
        
        self.machines_list = QListWidget()
        self.machines_list.setMaximumHeight(120)
        # Используем чекбоксы для множественного выбора
        self.machines_list.itemChanged.connect(self.on_machines_selection_changed)
        layout.addWidget(self.machines_list)
        
        # Кнопка создания новой машины
        machine_buttons_layout = QHBoxLayout()
        self.btn_create_machine = QPushButton("Создать новую машину")
        self.btn_create_machine.clicked.connect(self.create_machine)
        machine_buttons_layout.addWidget(self.btn_create_machine)
        machine_buttons_layout.addStretch()
        layout.addLayout(machine_buttons_layout)
        
        # Информация о выбранных изделиях
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
        
        self.btn_remove_row = QPushButton("Удалить строку")
        self.btn_remove_row.clicked.connect(self.remove_selected_row)
        add_layout.addWidget(self.btn_remove_row)
        
        self.btn_delete_part_from_db = QPushButton("Удалить деталь из БД")
        self.btn_delete_part_from_db.clicked.connect(self.delete_part_from_database)
        add_layout.addWidget(self.btn_delete_part_from_db)
        
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
        self.table.itemSelectionChanged.connect(self._on_table_selection_changed)
        
        layout.addWidget(self.table)
        
        # Загружаем список деталей при инициализации
        self.load_machines_list()
        self.load_parts_list()
    
    def load_machines_list(self):
        """Загрузить список машин (изделий) в виджет"""
        if not self.product_store:
            return
        
        self.machines_list.clear()
        products = self.product_store.get_all_products()
        
        # Получаем текущий список выбранных машин
        selected_products = set(self.document_data.products)
        
        for product_id, product_name in products:
            item = QListWidgetItem(product_name)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked if product_name in selected_products else Qt.Unchecked)
            item.setData(Qt.UserRole, product_name)
            self.machines_list.addItem(item)
        
        self.update_product_info_label()
    
    def on_machines_selection_changed(self, item: QListWidgetItem):
        """Обработка изменения выбора машин"""
        # Обновляем список выбранных машин
        selected_products = []
        for i in range(self.machines_list.count()):
            list_item = self.machines_list.item(i)
            if list_item and list_item.checkState() == Qt.Checked:
                product_name = list_item.data(Qt.UserRole)
                if product_name:
                    selected_products.append(product_name)
        
        self.document_data.products = selected_products
        self.update_product_info_label()
        # Обновляем фильтр деталей
        self.load_parts_list()
    
    def create_machine(self):
        """Создать новую машину (изделие)"""
        if not self.product_store:
            QMessageBox.warning(self, "Ошибка", "ProductStore не инициализирован")
            return
        
        from PySide6.QtWidgets import QInputDialog
        text, ok = QInputDialog.getText(
            self, "Создать машину", "Введите название машины (изделия):"
        )
        if not ok or not text.strip():
            return
        
        product_name = text.strip()
        
        # Проверяем, не существует ли уже такая машина
        existing_id = self.product_store.get_product_by_name(product_name)
        if existing_id:
            QMessageBox.information(
                self, "Информация", 
                f"Машина '{product_name}' уже существует"
            )
            # Устанавливаем её как выбранную
            for i in range(self.machines_list.count()):
                item = self.machines_list.item(i)
                if item and item.data(Qt.UserRole) == product_name:
                    item.setCheckState(Qt.Checked)
                    break
            return
        
        # Создаём новую машину
        product_id = self.product_store.add_product(product_name)
        if product_id:
            # Обновляем список
            self.load_machines_list()
            # Устанавливаем новую машину как выбранную
            for i in range(self.machines_list.count()):
                item = self.machines_list.item(i)
                if item and item.data(Qt.UserRole) == product_name:
                    item.setCheckState(Qt.Checked)
                    break
            QMessageBox.information(
                self, "Успех", 
                f"Машина '{product_name}' успешно создана"
            )
        else:
            QMessageBox.warning(self, "Ошибка", "Не удалось создать машину")
    
    def update_product_info_label(self):
        """Обновить информационную метку о выбранных машинах"""
        if self.document_data.products:
            products_text = ", ".join(self.document_data.products)
            self.product_info_label.setText(
                f"Выбрано машин: {len(self.document_data.products)} ({products_text})"
            )
        else:
            self.product_info_label.setText("Машины не выбраны. Показаны все детали.")
    
    def update_product_filter(self):
        """Обновить фильтр деталей по выбранному изделию"""
        self.load_parts_list()
    
    def load_parts_list(self):
        """Загрузить список деталей с учетом фильтра по изделиям"""
        try:
            selected_products = self.document_data.products
            logger.debug(f"Загрузка списка деталей для изделий: {selected_products}")
            
            if selected_products and self.product_store:
                # Загружаем детали, привязанные к выбранным изделиям
                all_parts = set()
                for product_name in selected_products:
                    parts = self.product_store.get_parts_by_product_name(product_name)
                    all_parts.update(parts)
                    logger.info(f"Загружено {len(parts)} деталей для изделия '{product_name}'")
                
                parts = sorted(list(all_parts))
                logger.info(f"Всего уникальных деталей для выбранных изделий: {len(parts)}")
            else:
                # Загружаем все детали
                parts = self.catalog_loader.get_all_parts()
                logger.info(f"Загружено всех деталей: {len(parts)}")
            
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
                selected_products = self.document_data.products
                logger.debug(f"Поиск для изделий: {selected_products}")
                
                if selected_products and self.product_store:
                    # Ищем среди деталей выбранных изделий
                    all_parts = set()
                    for product_name in selected_products:
                        product_parts = self.product_store.get_parts_by_product_name(product_name)
                        all_parts.update(product_parts)
                    logger.info(f"Деталей в выбранных изделиях: {len(all_parts)}")
                    
                    # Фильтруем по поисковому запросу
                    search_results = self.catalog_loader.search_parts(search_text)
                    logger.info(f"Результатов поиска по запросу '{search_text}': {len(search_results)}")
                    
                    # Пересечение: только те, что есть и в изделиях, и в результатах поиска
                    parts = [p for p in search_results if p in all_parts]
                    logger.info(f"Результатов после фильтрации по изделиям: {len(parts)}")
                else:
                    # Ищем среди всех деталей
                    logger.debug("Поиск среди всех деталей (изделия не выбраны или не найдены)")
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
        part = (text or "").strip()
        if part:
            self.part_selected.emit(part)

    def _on_table_selection_changed(self):
        """Обработка выбора строки в таблице (эмитим выбранную деталь)"""
        row = self.table.currentRow()
        if row < 0:
            return
        part_item = self.table.item(row, 0)
        if not part_item:
            return
        part = part_item.text().strip()
        if part:
            self.part_selected.emit(part)
    
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
                selected_products = self.document_data.products
                
                if selected_products and self.product_store:
                    all_parts = set()
                    for product_name in selected_products:
                        product_parts = self.product_store.get_parts_by_product_name(product_name)
                        all_parts.update(product_parts)
                    search_results = self.catalog_loader.search_parts(search_text)
                    parts = [p for p in search_results if p in all_parts]
                    logger.info(f"Поиск в комбобоксе: найдено {len(parts)} деталей для выбранных изделий")
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
                    
                    # Предлагаем привязать к выбранным изделиям
                    selected_products = self.document_data.products
                    if selected_products and self.product_store:
                        products_text = ", ".join(selected_products)
                        reply = QMessageBox.question(
                            self, "Привязать к изделиям?",
                            f"Привязать деталь '{part_code}' к выбранным изделиям ({products_text})?",
                            QMessageBox.Yes | QMessageBox.No
                        )
                        if reply == QMessageBox.Yes:
                            for product_name in selected_products:
                                self.product_store.link_part_to_product_by_name(product_name, part_code)
                    
                    # Обновляем список
                    self.load_parts_list()
                    # Устанавливаем новую деталь в комбобоксе
                    index = self.part_combo.findText(part_code)
                    if index >= 0:
                        self.part_combo.setCurrentIndex(index)
                else:
                    QMessageBox.warning(self, "Ошибка", "Не удалось сохранить набор материалов")
    
    def _deduplicate_materials(self, materials: List[CatalogEntry]) -> List[CatalogEntry]:
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
            else:
                logger.debug(f"Пропущен дубликат материала: {material.part}, {material.workshop}, {material.role}, {material.before_name}, {material.unit}")
        return unique_materials

    def _pick_matching_from_set(
        self,
        from_sets: List[MaterialReplacementSet],
        selected_to_set: Optional[MaterialReplacementSet],
    ) -> Optional[MaterialReplacementSet]:
        """
        Подобрать один набор 'from', соответствующий выбранному набору 'to'.

        Корневая причина дублей: раньше мы объединяли материалы из ВСЕХ from_sets.
        Теперь берём только один (наиболее подходящий) from_set.
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
    
    def add_part(self):
        """Добавить деталь в таблицу"""
        part = self.part_combo.currentText().strip()
        if not part:
            QMessageBox.warning(self, "Ошибка", "Укажите деталь")
            return
        
        # Проверяем, не добавлена ли уже эта деталь
        # Проверяем не только код детали, но и наличие материалов
        for part_change in self.document_data.part_changes:
            if part_change.part == part:
                materials_count = len(part_change.materials)
                QMessageBox.information(
                    self, 
                    "Информация", 
                    f"Деталь '{part}' уже добавлена ({materials_count} материалов)"
                )
                return
        
        # Проверяем наличие наборов замены для детали
        replacement_sets = self.catalog_loader.get_replacement_sets_by_part(part)
        
        if replacement_sets:
            # Есть наборы замены - используем их
            from_sets = [s for s in replacement_sets if s.set_type == 'from']
            to_sets = [s for s in replacement_sets if s.set_type == 'to']
            
            # Выбираем набор "to"
            selected_to_set: Optional[MaterialReplacementSet] = None
            if len(to_sets) == 1:
                # Один набор - используем автоматически
                selected_to_set = to_sets[0]
            elif len(to_sets) > 1:
                # Несколько наборов - показываем диалог выбора
                dialog = ReplacementSetSelectionDialog(to_sets, part, self, catalog_loader=self.catalog_loader)
                if dialog.exec():
                    selected_to_set = dialog.get_selected_set()
                else:
                    # Пользователь отменил выбор
                    return

            # Выбираем ОДИН набор "from", соответствующий выбранному "to"
            selected_from_set = self._pick_matching_from_set(from_sets, selected_to_set)
            if selected_from_set is None:
                QMessageBox.warning(self, "Ошибка", f"Для детали '{part}' не найден набор материалов 'до'")
                return

            from_materials = self._deduplicate_materials(selected_from_set.materials or [])
            logger.debug(
                f"Используем from_set_id={selected_from_set.id}, "
                f"to_set_id={selected_to_set.id if selected_to_set else None}. "
                f"Материалов 'до' после дедупликации: {len(from_materials)}"
            )
            
            # Создаём изменения для детали с наборами
            # Получаем текущий номер доп. страницы, если установлен
            additional_page = None
            if hasattr(self, 'get_current_additional_page') and callable(self.get_current_additional_page):
                additional_page = self.get_current_additional_page()
                logger.info(f"[add_part] Деталь '{part}': получен номер страницы из get_current_additional_page() = {additional_page}")
            else:
                logger.info(f"[add_part] Деталь '{part}': get_current_additional_page не доступен, additional_page = None")
            
            # ВАЖНО: Если номер страницы установлен, убеждаемся, что он больше максимального существующего
            # Это гарантирует, что деталь всегда попадает на новую страницу
            if additional_page is not None:
                max_existing_page = 0
                existing_pages = []
                all_parts_info = []
                for existing_part_change in self.document_data.part_changes:
                    all_parts_info.append((existing_part_change.part, existing_part_change.additional_page_number))
                    if existing_part_change.additional_page_number is not None:
                        max_existing_page = max(max_existing_page, existing_part_change.additional_page_number)
                        existing_pages.append((existing_part_change.part, existing_part_change.additional_page_number))
                logger.info(f"[add_part] Деталь '{part}': всего деталей в document_data = {len(self.document_data.part_changes)}")
                logger.info(f"[add_part] Деталь '{part}': все детали с их номерами страниц: {all_parts_info}")
                logger.info(f"[add_part] Деталь '{part}': максимальный существующий номер страницы = {max_existing_page}")
                logger.info(f"[add_part] Деталь '{part}': существующие детали по страницам: {existing_pages}")
                
                # Если установленный номер страницы не больше максимального, увеличиваем его
                if additional_page <= max_existing_page:
                    old_page = additional_page
                    additional_page = max_existing_page + 1
                    logger.warning(f"[add_part] Деталь '{part}': номер страницы увеличен с {old_page} до {additional_page} (был <= максимального {max_existing_page})")
                else:
                    logger.info(f"[add_part] Деталь '{part}': номер страницы {additional_page} корректен (больше максимального {max_existing_page})")
            
            logger.info(f"[add_part] Деталь '{part}': финальный номер страницы = {additional_page}")
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
                    logger.warning(f"Пропущен дубликат материала при добавлении в part_changes: {from_entry.part}, {from_entry.workshop}, {from_entry.role}, {from_entry.before_name}")
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
                        comment=""
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
                        comment=""
                    )
                    # Создаём MaterialChange только для правой колонки (is_changed=False, но after_name заполнен)
                    material_change = MaterialChange(
                        catalog_entry=dummy_entry,
                        is_changed=False,  # Не попадает в левую колонку
                        after_name=to_entry.before_name,  # Попадает в правую колонку
                        after_unit=to_entry.unit if to_entry.unit else template_entry.unit,
                        after_norm=to_entry.norm if to_entry.norm > 0 else None
                    )
                    part_changes.materials.append(material_change)
            
            self.document_data.part_changes.append(part_changes)
            logger.info(f"[add_part] Деталь '{part}' добавлена в document_data с номером страницы = {part_changes.additional_page_number}")
            logger.info(f"[add_part] Всего деталей в document_data: {len(self.document_data.part_changes)}")
            all_pages_after = [(pc.part, pc.additional_page_number) for pc in self.document_data.part_changes if pc.additional_page_number is not None]
            logger.info(f"[add_part] Все детали с доп. страницами после добавления: {all_pages_after}")
        else:
            # Нет наборов - используем старую логику
            entries = self.catalog_loader.get_entries_by_part(part)
            if not entries:
                QMessageBox.warning(self, "Ошибка", f"Деталь '{part}' не найдена в справочнике")
                return
            
            # Фильтруем материалы, которые входят в наборы (чтобы избежать дублирования)
            # Материалы из наборов должны использоваться только через наборы
            entries = [e for e in entries if not e.is_part_of_set]
            logger.debug(f"После фильтрации материалов из наборов осталось {len(entries)} записей из обычного каталога")
            
            # Дедупликация записей из каталога (на случай дубликатов в БД)
            entries = self._deduplicate_materials(entries)
            logger.debug(f"После дедупликации осталось {len(entries)} уникальных записей из каталога")
            
            # Создаём изменения для детали
            # Получаем текущий номер доп. страницы, если установлен
            additional_page = None
            if hasattr(self, 'get_current_additional_page') and callable(self.get_current_additional_page):
                additional_page = self.get_current_additional_page()
                logger.info(f"[add_part] Деталь '{part}' (без наборов): получен номер страницы из get_current_additional_page() = {additional_page}")
            else:
                logger.info(f"[add_part] Деталь '{part}' (без наборов): get_current_additional_page не доступен, additional_page = None")
            
            # ВАЖНО: Если номер страницы установлен, убеждаемся, что он больше максимального существующего
            # Это гарантирует, что деталь всегда попадает на новую страницу
            if additional_page is not None:
                max_existing_page = 0
                existing_pages = []
                all_parts_info = []
                for existing_part_change in self.document_data.part_changes:
                    all_parts_info.append((existing_part_change.part, existing_part_change.additional_page_number))
                    if existing_part_change.additional_page_number is not None:
                        max_existing_page = max(max_existing_page, existing_part_change.additional_page_number)
                        existing_pages.append((existing_part_change.part, existing_part_change.additional_page_number))
                logger.info(f"[add_part] Деталь '{part}' (без наборов): всего деталей в document_data = {len(self.document_data.part_changes)}")
                logger.info(f"[add_part] Деталь '{part}' (без наборов): все детали с их номерами страниц: {all_parts_info}")
                logger.info(f"[add_part] Деталь '{part}' (без наборов): максимальный существующий номер страницы = {max_existing_page}")
                logger.info(f"[add_part] Деталь '{part}' (без наборов): существующие детали по страницам: {existing_pages}")
                
                # Если установленный номер страницы не больше максимального, увеличиваем его
                if additional_page <= max_existing_page:
                    old_page = additional_page
                    additional_page = max_existing_page + 1
                    logger.warning(f"[add_part] Деталь '{part}' (без наборов): номер страницы увеличен с {old_page} до {additional_page} (был <= максимального {max_existing_page})")
                else:
                    logger.info(f"[add_part] Деталь '{part}' (без наборов): номер страницы {additional_page} корректен (больше максимального {max_existing_page})")
            
            logger.info(f"[add_part] Деталь '{part}' (без наборов): финальный номер страницы = {additional_page}")
            part_changes = PartChanges(part=part, additional_page_number=additional_page)
            for entry in entries:
                material_change = MaterialChange(catalog_entry=entry)
                part_changes.materials.append(material_change)
            
            self.document_data.part_changes.append(part_changes)
            logger.info(f"[add_part] Деталь '{part}' (без наборов) добавлена в document_data с номером страницы = {part_changes.additional_page_number}")
            logger.info(f"[add_part] Всего деталей в document_data: {len(self.document_data.part_changes)}")
            all_pages_after = [(pc.part, pc.additional_page_number) for pc in self.document_data.part_changes if pc.additional_page_number is not None]
            logger.info(f"[add_part] Все детали с доп. страницами после добавления: {all_pages_after}")
        
        # Автоматически привязываем деталь к изделию, если изделие выбрано
        self._link_part_to_product(part)
        
        self.refresh()
    
    def _link_part_to_product(self, part: str):
        """Привязать деталь к выбранным изделиям"""
        selected_products = self.document_data.products
        if selected_products and self.product_store:
            for product_name in selected_products:
                self.product_store.link_part_to_product_by_name(product_name, part)
    
    def remove_selected_part(self):
        """Удалить выбранную деталь (все строки детали)"""
        current_row = self.table.currentRow()
        if current_row < 0:
            QMessageBox.warning(self, "Ошибка", "Выберите строку для удаления")
            return
        
        # Находим деталь по строке
        part_item = self.table.item(current_row, 0)
        if not part_item:
            return
        
        part = part_item.text()
        
        # Подтверждение
        reply = QMessageBox.question(
            self,
            "Подтверждение",
            f"Удалить все строки детали '{part}'?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply != QMessageBox.Yes:
            return
        
        # Удаляем из данных
        self.document_data.part_changes = [
            pc for pc in self.document_data.part_changes
            if pc.part != part
        ]
        
        self.refresh()
    
    def remove_selected_row(self):
        """Удалить выбранную строку материала"""
        current_row = self.table.currentRow()
        if current_row < 0:
            QMessageBox.warning(self, "Ошибка", "Выберите строку для удаления")
            return
        
        # Получаем ссылки на part_change и material из первой колонки
        part_item = self.table.item(current_row, 0)
        if not part_item:
            return
        
        data = part_item.data(Qt.UserRole)
        # PySide/Qt могут вернуть list вместо tuple
        if isinstance(data, list):
            data = tuple(data)
        if not data or not isinstance(data, tuple) or len(data) != 2:
            QMessageBox.warning(self, "Ошибка", "Не удалось определить данные строки")
            return
        
        part_change, material = data
        
        # Подтверждение
        material_name = material.catalog_entry.before_name
        reply = QMessageBox.question(
            self,
            "Подтверждение",
            f"Удалить строку с материалом '{material_name}'?\n"
            f"Деталь: {part_change.part}",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply != QMessageBox.Yes:
            return
        
        # Удаляем материал из part_change
        if material in part_change.materials:
            part_change.materials.remove(material)
        
        # Если материалов не осталось, удаляем всю деталь
        if not part_change.materials:
            if part_change in self.document_data.part_changes:
                self.document_data.part_changes.remove(part_change)
        
        self.refresh()
    
    def refresh(self):
        """Обновить таблицу"""
        # Обновляем список машин
        self.load_machines_list()
        
        self.table.setRowCount(0)
        
        row = 0
        for part_change in self.document_data.part_changes:
            for material in part_change.materials:
                self.table.insertRow(row)
                
                # Деталь
                part_item = QTableWidgetItem(part_change.part)
                part_item.setFlags(part_item.flags() & ~Qt.ItemIsEditable)
                part_item.setData(Qt.UserRole, (part_change, material))  # Сохраняем ссылки для удаления строки
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
        # Обновляем список выбранных машин
        selected_products = []
        for i in range(self.machines_list.count()):
            item = self.machines_list.item(i)
            if item and item.checkState() == Qt.Checked:
                product_name = item.data(Qt.UserRole)
                if product_name:
                    selected_products.append(product_name)
        self.document_data.products = selected_products
        
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
            
            # Дедупликация материалов из наборов "from"
            from_materials = self._deduplicate_materials(from_materials)
            logger.debug(f"После дедупликации осталось {len(from_materials)} уникальных материалов из наборов 'from'")
            
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
                    
                    # Предлагаем привязать к выбранным изделиям
                    selected_products = self.document_data.products
                    if selected_products and self.product_store:
                        products_text = ", ".join(selected_products)
                        reply = QMessageBox.question(
                            self, "Привязать к изделиям?",
                            f"Привязать деталь '{part_code}' к выбранным изделиям ({products_text})?",
                            QMessageBox.Yes | QMessageBox.No
                        )
                        if reply == QMessageBox.Yes:
                            for product_name in selected_products:
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
            
            # Дедупликация материалов (для консистентности, хотя материалы уже из каталога)
            from_materials = self._deduplicate_materials(from_materials)
            logger.debug(f"После дедупликации осталось {len(from_materials)} уникальных материалов из каталога")
            
            # Для деталей без наборов создаём набор с одинаковыми материалами в "from" и "to"
            # Используем дедуплицированные материалы для создания to_materials
            to_materials = []
            for entry in from_materials:
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
                    
                    # Предлагаем привязать к выбранным изделиям
                    selected_products = self.document_data.products
                    if selected_products and self.product_store:
                        products_text = ", ".join(selected_products)
                        reply = QMessageBox.question(
                            self, "Привязать к изделиям?",
                            f"Привязать деталь '{part_code}' к выбранным изделиям ({products_text})?",
                            QMessageBox.Yes | QMessageBox.No
                        )
                        if reply == QMessageBox.Yes:
                            for product_name in selected_products:
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
            dialog = ReplacementSetSelectionDialog(to_sets, part, self, catalog_loader=self.catalog_loader)
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
    
    def delete_part_from_database(self):
        """Удалить деталь из базы данных"""
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
            QMessageBox.warning(self, "Ошибка", "Выберите деталь для удаления из базы данных")
            return
        
        # Проверяем, существует ли деталь в базе данных
        if not self.catalog_loader.part_exists(part):
            QMessageBox.warning(self, "Ошибка", f"Деталь '{part}' не найдена в базе данных")
            return
        
        # Показываем диалог подтверждения с предупреждением
        reply = QMessageBox.question(
            self, 
            "Подтверждение удаления",
            f"Вы уверены, что хотите полностью удалить деталь '{part}' из базы данных?\n\n"
            f"Это действие удалит:\n"
            f"- Все записи каталога для этой детали\n"
            f"- Все наборы материалов для замены\n"
            f"- Все связи с изделиями\n"
            f"- Историю замен\n\n"
            f"Это действие нельзя отменить!",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply != QMessageBox.Yes:
            return
        
        # Удаляем деталь из базы данных
        success = self.catalog_loader.delete_part(part)
        
        if success:
            # Удаляем деталь из текущего документа, если она там есть
            self.document_data.part_changes = [
                pc for pc in self.document_data.part_changes
                if pc.part != part
            ]
            
            # Обновляем список деталей
            self.load_parts_list()
            
            # Обновляем таблицу
            self.refresh()
            
            QMessageBox.information(
                self, 
                "Успех", 
                f"Деталь '{part}' успешно удалена из базы данных"
            )
        else:
            QMessageBox.warning(
                self, 
                "Ошибка", 
                f"Не удалось удалить деталь '{part}' из базы данных"
            )
