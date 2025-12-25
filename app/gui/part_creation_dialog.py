"""
Диалог создания новой детали
"""
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QFormLayout,
    QLineEdit, QPushButton, QMessageBox, QDoubleSpinBox, QComboBox,
    QRadioButton, QButtonGroup, QTableWidget, QTableWidgetItem,
    QHeaderView, QGroupBox, QLabel, QWidget
)
from PySide6.QtCore import Qt
from typing import List, Optional

from app.models import CatalogEntry
from app.config import WORKSHOPS
from app.catalog_loader import CatalogLoader
from app.gui.material_selection_dialog import MaterialSelectionDialog


class PartCreationDialog(QDialog):
    """Диалог для создания новой детали с материалами"""
    
    def __init__(self, catalog_loader: CatalogLoader = None, parent=None, 
                 edit_set_id: Optional[int] = None,
                 copy_from_materials: Optional[tuple] = None):
        super().__init__(parent)
        self.edit_set_id = edit_set_id
        self.edit_from_set_id: Optional[int] = None
        self.edit_to_set_id: Optional[int] = None
        self.is_edit_mode = edit_set_id is not None
        self.copy_from_materials = copy_from_materials  # (from_materials, to_materials)
        
        if self.is_edit_mode:
            self.setWindowTitle("Редактировать набор материалов")
        else:
            self.setWindowTitle("Создать новую деталь")
        
        self.setMinimumSize(900, 700)
        self.entry: CatalogEntry = None
        self.from_materials: List[CatalogEntry] = []
        self.to_materials: List[CatalogEntry] = []
        self.catalog_loader = catalog_loader or CatalogLoader()
        self.is_replacement_set_mode = False
        self.init_ui()
        
        # Если режим редактирования, загружаем данные
        if self.is_edit_mode:
            self.load_set_for_editing()
        # Если режим копирования, предзаполняем материалы
        elif self.copy_from_materials:
            from_materials, to_materials = self.copy_from_materials
            self.from_materials = from_materials.copy() if from_materials else []
            self.to_materials = to_materials.copy() if to_materials else []
            # Включаем режим набора материалов
            self.mode_set.setChecked(True)
            self.is_replacement_set_mode = True
            self.single_widget.setVisible(False)
            self.set_widget.setVisible(True)
            # Обновляем таблицы
            self.update_from_table()
            self.update_to_table()
    
    def init_ui(self):
        """Инициализация интерфейса"""
        layout = QVBoxLayout(self)
        
        # Переключатель режимов (скрыт в режиме редактирования)
        mode_group = QGroupBox("Режим создания")
        mode_layout = QVBoxLayout()
        self.mode_single = QRadioButton("Одиночный материал")
        self.mode_set = QRadioButton("Набор материалов для замены")
        self.mode_single.setChecked(True)
        self.mode_single.toggled.connect(self.on_mode_changed)
        self.mode_set.toggled.connect(self.on_mode_changed)
        mode_layout.addWidget(self.mode_single)
        mode_layout.addWidget(self.mode_set)
        mode_group.setLayout(mode_layout)
        
        # В режиме редактирования скрываем переключатель и сразу включаем режим набора
        if self.is_edit_mode:
            mode_group.setVisible(False)
            self.mode_set.setChecked(True)
            self.is_replacement_set_mode = True
        
        layout.addWidget(mode_group)
        
        # Виджет для одиночного материала
        self.single_widget = self.create_single_material_widget()
        layout.addWidget(self.single_widget)
        
        # Виджет для набора материалов
        self.set_widget = self.create_replacement_set_widget()
        if not self.is_edit_mode:
            self.set_widget.setVisible(False)
        else:
            self.set_widget.setVisible(True)
            # В режиме редактирования делаем поле кода детали только для чтения
            self.part_edit_set.setReadOnly(True)
        layout.addWidget(self.set_widget)
        
        # Кнопки
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        self.btn_cancel = QPushButton("Отмена")
        self.btn_cancel.clicked.connect(self.reject)
        button_layout.addWidget(self.btn_cancel)
        
        self.btn_save = QPushButton("Сохранить")
        self.btn_save.clicked.connect(self.accept_and_validate)
        self.btn_save.setDefault(True)
        button_layout.addWidget(self.btn_save)
        
        layout.addLayout(button_layout)
    
    def create_single_material_widget(self) -> QWidget:
        """Создать виджет для одиночного материала"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        form_layout = QFormLayout()
        
        # Код детали
        self.part_edit = QLineEdit()
        self.part_edit.setPlaceholderText("например: КИР 03.614")
        form_layout.addRow("Код детали:", self.part_edit)
        
        # Цех
        self.workshop_combo = QComboBox()
        self.workshop_combo.addItems(WORKSHOPS)
        self.workshop_combo.setEditable(True)
        form_layout.addRow("Цех:", self.workshop_combo)
        
        # Тип позиции
        self.role_edit = QLineEdit()
        self.role_edit.setPlaceholderText("например: краска, разбавитель, отвердитель")
        form_layout.addRow("Тип позиции:", self.role_edit)
        
        # Наименование материала
        self.before_name_edit = QLineEdit()
        self.before_name_edit.setPlaceholderText("Полное наименование материала")
        form_layout.addRow("Наименование материала:", self.before_name_edit)
        
        # Единица измерения
        self.unit_edit = QLineEdit()
        self.unit_edit.setPlaceholderText("например: кг, шт, м")
        form_layout.addRow("Ед. изм.:", self.unit_edit)
        
        # Норма
        self.norm_spin = QDoubleSpinBox()
        self.norm_spin.setMinimum(0.0)
        self.norm_spin.setMaximum(999999.9999)
        self.norm_spin.setDecimals(4)
        self.norm_spin.setValue(0.0)
        form_layout.addRow("Норма:", self.norm_spin)
        
        # Примечание
        self.comment_edit = QLineEdit()
        self.comment_edit.setPlaceholderText("Опционально")
        form_layout.addRow("Примечание:", self.comment_edit)
        
        layout.addLayout(form_layout)
        layout.addStretch()
        return widget
    
    def create_replacement_set_widget(self) -> QWidget:
        """Создать виджет для набора материалов"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        # Код детали
        part_layout = QHBoxLayout()
        part_layout.addWidget(QLabel("Код детали:"))
        self.part_edit_set = QLineEdit()
        self.part_edit_set.setPlaceholderText("например: КИР 03.614")
        part_layout.addWidget(self.part_edit_set)
        layout.addLayout(part_layout)
        
        # Набор материалов "до"
        from_group = QGroupBox("Набор материалов 'до' (что заменяем)")
        from_layout = QVBoxLayout()
        
        from_buttons = QHBoxLayout()
        self.btn_add_from_new = QPushButton("Добавить новый материал")
        self.btn_add_from_new.clicked.connect(self.add_from_material_new)
        from_buttons.addWidget(self.btn_add_from_new)
        
        self.btn_add_from_existing = QPushButton("Выбрать из каталога")
        self.btn_add_from_existing.clicked.connect(self.add_from_material_existing)
        from_buttons.addWidget(self.btn_add_from_existing)
        
        self.btn_remove_from = QPushButton("Удалить выбранный")
        self.btn_remove_from.clicked.connect(self.remove_from_material)
        from_buttons.addWidget(self.btn_remove_from)
        from_buttons.addStretch()
        from_layout.addLayout(from_buttons)
        
        self.from_table = QTableWidget()
        self.from_table.setColumnCount(5)
        self.from_table.setHorizontalHeaderLabels([
            "Цех", "Тип позиции", "Наименование", "Ед. изм.", "Норма"
        ])
        self.from_table.horizontalHeader().setStretchLastSection(False)
        self.from_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.from_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        # Подключаем обработчик изменений для обновления норм (один раз)
        self.from_table.itemChanged.connect(self.on_from_norm_changed)
        from_layout.addWidget(self.from_table)
        from_group.setLayout(from_layout)
        layout.addWidget(from_group)
        
        # Набор материалов "после"
        to_group = QGroupBox("Набор материалов 'после' (на что заменяем)")
        to_layout = QVBoxLayout()
        
        to_buttons = QHBoxLayout()
        self.btn_add_to_new = QPushButton("Добавить новый материал")
        self.btn_add_to_new.clicked.connect(self.add_to_material_new)
        to_buttons.addWidget(self.btn_add_to_new)
        
        self.btn_add_to_existing = QPushButton("Выбрать из каталога")
        self.btn_add_to_existing.clicked.connect(self.add_to_material_existing)
        to_buttons.addWidget(self.btn_add_to_existing)
        
        self.btn_remove_to = QPushButton("Удалить выбранный")
        self.btn_remove_to.clicked.connect(self.remove_to_material)
        to_buttons.addWidget(self.btn_remove_to)
        to_buttons.addStretch()
        to_layout.addLayout(to_buttons)
        
        self.to_table = QTableWidget()
        self.to_table.setColumnCount(5)
        self.to_table.setHorizontalHeaderLabels([
            "Цех", "Тип позиции", "Наименование", "Ед. изм.", "Норма"
        ])
        self.to_table.horizontalHeader().setStretchLastSection(False)
        self.to_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.to_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        # Подключаем обработчик изменений для обновления норм (один раз)
        self.to_table.itemChanged.connect(self.on_to_norm_changed)
        to_layout.addWidget(self.to_table)
        to_group.setLayout(to_layout)
        layout.addWidget(to_group)
        
        return widget
    
    def on_mode_changed(self):
        """Обработчик изменения режима"""
        self.is_replacement_set_mode = self.mode_set.isChecked()
        self.single_widget.setVisible(not self.is_replacement_set_mode)
        self.set_widget.setVisible(self.is_replacement_set_mode)
    
    def add_from_material_new(self):
        """Добавить новый материал в набор 'до'"""
        dialog = MaterialEntryDialog(self)
        if dialog.exec():
            entry = dialog.get_entry()
            if entry:
                self.from_materials.append(entry)
                self.update_from_table()
    
    def add_from_material_existing(self):
        """Выбрать существующий материал из каталога для набора 'до'"""
        part_code = self.part_edit_set.text().strip()
        if not part_code:
            QMessageBox.warning(self, "Ошибка", "Сначала укажите код детали")
            return
        
        dialog = MaterialSelectionDialog(part_code, self.catalog_loader, self)
        if dialog.exec():
            selected = dialog.get_selected_entries()
            self.from_materials.extend(selected)
            self.update_from_table()
    
    def add_to_material_new(self):
        """Добавить новый материал в набор 'после'"""
        dialog = MaterialEntryDialog(self)
        if dialog.exec():
            entry = dialog.get_entry()
            if entry:
                self.to_materials.append(entry)
                self.update_to_table()
    
    def add_to_material_existing(self):
        """Выбрать существующий материал из каталога для набора 'после'"""
        part_code = self.part_edit_set.text().strip()
        if not part_code:
            QMessageBox.warning(self, "Ошибка", "Сначала укажите код детали")
            return
        
        dialog = MaterialSelectionDialog(part_code, self.catalog_loader, self)
        if dialog.exec():
            selected = dialog.get_selected_entries()
            self.to_materials.extend(selected)
            self.update_to_table()
    
    def remove_from_material(self):
        """Удалить выбранный материал из набора 'до'"""
        row = self.from_table.currentRow()
        if row >= 0 and row < len(self.from_materials):
            self.from_materials.pop(row)
            self.update_from_table()
    
    def remove_to_material(self):
        """Удалить выбранный материал из набора 'после'"""
        row = self.to_table.currentRow()
        if row >= 0 and row < len(self.to_materials):
            self.to_materials.pop(row)
            self.update_to_table()
    
    def update_from_table(self):
        """Обновить таблицу материалов 'до'"""
        # Временно блокируем сигналы, чтобы избежать лишних вызовов при обновлении
        self.from_table.blockSignals(True)
        self.from_table.setRowCount(len(self.from_materials))
        for row, entry in enumerate(self.from_materials):
            self.from_table.setItem(row, 0, QTableWidgetItem(entry.workshop))
            self.from_table.setItem(row, 1, QTableWidgetItem(entry.role))
            self.from_table.setItem(row, 2, QTableWidgetItem(entry.before_name))
            self.from_table.setItem(row, 3, QTableWidgetItem(entry.unit))
            # Норма - редактируемая ячейка
            norm_item = QTableWidgetItem(f"{entry.norm:.4f}")
            norm_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            norm_item.setData(Qt.UserRole, entry)  # Сохраняем ссылку на entry
            self.from_table.setItem(row, 4, norm_item)
        
        # Разблокируем сигналы
        self.from_table.blockSignals(False)
    
    def update_to_table(self):
        """Обновить таблицу материалов 'после'"""
        # Временно блокируем сигналы, чтобы избежать лишних вызовов при обновлении
        self.to_table.blockSignals(True)
        self.to_table.setRowCount(len(self.to_materials))
        for row, entry in enumerate(self.to_materials):
            self.to_table.setItem(row, 0, QTableWidgetItem(entry.workshop))
            self.to_table.setItem(row, 1, QTableWidgetItem(entry.role))
            self.to_table.setItem(row, 2, QTableWidgetItem(entry.before_name))
            self.to_table.setItem(row, 3, QTableWidgetItem(entry.unit))
            # Норма - редактируемая ячейка
            norm_item = QTableWidgetItem(f"{entry.norm:.4f}")
            norm_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            norm_item.setData(Qt.UserRole, entry)  # Сохраняем ссылку на entry
            self.to_table.setItem(row, 4, norm_item)
        
        # Разблокируем сигналы
        self.to_table.blockSignals(False)
    
    def accept_and_validate(self):
        """Валидация и принятие диалога"""
        if not self.is_replacement_set_mode:
            # Валидация одиночного материала
            part = self.part_edit.text().strip()
            if not part:
                QMessageBox.warning(self, "Ошибка", "Укажите код детали")
                return
            
            workshop = self.workshop_combo.currentText().strip()
            if not workshop:
                QMessageBox.warning(self, "Ошибка", "Укажите цех")
                return
            
            role = self.role_edit.text().strip()
            if not role:
                QMessageBox.warning(self, "Ошибка", "Укажите тип позиции")
                return
            
            before_name = self.before_name_edit.text().strip()
            if not before_name:
                QMessageBox.warning(self, "Ошибка", "Укажите наименование материала")
                return
            
            unit = self.unit_edit.text().strip()
            if not unit:
                QMessageBox.warning(self, "Ошибка", "Укажите единицу измерения")
                return
            
            norm = self.norm_spin.value()
            if norm <= 0:
                QMessageBox.warning(self, "Ошибка", "Норма должна быть больше нуля")
                return
            
            # Создаём запись
            self.entry = CatalogEntry(
                part=part,
                workshop=workshop,
                role=role,
                before_name=before_name,
                unit=unit,
                norm=norm,
                comment=self.comment_edit.text().strip()
            )
        else:
            # Валидация набора материалов
            part = self.part_edit_set.text().strip()
            if not part:
                QMessageBox.warning(self, "Ошибка", "Укажите код детали")
                return
            
            if not self.from_materials:
                QMessageBox.warning(self, "Ошибка", "Добавьте хотя бы один материал в набор 'до'")
                return
            
            if not self.to_materials:
                QMessageBox.warning(self, "Ошибка", "Добавьте хотя бы один материал в набор 'после'")
                return
            
            # Валидация каждого материала
            for i, material in enumerate(self.from_materials):
                if not material.workshop or not material.role or not material.before_name or not material.unit:
                    QMessageBox.warning(self, "Ошибка", 
                                      f"Материал {i+1} в наборе 'до' имеет незаполненные обязательные поля")
                    return
                if material.norm <= 0:
                    QMessageBox.warning(self, "Ошибка", 
                                      f"Норма материала {i+1} в наборе 'до' должна быть больше нуля")
                    return
            
            for i, material in enumerate(self.to_materials):
                if not material.workshop or not material.role or not material.before_name or not material.unit:
                    QMessageBox.warning(self, "Ошибка", 
                                      f"Материал {i+1} в наборе 'после' имеет незаполненные обязательные поля")
                    return
                if material.norm <= 0:
                    QMessageBox.warning(self, "Ошибка", 
                                      f"Норма материала {i+1} в наборе 'после' должна быть больше нуля")
                    return
        
        self.accept()
    
    def get_entry(self) -> Optional[CatalogEntry]:
        """Получить созданную запись (для одиночного режима)"""
        return self.entry
    
    def load_set_for_editing(self):
        """Загрузить набор для редактирования"""
        if not self.edit_set_id:
            return
        
        # Получаем набор "to" по ID
        to_set = self.catalog_loader.get_replacement_set_by_id(self.edit_set_id)
        if not to_set:
            return
        
        # Находим соответствующий набор "from" для той же детали
        all_sets = self.catalog_loader.get_replacement_sets_by_part(to_set.part_code)
        from_set = None
        for s in all_sets:
            if s.set_type == 'from' and s.part_code == to_set.part_code:
                # Берем первый найденный набор "from" (обычно он один)
                from_set = s
                break
        
        # Заполняем форму
        self.part_edit_set.setText(to_set.part_code)
        
        if from_set:
            self.edit_from_set_id = from_set.id
            self.from_materials = from_set.materials.copy()
            self.update_from_table()
        
        self.edit_to_set_id = to_set.id
        self.to_materials = to_set.materials.copy()
        self.update_to_table()
    
    def get_replacement_set_data(self) -> tuple:
        """Получить данные набора замены: (part_code, from_materials, to_materials)"""
        return (self.part_edit_set.text().strip(), self.from_materials, self.to_materials)
    
    def get_edit_set_ids(self) -> tuple:
        """Получить ID наборов для редактирования: (from_set_id, to_set_id)"""
        return (self.edit_from_set_id, self.edit_to_set_id)
    
    def on_from_norm_changed(self, item: QTableWidgetItem):
        """Обработчик изменения нормы в таблице 'до'"""
        if item.column() == 4:  # Колонка "Норма"
            entry = item.data(Qt.UserRole)
            if entry:
                try:
                    new_norm = float(item.text().replace(',', '.'))
                    if new_norm > 0:
                        entry.norm = new_norm
                        # Обновляем форматирование
                        item.setText(f"{new_norm:.4f}")
                    else:
                        # Восстанавливаем старое значение
                        item.setText(f"{entry.norm:.4f}")
                except ValueError:
                    # Восстанавливаем старое значение при ошибке
                    item.setText(f"{entry.norm:.4f}")
    
    def on_to_norm_changed(self, item: QTableWidgetItem):
        """Обработчик изменения нормы в таблице 'после'"""
        if item.column() == 4:  # Колонка "Норма"
            entry = item.data(Qt.UserRole)
            if entry:
                try:
                    new_norm = float(item.text().replace(',', '.'))
                    if new_norm > 0:
                        entry.norm = new_norm
                        # Обновляем форматирование
                        item.setText(f"{new_norm:.4f}")
                    else:
                        # Восстанавливаем старое значение
                        item.setText(f"{entry.norm:.4f}")
                except ValueError:
                    # Восстанавливаем старое значение при ошибке
                    item.setText(f"{entry.norm:.4f}")


class MaterialEntryDialog(QDialog):
    """Диалог для ввода одного материала"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Добавить материал")
        self.setMinimumWidth(400)
        self.entry: CatalogEntry = None
        self.init_ui()
    
    def init_ui(self):
        """Инициализация интерфейса"""
        layout = QVBoxLayout(self)
        form_layout = QFormLayout()
        
        # Цех
        self.workshop_combo = QComboBox()
        self.workshop_combo.addItems(WORKSHOPS)
        self.workshop_combo.setEditable(True)
        form_layout.addRow("Цех:", self.workshop_combo)
        
        # Тип позиции
        self.role_edit = QLineEdit()
        self.role_edit.setPlaceholderText("например: краска, разбавитель, отвердитель")
        form_layout.addRow("Тип позиции:", self.role_edit)
        
        # Наименование материала
        self.before_name_edit = QLineEdit()
        self.before_name_edit.setPlaceholderText("Полное наименование материала")
        form_layout.addRow("Наименование материала:", self.before_name_edit)
        
        # Единица измерения
        self.unit_edit = QLineEdit()
        self.unit_edit.setPlaceholderText("например: кг, шт, м")
        form_layout.addRow("Ед. изм.:", self.unit_edit)
        
        # Норма
        self.norm_spin = QDoubleSpinBox()
        self.norm_spin.setMinimum(0.0)
        self.norm_spin.setMaximum(999999.9999)
        self.norm_spin.setDecimals(4)
        self.norm_spin.setValue(0.0)
        form_layout.addRow("Норма:", self.norm_spin)
        
        # Примечание
        self.comment_edit = QLineEdit()
        self.comment_edit.setPlaceholderText("Опционально")
        form_layout.addRow("Примечание:", self.comment_edit)
        
        layout.addLayout(form_layout)
        
        # Кнопки
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        btn_cancel = QPushButton("Отмена")
        btn_cancel.clicked.connect(self.reject)
        button_layout.addWidget(btn_cancel)
        
        btn_ok = QPushButton("ОК")
        btn_ok.clicked.connect(self.accept_and_validate)
        btn_ok.setDefault(True)
        button_layout.addWidget(btn_ok)
        
        layout.addLayout(button_layout)
    
    def accept_and_validate(self):
        """Валидация и принятие диалога"""
        workshop = self.workshop_combo.currentText().strip()
        if not workshop:
            QMessageBox.warning(self, "Ошибка", "Укажите цех")
            return
        
        role = self.role_edit.text().strip()
        if not role:
            QMessageBox.warning(self, "Ошибка", "Укажите тип позиции")
            return
        
        before_name = self.before_name_edit.text().strip()
        if not before_name:
            QMessageBox.warning(self, "Ошибка", "Укажите наименование материала")
            return
        
        unit = self.unit_edit.text().strip()
        if not unit:
            QMessageBox.warning(self, "Ошибка", "Укажите единицу измерения")
            return
        
        norm = self.norm_spin.value()
        if norm <= 0:
            QMessageBox.warning(self, "Ошибка", "Норма должна быть больше нуля")
            return
        
        self.entry = CatalogEntry(
            part="",  # Будет установлен позже
            workshop=workshop,
            role=role,
            before_name=before_name,
            unit=unit,
            norm=norm,
            comment=self.comment_edit.text().strip()
        )
        
        self.accept()
    
    def get_entry(self) -> CatalogEntry:
        """Получить созданную запись"""
        return self.entry
