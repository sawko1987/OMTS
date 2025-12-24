"""
Диалог выбора существующих материалов из каталога
"""
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton,
    QTableWidget, QTableWidgetItem, QCheckBox, QHeaderView,
    QMessageBox, QLabel, QComboBox, QLineEdit
)
from PySide6.QtCore import Qt
from typing import List, Optional

from app.models import CatalogEntry
from app.catalog_loader import CatalogLoader
from app.config import WORKSHOPS


class MaterialSelectionDialog(QDialog):
    """Диалог для выбора существующих материалов из каталога"""
    
    def __init__(self, part_code: str, catalog_loader: CatalogLoader = None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Выбор материалов из каталога")
        self.setMinimumSize(800, 600)
        self.part_code = part_code
        self.catalog_loader = catalog_loader or CatalogLoader()
        self.selected_entries: List[CatalogEntry] = []
        self.all_entries: List[CatalogEntry] = []
        self.init_ui()
        self.load_materials()
    
    def init_ui(self):
        """Инициализация интерфейса"""
        layout = QVBoxLayout(self)
        
        # Заголовок с кодом детали
        header_label = QLabel(f"Материалы для детали: {self.part_code}")
        header_label.setStyleSheet("font-weight: bold; font-size: 12pt;")
        layout.addWidget(header_label)
        
        # Фильтры
        filter_layout = QHBoxLayout()
        
        filter_layout.addWidget(QLabel("Цех:"))
        self.workshop_filter = QComboBox()
        self.workshop_filter.addItem("Все", "")
        self.workshop_filter.addItems(WORKSHOPS)
        self.workshop_filter.currentIndexChanged.connect(self.apply_filters)
        filter_layout.addWidget(self.workshop_filter)
        
        filter_layout.addWidget(QLabel("Тип позиции:"))
        self.role_filter = QLineEdit()
        self.role_filter.setPlaceholderText("Фильтр по типу позиции...")
        self.role_filter.textChanged.connect(self.apply_filters)
        filter_layout.addWidget(self.role_filter)
        
        filter_layout.addStretch()
        layout.addLayout(filter_layout)
        
        # Таблица материалов
        self.table = QTableWidget()
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels([
            "",  # Чекбокс
            "Цех",
            "Тип позиции",
            "Наименование материала",
            "Ед. изм.",
            "Норма",
            "Примечание"
        ])
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        
        # Настройка ширины колонок
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)  # Чекбокс
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)  # Цех
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)  # Тип позиции
        header.setSectionResizeMode(3, QHeaderView.Stretch)  # Наименование
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)  # Ед. изм.
        header.setSectionResizeMode(5, QHeaderView.ResizeToContents)  # Норма
        
        layout.addWidget(self.table)
        
        # Кнопки
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        self.btn_cancel = QPushButton("Отмена")
        self.btn_cancel.clicked.connect(self.reject)
        button_layout.addWidget(self.btn_cancel)
        
        self.btn_select = QPushButton("Выбрать")
        self.btn_select.clicked.connect(self.accept_selection)
        self.btn_select.setDefault(True)
        button_layout.addWidget(self.btn_select)
        
        layout.addLayout(button_layout)
    
    def load_materials(self):
        """Загрузить материалы для детали"""
        if not self.part_code:
            return
        
        self.all_entries = self.catalog_loader.get_entries_by_part(self.part_code)
        self.apply_filters()
    
    def apply_filters(self):
        """Применить фильтры к таблице"""
        # Получаем значения фильтров
        workshop_index = self.workshop_filter.currentIndex()
        workshop_filter = "" if workshop_index == 0 else self.workshop_filter.currentText()
        role_filter = self.role_filter.text().strip().lower()
        
        # Фильтруем записи
        filtered_entries = []
        for entry in self.all_entries:
            # Фильтр по цеху
            if workshop_filter and entry.workshop != workshop_filter:
                continue
            
            # Фильтр по типу позиции
            if role_filter and role_filter not in entry.role.lower():
                continue
            
            filtered_entries.append(entry)
        
        # Заполняем таблицу
        self.table.setRowCount(len(filtered_entries))
        
        for row, entry in enumerate(filtered_entries):
            # Чекбокс
            checkbox = QCheckBox()
            checkbox.setChecked(entry in self.selected_entries)
            self.table.setCellWidget(row, 0, checkbox)
            
            # Цех
            self.table.setItem(row, 1, QTableWidgetItem(entry.workshop))
            
            # Тип позиции
            self.table.setItem(row, 2, QTableWidgetItem(entry.role))
            
            # Наименование материала
            self.table.setItem(row, 3, QTableWidgetItem(entry.before_name))
            
            # Ед. изм.
            self.table.setItem(row, 4, QTableWidgetItem(entry.unit))
            
            # Норма
            norm_item = QTableWidgetItem(f"{entry.norm:.4f}")
            norm_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.table.setItem(row, 5, norm_item)
            
            # Примечание
            self.table.setItem(row, 6, QTableWidgetItem(entry.comment or ""))
            
            # Сохраняем entry в item для доступа
            for col in range(1, 7):
                item = self.table.item(row, col)
                if item:
                    item.setData(Qt.UserRole, entry)
    
    def accept_selection(self):
        """Принять выбор и вернуть выбранные материалы"""
        self.selected_entries = []
        
        for row in range(self.table.rowCount()):
            checkbox = self.table.cellWidget(row, 0)
            if isinstance(checkbox, QCheckBox) and checkbox.isChecked():
                # Получаем entry из любой ячейки строки
                item = self.table.item(row, 1)
                if item:
                    entry = item.data(Qt.UserRole)
                    if entry:
                        self.selected_entries.append(entry)
        
        if not self.selected_entries:
            QMessageBox.warning(self, "Предупреждение", "Выберите хотя бы один материал")
            return
        
        self.accept()
    
    def get_selected_entries(self) -> List[CatalogEntry]:
        """Получить выбранные материалы"""
        return self.selected_entries

