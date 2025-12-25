"""
Диалог выбора набора материалов для замены
"""
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton,
    QTableWidget, QTableWidgetItem, QHeaderView, QLabel, QMessageBox
)
from PySide6.QtCore import Qt
from typing import Optional, List

from app.models import MaterialReplacementSet


class ReplacementSetSelectionDialog(QDialog):
    """Диалог для выбора набора материалов 'to' из нескольких вариантов"""
    
    def __init__(self, sets: List[MaterialReplacementSet], part_code: str, parent=None):
        super().__init__(parent)
        self.sets = sets
        self.part_code = part_code
        self.selected_set: Optional[MaterialReplacementSet] = None
        self.setWindowTitle(f"Выбор набора материалов для детали {part_code}")
        self.setMinimumSize(700, 500)
        self.init_ui()
    
    def init_ui(self):
        """Инициализация интерфейса"""
        layout = QVBoxLayout(self)
        
        # Заголовок
        header_label = QLabel(f"Для детали '{self.part_code}' найдено {len(self.sets)} наборов 'после'.\nВыберите один набор:")
        header_label.setStyleSheet("font-weight: bold; font-size: 11pt;")
        layout.addWidget(header_label)
        
        # Таблица наборов
        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels([
            "ID", "Название", "Материалы", "Дата создания"
        ])
        
        # Настройка таблицы
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.Stretch)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        
        # Заполняем таблицу
        self.table.setRowCount(len(self.sets))
        for row, replacement_set in enumerate(self.sets):
            # ID
            id_item = QTableWidgetItem(str(replacement_set.id or ""))
            id_item.setData(Qt.UserRole, replacement_set)
            self.table.setItem(row, 0, id_item)
            
            # Название
            name = replacement_set.set_name or f"Набор #{replacement_set.id}"
            name_item = QTableWidgetItem(name)
            self.table.setItem(row, 1, name_item)
            
            # Материалы (краткая информация)
            materials_text = self._format_materials(replacement_set.materials)
            materials_item = QTableWidgetItem(materials_text)
            self.table.setItem(row, 2, materials_item)
            
            # Дата создания
            date_text = ""
            if replacement_set.created_at:
                if isinstance(replacement_set.created_at, str):
                    date_text = replacement_set.created_at
                else:
                    date_text = replacement_set.created_at.strftime("%Y-%m-%d %H:%M")
            date_item = QTableWidgetItem(date_text)
            self.table.setItem(row, 3, date_item)
        
        # Выбираем первую строку по умолчанию
        if len(self.sets) > 0:
            self.table.selectRow(0)
        
        layout.addWidget(self.table)
        
        # Кнопки
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        btn_cancel = QPushButton("Отмена")
        btn_cancel.clicked.connect(self.reject)
        button_layout.addWidget(btn_cancel)
        
        btn_ok = QPushButton("Выбрать")
        btn_ok.clicked.connect(self.accept_selection)
        btn_ok.setDefault(True)
        button_layout.addWidget(btn_ok)
        
        layout.addLayout(button_layout)
    
    def _format_materials(self, materials: List) -> str:
        """Форматировать список материалов для отображения"""
        if not materials:
            return "Нет материалов"
        
        # Показываем первые 3 материала, если больше - добавляем "..."
        parts = []
        for mat in materials[:3]:
            parts.append(f"{mat.workshop}/{mat.role}: {mat.before_name[:30]}")
        
        if len(materials) > 3:
            parts.append(f"... и ещё {len(materials) - 3}")
        
        return "; ".join(parts)
    
    def accept_selection(self):
        """Принять выбор набора"""
        current_row = self.table.currentRow()
        if current_row < 0:
            QMessageBox.warning(self, "Ошибка", "Выберите набор из списка")
            return
        
        id_item = self.table.item(current_row, 0)
        if id_item:
            self.selected_set = id_item.data(Qt.UserRole)
            self.accept()
        else:
            QMessageBox.warning(self, "Ошибка", "Не удалось определить выбранный набор")
    
    def get_selected_set(self) -> Optional[MaterialReplacementSet]:
        """Получить выбранный набор"""
        return self.selected_set

