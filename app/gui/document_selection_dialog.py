"""
Диалог выбора документа для открытия
"""
from pathlib import Path
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton,
    QTableWidget, QTableWidgetItem, QHeaderView, QLabel
)
from PySide6.QtCore import Qt
from datetime import datetime

from app.document_store import DocumentStore


class DocumentSelectionDialog(QDialog):
    """Диалог выбора документа для открытия"""
    
    def __init__(self, document_store: DocumentStore, parent=None):
        super().__init__(parent)
        self.document_store = document_store
        self.selected_document_number = None
        self.selected_year = None
        self.setWindowTitle("Открыть документ")
        self.setMinimumSize(500, 400)
        self.init_ui()
        self.load_documents()
    
    def init_ui(self):
        """Инициализация интерфейса"""
        layout = QVBoxLayout(self)
        
        # Заголовок
        title = QLabel("Выберите документ для открытия:")
        title.setStyleSheet("font-size: 14px; font-weight: bold;")
        layout.addWidget(title)
        
        # Таблица документов
        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels([
            "Номер",
            "Год",
            "Дата создания",
            "Файл"
        ])
        
        # Настройка таблицы
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.Stretch)
        
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.doubleClicked.connect(self.accept)
        
        layout.addWidget(self.table)
        
        # Кнопки
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        self.btn_cancel = QPushButton("Отмена")
        self.btn_cancel.clicked.connect(self.reject)
        button_layout.addWidget(self.btn_cancel)
        
        self.btn_open = QPushButton("Открыть")
        self.btn_open.clicked.connect(self.accept)
        self.btn_open.setDefault(True)
        button_layout.addWidget(self.btn_open)
        
        layout.addLayout(button_layout)
    
    def load_documents(self):
        """Загрузить список документов"""
        documents = self.document_store.get_all_documents()
        
        self.table.setRowCount(len(documents))
        
        for row, (doc_number, year, created_at, file_path) in enumerate(documents):
            # Номер
            number_item = QTableWidgetItem(str(doc_number))
            number_item.setData(Qt.UserRole, (doc_number, year))
            self.table.setItem(row, 0, number_item)
            
            # Год
            year_item = QTableWidgetItem(str(year))
            self.table.setItem(row, 1, year_item)
            
            # Дата создания
            if created_at:
                try:
                    if isinstance(created_at, str):
                        # Парсим строку даты
                        dt = datetime.fromisoformat(created_at.replace('Z', '+00:00'))
                        date_str = dt.strftime("%d.%m.%Y %H:%M")
                    else:
                        date_str = str(created_at)
                except:
                    date_str = str(created_at)
            else:
                date_str = ""
            date_item = QTableWidgetItem(date_str)
            self.table.setItem(row, 2, date_item)
            
            # Файл
            file_name = Path(file_path).name if file_path else ""
            file_item = QTableWidgetItem(file_name)
            self.table.setItem(row, 3, file_item)
        
        # Сортируем по номеру (уже отсортировано в запросе)
        if documents:
            self.table.selectRow(0)
    
    def get_selected_document(self):
        """Получить выбранный документ: (document_number, year)"""
        current_row = self.table.currentRow()
        if current_row >= 0:
            item = self.table.item(current_row, 0)
            if item:
                doc_number, year = item.data(Qt.UserRole)
                self.selected_document_number = doc_number
                self.selected_year = year
                return doc_number, year
        return None, None

