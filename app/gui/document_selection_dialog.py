"""
Диалог выбора документа для открытия
"""
from pathlib import Path
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton,
    QTableWidget, QTableWidgetItem, QHeaderView, QLabel, QLineEdit
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
        self.all_documents = []
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

        # Поиск по документам
        search_layout = QHBoxLayout()
        search_layout.addWidget(QLabel("Поиск:"))
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Введите номер, год или часть имени файла")
        self.search_edit.textChanged.connect(self.apply_filter)
        search_layout.addWidget(self.search_edit)
        layout.addLayout(search_layout)
        
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
        self.table.itemSelectionChanged.connect(self.update_open_button_state)
        
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
        self.all_documents = self.document_store.get_all_documents()
        self.populate_table(self.all_documents)

    def populate_table(self, documents):
        """Заполнить таблицу документами"""
        self.table.setRowCount(len(documents))

        for row, doc_data_row in enumerate(documents):
            doc_number, year, month, created_at, file_path = doc_data_row
            # Номер
            number_item = QTableWidgetItem(str(doc_number))
            number_item.setData(Qt.UserRole, (doc_number, year, month))
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

        if documents:
            self.table.selectRow(0)
        else:
            self.table.clearSelection()

        self.update_open_button_state()

    def apply_filter(self):
        """Отфильтровать документы по строке поиска"""
        search_text = self.search_edit.text().strip().casefold()

        if not search_text:
            filtered_documents = self.all_documents
        else:
            filtered_documents = [
                document for document in self.all_documents
                if self.document_matches_search(document, search_text)
            ]

        self.populate_table(filtered_documents)

    def document_matches_search(self, document, search_text: str) -> bool:
        """Проверить, соответствует ли документ строке поиска"""
        doc_number, year, month, _, file_path = document
        file_name = Path(file_path).name if file_path else ""

        searchable_values = (
            str(doc_number),
            str(year),
            file_name.casefold(),
        )
        return any(search_text in value for value in searchable_values)

    def update_open_button_state(self):
        """Обновить доступность кнопки открытия"""
        self.btn_open.setEnabled(self.table.currentRow() >= 0)
    
    def get_selected_document(self):
        """Получить выбранный документ: (document_number, year, month)"""
        current_row = self.table.currentRow()
        if current_row >= 0:
            item = self.table.item(current_row, 0)
            if item:
                doc_number, year, month = item.data(Qt.UserRole)
                self.selected_document_number = doc_number
                self.selected_year = year
                return doc_number, year, month
        return None, None, None

