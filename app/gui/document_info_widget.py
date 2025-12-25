"""
Виджет для ввода реквизитов документа
"""
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
    QLineEdit, QDateEdit, QLabel, QComboBox, QPushButton, QMessageBox, QInputDialog
)
from PySide6.QtCore import QDate, Qt, Signal
from datetime import date

from app.models import DocumentData
from app.numbering import NumberingManager
from app.product_store import ProductStore
from app.database import DatabaseManager


def get_current_month_name() -> str:
    """Получить название текущего месяца на русском языке"""
    months = {
        1: "Январь", 2: "Февраль", 3: "Март", 4: "Апрель",
        5: "Май", 6: "Июнь", 7: "Июль", 8: "Август",
        9: "Сентябрь", 10: "Октябрь", 11: "Ноябрь", 12: "Декабрь"
    }
    return months.get(date.today().month, "Неизвестно")


class DocumentInfoWidget(QWidget):
    """Виджет реквизитов документа"""
    
    product_changed = Signal(str)  # Сигнал об изменении изделия
    
    def __init__(self, document_data: DocumentData, product_store: ProductStore = None, db_manager: DatabaseManager = None):
        super().__init__()
        self.document_data = document_data
        self.db_manager = db_manager or DatabaseManager()
        self.numbering = NumberingManager(self.db_manager)
        self.product_store = product_store
        self.init_ui()
        self.load_products()
        self.refresh()
    
    def init_ui(self):
        """Инициализация интерфейса"""
        layout = QVBoxLayout(self)
        
        # Заголовок
        title = QLabel("Реквизиты документа")
        title.setStyleSheet("font-size: 16px; font-weight: bold;")
        layout.addWidget(title)
        
        # Форма
        form_layout = QFormLayout()
        
        # Номер документа (автоматический)
        self.number_label = QLabel()
        form_layout.addRow("Номер извещения:", self.number_label)
        
        # Дата внедрения замены
        self.date_edit = QDateEdit()
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDate(QDate.currentDate())
        form_layout.addRow("Дата внедрения замены:", self.date_edit)
        
        # Срок действия (партия)
        self.validity_edit = QLineEdit()
        self.validity_edit.setPlaceholderText("например: ноябрь")
        form_layout.addRow("Срок действия (партия):", self.validity_edit)
        
        # Изделие
        product_layout = QHBoxLayout()
        self.product_combo = QComboBox()
        self.product_combo.setEditable(True)
        self.product_combo.setMinimumWidth(200)
        self.product_combo.currentTextChanged.connect(self.on_product_changed)
        product_layout.addWidget(self.product_combo)
        
        self.btn_create_product = QPushButton("Создать")
        self.btn_create_product.clicked.connect(self.create_product)
        product_layout.addWidget(self.btn_create_product)
        
        product_widget = QWidget()
        product_widget.setLayout(product_layout)
        form_layout.addRow("Изделие:", product_widget)
        
        # Причина
        self.reason_edit = QLineEdit()
        self.reason_edit.setPlaceholderText("Укажите причину замены материалов")
        form_layout.addRow("Причина:", self.reason_edit)
        
        # Заключение ТКО
        self.tko_edit = QLineEdit()
        self.tko_edit.setPlaceholderText("допускается/не допускается")
        form_layout.addRow("Заключение ТКО:", self.tko_edit)
        
        layout.addLayout(form_layout)
        layout.addStretch()
    
    def load_products(self):
        """Загрузить список изделий из БД"""
        if not self.product_store:
            return
        
        self.product_combo.clear()
        products = self.product_store.get_all_products()
        for product_id, product_name in products:
            self.product_combo.addItem(product_name, product_id)
    
    def create_product(self):
        """Создать новое изделие"""
        if not self.product_store:
            QMessageBox.warning(self, "Ошибка", "ProductStore не инициализирован")
            return
        
        current_text = self.product_combo.currentText().strip()
        
        if not current_text:
            text, ok = QInputDialog.getText(
                self, "Создать изделие", "Введите название изделия:"
            )
            if not ok or not text.strip():
                return
            product_name = text.strip()
        else:
            product_name = current_text
        
        # Проверяем, не существует ли уже такое изделие
        existing_id = self.product_store.get_product_by_name(product_name)
        if existing_id:
            QMessageBox.information(
                self, "Информация", 
                f"Изделие '{product_name}' уже существует"
            )
            # Устанавливаем его в комбобоксе
            index = self.product_combo.findText(product_name)
            if index >= 0:
                self.product_combo.setCurrentIndex(index)
            return
        
        # Создаём новое изделие
        product_id = self.product_store.add_product(product_name)
        if product_id:
            # Обновляем список
            self.load_products()
            # Устанавливаем новое изделие
            index = self.product_combo.findText(product_name)
            if index >= 0:
                self.product_combo.setCurrentIndex(index)
            QMessageBox.information(
                self, "Успех", 
                f"Изделие '{product_name}' успешно создано"
            )
        else:
            QMessageBox.warning(self, "Ошибка", "Не удалось создать изделие")
    
    def on_product_changed(self, text: str):
        """Обработка изменения изделия"""
        self.product_changed.emit(text.strip())
    
    def refresh(self):
        """Обновить отображение"""
        # Номер документа - если уже установлен, используем его, иначе берём следующий
        if self.document_data.document_number:
            self.number_label.setText(str(self.document_data.document_number))
        else:
            next_num = self.numbering.get_current_number()
            self.number_label.setText(str(next_num))
            self.document_data.document_number = next_num
        
        # Дата внедрения
        if self.document_data.implementation_date:
            impl_date = self.document_data.implementation_date
            qdate = QDate(impl_date.year, impl_date.month, impl_date.day)
            self.date_edit.setDate(qdate)
        else:
            self.date_edit.setDate(QDate.currentDate())
        
        # Срок действия (партия)
        if self.document_data.validity_period:
            self.validity_edit.setText(self.document_data.validity_period)
        else:
            current_month = get_current_month_name()
            self.validity_edit.setText(current_month)
            self.document_data.validity_period = current_month
        
        # Изделие
        if self.document_data.product:
            index = self.product_combo.findText(self.document_data.product)
            if index >= 0:
                self.product_combo.setCurrentIndex(index)
            else:
                self.product_combo.setEditText(self.document_data.product)
        
        # Причина
        self.reason_edit.setText(self.document_data.reason)
        
        # Заключение ТКО
        self.tko_edit.setText(self.document_data.tko_conclusion)
    
    def refresh_number(self):
        """Обновить только номер документа"""
        # Пересоздаем NumberingManager, чтобы гарантировать чтение актуальных данных
        self.numbering = NumberingManager(self.db_manager)
        next_num = self.numbering.get_current_number()
        self.number_label.setText(str(next_num))
        # Обновляем document_number, чтобы следующий документ использовал новый номер
        self.document_data.document_number = next_num
    
    def update_document_data(self):
        """Обновить данные документа из полей"""
        # Номер (уже установлен при refresh)
        
        # Дата внедрения
        qdate = self.date_edit.date()
        self.document_data.implementation_date = qdate.toPython()
        
        # Остальные поля
        self.document_data.validity_period = self.validity_edit.text().strip() or None
        self.document_data.product = self.product_combo.currentText().strip()
        self.document_data.reason = self.reason_edit.text().strip()
        self.document_data.tko_conclusion = self.tko_edit.text().strip()

