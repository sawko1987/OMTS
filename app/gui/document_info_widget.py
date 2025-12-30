"""
Виджет для ввода реквизитов документа
"""
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QFormLayout,
    QLineEdit, QDateEdit, QLabel
)
from PySide6.QtCore import QDate
from datetime import date

from app.models import DocumentData
from app.numbering import NumberingManager
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
    
    def __init__(self, document_data: DocumentData, product_store=None, db_manager: DatabaseManager = None):
        super().__init__()
        self.document_data = document_data
        self.db_manager = db_manager or DatabaseManager()
        self.numbering = NumberingManager(self.db_manager)
        self.init_ui()
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
        
        # Причина
        self.reason_edit = QLineEdit()
        self.reason_edit.setPlaceholderText("Укажите причину замены материалов")
        form_layout.addRow("Причина:", self.reason_edit)
        
        layout.addLayout(form_layout)
        layout.addStretch()
    
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
        
        # Причина
        self.reason_edit.setText(self.document_data.reason)
    
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
        self.document_data.reason = self.reason_edit.text().strip()

