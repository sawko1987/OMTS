"""
Виджет для ввода реквизитов документа
"""
import logging
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QFormLayout,
    QLineEdit, QDateEdit, QLabel
)
from PySide6.QtCore import QDate
from datetime import date

from app.models import DocumentData
from app.numbering import NumberingManager
from app.database import DatabaseManager

logger = logging.getLogger(__name__)


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
        self._number_year = None  # Год, для которого был установлен текущий номер
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
        # Подключаем обработчик изменения даты для автоматического пересчета номера
        self.date_edit.dateChanged.connect(self._on_date_changed)
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
        logger.info(f"[refresh] Начало обновления. document_number={self.document_data.document_number}, _number_year={self._number_year}")
        
        # Дата внедрения (устанавливаем первой, чтобы номер рассчитывался на её основе)
        if self.document_data.implementation_date:
            impl_date = self.document_data.implementation_date
            qdate = QDate(impl_date.year, impl_date.month, impl_date.day)
            self.date_edit.setDate(qdate)
            logger.info(f"[refresh] Дата внедрения установлена: {impl_date}")
        else:
            self.date_edit.setDate(QDate.currentDate())
            logger.info(f"[refresh] Дата внедрения не установлена, используется текущая дата")
        
        # Номер документа - если уже установлен, используем его и сохраняем
        # Для существующих документов номер не должен изменяться
        if self.document_data.document_number:
            # Сохраняем существующий номер документа
            self.number_label.setText(str(self.document_data.document_number))
            # Сохраняем год, для которого был установлен номер (из даты внедрения)
            impl_date = self.document_data.implementation_date
            self._number_year = impl_date.year if impl_date else None
            logger.info(f"[refresh] Используется существующий номер: {self.document_data.document_number}, год: {self._number_year}")
        else:
            # Устанавливаем новый номер только для новых документов
            # Получаем год из даты внедрения или текущий год
            impl_date = self.document_data.implementation_date
            if impl_date:
                year = impl_date.year
            else:
                # Если дата не установлена, используем год из текущей даты виджета
                year = self.date_edit.date().year() if self.date_edit.date().isValid() else date.today().year
            next_num = self.numbering.get_current_number(year)
            self.number_label.setText(str(next_num))
            self.document_data.document_number = next_num
            self._number_year = year  # Сохраняем год, для которого установлен номер
            logger.info(f"[refresh] Установлен новый номер: {next_num} для года: {year}")
        
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
        logger.info(f"[refresh_number] Начало. document_number={self.document_data.document_number}, _number_year={self._number_year}")
        
        # Пересоздаем NumberingManager, чтобы гарантировать чтение актуальных данных
        self.numbering = NumberingManager(self.db_manager)
        # Получаем год из даты внедрения или текущий год
        impl_date = self.document_data.implementation_date
        year = impl_date.year if impl_date else None
        next_num = self.numbering.get_current_number(year)
        self.number_label.setText(str(next_num))
        # Обновляем document_number, чтобы следующий документ использовал новый номер
        self.document_data.document_number = next_num
        self._number_year = year  # Сохраняем год, для которого установлен номер
        
        logger.info(f"[refresh_number] Обновлен номер: {next_num} для года: {year}")
    
    def _on_date_changed(self, new_date: QDate):
        """Обработчик изменения даты внедрения - автоматически пересчитывает номер документа"""
        logger.info(f"[_on_date_changed] Вызван. Новая дата: {new_date.toPython()}")
        
        # Обновляем дату в document_data
        self.document_data.implementation_date = new_date.toPython()
        
        new_year = new_date.year()
        
        logger.info(f"[_on_date_changed] Текущий номер: {self.document_data.document_number}, "
                   f"текущий год (_number_year): {self._number_year}, новый год: {new_year}")
        
        # Пересчитываем номер документа на основе нового года
        # Если номер еще не установлен или год изменился - пересчитываем номер
        # Для существующих документов (загруженных из БД) сохраняем текущий номер
        if self.document_data.document_number is None:
            # Новый документ - получаем номер для нового года
            logger.info(f"[_on_date_changed] Условие 1: номер не установлен, получаем номер для года {new_year}")
            next_num = self.numbering.get_current_number(new_year)
            self.number_label.setText(str(next_num))
            self.document_data.document_number = next_num
            self._number_year = new_year
            logger.info(f"[_on_date_changed] Установлен номер: {next_num} для года: {new_year}")
        elif self._number_year is None:
            # _number_year не установлен (новый документ, но номер уже был установлен в refresh)
            # Пересчитываем номер для нового года
            logger.info(f"[_on_date_changed] Условие 2: _number_year не установлен (новый документ), пересчитываем номер для года {new_year}")
            next_num = self.numbering.get_current_number(new_year)
            self.number_label.setText(str(next_num))
            self.document_data.document_number = next_num
            self._number_year = new_year
            logger.info(f"[_on_date_changed] Обновлен номер: {next_num} для года: {new_year}")
        elif self._number_year != new_year:
            # Год изменился - пересчитываем номер для нового года
            # Это означает, что пользователь изменил год для нового документа
            logger.info(f"[_on_date_changed] Условие 3: год изменился ({self._number_year} -> {new_year}), пересчитываем номер")
            next_num = self.numbering.get_current_number(new_year)
            logger.info(f"[_on_date_changed] Получен номер {next_num} для года {new_year}, обновляем отображение")
            old_text = self.number_label.text()
            self.number_label.setText(str(next_num))
            self.number_label.update()  # Принудительное обновление виджета
            new_text = self.number_label.text()
            logger.info(f"[_on_date_changed] Текст в number_label изменен: '{old_text}' -> '{new_text}'")
            self.document_data.document_number = next_num
            self._number_year = new_year
            logger.info(f"[_on_date_changed] Обновлен номер: {next_num} для года: {new_year}, _number_year установлен в {self._number_year}")
        else:
            # Год не изменился - сохраняем текущий номер
            # Просто обновляем отображение, не меняя номер
            logger.info(f"[_on_date_changed] Условие 4: год не изменился ({self._number_year}). "
                       f"Сохраняем номер: {self.document_data.document_number}")
            self.number_label.setText(str(self.document_data.document_number))
    
    def update_document_data(self):
        """Обновить данные документа из полей"""
        logger.info(f"[update_document_data] Начало. document_number={self.document_data.document_number}, _number_year={self._number_year}")
        
        # Дата внедрения
        qdate = self.date_edit.date()
        self.document_data.implementation_date = qdate.toPython()
        year = qdate.year() if qdate.isValid() else None
        logger.info(f"[update_document_data] Дата: {self.document_data.implementation_date}, год: {year}")
        
        # Номер документа - сохраняем существующий номер, если он установлен
        # Устанавливаем новый номер только для новых документов (когда номер не установлен)
        if not self.document_data.document_number:
            # Новый документ - получаем номер на основе года из даты
            logger.info(f"[update_document_data] Номер не установлен, получаем номер для года {year}")
            next_num = self.numbering.get_current_number(year)
            self.document_data.document_number = next_num
            self.number_label.setText(str(next_num))
            self._number_year = year  # Сохраняем год, для которого установлен номер
            logger.info(f"[update_document_data] Установлен номер: {next_num} для года: {year}")
        else:
            # Существующий документ - сохраняем текущий номер
            # Убеждаемся, что отображение соответствует сохраненному номеру
            logger.info(f"[update_document_data] Номер уже установлен: {self.document_data.document_number}, сохраняем его")
            self.number_label.setText(str(self.document_data.document_number))
        
        # Остальные поля
        self.document_data.validity_period = self.validity_edit.text().strip() or None
        self.document_data.reason = self.reason_edit.text().strip()

