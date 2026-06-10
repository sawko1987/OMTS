"""
Виджет для ввода реквизитов документа
"""
import logging
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QFormLayout,
    QLineEdit, QDateEdit, QLabel, QComboBox, QSpinBox
)
from PySide6.QtCore import QDate
from datetime import date

from app.models import DocumentData
from app.numbering import NumberingManager
from app.database import DatabaseManager
from app.settings_manager import SettingsManager
from app.config import MONTHS

logger = logging.getLogger(__name__)


class DocumentInfoWidget(QWidget):
    """Виджет реквизитов документа"""
    
    def __init__(self, document_data: DocumentData, product_store=None, db_manager: DatabaseManager = None):
        super().__init__()
        self.document_data = document_data
        self.db_manager = db_manager or DatabaseManager()
        self.numbering = NumberingManager(self.db_manager)
        self.settings_manager = SettingsManager()
        self._number_year = None  # Год, для которого был установлен текущий номер
        self._number_month = None  # Месяц, для которого был установлен текущий номер
        self._updating = False  # Флаг для предотвращения рекурсивных вызовов
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
        
        # Номер документа (автоматический, формат ММ-ГГ-№)
        self.number_label = QLabel()
        self.number_label.setStyleSheet("font-size: 14px; font-weight: bold; color: #2a6da8;")
        form_layout.addRow("Номер извещения:", self.number_label)
        
        # Дата внедрения замены
        self.date_edit = QDateEdit()
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDate(QDate.currentDate())
        # Подключаем обработчик изменения даты для автоматического пересчета номера
        self.date_edit.dateChanged.connect(self._on_date_changed)
        form_layout.addRow("Дата внедрения замены:", self.date_edit)
        
        # Срок действия (партия) - выпадающий список месяцев
        self.validity_combo = QComboBox()
        for month_num in range(1, 13):
            self.validity_combo.addItem(MONTHS[month_num], month_num)
        self.validity_combo.currentIndexChanged.connect(self._on_month_changed)
        form_layout.addRow("Срок действия (партия):", self.validity_combo)
        
        # Год для нумерации (ГГ)
        self.year_spin = QSpinBox()
        current_year = date.today().year
        self.year_spin.setRange(current_year - 10, current_year + 10)
        self.year_spin.setValue(self.settings_manager.get_default_year())
        self.year_spin.valueChanged.connect(self._on_year_changed)
        form_layout.addRow("Год (для нумерации):", self.year_spin)
        
        # Причина
        self.reason_edit = QLineEdit()
        self.reason_edit.setPlaceholderText("Укажите причину замены материалов")
        form_layout.addRow("Причина:", self.reason_edit)
        
        layout.addLayout(form_layout)
        layout.addStretch()
    
    def _get_current_month_year(self) -> tuple:
        """Получить текущие месяц (1-12) и год из полей виджета"""
        month = self.validity_combo.currentData()
        year = self.year_spin.value()
        return month, year
    
    def refresh(self):
        """Обновить отображение"""
        logger.info(f"[refresh] Начало обновления. document_number={self.document_data.document_number}, _number_year={self._number_year}, _number_month={self._number_month}")
        
        self._updating = True
        
        # Дата внедрения (устанавливаем первой, чтобы номер рассчитывался на её основе)
        if self.document_data.implementation_date:
            impl_date = self.document_data.implementation_date
            qdate = QDate(impl_date.year, impl_date.month, impl_date.day)
            self.date_edit.setDate(qdate)
            logger.info(f"[refresh] Дата внедрения установлена: {impl_date}")
        else:
            self.date_edit.setDate(QDate.currentDate())
            logger.info(f"[refresh] Дата внедрения не установлена, используется текущая дата")
        
        # Срок действия (партия) - устанавливаем месяц в комбобоксе
        if self.document_data.document_month is not None:
            month = self.document_data.document_month
            idx = self.validity_combo.findData(month)
            if idx >= 0:
                self.validity_combo.setCurrentIndex(idx)
        elif self.document_data.validity_period:
            # Пытаемся найти месяц по названию (для обратной совместимости)
            month_name = self.document_data.validity_period
            for num, name in MONTHS.items():
                if name.lower() == month_name.lower():
                    idx = self.validity_combo.findData(num)
                    if idx >= 0:
                        self.validity_combo.setCurrentIndex(idx)
                    self.document_data.document_month = num
                    break
        else:
            # Устанавливаем текущий месяц
            current_month = date.today().month
            idx = self.validity_combo.findData(current_month)
            if idx >= 0:
                self.validity_combo.setCurrentIndex(idx)
            self.document_data.document_month = current_month
            self.document_data.validity_period = MONTHS[current_month]
        
        # Год для нумерации
        if self.document_data.document_year is not None:
            self.year_spin.setValue(self.document_data.document_year)
        else:
            default_year = self.settings_manager.get_default_year()
            self.year_spin.setValue(default_year)
            self.document_data.document_year = default_year
        
        self._updating = False
        
        # Для загруженного из БД документа — сохраняем его месяц/год,
        # чтобы _update_number_display() не перезаписал номер счётчиком
        if self.document_data.document_number is not None:
            if self.document_data.document_month is not None:
                self._number_month = self.document_data.document_month
            if self.document_data.document_year is not None:
                self._number_year = self.document_data.document_year
        
        # Номер документа
        self._update_number_display()
        
        # Причина
        self.reason_edit.setText(self.document_data.reason)
    
    def _update_number_display(self):
        """Обновить отображение номера документа в формате ММ-ГГ-№"""
        month, year = self._get_current_month_year()
        
        # Для существующих документов (загруженных из БД) сохраняем текущий номер
        if self.document_data.document_number is not None and self._number_month is not None and self._number_year is not None:
            # Показываем номер в формате ММ-ГГ-№
            display = NumberingManager.format_number(month, year, self.document_data.document_number)
            self.number_label.setText(display)
            logger.info(f"[_update_number_display] Существующий номер: {display}")
        else:
            # Для нового документа - получаем текущий номер для месяца+года
            next_num = self.numbering.get_current_number(year, month)
            self.document_data.document_number = next_num
            self.document_data.document_month = month
            self.document_data.document_year = year
            self._number_month = month
            self._number_year = year
            display = NumberingManager.format_number(month, year, next_num)
            self.number_label.setText(display)
            logger.info(f"[_update_number_display] Новый номер: {display}")
    
    def refresh_number(self):
        """Обновить только номер документа"""
        logger.info(f"[refresh_number] Начало. document_number={self.document_data.document_number}, _number_year={self._number_year}, _number_month={self._number_month}")
        
        # Пересоздаем NumberingManager, чтобы гарантировать чтение актуальных данных
        self.numbering = NumberingManager(self.db_manager)
        month, year = self._get_current_month_year()
        
        next_num = self.numbering.get_current_number(year, month)
        self.document_data.document_number = next_num
        self.document_data.document_month = month
        self.document_data.document_year = year
        self._number_month = month
        self._number_year = year
        
        display = NumberingManager.format_number(month, year, next_num)
        self.number_label.setText(display)
        
        logger.info(f"[refresh_number] Обновлен номер: {display}")
    
    def _on_date_changed(self, new_date: QDate):
        """Обработчик изменения даты внедрения"""
        logger.info(f"[_on_date_changed] Вызван. Новая дата: {new_date.toPython()}")
        
        # Обновляем дату в document_data
        self.document_data.implementation_date = new_date.toPython()
        
        # Не пересчитываем номер при изменении даты (номер привязан к месяцу+году из выпадающих списков)
        if not self._updating:
            self._update_number_display()
    
    def _on_month_changed(self, index):
        """Обработчик изменения месяца в выпадающем списке"""
        if self._updating:
            return
        
        month = self.validity_combo.currentData()
        self.document_data.document_month = month
        self.document_data.validity_period = MONTHS[month]
        
        logger.info(f"[_on_month_changed] Выбран месяц: {month} ({MONTHS[month]})")
        
        # Если у документа нет «родного» месяца (новый документ, месяц не совпадает) —
        # перезапрашиваем номер для нового месяца
        if (self.document_data.document_number is not None
                and self._number_month is not None
                and self._number_month == month):
            # Месяц не изменился относительно сохранённого — просто обновить отображение
            self._update_number_display()
        else:
            # Месяц изменился или это новый документ — получить номер для нового месяца
            self.refresh_number()
    
    def _on_year_changed(self, year):
        """Обработчик изменения года"""
        if self._updating:
            return
        
        self.document_data.document_year = year
        logger.info(f"[_on_year_changed] Выбран год: {year}")
        
        # Сохраняем выбранный год в настройки
        self.settings_manager.set_default_year(year)
        
        # Если год изменился — перезапрашиваем номер для нового года
        if (self.document_data.document_number is not None
                and self._number_year is not None
                and self._number_year == year):
            self._update_number_display()
        else:
            self.refresh_number()
    
    def update_document_data(self):
        """Обновить данные документа из полей"""
        logger.info(f"[update_document_data] Начало. document_number={self.document_data.document_number}")
        
        month, year = self._get_current_month_year()
        
        # Дата внедрения
        qdate = self.date_edit.date()
        self.document_data.implementation_date = qdate.toPython()
        
        # Месяц и год для нумерации
        self.document_data.document_month = month
        self.document_data.document_year = year
        self.document_data.validity_period = MONTHS.get(month)
        
        # Номер документа - сохраняем существующий номер, если он установлен
        if not self.document_data.document_number:
            # Новый документ - получаем номер
            next_num = self.numbering.get_current_number(year, month)
            self.document_data.document_number = next_num
            self._number_month = month
            self._number_year = year
            display = NumberingManager.format_number(month, year, next_num)
            self.number_label.setText(display)
            logger.info(f"[update_document_data] Установлен номер: {display}")
        else:
            # Существующий документ - сохраняем текущий номер
            display = NumberingManager.format_number(
                self._number_month or month,
                self._number_year or year,
                self.document_data.document_number
            )
            self.number_label.setText(display)
            logger.info(f"[update_document_data] Номер уже установлен: {display}")
        
        # Причина
        self.document_data.reason = self.reason_edit.text().strip()
