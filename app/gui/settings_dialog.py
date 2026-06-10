"""
Диалог настроек приложения
"""
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QLineEdit, QMessageBox, QFileDialog, QSpinBox, QCheckBox,
    QComboBox
)
from pathlib import Path

from app.settings_manager import SettingsManager
from app.config import PROJECT_ROOT, MONTHS
from app.numbering import NumberingManager


class SettingsDialog(QDialog):
    """Диалог настроек приложения"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.settings_manager = SettingsManager()
        self.numbering_manager = NumberingManager()
        self.init_ui()
        self.load_settings()
    
    def init_ui(self):
        """Инициализация интерфейса"""
        self.setWindowTitle("Настройки")
        self.setMinimumWidth(500)
        
        layout = QVBoxLayout(self)
        
        # Настройка папки сохранения
        output_layout = QVBoxLayout()
        output_label = QLabel("Папка для сохранения извещений:")
        output_layout.addWidget(output_label)
        
        path_layout = QHBoxLayout()
        self.path_edit = QLineEdit()
        self.path_edit.setReadOnly(True)
        path_layout.addWidget(self.path_edit)
        
        self.btn_browse = QPushButton("Обзор...")
        self.btn_browse.clicked.connect(self.browse_directory)
        path_layout.addWidget(self.btn_browse)
        
        output_layout.addLayout(path_layout)
        layout.addLayout(output_layout)
        
        # Настройка начального номера
        numbering_layout = QVBoxLayout()
        numbering_label = QLabel("Начальный номер для нумерации извещений:")
        numbering_layout.addWidget(numbering_label)
        
        self.starting_number_spin = QSpinBox()
        self.starting_number_spin.setMinimum(1)
        self.starting_number_spin.setMaximum(999999)
        numbering_layout.addWidget(self.starting_number_spin)
        
        numbering_info = QLabel("Номер, с которого будет начинаться нумерация при создании новой записи для года")
        numbering_info.setStyleSheet("color: gray; font-size: 10px;")
        numbering_layout.addWidget(numbering_info)
        
        layout.addLayout(numbering_layout)
        
        # Настройка года по умолчанию
        default_year_layout = QVBoxLayout()
        default_year_label = QLabel("Год по умолчанию для нумерации:")
        default_year_layout.addWidget(default_year_label)
        
        self.default_year_spin = QSpinBox()
        current_year = 2026  # будет переопределено в load_settings
        self.default_year_spin.setRange(current_year - 10, current_year + 10)
        default_year_layout.addWidget(self.default_year_spin)
        
        default_year_info = QLabel("Год, который будет подставляться по умолчанию в поле 'Год (для нумерации)' при создании нового документа")
        default_year_info.setStyleSheet("color: gray; font-size: 10px;")
        default_year_layout.addWidget(default_year_info)
        
        layout.addLayout(default_year_layout)
        
        # Настройка текущего номера (следующий номер, который будет использован)
        current_number_layout = QVBoxLayout()
        current_number_label = QLabel("Следующий номер извещения (текущий):")
        current_number_layout.addWidget(current_number_label)
        
        # Выбор месяца для просмотра/установки номера
        month_layout = QHBoxLayout()
        month_layout.addWidget(QLabel("Месяц:"))
        self.current_month_combo = QComboBox()
        for month_num in range(1, 13):
            self.current_month_combo.addItem(MONTHS[month_num], month_num)
        self.current_month_combo.currentIndexChanged.connect(self._on_month_changed)
        month_layout.addWidget(self.current_month_combo)
        current_number_layout.addLayout(month_layout)
        
        self.current_number_spin = QSpinBox()
        self.current_number_spin.setMinimum(1)
        self.current_number_spin.setMaximum(999999)
        current_number_layout.addWidget(self.current_number_spin)
        
        current_number_info = QLabel("Установите номер, с которого продолжить нумерацию для выбранного месяца. Будет применено немедленно.")
        current_number_info.setStyleSheet("color: gray; font-size: 10px;")
        current_number_layout.addWidget(current_number_info)
        
        layout.addLayout(current_number_layout)
        
        # Настройка автоматического открытия файла после генерации
        open_after_layout = QVBoxLayout()
        self.open_after_checkbox = QCheckBox("Автоматически открывать файл после генерации")
        open_after_layout.addWidget(self.open_after_checkbox)
        
        open_after_info = QLabel("После успешной генерации документа файл будет автоматически открыт в Excel для просмотра и печати")
        open_after_info.setStyleSheet("color: gray; font-size: 10px;")
        open_after_layout.addWidget(open_after_info)
        
        layout.addLayout(open_after_layout)
        
        layout.addStretch()
        
        # Кнопки
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        self.btn_ok = QPushButton("OK")
        self.btn_ok.clicked.connect(self.accept)
        button_layout.addWidget(self.btn_ok)
        
        self.btn_cancel = QPushButton("Отмена")
        self.btn_cancel.clicked.connect(self.reject)
        button_layout.addWidget(self.btn_cancel)
        
        layout.addLayout(button_layout)
    
    def load_settings(self):
        """Загрузить настройки"""
        output_dir = self.settings_manager.get_output_directory()
        if output_dir:
            self.path_edit.setText(output_dir)
        else:
            default_path = str(PROJECT_ROOT / "output")
            self.path_edit.setText(default_path)
            self.path_edit.setPlaceholderText("Не выбрано (будет использована папка по умолчанию)")
        
        # Загружаем начальный номер
        starting_number = self.settings_manager.get_starting_number()
        self.starting_number_spin.setValue(starting_number)
        
        # Загружаем год по умолчанию
        default_year = self.settings_manager.get_default_year()
        self.default_year_spin.setValue(default_year)
        self.default_year_spin.setRange(default_year - 10, default_year + 10)
        
        # Загружаем текущий месяц и номер для него
        from datetime import date
        current_month = date.today().month
        idx = self.current_month_combo.findData(current_month)
        if idx >= 0:
            self.current_month_combo.setCurrentIndex(idx)
        current_number = self.numbering_manager.get_current_number(year=default_year, month=current_month)
        self.current_number_spin.setValue(current_number)
        
        # Загружаем настройку автоматического открытия файла
        open_after = self.settings_manager.get_open_after_generate()
        self.open_after_checkbox.setChecked(open_after)
    
    def _on_month_changed(self, index):
        """Обновить отображение текущего номера при смене месяца"""
        month = self.current_month_combo.currentData()
        year = self.default_year_spin.value()
        current_number = self.numbering_manager.get_current_number(year=year, month=month)
        self.current_number_spin.setValue(current_number)
    
    def browse_directory(self):
        """Выбрать папку для сохранения"""
        current_path = self.path_edit.text() or str(PROJECT_ROOT / "output")
        
        selected_dir = QFileDialog.getExistingDirectory(
            self,
            "Выберите папку для сохранения извещений",
            current_path
        )
        
        if selected_dir:
            self.path_edit.setText(selected_dir)
    
    def accept(self):
        """Применить настройки"""
        path = self.path_edit.text().strip()
        
        if not path:
            QMessageBox.warning(self, "Ошибка", "Укажите папку для сохранения извещений")
            return
        
        path_obj = Path(path)
        if not path_obj.exists():
            reply = QMessageBox.question(
                self,
                "Подтверждение",
                f"Папка не существует:\n{path}\n\nСоздать её?",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply == QMessageBox.Yes:
                try:
                    path_obj.mkdir(parents=True, exist_ok=True)
                except Exception as e:
                    QMessageBox.critical(self, "Ошибка", f"Не удалось создать папку:\n{e}")
                    return
            else:
                return
        
        if not path_obj.is_dir():
            QMessageBox.warning(self, "Ошибка", "Указанный путь не является папкой")
            return
        
        try:
            self.settings_manager.set_output_directory(path)
            
            # Сохраняем начальный номер
            starting_number = self.starting_number_spin.value()
            self.settings_manager.set_starting_number(starting_number)
            
            # Сохраняем год по умолчанию
            default_year = self.default_year_spin.value()
            self.settings_manager.set_default_year(default_year)
            
            # Сохраняем настройку автоматического открытия файла
            open_after = self.open_after_checkbox.isChecked()
            self.settings_manager.set_open_after_generate(open_after)
            
            # Устанавливаем текущий номер для выбранного месяца (если он был изменен)
            month = self.current_month_combo.currentData()
            current_number = self.current_number_spin.value()
            original_current_number = self.numbering_manager.get_current_number(year=default_year, month=month)
            if current_number != original_current_number:
                self.numbering_manager.set_number(current_number, year=default_year, month=month)
            
            super().accept()
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить настройки:\n{e}")

