"""
Диалог настроек приложения
"""
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QLineEdit, QMessageBox, QFileDialog, QSpinBox, QCheckBox
)
from pathlib import Path

from app.settings_manager import SettingsManager
from app.config import PROJECT_ROOT
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
        
        # Настройка текущего номера (следующий номер, который будет использован)
        current_number_layout = QVBoxLayout()
        current_number_label = QLabel("Следующий номер извещения (текущий):")
        current_number_layout.addWidget(current_number_label)
        
        self.current_number_spin = QSpinBox()
        self.current_number_spin.setMinimum(1)
        self.current_number_spin.setMaximum(999999)
        current_number_layout.addWidget(self.current_number_spin)
        
        current_number_info = QLabel("Установите номер, с которого продолжить нумерацию. Будет применено немедленно для текущего года.")
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
        
        # Загружаем текущий номер (следующий номер, который будет использован)
        current_number = self.numbering_manager.get_current_number()
        self.current_number_spin.setValue(current_number)
        
        # Загружаем настройку автоматического открытия файла
        open_after = self.settings_manager.get_open_after_generate()
        self.open_after_checkbox.setChecked(open_after)
    
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
            
            # Сохраняем настройку автоматического открытия файла
            open_after = self.open_after_checkbox.isChecked()
            self.settings_manager.set_open_after_generate(open_after)
            
            # Устанавливаем текущий номер (если он был изменен)
            current_number = self.current_number_spin.value()
            original_current_number = self.numbering_manager.get_current_number()
            if current_number != original_current_number:
                self.numbering_manager.set_number(current_number)
                QMessageBox.information(
                    self,
                    "Номер установлен",
                    f"Следующий номер извещения установлен: {current_number}"
                )
            
            super().accept()
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить настройки:\n{e}")

