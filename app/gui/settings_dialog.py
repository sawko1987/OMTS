"""
Диалог настроек приложения
"""
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QLineEdit, QMessageBox, QFileDialog
)
from pathlib import Path

from app.settings_manager import SettingsManager
from app.config import PROJECT_ROOT


class SettingsDialog(QDialog):
    """Диалог настроек приложения"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.settings_manager = SettingsManager()
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
            super().accept()
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить настройки:\n{e}")

