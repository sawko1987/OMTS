"""
Диалог восстановления базы данных из резервной копии
"""
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton,
    QListWidget, QListWidgetItem, QLabel, QMessageBox,
    QFileDialog
)
from PySide6.QtCore import Qt
from pathlib import Path

from app.database_restore import (
    list_backup_files, get_backup_timestamp,
    restore_database, delete_backup
)
from app.config import DATABASE_PATH, DATA_DIR


class DatabaseRestoreDialog(QDialog):
    """Диалог для восстановления базы данных из резервной копии"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Восстановление базы данных")
        self.setMinimumSize(600, 500)
        self.selected_backup: Path = None
        self.init_ui()
        self.load_backups()
    
    def init_ui(self):
        """Инициализация интерфейса"""
        layout = QVBoxLayout(self)
        
        # Информация
        info_label = QLabel(
            "Выберите резервную копию базы данных для восстановления.\n"
            "Текущая БД будет сохранена перед восстановлением."
        )
        info_label.setStyleSheet("color: gray; font-size: 10px;")
        layout.addWidget(info_label)
        
        # Список резервных копий
        list_label = QLabel("Доступные резервные копии:")
        layout.addWidget(list_label)
        
        self.backups_list = QListWidget()
        self.backups_list.setSelectionMode(QListWidget.SingleSelection)
        self.backups_list.itemSelectionChanged.connect(self.on_selection_changed)
        layout.addWidget(self.backups_list)
        
        # Кнопки для работы с резервными копиями
        backup_buttons = QHBoxLayout()
        
        btn_refresh = QPushButton("Обновить список")
        btn_refresh.clicked.connect(self.load_backups)
        backup_buttons.addWidget(btn_refresh)
        
        btn_browse = QPushButton("Выбрать файл...")
        btn_browse.clicked.connect(self.browse_backup_file)
        backup_buttons.addWidget(btn_browse)
        
        btn_delete = QPushButton("Удалить выбранную")
        btn_delete.clicked.connect(self.delete_selected_backup)
        backup_buttons.addWidget(btn_delete)
        
        backup_buttons.addStretch()
        layout.addLayout(backup_buttons)
        
        # Информация о выбранной копии
        self.selected_info = QLabel("Выберите резервную копию из списка")
        self.selected_info.setStyleSheet("font-weight: bold;")
        layout.addWidget(self.selected_info)
        
        # Кнопки действий
        buttons_layout = QHBoxLayout()
        buttons_layout.addStretch()
        
        self.btn_restore = QPushButton("Восстановить")
        self.btn_restore.setEnabled(False)
        self.btn_restore.clicked.connect(self.restore_selected)
        buttons_layout.addWidget(self.btn_restore)
        
        btn_cancel = QPushButton("Отмена")
        btn_cancel.clicked.connect(self.reject)
        buttons_layout.addWidget(btn_cancel)
        
        layout.addLayout(buttons_layout)
    
    def load_backups(self):
        """Загрузить список резервных копий"""
        self.backups_list.clear()
        backups = list_backup_files()
        
        if not backups:
            item = QListWidgetItem("Резервные копии не найдены")
            item.setFlags(Qt.NoItemFlags)
            self.backups_list.addItem(item)
            return
        
        for backup_path in backups:
            timestamp = get_backup_timestamp(backup_path)
            size = backup_path.stat().st_size / 1024  # Размер в KB
            
            item_text = f"{timestamp} ({size:.1f} KB)"
            item = QListWidgetItem(item_text)
            item.setData(Qt.UserRole, backup_path)
            self.backups_list.addItem(item)
        
        # Выбираем первую копию по умолчанию
        if backups:
            self.backups_list.setCurrentRow(0)
    
    def on_selection_changed(self):
        """Обработка изменения выбора"""
        current_item = self.backups_list.currentItem()
        if current_item and current_item.data(Qt.UserRole):
            backup_path = current_item.data(Qt.UserRole)
            self.selected_backup = backup_path
            timestamp = get_backup_timestamp(backup_path)
            size = backup_path.stat().st_size / 1024
            
            self.selected_info.setText(
                f"Выбрана резервная копия: {timestamp}\n"
                f"Размер: {size:.1f} KB\n"
                f"Путь: {backup_path.name}"
            )
            self.btn_restore.setEnabled(True)
        else:
            self.selected_backup = None
            self.selected_info.setText("Выберите резервную копию из списка")
            self.btn_restore.setEnabled(False)
    
    def browse_backup_file(self):
        """Выбрать файл резервной копии вручную"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите файл резервной копии",
            str(DATA_DIR),
            "Backup files (*.bak-*);;All files (*)"
        )
        
        if file_path:
            backup_path = Path(file_path)
            if restore_database(backup_path):
                QMessageBox.information(
                    self,
                    "Успех",
                    f"База данных восстановлена из файла:\n{backup_path.name}\n\n"
                    "Перезапустите приложение для применения изменений."
                )
                self.accept()
            else:
                QMessageBox.warning(
                    self,
                    "Ошибка",
                    f"Не удалось восстановить базу данных из файла:\n{backup_path}"
                )
    
    def delete_selected_backup(self):
        """Удалить выбранную резервную копию"""
        current_item = self.backups_list.currentItem()
        if not current_item or not current_item.data(Qt.UserRole):
            QMessageBox.warning(self, "Ошибка", "Выберите резервную копию для удаления")
            return
        
        backup_path = current_item.data(Qt.UserRole)
        timestamp = get_backup_timestamp(backup_path)
        
        reply = QMessageBox.question(
            self,
            "Подтверждение удаления",
            f"Вы действительно хотите удалить резервную копию:\n{timestamp}?\n\n"
            "Это действие нельзя отменить.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            if delete_backup(backup_path):
                QMessageBox.information(self, "Успех", "Резервная копия удалена")
                self.load_backups()
            else:
                QMessageBox.warning(self, "Ошибка", "Не удалось удалить резервную копию")
    
    def restore_selected(self):
        """Восстановить базу данных из выбранной резервной копии"""
        if not self.selected_backup:
            QMessageBox.warning(self, "Ошибка", "Выберите резервную копию для восстановления")
            return
        
        timestamp = get_backup_timestamp(self.selected_backup)
        
        reply = QMessageBox.question(
            self,
            "Подтверждение восстановления",
            f"Восстановить базу данных из резервной копии:\n{timestamp}?\n\n"
            "Текущая база данных будет сохранена перед восстановлением.\n"
            "После восстановления необходимо перезапустить приложение.\n\n"
            "Продолжить?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply != QMessageBox.Yes:
            return
        
        if restore_database(self.selected_backup):
            QMessageBox.information(
                self,
                "Успех",
                f"База данных успешно восстановлена из резервной копии:\n{timestamp}\n\n"
                "Текущая БД сохранена как 'app.db.before-restore-...'\n\n"
                "⚠ ВАЖНО: Перезапустите приложение для применения изменений!"
            )
            self.accept()
        else:
            QMessageBox.critical(
                self,
                "Ошибка",
                "Не удалось восстановить базу данных.\n"
                "Проверьте права доступа к файлам."
            )

