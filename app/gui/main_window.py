"""
Главное окно приложения
"""
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QTabWidget,
    QPushButton, QLabel, QMessageBox, QFileDialog, QMenuBar, QMenu
)
from PySide6.QtCore import Qt
from pathlib import Path
from datetime import date
import re

from app.models import DocumentData
from app.gui.document_info_widget import DocumentInfoWidget
from app.gui.changes_table_widget import ChangesTableWidget
from app.gui.settings_dialog import SettingsDialog
from app.config import PROJECT_ROOT, DATABASE_PATH
from app.database import DatabaseManager
from app.product_store import ProductStore
from app.catalog_loader import CatalogLoader
from app.history_store import HistoryStore
from app.migrate_to_sqlite import migrate_all
from app.settings_manager import SettingsManager

class MainWindow(QMainWindow):
    """Главное окно приложения"""
    
    def __init__(self):
        super().__init__()
        self.document_data = DocumentData()
        
        # Инициализация БД (до создания UI)
        self.init_database()
        
        # Создаём менеджеры данных
        self.db_manager = DatabaseManager()
        self.product_store = ProductStore(self.db_manager)
        self.catalog_loader = CatalogLoader(self.db_manager)
        self.history_store = HistoryStore(self.db_manager)
        self.settings_manager = SettingsManager()
        
        self.init_ui()
        self.load_catalog()
    
    def init_database(self):
        """Инициализировать базу данных"""
        db_manager = DatabaseManager()
        
        # Проверяем, существует ли БД
        if not DATABASE_PATH.exists():
            # Первый запуск - создаём схему и мигрируем данные
            db_manager.initialize()
            migrate_all()
            # Показываем сообщение после создания UI
            # (будет показано в init_ui, если нужно)
        else:
            # БД существует, просто инициализируем схему (создаст таблицы, если их нет)
            db_manager.initialize()
    
    def init_ui(self):
        """Инициализация интерфейса"""
        self.setWindowTitle("Извещение на замену материалов")
        self.setMinimumSize(1200, 800)
        
        # Меню
        menubar = self.menuBar()
        settings_menu = menubar.addMenu("Настройки")
        settings_action = settings_menu.addAction("Настройки...")
        settings_action.triggered.connect(self.show_settings)
        
        # Центральный виджет
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Основной layout
        main_layout = QVBoxLayout(central_widget)
        
        # Вкладки
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)
        
        # Вкладка 1: Реквизиты документа
        self.doc_info_widget = DocumentInfoWidget(self.document_data, self.product_store, self.db_manager)
        self.doc_info_widget.product_changed.connect(self.on_product_changed)
        self.tabs.addTab(self.doc_info_widget, "Реквизиты документа")
        
        # Вкладка 2: Таблица изменений
        self.changes_widget = ChangesTableWidget(
            self.document_data,
            self.catalog_loader,
            self.history_store,
            self.product_store
        )
        self.tabs.addTab(self.changes_widget, "Изменения материалов")
        
        # Кнопки внизу
        button_layout = QHBoxLayout()
        
        self.btn_generate = QPushButton("Сгенерировать документ")
        self.btn_generate.clicked.connect(self.generate_document)
        button_layout.addWidget(self.btn_generate)
        
        self.btn_clear = QPushButton("Очистить")
        self.btn_clear.clicked.connect(self.clear_data)
        button_layout.addWidget(self.btn_clear)
        
        button_layout.addStretch()
        
        main_layout.addLayout(button_layout)
    
    def load_catalog(self):
        """Загрузить каталог и обновить виджеты"""
        try:
            self.catalog_loader.load()
            self.changes_widget.load_catalog()
        except Exception as e:
            QMessageBox.warning(self, "Ошибка", f"Не удалось загрузить справочник:\n{e}")
    
    def show_settings(self):
        """Показать диалог настроек"""
        dialog = SettingsDialog(self)
        if dialog.exec():
            # Настройки сохранены в диалоге
            pass
    
    def on_product_changed(self, product_name: str):
        """Обработка изменения изделия"""
        # Обновляем данные документа
        self.document_data.product = product_name
        # Обновляем список деталей на второй вкладке
        self.changes_widget.update_product_filter()
    
    def get_first_material_name(self) -> str:
        """Получить название первого заменяемого материала"""
        for part_change in self.document_data.part_changes:
            for material in part_change.materials:
                if material.is_changed and material.after_name and material.after_name.strip():
                    return material.after_name.strip()
        return "материал"
    
    def get_month_year_folder(self) -> str:
        """Сформировать название папки в формате 'Январь_2024' на основе даты внедрения"""
        if not self.document_data.implementation_date:
            # Если дата не указана, используем текущую дату
            impl_date = date.today()
        else:
            impl_date = self.document_data.implementation_date
        
        # Словарь названий месяцев на русском
        months = {
            1: "Январь", 2: "Февраль", 3: "Март", 4: "Апрель",
            5: "Май", 6: "Июнь", 7: "Июль", 8: "Август",
            9: "Сентябрь", 10: "Октябрь", 11: "Ноябрь", 12: "Декабрь"
        }
        
        month_name = months.get(impl_date.month, "Неизвестно")
        year = impl_date.year
        
        return f"{month_name}_{year}"
    
    def sanitize_filename(self, filename: str, max_length: int = 200) -> str:
        """Очистить название от недопустимых символов для имени файла"""
        # Удаляем недопустимые символы для Windows
        invalid_chars = r'[<>:"/\\|?*]'
        sanitized = re.sub(invalid_chars, '_', filename)
        
        # Удаляем ведущие и завершающие пробелы и точки
        sanitized = sanitized.strip(' .')
        
        # Ограничиваем длину
        if len(sanitized) > max_length:
            sanitized = sanitized[:max_length]
        
        # Если после очистки строка пустая, возвращаем значение по умолчанию
        if not sanitized:
            sanitized = "материал"
        
        return sanitized
    
    def generate_document(self):
        """Генерация документа"""
        # Обновляем данные из виджетов
        self.doc_info_widget.update_document_data()
        self.changes_widget.update_document_data()
        
        # Валидация
        if not self.document_data.implementation_date:
            QMessageBox.warning(self, "Ошибка", "Укажите дату внедрения замены")
            self.tabs.setCurrentIndex(0)
            return
        
        if not self.document_data.part_changes:
            QMessageBox.warning(self, "Ошибка", "Добавьте хотя бы одну деталь с изменениями")
            self.tabs.setCurrentIndex(1)
            return
        
        # Проверяем, что есть хотя бы одно изменение
        has_changes = False
        missing_after = []
        missing_checkbox = []
        
        for part_change in self.document_data.part_changes:
            for material in part_change.materials:
                if material.is_changed:
                    if not material.after_name or not material.after_name.strip():
                        missing_after.append(f"{part_change.part}: {material.catalog_entry.before_name}")
                    else:
                        has_changes = True
        
        # Более детальное сообщение об ошибке
        if not has_changes:
            error_msg = "Не найдено ни одного материала для замены.\n\n"
            if missing_after:
                error_msg += "Следующие материалы отмечены для замены, но не указан материал 'после':\n"
                for item in missing_after[:5]:  # Показываем первые 5
                    error_msg += f"  - {item}\n"
                if len(missing_after) > 5:
                    error_msg += f"  ... и ещё {len(missing_after) - 5}\n"
            else:
                error_msg += "Отметьте чекбокс 'Меняем' для хотя бы одного материала и укажите материал 'после'."
            
            QMessageBox.warning(self, "Ошибка", error_msg)
            self.tabs.setCurrentIndex(1)
            return
        
        # Получаем папку сохранения из настроек
        selected_dir = self.settings_manager.get_output_directory()
        
        # Если папка не настроена, предлагаем выбрать её
        if not selected_dir:
            output_dir = PROJECT_ROOT / "output"
            selected_dir = QFileDialog.getExistingDirectory(
                self,
                "Выберите папку для сохранения извещений",
                str(output_dir)
            )
            
            if not selected_dir:
                return
            
            # Сохраняем выбранную папку в настройки
            try:
                self.settings_manager.set_output_directory(selected_dir)
            except Exception as e:
                QMessageBox.warning(self, "Предупреждение", 
                                  f"Не удалось сохранить настройки:\n{e}\n\n"
                                  "Папка будет использована только для этого документа.")
        
        # Получаем номер документа заранее (для формирования имени файла)
        # Импортируем здесь, чтобы избежать циклических импортов
        from app.numbering import NumberingManager
        
        numbering = NumberingManager()
        if not self.document_data.document_number:
            # Получаем следующий номер и устанавливаем его в document_data
            # Генератор использует существующий номер и не будет вызывать get_next_number()
            doc_number = numbering.get_next_number()
            self.document_data.document_number = doc_number
        else:
            doc_number = self.document_data.document_number
        
        # Формируем структуру папок: месяц_год
        month_year_folder = self.get_month_year_folder()
        target_dir = Path(selected_dir) / month_year_folder
        
        # Создаём папку, если её нет
        target_dir.mkdir(parents=True, exist_ok=True)
        
        # Получаем название первого материала
        material_name = self.get_first_material_name()
        sanitized_material = self.sanitize_filename(material_name)
        
        # Формируем имя файла: номер_материал.xlsx
        filename = f"{doc_number}_{sanitized_material}.xlsx"
        file_path = target_dir / filename
        
        # Импортируем генератор здесь, чтобы избежать циклических импортов
        from app.excel_generator import ExcelGenerator
        
        try:
            generator = ExcelGenerator()
            generator.generate(self.document_data, file_path)
            QMessageBox.information(self, "Успех", f"Документ успешно создан:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось создать документ:\n{e}")
    
    def clear_data(self):
        """Очистить все данные"""
        reply = QMessageBox.question(
            self,
            "Подтверждение",
            "Очистить все введённые данные?",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            self.document_data = DocumentData()
            self.doc_info_widget.document_data = self.document_data
            self.changes_widget.document_data = self.document_data
            self.doc_info_widget.refresh()
            self.changes_widget.refresh()
            self.changes_widget.update_product_filter()

