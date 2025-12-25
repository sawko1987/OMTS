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
from app.document_store import DocumentStore

class MainWindow(QMainWindow):
    """Главное окно приложения"""
    
    def __init__(self):
        super().__init__()
        self.document_data = DocumentData()
        
        # Счётчик дополнительных страниц (0 = первая страница, 1 = 1+, 2 = 2+ и т.д.)
        self.current_additional_page = None  # None = данные идут на первую страницу
        
        # Инициализация БД (до создания UI)
        self.init_database()
        
        # Создаём менеджеры данных
        self.db_manager = DatabaseManager()
        self.product_store = ProductStore(self.db_manager)
        self.catalog_loader = CatalogLoader(self.db_manager)
        self.history_store = HistoryStore(self.db_manager)
        self.settings_manager = SettingsManager()
        self.document_store = DocumentStore(self.db_manager, self.catalog_loader)
        
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
        
        # Меню "Документ"
        document_menu = menubar.addMenu("Документ")
        new_doc_action = document_menu.addAction("Новый документ")
        new_doc_action.triggered.connect(self.new_document)
        open_doc_action = document_menu.addAction("Открыть документ...")
        open_doc_action.triggered.connect(self.open_document)
        
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
        # Передаём функцию для получения текущего номера доп. страницы
        self.changes_widget.get_current_additional_page = lambda: self.current_additional_page
        self.tabs.addTab(self.changes_widget, "Изменения материалов")
        
        # Кнопки внизу
        button_layout = QHBoxLayout()
        
        self.btn_generate = QPushButton("Сгенерировать документ")
        self.btn_generate.clicked.connect(self.generate_document)
        button_layout.addWidget(self.btn_generate)
        
        self.btn_add_additional_page = QPushButton("Добавить данные для доп. страницы")
        self.btn_add_additional_page.clicked.connect(self.add_additional_page_data)
        button_layout.addWidget(self.btn_add_additional_page)
        
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
            # Обновляем номер документа, если он был изменен
            self.doc_info_widget.refresh_number()
    
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
        else:
            # Проверяем доступность сохраненного пути
            path_obj = Path(selected_dir)
            if not path_obj.exists():
                # Путь не существует - пытаемся создать его
                try:
                    path_obj.mkdir(parents=True, exist_ok=True)
                except Exception as e:
                    # Не удалось создать - предлагаем выбрать другую папку
                    reply = QMessageBox.question(
                        self,
                        "Папка недоступна",
                        f"Сохраненная папка недоступна:\n{selected_dir}\n\n"
                        f"Ошибка: {e}\n\n"
                        "Выбрать другую папку?",
                        QMessageBox.Yes | QMessageBox.No
                    )
                    if reply == QMessageBox.Yes:
                        output_dir = PROJECT_ROOT / "output"
                        selected_dir = QFileDialog.getExistingDirectory(
                            self,
                            "Выберите папку для сохранения извещений",
                            str(output_dir)
                        )
                        if not selected_dir:
                            return
                        try:
                            self.settings_manager.set_output_directory(selected_dir)
                        except Exception as e2:
                            QMessageBox.warning(self, "Предупреждение", 
                                              f"Не удалось сохранить настройки:\n{e2}")
                    else:
                        return
            elif not path_obj.is_dir():
                # Путь существует, но это не папка
                QMessageBox.warning(self, "Ошибка", 
                                  f"Указанный путь не является папкой:\n{selected_dir}")
                return
        
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
        try:
            target_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", 
                               f"Не удалось создать папку для сохранения:\n{target_dir}\n\n"
                               f"Ошибка: {e}")
            return
        
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
            
            # Сохраняем документ в БД
            self.document_store.save_document(self.document_data, str(file_path))
            
            QMessageBox.information(self, "Успех", f"Документ успешно создан:\n{file_path}")
            
            # Создаём новый документ с новым номером
            self.new_document()
            
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось создать документ:\n{e}")
    
    def new_document(self):
        """Создать новый документ"""
        # Проверяем, есть ли несохранённые изменения
        if self.document_data.part_changes or self.document_data.product:
            reply = QMessageBox.question(
                self,
                "Подтверждение",
                "Создать новый документ? Текущие данные будут очищены.",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply != QMessageBox.Yes:
                return
        
        # Создаём новый документ
        self.document_data = DocumentData()
        self.current_additional_page = None  # Сбрасываем счётчик доп. страниц
        self.doc_info_widget.document_data = self.document_data
        self.changes_widget.document_data = self.document_data
        self.doc_info_widget.refresh()
        self.changes_widget.refresh()
        self.changes_widget.update_product_filter()
    
    def open_document(self):
        """Открыть существующий документ"""
        from app.gui.document_selection_dialog import DocumentSelectionDialog
        
        dialog = DocumentSelectionDialog(self.document_store, self)
        if dialog.exec():
            document_number, year = dialog.get_selected_document()
            if document_number:
                loaded_data = self.document_store.load_document(document_number, year)
                if loaded_data:
                    self.document_data = loaded_data
                    # Определяем максимальный номер доп. страницы для восстановления счётчика
                    max_page = 0
                    for part_change in self.document_data.part_changes:
                        if part_change.additional_page_number is not None:
                            max_page = max(max_page, part_change.additional_page_number)
                    self.current_additional_page = max_page if max_page > 0 else None
                    self.doc_info_widget.document_data = self.document_data
                    self.changes_widget.document_data = self.document_data
                    self.doc_info_widget.refresh()
                    self.changes_widget.refresh()
                    self.changes_widget.update_product_filter()
                    QMessageBox.information(self, "Успех", f"Документ №{document_number} загружен")
                else:
                    QMessageBox.warning(self, "Ошибка", f"Не удалось загрузить документ №{document_number}")
    
    def add_additional_page_data(self):
        """Добавить данные для дополнительной страницы"""
        # Определяем номер следующей доп. страницы
        # Находим максимальный номер доп. страницы среди существующих данных
        max_page = 0
        for part_change in self.document_data.part_changes:
            if part_change.additional_page_number is not None:
                max_page = max(max_page, part_change.additional_page_number)
        
        # Увеличиваем счётчик для новой страницы
        self.current_additional_page = max_page + 1
        
        page_label = f"{self.current_additional_page}+"
        
        message = (
            f"Режим добавления данных для страницы {page_label} активирован.\n\n"
            f"Все детали и материалы, которые вы добавите сейчас, будут размещены на странице {page_label}.\n\n"
            f"Переключить на вкладку 'Изменения материалов'?"
        )
        
        reply = QMessageBox.question(
            self,
            f"Добавить данные для страницы {page_label}",
            message,
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes
        )
        
        if reply == QMessageBox.Yes:
            # Переключаемся на вкладку изменений материалов
            self.tabs.setCurrentIndex(1)
            # Фокусируемся на виджете изменений
            self.changes_widget.setFocus()
    
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
            self.current_additional_page = None  # Сбрасываем счётчик доп. страниц
            self.doc_info_widget.document_data = self.document_data
            self.changes_widget.document_data = self.document_data
            self.doc_info_widget.refresh()
            self.changes_widget.refresh()
            self.changes_widget.update_product_filter()

