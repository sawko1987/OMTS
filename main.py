"""
Главный файл запуска приложения "Извещение на замену материалов"
"""
import sys
import logging
from pathlib import Path

from PySide6.QtWidgets import QApplication
from PySide6.QtCore import Qt

from app.gui.main_window import MainWindow
from app.config import PROJECT_ROOT

def setup_logging():
    """Настройка логирования для приложения"""
    # Настройка формата логов
    log_format = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    date_format = '%Y-%m-%d %H:%M:%S'
    
    # Настройка базового логирования
    logging.basicConfig(
        level=logging.INFO,
        format=log_format,
        datefmt=date_format,
        handlers=[
            logging.StreamHandler(sys.stdout)  # Вывод в консоль
        ]
    )
    
    # Устанавливаем уровень логирования для модулей приложения
    logging.getLogger('app').setLevel(logging.INFO)
    logging.getLogger('app.catalog_loader').setLevel(logging.INFO)
    logging.getLogger('app.gui').setLevel(logging.INFO)
    logging.getLogger('app.gui.changes_table_widget').setLevel(logging.INFO)
    logging.getLogger('app.excel_generator').setLevel(logging.INFO)
    
    logger = logging.getLogger(__name__)
    logger.info("Логирование инициализировано")

def main():
    """Главная функция"""
    # Настраиваем логирование
    setup_logging()
    logger = logging.getLogger(__name__)
    logger.info("Запуск приложения 'Извещение на замену материалов'")
    
    # Создаём приложение
    app = QApplication(sys.argv)
    app.setApplicationName("Извещение на замену материалов")
    
    # Создаём главное окно
    window = MainWindow()
    window.show()
    logger.info("Главное окно создано и отображено")
    
    # Запускаем цикл событий
    sys.exit(app.exec())

if __name__ == "__main__":
    main()

