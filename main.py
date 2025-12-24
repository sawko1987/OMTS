"""
Главный файл запуска приложения "Извещение на замену материалов"
"""
import sys
from pathlib import Path

from PySide6.QtWidgets import QApplication
from PySide6.QtCore import Qt

from app.gui.main_window import MainWindow
from app.config import PROJECT_ROOT

def main():
    """Главная функция"""
    # Создаём приложение
    app = QApplication(sys.argv)
    app.setApplicationName("Извещение на замену материалов")
    
    # Создаём главное окно
    window = MainWindow()
    window.show()
    
    # Запускаем цикл событий
    sys.exit(app.exec())

if __name__ == "__main__":
    main()

