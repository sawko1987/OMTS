"""
Менеджер настроек приложения
"""
import json
from pathlib import Path
from typing import Optional

from app.config import DATA_DIR

SETTINGS_FILE = DATA_DIR / "settings.json"


class SettingsManager:
    """Менеджер настроек приложения"""
    
    def __init__(self):
        self.settings_file = SETTINGS_FILE
        self._settings = self._load_settings()
    
    def _load_settings(self) -> dict:
        """Загрузить настройки из файла"""
        default_settings = {
            "output_directory": None,  # Путь к папке сохранения документов
            "starting_number": 1  # Начальный номер для нумерации извещений
        }
        
        if not self.settings_file.exists():
            return default_settings
        
        try:
            with open(self.settings_file, "r", encoding="utf-8") as f:
                settings = json.load(f)
                # Объединяем с настройками по умолчанию
                default_settings.update(settings)
                return default_settings
        except (json.JSONDecodeError, FileNotFoundError):
            return default_settings
    
    def _save_settings(self):
        """Сохранить настройки в файл"""
        self.settings_file.parent.mkdir(exist_ok=True, parents=True)
        with open(self.settings_file, "w", encoding="utf-8") as f:
            json.dump(self._settings, f, indent=2, ensure_ascii=False)
    
    def get_output_directory(self) -> Optional[str]:
        """Получить путь к папке сохранения документов"""
        path = self._settings.get("output_directory")
        if path:
            # Проверяем, что путь существует
            path_obj = Path(path)
            if path_obj.exists() and path_obj.is_dir():
                return str(path_obj)
        return None
    
    def set_output_directory(self, path: str):
        """Установить путь к папке сохранения документов"""
        path_obj = Path(path)
        if path_obj.exists() and path_obj.is_dir():
            self._settings["output_directory"] = str(path_obj.absolute())
            self._save_settings()
        else:
            raise ValueError(f"Путь не существует или не является папкой: {path}")
    
    def get_starting_number(self) -> int:
        """Получить начальный номер для нумерации извещений"""
        return self._settings.get("starting_number", 1)
    
    def set_starting_number(self, number: int):
        """Установить начальный номер для нумерации извещений"""
        if not isinstance(number, int) or number < 1:
            raise ValueError("Начальный номер должен быть положительным целым числом")
        self._settings["starting_number"] = number
        self._save_settings()

