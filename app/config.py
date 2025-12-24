"""
Конфигурация приложения
"""
from pathlib import Path

# Корневая директория проекта
PROJECT_ROOT = Path(__file__).parent.parent

# Пути к файлам
TEMPLATE_PATH = PROJECT_ROOT / "templates" / "Izveshchenie_template.xlsx"
CATALOG_PATH = PROJECT_ROOT / "catalog" / "catalog.xlsx"
DATA_DIR = PROJECT_ROOT / "data"
DATABASE_PATH = DATA_DIR / "app.db"
NUMBERING_FILE = DATA_DIR / "numbering.json"  # Для обратной совместимости
HISTORY_FILE = DATA_DIR / "history.json"  # Для обратной совместимости
TEMPLATE_CONFIG_FILE = DATA_DIR / "template_config.json"
SETTINGS_FILE = DATA_DIR / "settings.json"

# Создаём необходимые директории
DATA_DIR.mkdir(exist_ok=True, parents=True)
(PROJECT_ROOT / "catalog").mkdir(exist_ok=True, parents=True)
(PROJECT_ROOT / "templates").mkdir(exist_ok=True, parents=True)
(PROJECT_ROOT / "output").mkdir(exist_ok=True, parents=True)

# Цеха
WORKSHOPS = ["ПЗУ", "ЗМУ", "СУ", "ОСУ"]

# Маппинг цехов для блока "Вручено" (если нужен отдельный маппинг)
WORKSHOP_TO_VRUHCHENO = {
    "ПЗУ": "ПЗУ",
    "ЗМУ": "ЗМУ", 
    "СУ": "СУ",
    "ОСУ": "ОСУ"
}

