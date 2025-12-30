"""
Скрипт для восстановления нижней таблицы на первой странице шаблона
Копирует диапазон A43:N58 из Образец.xls в templates/Izveshchenie_template.xlsx
с сохранением всех форматов, объединений ячеек, границ и стилей
"""
import sys
from pathlib import Path

# Добавляем корневую директорию проекта в путь
PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from app.config import TEMPLATE_PATH, TEMPLATE_CONFIG_FILE
import json


def restore_bottom_block():
    """Восстановить нижнюю таблицу через Excel automation"""
    
    try:
        import win32com.client
    except ImportError:
        print("Ошибка: Не установлен pywin32. Установите его командой:")
        print("  pip install pywin32")
        sys.exit(1)
    
    sample_path = PROJECT_ROOT / "Образец.xls"
    template_path = TEMPLATE_PATH
    
    if not sample_path.exists():
        print(f"Ошибка: Образец не найден: {sample_path}")
        sys.exit(1)
    
    if not template_path.exists():
        print(f"Ошибка: Шаблон не найден: {template_path}")
        sys.exit(1)
    
    # Загружаем конфигурацию для определения основного листа
    if TEMPLATE_CONFIG_FILE.exists():
        with open(TEMPLATE_CONFIG_FILE, "r", encoding="utf-8") as f:
            config = json.load(f)
        main_sheet_name = config.get("main_sheet", "Лист1")
    else:
        main_sheet_name = "Лист1"
    
    print(f"Образец: {sample_path}")
    print(f"Шаблон: {template_path}")
    print(f"Основной лист: {main_sheet_name}")
    print()
    
    # Запускаем Excel
    print("Запускаю Excel...")
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False  # Не показывать Excel
    excel.DisplayAlerts = False  # Отключить предупреждения
    
    try:
        # Открываем образец
        print("Открываю образец...")
        sample_wb = excel.Workbooks.Open(str(sample_path.absolute()))
        sample_sheet = sample_wb.Worksheets(1)  # Первый лист
        
        # Открываем шаблон
        print("Открываю шаблон...")
        template_wb = excel.Workbooks.Open(str(template_path.absolute()))
        
        # Находим нужный лист в шаблоне
        template_sheet = None
        for sheet in template_wb.Worksheets:
            if sheet.Name == main_sheet_name:
                template_sheet = sheet
                break
        
        if template_sheet is None:
            print(f"Ошибка: Лист '{main_sheet_name}' не найден в шаблоне")
            print(f"Доступные листы: {[s.Name for s in template_wb.Worksheets]}")
            template_wb.Close(SaveChanges=False)
            sample_wb.Close(SaveChanges=False)
            excel.Quit()
            sys.exit(1)
        
        print(f"Найден лист в шаблоне: {template_sheet.Name}")
        
        # Копируем диапазон A43:N58 из образца
        source_range = sample_sheet.Range("A43:N58")
        print(f"Копирую диапазон {source_range.Address} из образца...")
        source_range.Copy()
        
        # Вставляем в шаблон на тот же диапазон A43:N58
        target_range = template_sheet.Range("A43:N58")
        print(f"Вставляю в диапазон {target_range.Address} шаблона...")
        
        # Вставляем с сохранением всех форматов (значения + форматы)
        target_range.PasteSpecial(
            Paste=-4163,  # xlPasteValuesAndNumberFormats
            Operation=-4142,  # xlNone
            SkipBlanks=False,
            Transpose=False
        )
        
        # Также копируем форматы отдельно для полного соответствия
        source_range.Copy()
        target_range.PasteSpecial(
            Paste=-4122,  # xlPasteFormats
            Operation=-4142,  # xlNone
            SkipBlanks=False,
            Transpose=False
        )
        
        # Очищаем буфер обмена
        excel.CutCopyMode = False
        
        print("Копирование завершено")
        
        # Сохраняем шаблон
        print("Сохраняю шаблон...")
        template_wb.Save()
        
        # Закрываем файлы
        template_wb.Close()
        sample_wb.Close()
        
        print("Готово! Нижняя таблица восстановлена из образца.")
        print("Все форматы, объединения ячеек и стили скопированы.")
        
    except Exception as e:
        print(f"Ошибка при работе с Excel: {e}")
        import traceback
        traceback.print_exc()
        try:
            excel.Quit()
        except:
            pass
        sys.exit(1)
    finally:
        # Закрываем Excel
        try:
            excel.Quit()
        except:
            pass


if __name__ == "__main__":
    try:
        restore_bottom_block()
    except KeyboardInterrupt:
        print("\nПрервано пользователем")
        sys.exit(1)
    except Exception as e:
        print(f"Критическая ошибка: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

