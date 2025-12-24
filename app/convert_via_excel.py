"""
Конвертация шаблона .xls в .xlsx через установленный Excel
для полного сохранения форматирования
"""
import sys
from pathlib import Path
import win32com.client
import shutil
import time

def convert_xls_to_xlsx_via_excel(xls_path: Path, xlsx_path: Path):
    """Конвертирует .xls в .xlsx через Excel с сохранением форматирования"""
    
    print(f"Конвертация через Excel: {xls_path} -> {xlsx_path}")
    
    try:
        # Запускаем Excel
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        # Открываем .xls файл
        print(f"Открытие файла: {xls_path}")
        workbook = excel.Workbooks.Open(str(xls_path.absolute()))
        
        # Сохраняем как .xlsx (формат 51 = xlOpenXMLWorkbook)
        print(f"Сохранение как .xlsx: {xlsx_path}")
        xlsx_path.parent.mkdir(exist_ok=True, parents=True)
        
        # Используем временное имя, чтобы избежать конфликтов
        temp_path = xlsx_path.parent / f"temp_{xlsx_path.name}"
        
        # Удаляем временный файл, если существует
        if temp_path.exists():
            try:
                temp_path.unlink()
            except:
                pass
        
        # Сохраняем во временный файл
        workbook.SaveAs(
            str(temp_path.absolute()),
            FileFormat=51  # xlOpenXMLWorkbook (.xlsx)
        )
        
        # Закрываем книгу
        workbook.Close(SaveChanges=False)
        
        # Ждём немного, чтобы файл был полностью записан
        time.sleep(0.5)
        
        # Копируем через shutil вместо rename
        if xlsx_path.exists():
            try:
                xlsx_path.unlink()
            except:
                # Если не удалось удалить, пробуем переименовать старый
                backup_path = xlsx_path.parent / f"backup_{xlsx_path.name}"
                try:
                    if backup_path.exists():
                        backup_path.unlink()
                    xlsx_path.rename(backup_path)
                except:
                    pass
        
        shutil.copy2(temp_path, xlsx_path)
        
        # Удаляем временный файл
        try:
            temp_path.unlink()
        except:
            pass
        
        # Закрываем
        excel.Quit()
        
        print(f"[OK] Конвертация завершена: {xlsx_path}")
        return True
        
    except Exception as e:
        print(f"[ERROR] Ошибка конвертации: {e}")
        try:
            excel.Quit()
        except:
            pass
        return False

if __name__ == "__main__":
    xls_path = Path(__file__).parent.parent / "Образец.xls"
    xlsx_path = Path(__file__).parent.parent / "templates" / "Izveshchenie_template.xlsx"
    
    if not xls_path.exists():
        print(f"Файл не найден: {xls_path}")
        sys.exit(1)
    
    success = convert_xls_to_xlsx_via_excel(xls_path, xlsx_path)
    sys.exit(0 if success else 1)

