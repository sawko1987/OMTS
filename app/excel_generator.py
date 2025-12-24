"""
Генератор Excel документов по шаблону
"""
from pathlib import Path
from datetime import date
import json
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from app.config import TEMPLATE_PATH, TEMPLATE_CONFIG_FILE
from app.models import DocumentData, PartChanges
from app.numbering import NumberingManager


def get_merged_cell_value(sheet, row, col):
    """
    Получить ячейку для записи, учитывая объединенные ячейки.
    Если ячейка является частью объединенной области, возвращает главную ячейку.
    """
    cell = sheet.cell(row, col)
    
    # Проверяем, является ли ячейка частью объединенной области
    for merged_range in sheet.merged_cells.ranges:
        # Проверяем, находится ли ячейка в диапазоне объединенных ячеек
        if (merged_range.min_row <= row <= merged_range.max_row and
            merged_range.min_col <= col <= merged_range.max_col):
            # Возвращаем главную ячейку (верхняя левая ячейка объединенной области)
            return sheet.cell(merged_range.min_row, merged_range.min_col)
    
    return cell


class ExcelGenerator:
    """Генератор Excel документов"""
    
    def __init__(self):
        self.template_path = TEMPLATE_PATH
        self.config = self.load_config()
        self.numbering = NumberingManager()
    
    def load_config(self) -> dict:
        """Загрузить конфигурацию шаблона"""
        if not TEMPLATE_CONFIG_FILE.exists():
            raise FileNotFoundError(f"Конфигурация шаблона не найдена: {TEMPLATE_CONFIG_FILE}")
        
        with open(TEMPLATE_CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    
    def generate(self, document_data: DocumentData, output_path: Path):
        """Сгенерировать документ"""
        if not self.template_path.exists():
            raise FileNotFoundError(f"Шаблон не найден: {self.template_path}")
        
        # Загружаем шаблон
        wb = load_workbook(self.template_path)
        main_sheet = wb[self.config["main_sheet"]]
        
        # Получаем номер документа
        if not document_data.document_number:
            document_data.document_number = self.numbering.get_next_number()
        else:
            # Увеличиваем счётчик, если номер уже был установлен
            self.numbering.set_number(document_data.document_number)
            self.numbering.get_next_number()
        
        # Очищаем старые данные из шаблона
        self.clear_template_data(main_sheet)
        
        # Заполняем шапку
        self.fill_header(main_sheet, document_data)
        
        # Заполняем блок "Вручено"
        self.fill_vruhcheno(main_sheet, document_data)
        
        # Заполняем таблицу изменений
        remaining_data = self.fill_changes_table(main_sheet, document_data)
        
        # Обрабатываем дополнительные листы, если нужно
        if remaining_data:
            self.handle_additional_sheets(wb, remaining_data, document_data)
        
        # Сохраняем
        wb.save(output_path)
        wb.close()
    
    def clear_template_data(self, sheet):
        """Очистить старые данные из шаблона перед заполнением"""
        # Очищаем область таблицы изменений (начиная со строки 16 до конца данных)
        data_start_row = 16
        max_row = sheet.max_row
        
        # Очищаем данные в области таблицы (колонки A-N, строки 16-42)
        # Нижняя часть документа (строка 43 и далее) предназначена для ручного заполнения
        for row in range(data_start_row, min(max_row + 1, 43)):
            for col in range(1, 15):  # Колонки A-N
                cell = get_merged_cell_value(sheet, row, col)
                # Очищаем только значение, сохраняя форматирование
                if cell.value is not None:
                    cell.value = None
        
        # Очищаем блок "Вручено" (строка 7)
        for col in range(1, 16):
            cell = get_merged_cell_value(sheet, 7, col)
            if cell.value and str(cell.value).strip().lower() in ['х', 'x']:
                cell.value = None
        
        # Очищаем поля шапки, которые будут заполняться
        # Дата внедрения (строка 9, колонка E)
        cell = get_merged_cell_value(sheet, 9, 5)
        if cell.value:
            cell.value = None
        
        # Срок действия (строка 9, колонка K)
        cell = get_merged_cell_value(sheet, 9, 11)
        if cell.value:
            cell.value = None
        
        # Изделие (строка 10, колонка B)
        cell = get_merged_cell_value(sheet, 10, 2)
        if cell.value:
            cell.value = None
        
        # Причина (строка 11, колонка B)
        cell = get_merged_cell_value(sheet, 11, 2)
        if cell.value:
            cell.value = None
        
        # Заключение ТКО (строка 10, колонка I)
        cell = get_merged_cell_value(sheet, 10, 9)
        if cell.value:
            cell.value = None
    
    def fill_header(self, sheet, doc_data: DocumentData):
        """Заполнить шапку документа"""
        # Номер документа (в заголовке, строка 5)
        for col in range(1, 16):
            cell = get_merged_cell_value(sheet, 5, col)
            if cell.value and "извещение" in str(cell.value).lower():
                text = str(cell.value)
                if "№" in text:
                    import re
                    new_text = re.sub(r'№\s*\d+', f'№ {doc_data.document_number}', text)
                    cell.value = new_text
                break
        
        # Дата внедрения (строка 9, колонка E)
        if doc_data.implementation_date:
            cell = get_merged_cell_value(sheet, 9, 5)
            # Сохраняем существующий формат, если он есть, иначе устанавливаем формат даты
            if not cell.number_format or cell.number_format == 'General':
                cell.number_format = "DD.MM.YYYY"
            cell.value = doc_data.implementation_date
        
        # Срок действия (строка 9, колонка K)
        if doc_data.validity_period:
            cell = get_merged_cell_value(sheet, 9, 11)
            cell.value = doc_data.validity_period
        
        # Изделие (строка 10, колонка B)
        if doc_data.product:
            cell = get_merged_cell_value(sheet, 10, 2)
            cell.value = doc_data.product
        
        # Причина (строка 11, колонка B)
        if doc_data.reason:
            cell = get_merged_cell_value(sheet, 11, 2)
            cell.value = doc_data.reason
        
        # Заключение ТКО (строка 10, колонка I)
        if doc_data.tko_conclusion:
            cell = get_merged_cell_value(sheet, 10, 9)
            cell.value = doc_data.tko_conclusion
    
    def fill_vruhcheno(self, sheet, doc_data: DocumentData):
        """Заполнить блок 'Вручено'"""
        # Получаем все цеха из документа
        workshops_in_doc = doc_data.get_all_workshops()
        
        # Из анализа шаблона: строка 6 - заголовки, строка 7 - отметки
        header_row = 6
        mark_row = 7
        
        workshop_cols = {}
        
        # Ищем заголовки цехов в строке 6
        for col in range(1, 16):
            cell = get_merged_cell_value(sheet, header_row, col)
            if cell.value:
                val = str(cell.value).strip().upper()
                # Ищем цеха (ПЗУ может быть как ПДО)
                if "ПЗУ" in val or ("ПДО" in val and "КЛАДОВАЯ" not in val):
                    workshop_cols["ПЗУ"] = col
                elif "ЗМУ" in val:
                    workshop_cols["ЗМУ"] = col
                elif "СУ" in val and "ОСУ" not in val:
                    workshop_cols["СУ"] = col
                elif "ОСУ" in val:
                    workshop_cols["ОСУ"] = col
        
        # Если не нашли через заголовки, используем конфиг
        if not workshop_cols:
            findings = self.config["findings"]
            vruhcheno = findings.get("vruhcheno_marks", {})
            
            # Маппинг из конфига
            config_mapping = {
                "ПЗУ": ["ПДО", "Кладовая ПДО"],
                "ЗМУ": ["ЗМУ"],
                "СУ": ["СУ"],
                "ОСУ": ["ОСУ"]
            }
            
            for workshop in workshops_in_doc:
                if workshop in config_mapping:
                    for alt_name in config_mapping[workshop]:
                        if alt_name in vruhcheno:
                            coord_data = vruhcheno[alt_name]
                            if isinstance(coord_data, list) and len(coord_data) >= 3:
                                coord = coord_data[2]
                                sheet[coord].value = "х"
                                break
        else:
            # Ставим отметки по найденным колонкам
            for workshop in workshops_in_doc:
                if workshop in workshop_cols:
                    col = workshop_cols[workshop]
                    cell = get_merged_cell_value(sheet, mark_row, col)
                    cell.value = "х"
    
    def fill_changes_table(self, sheet, doc_data: DocumentData):
        """Заполнить таблицу изменений на первом листе
        
        Записывает ТОЛЬКО значения, не изменяя стили, форматирование, ширину колонок и т.д.
        
        Returns:
            list: Оставшиеся данные, которые не поместились (для доп. листов)
        """
        # Данные начинаются со строки 16 (из анализа шаблона)
        data_start_row = 16
        current_row = data_start_row
        
        # Максимальная строка на первом листе (до нижней части для ручного заполнения, строка 42)
        # Нижняя часть документа (строка 43 и далее) предназначена для ручного заполнения
        max_row_first_sheet = 42
        
        remaining_data = []
        current_part_data = None
        
        for part_change in doc_data.part_changes:
            # Проверяем, есть ли изменения для этой детали
            changed_materials = [m for m in part_change.materials if m.is_changed and m.after_name]
            if not changed_materials:
                continue
            
            # Если выходим за пределы первого листа, сохраняем для доп. листов
            if current_row >= max_row_first_sheet:
                remaining_data.append(part_change)
                continue
            
            # Записываем деталь в колонку A (первая строка группы)
            # Только значение, стиль не трогаем
            cell = get_merged_cell_value(sheet, current_row, 1)
            cell.value = part_change.part
            
            # Записываем деталь в колонку G (правая часть)
            cell = get_merged_cell_value(sheet, current_row, 7)
            cell.value = part_change.part
            
            current_row += 1
            
            # Записываем материалы для этой детали
            for material in changed_materials:
                if current_row >= max_row_first_sheet:
                    # Сохраняем оставшиеся материалы для доп. листов
                    if not current_part_data:
                        current_part_data = PartChanges(part=part_change.part)
                        remaining_data.append(current_part_data)
                    current_part_data.materials.append(material)
                    continue
                
                # Левая часть: "Подлежит изменению"
                # Колонка A: Наименование материала "до"
                cell = get_merged_cell_value(sheet, current_row, 1)
                cell.value = material.catalog_entry.before_name
                
                # Колонка D (4): Ед. изм. "до"
                cell = get_merged_cell_value(sheet, current_row, 4)
                cell.value = material.catalog_entry.unit
                
                # Колонка E (5): Норма "до"
                cell = get_merged_cell_value(sheet, current_row, 5)
                cell.value = material.catalog_entry.norm
                # Сохраняем формат числа, если он был в шаблоне
                # Не устанавливаем новый формат, чтобы не перезаписать существующий
                
                # Правая часть: "Вносимые изменения"
                # Колонка G (7): Наименование материала "после"
                cell = get_merged_cell_value(sheet, current_row, 7)
                cell.value = material.after_name
                
                # Колонка I (9): Ед. изм. "после"
                cell = get_merged_cell_value(sheet, current_row, 9)
                if material.after_unit:
                    cell.value = material.after_unit
                else:
                    cell.value = material.catalog_entry.unit
                
                # Колонка J (10): Норма "после" (если указана)
                if material.after_norm is not None:
                    cell = get_merged_cell_value(sheet, current_row, 10)
                    cell.value = material.after_norm
                
                current_row += 1
                current_part_data = None
        
        return remaining_data
    
    def handle_additional_sheets(self, wb, remaining_data, doc_data: DocumentData):
        """Обработать дополнительные листы при переполнении
        
        Записывает ТОЛЬКО значения, сохраняя разметку шаблона
        """
        if not remaining_data:
            return
        
        # Ищем шаблон дополнительного листа (обычно "1+", "2+" и т.д.)
        additional_sheet_names = [name for name in wb.sheetnames if "+" in name]
        
        if not additional_sheet_names:
            # Если нет готовых листов, создаём новый на основе первого
            # Но лучше использовать существующие шаблоны
            return
        
        # Используем первый доступный дополнительный лист
        sheet = wb[additional_sheet_names[0]]
        
        # Очищаем старые данные на дополнительном листе
        self.clear_additional_sheet_data(sheet)
        
        # Данные на доп. листе начинаются примерно со строки 7 (из анализа)
        data_start_row = 7
        current_row = data_start_row
        
        for part_change in remaining_data:
            changed_materials = [m for m in part_change.materials if m.is_changed and m.after_name]
            if not changed_materials:
                continue
            
            # Записываем деталь (только значения)
            cell = get_merged_cell_value(sheet, current_row, 1)
            cell.value = part_change.part
            cell = get_merged_cell_value(sheet, current_row, 7)
            cell.value = part_change.part
            current_row += 1
            
            # Записываем материалы (только значения)
            for material in changed_materials:
                # Левая часть
                cell = get_merged_cell_value(sheet, current_row, 1)
                cell.value = material.catalog_entry.before_name
                cell = get_merged_cell_value(sheet, current_row, 4)
                cell.value = material.catalog_entry.unit
                cell = get_merged_cell_value(sheet, current_row, 5)
                cell.value = material.catalog_entry.norm
                
                # Правая часть
                cell = get_merged_cell_value(sheet, current_row, 7)
                cell.value = material.after_name
                cell = get_merged_cell_value(sheet, current_row, 9)
                if material.after_unit:
                    cell.value = material.after_unit
                else:
                    cell.value = material.catalog_entry.unit
                if material.after_norm is not None:
                    cell = get_merged_cell_value(sheet, current_row, 10)
                    cell.value = material.after_norm
                
                current_row += 1

    
    def clear_additional_sheet_data(self, sheet):
        """Очистить старые данные на дополнительном листе"""
        # Данные начинаются со строки 7
        data_start_row = 7
        max_row = sheet.max_row
        
        # Очищаем данные в области таблицы (колонки A-N)
        for row in range(data_start_row, min(max_row + 1, 200)):
            for col in range(1, 15):  # Колонки A-N
                cell = get_merged_cell_value(sheet, row, col)
                # Очищаем только значение, сохраняя форматирование
                if cell.value is not None:
                    cell.value = None
