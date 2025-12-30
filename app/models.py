"""
Модели данных для приложения
"""
from dataclasses import dataclass, field
from typing import List, Optional
from datetime import date, datetime

@dataclass
class CatalogEntry:
    """Запись из справочника каталога"""
    part: str  # Деталь
    workshop: str  # Цех
    role: str  # Тип позиции (краска, разбавитель и т.д.)
    before_name: str  # Наименование материала "до"
    unit: str  # Ед. изм.
    norm: float  # Норма
    comment: str = ""  # Примечание
    id: Optional[int] = None  # ID записи в БД
    is_part_of_set: bool = False  # Флаг, что материал входит в набор
    replacement_set_id: Optional[int] = None  # ID набора замены, если материал в наборе

@dataclass
class MaterialChange:
    """Одна строка изменения материала"""
    catalog_entry: CatalogEntry  # Исходная запись из каталога
    is_changed: bool = False  # Меняем ли эту позицию
    after_name: Optional[str] = None  # Наименование материала "после"
    after_unit: Optional[str] = None  # Ед. изм. "после" (если отличается)
    after_norm: Optional[float] = None  # Норма "после" (если отличается)

@dataclass
class PartChanges:
    """Изменения для одной детали"""
    part: str  # Код детали
    materials: List[MaterialChange] = field(default_factory=list)  # Список материалов
    additional_page_number: Optional[int] = None  # Номер доп. страницы (None = первая страница, 1 = 1+, 2 = 2+ и т.д.)

@dataclass
class DocumentData:
    """Данные документа "Извещение на замену материалов" """
    # Реквизиты документа
    document_number: Optional[int] = None  # Номер извещения (автоматически)
    implementation_date: Optional[date] = None  # Дата внедрения замены
    validity_period: Optional[str] = None  # Срок действия (партия)
    products: List[str] = field(default_factory=list)  # Изделия (машины)
    reason: str = ""  # Причина
    tko_conclusion: str = ""  # Заключение ТКО
    
    # Изменения по деталям
    part_changes: List[PartChanges] = field(default_factory=list)
    
    def get_all_workshops(self) -> List[str]:
        """Получить список всех цехов, встречающихся в документе"""
        workshops = set()
        for part_change in self.part_changes:
            for material in part_change.materials:
                if material.is_changed:
                    workshops.add(material.catalog_entry.workshop)
        return sorted(list(workshops))

@dataclass
class MaterialReplacementSet:
    """Набор материалов для замены"""
    id: Optional[int] = None  # ID набора в БД
    part_code: str = ""  # Код детали
    set_type: str = ""  # Тип набора: 'from' (что заменяем) или 'to' (на что заменяем)
    set_name: Optional[str] = None  # Опциональное название набора
    created_at: Optional[datetime] = None  # Дата создания
    materials: List['CatalogEntry'] = field(default_factory=list)  # Материалы в наборе

@dataclass
class MaterialSetItem:
    """Элемент набора материалов"""
    id: Optional[int] = None  # ID элемента в БД
    set_id: int = 0  # ID набора
    catalog_entry_id: int = 0  # ID записи каталога
    order_index: int = 0  # Порядок материала в наборе

