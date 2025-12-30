"""
Сериализация и десериализация DocumentData в JSON
"""
import json
from datetime import date, datetime
from typing import Any, Dict, List

from app.models import DocumentData, PartChanges, MaterialChange, CatalogEntry


class DocumentSerializer:
    """Сериализатор для DocumentData"""
    
    @staticmethod
    def serialize(document_data: DocumentData) -> str:
        """Сериализовать DocumentData в JSON строку"""
        data = {
            "document_number": document_data.document_number,
            "implementation_date": document_data.implementation_date.isoformat() if document_data.implementation_date else None,
            "validity_period": document_data.validity_period,
            "products": document_data.products,  # Список продуктов
            "reason": document_data.reason,
            "tko_conclusion": document_data.tko_conclusion,
            "part_changes": []
        }
        
        for part_change in document_data.part_changes:
            part_data = {
                "part": part_change.part,
                "additional_page_number": part_change.additional_page_number,
                "materials": []
            }
            
            for material in part_change.materials:
                material_data = {
                    "is_changed": material.is_changed,
                    "after_name": material.after_name,
                    "after_unit": material.after_unit,
                    "after_norm": material.after_norm,
                    "catalog_entry": {
                        "id": material.catalog_entry.id,
                        "part": material.catalog_entry.part,
                        "workshop": material.catalog_entry.workshop,
                        "role": material.catalog_entry.role,
                        "before_name": material.catalog_entry.before_name,
                        "unit": material.catalog_entry.unit,
                        "norm": material.catalog_entry.norm,
                        "comment": material.catalog_entry.comment,
                        "is_part_of_set": material.catalog_entry.is_part_of_set,
                        "replacement_set_id": material.catalog_entry.replacement_set_id
                    }
                }
                part_data["materials"].append(material_data)
            
            data["part_changes"].append(part_data)
        
        return json.dumps(data, ensure_ascii=False, indent=2)
    
    @staticmethod
    def deserialize(json_str: str, catalog_loader=None) -> DocumentData:
        """Десериализовать JSON строку в DocumentData"""
        data = json.loads(json_str)
        
        # Обратная совместимость: если есть старое поле "product" (строка), преобразуем в список
        products = data.get("products", [])
        if not products and "product" in data:
            # Старый формат - одна строка
            old_product = data.get("product", "")
            if old_product:
                products = [old_product]
        
        document_data = DocumentData(
            document_number=data.get("document_number"),
            implementation_date=date.fromisoformat(data["implementation_date"]) if data.get("implementation_date") else None,
            validity_period=data.get("validity_period"),
            products=products,
            reason=data.get("reason", ""),
            tko_conclusion=data.get("tko_conclusion", "")
        )
        
        # Загружаем каталог, если передан загрузчик
        catalog_entries_by_id = {}
        if catalog_loader:
            # Загружаем все записи каталога для восстановления CatalogEntry
            all_entries = catalog_loader.load()
            catalog_entries_by_id = {entry.id: entry for entry in all_entries if entry.id}
        
        for part_data in data.get("part_changes", []):
            part_change = PartChanges(
                part=part_data.get("part", ""),
                additional_page_number=part_data.get("additional_page_number")
            )
            
            for material_data in part_data.get("materials", []):
                catalog_entry_data = material_data.get("catalog_entry", {})
                
                # Пытаемся восстановить CatalogEntry из каталога
                catalog_entry = None
                if catalog_loader and catalog_entry_data.get("id"):
                    catalog_entry = catalog_entries_by_id.get(catalog_entry_data["id"])
                
                # Если не нашли в каталоге, создаём новый объект из данных
                if not catalog_entry:
                    catalog_entry = CatalogEntry(
                        id=catalog_entry_data.get("id"),
                        part=catalog_entry_data.get("part", ""),
                        workshop=catalog_entry_data.get("workshop", ""),
                        role=catalog_entry_data.get("role", ""),
                        before_name=catalog_entry_data.get("before_name", ""),
                        unit=catalog_entry_data.get("unit", ""),
                        norm=catalog_entry_data.get("norm", 0.0),
                        comment=catalog_entry_data.get("comment", ""),
                        is_part_of_set=catalog_entry_data.get("is_part_of_set", False),
                        replacement_set_id=catalog_entry_data.get("replacement_set_id")
                    )
                
                material_change = MaterialChange(
                    catalog_entry=catalog_entry,
                    is_changed=material_data.get("is_changed", False),
                    after_name=material_data.get("after_name"),
                    after_unit=material_data.get("after_unit"),
                    after_norm=material_data.get("after_norm")
                )
                
                part_change.materials.append(material_change)
            
            document_data.part_changes.append(part_change)
        
        return document_data

