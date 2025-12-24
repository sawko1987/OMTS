"""
Загрузка справочника каталога из SQLite
"""
from pathlib import Path
from typing import List, Dict, Optional
import sqlite3
from datetime import datetime
import logging
import time

from app.config import CATALOG_PATH
from app.models import CatalogEntry, MaterialReplacementSet, MaterialSetItem
from app.database import DatabaseManager

# Настройка логирования
logger = logging.getLogger(__name__)


class CatalogLoader:
    """Загрузчик справочника каталога из SQLite"""
    
    def __init__(self, db_manager: DatabaseManager = None):
        self.db_manager = db_manager or DatabaseManager()
        self._entries: List[CatalogEntry] = []
        self._by_part: Dict[str, List[CatalogEntry]] = {}
        self._parts_cache: Optional[List[str]] = None
    
    def load(self) -> List[CatalogEntry]:
        """Загрузить справочник из БД"""
        with self.db_manager.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT id, part_code, workshop, role, before_name, unit, norm, comment,
                       is_part_of_set, replacement_set_id
                FROM catalog_entries
                ORDER BY part_code, workshop, role
            """)
            
            entries = []
            for row in cursor.fetchall():
                entry = CatalogEntry(
                    id=row['id'],
                    part=row['part_code'],
                    workshop=row['workshop'],
                    role=row['role'],
                    before_name=row['before_name'],
                    unit=row['unit'],
                    norm=row['norm'],
                    comment=row['comment'] or "",
                    is_part_of_set=bool(row['is_part_of_set'] or 0),
                    replacement_set_id=row['replacement_set_id']
                )
                entries.append(entry)
            
            self._entries = entries
            
            # Индексируем по деталям
            self._by_part = {}
            for entry in entries:
                if entry.part not in self._by_part:
                    self._by_part[entry.part] = []
                self._by_part[entry.part].append(entry)
            
            # Сбрасываем кэш списка деталей
            self._parts_cache = None
            
            return entries
    
    def get_entries_by_part(self, part: str) -> List[CatalogEntry]:
        """Получить все записи для детали"""
        if not self._by_part:
            self.load()
        return self._by_part.get(part, [])
    
    def get_all_parts(self) -> List[str]:
        """Получить список всех деталей (с кэшированием)"""
        if self._parts_cache is not None:
            return self._parts_cache
        
        with self.db_manager.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT DISTINCT part_code
                FROM catalog_entries
                ORDER BY part_code
            """)
            parts = [row['part_code'] for row in cursor.fetchall()]
            self._parts_cache = parts
            return parts
    
    def get_entries_by_part_and_workshop(self, part: str, workshop: str) -> List[CatalogEntry]:
        """Получить записи для детали и цеха"""
        entries = self.get_entries_by_part(part)
        return [e for e in entries if e.workshop == workshop]
    
    def search_parts(self, query: str) -> List[str]:
        """Быстрый поиск деталей по подстроке (case-insensitive)"""
        start_time = time.time()
        logger.info(f"Поиск деталей: запрос='{query}'")
        
        if not query:
            logger.info("Пустой запрос, возвращаем все детали")
            return self.get_all_parts()
        
        try:
            # Преобразуем запрос в верхний регистр для регистронезависимого поиска
            # Это необходимо, так как COLLATE NOCASE не работает для кириллицы в SQLite
            query_upper = query.upper()
            search_pattern = f"%{query_upper}%"
            logger.debug(f"Паттерн поиска (верхний регистр): '{search_pattern}'")
            
            with self.db_manager.get_connection() as conn:
                cursor = conn.cursor()
                # Используем UPPER() для колонки и передаем уже преобразованный паттерн
                # Это обеспечивает регистронезависимый поиск для кириллицы
                cursor.execute("""
                    SELECT DISTINCT part_code
                    FROM catalog_entries
                    WHERE UPPER(part_code) LIKE ?
                    ORDER BY part_code
                """, (search_pattern,))
                
                results = [row['part_code'] for row in cursor.fetchall()]
                elapsed_time = time.time() - start_time
                
                logger.info(f"Поиск завершен: найдено {len(results)} деталей за {elapsed_time:.3f} сек")
                if results:
                    logger.debug(f"Найденные детали (первые 10): {results[:10]}")
                if len(results) > 1000:
                    logger.warning(f"Найдено большое количество результатов ({len(results)}), это может замедлить работу интерфейса")
                
                return results
        except sqlite3.Error as e:
            elapsed_time = time.time() - start_time
            logger.error(f"Ошибка при поиске деталей (запрос='{query}'): {e}, время выполнения: {elapsed_time:.3f} сек")
            raise
        except Exception as e:
            elapsed_time = time.time() - start_time
            logger.error(f"Неожиданная ошибка при поиске деталей (запрос='{query}'): {e}, время выполнения: {elapsed_time:.3f} сек")
            raise
    
    def _add_entry_internal(self, cursor: sqlite3.Cursor, entry: CatalogEntry) -> Optional[int]:
        """Внутренний метод для добавления записи в каталог (используется внутри транзакции)"""
        try:
            # Добавляем деталь в таблицу parts, если её ещё нет
            cursor.execute(
                "INSERT OR IGNORE INTO parts (code) VALUES (?)",
                (entry.part,)
            )
            
            # Добавляем запись каталога
            cursor.execute("""
                INSERT INTO catalog_entries 
                (part_code, workshop, role, before_name, unit, norm, comment,
                 is_part_of_set, replacement_set_id)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                entry.part,
                entry.workshop,
                entry.role,
                entry.before_name,
                entry.unit,
                entry.norm,
                entry.comment,
                1 if entry.is_part_of_set else 0,
                entry.replacement_set_id
            ))
            
            entry_id = cursor.lastrowid
            entry.id = entry_id
            return entry_id
        except sqlite3.IntegrityError as e:
            print(f"Ошибка при добавлении записи: {e}")
            return None
    
    def add_entry(self, entry: CatalogEntry) -> Optional[int]:
        """Добавить новую запись в каталог. Возвращает ID созданной записи или None при ошибке"""
        try:
            with self.db_manager.get_connection() as conn:
                cursor = conn.cursor()
                entry_id = self._add_entry_internal(cursor, entry)
                
                if entry_id:
                    # Обновляем кэш
                    self._parts_cache = None
                    if entry.part not in self._by_part:
                        self._by_part[entry.part] = []
                    self._by_part[entry.part].append(entry)
                    self._entries.append(entry)
                
                return entry_id
        except sqlite3.Error as e:
            print(f"Ошибка при добавлении записи: {e}")
            return None
    
    def update_entry(self, entry_id: int, entry: CatalogEntry) -> bool:
        """Обновить запись каталога"""
        try:
            with self.db_manager.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("""
                    UPDATE catalog_entries
                    SET part_code = ?, workshop = ?, role = ?, before_name = ?,
                        unit = ?, norm = ?, comment = ?
                    WHERE id = ?
                """, (
                    entry.part, entry.workshop, entry.role, entry.before_name,
                    entry.unit, entry.norm, entry.comment, entry_id
                ))
                conn.commit()
                
                # Обновляем кэш
                self._parts_cache = None
                self._by_part = {}
                self._entries = []
                self.load()
                
                return True
        except sqlite3.Error as e:
            print(f"Ошибка при обновлении записи: {e}")
            return False
    
    def delete_entry(self, entry_id: int) -> bool:
        """Удалить запись каталога"""
        try:
            with self.db_manager.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("DELETE FROM catalog_entries WHERE id = ?", (entry_id,))
                conn.commit()
                
                # Обновляем кэш
                self._parts_cache = None
                self._by_part = {}
                self._entries = []
                self.load()
                
                return True
        except sqlite3.Error as e:
            print(f"Ошибка при удалении записи: {e}")
            return False
    
    def part_exists(self, part_code: str) -> bool:
        """Проверить существование детали"""
        with self.db_manager.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT COUNT(*) FROM catalog_entries WHERE part_code = ?
            """, (part_code,))
            return cursor.fetchone()[0] > 0
    
    def add_replacement_set(self, part_code: str, from_materials: List[CatalogEntry], 
                           to_materials: List[CatalogEntry], set_name: Optional[str] = None) -> Optional[int]:
        """Создать набор замены материалов. Возвращает ID созданного набора или None при ошибке"""
        try:
            with self.db_manager.get_connection() as conn:
                cursor = conn.cursor()
                
                # Добавляем деталь в таблицу parts, если её ещё нет
                cursor.execute(
                    "INSERT OR IGNORE INTO parts (code) VALUES (?)",
                    (part_code,)
                )
                
                # Создаём набор "from" (что заменяем)
                cursor.execute("""
                    INSERT INTO material_replacement_sets 
                    (part_code, set_type, set_name)
                    VALUES (?, 'from', ?)
                """, (part_code, set_name))
                from_set_id = cursor.lastrowid
                
                # Создаём набор "to" (на что заменяем)
                cursor.execute("""
                    INSERT INTO material_replacement_sets 
                    (part_code, set_type, set_name)
                    VALUES (?, 'to', ?)
                """, (part_code, set_name))
                to_set_id = cursor.lastrowid
                
                # Добавляем материалы "from" в каталог и связываем с набором
                for order_idx, material in enumerate(from_materials):
                    material.part = part_code
                    material.is_part_of_set = True
                    material.replacement_set_id = from_set_id
                    
                    entry_id = self._add_entry_internal(cursor, material)
                    if entry_id:
                        cursor.execute("""
                            INSERT INTO material_set_items 
                            (set_id, catalog_entry_id, order_index)
                            VALUES (?, ?, ?)
                        """, (from_set_id, entry_id, order_idx))
                        # Обновляем кэш
                        if material.part not in self._by_part:
                            self._by_part[material.part] = []
                        self._by_part[material.part].append(material)
                        self._entries.append(material)
                
                # Добавляем материалы "to" в каталог и связываем с набором
                for order_idx, material in enumerate(to_materials):
                    material.part = part_code
                    material.is_part_of_set = True
                    material.replacement_set_id = to_set_id
                    
                    entry_id = self._add_entry_internal(cursor, material)
                    if entry_id:
                        cursor.execute("""
                            INSERT INTO material_set_items 
                            (set_id, catalog_entry_id, order_index)
                            VALUES (?, ?, ?)
                        """, (to_set_id, entry_id, order_idx))
                        # Обновляем кэш
                        if material.part not in self._by_part:
                            self._by_part[material.part] = []
                        self._by_part[material.part].append(material)
                        self._entries.append(material)
                
                conn.commit()
                
                # Обновляем кэш
                self._parts_cache = None
                
                return from_set_id  # Возвращаем ID первого набора
        except sqlite3.Error as e:
            print(f"Ошибка при создании набора замены: {e}")
            return None
    
    def get_replacement_sets_by_part(self, part_code: str) -> List[MaterialReplacementSet]:
        """Получить все наборы замены для детали"""
        sets = []
        try:
            with self.db_manager.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("""
                    SELECT id, part_code, set_type, set_name, created_at
                    FROM material_replacement_sets
                    WHERE part_code = ?
                    ORDER BY created_at DESC
                """, (part_code,))
                
                for row in cursor.fetchall():
                    created_at_value = row['created_at']
                    # SQLite возвращает строку для TIMESTAMP, оставляем как есть или конвертируем при необходимости
                    replacement_set = MaterialReplacementSet(
                        id=row['id'],
                        part_code=row['part_code'],
                        set_type=row['set_type'],
                        set_name=row['set_name'],
                        created_at=created_at_value
                    )
                    # Загружаем материалы набора
                    replacement_set.materials = self.get_materials_in_set(row['id'])
                    sets.append(replacement_set)
        except sqlite3.Error as e:
            print(f"Ошибка при получении наборов замены: {e}")
        
        return sets
    
    def get_materials_in_set(self, set_id: int) -> List[CatalogEntry]:
        """Получить материалы в наборе"""
        materials = []
        try:
            with self.db_manager.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("""
                    SELECT ce.id, ce.part_code, ce.workshop, ce.role, ce.before_name,
                           ce.unit, ce.norm, ce.comment, ce.is_part_of_set, ce.replacement_set_id,
                           msi.order_index
                    FROM catalog_entries ce
                    INNER JOIN material_set_items msi ON ce.id = msi.catalog_entry_id
                    WHERE msi.set_id = ?
                    ORDER BY msi.order_index
                """, (set_id,))
                
                for row in cursor.fetchall():
                    entry = CatalogEntry(
                        id=row['id'],
                        part=row['part_code'],
                        workshop=row['workshop'],
                        role=row['role'],
                        before_name=row['before_name'],
                        unit=row['unit'],
                        norm=row['norm'],
                        comment=row['comment'] or "",
                        is_part_of_set=bool(row['is_part_of_set'] or 0),
                        replacement_set_id=row['replacement_set_id']
                    )
                    materials.append(entry)
        except sqlite3.Error as e:
            print(f"Ошибка при получении материалов набора: {e}")
        
        return materials
    
    def get_replacement_set_by_id(self, set_id: int) -> Optional[MaterialReplacementSet]:
        """Получить набор замены по ID"""
        try:
            with self.db_manager.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("""
                    SELECT id, part_code, set_type, set_name, created_at
                    FROM material_replacement_sets
                    WHERE id = ?
                """, (set_id,))
                
                row = cursor.fetchone()
                if row:
                    created_at_value = row['created_at']
                    replacement_set = MaterialReplacementSet(
                        id=row['id'],
                        part_code=row['part_code'],
                        set_type=row['set_type'],
                        set_name=row['set_name'],
                        created_at=created_at_value
                    )
                    # Загружаем материалы набора
                    replacement_set.materials = self.get_materials_in_set(row['id'])
                    return replacement_set
        except sqlite3.Error as e:
            print(f"Ошибка при получении набора замены: {e}")
        
        return None
    
    def update_replacement_set(self, set_id: int, materials: List[CatalogEntry]) -> bool:
        """Обновить материалы в наборе замены"""
        try:
            with self.db_manager.get_connection() as conn:
                cursor = conn.cursor()
                
                # Получаем информацию о наборе
                cursor.execute("""
                    SELECT part_code, set_type FROM material_replacement_sets WHERE id = ?
                """, (set_id,))
                set_info = cursor.fetchone()
                if not set_info:
                    return False
                
                part_code = set_info['part_code']
                set_type = set_info['set_type']
                
                # Удаляем старые материалы набора из material_set_items и catalog_entries
                cursor.execute("""
                    SELECT catalog_entry_id FROM material_set_items WHERE set_id = ?
                """, (set_id,))
                old_entry_ids = [row['catalog_entry_id'] for row in cursor.fetchall()]
                
                # Удаляем связи
                cursor.execute("DELETE FROM material_set_items WHERE set_id = ?", (set_id,))
                
                # Удаляем записи каталога
                for entry_id in old_entry_ids:
                    cursor.execute("DELETE FROM catalog_entries WHERE id = ?", (entry_id,))
                
                # Добавляем новые материалы
                for order_idx, material in enumerate(materials):
                    material.part = part_code
                    material.is_part_of_set = True
                    material.replacement_set_id = set_id
                    
                    entry_id = self._add_entry_internal(cursor, material)
                    if entry_id:
                        cursor.execute("""
                            INSERT INTO material_set_items 
                            (set_id, catalog_entry_id, order_index)
                            VALUES (?, ?, ?)
                        """, (set_id, entry_id, order_idx))
                
                conn.commit()
                
                # Обновляем кэш
                self._parts_cache = None
                self._by_part = {}
                self._entries = []
                self.load()
                
                return True
        except sqlite3.Error as e:
            print(f"Ошибка при обновлении набора замены: {e}")
            return False
    
    def delete_replacement_set(self, set_id: int) -> bool:
        """Удалить набор замены и все связанные материалы"""
        try:
            with self.db_manager.get_connection() as conn:
                cursor = conn.cursor()
                
                # Получаем ID записей каталога, связанных с набором
                cursor.execute("""
                    SELECT catalog_entry_id FROM material_set_items WHERE set_id = ?
                """, (set_id,))
                entry_ids = [row['catalog_entry_id'] for row in cursor.fetchall()]
                
                # Удаляем связи
                cursor.execute("DELETE FROM material_set_items WHERE set_id = ?", (set_id,))
                
                # Удаляем записи каталога
                for entry_id in entry_ids:
                    cursor.execute("DELETE FROM catalog_entries WHERE id = ?", (entry_id,))
                
                # Удаляем сам набор
                cursor.execute("DELETE FROM material_replacement_sets WHERE id = ?", (set_id,))
                
                conn.commit()
                
                # Обновляем кэш
                self._parts_cache = None
                self._by_part = {}
                self._entries = []
                self.load()
                
                return True
        except sqlite3.Error as e:
            print(f"Ошибка при удалении набора замены: {e}")
            return False