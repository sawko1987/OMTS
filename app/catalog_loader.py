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
        entries = self._by_part.get(part, [])
        return self._deduplicate_materials_by_name(entries)

    def _deduplicate_materials_by_name(self, entries: List[CatalogEntry]) -> List[CatalogEntry]:
        """
        Дедупликация материалов по ключу (workshop, before_name, unit).

        Правила:
        - группируем по (workshop, before_name, unit) - материалы с одинаковым названием,
          но разными единицами измерения считаются разными материалами
        - если в группе есть записи с заполненным цехом (не NULL/не пустая) — берём любую из них (первую)
        - иначе берём любую без цеха (первую)
        """
        if not entries:
            return []

        def _norm_text(s: object) -> str:
            return (str(s) if s is not None else "").strip()

        def _has_workshop(e: CatalogEntry) -> bool:
            return bool(_norm_text(e.workshop))

        # Ключ: (workshop, before_name, unit)
        def _make_key(e: CatalogEntry) -> tuple[str, str, str]:
            return (
                _norm_text(e.workshop),
                _norm_text(e.before_name),
                _norm_text(e.unit)
            )

        chosen_by_key: dict[tuple[str, str, str], CatalogEntry] = {}
        chosen_has_workshop: dict[tuple[str, str, str], bool] = {}
        order: list[tuple[str, str, str]] = []

        for e in entries:
            key = _make_key(e)
            if key not in chosen_by_key:
                chosen_by_key[key] = e
                chosen_has_workshop[key] = _has_workshop(e)
                order.append(key)
                continue

            # Если ранее выбран вариант без цеха, а текущий — с цехом, то заменяем
            if not chosen_has_workshop[key] and _has_workshop(e):
                chosen_by_key[key] = e
                chosen_has_workshop[key] = True

        return [chosen_by_key[k] for k in order]

    def get_entries_by_replacement_set_id(self, set_id: int) -> List[CatalogEntry]:
        """
        Получить материалы конкретного набора (ordered).

        Это обёртка над get_materials_in_set(), добавлена для явности в местах UI.
        """
        return self.get_materials_in_set(set_id)

    def get_entries_by_set_type(self, set_type: str) -> List[CatalogEntry]:
        """
        Получить "каталог" материалов по типу набора ('from'/'to') для ВСЕХ деталей.

        В БД материалы наборов хранятся как отдельные записи catalog_entries с is_part_of_set=1.
        Пользователь ожидает общий список материалов "до"/"после", доступный при работе с любой деталью.
        """
        if set_type not in ("from", "to"):
            return []

        materials: List[CatalogEntry] = []
        try:
            with self.db_manager.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(
                    """
                    SELECT
                        ce.id, ce.part_code, ce.workshop, ce.role, ce.before_name,
                        ce.unit, ce.norm, ce.comment, ce.is_part_of_set, ce.replacement_set_id,
                        mrs.id AS set_id,
                        msi.order_index
                    FROM material_replacement_sets mrs
                    INNER JOIN material_set_items msi ON msi.set_id = mrs.id
                    INNER JOIN catalog_entries ce ON ce.id = msi.catalog_entry_id
                    WHERE mrs.set_type = ?
                    ORDER BY mrs.id DESC, msi.order_index ASC
                    """,
                    (set_type,),
                )
                for row in cursor.fetchall():
                    entry = CatalogEntry(
                        id=row["id"],
                        part=row["part_code"],
                        workshop=row["workshop"],
                        role=row["role"],
                        before_name=row["before_name"],
                        unit=row["unit"],
                        norm=row["norm"],
                        comment=row["comment"] or "",
                        is_part_of_set=bool(row["is_part_of_set"] or 0),
                        replacement_set_id=row["replacement_set_id"],
                    )
                    materials.append(entry)
        except sqlite3.Error as e:
            print(f"Ошибка при получении материалов по типу набора: {e}")
            return []

        return self._deduplicate_materials_by_name(materials)

    def get_entries_by_part_and_set_type(self, part_code: str, set_type: str) -> List[CatalogEntry]:
        """
        Получить материалы детали, относящиеся к наборам заданного типа ('from'/'to').

        Используется в диалоге "Выбрать из каталога" для раздельных каталогов
        материалов "до" и "после".
        """
        part_code = (part_code or "").strip()
        if not part_code:
            return []
        if set_type not in ("from", "to"):
            return []

        materials: List[CatalogEntry] = []
        try:
            with self.db_manager.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(
                    """
                    SELECT
                        ce.id, ce.part_code, ce.workshop, ce.role, ce.before_name,
                        ce.unit, ce.norm, ce.comment, ce.is_part_of_set, ce.replacement_set_id,
                        mrs.id AS set_id,
                        msi.order_index
                    FROM material_replacement_sets mrs
                    INNER JOIN material_set_items msi ON msi.set_id = mrs.id
                    INNER JOIN catalog_entries ce ON ce.id = msi.catalog_entry_id
                    WHERE mrs.part_code = ? AND mrs.set_type = ?
                    ORDER BY mrs.id DESC, msi.order_index ASC
                    """,
                    (part_code, set_type),
                )
                for row in cursor.fetchall():
                    entry = CatalogEntry(
                        id=row["id"],
                        part=row["part_code"],
                        workshop=row["workshop"],
                        role=row["role"],
                        before_name=row["before_name"],
                        unit=row["unit"],
                        norm=row["norm"],
                        comment=row["comment"] or "",
                        is_part_of_set=bool(row["is_part_of_set"] or 0),
                        replacement_set_id=row["replacement_set_id"],
                    )
                    materials.append(entry)
        except sqlite3.Error as e:
            print(f"Ошибка при получении материалов по типу набора: {e}")
            return []

        return self._deduplicate_materials_by_name(materials)

    # ----------------------------
    # Глобальный словарь замен (до -> варианты после)
    # ----------------------------

    def get_replacement_dictionary_options(self, from_entry: CatalogEntry) -> List[CatalogEntry]:
        """
        Получить варианты материалов 'после' для заданного материала 'до'
        (ключ: цех + наименование + ед.изм.).
        """
        fw = (from_entry.workshop or "").strip()
        fn = (from_entry.before_name or "").strip()
        fu = (from_entry.unit or "").strip()
        if not (fw and fn and fu):
            return []

        results: List[CatalogEntry] = []
        try:
            with self.db_manager.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(
                    """
                    SELECT
                        to_workshop, to_name, to_unit, to_norm, to_comment
                    FROM material_replacement_dictionary
                    WHERE from_workshop = ? AND from_name = ? AND from_unit = ?
                    ORDER BY created_at DESC, id DESC
                    """,
                    (fw, fn, fu),
                )
                for row in cursor.fetchall():
                    results.append(
                        CatalogEntry(
                            part="",
                            workshop=row["to_workshop"] or "",
                            role="",
                            before_name=row["to_name"] or "",
                            unit=row["to_unit"] or "",
                            norm=float(row["to_norm"] or 0.0),
                            comment=row["to_comment"] or "",
                            id=None,
                            is_part_of_set=False,
                            replacement_set_id=None,
                        )
                    )
        except sqlite3.Error as e:
            print(f"Ошибка при получении словаря замен: {e}")
            return []

        return results

    def add_replacement_dictionary_link(self, from_entry: CatalogEntry, to_entry: CatalogEntry) -> bool:
        """
        Добавить связь в словарь замен: материал 'до' -> вариант материала 'после'.
        Дубликаты подавляются UNIQUE индексом.
        """
        fw = (from_entry.workshop or "").strip()
        fn = (from_entry.before_name or "").strip()
        fu = (from_entry.unit or "").strip()

        tw = (to_entry.workshop or "").strip()
        tn = (to_entry.before_name or "").strip()
        tu = (to_entry.unit or "").strip()
        tnorm = float(to_entry.norm or 0.0)
        tcomment = (to_entry.comment or "").strip()

        if not (fw and fn and fu and tw and tn and tu):
            return False

        try:
            with self.db_manager.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(
                    """
                    INSERT OR IGNORE INTO material_replacement_dictionary
                    (from_workshop, from_name, from_unit, to_workshop, to_name, to_unit, to_norm, to_comment)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (fw, fn, fu, tw, tn, tu, tnorm, tcomment),
                )
                conn.commit()
                # INSERT OR IGNORE: rowcount==1 только если реально вставили
                return cursor.rowcount > 0
        except sqlite3.Error as e:
            print(f"Ошибка при добавлении связи в словарь замен: {e}")
            return False

    def delete_replacement_dictionary_link(self, from_entry: CatalogEntry, to_entry: CatalogEntry) -> bool:
        """
        Удалить связь из словаря замен: материал 'до' -> вариант материала 'после'.
        """
        fw = (from_entry.workshop or "").strip()
        fn = (from_entry.before_name or "").strip()
        fu = (from_entry.unit or "").strip()

        tw = (to_entry.workshop or "").strip()
        tn = (to_entry.before_name or "").strip()
        tu = (to_entry.unit or "").strip()
        tnorm = float(to_entry.norm or 0.0)
        tcomment = (to_entry.comment or "").strip()

        if not (fw and fn and fu and tw and tn and tu):
            return False

        try:
            with self.db_manager.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(
                    """
                    DELETE FROM material_replacement_dictionary
                    WHERE
                        from_workshop = ? AND from_name = ? AND from_unit = ?
                        AND to_workshop = ? AND to_name = ? AND to_unit = ?
                        AND to_norm = ? AND COALESCE(to_comment, '') = ?
                    """,
                    (fw, fn, fu, tw, tn, tu, tnorm, tcomment),
                )
                conn.commit()
                return cursor.rowcount > 0
        except sqlite3.Error as e:
            print(f"Ошибка при удалении связи из словаря замен: {e}")
            return False
    
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
        return [e for e in entries if (e.workshop or "").strip() == (workshop or "").strip()]
    
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
            # Норма обнуляется только для записей справочника (is_part_of_set=0),
            # для материалов в наборах (is_part_of_set=1) норма сохраняется как есть
            is_part_of_set = 1 if entry.is_part_of_set else 0
            norm_value = 0.0 if not entry.is_part_of_set else entry.norm
            
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
                norm_value,
                entry.comment,
                is_part_of_set,
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
                
                # Проверяем, является ли запись частью набора
                cursor.execute("""
                    SELECT is_part_of_set FROM catalog_entries WHERE id = ?
                """, (entry_id,))
                row = cursor.fetchone()
                is_part_of_set = row['is_part_of_set'] if row else 0
                
                # Норма обнуляется только для записей справочника (is_part_of_set=0),
                # для материалов в наборах (is_part_of_set=1) норма сохраняется как есть
                norm_value = 0.0 if not is_part_of_set else entry.norm
                
                cursor.execute("""
                    UPDATE catalog_entries
                    SET part_code = ?, workshop = ?, role = ?, before_name = ?,
                        unit = ?, norm = ?, comment = ?
                    WHERE id = ?
                """, (
                    entry.part, entry.workshop, entry.role, entry.before_name,
                    entry.unit, norm_value, entry.comment, entry_id
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
    
    def delete_part(self, part_code: str) -> bool:
        """Полностью удалить деталь из базы данных со всеми связанными данными"""
        try:
            with self.db_manager.get_connection() as conn:
                cursor = conn.cursor()
                
                # Проверяем, существует ли деталь
                cursor.execute("""
                    SELECT COUNT(*) FROM catalog_entries WHERE part_code = ?
                """, (part_code,))
                if cursor.fetchone()[0] == 0:
                    logger.warning(f"Деталь '{part_code}' не найдена в базе данных")
                    return False
                
                # Удаляем все записи из catalog_entries для детали
                # (это автоматически удалит material_set_items благодаря CASCADE)
                cursor.execute("""
                    DELETE FROM catalog_entries WHERE part_code = ?
                """, (part_code,))
                deleted_entries = cursor.rowcount
                logger.info(f"Удалено {deleted_entries} записей из catalog_entries для детали '{part_code}'")
                
                # Удаляем все наборы замены (material_set_items удалятся автоматически благодаря CASCADE)
                cursor.execute("""
                    DELETE FROM material_replacement_sets WHERE part_code = ?
                """, (part_code,))
                deleted_sets = cursor.rowcount
                logger.info(f"Удалено {deleted_sets} наборов замены для детали '{part_code}'")
                
                # Удаляем связи с изделиями
                cursor.execute("""
                    DELETE FROM product_parts WHERE part_code = ?
                """, (part_code,))
                deleted_links = cursor.rowcount
                logger.info(f"Удалено {deleted_links} связей с изделиями для детали '{part_code}'")
                
                # Удаляем историю замен
                cursor.execute("""
                    DELETE FROM material_replacements WHERE part_code = ?
                """, (part_code,))
                deleted_history = cursor.rowcount
                logger.info(f"Удалено {deleted_history} записей истории замен для детали '{part_code}'")
                
                # Удаляем запись из parts
                cursor.execute("""
                    DELETE FROM parts WHERE code = ?
                """, (part_code,))
                deleted_part = cursor.rowcount
                logger.info(f"Удалена запись из parts для детали '{part_code}'")
                
                conn.commit()
                
                # Обновляем кэш
                self._parts_cache = None
                self._by_part = {}
                self._entries = []
                self.load()
                
                logger.info(f"Деталь '{part_code}' успешно удалена из базы данных")
                return True
        except sqlite3.Error as e:
            logger.error(f"Ошибка при удалении детали '{part_code}': {e}", exc_info=True)
            return False

    # ----------------------------
    # Операции с парами наборов (from+to)
    # ----------------------------

    def _find_matching_from_set(
        self,
        from_sets: List[MaterialReplacementSet],
        selected_to_set: Optional[MaterialReplacementSet],
    ) -> Optional[MaterialReplacementSet]:
        """
        Подобрать один набор 'from', соответствующий выбранному набору 'to'.

        Правила:
        - если to_set задан и есть совпадение по set_name — берём его
        - иначе если to_set задан и существует from с id == to_id-1 — берём его
        - иначе фоллбек: ближайший по id
        """
        if not from_sets:
            return None
        if selected_to_set is None:
            return from_sets[0]

        to_name = selected_to_set.set_name or ""
        if to_name:
            same_name = [s for s in from_sets if (s.set_name or "") == to_name]
            if same_name:
                return same_name[0]

        if selected_to_set.id is not None:
            id_match = [s for s in from_sets if s.id == (selected_to_set.id - 1)]
            if id_match:
                return id_match[0]

            def _dist(s: MaterialReplacementSet) -> int:
                return abs((s.id or 0) - selected_to_set.id)  # type: ignore[arg-type]

            return sorted(from_sets, key=_dist)[0]

        return from_sets[0]

    def find_matching_from_set(
        self,
        from_sets: List[MaterialReplacementSet],
        selected_to_set: Optional[MaterialReplacementSet],
    ) -> Optional[MaterialReplacementSet]:
        """Публичная обёртка для подбора from набора под выбранный to набор."""
        return self._find_matching_from_set(from_sets, selected_to_set)

    def get_replacement_pair_by_to_id(
        self, to_set_id: int
    ) -> tuple[Optional[MaterialReplacementSet], Optional[MaterialReplacementSet]]:
        """
        Получить пару (from_set, to_set) по ID набора 'to'.
        Возвращает (from_set, to_set). Если to_set не найден — (None, None).
        """
        to_set = self.get_replacement_set_by_id(to_set_id)
        if not to_set:
            return (None, None)

        all_sets = self.get_replacement_sets_by_part(to_set.part_code)
        from_sets = [s for s in all_sets if s.set_type == "from"]
        from_set = self._find_matching_from_set(from_sets, to_set)
        return (from_set, to_set)

    def update_replacement_set_name(self, set_id: int, set_name: Optional[str]) -> bool:
        """Обновить set_name у одного набора"""
        try:
            with self.db_manager.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(
                    "UPDATE material_replacement_sets SET set_name = ? WHERE id = ?",
                    (set_name, set_id),
                )
                conn.commit()
                return cursor.rowcount > 0
        except sqlite3.Error as e:
            print(f"Ошибка при обновлении названия набора: {e}")
            return False

    def update_set_name_for_pair(self, to_set_id: int, set_name: Optional[str]) -> bool:
        """Обновить set_name для пары наборов (from+to)"""
        from_set, to_set = self.get_replacement_pair_by_to_id(to_set_id)
        if not to_set:
            return False

        ok = True
        if from_set and from_set.id:
            ok = ok and self.update_replacement_set_name(from_set.id, set_name)
        if to_set.id:
            ok = ok and self.update_replacement_set_name(to_set.id, set_name)
        return ok

    def delete_replacement_pair(self, to_set_id: int) -> bool:
        """Удалить пару наборов (from+to) по ID набора 'to'"""
        from_set, to_set = self.get_replacement_pair_by_to_id(to_set_id)
        if not to_set or not to_set.id:
            return False

        ok = True
        # Сначала удаляем to, затем from (порядок не критичен)
        ok = ok and self.delete_replacement_set(to_set.id)
        if from_set and from_set.id:
            ok = ok and self.delete_replacement_set(from_set.id)
        return ok

    def _clone_materials(self, materials: List[CatalogEntry]) -> List[CatalogEntry]:
        """Создать копии материалов для создания нового набора"""
        cloned: List[CatalogEntry] = []
        for m in materials or []:
            cloned.append(
                CatalogEntry(
                    part="",
                    workshop=m.workshop,
                    role=m.role,
                    before_name=m.before_name,
                    unit=m.unit,
                    norm=m.norm,
                    comment=m.comment,
                    id=None,
                    is_part_of_set=False,
                    replacement_set_id=None,
                )
            )
        return cloned

    def clone_replacement_pair(self, to_set_id: int, new_set_name: Optional[str]) -> Optional[int]:
        """
        Клонировать пару (from+to). Возвращает ID нового набора 'to' или None.
        """
        from_set, to_set = self.get_replacement_pair_by_to_id(to_set_id)
        if not to_set:
            return None

        from_materials = self._clone_materials(from_set.materials if from_set else [])
        to_materials = self._clone_materials(to_set.materials or [])

        new_from_id = self.add_replacement_set(to_set.part_code, from_materials, to_materials, set_name=new_set_name)
        if not new_from_id:
            return None
        return new_from_id + 1  # новый to создается сразу после from

    def split_replacement_pair_copy(
        self,
        to_set_id: int,
        from_indexes: List[int],
        to_indexes: List[int],
        new_set_name: Optional[str],
    ) -> Optional[int]:
        """
        Создать новую пару (from+to) из поднабора материалов исходной пары, копированием.
        from_indexes/to_indexes — индексы строк в соответствующих списках materials.
        Возвращает ID нового набора 'to' или None.
        """
        from_set, to_set = self.get_replacement_pair_by_to_id(to_set_id)
        if not to_set:
            return None

        base_from = from_set.materials if from_set else []
        base_to = to_set.materials or []

        picked_from = [base_from[i] for i in from_indexes if 0 <= i < len(base_from)]
        picked_to = [base_to[i] for i in to_indexes if 0 <= i < len(base_to)]

        new_from_materials = self._clone_materials(picked_from)
        new_to_materials = self._clone_materials(picked_to)

        new_from_id = self.add_replacement_set(
            to_set.part_code,
            new_from_materials,
            new_to_materials,
            set_name=new_set_name,
        )
        if not new_from_id:
            return None
        return new_from_id + 1

    def clear_norms_in_catalog(self) -> bool:
        """
        Обнулить все нормы в справочнике каталога.
        Норма не должна храниться в справочнике, так как она разная для каждой детали.
        """
        try:
            with self.db_manager.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("""
                    UPDATE catalog_entries
                    SET norm = 0.0
                    WHERE norm != 0.0
                """)
                conn.commit()
                
                # Обновляем кэш
                self._parts_cache = None
                self._by_part = {}
                self._entries = []
                self.load()
                
                logger.info(f"Обнулено норм в справочнике: {cursor.rowcount}")
                return True
        except sqlite3.Error as e:
            logger.error(f"Ошибка при обнулении норм в справочнике: {e}")
            return False