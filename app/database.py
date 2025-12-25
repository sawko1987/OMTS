"""
Единый менеджер базы данных SQLite
"""
import sqlite3
from pathlib import Path
from typing import Optional
from contextlib import contextmanager

from app.config import DATABASE_PATH


class DatabaseManager:
    """Менеджер базы данных SQLite (синглтон)"""
    
    _instance: Optional['DatabaseManager'] = None
    _initialized: bool = False
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance
    
    def __init__(self):
        if not self._initialized:
            self.db_path = DATABASE_PATH
            self._initialized = True
    
    @contextmanager
    def get_connection(self):
        """Получить соединение с БД (контекстный менеджер)"""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        try:
            yield conn
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()
    
    def initialize(self):
        """Инициализировать схему базы данных"""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            
            # Таблица изделий
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS products (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL UNIQUE,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)
            
            # Таблица деталей
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS parts (
                    code TEXT PRIMARY KEY
                )
            """)
            
            # Таблица записей каталога
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS catalog_entries (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    part_code TEXT NOT NULL,
                    workshop TEXT NOT NULL,
                    role TEXT NOT NULL,
                    before_name TEXT NOT NULL,
                    unit TEXT NOT NULL,
                    norm REAL NOT NULL,
                    comment TEXT DEFAULT '',
                    is_part_of_set INTEGER DEFAULT 0,
                    replacement_set_id INTEGER,
                    FOREIGN KEY (part_code) REFERENCES parts(code)
                )
            """)
            
            # Таблица связей изделие-деталь
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS product_parts (
                    product_id INTEGER NOT NULL,
                    part_code TEXT NOT NULL,
                    PRIMARY KEY (product_id, part_code),
                    FOREIGN KEY (product_id) REFERENCES products(id),
                    FOREIGN KEY (part_code) REFERENCES parts(code)
                )
            """)
            
            # Таблица истории замен материалов
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS material_replacements (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    part_code TEXT NOT NULL,
                    workshop TEXT NOT NULL,
                    role TEXT NOT NULL,
                    before_name TEXT NOT NULL,
                    after_name TEXT NOT NULL,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)
            
            # Таблица нумерации документов
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS numbering (
                    year INTEGER PRIMARY KEY,
                    last_number INTEGER NOT NULL DEFAULT 0
                )
            """)
            
            # Таблица наборов материалов для замены
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS material_replacement_sets (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    part_code TEXT NOT NULL,
                    set_type TEXT NOT NULL CHECK(set_type IN ('from', 'to')),
                    set_name TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (part_code) REFERENCES parts(code)
                )
            """)
            
            # Таблица элементов наборов материалов
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS material_set_items (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    set_id INTEGER NOT NULL,
                    catalog_entry_id INTEGER NOT NULL,
                    order_index INTEGER NOT NULL DEFAULT 0,
                    FOREIGN KEY (set_id) REFERENCES material_replacement_sets(id) ON DELETE CASCADE,
                    FOREIGN KEY (catalog_entry_id) REFERENCES catalog_entries(id) ON DELETE CASCADE
                )
            """)
            
            # Таблица документов (извещений)
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS documents (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    document_number INTEGER NOT NULL,
                    year INTEGER NOT NULL,
                    data_json TEXT NOT NULL,
                    output_file_path TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE(document_number, year)
                )
            """)
            
            # Миграция существующих таблиц
            self._migrate_schema(cursor)
            
            # Создание индексов для оптимизации
            self._create_indexes(cursor)
            
            conn.commit()
    
    def _migrate_schema(self, cursor: sqlite3.Cursor):
        """Миграция схемы для существующих БД"""
        # Проверяем наличие новых полей в catalog_entries
        cursor.execute("PRAGMA table_info(catalog_entries)")
        columns = [row[1] for row in cursor.fetchall()]
        
        if 'is_part_of_set' not in columns:
            cursor.execute("ALTER TABLE catalog_entries ADD COLUMN is_part_of_set INTEGER DEFAULT 0")
        
        if 'replacement_set_id' not in columns:
            cursor.execute("ALTER TABLE catalog_entries ADD COLUMN replacement_set_id INTEGER")
    
    def _create_indexes(self, cursor: sqlite3.Cursor):
        """Создать индексы для оптимизации запросов"""
        indexes = [
            "CREATE INDEX IF NOT EXISTS idx_catalog_part_code ON catalog_entries(part_code)",
            "CREATE INDEX IF NOT EXISTS idx_catalog_search ON catalog_entries(part_code COLLATE NOCASE)",
            "CREATE INDEX IF NOT EXISTS idx_product_parts_product ON product_parts(product_id)",
            "CREATE INDEX IF NOT EXISTS idx_product_parts_part ON product_parts(part_code)",
            "CREATE INDEX IF NOT EXISTS idx_replacements_lookup ON material_replacements(part_code, workshop, role, before_name)",
            "CREATE INDEX IF NOT EXISTS idx_catalog_is_part_of_set ON catalog_entries(is_part_of_set)",
            "CREATE INDEX IF NOT EXISTS idx_catalog_replacement_set_id ON catalog_entries(replacement_set_id)",
            "CREATE INDEX IF NOT EXISTS idx_replacement_sets_part_code ON material_replacement_sets(part_code)",
            "CREATE INDEX IF NOT EXISTS idx_set_items_set_id ON material_set_items(set_id)",
            "CREATE INDEX IF NOT EXISTS idx_set_items_entry_id ON material_set_items(catalog_entry_id)",
            "CREATE INDEX IF NOT EXISTS idx_documents_number_year ON documents(document_number, year)",
            "CREATE INDEX IF NOT EXISTS idx_documents_year ON documents(year)",
        ]
        
        for index_sql in indexes:
            cursor.execute(index_sql)
    
    def table_exists(self, table_name: str) -> bool:
        """Проверить существование таблицы"""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT name FROM sqlite_master 
                WHERE type='table' AND name=?
            """, (table_name,))
            return cursor.fetchone() is not None
    
    def has_data(self, table_name: str) -> bool:
        """Проверить наличие данных в таблице"""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(f"SELECT COUNT(*) FROM {table_name}")
            count = cursor.fetchone()[0]
            return count > 0

