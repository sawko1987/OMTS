"""
Хранение истории замен материалов в SQLite
"""
from pathlib import Path
from typing import List
import sqlite3

from app.config import HISTORY_FILE
from app.models import CatalogEntry
from app.database import DatabaseManager


class HistoryStore:
    """Хранилище истории замен в SQLite"""
    
    def __init__(self, db_manager: DatabaseManager = None):
        self.db_manager = db_manager or DatabaseManager()
    
    def add_replacement(self, entry: CatalogEntry, after_name: str):
        """Добавить вариант замены в историю"""
        if not after_name or not after_name.strip():
            return
        
        try:
            with self.db_manager.get_connection() as conn:
                cursor = conn.cursor()
                
                # Проверяем, нет ли уже такой замены
                cursor.execute("""
                    SELECT COUNT(*) FROM material_replacements
                    WHERE part_code = ? AND workshop = ? AND role = ? 
                    AND before_name = ? AND after_name = ?
                """, (entry.part, entry.workshop, entry.role, entry.before_name, after_name))
                
                if cursor.fetchone()[0] == 0:
                    # Добавляем только если такого варианта ещё нет
                    cursor.execute("""
                        INSERT INTO material_replacements
                        (part_code, workshop, role, before_name, after_name)
                        VALUES (?, ?, ?, ?, ?)
                    """, (entry.part, entry.workshop, entry.role, entry.before_name, after_name))
                    conn.commit()
        except sqlite3.Error as e:
            print(f"Ошибка при добавлении замены в историю: {e}")
    
    def get_suggestions(self, entry: CatalogEntry) -> List[str]:
        """Получить предложения замен для записи каталога"""
        with self.db_manager.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT DISTINCT after_name
                FROM material_replacements
                WHERE part_code = ? AND workshop = ? AND role = ? AND before_name = ?
                ORDER BY created_at DESC
                LIMIT 10
            """, (entry.part, entry.workshop, entry.role, entry.before_name))
            
            return [row['after_name'] for row in cursor.fetchall()]
    
    def get_suggestions_for_part_role(self, part: str, workshop: str, role: str) -> List[str]:
        """Получить все предложения для детали и типа позиции (по всем материалам 'до')"""
        with self.db_manager.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT DISTINCT after_name
                FROM material_replacements
                WHERE part_code = ? AND workshop = ? AND role = ?
                ORDER BY created_at DESC
            """, (part, workshop, role))
            
            return [row['after_name'] for row in cursor.fetchall()]
