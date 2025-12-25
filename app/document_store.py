"""
Хранилище документов в SQLite
"""
from pathlib import Path
from typing import List, Optional, Tuple
from datetime import date
import sqlite3

from app.database import DatabaseManager
from app.models import DocumentData
from app.serialization import DocumentSerializer
from app.catalog_loader import CatalogLoader


class DocumentStore:
    """Хранилище документов"""
    
    def __init__(self, db_manager: DatabaseManager = None, catalog_loader: CatalogLoader = None):
        self.db_manager = db_manager or DatabaseManager()
        self.catalog_loader = catalog_loader or CatalogLoader(self.db_manager)
        self.serializer = DocumentSerializer()
    
    def save_document(self, document_data: DocumentData, output_file_path: Optional[str] = None) -> Optional[int]:
        """Сохранить документ в БД. Возвращает ID сохранённого документа или None при ошибке."""
        if not document_data.document_number:
            raise ValueError("Документ должен иметь номер")
        
        current_year = date.today().year
        data_json = self.serializer.serialize(document_data)
        
        try:
            with self.db_manager.get_connection() as conn:
                cursor = conn.cursor()
                
                # Проверяем, существует ли уже документ с таким номером и годом
                cursor.execute("""
                    SELECT id FROM documents 
                    WHERE document_number = ? AND year = ?
                """, (document_data.document_number, current_year))
                
                existing = cursor.fetchone()
                
                if existing:
                    # Обновляем существующий документ
                    cursor.execute("""
                        UPDATE documents 
                        SET data_json = ?, output_file_path = ?, updated_at = CURRENT_TIMESTAMP
                        WHERE id = ?
                    """, (data_json, output_file_path, existing['id']))
                    return existing['id']
                else:
                    # Создаём новый документ
                    cursor.execute("""
                        INSERT INTO documents (document_number, year, data_json, output_file_path)
                        VALUES (?, ?, ?, ?)
                    """, (document_data.document_number, current_year, data_json, output_file_path))
                    return cursor.lastrowid
        except sqlite3.Error as e:
            print(f"Ошибка при сохранении документа: {e}")
            return None
    
    def load_document(self, document_number: int, year: Optional[int] = None) -> Optional[DocumentData]:
        """Загрузить документ по номеру. Если год не указан, используется текущий год."""
        if year is None:
            year = date.today().year
        
        try:
            with self.db_manager.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("""
                    SELECT data_json FROM documents 
                    WHERE document_number = ? AND year = ?
                """, (document_number, year))
                
                row = cursor.fetchone()
                if row:
                    return self.serializer.deserialize(row['data_json'], self.catalog_loader)
        except sqlite3.Error as e:
            print(f"Ошибка при загрузке документа: {e}")
        except Exception as e:
            print(f"Ошибка при десериализации документа: {e}")
        
        return None
    
    def get_all_documents(self, year: Optional[int] = None) -> List[Tuple[int, int, str, Optional[str]]]:
        """Получить список всех документов: (document_number, year, created_at, output_file_path).
        Если год не указан, возвращаются документы за все годы."""
        try:
            with self.db_manager.get_connection() as conn:
                cursor = conn.cursor()
                
                if year:
                    cursor.execute("""
                        SELECT document_number, year, created_at, output_file_path
                        FROM documents
                        WHERE year = ?
                        ORDER BY document_number DESC
                    """, (year,))
                else:
                    cursor.execute("""
                        SELECT document_number, year, created_at, output_file_path
                        FROM documents
                        ORDER BY year DESC, document_number DESC
                    """)
                
                return [(row['document_number'], row['year'], row['created_at'], row['output_file_path']) 
                        for row in cursor.fetchall()]
        except sqlite3.Error as e:
            print(f"Ошибка при получении списка документов: {e}")
            return []
    
    def document_exists(self, document_number: int, year: Optional[int] = None) -> bool:
        """Проверить существование документа"""
        if year is None:
            year = date.today().year
        
        try:
            with self.db_manager.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("""
                    SELECT COUNT(*) FROM documents 
                    WHERE document_number = ? AND year = ?
                """, (document_number, year))
                return cursor.fetchone()[0] > 0
        except sqlite3.Error:
            return False
    
    def delete_document(self, document_number: int, year: Optional[int] = None) -> bool:
        """Удалить документ"""
        if year is None:
            year = date.today().year
        
        try:
            with self.db_manager.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("""
                    DELETE FROM documents 
                    WHERE document_number = ? AND year = ?
                """, (document_number, year))
                return cursor.rowcount > 0
        except sqlite3.Error as e:
            print(f"Ошибка при удалении документа: {e}")
            return False

