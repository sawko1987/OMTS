"""
Менеджер изделий и связей изделие-деталь
"""
from typing import List, Optional
import sqlite3

from app.database import DatabaseManager


class ProductStore:
    """Хранилище изделий и связей с деталями"""
    
    def __init__(self, db_manager: DatabaseManager = None):
        self.db_manager = db_manager or DatabaseManager()
    
    def get_all_products(self) -> List[tuple]:
        """Получить список всех изделий (id, name)"""
        with self.db_manager.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT id, name FROM products ORDER BY name
            """)
            return [(row['id'], row['name']) for row in cursor.fetchall()]
    
    def get_product_by_name(self, name: str) -> Optional[int]:
        """Получить ID изделия по названию"""
        with self.db_manager.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT id FROM products WHERE name = ?", (name,))
            row = cursor.fetchone()
            return row['id'] if row else None
    
    def add_product(self, name: str) -> Optional[int]:
        """Добавить новое изделие"""
        if not name or not name.strip():
            return None
        
        name = name.strip()
        
        try:
            with self.db_manager.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO products (name) VALUES (?)
                """, (name,))
                conn.commit()
                return cursor.lastrowid
        except sqlite3.IntegrityError:
            # Изделие уже существует
            return self.get_product_by_name(name)
        except sqlite3.Error as e:
            print(f"Ошибка при добавлении изделия: {e}")
            return None
    
    def get_parts_by_product(self, product_id: int) -> List[str]:
        """Получить список деталей, привязанных к изделию"""
        with self.db_manager.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT pp.part_code
                FROM product_parts pp
                WHERE pp.product_id = ?
                ORDER BY pp.part_code
            """, (product_id,))
            return [row['part_code'] for row in cursor.fetchall()]
    
    def get_parts_by_product_name(self, product_name: str) -> List[str]:
        """Получить список деталей по названию изделия"""
        product_id = self.get_product_by_name(product_name)
        if product_id is None:
            return []
        return self.get_parts_by_product(product_id)
    
    def link_part_to_product(self, product_id: int, part_code: str) -> bool:
        """Привязать деталь к изделию"""
        try:
            with self.db_manager.get_connection() as conn:
                cursor = conn.cursor()
                
                # Убеждаемся, что деталь существует в таблице parts
                cursor.execute("INSERT OR IGNORE INTO parts (code) VALUES (?)", (part_code,))
                
                # Создаём связь
                cursor.execute("""
                    INSERT OR IGNORE INTO product_parts (product_id, part_code)
                    VALUES (?, ?)
                """, (product_id, part_code))
                conn.commit()
                return True
        except sqlite3.Error as e:
            print(f"Ошибка при привязке детали к изделию: {e}")
            return False
    
    def link_part_to_product_by_name(self, product_name: str, part_code: str) -> bool:
        """Привязать деталь к изделию по названию"""
        product_id = self.get_product_by_name(product_name)
        if product_id is None:
            return False
        return self.link_part_to_product(product_id, part_code)
    
    def unlink_part_from_product(self, product_id: int, part_code: str) -> bool:
        """Отвязать деталь от изделия"""
        try:
            with self.db_manager.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("""
                    DELETE FROM product_parts
                    WHERE product_id = ? AND part_code = ?
                """, (product_id, part_code))
                conn.commit()
                return True
        except sqlite3.Error as e:
            print(f"Ошибка при отвязке детали от изделия: {e}")
            return False
    
    def is_part_linked_to_product(self, product_id: int, part_code: str) -> bool:
        """Проверить, привязана ли деталь к изделию"""
        with self.db_manager.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT COUNT(*) FROM product_parts
                WHERE product_id = ? AND part_code = ?
            """, (product_id, part_code))
            return cursor.fetchone()[0] > 0
    
    def bulk_link_parts_to_products(self, product_ids: List[int], part_codes: List[str]) -> int:
        """
        Массовая привязка деталей к изделиям
        
        Args:
            product_ids: Список ID изделий
            part_codes: Список кодов деталей
        
        Returns:
            Количество созданных связей
        """
        if not product_ids or not part_codes:
            return 0
        
        try:
            with self.db_manager.get_connection() as conn:
                cursor = conn.cursor()
                
                # Убеждаемся, что все детали существуют в таблице parts
                for part_code in part_codes:
                    cursor.execute("INSERT OR IGNORE INTO parts (code) VALUES (?)", (part_code,))
                
                # Создаём все связи
                created_count = 0
                for product_id in product_ids:
                    for part_code in part_codes:
                        cursor.execute("""
                            INSERT OR IGNORE INTO product_parts (product_id, part_code)
                            VALUES (?, ?)
                        """, (product_id, part_code))
                        if cursor.rowcount > 0:
                            created_count += 1
                
                conn.commit()
                return created_count
        except sqlite3.Error as e:
            print(f"Ошибка при массовой привязке деталей к изделиям: {e}")
            return 0
    
    def bulk_unlink_parts_from_products(self, product_ids: List[int], part_codes: List[str]) -> int:
        """
        Массовая отвязка деталей от изделий
        
        Args:
            product_ids: Список ID изделий
            part_codes: Список кодов деталей
        
        Returns:
            Количество удалённых связей
        """
        if not product_ids or not part_codes:
            return 0
        
        try:
            with self.db_manager.get_connection() as conn:
                cursor = conn.cursor()
                
                deleted_count = 0
                for product_id in product_ids:
                    for part_code in part_codes:
                        cursor.execute("""
                            DELETE FROM product_parts
                            WHERE product_id = ? AND part_code = ?
                        """, (product_id, part_code))
                        if cursor.rowcount > 0:
                            deleted_count += 1
                
                conn.commit()
                return deleted_count
        except sqlite3.Error as e:
            print(f"Ошибка при массовой отвязке деталей от изделий: {e}")
            return 0
    
    def get_all_parts(self) -> List[str]:
        """Получить список всех деталей в базе"""
        with self.db_manager.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT DISTINCT code FROM parts ORDER BY code
            """)
            return [row['code'] for row in cursor.fetchall()]

