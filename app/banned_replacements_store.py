"""
Хранение информации о запрещённых заменах материалов
"""
from typing import Optional

from app.database import DatabaseManager
from app.models import CatalogEntry, BannedReplacementInfo


class BannedReplacementsStore:
    """Хранилище запрещённых замен материалов в SQLite"""

    def __init__(self, db_manager: DatabaseManager = None):
        self.db_manager = db_manager or DatabaseManager()

    def ban(self, entry: CatalogEntry, after_name: str = "", reason: str = "") -> bool:
        """Запретить замену материала (конкретную пару before_name -> after_name)"""
        try:
            with self.db_manager.get_connection() as conn:
                conn.execute("""
                    INSERT OR REPLACE INTO banned_replacements
                    (part_code, workshop, role, before_name, after_name, reason)
                    VALUES (?, ?, ?, ?, ?, ?)
                """, (entry.part, entry.workshop, entry.role, entry.before_name, after_name, reason))
                return True
        except Exception as e:
            print(f"Ошибка при добавлении запрета замены: {e}")
            return False

    def unban(self, entry: CatalogEntry, after_name: Optional[str] = None) -> bool:
        """
        Снять запрет с материала.
        Если after_name указан — снять только для конкретной пары.
        Если after_name = None — снять все запреты для этого материала (entry).
        """
        try:
            with self.db_manager.get_connection() as conn:
                if after_name is not None:
                    conn.execute("""
                        DELETE FROM banned_replacements
                        WHERE part_code = ? AND workshop = ? AND role = ? AND before_name = ? AND after_name = ?
                    """, (entry.part, entry.workshop, entry.role, entry.before_name, after_name))
                else:
                    conn.execute("""
                        DELETE FROM banned_replacements
                        WHERE part_code = ? AND workshop = ? AND role = ? AND before_name = ?
                    """, (entry.part, entry.workshop, entry.role, entry.before_name))
                return True
        except Exception as e:
            print(f"Ошибка при снятии запрета замены: {e}")
            return False

    def is_banned(self, entry: CatalogEntry, after_name: Optional[str] = None) -> bool:
        """
        Проверить, запрещена ли замена материала.
        Если after_name указан — проверить конкретную пару.
        Если after_name = None — проверить, есть ли хоть какой-то запрет для этого entry.
        """
        with self.db_manager.get_connection() as conn:
            if after_name is not None:
                cursor = conn.execute("""
                    SELECT COUNT(*) FROM banned_replacements
                    WHERE part_code = ? AND workshop = ? AND role = ? AND before_name = ? AND after_name = ?
                """, (entry.part, entry.workshop, entry.role, entry.before_name, after_name))
            else:
                cursor = conn.execute("""
                    SELECT COUNT(*) FROM banned_replacements
                    WHERE part_code = ? AND workshop = ? AND role = ? AND before_name = ?
                """, (entry.part, entry.workshop, entry.role, entry.before_name))
            return cursor.fetchone()[0] > 0

    def get_ban_info(self, entry: CatalogEntry, after_name: Optional[str] = None) -> Optional[BannedReplacementInfo]:
        """
        Получить информацию о запрете.
        Если after_name указан — искать конкретную пару.
        Если after_name = None — вернуть первый найденный запрет для entry.
        """
        with self.db_manager.get_connection() as conn:
            if after_name is not None:
                cursor = conn.execute("""
                    SELECT reason, banned_at FROM banned_replacements
                    WHERE part_code = ? AND workshop = ? AND role = ? AND before_name = ? AND after_name = ?
                    LIMIT 1
                """, (entry.part, entry.workshop, entry.role, entry.before_name, after_name))
            else:
                cursor = conn.execute("""
                    SELECT reason, banned_at FROM banned_replacements
                    WHERE part_code = ? AND workshop = ? AND role = ? AND before_name = ?
                    LIMIT 1
                """, (entry.part, entry.workshop, entry.role, entry.before_name))
            row = cursor.fetchone()
            if row:
                return BannedReplacementInfo(reason=row['reason'] or "", banned_at=row['banned_at'])
            return None

    def get_banned_for_part(self, part_code: str):
        """Получить все запреты для детали"""
        with self.db_manager.get_connection() as conn:
            cursor = conn.execute("""
                SELECT * FROM banned_replacements
                WHERE part_code = ?
                ORDER BY banned_at DESC
            """, (part_code,))
            return cursor.fetchall()

    def get_all_banned(self):
        """Получить все запреты"""
        with self.db_manager.get_connection() as conn:
            cursor = conn.execute("""
                SELECT * FROM banned_replacements
                ORDER BY part_code, banned_at DESC
            """)
            return cursor.fetchall()
