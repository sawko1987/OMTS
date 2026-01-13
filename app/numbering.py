"""
Управление нумерацией документов
"""
import json
import logging
from pathlib import Path
from datetime import date
from typing import Optional

from app.config import NUMBERING_FILE, DATABASE_PATH
from app.database import DatabaseManager
from app.settings_manager import SettingsManager

logger = logging.getLogger(__name__)


class NumberingManager:
    """Менеджер нумерации документов (работает с БД, с fallback на JSON)"""
    
    def __init__(self, db_manager: DatabaseManager = None, numbering_file: Path = None):
        self.db_manager = db_manager or DatabaseManager()
        self.numbering_file = numbering_file or NUMBERING_FILE
        self._use_db = DATABASE_PATH.exists()
        self.settings_manager = SettingsManager()
    
    def get_next_number(self, year: Optional[int] = None) -> int:
        """Получить следующий номер документа
        
        Args:
            year: Год для нумерации. Если не указан, используется текущий год.
        """
        target_year = year if year is not None else date.today().year
        
        if self._use_db:
            try:
                with self.db_manager.get_connection() as conn:
                    cursor = conn.cursor()
                    
                    # Проверяем, есть ли запись для указанного года
                    cursor.execute("SELECT last_number FROM numbering WHERE year = ?", (target_year,))
                    row = cursor.fetchone()
                    
                    if row:
                        last_number = row['last_number']
                        new_number = last_number + 1
                        cursor.execute(
                            "UPDATE numbering SET last_number = ? WHERE year = ?",
                            (new_number, target_year)
                        )
                    else:
                        # Создаём запись для нового года с начальным номером из настроек
                        starting_number = self.settings_manager.get_starting_number()
                        new_number = starting_number
                        cursor.execute(
                            "INSERT INTO numbering (year, last_number) VALUES (?, ?)",
                            (target_year, new_number)
                        )
                    
                    conn.commit()
                    return new_number
            except Exception:
                # Fallback на JSON
                self._use_db = False
                return self._get_next_number_json(target_year)
        else:
            return self._get_next_number_json(target_year)
    
    def get_current_number(self, year: Optional[int] = None) -> int:
        """Получить текущий номер (без увеличения)
        
        Проверяет как таблицу numbering, так и реальные документы в таблице documents,
        чтобы вернуть правильный следующий номер.
        
        Args:
            year: Год для нумерации. Если не указан, используется текущий год.
        """
        target_year = year if year is not None else date.today().year
        logger.info(f"[get_current_number] Запрос номера для года: {target_year} (передан year={year})")
        
        if self._use_db:
            try:
                with self.db_manager.get_connection() as conn:
                    cursor = conn.cursor()
                    
                    # Получаем номер из таблицы numbering
                    cursor.execute("SELECT last_number FROM numbering WHERE year = ?", (target_year,))
                    numbering_row = cursor.fetchone()
                    numbering_number = numbering_row['last_number'] + 1 if numbering_row else None
                    
                    # Проверяем реальные документы в таблице documents
                    cursor.execute("""
                        SELECT MAX(document_number) as max_number 
                        FROM documents 
                        WHERE year = ?
                    """, (target_year,))
                    documents_row = cursor.fetchone()
                    documents_max = documents_row['max_number'] if documents_row and documents_row['max_number'] else None
                    
                    # Используем максимальное значение из двух источников
                    if numbering_number is not None and documents_max is not None:
                        result = max(numbering_number, documents_max + 1)
                        logger.info(f"[get_current_number] Для года {target_year}: numbering={numbering_number}, documents_max={documents_max}, возвращаем {result}")
                    elif numbering_number is not None:
                        result = numbering_number
                        logger.info(f"[get_current_number] Для года {target_year}: только numbering={numbering_number}, возвращаем {result}")
                    elif documents_max is not None:
                        result = documents_max + 1
                        logger.info(f"[get_current_number] Для года {target_year}: только documents_max={documents_max}, возвращаем {result}")
                    else:
                        # Если записей нет, возвращаем начальный номер из настроек
                        starting_number = self.settings_manager.get_starting_number()
                        logger.info(f"[get_current_number] Запись для года {target_year} не найдена, возвращаем начальный номер: {starting_number}")
                        return starting_number
                    
                    return result
            except Exception as e:
                logger.warning(f"[get_current_number] Ошибка при работе с БД: {e}, переключаемся на JSON")
                self._use_db = False
                return self._get_current_number_json(target_year)
        else:
            logger.info(f"[get_current_number] Используется JSON fallback для года {target_year}")
            return self._get_current_number_json(target_year)
    
    def _get_next_number_json(self, year: int) -> int:
        """Получить следующий номер из JSON (fallback)
        
        Args:
            year: Год для нумерации
        """
        starting_number = self.settings_manager.get_starting_number()
        data = {"year": year, "last": starting_number - 1}
        
        if self.numbering_file.exists():
            try:
                with open(self.numbering_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except (json.JSONDecodeError, FileNotFoundError):
                pass
        
        # Если год изменился, сбрасываем счётчик на начальный номер из настроек
        if data.get("year") != year:
            data["year"] = year
            data["last"] = starting_number - 1
        
        # Увеличиваем номер
        data["last"] += 1
        self._save_json(data)
        
        return data["last"]
    
    def _get_current_number_json(self, year: int) -> int:
        """Получить текущий номер из JSON (fallback)
        
        Args:
            year: Год для нумерации
        """
        starting_number = self.settings_manager.get_starting_number()
        data = {"year": year, "last": starting_number - 1}
        
        if self.numbering_file.exists():
            try:
                with open(self.numbering_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except (json.JSONDecodeError, FileNotFoundError):
                pass
        
        if data.get("year") != year:
            return starting_number
        
        return data.get("last", starting_number - 1) + 1
    
    def _save_json(self, data: dict):
        """Сохранить данные в JSON"""
        self.numbering_file.parent.mkdir(exist_ok=True, parents=True)
        with open(self.numbering_file, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
    
    def set_number(self, number: int, year: Optional[int] = None):
        """Установить номер вручную
        
        Args:
            number: Номер для установки
            year: Год для нумерации. Если не указан, используется текущий год.
        """
        target_year = year if year is not None else date.today().year
        
        if self._use_db:
            try:
                with self.db_manager.get_connection() as conn:
                    cursor = conn.cursor()
                    cursor.execute(
                        "UPDATE numbering SET last_number = ? WHERE year = ?",
                        (number - 1, target_year)
                    )
                    if cursor.rowcount == 0:
                        cursor.execute(
                            "INSERT INTO numbering (year, last_number) VALUES (?, ?)",
                            (target_year, number - 1)
                        )
                    conn.commit()
            except Exception:
                self._use_db = False
                self._set_number_json(number, target_year)
        else:
            self._set_number_json(number, target_year)
    
    def _set_number_json(self, number: int, year: int):
        """Установить номер в JSON
        
        Args:
            number: Номер для установки
            year: Год для нумерации
        """
        data = {"year": year, "last": number - 1}
        self._save_json(data)
    
    def mark_number_as_used(self, number: int, year: Optional[int] = None):
        """Пометить номер как использованный (сохранить как последний использованный номер)
        
        Args:
            number: Номер документа
            year: Год для нумерации. Если не указан, используется текущий год.
        """
        target_year = year if year is not None else date.today().year
        
        if self._use_db:
            try:
                with self.db_manager.get_connection() as conn:
                    cursor = conn.cursor()
                    cursor.execute(
                        "UPDATE numbering SET last_number = ? WHERE year = ?",
                        (number, target_year)
                    )
                    if cursor.rowcount == 0:
                        cursor.execute(
                            "INSERT INTO numbering (year, last_number) VALUES (?, ?)",
                            (target_year, number)
                        )
                    conn.commit()
            except Exception:
                self._use_db = False
                self._mark_number_as_used_json(number, target_year)
        else:
            self._mark_number_as_used_json(number, target_year)
    
    def _mark_number_as_used_json(self, number: int, year: int):
        """Пометить номер как использованный в JSON
        
        Args:
            number: Номер документа
            year: Год для нумерации
        """
        data = {"year": year, "last": number}
        self._save_json(data)