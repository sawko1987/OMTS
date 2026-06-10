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
    
    @staticmethod
    def format_number(month: int, year: int, seq_number: int) -> str:
        """Отформатировать номер документа в формате ММ-ГГ-№
        
        Args:
            month: Месяц (1-12)
            year: Полный год (напр. 2026)
            seq_number: Последовательный номер
        
        Returns:
            str: Отформатированный номер, напр. "01-26-5"
        """
        mm = f"{month:02d}"
        yy = f"{year % 100:02d}"
        return f"{mm}-{yy}-{seq_number}"
    
    def _make_key(self, year: int, month: int) -> str:
        """Создать ключ для JSON: 'год_месяц'"""
        return f"{year}_{month}"
    
    def get_next_number(self, year: Optional[int] = None, month: Optional[int] = None) -> int:
        """Получить следующий номер документа
        
        Args:
            year: Год для нумерации. Если не указан, используется текущий год.
            month: Месяц для нумерации (1-12). Если не указан, используется текущий месяц.
        """
        target_year = year if year is not None else date.today().year
        target_month = month if month is not None else date.today().month
        
        if self._use_db:
            try:
                with self.db_manager.get_connection() as conn:
                    cursor = conn.cursor()
                    
                    # Проверяем, есть ли запись для указанного года и месяца
                    cursor.execute(
                        "SELECT last_number FROM numbering WHERE year = ? AND month = ?",
                        (target_year, target_month)
                    )
                    row = cursor.fetchone()
                    
                    if row:
                        last_number = row['last_number']
                        new_number = last_number + 1
                        cursor.execute(
                            "UPDATE numbering SET last_number = ? WHERE year = ? AND month = ?",
                            (new_number, target_year, target_month)
                        )
                    else:
                        # Создаём запись для нового месяца с начальным номером из настроек
                        starting_number = self.settings_manager.get_starting_number()
                        new_number = starting_number
                        cursor.execute(
                            "INSERT INTO numbering (year, month, last_number) VALUES (?, ?, ?)",
                            (target_year, target_month, new_number)
                        )
                    
                    conn.commit()
                    return new_number
            except Exception:
                # Fallback на JSON
                self._use_db = False
                return self._get_next_number_json(target_year, target_month)
        else:
            return self._get_next_number_json(target_year, target_month)
    
    def get_current_number(self, year: Optional[int] = None, month: Optional[int] = None) -> int:
        """Получить текущий номер (без увеличения)
        
        Возвращает следующий номер из счётчика numbering.
        
        Args:
            year: Год для нумерации. Если не указан, используется текущий год.
            month: Месяц для нумерации (1-12). Если не указан, используется текущий месяц.
        """
        target_year = year if year is not None else date.today().year
        target_month = month if month is not None else date.today().month
        logger.info(f"[get_current_number] Запрос номера для года: {target_year}, месяца: {target_month}")
        
        if self._use_db:
            try:
                with self.db_manager.get_connection() as conn:
                    cursor = conn.cursor()
                    
                    # Получаем номер из таблицы numbering
                    cursor.execute(
                        "SELECT last_number FROM numbering WHERE year = ? AND month = ?",
                        (target_year, target_month)
                    )
                    numbering_row = cursor.fetchone()
                    numbering_number = numbering_row['last_number'] + 1 if numbering_row else None
                    
                    if numbering_number is not None:
                        result = numbering_number
                        logger.info(f"[get_current_number] Для года {target_year}, месяца {target_month}: numbering={numbering_number}, возвращаем {result}")
                    else:
                        # Если записей нет, возвращаем начальный номер из настроек
                        starting_number = self.settings_manager.get_starting_number()
                        logger.info(f"[get_current_number] Запись для года {target_year}, месяца {target_month} не найдена, возвращаем начальный номер: {starting_number}")
                        return starting_number
                    
                    return result
            except Exception as e:
                logger.warning(f"[get_current_number] Ошибка при работе с БД: {e}, переключаемся на JSON")
                self._use_db = False
                return self._get_current_number_json(target_year, target_month)
        else:
            logger.info(f"[get_current_number] Используется JSON fallback для года {target_year}, месяца {target_month}")
            return self._get_current_number_json(target_year, target_month)
    
    def _get_next_number_json(self, year: int, month: int) -> int:
        """Получить следующий номер из JSON (fallback)
        
        Args:
            year: Год для нумерации
            month: Месяц для нумерации (1-12)
        """
        starting_number = self.settings_manager.get_starting_number()
        key = self._make_key(year, month)
        data = {}
        
        if self.numbering_file.exists():
            try:
                with open(self.numbering_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except (json.JSONDecodeError, FileNotFoundError):
                pass
        
        last = data.get(key, starting_number - 1)
        new_number = last + 1
        data[key] = new_number
        self._save_json(data)
        
        return new_number
    
    def _get_current_number_json(self, year: int, month: int) -> int:
        """Получить текущий номер из JSON (fallback)
        
        Args:
            year: Год для нумерации
            month: Месяц для нумерации (1-12)
        """
        starting_number = self.settings_manager.get_starting_number()
        key = self._make_key(year, month)
        data = {}
        
        if self.numbering_file.exists():
            try:
                with open(self.numbering_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except (json.JSONDecodeError, FileNotFoundError):
                pass
        
        last = data.get(key, starting_number - 1)
        return last + 1
    
    def _save_json(self, data: dict):
        """Сохранить данные в JSON"""
        self.numbering_file.parent.mkdir(exist_ok=True, parents=True)
        with open(self.numbering_file, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
    
    def set_number(self, number: int, year: Optional[int] = None, month: Optional[int] = None):
        """Установить номер вручную
        
        Args:
            number: Номер для установки
            year: Год для нумерации. Если не указан, используется текущий год.
            month: Месяц для нумерации (1-12). Если не указан, используется текущий месяц.
        """
        target_year = year if year is not None else date.today().year
        target_month = month if month is not None else date.today().month
        
        if self._use_db:
            try:
                with self.db_manager.get_connection() as conn:
                    cursor = conn.cursor()
                    cursor.execute(
                        "UPDATE numbering SET last_number = ? WHERE year = ? AND month = ?",
                        (number - 1, target_year, target_month)
                    )
                    if cursor.rowcount == 0:
                        cursor.execute(
                            "INSERT INTO numbering (year, month, last_number) VALUES (?, ?, ?)",
                            (target_year, target_month, number - 1)
                        )
                    conn.commit()
            except Exception:
                self._use_db = False
                self._set_number_json(number, target_year, target_month)
        else:
            self._set_number_json(number, target_year, target_month)
    
    def _set_number_json(self, number: int, year: int, month: int):
        """Установить номер в JSON
        
        Args:
            number: Номер для установки
            year: Год для нумерации
            month: Месяц для нумерации (1-12)
        """
        key = self._make_key(year, month)
        data = {}
        if self.numbering_file.exists():
            try:
                with open(self.numbering_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except (json.JSONDecodeError, FileNotFoundError):
                pass
        data[key] = number - 1
        self._save_json(data)
    
    def mark_number_as_used(self, number: int, year: Optional[int] = None, month: Optional[int] = None):
        """Пометить номер как использованный (сохранить как последний использованный номер)
        
        Args:
            number: Номер документа
            year: Год для нумерации. Если не указан, используется текущий год.
            month: Месяц для нумерации (1-12). Если не указан, используется текущий месяц.
        """
        target_year = year if year is not None else date.today().year
        target_month = month if month is not None else date.today().month
        
        if self._use_db:
            try:
                with self.db_manager.get_connection() as conn:
                    cursor = conn.cursor()
                    cursor.execute(
                        "SELECT last_number FROM numbering WHERE year = ? AND month = ?",
                        (target_year, target_month)
                    )
                    row = cursor.fetchone()
                    last_number = row['last_number'] if row else 0
                    number_to_store = max(last_number, number)

                    cursor.execute(
                        "UPDATE numbering SET last_number = ? WHERE year = ? AND month = ?",
                        (number_to_store, target_year, target_month)
                    )
                    if cursor.rowcount == 0:
                        cursor.execute(
                            "INSERT INTO numbering (year, month, last_number) VALUES (?, ?, ?)",
                            (target_year, target_month, number_to_store)
                        )
                    conn.commit()
            except Exception:
                self._use_db = False
                self._mark_number_as_used_json(number, target_year, target_month)
        else:
            self._mark_number_as_used_json(number, target_year, target_month)
    
    def _mark_number_as_used_json(self, number: int, year: int, month: int):
        """Пометить номер как использованный в JSON
        
        Args:
            number: Номер документа
            year: Год для нумерации
            month: Месяц для нумерации (1-12)
        """
        current_number = self._get_current_number_json(year, month)
        key = self._make_key(year, month)
        data = {}
        if self.numbering_file.exists():
            try:
                with open(self.numbering_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except (json.JSONDecodeError, FileNotFoundError):
                pass
        data[key] = max(current_number - 1, number)
        self._save_json(data)
