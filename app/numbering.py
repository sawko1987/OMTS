"""
Управление нумерацией документов
"""
import json
from pathlib import Path
from datetime import date
from typing import Optional

from app.config import NUMBERING_FILE, DATABASE_PATH
from app.database import DatabaseManager
from app.settings_manager import SettingsManager


class NumberingManager:
    """Менеджер нумерации документов (работает с БД, с fallback на JSON)"""
    
    def __init__(self, db_manager: DatabaseManager = None, numbering_file: Path = None):
        self.db_manager = db_manager or DatabaseManager()
        self.numbering_file = numbering_file or NUMBERING_FILE
        self._use_db = DATABASE_PATH.exists()
        self.settings_manager = SettingsManager()
    
    def get_next_number(self) -> int:
        """Получить следующий номер документа"""
        current_year = date.today().year
        
        if self._use_db:
            try:
                with self.db_manager.get_connection() as conn:
                    cursor = conn.cursor()
                    
                    # Проверяем, есть ли запись для текущего года
                    cursor.execute("SELECT last_number FROM numbering WHERE year = ?", (current_year,))
                    row = cursor.fetchone()
                    
                    if row:
                        last_number = row['last_number']
                        new_number = last_number + 1
                        cursor.execute(
                            "UPDATE numbering SET last_number = ? WHERE year = ?",
                            (new_number, current_year)
                        )
                    else:
                        # Создаём запись для нового года с начальным номером из настроек
                        starting_number = self.settings_manager.get_starting_number()
                        new_number = starting_number
                        cursor.execute(
                            "INSERT INTO numbering (year, last_number) VALUES (?, ?)",
                            (current_year, new_number)
                        )
                    
                    conn.commit()
                    return new_number
            except Exception:
                # Fallback на JSON
                self._use_db = False
                return self._get_next_number_json()
        else:
            return self._get_next_number_json()
    
    def get_current_number(self) -> int:
        """Получить текущий номер (без увеличения)"""
        current_year = date.today().year
        
        if self._use_db:
            try:
                with self.db_manager.get_connection() as conn:
                    cursor = conn.cursor()
                    cursor.execute("SELECT last_number FROM numbering WHERE year = ?", (current_year,))
                    row = cursor.fetchone()
                    if row:
                        return row['last_number'] + 1
                    # Если записи нет, возвращаем начальный номер из настроек
                    return self.settings_manager.get_starting_number()
            except Exception:
                self._use_db = False
                return self._get_current_number_json()
        else:
            return self._get_current_number_json()
    
    def _get_next_number_json(self) -> int:
        """Получить следующий номер из JSON (fallback)"""
        current_year = date.today().year
        starting_number = self.settings_manager.get_starting_number()
        data = {"year": current_year, "last": starting_number - 1}
        
        if self.numbering_file.exists():
            try:
                with open(self.numbering_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except (json.JSONDecodeError, FileNotFoundError):
                pass
        
        # Если год изменился, сбрасываем счётчик на начальный номер из настроек
        if data.get("year") != current_year:
            data["year"] = current_year
            data["last"] = starting_number - 1
        
        # Увеличиваем номер
        data["last"] += 1
        self._save_json(data)
        
        return data["last"]
    
    def _get_current_number_json(self) -> int:
        """Получить текущий номер из JSON (fallback)"""
        current_year = date.today().year
        starting_number = self.settings_manager.get_starting_number()
        data = {"year": current_year, "last": starting_number - 1}
        
        if self.numbering_file.exists():
            try:
                with open(self.numbering_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except (json.JSONDecodeError, FileNotFoundError):
                pass
        
        if data.get("year") != current_year:
            return starting_number
        
        return data.get("last", starting_number - 1) + 1
    
    def _save_json(self, data: dict):
        """Сохранить данные в JSON"""
        self.numbering_file.parent.mkdir(exist_ok=True, parents=True)
        with open(self.numbering_file, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
    
    def set_number(self, number: int):
        """Установить номер вручную"""
        current_year = date.today().year
        
        if self._use_db:
            try:
                with self.db_manager.get_connection() as conn:
                    cursor = conn.cursor()
                    cursor.execute(
                        "UPDATE numbering SET last_number = ? WHERE year = ?",
                        (number - 1, current_year)
                    )
                    if cursor.rowcount == 0:
                        cursor.execute(
                            "INSERT INTO numbering (year, last_number) VALUES (?, ?)",
                            (current_year, number - 1)
                        )
                    conn.commit()
            except Exception:
                self._use_db = False
                self._set_number_json(number)
        else:
            self._set_number_json(number)
    
    def _set_number_json(self, number: int):
        """Установить номер в JSON"""
        current_year = date.today().year
        data = {"year": current_year, "last": number - 1}
        self._save_json(data)
    
    def mark_number_as_used(self, number: int):
        """Пометить номер как использованный (сохранить как последний использованный номер)"""
        current_year = date.today().year
        
        if self._use_db:
            try:
                with self.db_manager.get_connection() as conn:
                    cursor = conn.cursor()
                    cursor.execute(
                        "UPDATE numbering SET last_number = ? WHERE year = ?",
                        (number, current_year)
                    )
                    if cursor.rowcount == 0:
                        cursor.execute(
                            "INSERT INTO numbering (year, last_number) VALUES (?, ?)",
                            (current_year, number)
                        )
                    conn.commit()
            except Exception:
                self._use_db = False
                self._mark_number_as_used_json(number)
        else:
            self._mark_number_as_used_json(number)
    
    def _mark_number_as_used_json(self, number: int):
        """Пометить номер как использованный в JSON"""
        current_year = date.today().year
        data = {"year": current_year, "last": number}
        self._save_json(data)