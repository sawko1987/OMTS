"""
Журнал действий пользователя (audit log).

Хранит последние N событий в памяти и (опционально) в файле.
Используется для отображения истории сеанса в диалоге «Журнал действий».

Сейчас — только in-memory. При желании легко добавить file-based persistence.
"""
from __future__ import annotations

import logging
from collections import deque
from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional, Deque, List

logger = logging.getLogger(__name__)

MAX_ENTRIES_DEFAULT = 200


@dataclass
class LogEntry:
    """Одна запись журнала."""
    timestamp: datetime
    action: str  # краткое имя действия, например "add_part"
    details: str = ""  # детали, например "Деталь КИР 03.614 (3 материала)"
    level: str = "info"  # info | warning | error

    def format(self) -> str:
        ts = self.timestamp.strftime("%H:%M:%S")
        prefix = {
            "info": "ℹ",
            "warning": "⚠",
            "error": "✗",
            "success": "✓",
        }.get(self.level, "·")
        return f"[{ts}] {prefix} {self.action}: {self.details}"


class AuditLog:
    """Простой журнал действий с кольцевым буфером."""

    def __init__(self, max_entries: int = MAX_ENTRIES_DEFAULT):
        self._entries: Deque[LogEntry] = deque(maxlen=max_entries)

    def info(self, action: str, details: str = "") -> None:
        self._add(LogEntry(datetime.now(), action, details, "info"))

    def success(self, action: str, details: str = "") -> None:
        self._add(LogEntry(datetime.now(), action, details, "success"))

    def warning(self, action: str, details: str = "") -> None:
        self._add(LogEntry(datetime.now(), action, details, "warning"))

    def error(self, action: str, details: str = "") -> None:
        self._add(LogEntry(datetime.now(), action, details, "error"))

    def _add(self, entry: LogEntry) -> None:
        self._entries.append(entry)
        # Дублируем в logging для консольного вывода
        msg = entry.format()
        if entry.level == "error":
            logger.error(msg)
        elif entry.level == "warning":
            logger.warning(msg)
        else:
            logger.info(msg)

    def entries(self) -> List[LogEntry]:
        """Получить копию списка записей (от старых к новым)."""
        return list(self._entries)

    def last(self) -> Optional[LogEntry]:
        """Получить последнюю запись или None."""
        return self._entries[-1] if self._entries else None

    def clear(self) -> None:
        self._entries.clear()