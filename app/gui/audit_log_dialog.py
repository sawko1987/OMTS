"""
Диалог просмотра журнала действий.

Показывает последние N операций пользователя в текущей сессии.
"""
from __future__ import annotations

from typing import Optional

from PySide6.QtCore import Qt
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QPlainTextEdit, QDialogButtonBox,
)

from app.audit_log import AuditLog


class AuditLogDialog(QDialog):
    """Просмотр журнала действий текущей сессии."""

    def __init__(self, audit_log: AuditLog, parent: Optional[object] = None):
        super().__init__(parent)
        self.audit_log = audit_log
        self.setWindowTitle("Журнал действий")
        self.setMinimumSize(600, 400)

        layout = QVBoxLayout(self)

        header = QLabel("📋  Журнал действий текущей сессии")
        header.setStyleSheet("font-size: 14px; font-weight: 600;")
        layout.addWidget(header)

        info = QLabel(
            "Здесь показаны последние операции: добавление и удаление деталей, "
            "изменения материалов, генерация документов и т.д. "
            "Журнал хранится только в памяти и очищается при перезапуске приложения."
        )
        info.setWordWrap(True)
        info.setStyleSheet("color: #666; font-size: 11px;")
        layout.addWidget(info)

        # Текстовое поле с моноширинным шрифтом
        self.text_edit = QPlainTextEdit()
        self.text_edit.setReadOnly(True)
        mono = QFont("Consolas")
        mono.setStyleHint(QFont.Monospace)
        self.text_edit.setFont(mono)
        self.text_edit.setStyleSheet(
            "QPlainTextEdit {"
            "  background-color: #f8f8f8; border: 1px solid #d0d7de;"
            "}"
        )
        layout.addWidget(self.text_edit, 1)

        # Кнопки
        button_layout = QHBoxLayout()

        self.lbl_count = QLabel("")
        self.lbl_count.setStyleSheet("color: #666; font-size: 11px;")
        button_layout.addWidget(self.lbl_count)
        button_layout.addStretch()

        self.btn_clear = QPushButton("🧹 Очистить журнал")
        self.btn_clear.clicked.connect(self._on_clear)
        button_layout.addWidget(self.btn_clear)

        btn_refresh = QPushButton("🔄 Обновить")
        btn_refresh.clicked.connect(self.refresh)
        button_layout.addWidget(btn_refresh)

        btn_close = QPushButton("Закрыть")
        btn_close.clicked.connect(self.accept)
        button_layout.addWidget(btn_close)

        layout.addLayout(button_layout)

        self.refresh()

    def refresh(self) -> None:
        """Перечитать записи из журнала и отобразить."""
        entries = self.audit_log.entries()
        self.lbl_count.setText(f"{len(entries)} записей")

        if not entries:
            self.text_edit.setPlainText("(журнал пуст)")
            return

        lines = [e.format() for e in entries]
        self.text_edit.setPlainText("\n".join(lines))
        # Прокрутить к концу
        cursor = self.text_edit.textCursor()
        cursor.movePosition(cursor.End)
        self.text_edit.setTextCursor(cursor)

    def _on_clear(self) -> None:
        self.audit_log.clear()
        self.refresh()