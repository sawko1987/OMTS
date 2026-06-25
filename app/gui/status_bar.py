"""
Виджет статус-бара для главного окна.

Показывает текущие счётчики документа и состояние приложения:
- количество деталей
- количество выбранных машин (изделий)
- количество замен (материалы с галкой "Меняем")
- количество активных запретов замены в документе
- статус БД
"""
from __future__ import annotations

from typing import Optional

from PySide6.QtWidgets import (
    QWidget, QHBoxLayout, QLabel, QFrame, QSizePolicy,
)
from PySide6.QtCore import Qt

from app.models import DocumentData
from app.banned_replacements_store import BannedReplacementsStore


class StatusBarWidget(QWidget):
    """Компактная панель статуса документа и приложения."""

    def __init__(
        self,
        document_data: DocumentData,
        banned_store: Optional[BannedReplacementsStore] = None,
        parent: Optional[QWidget] = None,
    ):
        super().__init__(parent)
        self.document_data = document_data
        self.banned_store = banned_store

        layout = QHBoxLayout(self)
        layout.setContentsMargins(8, 2, 8, 2)
        layout.setSpacing(14)

        self.db_status_label = self._make_label("БД: …", "db")
        layout.addWidget(self.db_status_label)
        layout.addWidget(self._separator())

        self.parts_label = self._make_label("Деталей: 0", "parts")
        layout.addWidget(self.parts_label)

        self.machines_label = self._make_label("Машин: 0", "machines")
        layout.addWidget(self.machines_label)

        self.replacements_label = self._make_label("Замен: 0", "replacements")
        layout.addWidget(self.replacements_label)

        self.bans_label = self._make_label("Запретов: 0", "bans")
        layout.addWidget(self.bans_label)

        layout.addStretch(1)

        self.ready_label = self._make_label("", "ready")
        self.ready_label.setStyleSheet("color: #666; font-size: 11px;")
        layout.addWidget(self.ready_label)

        self.setStyleSheet(
            "StatusBarWidget { background-color: #f0f4f8; "
            "border-top: 1px solid #d0d7de; }"
        )

        self.set_db_status(True)
        self.refresh()

    @staticmethod
    def _separator() -> QFrame:
        sep = QFrame()
        sep.setFrameShape(QFrame.VLine)
        sep.setFrameShadow(QFrame.Sunken)
        sep.setStyleSheet("color: #d0d7de;")
        sep.setFixedHeight(14)
        return sep

    @staticmethod
    def _make_label(text: str, name: str) -> QLabel:
        lbl = QLabel(text)
        lbl.setObjectName(f"status_{name}")
        lbl.setStyleSheet("color: #444; font-size: 11px;")
        lbl.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Preferred)
        return lbl

    def set_db_status(self, ok: bool, detail: str = "") -> None:
        if ok:
            text = "БД: ✓"
            if detail:
                text = f"БД: ✓ ({detail})"
            self.db_status_label.setStyleSheet("color: #3a7a2a; font-size: 11px; font-weight: 600;")
        else:
            text = f"БД: ✗ {detail}".strip()
            self.db_status_label.setStyleSheet("color: #c00; font-size: 11px; font-weight: 600;")
        self.db_status_label.setText(text)

    def set_ready_text(self, text: str) -> None:
        self.ready_label.setText(text)

    def refresh(self) -> None:
        """Пересчитать все счётчики на основе текущего состояния документа."""
        data = self.document_data
        parts_count = len(data.part_changes)
        machines_count = len(data.products)
        replacements_count = sum(
            1 for pc in data.part_changes for m in pc.materials if m.is_changed
        )

        self.parts_label.setText(f"📦 Деталей: {parts_count}")
        self.machines_label.setText(f"🚜 Машин: {machines_count}")
        self.replacements_label.setText(f"🔄 Замен: {replacements_count}")

        ban_count = self._count_active_bans(data)
        self.bans_label.setText(f"⛔ Запретов: {ban_count}")

        if parts_count == 0:
            self.set_ready_text("Добавьте деталь, чтобы начать")
        elif replacements_count == 0:
            self.set_ready_text("Отметьте материалы для замены")
        elif machines_count == 0:
            self.set_ready_text("Выберите изделия (машины)")
        elif ban_count > 0:
            self.set_ready_text(f"⚠ {ban_count} активных запрет(ов) — нужно снять или убрать замену")
        else:
            self.set_ready_text("✓ Готово к генерации")

    def _count_active_bans(self, data: DocumentData) -> int:
        """Подсчитать запреты, которые сейчас блокируют генерацию."""
        if not self.banned_store:
            return 0
        count = 0
        for pc in data.part_changes:
            for m in pc.materials:
                if m.is_changed and m.after_name and m.after_name.strip():
                    if self.banned_store.is_banned(m.catalog_entry, m.after_name.strip()):
                        count += 1
        return count