"""
Виджет предпросмотра документа (Шаг 4 визарда).

Показывает сводку перед генерацией:
- реквизиты (номер, дата, причина, машины)
- список деталей по страницам (1, 2+, 3+...)
- какие материалы заменены, на что
- список активных запретов с возможностью перехода к ним
- примерный размер итогового файла
"""
from __future__ import annotations

from typing import Optional, Callable

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QColor, QFont
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QGroupBox, QFrame,
    QScrollArea, QSizePolicy, QTreeWidget, QTreeWidgetItem, QHeaderView,
    QPushButton, QGridLayout,
)

from app.models import DocumentData
from app.banned_replacements_store import BannedReplacementsStore


class PreviewWidget(QWidget):
    """Шаг 4 визарда: что попадёт в итоговый Excel-файл."""

    # Сигнал: пользователь нажал "Перейти к материалам" (шаг 2)
    goto_materials = Signal()
    # Сигнал: пользователь нажал "Сгенерировать документ"
    request_generate = Signal()

    def __init__(
        self,
        document_data: DocumentData,
        banned_store: Optional[BannedReplacementsStore] = None,
        parent: Optional[QWidget] = None,
    ):
        super().__init__(parent)
        self.document_data = document_data
        self.banned_store = banned_store

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        outer.addWidget(scroll)

        container = QWidget()
        scroll.setWidget(container)
        layout = QVBoxLayout(container)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(16)

        # === 1. Шапка ===
        self.header_group = self._build_header_group()
        layout.addWidget(self.header_group)

        # === 2. Машины ===
        self.machines_group = self._build_machines_group()
        layout.addWidget(self.machines_group)

        # === 3. Детали по страницам ===
        self.parts_group = self._build_parts_group()
        layout.addWidget(self.parts_group)

        # === 4. Запреты (если есть) ===
        self.bans_group = self._build_bans_group()
        layout.addWidget(self.bans_group)

        # === 5. Действия ===
        actions = QHBoxLayout()
        actions.addStretch(1)
        self.btn_goto_materials = QPushButton("← Перейти к материалам")
        self.btn_goto_materials.clicked.connect(self.goto_materials.emit)
        actions.addWidget(self.btn_goto_materials)

        self.btn_generate = QPushButton("📄 Сгенерировать документ")
        self.btn_generate.setProperty("primary", True)
        self.btn_generate.setStyleSheet(
            "QPushButton {"
            "  background-color: #2a6da8; color: #ffffff; "
            "  border: 1px solid #1f568a; border-radius: 4px;"
            "  padding: 8px 18px; font-weight: 600;"
            "}"
            "QPushButton:hover { background-color: #3480c2; }"
            "QPushButton:pressed { background-color: #1f568a; }"
            "QPushButton:disabled { background-color: #c0c8d0; color: #666; }"
        )
        self.btn_generate.clicked.connect(self.request_generate.emit)
        actions.addWidget(self.btn_generate)
        layout.addLayout(actions)

        layout.addStretch(1)

    # ----- Группы -----

    def _build_header_group(self) -> QGroupBox:
        group = QGroupBox("📋 Реквизиты документа")
        layout = QGridLayout(group)
        layout.setColumnStretch(1, 1)

        self.lbl_number = QLabel("—")
        self.lbl_date = QLabel("—")
        self.lbl_period = QLabel("—")
        self.lbl_reason = QLabel("—")

        layout.addWidget(QLabel("Номер:"), 0, 0)
        layout.addWidget(self.lbl_number, 0, 1)
        layout.addWidget(QLabel("Дата внедрения:"), 1, 0)
        layout.addWidget(self.lbl_date, 1, 1)
        layout.addWidget(QLabel("Срок действия:"), 2, 0)
        layout.addWidget(self.lbl_period, 2, 1)
        layout.addWidget(QLabel("Причина:"), 3, 0)
        layout.addWidget(self.lbl_reason, 3, 1)

        return group

    def _build_machines_group(self) -> QGroupBox:
        group = QGroupBox("🚜 Изделия (машины)")
        layout = QVBoxLayout(group)
        self.lbl_machines = QLabel("—")
        self.lbl_machines.setWordWrap(True)
        layout.addWidget(self.lbl_machines)
        return group

    def _build_parts_group(self) -> QGroupBox:
        group = QGroupBox("📦 Детали и замены")
        v = QVBoxLayout(group)

        self.lbl_parts_summary = QLabel("—")
        self.lbl_parts_summary.setStyleSheet("color: #555; font-size: 11px;")
        v.addWidget(self.lbl_parts_summary)

        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["Деталь / Стр. / Цех", "Материал 'до'", "Материал 'после'", "Статус"])
        self.tree.setRootIsDecorated(True)
        self.tree.setAlternatingRowColors(True)
        self.tree.setUniformRowHeights(True)
        header = self.tree.header()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.Stretch)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        v.addWidget(self.tree, 1)

        return group

    def _build_bans_group(self) -> QGroupBox:
        group = QGroupBox("⛔ Активные запреты замены")
        v = QVBoxLayout(group)
        self.lbl_bans = QLabel("—")
        self.lbl_bans.setWordWrap(True)
        self.lbl_bans.setStyleSheet("color: #a02020;")
        v.addWidget(self.lbl_bans)
        group.hide()  # показывается только если есть запреты
        return group

    # ----- Публичные методы -----

    def refresh(self) -> None:
        """Пересобрать все секции из текущего состояния DocumentData."""
        data = self.document_data

        # Шапка
        self.lbl_number.setText(data.get_display_number() or "—")
        if data.implementation_date:
            self.lbl_date.setText(data.implementation_date.strftime("%d.%m.%Y"))
        else:
            self.lbl_date.setText("❌ не указана")
            self.lbl_date.setStyleSheet("color: #a02020;")
        self.lbl_period.setText(data.validity_period or "—")
        self.lbl_reason.setText(data.reason or "—")

        # Машины
        if data.products:
            self.lbl_machines.setText(", ".join(data.products))
            self.lbl_machines.setStyleSheet("")
        else:
            self.lbl_machines.setText("❌ машины не выбраны — генерация невозможна")
            self.lbl_machines.setStyleSheet("color: #a02020;")

        # Детали — дерево по страницам
        self.tree.clear()

        # Группировка по страницам: None → "1", 1 → "2+", 2 → "3+", ...
        pages: dict[int, list] = {}
        for pc in data.part_changes:
            pages.setdefault(pc.additional_page_number if pc.additional_page_number is not None else 0, []).append(pc)

        total_changes = 0
        for page_num in sorted(pages.keys()):
            page_label = f"Страница {page_num + 1}" if page_num == 0 else f"Страница {page_num + 1}+"
            page_item = QTreeWidgetItem([page_label, "", "", ""])
            page_font = QFont()
            page_font.setBold(True)
            page_item.setFont(0, page_font)
            self.tree.addTopLevelItem(page_item)

            for pc in pages[page_num]:
                for mat in pc.materials:
                    status_text = ""
                    if mat.is_changed:
                        if mat.after_name and mat.after_name.strip():
                            total_changes += 1
                            after_disp = mat.after_name.strip()
                            if self.banned_store and self.banned_store.is_banned(mat.catalog_entry, after_disp):
                                status_text = "⛔ запрет"
                            else:
                                status_text = "🔄 меняем"
                        else:
                            status_text = "⚠ нет 'после'"
                    else:
                        status_text = "— пропущено"

                    child = QTreeWidgetItem([
                        pc.part,
                        f"{mat.catalog_entry.workshop} | {mat.catalog_entry.before_name}",
                        (mat.after_name or "").strip() if mat.is_changed else "",
                        status_text,
                    ])
                    # Цвет строки
                    if "⛔" in status_text:
                        for c in range(4):
                            child.setForeground(c, QColor("#a02020"))
                    elif "⚠" in status_text:
                        for c in range(4):
                            child.setForeground(c, QColor("#a06000"))
                    elif mat.is_changed:
                        for c in range(4):
                            child.setForeground(c, QColor("#1f568a"))
                    page_item.addChild(child)

            page_item.setExpanded(True)

        parts_count = len(data.part_changes)
        self.lbl_parts_summary.setText(
            f"Всего деталей: {parts_count} | Замен: {total_changes}"
        )

        # Запреты
        if self.banned_store:
            active_bans = []
            for pc in data.part_changes:
                for mat in pc.materials:
                    if mat.is_changed and mat.after_name and mat.after_name.strip():
                        if self.banned_store.is_banned(mat.catalog_entry, mat.after_name.strip()):
                            active_bans.append(
                                f"• {pc.part}: {mat.catalog_entry.before_name} → {mat.after_name.strip()}"
                            )
            if active_bans:
                self.bans_group.show()
                self.lbl_bans.setText(
                    f"Обнаружено {len(active_bans)} запрещённых замен:\n"
                    + "\n".join(active_bans[:10])
                    + (f"\n…и ещё {len(active_bans) - 10}" if len(active_bans) > 10 else "")
                )
            else:
                self.bans_group.hide()
        else:
            self.bans_group.hide()

        # Готова ли кнопка генерации
        can_generate = bool(data.implementation_date and data.part_changes and data.products)
        self.btn_generate.setEnabled(can_generate)
        if not can_generate:
            if not data.implementation_date:
                self.btn_generate.setToolTip("Укажите дату внедрения")
            elif not data.part_changes:
                self.btn_generate.setToolTip("Добавьте хотя бы одну деталь")
            elif not data.products:
                self.btn_generate.setToolTip("Выберите хотя бы одну машину")
        else:
            self.btn_generate.setToolTip("")