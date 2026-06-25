"""
Диалог управления запретами замены материалов.

Позволяет:
- просмотреть все запреты в одном списке
- фильтровать по детали/цеху/материалу
- снимать запреты (отдельно для конкретной пары before→after или для всех пар материала)
"""
from __future__ import annotations

from typing import Optional, List

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QLineEdit, QComboBox, QTableWidget, QTableWidgetItem, QHeaderView,
    QMessageBox, QFrame,
)
from PySide6.QtGui import QColor

from app.banned_replacements_store import BannedReplacementsStore
from app.config import WORKSHOPS
from app.gui.toast import show_toast, ToastLevel


class BannedReplacementsDialog(QDialog):
    """Окно управления всеми запретами замены."""

    def __init__(self, banned_store: BannedReplacementsStore, parent=None):
        super().__init__(parent)
        self.banned_store = banned_store
        self.setWindowTitle("Запреты замены материалов")
        self.setMinimumSize(900, 550)
        self._all_rows: list[dict] = []  # кеш всех строк для фильтрации
        self._init_ui()
        self._load_all()

    def _init_ui(self) -> None:
        layout = QVBoxLayout(self)

        # Заголовок
        header = QLabel("⛔  Запрещённые замены материалов")
        header.setStyleSheet("font-size: 15px; font-weight: 600; color: #a02020;")
        layout.addWidget(header)

        info = QLabel(
            "Запрет ставит конструктор, когда замена материала недопустима по техническим "
            "или нормативным причинам. При генерации документа запрещённые замены блокируют "
            "выпуск извещения."
        )
        info.setWordWrap(True)
        info.setStyleSheet("color: #666; font-size: 11px;")
        layout.addWidget(info)

        # === Фильтры ===
        filter_layout = QHBoxLayout()

        filter_layout.addWidget(QLabel("Поиск:"))
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText(
            "Код детали, материал 'до', материал 'после' или причина…"
        )
        self.search_edit.textChanged.connect(self._apply_filters)
        filter_layout.addWidget(self.search_edit, 1)

        filter_layout.addWidget(QLabel("Цех:"))
        self.workshop_filter = QComboBox()
        self.workshop_filter.addItem("Все цеха", "")
        self.workshop_filter.addItems(WORKSHOPS)
        self.workshop_filter.currentIndexChanged.connect(self._apply_filters)
        filter_layout.addWidget(self.workshop_filter)

        layout.addLayout(filter_layout)

        # === Таблица ===
        self.table = QTableWidget()
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels([
            "Деталь", "Цех", "Материал 'до'", "Запрещённая замена 'после'",
            "Причина", "Дата", "",
        ])
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)

        header_view = self.table.horizontalHeader()
        header_view.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header_view.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header_view.setSectionResizeMode(2, QHeaderView.Stretch)
        header_view.setSectionResizeMode(3, QHeaderView.Stretch)
        header_view.setSectionResizeMode(4, QHeaderView.Stretch)
        header_view.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        header_view.setSectionResizeMode(6, QHeaderView.ResizeToContents)

        layout.addWidget(self.table, 1)

        # === Подвал: счётчик + кнопки ===
        footer = QHBoxLayout()

        self.lbl_count = QLabel("0 запретов")
        self.lbl_count.setStyleSheet("color: #666; font-size: 11px;")
        footer.addWidget(self.lbl_count)
        footer.addStretch()

        self.btn_unban = QPushButton("🗑 Снять запрет (конкретная пара)")
        self.btn_unban.setToolTip(
            "Снять запрет только для выбранной пары «до → после». "
            "Если для этого материала есть другие запреты — они сохранятся."
        )
        self.btn_unban.clicked.connect(lambda: self._unban(specific=True))
        footer.addWidget(self.btn_unban)

        self.btn_unban_all = QPushButton("🗑 Снять ВСЕ запреты для этого материала")
        self.btn_unban_all.setToolTip(
            "Снять все запреты для выбранного материала 'до', "
            "включая пары с другими значениями 'после'."
        )
        self.btn_unban_all.clicked.connect(lambda: self._unban(specific=False))
        footer.addWidget(self.btn_unban_all)

        btn_close = QPushButton("Закрыть")
        btn_close.clicked.connect(self.accept)
        footer.addWidget(btn_close)

        layout.addLayout(footer)

        # Начальное состояние кнопок
        self._update_buttons_state()

    # === Загрузка и отображение ===

    def _load_all(self) -> None:
        """Загрузить все запреты из БД."""
        try:
            rows = self.banned_store.get_all_banned()
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось загрузить запреты:\n{e}")
            self._all_rows = []
            self._render([])
            return

        # rows — список sqlite3.Row, конвертируем в dict
        self._all_rows = [
            {
                "part_code": r["part_code"],
                "workshop": r["workshop"] or "",
                "role": r["role"] or "",
                "before_name": r["before_name"],
                "after_name": r["after_name"] or "",
                "reason": r["reason"] or "",
                "banned_at": r["banned_at"] or "",
            }
            for r in rows
        ]
        self._apply_filters()

    def _apply_filters(self) -> None:
        """Применить поисковый фильтр и фильтр по цеху, отрисовать таблицу."""
        search = self.search_edit.text().strip().casefold()
        workshop = self.workshop_filter.currentData() or ""

        def matches(r: dict) -> bool:
            if workshop and r["workshop"] != workshop:
                return False
            if not search:
                return True
            haystack = " ".join((
                r["part_code"], r["workshop"], r["before_name"],
                r["after_name"], r["reason"],
            )).casefold()
            return search in haystack

        filtered = [r for r in self._all_rows if matches(r)]
        self._render(filtered)

    def _render(self, rows: list[dict]) -> None:
        """Заполнить таблицу строками."""
        self.table.setRowCount(len(rows))
        for row_idx, r in enumerate(rows):
            self.table.setItem(row_idx, 0, QTableWidgetItem(r["part_code"]))
            self.table.setItem(row_idx, 1, QTableWidgetItem(r["workshop"]))
            self.table.setItem(row_idx, 2, QTableWidgetItem(r["before_name"]))
            self.table.setItem(row_idx, 3, QTableWidgetItem(r["after_name"] or "(любая замена)"))
            self.table.setItem(row_idx, 4, QTableWidgetItem(r["reason"]))
            self.table.setItem(row_idx, 5, QTableWidgetItem(r["banned_at"]))

            # Кнопка "Снять" в последней колонке
            btn = QPushButton("✕")
            btn.setToolTip("Снять запрет для этой пары")
            btn.setFixedWidth(32)
            btn.setStyleSheet(
                "QPushButton { color: #a02020; border: 1px solid #f0a0a0; "
                "border-radius: 4px; padding: 2px 4px; }"
                "QPushButton:hover { background-color: #fde0e0; }"
            )
            # Замыкание на текущий row через default-аргумент
            btn.clicked.connect(lambda _checked=False, rr=r: self._unban_row(rr))
            self.table.setCellWidget(row_idx, 6, btn)

        self.lbl_count.setText(
            f"{len(rows)} запрет(ов)" + (f" из {len(self._all_rows)}" if len(rows) != len(self._all_rows) else "")
        )
        self._update_buttons_state()

        # Подсветка запрещённых строк светло-красным
        for row_idx in range(self.table.rowCount()):
            for col in range(7):
                item = self.table.item(row_idx, col)
                if item:
                    item.setBackground(QColor(255, 240, 240))

    def _update_buttons_state(self) -> None:
        """Включить/выключить кнопки в зависимости от выбора."""
        has_selection = self.table.currentRow() >= 0
        self.btn_unban.setEnabled(has_selection)
        self.btn_unban_all.setEnabled(has_selection)

    # === Действия ===

    def _get_selected_row_dict(self) -> Optional[dict]:
        """Получить данные выбранной строки."""
        row_idx = self.table.currentRow()
        if row_idx < 0:
            return None
        if row_idx >= len(self._visible_rows):
            return None
        return self._visible_rows[row_idx]

    def _unban(self, specific: bool) -> None:
        """Снять запрет для выбранной строки."""
        # Сохраняем видимые строки чтобы маппить индекс
        # (мы рендерим filtered, но self._all_rows содержит оригиналы)
        row_idx = self.table.currentRow()
        if row_idx < 0:
            QMessageBox.information(self, "Информация", "Выберите строку для снятия запрета")
            return

        # Достаём данные из ячеек напрямую
        part = self.table.item(row_idx, 0).text() if self.table.item(row_idx, 0) else ""
        workshop = self.table.item(row_idx, 1).text() if self.table.item(row_idx, 1) else ""
        before_name = self.table.item(row_idx, 2).text() if self.table.item(row_idx, 2) else ""
        after_name = self.table.item(row_idx, 3).text() if self.table.item(row_idx, 3) else ""
        if after_name == "(любая замена)":
            after_name = ""

        # Конструируем CatalogEntry (без полного norm/unit, не нужны для операции)
        from app.models import CatalogEntry
        entry = CatalogEntry(
            part=part, workshop=workshop, role="",
            before_name=before_name, unit="", norm=0.0,
        )

        if specific and after_name:
            msg = f"Снять запрет для пары:\n{before_name} → {after_name}\n(деталь: {part})?"
        else:
            msg = (
                f"Снять ВСЕ запреты для материала «{before_name}» "
                f"(деталь: {part}, цех: {workshop})?\n\n"
                "Это удалит запреты и для других пар 'после'."
            )

        reply = QMessageBox.question(
            self, "Подтверждение", msg,
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            return

        ok = self.banned_store.unban(entry, after_name if specific else None)
        if ok:
            show_toast(
                self.window() if self.parent() is None else self.parent(),
                "Запрет снят",
                ToastLevel.SUCCESS,
            )
            self._load_all()
        else:
            QMessageBox.warning(self, "Ошибка", "Не удалось снять запрет")

    def _unban_row(self, row_data: dict) -> None:
        """Снять запрет для строки, переданной из inline-кнопки."""
        from app.models import CatalogEntry
        entry = CatalogEntry(
            part=row_data["part_code"],
            workshop=row_data["workshop"],
            role=row_data["role"],
            before_name=row_data["before_name"],
            unit="", norm=0.0,
        )
        after_name = row_data["after_name"] or None

        ok = self.banned_store.unban(entry, after_name)
        if ok:
            show_toast(
                self.window() if self.parent() is None else self.parent(),
                "Запрет снят",
                ToastLevel.SUCCESS,
            )
            self._load_all()
        else:
            QMessageBox.warning(self, "Ошибка", "Не удалось снять запрет")

    # Кеш видимых строк для маппинга индексов (на случай если потом понадобится)
    @property
    def _visible_rows(self) -> list[dict]:
        """Получить список строк, видимых в таблице в текущем порядке."""
        rows = []
        for row_idx in range(self.table.rowCount()):
            part = self.table.item(row_idx, 0).text() if self.table.item(row_idx, 0) else ""
            workshop = self.table.item(row_idx, 1).text() if self.table.item(row_idx, 1) else ""
            before_name = self.table.item(row_idx, 2).text() if self.table.item(row_idx, 2) else ""
            after_name = self.table.item(row_idx, 3).text() if self.table.item(row_idx, 3) else ""
            if after_name == "(любая замена)":
                after_name = ""
            rows.append({
                "part_code": part,
                "workshop": workshop,
                "before_name": before_name,
                "after_name": after_name,
                "role": "",
                "reason": "",
                "banned_at": "",
            })
        return rows