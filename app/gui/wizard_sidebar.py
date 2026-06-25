"""
Wizard-сайдбар для главного окна.

Вертикальный список кнопок-шагов с иконками статуса (готово/текущий/предупреждение).
Клик по шагу переключает QStackedWidget в главном окне.
"""
from __future__ import annotations

from enum import Enum
from typing import List, Optional

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QPushButton, QLabel, QFrame, QSizePolicy,
)


class StepStatus(Enum):
    PENDING = "pending"     # серый кружок
    ACTIVE = "active"       # синий кружок
    DONE = "done"           # зелёная галка
    WARNING = "warning"     # оранжевый треугольник
    ERROR = "error"         # красный крест


_STATUS_TEXT = {
    StepStatus.PENDING: "○",
    StepStatus.ACTIVE: "●",
    StepStatus.DONE: "✓",
    StepStatus.WARNING: "⚠",
    StepStatus.ERROR: "✗",
}

_STATUS_COLOR = {
    StepStatus.PENDING: "#a0a8b0",
    StepStatus.ACTIVE: "#2a6da8",
    StepStatus.DONE: "#3a7a2a",
    StepStatus.WARNING: "#a06000",
    StepStatus.ERROR: "#a02020",
}


class WizardStepButton(QPushButton):
    """Кнопка-шаг в сайдбаре: номер + статус + название."""

    def __init__(self, index: int, title: str, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.index = index
        self._title = title
        self._status = StepStatus.PENDING

        self.setCheckable(True)
        self.setFlat(True)
        self.setCursor(Qt.PointingHandCursor)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.setMinimumHeight(48)
        self._refresh_text()
        self._refresh_style()

    def set_status(self, status: StepStatus) -> None:
        self._status = status
        self._refresh_text()
        self._refresh_style()

    def status(self) -> StepStatus:
        return self._status

    def _refresh_text(self) -> None:
        prefix = _STATUS_TEXT[self._status]
        self.setText(f"  {self.index + 1}.  {prefix}  {self._title}")

    def _refresh_style(self) -> None:
        color = _STATUS_COLOR[self._status]
        # Когда шаг выбран — жирный шрифт и активный фон
        font_weight = "600" if self._status == StepStatus.ACTIVE else "400"
        bg = "#ffffff" if self._status == StepStatus.ACTIVE else "transparent"
        border_color = "#2a6da8" if self._status == StepStatus.ACTIVE else "transparent"
        self.setStyleSheet(
            f"QPushButton {{"
            f"  text-align: left;"
            f"  padding: 8px 12px;"
            f"  border: none;"
            f"  border-left: 3px solid {border_color};"
            f"  background-color: {bg};"
            f"  color: {color};"
            f"  font-weight: {font_weight};"
            f"  font-size: 13px;"
            f"}}"
            f"QPushButton:hover {{"
            f"  background-color: #f0f4f8;"
            f"}}"
        )


class WizardSidebar(QWidget):
    """Сайдбар с шагами визарда."""

    # Сигнал: пользователь выбрал шаг (0-based индекс)
    step_selected = Signal(int)

    def __init__(
        self,
        titles: List[str],
        parent: Optional[QWidget] = None,
    ):
        super().__init__(parent)
        self._buttons: List[WizardStepButton] = []
        self._current_index: int = -1

        self.setObjectName("WizardSidebar")
        self.setFixedWidth(220)
        self.setStyleSheet(
            "#WizardSidebar {"
            "  background-color: #f8fafc;"
            "  border-right: 1px solid #d0d7de;"
            "}"
        )

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 16, 0, 16)
        layout.setSpacing(4)

        title = QLabel("Создание документа")
        title.setStyleSheet(
            "color: #666; font-size: 11px; font-weight: 600; "
            "padding: 0 12px 8px 12px; text-transform: uppercase; letter-spacing: 0.5px;"
        )
        layout.addWidget(title)

        for i, t in enumerate(titles):
            btn = WizardStepButton(i, t)
            btn.clicked.connect(lambda _checked=False, idx=i: self._on_clicked(idx))
            self._buttons.append(btn)
            layout.addWidget(btn)

        layout.addStretch(1)

        # Нижняя подсказка
        self.hint_label = QLabel("")
        self.hint_label.setStyleSheet(
            "color: #888; font-size: 10px; padding: 8px 12px;"
        )
        self.hint_label.setWordWrap(True)
        layout.addWidget(self.hint_label)

        if self._buttons:
            self.set_current(0)

    def _on_clicked(self, index: int) -> None:
        self.set_current(index)
        self.step_selected.emit(index)

    def set_current(self, index: int) -> None:
        """Подсветить выбранный шаг, остальные — pending."""
        self._current_index = index
        for i, btn in enumerate(self._buttons):
            if i == index:
                if btn.status() != StepStatus.WARNING and btn.status() != StepStatus.ERROR:
                    btn.set_status(StepStatus.ACTIVE)
                btn.setChecked(True)
            else:
                if btn.status() != StepStatus.WARNING and btn.status() != StepStatus.ERROR:
                    btn.set_status(StepStatus.PENDING)
                btn.setChecked(False)

    def set_step_status(self, index: int, status: StepStatus) -> None:
        """Установить статус шага (готово/предупреждение/ошибка)."""
        if 0 <= index < len(self._buttons):
            btn = self._buttons[index]
            btn.set_status(status)
            # Если активный шаг получает warning/error, оставляем его активным (визуальный сигнал)
            if index == self._current_index and status in (StepStatus.WARNING, StepStatus.ERROR):
                pass  # статус уже установлен

    def set_hint(self, text: str) -> None:
        self.hint_label.setText(text)

    def current_index(self) -> int:
        return self._current_index