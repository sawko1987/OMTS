"""
Toast-виджет — компактное ненавязчивое уведомление.

Используется вместо QMessageBox для информационных сообщений:
не блокирует UI, не требует подтверждения, исчезает через N секунд.
Для ошибок и предупреждений, требующих внимания пользователя,
по-прежнему используется QMessageBox.
"""
from __future__ import annotations

from enum import Enum
from typing import Optional

from PySide6.QtCore import Qt, QTimer, QPropertyAnimation, QEasingCurve, QPoint, QAbstractAnimation
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QWidget, QLabel, QHBoxLayout, QPushButton, QGraphicsOpacityEffect,
    QSizePolicy,
)


class ToastLevel(Enum):
    INFO = "info"
    SUCCESS = "success"
    WARNING = "warning"
    ERROR = "error"


_LEVEL_STYLE = {
    ToastLevel.INFO: ("ℹ", "#2a6da8", "#e8f0fa", "#b8d0e8"),
    ToastLevel.SUCCESS: ("✓", "#3a7a2a", "#e8f5e0", "#c0e0a0"),
    ToastLevel.WARNING: ("⚠", "#a06000", "#fff5e0", "#f0c878"),
    ToastLevel.ERROR: ("✗", "#a02020", "#fde0e0", "#f0a0a0"),
}


class Toast(QWidget):
    """Одиночный тост. Автоматически скрывается через `timeout_ms`."""

    DEFAULT_TIMEOUT_MS = 3500

    def __init__(
        self,
        text: str,
        level: ToastLevel = ToastLevel.INFO,
        parent: Optional[QWidget] = None,
        timeout_ms: int = DEFAULT_TIMEOUT_MS,
    ):
        super().__init__(parent)
        self.setWindowFlags(Qt.ToolTip | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        self.setAttribute(Qt.WA_ShowWithoutActivating, True)

        # Контейнер с фоном и рамкой (помещаем во внутренний QWidget, чтобы
        # прозрачность работала на самом окне, а фон был у потомка).
        container = QWidget(self)
        container.setObjectName("ToastContainer")

        icon_char, fg, bg, border = _LEVEL_STYLE[level]
        container.setStyleSheet(
            f"#ToastContainer {{"
            f"background-color: {bg};"
            f"border: 1px solid {border};"
            f"border-left: 4px solid {fg};"
            f"border-radius: 6px;"
            f"}}"
            f"#ToastContainer QLabel {{"
            f"color: {fg};"
            f"background: transparent;"
            f"}}"
            f"#ToastContainer QPushButton {{"
            f"color: {fg};"
            f"background: transparent;"
            f"border: none;"
            f"font-weight: bold;"
            f"padding: 0 4px;"
            f"}}"
        )

        layout = QHBoxLayout(container)
        layout.setContentsMargins(10, 8, 6, 8)
        layout.setSpacing(8)

        icon_label = QLabel(icon_char)
        icon_label.setStyleSheet(f"color: {fg}; font-size: 16px; font-weight: bold;")
        layout.addWidget(icon_label)

        text_label = QLabel(text)
        text_label.setWordWrap(True)
        text_label.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Preferred)
        text_label.setMinimumWidth(200)
        text_label.setMaximumWidth(420)
        layout.addWidget(text_label, 1)

        close_btn = QPushButton("✕")
        close_btn.setFixedSize(20, 20)
        close_btn.setCursor(Qt.PointingHandCursor)
        close_btn.clicked.connect(self._hide_and_close)
        layout.addWidget(close_btn, 0, Qt.AlignTop)

        outer = QHBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(container)

        # Эффект прозрачности для плавного появления/исчезновения
        self._opacity_effect = QGraphicsOpacityEffect(self)
        self._opacity_effect.setOpacity(0.0)
        self.setGraphicsEffect(self._opacity_effect)

        self._fade_in = QPropertyAnimation(self._opacity_effect, b"opacity", self)
        self._fade_in.setDuration(180)
        self._fade_in.setStartValue(0.0)
        self._fade_in.setEndValue(1.0)
        self._fade_in.setEasingCurve(QEasingCurve.OutCubic)

        self._fade_out = QPropertyAnimation(self._opacity_effect, b"opacity", self)
        self._fade_out.setDuration(220)
        self._fade_out.setStartValue(1.0)
        self._fade_out.setEndValue(0.0)
        self._fade_out.setEasingCurve(QEasingCurve.InCubic)
        self._fade_out.finished.connect(self.deleteLater)

        # Таймер авто-скрытия
        self._timer = QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.timeout.connect(self._hide_and_close)
        self._timer.start(timeout_ms)

        self.adjustSize()

    def show_at(self, global_pos: QPoint) -> None:
        """Показать тост с якорем в правом-нижнем углу родителя (глобальная координата)."""
        self.move(global_pos)
        self.show()
        self._fade_in.start()

    def _hide_and_close(self) -> None:
        if self._timer.isActive():
            self._timer.stop()
        if self._fade_out.state() != QAbstractAnimation.Running:
            self._fade_out.start()


def show_toast(
    parent: QWidget,
    text: str,
    level: ToastLevel = ToastLevel.INFO,
    timeout_ms: int = Toast.DEFAULT_TIMEOUT_MS,
) -> Toast:
    """Утилита: показать тост, привязанный к правому-нижнему углу parent.

    Возвращает созданный тост — можно подключить к finished(), если нужно.
    """
    toast = Toast(text, level=level, parent=parent, timeout_ms=timeout_ms)
    # Размещаем в правом-нижнем углу parent с отступом 16px
    parent_rect = parent.rect()
    bottom_right = parent.mapToGlobal(
        parent_rect.bottomRight() - QPoint(toast.width() + 16, toast.height() + 16)
    )
    toast.show_at(bottom_right)
    return toast