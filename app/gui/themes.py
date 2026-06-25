"""
Управление темами приложения: light / dark / system.

Темы применяются через QApplication.setStyleSheet().
- "light" — стиль style.qss (по умолчанию)
- "dark" — стиль style_dark.qss
- "system" — пытается определить тёмную тему Windows через палитру Qt
"""
from __future__ import annotations

import logging
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QGuiApplication, QPalette
from PySide6.QtWidgets import QApplication

logger = logging.getLogger(__name__)


THEME_LIGHT = "light"
THEME_DARK = "dark"
THEME_SYSTEM = "system"
VALID_THEMES = (THEME_LIGHT, THEME_DARK, THEME_SYSTEM)

# Папка, где лежат QSS-файлы (рядом с этим модулем)
_QSS_DIR = Path(__file__).parent
_LIGHT_QSS = _QSS_DIR / "style.qss"
_DARK_QSS = _QSS_DIR / "style_dark.qss"


def _detect_system_dark() -> bool:
    """Определить, использует ли ОС тёмную тему.

    Используем Qt-палитру: если Window текст светлее фона, система в тёмной теме.
    Также проверяем Qt.ColorScheme если доступен (Qt 6.5+).
    """
    app = QApplication.instance()
    if app is None:
        return False

    # Qt 6.5+: явный API
    try:
        scheme = app.styleHints().colorScheme()
        if hasattr(scheme, "value"):
            # Qt.ColorScheme.Dark == 1, Unknown == 0, Light == 2
            if scheme.value == 1:
                return True
            if scheme.value == 2:
                return False
    except (AttributeError, Exception):
        pass

    # Fallback: палитра
    palette = app.palette()
    bg = palette.color(QPalette.Window)
    text_color = palette.color(QPalette.WindowText)
    # Простая эвристика: если фон тёмный (luminance < 0.5) — тёмная тема
    luminance = 0.299 * bg.red() + 0.587 * bg.green() + 0.114 * bg.blue()
    return luminance < 128


def _load_qss(theme: str) -> str:
    """Загрузить QSS для указанной темы (с учётом 'system')."""
    effective = theme
    if theme == THEME_SYSTEM:
        effective = THEME_DARK if _detect_system_dark() else THEME_LIGHT

    qss_path = _DARK_QSS if effective == THEME_DARK else _LIGHT_QSS
    try:
        with open(qss_path, "r", encoding="utf-8") as f:
            content = f.read()
        logger.debug(f"Загружен QSS для темы {effective} из {qss_path}")
        return content
    except OSError as e:
        logger.warning(f"Не удалось загрузить QSS {qss_path}: {e}")
        return ""


def apply_theme(theme: str) -> str:
    """Применить тему к QApplication.

    Args:
        theme: одна из THEME_LIGHT, THEME_DARK, THEME_SYSTEM

    Returns:
        Эффективная тема (light/dark), применённая в результате
    """
    if theme not in VALID_THEMES:
        logger.warning(f"Неизвестная тема '{theme}', используется light")
        theme = THEME_LIGHT

    qss = _load_qss(theme)
    app = QApplication.instance()
    if app is not None:
        app.setStyleSheet(qss)
    effective = theme if theme != THEME_SYSTEM else (
        THEME_DARK if _detect_system_dark() else THEME_LIGHT
    )
    return effective


def get_effective_theme(theme: str) -> str:
    """Получить эффективную тему (раскрыть 'system' в light/dark)."""
    if theme == THEME_SYSTEM:
        return THEME_DARK if _detect_system_dark() else THEME_LIGHT
    return theme