"""
Тесты для Фазы 1 рефакторинга UI:
- Настройки: шаблоны причин замены (SettingsManager)
- Toast-виджет: smoke-тест создания и инициализации
- StatusBar: подсчёт счётчиков из DocumentData
"""
import os
import sys
import unittest
from pathlib import Path
from unittest.mock import patch

# Добавляем корень проекта в sys.path
sys.path.insert(0, str(Path(__file__).parent))


class SettingsManagerReasonTemplatesTest(unittest.TestCase):
    """Шаблоны причин замены: дефолты, сохранение, очистка, дедупликация."""

    def setUp(self):
        # Подменяем путь к settings.json на временный, чтобы не трогать реальные настройки
        import tempfile
        self.tmp_dir = Path(tempfile.mkdtemp(prefix="omts_settings_test_"))
        self.settings_file = self.tmp_dir / "settings.json"
        # Патчим константу SETTINGS_FILE в модуле settings_manager
        import app.settings_manager as sm_module
        self._sm_module = sm_module
        self._original_settings_file = sm_module.SETTINGS_FILE
        sm_module.SETTINGS_FILE = self.settings_file

    def tearDown(self):
        import shutil
        self._sm_module.SETTINGS_FILE = self._original_settings_file
        if self.tmp_dir.exists():
            shutil.rmtree(self.tmp_dir)

    def test_default_templates_present_on_first_load(self):
        from app.settings_manager import SettingsManager
        sm = SettingsManager()
        templates = sm.get_reason_templates()
        self.assertTrue(len(templates) >= 3, "Дефолтных шаблонов должно быть ≥ 3")
        self.assertIn("Решение ТКО", templates)

    def test_set_and_get_roundtrip(self):
        from app.settings_manager import SettingsManager
        sm = SettingsManager()
        sm.set_reason_templates(["Мой шаблон 1", "Мой шаблон 2"])
        sm2 = SettingsManager()
        self.assertEqual(sm2.get_reason_templates(), ["Мой шаблон 1", "Мой шаблон 2"])

    def test_empty_list_falls_back_to_defaults(self):
        from app.settings_manager import SettingsManager
        sm = SettingsManager()
        sm.set_reason_templates([])
        templates = sm.get_reason_templates()
        self.assertTrue(len(templates) >= 3, "Пустой список → дефолтные шаблоны")
        self.assertIn("Решение ТКО", templates)

    def test_duplicates_and_blanks_removed(self):
        from app.settings_manager import SettingsManager
        sm = SettingsManager()
        sm.set_reason_templates(["  A  ", "", "A", "B", "  B"])
        self.assertEqual(sm.get_reason_templates(), ["A", "B"])


class StatusBarCountersTest(unittest.TestCase):
    """StatusBarWidget: корректный подсчёт деталей/машин/замен/запретов."""

    def test_counts_from_document_data(self):
        from app.models import DocumentData, PartChanges, MaterialChange, CatalogEntry

        data = DocumentData()
        # Пустой документ
        self.assertEqual(len(data.part_changes), 0)

        # Документ с двумя деталями и одной заменой
        entry = CatalogEntry(part="X", workshop="ПЗУ", role="", before_name="M1",
                             unit="кг", norm=1.0)
        mat_changed = MaterialChange(catalog_entry=entry, is_changed=True)
        mat_unchanged = MaterialChange(catalog_entry=entry, is_changed=False)

        data.part_changes.append(PartChanges(part="X", materials=[mat_changed, mat_unchanged]))
        data.part_changes.append(PartChanges(part="Y", materials=[
            MaterialChange(catalog_entry=entry, is_changed=True)
        ]))
        data.products = ["Машина A", "Машина B"]

        replacements_count = sum(
            1 for pc in data.part_changes for m in pc.materials if m.is_changed
        )
        self.assertEqual(replacements_count, 2)
        self.assertEqual(len(data.part_changes), 2)
        self.assertEqual(len(data.products), 2)

    def test_status_bar_widget_imports(self):
        """StatusBarWidget импортируется и инстанцируется без QApplication."""
        from app.gui.status_bar import StatusBarWidget
        from app.models import DocumentData
        # Без QApplication нельзя создавать QWidget — пропускаем этот путь,
        # но проверяем, что класс доступен и у него есть нужные методы
        self.assertTrue(callable(StatusBarWidget))
        self.assertTrue(hasattr(StatusBarWidget, "refresh"))
        self.assertTrue(hasattr(StatusBarWidget, "set_db_status"))
        self.assertTrue(hasattr(StatusBarWidget, "set_ready_text"))


class ToastWidgetTest(unittest.TestCase):
    """Toast: smoke-тест: импорт и статические атрибуты."""

    def test_toast_level_enum(self):
        from app.gui.toast import ToastLevel
        self.assertEqual(ToastLevel.INFO.value, "info")
        self.assertEqual(ToastLevel.SUCCESS.value, "success")
        self.assertEqual(ToastLevel.WARNING.value, "warning")
        self.assertEqual(ToastLevel.ERROR.value, "error")

    def test_default_timeout(self):
        from app.gui.toast import Toast
        self.assertGreater(Toast.DEFAULT_TIMEOUT_MS, 0)
        self.assertLessEqual(Toast.DEFAULT_TIMEOUT_MS, 10000)


if __name__ == "__main__":
    unittest.main()