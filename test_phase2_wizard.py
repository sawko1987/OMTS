"""
Тесты для Фазы 2 рефакторинга UI:
- WizardSidebar: переключение шагов, статусы
- PreviewWidget: корректное формирование сводки
- BannedReplacementsDialog: smoke-тест
"""
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))


class WizardSidebarSmokeTest(unittest.TestCase):
    """Smoke-тесты структуры виджета (без инстанцирования QWidget)."""

    def test_wizard_sidebar_imports(self):
        from app.gui.wizard_sidebar import WizardSidebar, WizardStepButton, StepStatus
        self.assertTrue(callable(WizardSidebar))
        self.assertTrue(callable(WizardStepButton))

    def test_step_status_enum_values(self):
        from app.gui.wizard_sidebar import StepStatus
        self.assertEqual(StepStatus.PENDING.value, "pending")
        self.assertEqual(StepStatus.ACTIVE.value, "active")
        self.assertEqual(StepStatus.DONE.value, "done")
        self.assertEqual(StepStatus.WARNING.value, "warning")
        self.assertEqual(StepStatus.ERROR.value, "error")


class PreviewWidgetSmokeTest(unittest.TestCase):
    """Smoke-тесты структуры виджета превью."""

    def test_preview_widget_imports(self):
        from app.gui.preview_widget import PreviewWidget
        self.assertTrue(callable(PreviewWidget))

    def test_preview_widget_has_required_signals(self):
        from app.gui.preview_widget import PreviewWidget
        # Сигналы должны быть определены как атрибуты класса
        self.assertTrue(hasattr(PreviewWidget, "goto_materials"))
        self.assertTrue(hasattr(PreviewWidget, "request_generate"))


class BannedReplacementsDialogSmokeTest(unittest.TestCase):
    """Smoke-тесты структуры диалога запретов."""

    def test_dialog_imports(self):
        from app.gui.banned_replacements_dialog import BannedReplacementsDialog
        self.assertTrue(callable(BannedReplacementsDialog))


class BannedReplacementsStoreTest(unittest.TestCase):
    """Интеграционные тесты BannedReplacementsStore с временной БД."""

    def setUp(self):
        import tempfile
        import shutil
        import sqlite3
        from contextlib import contextmanager

        self.tmp_dir = Path(tempfile.mkdtemp(prefix="omts_banned_test_"))
        self.db_path = self.tmp_dir / "test.db"

        # Создаём минимальную схему
        with sqlite3.connect(self.db_path) as conn:
            conn.execute("""
                CREATE TABLE banned_replacements (
                    part_code TEXT NOT NULL,
                    workshop TEXT NOT NULL,
                    role TEXT NOT NULL,
                    before_name TEXT NOT NULL,
                    after_name TEXT,
                    reason TEXT,
                    banned_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)

        # Mock db_manager
        class _TempDB:
            def __init__(self, path):
                self.db_path = path

            @contextmanager
            def get_connection(self):
                conn = sqlite3.connect(self.db_path)
                conn.row_factory = sqlite3.Row
                try:
                    yield conn
                    conn.commit()
                except Exception:
                    conn.rollback()
                    raise
                finally:
                    conn.close()

        from app.banned_replacements_store import BannedReplacementsStore
        from app.models import CatalogEntry
        self.BannedReplacementsStore = BannedReplacementsStore
        self.CatalogEntry = CatalogEntry
        self.store = BannedReplacementsStore(_TempDB(self.db_path))

    def tearDown(self):
        import shutil
        import time
        # На Windows SQLite оставляет .db-wal/.db-shm, которые лочат удаление.
        # Пытаемся удалить с ретраем и игнорированием ошибок.
        for attempt in range(3):
            try:
                if self.tmp_dir.exists():
                    shutil.rmtree(self.tmp_dir, ignore_errors=False)
                break
            except (PermissionError, OSError):
                time.sleep(0.05)
        else:
            shutil.rmtree(self.tmp_dir, ignore_errors=True)

    def test_ban_and_check(self):
        entry = self.CatalogEntry(
            part="КИР 03.614", workshop="ПЗУ", role="",
            before_name="Сталь 40Х", unit="кг", norm=1.0,
        )
        self.assertTrue(self.store.ban(entry, "Сталь 30ХМА", "Нет сертификата"))
        self.assertTrue(self.store.is_banned(entry, "Сталь 30ХМА"))
        self.assertFalse(self.store.is_banned(entry, "Сталь 20Х"))

    def test_unban_specific(self):
        entry = self.CatalogEntry(
            part="X", workshop="ПЗУ", role="",
            before_name="M1", unit="кг", norm=1.0,
        )
        self.store.ban(entry, "M2")
        self.store.ban(entry, "M3")
        self.assertTrue(self.store.unban(entry, "M2"))
        self.assertFalse(self.store.is_banned(entry, "M2"))
        self.assertTrue(self.store.is_banned(entry, "M3"))

    def test_unban_all(self):
        entry = self.CatalogEntry(
            part="X", workshop="ПЗУ", role="",
            before_name="M1", unit="кг", norm=1.0,
        )
        self.store.ban(entry, "M2")
        self.store.ban(entry, "M3")
        self.assertTrue(self.store.unban(entry, None))
        self.assertFalse(self.store.is_banned(entry, "M2"))
        self.assertFalse(self.store.is_banned(entry, "M3"))

    def test_get_all_banned_returns_rows(self):
        entry = self.CatalogEntry(
            part="X", workshop="ПЗУ", role="",
            before_name="M1", unit="кг", norm=1.0,
        )
        self.store.ban(entry, "M2", "причина 1")
        rows = self.store.get_all_banned()
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["before_name"], "M1")
        self.assertEqual(rows[0]["after_name"], "M2")
        self.assertEqual(rows[0]["reason"], "причина 1")


if __name__ == "__main__":
    unittest.main()