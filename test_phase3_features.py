"""
Тесты для Фазы 3 — новые функции:
- DraftManager: save/load/discard/import/export
- Themes: загрузка QSS, определение темы системы
- AuditLog: добавление записей, формат
"""
import json
import sys
import unittest
from pathlib import Path
from datetime import datetime, timedelta
from unittest.mock import MagicMock, patch

sys.path.insert(0, str(Path(__file__).parent))


class DraftManagerTest(unittest.TestCase):
    """DraftManager: save, load, discard, export/import."""

    def setUp(self):
        import tempfile
        self.tmp_dir = Path(tempfile.mkdtemp(prefix="omts_draft_test_"))
        from app.draft_manager import DraftManager
        self.DraftManager = DraftManager
        self.dm = DraftManager(drafts_dir=self.tmp_dir)

    def tearDown(self):
        import shutil
        for attempt in range(3):
            try:
                if self.tmp_dir.exists():
                    shutil.rmtree(self.tmp_dir, ignore_errors=False)
                break
            except (PermissionError, OSError):
                import time
                time.sleep(0.05)
        else:
            shutil.rmtree(self.tmp_dir, ignore_errors=True)

    def test_no_draft_initially(self):
        self.assertFalse(self.dm.has_draft())
        self.assertIsNone(self.dm.get_draft_info())
        self.assertIsNone(self.dm.load())

    def test_save_and_load(self):
        from app.models import DocumentData, PartChanges, MaterialChange, CatalogEntry

        data = DocumentData(
            implementation_date=datetime(2026, 6, 25).date(),
            reason="Тест",
            document_month=6,
            document_year=2026,
        )
        entry = CatalogEntry(
            part="TEST-001", workshop="ПЗУ", role="",
            before_name="M1", unit="кг", norm=1.0,
        )
        data.part_changes.append(PartChanges(
            part="TEST-001",
            materials=[MaterialChange(catalog_entry=entry, is_changed=True)],
        ))

        path = self.dm.save(data, source="manual")
        self.assertTrue(path.exists())
        self.assertTrue(self.dm.has_draft())

        info = self.dm.get_draft_info()
        self.assertIsNotNone(info)
        self.assertEqual(info.source, "manual")
        self.assertEqual(info.part_count, 1)

        loaded = self.dm.load()
        self.assertIsNotNone(loaded)
        self.assertEqual(len(loaded.part_changes), 1)
        self.assertEqual(loaded.part_changes[0].part, "TEST-001")
        self.assertEqual(loaded.reason, "Тест")

    def test_discard(self):
        from app.models import DocumentData
        self.dm.save(DocumentData(), source="manual")
        self.assertTrue(self.dm.has_draft())
        self.assertTrue(self.dm.discard())
        self.assertFalse(self.dm.has_draft())

    def test_export_and_import(self):
        from app.models import DocumentData, PartChanges

        data = DocumentData(reason="Экспорт-тест")
        data.part_changes.append(PartChanges(part="X"))
        data.products = ["Машина 1"]

        export_path = self.tmp_dir / "exported.json"
        self.dm.export_to(data, export_path)
        self.assertTrue(export_path.exists())

        # Проверяем структуру файла
        with open(export_path, "r", encoding="utf-8") as f:
            payload = json.load(f)
        self.assertEqual(payload["source"], "export")
        self.assertIn("document_data", payload)
        self.assertEqual(payload["document_data"]["reason"], "Экспорт-тест")
        self.assertEqual(payload["document_data"]["products"], ["Машина 1"])

        # Импортируем обратно
        imported = self.dm.import_from(export_path)
        self.assertIsNotNone(imported)
        self.assertEqual(imported.reason, "Экспорт-тест")
        self.assertEqual(imported.products, ["Машина 1"])
        self.assertEqual(len(imported.part_changes), 1)

    def test_save_creates_tmp_file_atomically(self):
        from app.models import DocumentData
        self.dm.save(DocumentData(), source="manual")
        # Tmp-файл не должен остаться
        tmp_path = self.tmp_dir / "current_draft.json.tmp"
        self.assertFalse(tmp_path.exists())


class DraftInfoTest(unittest.TestCase):
    """DraftInfo: human-readable возраст."""

    def test_age_human_recent(self):
        from app.draft_manager import DraftInfo
        info = DraftInfo(
            path=Path("/tmp/x.json"),
            saved_at=datetime.now() - timedelta(seconds=30),
            source="auto",
            part_count=0,
        )
        self.assertEqual(info.age_human(), "только что")

    def test_age_human_minutes(self):
        from app.draft_manager import DraftInfo
        info = DraftInfo(
            path=Path("/tmp/x.json"),
            saved_at=datetime.now() - timedelta(minutes=5),
            source="auto",
            part_count=0,
        )
        self.assertEqual(info.age_human(), "5 мин назад")

    def test_age_human_hours(self):
        from app.draft_manager import DraftInfo
        info = DraftInfo(
            path=Path("/tmp/x.json"),
            saved_at=datetime.now() - timedelta(hours=3),
            source="auto",
            part_count=0,
        )
        self.assertEqual(info.age_human(), "3 ч назад")


class AuditLogTest(unittest.TestCase):
    """AuditLog: добавление записей, формат, last, clear."""

    def test_add_and_retrieve(self):
        from app.audit_log import AuditLog
        log = AuditLog()
        log.info("action1", "details1")
        log.success("action2")
        log.warning("action3", "warn details")

        entries = log.entries()
        self.assertEqual(len(entries), 3)
        self.assertEqual(entries[0].action, "action1")
        self.assertEqual(entries[0].level, "info")
        self.assertEqual(entries[1].action, "action2")
        self.assertEqual(entries[1].level, "success")
        self.assertEqual(entries[2].level, "warning")

    def test_max_entries_capped(self):
        from app.audit_log import AuditLog
        log = AuditLog(max_entries=5)
        for i in range(10):
            log.info(f"action_{i}")
        entries = log.entries()
        self.assertEqual(len(entries), 5)
        # Старые записи отбрасываются, остаются последние
        self.assertEqual(entries[-1].action, "action_9")
        self.assertEqual(entries[0].action, "action_5")

    def test_format(self):
        from app.audit_log import LogEntry
        entry = LogEntry(
            timestamp=datetime(2026, 6, 25, 14, 38, 0),
            action="add_part",
            details="КИР 03.614",
            level="info",
        )
        formatted = entry.format()
        self.assertIn("14:38:00", formatted)
        self.assertIn("add_part", formatted)
        self.assertIn("КИР 03.614", formatted)
        self.assertIn("ℹ", formatted)

    def test_last_returns_most_recent(self):
        from app.audit_log import AuditLog
        log = AuditLog()
        self.assertIsNone(log.last())
        log.info("first")
        log.info("second")
        log.info("third")
        self.assertEqual(log.last().action, "third")

    def test_clear(self):
        from app.audit_log import AuditLog
        log = AuditLog()
        log.info("a")
        log.info("b")
        self.assertEqual(len(log.entries()), 2)
        log.clear()
        self.assertEqual(len(log.entries()), 0)


class ThemesTest(unittest.TestCase):
    """Themes: загрузка QSS, валидация."""

    def test_valid_themes_constant(self):
        from app.gui.themes import VALID_THEMES, THEME_LIGHT, THEME_DARK, THEME_SYSTEM
        self.assertIn(THEME_LIGHT, VALID_THEMES)
        self.assertIn(THEME_DARK, VALID_THEMES)
        self.assertIn(THEME_SYSTEM, VALID_THEMES)

    def test_load_qss_returns_string(self):
        from app.gui.themes import _load_qss
        # Должен вернуть непустую строку для light и dark
        light_qss = _load_qss("light")
        self.assertIsInstance(light_qss, str)
        self.assertGreater(len(light_qss), 100)

        dark_qss = _load_qss("dark")
        self.assertIsInstance(dark_qss, str)
        self.assertGreater(len(dark_qss), 100)

    def test_get_effective_theme_expands_system(self):
        from app.gui.themes import get_effective_theme
        # В headless-окружении system скорее всего даст light или dark
        eff = get_effective_theme("system")
        self.assertIn(eff, ("light", "dark"))

        self.assertEqual(get_effective_theme("light"), "light")
        self.assertEqual(get_effective_theme("dark"), "dark")


if __name__ == "__main__":
    unittest.main()