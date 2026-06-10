import sqlite3
import unittest
from contextlib import contextmanager
from pathlib import Path
from uuid import uuid4

from app.numbering import NumberingManager


class TempDatabaseManager:
    def __init__(self, db_path: Path):
        self.db_path = db_path
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(
                """
                CREATE TABLE numbering (
                    year INTEGER NOT NULL,
                    month INTEGER NOT NULL,
                    last_number INTEGER NOT NULL DEFAULT 0,
                    PRIMARY KEY (year, month)
                )
                """
            )
            cursor.execute(
                """
                CREATE TABLE documents (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    document_number INTEGER NOT NULL,
                    year INTEGER NOT NULL,
                    month INTEGER NOT NULL DEFAULT 1,
                    data_json TEXT NOT NULL,
                    output_file_path TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE(document_number, year, month)
                )
                """
            )

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


class NumberingManagerTest(unittest.TestCase):
    def setUp(self):
        self.temp_dir = Path(__file__).parent / ".tmp_tests" / uuid4().hex
        self.temp_dir.mkdir(parents=True, exist_ok=True)
        self.db_path = self.temp_dir / "app.db"
        self.db_manager = TempDatabaseManager(self.db_path)
        self.numbering = NumberingManager(self.db_manager)
        self.numbering._use_db = True

    def tearDown(self):
        import shutil
        if self.temp_dir.exists():
            shutil.rmtree(self.temp_dir)

    def test_current_number_respects_manual_counter_below_saved_documents(self):
        with self.db_manager.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("INSERT INTO numbering (year, month, last_number) VALUES (?, ?, ?)", (2026, 1, 130))
            cursor.execute(
                "INSERT INTO documents (document_number, year, month, data_json) VALUES (?, ?, ?, ?)",
                (133, 2026, 1, "{}"),
            )

        self.assertEqual(self.numbering.get_current_number(2026, 1), 131)

        with self.db_manager.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT last_number FROM numbering WHERE year = ? AND month = ?", (2026, 1))
            self.assertEqual(cursor.fetchone()["last_number"], 130)

    def test_mark_number_as_used_does_not_lower_counter(self):
        with self.db_manager.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("INSERT INTO numbering (year, month, last_number) VALUES (?, ?, ?)", (2026, 1, 133))

        self.numbering.mark_number_as_used(131, 2026, 1)

        with self.db_manager.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT last_number FROM numbering WHERE year = ? AND month = ?", (2026, 1))
            self.assertEqual(cursor.fetchone()["last_number"], 133)

    def test_next_number_uses_manual_counter_even_below_saved_documents(self):
        with self.db_manager.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("INSERT INTO numbering (year, month, last_number) VALUES (?, ?, ?)", (2026, 1, 130))
            cursor.execute(
                "INSERT INTO documents (document_number, year, month, data_json) VALUES (?, ?, ?, ?)",
                (133, 2026, 1, "{}"),
            )

        self.assertEqual(self.numbering.get_next_number(2026, 1), 131)

    def test_set_number_can_lower_counter(self):
        with self.db_manager.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("INSERT INTO numbering (year, month, last_number) VALUES (?, ?, ?)", (2026, 1, 133))

        self.numbering.set_number(131, 2026, 1)

        self.assertEqual(self.numbering.get_current_number(2026, 1), 131)

    def test_different_months_have_independent_counters(self):
        with self.db_manager.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("INSERT INTO numbering (year, month, last_number) VALUES (?, ?, ?)", (2026, 1, 5))
            cursor.execute("INSERT INTO numbering (year, month, last_number) VALUES (?, ?, ?)", (2026, 2, 10))

        self.assertEqual(self.numbering.get_current_number(2026, 1), 6)
        self.assertEqual(self.numbering.get_current_number(2026, 2), 11)

    def test_next_number_per_month(self):
        jan_num = self.numbering.get_next_number(2026, 1)
        feb_num = self.numbering.get_next_number(2026, 2)

        self.assertEqual(jan_num, 1)
        self.assertEqual(feb_num, 1)

        jan_num_2 = self.numbering.get_next_number(2026, 1)
        self.assertEqual(jan_num_2, 2)

    def test_format_number(self):
        self.assertEqual(NumberingManager.format_number(1, 2026, 5), "01-26-5")
        self.assertEqual(NumberingManager.format_number(12, 2025, 1), "12-25-1")
        self.assertEqual(NumberingManager.format_number(3, 2024, 100), "03-24-100")

    def test_get_display_number_from_model(self):
        from app.models import DocumentData
        doc = DocumentData(document_number=5, document_month=1, document_year=2026)
        self.assertEqual(doc.get_display_number(), "01-26-5")


if __name__ == "__main__":
    unittest.main()
