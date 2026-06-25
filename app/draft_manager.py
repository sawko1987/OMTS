"""
Менеджер черновиков документа.

Сохраняет текущее состояние DocumentData в JSON-файл в data/drafts/.
Используется для автосохранения (каждые N секунд) и ручного сохранения (Ctrl+S).

Структура файла:
{
    "saved_at": "2026-06-25T14:38:00",
    "document_data": {...},     // сериализованный DocumentData
    "schema_version": 1,
    "source": "auto" | "manual"
}

При запуске приложения MainWindow спрашивает пользователя, восстанавливать
ли последний черновик (если он свежее последнего сгенерированного документа).
"""
from __future__ import annotations

import json
import logging
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Optional

from app.config import DATA_DIR
from app.models import DocumentData
from app.serialization import DocumentSerializer

logger = logging.getLogger(__name__)

SCHEMA_VERSION = 1


@dataclass
class DraftInfo:
    """Метаданные черновика (без самих данных документа)."""
    path: Path
    saved_at: datetime
    source: str  # "auto" | "manual"
    part_count: int

    def age_seconds(self) -> float:
        return (datetime.now() - self.saved_at).total_seconds()

    def age_human(self) -> str:
        seconds = self.age_seconds()
        if seconds < 60:
            return "только что"
        if seconds < 3600:
            m = int(seconds // 60)
            return f"{m} мин назад"
        if seconds < 86400:
            h = int(seconds // 3600)
            return f"{h} ч назад"
        d = int(seconds // 86400)
        return f"{d} дн назад"


class DraftManager:
    """Управление черновиками документа."""

    DEFAULT_DRAFT_FILENAME = "current_draft.json"

    def __init__(self, drafts_dir: Optional[Path] = None):
        self.drafts_dir = drafts_dir or (DATA_DIR / "drafts")
        self.drafts_dir.mkdir(parents=True, exist_ok=True)
        self.default_draft_path = self.drafts_dir / self.DEFAULT_DRAFT_FILENAME

    # ---- Сохранение ----

    def save(self, document_data: DocumentData, source: str = "auto") -> Path:
        """
        Сохранить текущий документ как черновик.

        Args:
            document_data: данные документа
            source: "auto" для автосохранения, "manual" для Ctrl+S

        Returns:
            Путь к сохранённому файлу
        """
        try:
            payload = {
                "schema_version": SCHEMA_VERSION,
                "saved_at": datetime.now().isoformat(),
                "source": source,
                "document_data": json.loads(DocumentSerializer.serialize(document_data)),
            }
        except Exception as e:
            logger.error(f"Не удалось сериализовать документ для черновика: {e}")
            raise

        # Пишем во временный файл, потом rename — чтобы не оставлять половинчатый файл
        tmp_path = self.default_draft_path.with_suffix(".json.tmp")
        try:
            with open(tmp_path, "w", encoding="utf-8") as f:
                json.dump(payload, f, ensure_ascii=False, indent=2)
            tmp_path.replace(self.default_draft_path)
            logger.debug(f"Черновик сохранён: {self.default_draft_path}")
            return self.default_draft_path
        except Exception as e:
            logger.error(f"Ошибка записи черновика: {e}")
            if tmp_path.exists():
                try:
                    tmp_path.unlink()
                except OSError:
                    pass
            raise

    # ---- Загрузка ----

    def has_draft(self) -> bool:
        return self.default_draft_path.exists()

    def get_draft_info(self) -> Optional[DraftInfo]:
        """Получить метаданные черновика без загрузки самих данных."""
        if not self.has_draft():
            return None
        try:
            with open(self.default_draft_path, "r", encoding="utf-8") as f:
                payload = json.load(f)
            saved_at = datetime.fromisoformat(payload["saved_at"])
            doc_data = payload.get("document_data", {})
            part_count = len(doc_data.get("part_changes", []))
            return DraftInfo(
                path=self.default_draft_path,
                saved_at=saved_at,
                source=payload.get("source", "auto"),
                part_count=part_count,
            )
        except Exception as e:
            logger.warning(f"Не удалось прочитать метаданные черновика: {e}")
            return None

    def load(self, catalog_loader=None) -> Optional[DocumentData]:
        """Загрузить черновик. Возвращает None если файл отсутствует или повреждён."""
        if not self.has_draft():
            return None
        try:
            with open(self.default_draft_path, "r", encoding="utf-8") as f:
                payload = json.load(f)
            schema_version = payload.get("schema_version", 0)
            if schema_version > SCHEMA_VERSION:
                logger.warning(
                    f"Черновик создан в более новой версии (schema={schema_version}), "
                    f"текущая {SCHEMA_VERSION}. Загрузка может быть неточной."
                )
            doc_json = json.dumps(payload.get("document_data", {}), ensure_ascii=False)
            return DocumentSerializer.deserialize(doc_json, catalog_loader=catalog_loader)
        except Exception as e:
            logger.error(f"Не удалось загрузить черновик: {e}")
            return None

    # ---- Удаление ----

    def discard(self) -> bool:
        """Удалить текущий черновик. Возвращает True если файл был удалён."""
        if self.default_draft_path.exists():
            try:
                self.default_draft_path.unlink()
                return True
            except OSError as e:
                logger.warning(f"Не удалось удалить черновик: {e}")
                return False
        return False

    # ---- Импорт/экспорт ----

    def export_to(self, document_data: DocumentData, file_path: Path) -> None:
        """Сохранить документ в произвольный JSON-файл (экспорт черновика)."""
        file_path.parent.mkdir(parents=True, exist_ok=True)
        payload = {
            "schema_version": SCHEMA_VERSION,
            "saved_at": datetime.now().isoformat(),
            "source": "export",
            "document_data": json.loads(DocumentSerializer.serialize(document_data)),
        }
        with open(file_path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)

    def import_from(self, file_path: Path, catalog_loader=None) -> Optional[DocumentData]:
        """Загрузить документ из произвольного JSON-файла."""
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                payload = json.load(f)
            doc_json = json.dumps(payload.get("document_data", {}), ensure_ascii=False)
            return DocumentSerializer.deserialize(doc_json, catalog_loader=catalog_loader)
        except Exception as e:
            logger.error(f"Не удалось импортировать файл {file_path}: {e}")
            return None