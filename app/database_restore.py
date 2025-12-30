"""
Утилита для восстановления базы данных из резервной копии
"""
import shutil
from pathlib import Path
from typing import List, Optional
from datetime import datetime

from app.config import DATABASE_PATH, DATA_DIR


def list_backup_files() -> List[Path]:
    """Получить список всех резервных копий БД, отсортированных по дате (новые первыми)"""
    backups = []
    
    if not DATA_DIR.exists():
        return backups
    
    # Ищем все файлы с расширением .bak-*
    for file_path in DATA_DIR.glob("app.db.bak-*"):
        if file_path.is_file():
            backups.append(file_path)
    
    # Сортируем по времени модификации (новые первыми)
    backups.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    
    return backups


def get_backup_timestamp(backup_path: Path) -> Optional[str]:
    """Извлечь timestamp из имени файла резервной копии"""
    try:
        # Формат: app.db.bak-YYYYMMDD-HHMMSS
        name = backup_path.name
        if name.startswith("app.db.bak-"):
            timestamp_str = name[12:]  # Убираем "app.db.bak-"
            # Парсим дату для красивого отображения
            if len(timestamp_str) >= 15:  # YYYYMMDD-HHMMSS
                year = timestamp_str[0:4]
                month = timestamp_str[4:6]
                day = timestamp_str[6:8]
                hour = timestamp_str[9:11]
                minute = timestamp_str[11:13]
                second = timestamp_str[13:15] if len(timestamp_str) > 13 else "00"
                return f"{day}.{month}.{year} {hour}:{minute}:{second}"
    except Exception:
        pass
    
    # Если не удалось распарсить, возвращаем время модификации файла
    try:
        mtime = backup_path.stat().st_mtime
        dt = datetime.fromtimestamp(mtime)
        return dt.strftime("%d.%m.%Y %H:%M:%S")
    except Exception:
        return backup_path.name


def restore_database(backup_path: Path) -> bool:
    """
    Восстановить базу данных из резервной копии
    
    Args:
        backup_path: Путь к файлу резервной копии
    
    Returns:
        True если восстановление успешно, False в противном случае
    """
    if not backup_path.exists():
        return False
    
    try:
        # Создаём резервную копию текущей БД перед восстановлением
        if DATABASE_PATH.exists():
            current_backup = DATABASE_PATH.parent / f"app.db.before-restore-{datetime.now().strftime('%Y%m%d-%H%M%S')}"
            shutil.copy2(DATABASE_PATH, current_backup)
        
        # Восстанавливаем из резервной копии
        shutil.copy2(backup_path, DATABASE_PATH)
        return True
    except Exception as e:
        print(f"Ошибка при восстановлении БД: {e}")
        return False


def delete_backup(backup_path: Path) -> bool:
    """Удалить резервную копию"""
    try:
        if backup_path.exists():
            backup_path.unlink()
            return True
        return False
    except Exception as e:
        print(f"Ошибка при удалении резервной копии: {e}")
        return False

