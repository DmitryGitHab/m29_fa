# config.py
import os
from pathlib import Path

# Путь к папке uploads
UPLOADS_DIR = Path("uploads")

# Интервал очистки в секундах (по умолчанию 24 часа)
CLEANUP_INTERVAL = 24 * 3600  # 24 часа

# Создаем папку uploads, если она не существует
if not UPLOADS_DIR.exists():
    os.makedirs(UPLOADS_DIR)