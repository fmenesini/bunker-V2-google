import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent

PRODUCTION_MODE = False

DB_FILE = os.getenv("BUNKER_DB_FILE", str(BASE_DIR / "bunker_commerciale.db"))
DATABASE_URI = f"sqlite:///{DB_FILE}" if not PRODUCTION_MODE else os.getenv("DATABASE_URL")

DECIMAL_PRECISION = 5
DISPLAY_PRECISION = 3
