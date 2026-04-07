import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parents[1]
ARTFATOS_DIR = BASE_DIR / "artfatos"
DB_PATH = Path(os.getenv("DB_PATH", str(BASE_DIR / "comissoes" / "comissoes.db")))
SMTP_HOST = ""
SMTP_PORT = 587
SMTP_USER = ""
SMTP_PASS = ""
SMTP_FROM = ""
