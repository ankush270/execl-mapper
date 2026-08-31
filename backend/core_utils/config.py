import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

class Config:
    BASE_DIR = Path(__file__).resolve().parent.parent
    MONGO_URI = os.getenv("MONGO_URI", "mongodb://localhost:27017")
    DATABASE_NAME = os.getenv("DATABASE_NAME", "json_excel_db")
    TEMPLATE_FOLDER = BASE_DIR / "templates"
    GENERATED_FOLDER = BASE_DIR / "generated_excels"
    MAX_CONTENT_LENGTH = 2 * 1024 * 1024 * 1024