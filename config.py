import os
from dotenv import load_dotenv

load_dotenv()

EXCEL_FILE_PATH = os.getenv("EXCEL_FILE_PATH")

if not EXCEL_FILE_PATH:
    raise ValueError("Переменная окружения BOT_TOKEN не установлена")