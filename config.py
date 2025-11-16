import os
from dotenv import load_dotenv

load_dotenv()

TELEGRAM_BOT_TOKEN = os.getenv('TELEGRAM_BOT_TOKEN')
OLLAMA_API_KEY = os.getenv('OLLAMA_API_KEY')
OLLAMA_MODEL = "kimi-k2:1t-cloud"