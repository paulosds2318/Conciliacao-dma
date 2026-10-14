import os
from dotenv import load_dotenv

load_dotenv()

CORAL_API_URL = os.getenv("CORAL_API_URL", "https://api.coral.exemplo.com/v1/dma/contratos")
CORAL_AUTH_TOKEN = os.getenv("CORAL_AUTH_TOKEN", "")

MODO_TESTE_API = os.getenv("MODO_TESTE_API", "true").strip().lower() in {"1", "true", "yes", "sim"}
