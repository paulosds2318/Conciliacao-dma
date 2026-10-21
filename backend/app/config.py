import os
from dotenv import load_dotenv
import sys
from pathlib import Path

_env_path = (
    Path(sys.executable).resolve().parent / ".env"
    if getattr(sys, "frozen", False)
    else Path(__file__).resolve().parents[1] / ".env"
)
load_dotenv(dotenv_path=_env_path)

CORAL_API_URL = os.getenv("CORAL_API_URL", "https://api.coral.exemplo.com/v1/dma/contratos")
def get_coral_auth_token():
    return os.getenv("CORAL_AUTH_TOKEN", "").strip()

MODO_TESTE_API = os.getenv("MODO_TESTE_API", "true").strip().lower() in {"1", "true", "yes", "sim"}
