"""关键词本地持久化：下次启动自动载入。"""
import json
from pathlib import Path

CONFIG_PATH = Path.home() / ".chat_audit_config.json"


def load_keywords() -> str:
    try:
        if CONFIG_PATH.exists():
            data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            return data.get("keywords", "") or ""
    except Exception:
        pass
    return ""


def save_keywords(text: str) -> None:
    try:
        CONFIG_PATH.write_text(
            json.dumps({"keywords": text}, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except Exception:
        pass
