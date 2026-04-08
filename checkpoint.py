import json
from pathlib import Path
from typing import Any

from app_types import RunCheckpoint
from configuration import OUTPUT_DIR
from microsoft.types import GraphDriveItem

CHECKPOINT_FILENAME = "workshop_catalog_checkpoint.json"
CHECKPOINT_PATH = Path(OUTPUT_DIR) / CHECKPOINT_FILENAME

def checkpoint_exists() -> bool:
    return CHECKPOINT_PATH.exists()


def load_checkpoint() -> RunCheckpoint:
    with CHECKPOINT_PATH.open("r", encoding="utf-8") as checkpoint_file:
        data = json.load(checkpoint_file)
    return {
        "processed_rows": data.get("processed_rows", []),
        "pending_items": data.get("pending_items", []),
    }


def save_checkpoint(
    processed_rows: list[dict[str, Any]],
    pending_items: list[GraphDriveItem],
) -> None:
    CHECKPOINT_PATH.parent.mkdir(parents=True, exist_ok=True)
    temp_path = CHECKPOINT_PATH.with_suffix(".tmp")
    payload: RunCheckpoint = {
        "processed_rows": processed_rows,
        "pending_items": pending_items,
    }
    with temp_path.open("w", encoding="utf-8") as checkpoint_file:
        json.dump(payload, checkpoint_file, ensure_ascii=True, indent=2)
    temp_path.replace(CHECKPOINT_PATH)


def clear_checkpoint() -> None:
    if CHECKPOINT_PATH.exists():
        CHECKPOINT_PATH.unlink()
