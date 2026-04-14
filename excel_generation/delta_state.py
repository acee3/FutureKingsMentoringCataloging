"""Persist Microsoft Graph delta links between Excel generation runs."""

import json
from pathlib import Path

from app_types import DeltaState, DeltaSourceState, DriveSource
from configuration import OUTPUT_DIR

DELTA_STATE_FILENAME = "workshop_catalog_delta_state.json"
DELTA_STATE_PATH = Path(OUTPUT_DIR) / DELTA_STATE_FILENAME


def get_delta_source_key(source: DriveSource) -> str:
    """Return the stable key used to store delta state for one source."""
    return f"{source['drive_id']}:{source.get('folder_id', 'root')}"


def delta_state_exists() -> bool:
    """Return `True` when saved delta state is available."""
    return DELTA_STATE_PATH.exists()


def load_delta_state() -> DeltaState:
    """Load saved Microsoft Graph delta state."""
    if not DELTA_STATE_PATH.exists():
        return {"sources": {}}

    with DELTA_STATE_PATH.open("r", encoding="utf-8") as state_file:
        data = json.load(state_file)
    return {"sources": data.get("sources", {})}


def save_delta_state(state: DeltaState) -> None:
    """Atomically save Microsoft Graph delta state."""
    DELTA_STATE_PATH.parent.mkdir(parents=True, exist_ok=True)
    temp_path = DELTA_STATE_PATH.with_suffix(".tmp")
    with temp_path.open("w", encoding="utf-8") as state_file:
        json.dump(state, state_file, ensure_ascii=True, indent=2)
    temp_path.replace(DELTA_STATE_PATH)


def clear_delta_state() -> None:
    """Delete saved delta state."""
    if DELTA_STATE_PATH.exists():
        DELTA_STATE_PATH.unlink()


def build_delta_source_state(source: DriveSource, delta_link: str) -> DeltaSourceState:
    """Build the saved state object for one source."""
    source_state: DeltaSourceState = {
        "source_name": source["name"],
        "drive_id": source["drive_id"],
        "delta_link": delta_link,
    }
    if source.get("folder"):
        source_state["source_folder"] = source["folder"]
    if source.get("folder_id"):
        source_state["folder_id"] = source["folder_id"]
    return source_state


def merge_delta_links(
    state: DeltaState,
    drive_sources: list[DriveSource],
    delta_links: dict[str, str],
) -> DeltaState:
    """Return state updated with newly collected delta links."""
    next_state: DeltaState = {"sources": dict(state.get("sources", {}))}
    sources_by_key = {get_delta_source_key(source): source for source in drive_sources}

    for source_key, delta_link in delta_links.items():
        source = sources_by_key[source_key]
        next_state["sources"][source_key] = build_delta_source_state(
            source,
            delta_link,
        )

    return next_state
