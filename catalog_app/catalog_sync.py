"""Manual SharePoint-to-Postgres reload workflow."""

from __future__ import annotations

import logging
from typing import Any

from dotenv import load_dotenv

from catalog_app.db import (
    count_presentations,
    delete_presentations,
    get_source_delta_links,
    get_sync_status,
    update_sync_status,
    upsert_presentation_source,
    upsert_presentations,
)
from catalog_app.embeddings import embed_texts
from catalog_app.generation.column_helpers import (
    build_presentation_row,
    create_presentation_metadata_model,
    get_ai_generation_inputs,
)
from catalog_app.generation.configuration import get_presentation_columns
from catalog_app.generation.generators import (
    GeneratorRegistry,
    get_configured_source_path,
)
from catalog_app.generation.llm_work import generate_ai_metadata, get_openai_client
from catalog_app.generation.microsoft import collect_drive_delta, excel_setup
from catalog_app.generation.microsoft.types import GraphDriveItem
from catalog_app.app_types import IndexedWorkbookRow

load_dotenv()

logger = logging.getLogger(__name__)

WORKBOOK_NAME = "workshop_catalog.xlsx"
SHEET_NAME = "Presentations"


def _get_source_key(source: dict[str, Any]) -> str:
    return f"{source['drive_id']}:{source.get('folder', 'root')}"


def _get_item_drive_id(item: GraphDriveItem, fallback_drive_id: str) -> str:
    return item.get("parentReference", {}).get("driveId") or fallback_drive_id


def _get_source_id(item: GraphDriveItem, fallback_drive_id: str) -> str:
    return f"{_get_item_drive_id(item, fallback_drive_id)}:{item['id']}"


def _is_processable_pptx_item(item: GraphDriveItem) -> bool:
    return "file" in item and item.get("name", "").lower().endswith(".pptx")


def _attach_source_context(
    item: GraphDriveItem,
    *,
    source_name: str,
    source_folder: str = "",
) -> GraphDriveItem:
    item["configuredSourceName"] = source_name
    if source_folder:
        item["configuredSourceFolder"] = source_folder
    return item


def _build_searchable_text(metadata: dict[str, Any]) -> str:
    preferred_fields = [
        "name",
        "theme*",
        "subtheme*",
        "description*",
        "audience*",
        "slide_texts",
        "presentation_path",
    ]
    ordered_keys: list[str] = []
    seen_keys: set[str] = set()
    for key in preferred_fields + list(metadata):
        if key in seen_keys:
            continue
        seen_keys.add(key)
        ordered_keys.append(key)

    parts: list[str] = []
    for key in ordered_keys:
        value = metadata.get(key)
        if value in (None, ""):
            continue
        if isinstance(value, list):
            text_value = "\n".join(str(item) for item in value if item not in (None, ""))
        else:
            text_value = str(value)
        if text_value.strip():
            parts.append(f"{key}: {text_value.strip()}")
    return "\n".join(parts)


def _collect_delta_items(
    drive_sources: list[dict[str, Any]],
    headers: dict[str, str],
) -> tuple[list[GraphDriveItem], set[str], dict[str, str]]:
    logger.info(
        "Gathering SharePoint presentation rows from %s configured source(s).",
        len(drive_sources),
    )
    saved_delta_links = get_source_delta_links()
    latest_items_by_source_id: dict[str, GraphDriveItem] = {}
    removed_source_ids: set[str] = set()
    next_delta_links: dict[str, str] = {}

    for source_index, source in enumerate(drive_sources, start=1):
        source_key = _get_source_key(source)
        source_label = (
            f"{source['name']}/{source['folder']}"
            if source.get("folder")
            else source["name"]
        )
        logger.info(
            "Gathering SharePoint presentation rows from source %s/%s: %s.",
            source_index,
            len(drive_sources),
            source_label,
        )
        upsert_presentation_source(
            source_key=source_key,
            source_name=source["name"],
            drive_id=source["drive_id"],
            folder_id=source.get("folder_id"),
            folder_path=source.get("folder"),
            delta_link=saved_delta_links.get(source_key),
        )

        result = collect_drive_delta(
            source,
            headers,
            saved_delta_links.get(source_key),
        )
        next_delta_links[source_key] = result["delta_link"]

        for item in result["items"]:
            if not item.get("id"):
                continue
            item = _attach_source_context(
                item,
                source_name=source["name"],
                source_folder=source.get("folder", ""),
            )
            source_id = _get_source_id(item, source["drive_id"])
            if "deleted" in item:
                removed_source_ids.add(source_id)
                latest_items_by_source_id.pop(source_id, None)
                continue
            if _is_processable_pptx_item(item):
                latest_items_by_source_id[source_id] = item

        logger.info(
            "Collected %s delta items from %s%s.",
            len(result["items"]),
            source["name"],
            f"/{source['folder']}" if source.get("folder") else "",
        )

    logger.info(
        "Finished gathering SharePoint presentation rows: %s changed presentation(s), %s removed item(s).",
        len(latest_items_by_source_id),
        len(removed_source_ids),
    )
    return list(latest_items_by_source_id.values()), removed_source_ids, next_delta_links


def _build_index_row(
    *,
    item: GraphDriveItem,
    row_metadata: dict[str, Any],
    source_id: str,
    row_number: int,
) -> IndexedWorkbookRow:
    drive_id = _get_item_drive_id(item, "")
    source_key = f"{drive_id}:{item.get('configuredSourceFolder', 'root')}"
    return IndexedWorkbookRow(
        source_id=source_id,
        title=str(row_metadata.get("name") or item.get("name") or source_id),
        workbook_name=WORKBOOK_NAME,
        sheet_name=SHEET_NAME,
        row_number=row_number,
        metadata=row_metadata,
        searchable_text=_build_searchable_text(row_metadata),
        source_key=source_key,
        drive_id=drive_id,
        item_id=item["id"],
        web_url=item.get("webUrl"),
        last_modified_at=item.get("lastModifiedDateTime"),
    )


def run_catalog_reload(*, mark_started: bool = True) -> None:
    """Reload changed SharePoint PowerPoints into the current Postgres catalog."""
    processed_items = 0
    removed_rows = 0
    pending_items: list[GraphDriveItem] = []

    if mark_started:
        update_sync_status(status="running", started=True)
    try:
        openai_client = get_openai_client()
        setup = excel_setup()
        headers = setup["headers"]
        drive_sources = setup["drive_sources"]
        if not drive_sources:
            raise ValueError("configuration.DRIVE_SOURCES must contain at least one drive")

        default_drive_source = next(
            (source for source in drive_sources if source.get("is_default")),
            drive_sources[0],
        )
        generator_registry = GeneratorRegistry(
            default_drive_id=default_drive_source["drive_id"],
            headers=headers,
        )
        presentation_columns = get_presentation_columns(generator_registry)
        metadata_model = create_presentation_metadata_model(presentation_columns)

        pending_items, removed_source_ids, next_delta_links = _collect_delta_items(
            drive_sources,
            headers,
        )
        removed_rows = delete_presentations(sorted(removed_source_ids))
        logger.info(
            "Preparing to build %s presentation row(s); removed %s stale row(s).",
            len(pending_items),
            removed_rows,
        )
        update_sync_status(
            status="running",
            total_items=len(pending_items),
            processed_items=0,
            indexed_rows=count_presentations(),
            removed_rows=removed_rows,
        )

        total_items = len(pending_items)
        for item in pending_items:
            source_id = _get_source_id(item, default_drive_source["drive_id"])
            logger.debug(
                "Building presentation row %s/%s: %s.",
                processed_items + 1,
                total_items,
                item.get("name", source_id),
            )
            slide_texts, number_of_slides, average_words_per_slide = (
                get_ai_generation_inputs(item, generator_registry)
            )
            ai_metadata = generate_ai_metadata(
                openai_client,
                name=item["name"],
                presentation_path=get_configured_source_path(item),
                slide_texts=slide_texts,
                number_of_slides=number_of_slides,
                average_words_per_slide=average_words_per_slide,
                response_model=metadata_model,
            )
            row_metadata = build_presentation_row(
                item,
                presentation_columns,
                ai_metadata,
            )
            index_row = _build_index_row(
                item=item,
                row_metadata=row_metadata,
                source_id=source_id,
                row_number=processed_items + 1,
            )
            embeddings = embed_texts([index_row.searchable_text])
            upsert_presentations([index_row], embeddings)

            processed_items += 1
            update_sync_status(
                status="running",
                total_items=len(pending_items),
                processed_items=processed_items,
                indexed_rows=count_presentations(),
                removed_rows=removed_rows,
            )
            if processed_items == 1 or processed_items % 5 == 0 or processed_items == total_items:
                logger.info(
                    "Built %s/%s presentation row(s).",
                    processed_items,
                    total_items,
                )

        for source in drive_sources:
            source_key = _get_source_key(source)
            upsert_presentation_source(
                source_key=source_key,
                source_name=source["name"],
                drive_id=source["drive_id"],
                folder_id=source.get("folder_id"),
                folder_path=source.get("folder"),
                delta_link=next_delta_links[source_key],
            )

        update_sync_status(
            status="succeeded",
            total_items=len(pending_items),
            processed_items=processed_items,
            indexed_rows=count_presentations(),
            removed_rows=removed_rows,
            finished=True,
        )
    except Exception as exc:
        logger.exception("Catalog reload failed.")
        update_sync_status(
            status="failed",
            total_items=len(pending_items),
            processed_items=processed_items,
            indexed_rows=count_presentations(),
            removed_rows=removed_rows,
            finished=True,
            error=str(exc),
        )
        raise


def is_reload_running() -> bool:
    return get_sync_status().status == "running"
