"""Command-line entry point for building the workshop catalog Excel file.

This file coordinates the whole export process:

1. Read configuration and authenticate with Microsoft Graph.
2. Discover PowerPoint files in the configured drives/folders.
3. Extract slide text and ask OpenAI for structured metadata.
4. Build rows for the Excel sheet.
5. Save progress to a checkpoint so the run can resume after interruptions.

If you are new to Python, `main()` is the best place to start reading because it
shows the full sequence of steps at a high level.
"""

import argparse
import logging

from checkpoint import (
    checkpoint_exists,
    clear_checkpoint,
    load_checkpoint,
    save_checkpoint,
)
from column_helpers import (
    build_presentation_row,
    create_presentation_metadata_model,
    get_excel_column_names,
    get_ai_generation_inputs,
)
from delta_state import (
    clear_delta_state,
    delta_state_exists,
    get_delta_source_key,
    load_delta_state,
    merge_delta_links,
    save_delta_state,
)
from dotenv import load_dotenv
from configuration import get_presentation_columns
from excel_maker import write_objects_to_excel
from generators import GeneratorRegistry, get_configured_source_path
from llm_work import generate_ai_metadata, get_openai_client
from microsoft import (
    collect_drive_delta,
    excel_setup,
    get_all_pptx_files,
    get_pptx_file,
)
from microsoft.types import GraphDriveItem
from workbook_rows import load_existing_workbook_rows

load_dotenv()
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

logger = logging.getLogger(__name__)


def parse_args() -> argparse.Namespace:
    """Read command-line flags for the script.

    Returns:
        argparse.Namespace: Parsed command-line options for discovery mode,
            checkpoint reset, and delta baseline behavior.
    """
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--restart-from-scratch",
        action="store_true",
        help=(
            "Ignore saved checkpoint and delta state, then rebuild the export "
            "from the beginning."
        ),
    )
    parser.add_argument(
        "--baseline-from-now",
        action="store_true",
        help=(
            "Save Microsoft Graph delta tokens from the current moment without "
            "modifying the workbook."
        ),
    )
    parser.add_argument(
        "--legacy-full-scan",
        action="store_true",
        help="Use the old recursive folder scan instead of Microsoft Graph delta.",
    )
    return parser.parse_args()


def dedupe_pptx_files(items: list[GraphDriveItem]) -> list[GraphDriveItem]:
    """Remove duplicate PowerPoint items from the discovery list.

    Microsoft Graph can occasionally surface the same presentation more than
    once depending on how folders are traversed or shared content is exposed.
    This function keeps the first copy it sees and ignores later duplicates.

    It deduplicates in two ways:
    - Exact item ID matches.
    - "Looks like the same file" matches based on name, size, and last
      modification time.

    Args:
        items: Raw PowerPoint metadata records returned from Graph.

    Returns:
        A new list with duplicates removed while preserving order.
    """
    unique_items: list[GraphDriveItem] = []
    seen_ids: set[str] = set()
    seen_signatures: set[tuple[str, int | None, str]] = set()

    for item in items:
        item_id = item["id"]
        if item_id in seen_ids:
            continue

        signature = (
            item["name"].strip().lower(),
            item.get("size"),
            item.get("lastModifiedDateTime", ""),
        )
        if signature in seen_signatures:
            continue

        seen_ids.add(item_id)
        seen_signatures.add(signature)
        unique_items.append(item)

    return unique_items


def get_item_row_key(item: GraphDriveItem) -> str:
    """Return the stable row key used for delta dedupe in memory."""
    drive_id = item.get("parentReference", {}).get("driveId", "")
    return f"{drive_id}:{item['id']}" if drive_id else item["id"]


def is_processable_pptx_item(item: GraphDriveItem) -> bool:
    """Return `True` when a delta item should produce an Excel row."""
    return "file" in item and item.get("name", "").lower().endswith(".pptx")


def attach_configured_source(
    item: GraphDriveItem,
    source_name: str,
    source_folder: str = "",
) -> GraphDriveItem:
    """Attach human-readable source context to a Graph item."""
    item["configuredSourceName"] = source_name
    if source_folder:
        item["configuredSourceFolder"] = source_folder
    return item


def classify_delta_items(
    items: list[GraphDriveItem],
) -> tuple[list[GraphDriveItem], set[str]]:
    """Split delta items into changed PowerPoints and deleted row IDs."""
    latest_by_item_key: dict[str, GraphDriveItem] = {}
    for item in items:
        item_id = item.get("id")
        if not item_id:
            continue
        latest_by_item_key[get_item_row_key(item)] = item

    changed_pptx_files: list[GraphDriveItem] = []
    removed_row_ids: set[str] = set()
    for item in latest_by_item_key.values():
        if "deleted" in item:
            removed_row_ids.add(item["id"])
            continue
        if is_processable_pptx_item(item):
            changed_pptx_files.append(item)

    return changed_pptx_files, removed_row_ids


def remove_rows_by_id(rows: list[dict], removed_row_ids: set[str]) -> list[dict]:
    """Remove rows whose Graph item IDs are no longer in scope."""
    if not removed_row_ids:
        return rows
    return [row for row in rows if str(row.get("id", "")) not in removed_row_ids]


def replace_row_by_id(rows: list[dict], next_row: dict) -> None:
    """Replace one row by Graph item ID, appending when it is new."""
    next_row_id = str(next_row.get("id", ""))
    for index, row in enumerate(rows):
        if str(row.get("id", "")) == next_row_id:
            rows[index] = next_row
            return
    rows.append(next_row)


def collect_delta_changes(
    drive_sources,
    headers,
) -> tuple[list[GraphDriveItem], set[str], dict[str, str], bool]:
    """Collect delta changes for all configured sources."""
    state = load_delta_state()
    has_any_saved_state = bool(state.get("sources"))
    pending_pptx_files: list[GraphDriveItem] = []
    removed_row_ids: set[str] = set()
    delta_links: dict[str, str] = {}

    for source in drive_sources:
        source_key = get_delta_source_key(source)
        source_state = state["sources"].get(source_key)
        result = collect_drive_delta(
            source,
            headers,
            source_state["delta_link"] if source_state else None,
        )
        delta_links[source_key] = result["delta_link"]

        source_items = [
            attach_configured_source(
                item,
                source["name"],
                source.get("folder", ""),
            )
            for item in result["items"]
        ]
        changed_items, deleted_ids = classify_delta_items(source_items)
        pending_pptx_files.extend(changed_items)
        removed_row_ids.update(deleted_ids)

        logger.info(
            "Collected %s delta items from %s%s.",
            len(result["items"]),
            source["name"],
            f"/{source['folder']}" if source.get("folder") else "",
        )

    return (
        dedupe_pptx_files(pending_pptx_files),
        removed_row_ids,
        delta_links,
        has_any_saved_state,
    )


def save_baseline_delta_state(drive_sources, headers) -> None:
    """Save Graph delta tokens from now without changing the workbook."""
    state = load_delta_state()
    delta_links: dict[str, str] = {}
    for source in drive_sources:
        result = collect_drive_delta(source, headers, token="latest")
        delta_links[get_delta_source_key(source)] = result["delta_link"]
        logger.info(
            "Saved baseline delta token for %s%s.",
            source["name"],
            f"/{source['folder']}" if source.get("folder") else "",
        )
    save_delta_state(merge_delta_links(state, drive_sources, delta_links))


def main() -> None:
    """Run the complete export workflow.

    The function is intentionally linear so it is easy to trace:
    - set up clients and configuration
    - load or create a checkpoint
    - process one PowerPoint at a time
    - write the current results to Excel as progress is made

    The workbook is rewritten after each file so the output on disk stays
    current even during a long run.
    """
    args = parse_args()

    openai_client = get_openai_client()

    setup = excel_setup()
    headers = setup["headers"]
    drive_sources = setup["drive_sources"]
    if not drive_sources:
        raise ValueError("configuration.DRIVE_SOURCES must contain at least one drive")

    # Backward-compatible alias for the old single-drive workflow. This is
    # mainly useful for the commented testing block lower in this file.
    library_drive_id = [s["drive_id"] for s in drive_sources if s.get("is_default")][0]
    library_drive_source = next(s for s in drive_sources if s.get("is_default"))

    generator_registry = GeneratorRegistry(
        default_drive_id=library_drive_id,
        headers=headers,
    )
    presentation_columns = get_presentation_columns(generator_registry)
    metadata_model = create_presentation_metadata_model(presentation_columns)
    excel_column_names = get_excel_column_names(presentation_columns)

    if args.restart_from_scratch and checkpoint_exists():
        clear_checkpoint()
        logger.info("Cleared saved checkpoint and restarting from scratch.")
    if args.restart_from_scratch and delta_state_exists():
        clear_delta_state()
        logger.info("Cleared saved Microsoft Graph delta state.")

    if args.baseline_from_now:
        save_baseline_delta_state(drive_sources, headers)
        logger.info("Saved baseline delta state without modifying the workbook.")
        return

    final_pptx_objects: list[dict] = []
    pending_pptx_files: list[GraphDriveItem]
    pending_delta_links: dict[str, str] = {}

    if checkpoint_exists():
        checkpoint = load_checkpoint()
        final_pptx_objects = checkpoint["processed_rows"]
        pending_pptx_files = checkpoint["pending_items"]
        pending_delta_links = checkpoint.get("delta_links", {})
        logger.info(
            "Resuming from checkpoint with %s processed rows and %s remaining files.",
            len(final_pptx_objects),
            len(pending_pptx_files),
        )
    elif args.legacy_full_scan:
        pending_pptx_files = []
        final_pptx_objects = []
        for source in drive_sources:
            pending_pptx_files.extend(
                get_all_pptx_files(
                    source["drive_id"],
                    headers,
                    source.get("folder_id", ""),
                    source["name"],
                    source.get("folder", ""),
                )
            )
        pending_pptx_files = dedupe_pptx_files(pending_pptx_files)
        save_checkpoint(final_pptx_objects, pending_pptx_files)
        logger.info(
            "Gathered %s presentation files for legacy Excel export.",
            len(pending_pptx_files),
        )
    else:
        (
            pending_pptx_files,
            removed_row_ids,
            pending_delta_links,
            has_any_saved_delta_state,
        ) = collect_delta_changes(drive_sources, headers)
        final_pptx_objects = (
            load_existing_workbook_rows() if has_any_saved_delta_state else []
        )
        final_pptx_objects = remove_rows_by_id(final_pptx_objects, removed_row_ids)
        # Example test-only shortcut: replace the full discovery result with a
        # hand-picked list of file IDs when you want to debug one deck quickly.
        # pending_pptx_files = [
        #     get_pptx_file(
        #         library_drive_id,
        #         item_id,
        #         headers,
        #         library_drive_source["name"],
        #         library_drive_source.get("folder", ""),
        #     )
        #     for item_id in [
        #         "01I7HKCO3RVKMEHQRDR5GZJS6QR56L6LCY",
        #         # "01I7HKCO4N6ZP6BHCCCVBJDSLM4WQMMQ3Q",
        #         # "01I7HKCO5IQU3OKDVXEJHYBUP7LAFU4UHH",
        #         # "01I7HKCO4SPWFCOVNN7JAL4QDULH5PPCYJ",
        #         # "01I7HKCO6FJRJTI7CXJRAISZFRGAYIPMPU",
        #         # "01I7HKCO7VKOUI5SISPVGITFBOSTIWTI3H",
        #     ]
        # ]
        save_checkpoint(
            final_pptx_objects,
            pending_pptx_files,
            pending_delta_links,
        )
        logger.info(
            "Gathered %s changed presentation files and %s removed rows for Excel export.",
            len(pending_pptx_files),
            len(removed_row_ids),
        )

    total_files = len(pending_pptx_files)
    processed_this_run = 0
    while pending_pptx_files:
        processed_this_run += 1
        pptx_file = pending_pptx_files[0]
        logger.info(
            "Processing %s/%s: %s (%s)",
            processed_this_run,
            total_files,
            pptx_file["name"],
            pptx_file["id"],
        )
        slide_texts, number_of_slides, average_words_per_slide = (
            get_ai_generation_inputs(pptx_file, generator_registry)
        )
        ai_metadata = generate_ai_metadata(
            openai_client,
            name=pptx_file["name"],
            presentation_path=get_configured_source_path(pptx_file),
            slide_texts=slide_texts,
            number_of_slides=number_of_slides,
            average_words_per_slide=average_words_per_slide,
            response_model=metadata_model,
        )

        replace_row_by_id(
            final_pptx_objects,
            build_presentation_row(pptx_file, presentation_columns, ai_metadata),
        )
        pending_pptx_files.pop(0)
        # Persist progress regularly so an interrupted run can resume instead of
        # starting over from the first file again.
        should_save_checkpoint = processed_this_run % 5 == 0 or not pending_pptx_files
        if should_save_checkpoint:
            save_checkpoint(
                final_pptx_objects,
                pending_pptx_files,
                pending_delta_links or None,
            )
        write_objects_to_excel(final_pptx_objects, headers=excel_column_names)

        if processed_this_run % 5 == 0 or processed_this_run == total_files:
            logger.info(
                "Processed %s/%s changed rows for Excel export.",
                processed_this_run,
                total_files,
            )

    write_objects_to_excel(final_pptx_objects, headers=excel_column_names)
    if pending_delta_links:
        save_delta_state(
            merge_delta_links(
                load_delta_state(),
                drive_sources,
                pending_delta_links,
            )
        )
    clear_checkpoint()


if __name__ == "__main__":
    main()
