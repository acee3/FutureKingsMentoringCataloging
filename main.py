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
from dotenv import load_dotenv
from configuration import get_presentation_columns
from excel_maker import write_objects_to_excel
from generators import GeneratorRegistry
from llm_work import generate_ai_metadata, get_openai_client
from microsoft import (
    excel_setup,
    get_all_pptx_files,
    get_pptx_file,
)
from microsoft.types import GraphDriveItem

load_dotenv()
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

logger = logging.getLogger(__name__)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--restart-from-scratch",
        action="store_true",
        help="Ignore any saved checkpoint and rebuild the export from the beginning.",
    )
    return parser.parse_args()


def dedupe_pptx_files(items: list[GraphDriveItem]) -> list[GraphDriveItem]:
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


def main():
    args = parse_args()

    openai_client = get_openai_client()

    setup = excel_setup()
    headers = setup["headers"]
    drive_sources = setup["drive_sources"]
    if not drive_sources:
        raise ValueError("configuration.DRIVE_SOURCES must contain at least one drive")

    generator_registry = GeneratorRegistry(
        default_drive_id=drive_sources[0]["drive_id"],
        headers=headers,
    )
    presentation_columns = get_presentation_columns(generator_registry)
    metadata_model = create_presentation_metadata_model(presentation_columns)
    excel_column_names = get_excel_column_names(presentation_columns)

    if args.restart_from_scratch and checkpoint_exists():
        clear_checkpoint()
        logger.info("Cleared saved checkpoint and restarting from scratch.")

    final_pptx_objects: list[dict] = []
    pending_pptx_files: list[GraphDriveItem]

    if checkpoint_exists():
        checkpoint = load_checkpoint()
        final_pptx_objects = checkpoint["processed_rows"]
        pending_pptx_files = checkpoint["pending_items"]
        logger.info(
            "Resuming from checkpoint with %s processed rows and %s remaining files.",
            len(final_pptx_objects),
            len(pending_pptx_files),
        )
    else:
        pending_pptx_files = []
        for source in drive_sources:
            pending_pptx_files.extend(
                get_all_pptx_files(
                    source["drive_id"],
                    headers,
                    source.get("folder_id", ""),
                )
            )
        pending_pptx_files = dedupe_pptx_files(pending_pptx_files)
        # Use for testing.
        # pending_pptx_files = [
        #     get_pptx_file(library_drive_id, item_id, headers)
        #     for item_id in [
        #         "01I7HKCO3RVKMEHQRDR5GZJS6QR56L6LCY",
        #         # "01I7HKCO4N6ZP6BHCCCVBJDSLM4WQMMQ3Q",
        #         # "01I7HKCO5IQU3OKDVXEJHYBUP7LAFU4UHH",
        #         # "01I7HKCO4SPWFCOVNN7JAL4QDULH5PPCYJ",
        #         # "01I7HKCO6FJRJTI7CXJRAISZFRGAYIPMPU",
        #         # "01I7HKCO7VKOUI5SISPVGITFBOSTIWTI3H",
        #     ]
        # ]
        save_checkpoint(final_pptx_objects, pending_pptx_files)
        logger.info(
            "Gathered %s presentation files for Excel export.",
            len(pending_pptx_files),
        )

    total_files = len(final_pptx_objects) + len(pending_pptx_files)
    while pending_pptx_files:
        index = len(final_pptx_objects) + 1
        pptx_file = pending_pptx_files[0]
        logger.info(
            "Processing %s/%s: %s (%s)",
            index,
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
            slide_texts=slide_texts,
            number_of_slides=number_of_slides,
            average_words_per_slide=average_words_per_slide,
            response_model=metadata_model,
        )

        final_pptx_objects.append(
            build_presentation_row(pptx_file, presentation_columns, ai_metadata)
        )
        pending_pptx_files.pop(0)
        should_save_checkpoint = index % 5 == 0 or not pending_pptx_files
        if should_save_checkpoint:
            save_checkpoint(final_pptx_objects, pending_pptx_files)
        write_objects_to_excel(final_pptx_objects, headers=excel_column_names)

        if index % 5 == 0 or index == total_files:
            logger.info("Processed %s/%s rows for Excel export.", index, total_files)

    write_objects_to_excel(final_pptx_objects, headers=excel_column_names)
    clear_checkpoint()


if __name__ == "__main__":
    main()
