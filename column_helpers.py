"""Helpers for turning configured columns into data rows and AI schemas."""

from typing import Any

from app_types import PresentationColumn
from generators import GeneratorRegistry
from pydantic import BaseModel, Field, create_model


def _is_ai_column(column: PresentationColumn) -> bool:
    """Return `True` when a column is backed by OpenAI output."""
    return bool(getattr(column["generator"], "is_ai", False))


def get_ai_columns(columns: list[PresentationColumn]) -> list[PresentationColumn]:
    """Filter the full column list down to AI-generated columns only."""
    return [column for column in columns if _is_ai_column(column)]


def get_excel_column_names(columns: list[PresentationColumn]) -> list[str]:
    """Return the Excel header row in the same order as the column config."""
    return [column["name"] for column in columns]


def create_presentation_metadata_model(
    columns: list[PresentationColumn],
) -> type[BaseModel]:
    """Create the Pydantic schema expected from OpenAI.

    Every AI column contributes one field to the model. This means the
    configuration in `configuration.py` automatically controls the structure of
    the AI response without needing a second schema definition elsewhere.
    """
    field_definitions = {
        column["generator"].field_name: (
            column["generator"].output_type,
            Field(description=column["generator"].description),
        )
        for column in get_ai_columns(columns)
    }
    return create_model("PresentationMetadata", **field_definitions)


def build_presentation_row(
    item: dict[str, Any],
    columns: list[PresentationColumn],
    ai_metadata: BaseModel | None = None,
) -> dict[str, Any]:
    """Build one final Excel row from raw item data plus optional AI metadata.

    Args:
        item: Raw Graph metadata for the PowerPoint file.
        columns: Ordered column definitions from configuration.
        ai_metadata: Parsed OpenAI response for AI columns.

    Returns:
        Dictionary keyed by Excel column name.
    """
    row: dict[str, Any] = {}

    for column in columns:
        if _is_ai_column(column):
            if ai_metadata is None:
                raise ValueError("AI metadata is required for AI-generated columns")
            row[column["name"]] = getattr(
                ai_metadata, column["generator"].field_name
            )
            continue
        row[column["name"]] = column["generator"](item)

    return row


def get_ai_generation_inputs(
    item: dict[str, Any],
    registry: GeneratorRegistry,
) -> tuple[list[str], int, float]:
    """Collect the exact inputs used in the OpenAI prompt.

    Keeping this calculation in one helper avoids subtle mismatches between the
    prompt and the exported non-AI columns.
    """
    drive_id = item.get("parentReference", {}).get("driveId", registry.default_drive_id)
    slide_texts = list(registry.get_slide_texts_for_item(drive_id, item["id"]))
    number_of_slides = len(slide_texts)
    average_words_per_slide = (
        sum(len(slide.split()) for slide in slide_texts) / number_of_slides
        if number_of_slides
        else 0.0
    )
    return slide_texts, number_of_slides, average_words_per_slide
