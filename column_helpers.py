from typing import Any

from configuration import PresentationColumn
from generators import GeneratorRegistry
from pydantic import BaseModel, Field, create_model


def _is_ai_column(column: PresentationColumn) -> bool:
    return bool(getattr(column["generator"], "is_ai", False))


def get_ai_columns(columns: list[PresentationColumn]) -> list[PresentationColumn]:
    return [column for column in columns if _is_ai_column(column)]


def get_excel_column_names(columns: list[PresentationColumn]) -> list[str]:
    return [column["name"] for column in columns]


def create_presentation_metadata_model(
    columns: list[PresentationColumn],
) -> type[BaseModel]:
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
    drive_id = item.get("parentReference", {}).get("driveId", registry.default_drive_id)
    slide_texts = list(registry.get_slide_texts_for_item(drive_id, item["id"]))
    number_of_slides = len(slide_texts)
    average_words_per_slide = (
        sum(len(slide.split()) for slide in slide_texts) / number_of_slides
        if number_of_slides
        else 0.0
    )
    return slide_texts, number_of_slides, average_words_per_slide
