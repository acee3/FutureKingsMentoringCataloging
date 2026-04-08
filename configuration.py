from typing import Any, Callable, Literal, NotRequired, TypedDict

from generators import GeneratorRegistry

EXCEL_CELL_CHARACTER_LIMIT = 32767
SLIDE_BREAK = "\n\n--- SLIDE BREAK ---\n\n"
DEFAULT_DATA_ROW_HEIGHT = 15
MINIMUM_COLUMN_WIDTH = 15
GENERATED_BY_AI_SUFFIX = "*"
OUTPUT_DIR = "output"


class ConfiguredDriveSource(TypedDict):
    name: str
    folder: NotRequired[str]


DRIVE_SOURCES: list[ConfiguredDriveSource] = [
    {"name": "Documents", "folder": "III. Partnerships"},
    {"name": "Documents", "folder": "I. Core King Content/II. Workshops"},
    {"name": "Workshops"},
]


class PresentationColumn(TypedDict):
    name: str
    generator: Callable[[dict[str, Any]], Any]


def get_presentation_columns(registry: GeneratorRegistry) -> list[PresentationColumn]:
    return [
        {"name": "id", "generator": registry.identity_generator("id")},
        {"name": "name", "generator": registry.identity_generator("name")},
        {"name": "web_url", "generator": registry.identity_generator("webUrl")},
        {"name": "slide_texts", "generator": registry.slide_texts_generator()},
        {
            "name": "last_modified",
            "generator": registry.identity_generator("lastModifiedDateTime"),
        },
        {
            "name": "number_of_slides",
            "generator": registry.number_of_slides_generator(),
        },
        {
            "name": "average_words_per_slide",
            "generator": registry.average_words_per_slide_generator(),
        },
        {
            "name": f"theme{GENERATED_BY_AI_SUFFIX}",
            "generator": registry.ai_generator(
                "theme",
                list[
                    Literal[
                        "Confidence & Leadership",
                        "Financial Literacy",
                        "College & Career Prep",
                    ]
                ],
                "One or more applicable themes. Only use multiple themes when necessary.",
            ),
        },
        {
            "name": f"description{GENERATED_BY_AI_SUFFIX}",
            "generator": registry.ai_generator(
                "description",
                str,
                "A concise description of the presentation in one or two sentences.",
            ),
        },
        {
            "name": f"duration_estimate_mins{GENERATED_BY_AI_SUFFIX}",
            "generator": registry.ai_generator(
                "duration_estimate_minutes",
                int,
                (
                    "Estimated total presentation duration in minutes, rounded to the "
                    "nearest 15 minutes unless over 120 minutes, in which case round "
                    "to the nearest hour."
                ),
            ),
        },
        {
            "name": f"audience{GENERATED_BY_AI_SUFFIX}",
            "generator": registry.ai_generator(
                "audience",
                Literal["Middle school", "High school", "College"],
                "The primary intended audience for the presentation.",
            ),
        },
        {
            "name": f"activity_length_mins{GENERATED_BY_AI_SUFFIX}",
            "generator": registry.ai_generator(
                "activity_length_minutes",
                int,
                (
                    "Approximate minutes of activity time in the presentation, using "
                    "the same rounding rules as duration_estimate_mins."
                ),
            ),
        },
    ]
