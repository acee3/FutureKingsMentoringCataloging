from typing import Literal

from app_types import ConfiguredDriveSource, PresentationColumn
from generators import GeneratorRegistry

EXCEL_CELL_CHARACTER_LIMIT = 32767
SLIDE_BREAK = "\n\n--- SLIDE BREAK ---\n\n"
DEFAULT_DATA_ROW_HEIGHT = 15
MINIMUM_COLUMN_WIDTH = 15
GENERATED_BY_AI_SUFFIX = "*"
OUTPUT_DIR = "output"
NON_PARTNERSHIP_TYPICAL_SLIDES = 20
NON_PARTNERSHIP_TYPICAL_WORDS_PER_SLIDE = 12.91


DRIVE_SOURCES: list[ConfiguredDriveSource] = [
    {"name": "Documents", "folder": "III. Partnerships"},
    {"name": "Documents", "folder": "I. Core King Content/II. Workshops"},
    {"name": "Workshops"},
]


def get_presentation_columns(registry: GeneratorRegistry) -> list[PresentationColumn]:
    return [
        {"name": "id", "generator": registry.identity_generator("id")},
        {"name": "name", "generator": registry.identity_generator("name")},
        {"name": "web_url", "generator": registry.identity_generator("webUrl")},
        {
            "name": "presentation_path",
            "generator": registry.presentation_path_generator(),
        },
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
                    "nearest multiple of 5. Evaluate the actual deck content first, "
                    "including size, density, pacing, and activity load. Use the "
                    "folder's historical norm as context for what a typical deck "
                    "usually runs: presentations in III. Partnerships are generally "
                    "around 120 minutes, while presentations in the other workshop "
                    "folders are generally around 90 minutes. Keep the estimate "
                    "reasonably close to that norm when the deck seems typical for "
                    "that folder, but move meaningfully shorter or longer when the "
                    "actual content is clearly much smaller or larger than typical. "
                    f"A typical non-partnership workshop in the current catalog is "
                    f"about {NON_PARTNERSHIP_TYPICAL_SLIDES} slides and about "
                    f"{NON_PARTNERSHIP_TYPICAL_WORDS_PER_SLIDE:.2f} average words per "
                    "slide, and decks around that size should usually stay close to "
                    "90 minutes unless their actual content suggests otherwise."
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
                    "Approximate minutes of activity time in the presentation, rounded "
                    "to the nearest multiple of 5 using the same content-sensitive "
                    "rounding approach as duration_estimate_mins."
                ),
            ),
        },
    ]
