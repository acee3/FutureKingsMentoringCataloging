"""Factories for column generator functions.

The export configuration does not call Microsoft Graph or OpenAI directly.
Instead, it asks `GeneratorRegistry` for small generator functions. This keeps
the column configuration readable while centralizing the implementation details.
"""

from functools import cache
from typing import Any

from microsoft.graph import download_pptx_file_content
from microsoft.types import GraphDriveItem, GraphHeaders
from presentation_reader import extract_slide_text_from_pptx_bytes


class GeneratorRegistry:
    """Create reusable generator callables for presentation columns.

    Args:
        default_drive_id: Fallback drive ID used when an item does not include
            one in its parent reference.
        headers: Auth headers for Microsoft Graph requests.
    """

    def __init__(self, *, default_drive_id: str, headers: GraphHeaders):
        self.default_drive_id = default_drive_id
        self.headers = headers

    @cache
    def get_slide_texts_for_item(self, drive_id: str, item_id: str) -> tuple[str, ...]:
        """Download and parse slide text for one PowerPoint item.

        The `@cache` decorator means repeated requests for the same file reuse
        the first result instead of downloading and parsing the file again.
        """
        pptx_bytes = download_pptx_file_content(drive_id, item_id, self.headers)
        return tuple(extract_slide_text_from_pptx_bytes(pptx_bytes))

    def _get_drive_id_for_item(self, item: GraphDriveItem) -> str:
        """Return the drive ID associated with a Graph item."""
        return item.get("parentReference", {}).get("driveId", self.default_drive_id)

    def identity_generator(self, source_key: str):
        """Build a generator that copies one field directly from Graph metadata."""

        def generate(item: GraphDriveItem) -> Any:
            return item[source_key]

        generate.is_ai = False
        return generate

    def presentation_path_generator(self):
        """Build a generator that returns the SharePoint/OneDrive folder path."""

        def generate(item: GraphDriveItem) -> str:
            return item.get("parentReference", {}).get("path", "")

        generate.is_ai = False
        return generate

    def slide_texts_generator(self):
        """Build a generator that returns slide text for the presentation."""

        def generate(item: GraphDriveItem) -> list[str]:
            drive_id = self._get_drive_id_for_item(item)
            return list(self.get_slide_texts_for_item(drive_id, item["id"]))

        generate.is_ai = False
        return generate

    def number_of_slides_generator(self):
        """Build a generator that counts slides using extracted slide text."""
        slide_texts = self.slide_texts_generator()

        def generate(item: GraphDriveItem) -> int:
            return len(slide_texts(item))

        generate.is_ai = False
        return generate

    def average_words_per_slide_generator(self):
        """Build a generator that measures how text-dense the deck is."""
        slide_texts = self.slide_texts_generator()

        def generate(item: GraphDriveItem) -> float:
            slides = slide_texts(item)
            if not slides:
                return 0.0
            return sum(len(slide.split()) for slide in slides) / len(slides)

        generate.is_ai = False
        return generate

    def ai_generator(self, field_name: str, output_type: Any, description: str):
        """Build a placeholder generator for an AI-produced column.

        AI columns are not generated directly by this function. Instead, the
        returned callable acts as a container for metadata:
        - `field_name`: internal structured output field name
        - `output_type`: expected type in the OpenAI response
        - `description`: prompt guidance for that field

        The actual values are filled later in `build_presentation_row(...)`
        after OpenAI returns a validated Pydantic model.
        """

        def generate(item: GraphDriveItem) -> None:
            return None

        generate.is_ai = True
        generate.field_name = field_name
        generate.output_type = output_type
        generate.description = description
        return generate
