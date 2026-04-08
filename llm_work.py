"""OpenAI-specific helpers.

This module is responsible for two things only:
- creating the OpenAI client from environment variables
- requesting structured metadata for one presentation at a time
"""

import os

from configuration import SLIDE_BREAK
from openai import OpenAI
from pydantic import BaseModel


def get_openai_client() -> OpenAI:
    """Create an authenticated OpenAI client.

    Returns:
        OpenAI: Client object used for API requests.

    Raises:
        ValueError: If `OPENAI_API_KEY` is missing.
    """
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise ValueError("OPENAI_API_KEY is missing from the environment")
    return OpenAI(api_key=api_key)


def generate_ai_metadata(
    client: OpenAI,
    *,
    name: str,
    presentation_path: str,
    slide_texts: list[str],
    number_of_slides: int,
    average_words_per_slide: float,
    response_model: type[BaseModel],
) -> BaseModel:
    """Ask OpenAI to produce structured metadata for one presentation.

    The caller passes a Pydantic model type in `response_model`. That model is
    built from the configured AI columns, so OpenAI is forced to return data in
    the exact shape the Excel export expects.

    Args:
        client: Authenticated OpenAI client.
        name: Presentation file name.
        presentation_path: Folder path in SharePoint/OneDrive.
        slide_texts: Plain text extracted from every slide.
        number_of_slides: Count of slides in the deck.
        average_words_per_slide: Simple density signal used in the prompt.
        response_model: Pydantic model describing the required AI output.

    Returns:
        A populated Pydantic model instance containing AI-generated fields.

    Raises:
        ValueError: If the API response does not contain parsed structured data.
    """
    entire_slide_text = SLIDE_BREAK.join(slide_texts)

    response = client.responses.parse(
        model="gpt-4.1",
        input=[
            {
                "role": "system",
                "content": (
                    "You classify educational presentation decks. "
                    "Return only schema-valid structured data. "
                    "Do not invent details that are not supported by the deck."
                ),
            },
            {
                "role": "user",
                "content": (
                    f"Presentation name: {name}\n"
                    f"Presentation path: {presentation_path or 'Unknown'}\n"
                    f"Number of slides: {number_of_slides}\n"
                    f"Average words per slide: {average_words_per_slide:.2f}\n"
                    "Entire slide text:\n"
                    f"{entire_slide_text}"
                ),
            },
        ],
        text_format=response_model,
    )

    if response.output_parsed is None:
        raise ValueError(f"OpenAI did not return structured metadata for {name}")

    return response.output_parsed
