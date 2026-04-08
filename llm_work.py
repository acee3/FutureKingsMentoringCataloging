import os

from configuration import SLIDE_BREAK
from openai import OpenAI
from pydantic import BaseModel


def get_openai_client() -> OpenAI:
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
