import os
from typing import Literal

from configuration import SLIDE_BREAK
from openai import OpenAI
from pydantic import BaseModel, Field


class PresentationMetadata(BaseModel):
    theme: list[
        Literal[
            "Confidence & Leadership",
            "Financial Literacy",
            "College & Career Prep",
        ]
    ] = Field(
        description=(
            "One or more applicable themes. Only use multiple themes when necessary."
        )
    )
    description: str = Field(
        description="A concise description of the presentation in one or two sentences."
    )
    duration_estimate_minutes: int = Field(
        description=(
            "Estimated total presentation duration in minutes, rounded to the nearest "
            "15 minutes unless over 120 minutes, in which case round to the nearest hour."
        )
    )
    audience: Literal["Middle school", "High school", "College"]
    activity_length_minutes: int = Field(
        description=(
            "Approximate minutes of activity time in the presentation, using the same "
            "rounding rules as duration_estimate_minutes."
        )
    )


def get_openai_client() -> OpenAI:
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise ValueError("OPENAI_API_KEY is missing from the environment")
    return OpenAI(api_key=api_key)


def generate_ai_metadata(
    client: OpenAI,
    *,
    name: str,
    slide_texts: list[str],
    number_of_slides: int,
    average_words_per_slide: float,
) -> PresentationMetadata:
    entire_slide_text = SLIDE_BREAK.join(slide_texts)

    response = client.responses.parse(
        model="gpt-4.1",
        input=[
            {
                "role": "system",
                "content": (
                    "You classify educational presentation decks. "
                    "Return only schema-valid structured data."
                ),
            },
            {
                "role": "user",
                "content": (
                    f"Presentation name: {name}\n"
                    f"Number of slides: {number_of_slides}\n"
                    f"Average words per slide: {average_words_per_slide:.2f}\n"
                    "Entire slide text:\n"
                    f"{entire_slide_text}"
                ),
            },
        ],
        text_format=PresentationMetadata,
    )

    if response.output_parsed is None:
        raise ValueError(f"OpenAI did not return structured metadata for {name}")

    return response.output_parsed
