from dotenv import load_dotenv
from configuration import GENERATED_BY_AI_SUFFIX
from excel_maker import write_objects_to_excel
from llm_work import generate_ai_metadata, get_openai_client
from microsoft import (
    download_pptx_file_content,
    excel_setup,
    get_file,
)
from presentation_reader import extract_slide_text_from_pptx_bytes

load_dotenv()


def main():
    openai_client = get_openai_client()

    headers, library_drive_id = excel_setup()

    # raw_pptx_files = get_all_pptx_files(library_drive_id, headers)
    raw_pptx_files = [
        get_file(library_drive_id, item_id, headers)
        for item_id in [
            "01I7HKCO3RVKMEHQRDR5GZJS6QR56L6LCY",
            # "01I7HKCO4N6ZP6BHCCCVBJDSLM4WQMMQ3Q",
            # "01I7HKCO5IQU3OKDVXEJHYBUP7LAFU4UHH",
            # "01I7HKCO4SPWFCOVNN7JAL4QDULH5PPCYJ",
            # "01I7HKCO6FJRJTI7CXJRAISZFRGAYIPMPU",
            # "01I7HKCO7VKOUI5SISPVGITFBOSTIWTI3H",
        ]
    ]

    base_pptx_slide_data = [
        {
            "id": pptx_file["id"],
            "name": pptx_file["name"],
            "web_url": pptx_file["webUrl"],
            "slide_texts": extract_slide_text_from_pptx_bytes(
                download_pptx_file_content(library_drive_id, pptx_file["id"], headers)
            ),
            "last_modified": pptx_file["lastModifiedDateTime"],
        }
        for pptx_file in raw_pptx_files
    ]

    final_pptx_objects = []
    for pptx_file in base_pptx_slide_data:
        number_of_slides = len(pptx_file["slide_texts"])
        average_words_per_slide = (
            sum(len(slide.split()) for slide in pptx_file["slide_texts"])
            / number_of_slides
        )
        ai_metadata = generate_ai_metadata(
            openai_client,
            name=pptx_file["name"],
            slide_texts=pptx_file["slide_texts"],
            number_of_slides=number_of_slides,
            average_words_per_slide=average_words_per_slide,
        )

        final_pptx_objects.append(
            {
                **pptx_file,
                "number_of_slides": number_of_slides,
                "average_words_per_slide": average_words_per_slide,
                f"theme{GENERATED_BY_AI_SUFFIX}": ai_metadata.theme,
                f"description{GENERATED_BY_AI_SUFFIX}": ai_metadata.description,
                f"duration_estimate_mins{GENERATED_BY_AI_SUFFIX}": ai_metadata.duration_estimate_minutes,
                f"audience{GENERATED_BY_AI_SUFFIX}": ai_metadata.audience,
                f"activity_length_mins{GENERATED_BY_AI_SUFFIX}": ai_metadata.activity_length_minutes,
            }
        )

    write_objects_to_excel(final_pptx_objects)


if __name__ == "__main__":
    main()
