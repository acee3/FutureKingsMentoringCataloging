import logging

from dotenv import load_dotenv
from column_helpers import (
    build_presentation_row,
    create_presentation_metadata_model,
    get_excel_column_names,
    get_ai_generation_inputs,
)
from configuration import get_presentation_columns
from excel_maker import write_objects_to_excel
from generators import GeneratorRegistry
from llm_work import generate_ai_metadata, get_openai_client
from microsoft import (
    excel_setup,
    get_all_pptx_files,
    get_pptx_file,
)

load_dotenv()
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

logger = logging.getLogger(__name__)


def main():
    openai_client = get_openai_client()

    headers, library_drive_id = excel_setup()
    generator_registry = GeneratorRegistry(drive_id=library_drive_id, headers=headers)
    presentation_columns = get_presentation_columns(generator_registry)
    metadata_model = create_presentation_metadata_model(presentation_columns)
    excel_column_names = get_excel_column_names(presentation_columns)

    raw_pptx_files = get_all_pptx_files(library_drive_id, headers)
    logger.info(
        "Gathered %s presentation files for Excel export.",
        len(raw_pptx_files),
    )
    # Use for testing.
    # raw_pptx_files = [
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

    final_pptx_objects = []
    total_files = len(raw_pptx_files)
    for index, pptx_file in enumerate(raw_pptx_files, start=1):
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

        if index % 5 == 0 or index == total_files:
            logger.info("Processed %s/%s rows for Excel export.", index, total_files)

    write_objects_to_excel(final_pptx_objects, headers=excel_column_names)


if __name__ == "__main__":
    main()
