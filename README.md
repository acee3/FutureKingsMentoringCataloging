# fkm

Builds an Excel catalog of workshop PowerPoints with a mix of direct metadata and AI-generated fields.

## What This Project Does

The script scans one or more Microsoft drives/folders for `.pptx` files, reads the text from each slide, asks OpenAI to classify or summarize each presentation, and writes the results to an Excel workbook.

The main output is:

- `output/workshop_catalog.xlsx`: the Excel file you care about
- `output/workshop_catalog_checkpoint.json`: a progress file used to resume long runs

## How The Code Is Organized

- [`main.py`](/Users/acheung/Desktop/fkm/main.py): the full workflow from start to finish
- [`configuration.py`](/Users/acheung/Desktop/fkm/configuration.py): the safest place to edit drives, folders, and Excel columns
- [`generators.py`](/Users/acheung/Desktop/fkm/generators.py): reusable helpers that know how to fill column values
- [`column_helpers.py`](/Users/acheung/Desktop/fkm/column_helpers.py): turns the configured columns into rows and an AI schema
- [`llm_work.py`](/Users/acheung/Desktop/fkm/llm_work.py): OpenAI client setup and AI metadata generation
- [`presentation_reader.py`](/Users/acheung/Desktop/fkm/presentation_reader.py): extracts text from PowerPoint files
- [`excel_maker.py`](/Users/acheung/Desktop/fkm/excel_maker.py): writes the Excel workbook
- [`checkpoint.py`](/Users/acheung/Desktop/fkm/checkpoint.py): saves and restores progress
- [`app_types.py`](/Users/acheung/Desktop/fkm/app_types.py): shared project types
- [`microsoft/`](/Users/acheung/Desktop/fkm/microsoft): Microsoft Graph authentication, requests, and related types

## How To Run It

1. Make sure the required environment variables are set:
   - `OPENAI_API_KEY`
   - `TENANT_ID`
   - `CLIENT_ID`
   - `CLIENT_SECRET_VALUE`
   - `SITE_HOSTNAME`
   - `SITE_PATH`
2. Run the program.

```bash
uv run python main.py
```

If a previous run stopped in the middle and you want to continue, just run the same command again. It will use the checkpoint automatically.

If you want to ignore the checkpoint and start over:

```bash
uv run python main.py --restart-from-scratch
```

## How To Change Things

### Add A Column

- Open [`configuration.py`](/Users/acheung/Desktop/fkm/configuration.py).
- Add a new entry in `get_presentation_columns(...)`.
- For a direct Microsoft field, use `registry.identity_generator(...)`.
- For nested/non-standard values, add a generator method in [`generators.py`](/Users/acheung/Desktop/fkm/generators.py) and use it in the column list.
- For an AI field, use `registry.ai_generator(field_name, output_type, description)`.
- Put all shared typed structures in [`app_types.py`](/Users/acheung/Desktop/fkm/app_types.py) if a new one is needed.

### Add A Folder To Check

- Open [`configuration.py`](/Users/acheung/Desktop/fkm/configuration.py).
- Add an entry to `DRIVE_SOURCES`.
- Use `{"name": "<drive name>"}` to scan an entire drive.
- Use `{"name": "<drive name>", "folder": "<folder path>"}` to scan one folder tree inside that drive.
- Run the export again; folder IDs are resolved from `DRIVE_SOURCES` at startup.

### Change How A Column Is Calculated

1. Open [`generators.py`](/Users/acheung/Desktop/fkm/generators.py).
2. Add a new method on `GeneratorRegistry` that returns a small `generate(...)` function.
3. Use that new generator inside [`configuration.py`](/Users/acheung/Desktop/fkm/configuration.py).

This pattern may look unusual if you are new to Python. The short version is: the registry methods build tiny helper functions so configuration stays simple.

## Beginner Editing Notes

- Start from [`configuration.py`](/Users/acheung/Desktop/fkm/configuration.py) unless you know you need deeper changes.
- Read [`main.py`](/Users/acheung/Desktop/fkm/main.py) top to bottom once before changing behavior. It gives you the full mental model.
- Shared types belong in [`app_types.py`](/Users/acheung/Desktop/fkm/app_types.py), not scattered across feature files.
- AI columns are marked with `*` in the column name because `GENERATED_BY_AI_SUFFIX = "*"` in configuration.
- The Excel file is rewritten often during a run. That is intentional so you can inspect progress.

## Troubleshooting

- If Microsoft authentication fails, check the Microsoft environment variables first.
- If OpenAI fails, check `OPENAI_API_KEY`.
- If the script stops partway through, rerun it and it should resume from the checkpoint.
- If you make a bad checkpoint and want a clean rerun, use `--restart-from-scratch`.
