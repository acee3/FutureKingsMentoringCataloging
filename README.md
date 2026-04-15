# fkm

Builds an Excel catalog of workshop PowerPoints with a mix of direct metadata and AI-generated fields, and now also includes a workbook-backed vector search app.

## What This Project Does

The script scans one or more Microsoft drives/folders for `.pptx` files, reads the text from each slide, asks OpenAI to classify or summarize each presentation, and writes the results to an Excel workbook.

The main output is:

- `output/workshop_catalog.xlsx`: the Excel file you care about
- `output/workshop_catalog_checkpoint.json`: a progress file used to resume long runs

The vector search app accepts an uploaded `.xlsx` workbook, turns each workbook row into searchable text, stores embeddings in Postgres with `pgvector`, and serves a spreadsheet-like UI plus semantic search.

## How The Code Is Organized

- [`excel_generation/main.py`](/Users/acheung/Desktop/fkm/excel_generation/main.py): the full workflow from start to finish
- [`excel_generation/configuration.py`](/Users/acheung/Desktop/fkm/excel_generation/configuration.py): the safest place to edit drives, folders, and Excel columns
- [`excel_generation/generators.py`](/Users/acheung/Desktop/fkm/excel_generation/generators.py): reusable helpers that know how to fill column values
- [`excel_generation/column_helpers.py`](/Users/acheung/Desktop/fkm/excel_generation/column_helpers.py): turns the configured columns into rows and an AI schema
- [`excel_generation/llm_work.py`](/Users/acheung/Desktop/fkm/excel_generation/llm_work.py): OpenAI client setup and AI metadata generation
- [`excel_generation/presentation_reader.py`](/Users/acheung/Desktop/fkm/excel_generation/presentation_reader.py): extracts text from PowerPoint files
- [`excel_generation/excel_maker.py`](/Users/acheung/Desktop/fkm/excel_generation/excel_maker.py): writes the Excel workbook
- [`excel_generation/checkpoint.py`](/Users/acheung/Desktop/fkm/excel_generation/checkpoint.py): saves and restores progress
- [`excel_generation/app_types.py`](/Users/acheung/Desktop/fkm/excel_generation/app_types.py): shared Excel-generation types
- [`excel_generation/microsoft/`](/Users/acheung/Desktop/fkm/excel_generation/microsoft): Microsoft Graph authentication, requests, and related types
- [`vector_search_app/`](/Users/acheung/Desktop/fkm/vector_search_app): workbook indexing, FastAPI service, `pgvector` storage, and static UI

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
uv run python excel_generation/main.py
```

If a previous run stopped in the middle and you want to continue, just run the same command again. It will use the checkpoint automatically.

If you want to ignore the checkpoint and start over:

```bash
uv run python excel_generation/main.py --restart-from-scratch
```

## Vector Search App

### How Data Gets Into The App

The vector search app does not scrape PowerPoints or talk to Microsoft Graph directly. It accepts an uploaded `.xlsx` workbook through the browser or CLI:

1. Upload a workbook in the UI, or pass one to the CLI.
2. Use the first row as column headers.
3. Convert each later row into:
   - row metadata as JSON
   - one combined searchable text string built from the most useful columns
4. Generate an embedding for that searchable text.
5. Upsert the row plus embedding into Postgres with `pgvector`.

That means the workbook remains the source of truth, while Postgres is only the search index.

### Local Commands

Index the workbook into Postgres:

```bash
uv run python -m vector_search_app.cli --workbook-path output/workshop_catalog.xlsx
```

Run the API locally:

```bash
uv run uvicorn vector_search_app.api:app --reload
```

### Docker Compose

Start Postgres and the vector search service:

```bash
docker compose up --build
```

Open [http://localhost:8000](http://localhost:8000).
Upload a workbook in the page and the app will index that file.

## How To Change Things

The safest approach for non-coders is to ask a coding agent to make changes
instead of manually editing Python files. GitHub Copilot in Visual Studio Code is
one option. Use the agent to edit the project, then rerun the Excel generator.

Good agent prompts are specific and mention the files below:

- "In `excel_generation/configuration.py`, add a new AI-generated Excel column
  called `audience_level*`. It should classify each presentation as Elementary,
  Middle School, High School, College, Parent, or General. Then explain how to
  rerun the workbook."
- "In `excel_generation/configuration.py`, add a new source folder under
  `DRIVE_SOURCES` for the Documents drive at `<folder path>`."
- "Change the way `duration_estimate_mins*` is calculated. Keep the Excel column
  name the same, but update the AI instructions to prefer shorter estimates for
  decks with fewer than 10 slides."

After an agent changes columns, run this so every row is rebuilt with the new
column values:

```bash
uv run python excel_generation/main.py --restart-from-scratch
```

### Add A Column

- Open [`excel_generation/configuration.py`](/Users/acheung/Desktop/fkm/excel_generation/configuration.py).
- Add a new entry in `get_presentation_columns(...)`.
- For a direct Microsoft field, use `registry.identity_generator(...)`.
- For nested/non-standard values, add a generator method in [`excel_generation/generators.py`](/Users/acheung/Desktop/fkm/excel_generation/generators.py) and use it in the column list.
- For an AI field, use `registry.ai_generator(field_name, output_type, description)`.
- Put all shared typed structures in [`excel_generation/app_types.py`](/Users/acheung/Desktop/fkm/excel_generation/app_types.py) if a new one is needed.

After an agent changes columns, run this so every row is rebuilt with the new
column values:

```bash
uv run python excel_generation/main.py --restart-from-scratch
```

For most new spreadsheet columns, ask the agent for an AI-generated column. That
usually only requires a change in `excel_generation/configuration.py`.

### Add A Folder To Check

- Open [`excel_generation/configuration.py`](/Users/acheung/Desktop/fkm/excel_generation/configuration.py).
- Add an entry to `DRIVE_SOURCES`.
- Use `{"name": "<drive name>"}` to scan an entire drive.
- Use `{"name": "<drive name>", "folder": "<folder path>"}` to scan one folder tree inside that drive.
- Run the export again; folder IDs are resolved from `DRIVE_SOURCES` at startup.

### Change How A Column Is Calculated

1. Open [`excel_generation/generators.py`](/Users/acheung/Desktop/fkm/excel_generation/generators.py).
2. Add a new method on `GeneratorRegistry` that returns a small `generate(...)` function.
3. Use that new generator inside [`excel_generation/configuration.py`](/Users/acheung/Desktop/fkm/excel_generation/configuration.py).

This pattern may look unusual if you are new to Python. The short version is: the registry methods build tiny helper functions so configuration stays simple.

## Beginner Editing Notes

- Start from [`excel_generation/configuration.py`](/Users/acheung/Desktop/fkm/excel_generation/configuration.py) unless you know you need deeper changes.
- Read [`excel_generation/main.py`](/Users/acheung/Desktop/fkm/excel_generation/main.py) top to bottom once before changing behavior. It gives you the full mental model.
- Shared Excel-generation types belong in [`excel_generation/app_types.py`](/Users/acheung/Desktop/fkm/excel_generation/app_types.py), not scattered across feature files.
- AI columns are marked with `*` in the column name because `GENERATED_BY_AI_SUFFIX = "*"` in configuration.
- The Excel file is rewritten often during a run. That is intentional so you can inspect progress.

## Troubleshooting

- If Microsoft authentication fails, check the Microsoft environment variables first.
- If OpenAI fails, check `OPENAI_API_KEY`.
- If the script stops partway through, rerun it and it should resume from the checkpoint.
- If you make a bad checkpoint and want a clean rerun, use `--restart-from-scratch`.
