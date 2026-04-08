# fkm

Builds an Excel catalog of workshop PowerPoints with a mix of direct metadata and AI-generated fields.

## Add A Column

- Open [`configuration.py`](/Users/acheung/Desktop/fkm/configuration.py).
- Add a new entry in `get_presentation_columns(...)`.
- For a direct Microsoft field, use `registry.identity_generator(...)`.
- For nested/non-standard values, add a generator method in [`generators.py`](/Users/acheung/Desktop/fkm/generators.py) and use it in the column list.
- For an AI field, use `registry.ai_generator(field_name, output_type, description)`.
- Put all shared typed structures in [`app_types.py`](/Users/acheung/Desktop/fkm/app_types.py) if a new one is needed.

## Add A Folder To Check

- Open [`configuration.py`](/Users/acheung/Desktop/fkm/configuration.py).
- Add an entry to `DRIVE_SOURCES`.
- Use `{"name": "<drive name>"}` to scan an entire drive.
- Use `{"name": "<drive name>", "folder": "<folder path>"}` to scan one folder tree inside that drive.
- Run the export again; folder IDs are resolved from `DRIVE_SOURCES` at startup.
