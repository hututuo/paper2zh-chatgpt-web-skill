You will receive one JSON object named input.json.

Return exactly one JSON object conforming to translation_output.schema.json, and deliver it as a downloadable JSON file using the exact filename specified below.

Rules:
- Do not return Markdown.
- Do not wrap JSON in code fences.
- Create/download the response as a `.json` file. The file content must be the raw JSON object only.
- Use the exact expected output filename shown below; do not rename it and do not add suffixes such as `(1)`, `.txt`, or `.md`.
- Preserve schema_version, job_id, document_id, item_id, item_type, and original_sha256 exactly.
- Translate every item faithfully into zh-CN. If `translation_source_text` is present, translate that field rather than plain `original_text`.
- Preserve formula symbols, variables, units, citations, and inline math as faithfully as possible in translated_text.
- Preserve every rich-text token exactly, such as `[[RT:S001:001]]`. Do not translate, delete, split, or modify these tokens.
- Preserve HTML inline formatting tags (`<b>`, `<i>`, `<u>`, `<s>`, `<sup>`, `<sub>` and closing tags). Apply each tag to the semantically corresponding translated words, not to the original English. Do not drop, rename, escape, or invent these tags.
- Do not summarize, omit, merge, split, or reorder items.
- If an item cannot be translated, keep the item and set translated_text to null with a structured error.
- For `item_type: "table"`, translate only the provided `cells[]` entries. Return the same table item once with `translated_cells[]`; each translated cell must preserve `cell_id` and `original_sha256` exactly and put the translation in `translated_text`.
- Do not invent translations for skipped/non-provided table cells.
- Output only valid JSON.
