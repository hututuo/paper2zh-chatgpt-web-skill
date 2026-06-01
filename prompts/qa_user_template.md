You will receive one JSON object named qa_job_input.json.

Compare original_text and translated_text.

Return exactly one JSON object conforming to qa_output.schema.json, and deliver it as a downloadable JSON file using the exact filename specified below.

Rules:
- Do not return Markdown.
- Do not wrap JSON in code fences.
- Create/download the response as a `.json` file. The file content must be the raw JSON object only.
- Use the exact expected output filename shown below; do not rename it and do not add suffixes such as `(1)`, `.txt`, or `.md`.
- Do not rewrite the whole document.
- Emit only structured findings.
- Review more broadly than the first translation pass: check omissions, mistranslations, terminology drift, table-cell values/labels, references/citations, units, math symbols, rich-text tokens, and Chinese readability.
- Set apply=true when suggested_translation can safely replace the exact item_id or exact table cell without needing full-document rewriting.
- For table-cell fixes, set `item_id` to the table item id and include the exact `cell_id`; `suggested_translation` replaces only that cell translation.
- If suggested_translation contains or edits an item with rich-text tokens such as `[[RT:S001:001]]`, preserve every token exactly.
- If suggested_translation contains or edits HTML inline formatting tags (`<b>`, `<i>`, `<u>`, `<s>`, `<sup>`, `<sub>`), preserve balanced tags and place them around the semantically corresponding Chinese text.
- Output only valid JSON.
