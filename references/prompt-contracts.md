# Prompt Contracts

All model jobs must be self-contained.

Translation outputs must:

- Return one JSON object only.
- Prefer raw JSON with no Markdown code fence. The local parser tolerates fenced JSON defensively, but prompts should still ask for plain JSON to reduce manual cleanup.
- Deliver the result as a `.json` file with the exact expected filename, such as `translation_job_001_output.json`.
- Preserve `schema_version`, `job_id`, `document_id`, `item_id`, `item_type`, and `original_sha256`.
- Use `translation_source_text` when present. Preserve formula symbols, variables, units, citations, inline math, every `[[RT:...]]` rich-text token, and balanced HTML inline formatting tags (`<b>`, `<i>`, `<u>`, `<s>`, `<sup>`, `<sub>`) in `translated_text`.
- Include every input item exactly once.
- Never merge, split, omit, or reorder items.
- Put the translated text in `translated_text`.
- Use structured errors instead of dropping items.

QA outputs must:

- Return one JSON object only.
- Prefer raw JSON with no Markdown code fence. Fenced JSON may be accepted locally as a recovery convenience, not as the desired output format.
- Deliver the result as a `.json` file with the exact expected filename, such as `qa_job_001_output.json`.
- Preserve `job_id` and `document_id`.
- Emit structured `findings`.
- Review broadly: omissions, mistranslations, terminology drift, table-cell values/labels, references/citations, units, math symbols, rich-text tokens, HTML inline formatting tags, and Chinese readability.
- Set `apply: true` only when `suggested_translation` is a direct replacement for one exact `item_id` or one exact `item_id + cell_id`.
- Avoid rewriting the whole document.
