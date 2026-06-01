# Output Spec

Bundle layout:

```text
reader_task/
  bundle_manifest.json
  source.docx
  document_ir.json
  source_map.json
  media_map.json
  jobs/
    translation/
      translation_job_001_input.json
      translation_job_001_prompt.md
      manifests/
        translation_job_001_exchange.json
      translation_job_001_output.raw.json
      translation_job_001_output.json
      translation_job_001_validation.json
      downloads/
    qa/
      qa_job_001_input.json
      qa_job_001_prompt.md
      manifests/
        qa_job_001_exchange.json
      qa_job_001_output.raw.json
      qa_job_001_output.json
      qa_job_001_validation.json
      downloads/
  logs/
    browser_exchange_log.jsonl
    pipeline_status.json
  assembled/
    draft_assembled.json
    fast_finish_report.json
    review_applied.json
    final_payload.json
  outputs/
    draft_text.md
    final_text.md
    final_text.docx
    final_payload.json
    validation_report.json
```

`*.raw.json` is the downloaded model output. `*_output.json` is the normalized accepted artifact used by local scripts.

The final delivery files are in `outputs/`.

If `finish-fast` is used, QA files are not created. `assembled/final_payload.json` is generated directly from `draft_assembled.json` with `status: "qa_skipped_fast"`, and `assembled/fast_finish_report.json` records that QA was deliberately skipped.

## IR Notes

`document_ir.json` is the local source of truth for deterministic merge and render steps.

- `items` contains all translatable units in document order.
- Tables are compact top-level `table` items. The model receives only translatable cells inside `items[].cells`; non-translatable cells remain only in `tables[].cells` and are not sent to ChatGPT.
- Paragraph-like items may include `heading_confidence`: `style`, `heuristic`, or `none`.
- `section_path` and job-level `section_context` provide the nearest heading context for translation and QA.
- `blocks` contains top-level paragraph-derived items only. Tables are represented separately in `tables` and as compact `table` entries in `items`.
- `tables[].cells` preserves physical table position and merged-cell canonical IDs; merged aliases are listed in `tables[].merged_aliases`.

## Chunking Notes

Translation and QA jobs include a `chunking` object.

- Translation jobs also include `translation_lane`: `body` or `auxiliary`.
- Body translation jobs use `max_items=100` and `max_chars=60000` by default.
- Auxiliary translation jobs use `max_items=500` and `max_chars=60000` by default.
- Auxiliary jobs contain tables plus references/appendix/supplementary/administrative sections. Body prose is packed separately so high-value narrative text does not share model context with fragmented table/reference material.
- Translation chunks use `chunker: "balanced_contiguous"` when lane content must be split. This balances item count and character count while preserving document order and keeping headings attached to their first paragraph.
- `contains_heading` is true when the job includes a heading item.
- `table_segments` lists table row ranges present in the job, for example `{"table_id":"T001","row_start":1,"row_end":8}`.
- Tables are counted as one top-level item for manual job splitting.
- Table `char_count` is computed from translatable cell text, not from a duplicated Markdown table.
- Large tables may occupy a whole translation job. The compact table item keeps `cell_id`, row/col coordinates, `translation_source_text`, `inline_tokens`, and `original_sha256` for each translatable cell.

## Rich Text Tokens

When the source DOCX contains non-translatable inline objects such as Office Math, drawings, or hyperlinks, the model input may contain protected tokens:

```text
[[RT:S001:001]]
```

The model must copy these tokens exactly into `translated_text`. During DOCX rendering, local Python replaces each token with a clone of the original inline OOXML. This preserves the original non-translatable rich element in the translated paragraph.

Translatable run formatting is not hidden behind tokens. Bold, italic, underline, strike, superscript, and subscript runs are exposed as HTML inline tags such as `<b>`, `<i>`, `<sup>`, and `<sub>` around the visible source text. ChatGPT should translate the text and move the same balanced tags onto the semantically corresponding Chinese text. DOCX rendering maps those tags back to Word run formatting. Color, highlight, all-caps, and small-caps styling is intentionally not marked for translation.

For table items, rich text tokens remain cell-scoped. The validator checks `translated_cells[].translated_text` against the corresponding cell's `inline_tokens`, and DOCX rendering clones the original cell inline XML by `cell_id`.

## LaTeX Text-Symbol Visualization

During final DOCX rendering, plain text snippets wrapped in dollar signs are lightly normalized with a built-in symbol map, for example `\pm` to `±`, `\le` to `≤`, and `^\circ C` to `°C`.

This is a final Word readability pass only. It does not change JSON translation payloads, does not replace native Office Math XML, and preserves the run's existing Word styling because it only updates `run.text`. The DOCX render report includes `latex_symbol_replacements`.
