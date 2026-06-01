# Recovery

## Source DOCX Failed Preflight

If `prepare-translation` or `preflight-source` reports `source DOCX failed quality preflight`, the source file likely came from physical PDF text extraction rather than structural document conversion.

Do this:

1. Return to the original PDF.
2. Run PaddleOCR MCP:

```text
mcp__paddleocr__ocr_extract(file_path="/path/to/source.pdf", output_format="docx", output_dir="/path/to/reader_task/source_mcp")
```

3. Use the generated DOCX as `--source-docx`.

Do not repair this by using `pdftotext`, copying text into Word, or forcing `--allow-layout-risk`, unless the user explicitly wants a rough text-only diagnostic draft.

`--allow-layout-risk` exists only for deliberate recovery/debugging runs. It is not acceptable for a normal final bilingual reader.

## Model JSON Issues

Common failures:

| Failure | Local detection | Recovery |
|---|---|---|
| Extra prose around JSON | JSON extraction fails | Run `make-repair-prompt` and ask ChatGPT to return JSON only |
| Missing item | Coverage check fails | Re-run that job or ask for a repair output covering missing items |
| Hash mismatch | `original_sha256` differs | Reject the output and re-run the original job |
| Extra item | Output contains unknown `item_id` | Ask ChatGPT to remove unknown items |
| Missing rich-text token | Token coverage check fails | Ask ChatGPT to return corrected JSON while preserving every `[[RT:...]]` token exactly |
| Wrong filename | `accept-downloads` cannot match | Rename into the stage `downloads/` folder or pass `--file` |
| QA patch too broad | `apply=true` without replacement | Keep finding as advisory; script will not apply it |

Do not repair translation semantics in Python. Repair by asking ChatGPT Web for corrected JSON, then re-run validation.
