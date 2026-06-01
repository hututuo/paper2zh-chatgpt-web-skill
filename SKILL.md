---
name: chatgpt-web-json-reader
description: Build bilingual DOCX/Markdown readers through a local JSON pipeline and manual ChatGPT Web upload/download. Use when the user wants to translate or QA a DOCX/PDF-derived Word document with structured JSON inputs, human-operated ChatGPT Web exchange, deterministic validation, merge, QA, and final render steps.
metadata:
  short-description: DOCX reader via manual ChatGPT Web JSON jobs
---

# ChatGPT Web JSON Reader

Use this skill when a document should be translated or QA'd through ChatGPT Web, while Codex keeps all deterministic work local.

Core boundary:

- Python scripts create `source.docx`, `document_ir.json`, model input JSON, prompts, exchange manifests, validation reports, merged drafts, and final outputs.
- At the start of a translation request, create one dedicated temporary task folder, then put every source conversion artifact, job file, download, log, draft, and output inside that folder. Do not scatter logs or intermediate files into the larger project context.
- ChatGPT Web only returns strict JSON outputs.
- Upload and download may be manual. Do not spend tokens operating the browser unless the user explicitly asks.
- Do not generate translation text in Python. Python may only validate, merge, apply QA patches, and render files.
- PDF inputs must first be converted to a structurally usable DOCX through MCP. When `mcp__paddleocr__ocr_extract` is available, call it with `output_format: "docx"` before any `readerctl` translation step.
- This repository includes the PaddleOCR MCP server source under `mcp/paddleocr/`, but the AI Studio access token must be configured locally through environment variables or a private `.env` file. `readerctl.py` consumes the generated DOCX; it does not start the MCP server itself.
- Do not substitute `pdftotext`, ad hoc Python DOCX generation, or physical-layout PDF-to-Word extraction for the MCP conversion step unless the user explicitly asks for a rough text-only draft.
- `prepare-translation` runs a source DOCX quality preflight before creating model jobs. If it reports layout risk, regenerate the DOCX via MCP/OCR rather than continuing.
- Rich text uses two mechanisms. Non-translatable inline objects such as Office Math, drawings, and hyperlinks are protected as `[[RT:...]]` tokens and cloned back from the original DOCX. Translatable run formatting such as bold, italic, underline, strike, superscript, and subscript is exposed to ChatGPT as HTML tags (`<b>`, `<i>`, `<u>`, `<s>`, `<sup>`, `<sub>`) so the text remains visible and can be translated.
- Tables are compact top-level items. ChatGPT receives only translatable cells inside a table item and returns `translated_cells[]`; skipped numeric/symbol cells are never sent to the model.
- Translation jobs use two lanes: body text is split separately at `100 items / 60000 chars`; tables plus references/appendix/supplementary/administrative sections are packed in an auxiliary lane at `500 items / 60000 chars`. Do not mix body prose with table/reference fragments in the same translation job. When a lane must be split into multiple jobs, use balanced contiguous chunks so paragraph count and character count are both reasonably even.
- Final DOCX rendering includes a lightweight LaTeX text-symbol visualization pass for common inline snippets such as `\pm`, `\le`, `\alpha`, and `^\circ C`; it runs after translation merge and never alters native Office Math XML.

## Default Workflow

If the user provides a PDF, first run PaddleOCR MCP:

```text
mcp__paddleocr__ocr_extract(file_path="/path/to/source.pdf", output_format="docx", output_dir="/path/to/reader_task/source_mcp")
```

Use the generated `.docx` as `--source-docx`. If MCP is unavailable or fails, stop and report that the PDF-to-DOCX source step is blocked; do not fall back to `pdftotext`.

Create a dedicated task folder, for example:

```bash
mkdir -p "/Users/huyiyang/AI agent/Codex/_tmp/YYYYMMDD-HHMM_chatgpt-web-json-reader-brief-name/reader_task"
```

Use that `reader_task` directory as `--bundle`; keep all logs and temporary artifacts under it.

Then run the bundled CLI:

```bash
python /Users/huyiyang/.codex/skills/chatgpt-web-json-reader/scripts/readerctl.py prepare-translation \
  --bundle /path/to/reader_task \
  --source-docx /path/to/source.docx
```

If this fails with `source DOCX failed quality preflight`, read `logs/source_quality_report.json`, regenerate the source DOCX from the original PDF with the MCP/OCR path, then rerun the command. Use `--allow-layout-risk` only when the user explicitly wants a physical-layout draft.

This creates translation `*_input.json` and `*_prompt.md` files, then opens:

- `jobs/translation/` for files to upload or paste into ChatGPT Web
- `jobs/translation/downloads/` for files the user should save/download back into

`*_exchange.json` files are internal automation manifests only. They are kept under `jobs/<stage>/manifests/` and should not be uploaded or pasted during manual exchange.

After the user downloads ChatGPT outputs:

```bash
python /Users/huyiyang/.codex/skills/chatgpt-web-json-reader/scripts/readerctl.py accept-downloads --bundle /path/to/reader_task --stage translation
python /Users/huyiyang/.codex/skills/chatgpt-web-json-reader/scripts/readerctl.py validate-model-outputs --bundle /path/to/reader_task --stage translation
python /Users/huyiyang/.codex/skills/chatgpt-web-json-reader/scripts/readerctl.py merge-draft --bundle /path/to/reader_task
python /Users/huyiyang/.codex/skills/chatgpt-web-json-reader/scripts/readerctl.py make-qa-jobs --bundle /path/to/reader_task
```

For a quick reader without QA, run this instead after saving translation downloads:

```bash
python /Users/huyiyang/.codex/skills/chatgpt-web-json-reader/scripts/readerctl.py finish-fast --bundle /path/to/reader_task --format md,docx,json
```

`finish-fast` accepts and validates translation downloads, merges the draft, marks QA as skipped, renders final files, runs gate, and opens `outputs/`.

`make-qa-jobs` opens the QA upload and download folders. After the user downloads QA outputs:

By default QA is a single aggregated job (`qa_job_001_input.json`) so the model can review cross-document consistency. Split QA only when the user explicitly asks for smaller QA jobs.

```bash
python /Users/huyiyang/.codex/skills/chatgpt-web-json-reader/scripts/readerctl.py accept-downloads --bundle /path/to/reader_task --stage qa
python /Users/huyiyang/.codex/skills/chatgpt-web-json-reader/scripts/readerctl.py validate-model-outputs --bundle /path/to/reader_task --stage qa
python /Users/huyiyang/.codex/skills/chatgpt-web-json-reader/scripts/readerctl.py apply-qa --bundle /path/to/reader_task
python /Users/huyiyang/.codex/skills/chatgpt-web-json-reader/scripts/readerctl.py render-final --bundle /path/to/reader_task --format md,docx,json
python /Users/huyiyang/.codex/skills/chatgpt-web-json-reader/scripts/readerctl.py gate --bundle /path/to/reader_task --json
```

## Manual Exchange Rules

For each job, give ChatGPT Web the matching prompt and input JSON:

- `translation_job_001_prompt.md`
- `translation_job_001_input.json`

Ask ChatGPT to create/download the exact expected output filename, for example:

- `translation_job_001_output.json`
- `qa_job_001_output.json`

Save downloaded JSON files into the opened `downloads/` folder for that stage. Then run `accept-downloads`.

The prompt files explicitly require ChatGPT Web to produce a downloadable raw JSON file with the exact filename. If ChatGPT shows JSON in chat instead of downloading it, save the raw JSON text into that exact filename under the stage `downloads/` folder.

If validation reports `missing rich-text token`, send the repair prompt to ChatGPT and require it to keep every `[[RT:...]]` token unchanged.

## References

- Read `references/runbook.md` for the full command sequence and pause/resume points.
- Read `references/output-spec.md` for output file meanings.
- Read `references/prompt-contracts.md` before editing prompts.
- Read `references/recovery.md` when validation fails.
