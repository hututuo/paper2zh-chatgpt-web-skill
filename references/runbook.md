# Runbook

## Source Stage

Create one task folder first. Keep every source conversion artifact, job file, download, log, draft, and final output inside this folder:

```bash
mkdir -p "/Users/huyiyang/AI agent/Codex/_tmp/YYYYMMDD-HHMM_chatgpt-web-json-reader-brief-name/reader_task"
```

Use that `reader_task` path as `--bundle`. Do not write temporary translation logs or generated job files into the larger project context.

For PDF input, create `source.docx` with PaddleOCR MCP before running local translation scripts:

```text
mcp__paddleocr__ocr_extract(file_path="/path/to/source.pdf", output_format="docx", output_dir="/path/to/reader_task/source_mcp")
```

Then pass the generated DOCX to `prepare-translation`.

Do not use `pdftotext`, manual text-layer extraction, or ad hoc `python-docx` reconstruction as a substitute for this source step. Those methods produce physical text blocks that can pass JSON validation while yielding unreadable bilingual output.

For an already-usable DOCX source, run a preflight check:

```bash
python /Users/huyiyang/.codex/skills/chatgpt-web-json-reader/scripts/readerctl.py preflight-source --bundle /path/to/reader_task --source-docx /path/to/source.docx
```

If preflight reports layout risk, regenerate the DOCX from the original PDF with PaddleOCR MCP.

## Translation Stage

1. Prepare the bundle:

```bash
python /Users/huyiyang/.codex/skills/chatgpt-web-json-reader/scripts/readerctl.py prepare-translation --bundle /path/to/reader_task --source-docx /path/to/source.docx
```

Default translation packing is lane-based:

- Body lane: `--max-items 100 --max-chars 60000`.
- Auxiliary lane: `--aux-max-items 500 --aux-max-chars 60000`.

Tables, references, appendix, supplementary, acknowledgements, funding, author contributions, conflicts, data availability, and ethics sections go into the auxiliary lane. Body prose is never mixed with those auxiliary fragments in the same translation job. Because tables are compact top-level items, a table with many translatable cells normally counts as one item and is limited by total cell text size instead of cell count.

When a lane exceeds its limits and needs more than one job, the chunker balances contiguous semantic groups by both paragraph/item count and character count. Headings remain attached to their first following paragraph, but the rest of a long section can be split evenly, avoiding lopsided body jobs such as 96 paragraphs followed by 31 paragraphs.

2. In the opened `jobs/translation/` folder, process each pair:

```text
translation_job_001_prompt.md
translation_job_001_input.json
```

Do not upload `manifests/*_exchange.json`. Those files are internal automation manifests for future browser/bridge tooling.

3. Save ChatGPT Web outputs into:

```text
jobs/translation/downloads/
```

Use the expected filename:

```text
translation_job_001_output.json
```

The prompt asks ChatGPT Web to create/download a raw JSON file with that exact filename. If the UI only prints JSON in the chat, save that raw JSON into the expected filename manually.

4. Continue locally:

```bash
python /Users/huyiyang/.codex/skills/chatgpt-web-json-reader/scripts/readerctl.py accept-downloads --bundle /path/to/reader_task --stage translation
python /Users/huyiyang/.codex/skills/chatgpt-web-json-reader/scripts/readerctl.py validate-model-outputs --bundle /path/to/reader_task --stage translation
python /Users/huyiyang/.codex/skills/chatgpt-web-json-reader/scripts/readerctl.py merge-draft --bundle /path/to/reader_task
```

## Fast Mode

Use this when the user wants a quick bilingual reader and explicitly does not need QA:

```bash
python /Users/huyiyang/.codex/skills/chatgpt-web-json-reader/scripts/readerctl.py finish-fast --bundle /path/to/reader_task --format md,docx,json
```

Run it after all translation output JSON files have been saved into `jobs/translation/downloads/`. It performs translation accept/validate, merges the draft, writes `assembled/final_payload.json` with `status: "qa_skipped_fast"`, renders final outputs, runs gate, and opens `outputs/`.

## QA Stage

1. Build QA jobs:

```bash
python /Users/huyiyang/.codex/skills/chatgpt-web-json-reader/scripts/readerctl.py make-qa-jobs --bundle /path/to/reader_task
```

`make-review-job` is accepted as an alias for `make-qa-jobs` when the task is described as a review pass.

By default this creates one aggregated QA input (`qa_job_001_input.json`) with a large review budget (`--max-items 200000 --max-chars 2000000`) so GPT can check more cross-document consistency. Pass smaller `--max-items` or `--max-chars` only when the user explicitly wants split QA jobs.

2. In the opened `jobs/qa/` folder, process each pair:

```text
qa_job_001_prompt.md
qa_job_001_input.json
```

3. Save ChatGPT Web outputs into:

```text
jobs/qa/downloads/
```

4. Continue locally:

```bash
python /Users/huyiyang/.codex/skills/chatgpt-web-json-reader/scripts/readerctl.py accept-downloads --bundle /path/to/reader_task --stage qa
python /Users/huyiyang/.codex/skills/chatgpt-web-json-reader/scripts/readerctl.py validate-model-outputs --bundle /path/to/reader_task --stage qa
python /Users/huyiyang/.codex/skills/chatgpt-web-json-reader/scripts/readerctl.py apply-qa --bundle /path/to/reader_task
python /Users/huyiyang/.codex/skills/chatgpt-web-json-reader/scripts/readerctl.py render-final --bundle /path/to/reader_task --format md,docx,json
python /Users/huyiyang/.codex/skills/chatgpt-web-json-reader/scripts/readerctl.py gate --bundle /path/to/reader_task --json
```

`apply-review` is accepted as an alias for `apply-qa`.

## Useful Checks

```bash
python /Users/huyiyang/.codex/skills/chatgpt-web-json-reader/scripts/readerctl.py status --bundle /path/to/reader_task --json
```

If a downloaded JSON is malformed, run:

```bash
python /Users/huyiyang/.codex/skills/chatgpt-web-json-reader/scripts/readerctl.py make-repair-prompt --bundle /path/to/reader_task --stage translation --job-id translation_job_001
```
