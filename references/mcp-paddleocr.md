# PaddleOCR MCP Boundary

This skill uses the bundled PaddleOCR MCP server as the recommended PDF-to-DOCX source conversion step.

The MCP server source is included under `mcp/paddleocr/`. The local agent environment must configure and expose a callable tool equivalent to:

```text
mcp__paddleocr__ocr_extract(
  file_path="/path/to/source.pdf",
  output_format="docx",
  output_dir="/path/to/reader_task/source_mcp"
)
```

## Why MCP Is Required For PDF

`readerctl.py` works from a structurally usable DOCX. PDF text extraction tools such as `pdftotext` often produce physical line blocks rather than document structure. That can make JSON validation pass while producing a poor bilingual reader.

For PDF input, the expected source path is:

```text
PDF paper
  -> PaddleOCR MCP with output_format="docx"
  -> source.docx
  -> readerctl.py prepare-translation
```

For already usable DOCX input, PaddleOCR MCP is not required.

## Codex Configuration Shape

Exact MCP installation paths depend on the user's machine. In Codex, point the MCP server entry at the bundled server script, for example in `~/.codex/config.toml`:

```toml
[mcp_servers.paddleocr]
type = "stdio"
command = "/absolute/path/to/paper2zh-chatgpt-web-skill/.venv-paddleocr-mcp/bin/python"
args = ["/absolute/path/to/paper2zh-chatgpt-web-skill/mcp/paddleocr/server.py"]
```

After the MCP server is available, the agent should call the MCP tool first and pass the generated DOCX to this skill's local CLI.

The access token must stay outside committed source. Use one of:

- `PADDLEOCR_ACCESS_TOKEN`
- `PADDLEOCR_MCP_AISTUDIO_ACCESS_TOKEN`
- `mcp/paddleocr/.env`

## Failure Policy

If the PDF-to-DOCX MCP step is unavailable or fails, stop and report the source conversion failure. Do not silently fall back to `pdftotext`, manual copy-paste, or ad hoc DOCX reconstruction unless the user explicitly asks for a rough text-only diagnostic draft.
