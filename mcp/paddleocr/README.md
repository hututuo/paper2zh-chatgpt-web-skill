# Bundled PaddleOCR MCP Server

This directory contains the PaddleOCR MCP server used by Paper2ZH for the PDF-to-DOCX source conversion step.

The server calls the PaddleOCR AI Studio hosted API and exposes one MCP tool:

```text
mcp__paddleocr__ocr_extract(file_path, output_format, output_dir)
```

Use `output_format="docx"` before running the Paper2ZH translation workflow on a PDF.

## Get An Access Token

1. Open the AI Studio Access Token page:
   <https://aistudio.baidu.com/index/accessToken>
2. Copy your access token.
3. Check the current free quota shown by your AI Studio/PaddleOCR account page.
4. Put it in a private `.env` file or inject it through the MCP host environment.

The official PaddleOCR docs also describe the same `PADDLEOCR_ACCESS_TOKEN` environment-variable convention:

- <https://www.paddleocr.ai/latest/version3.x/inference_deployment/serving/paddleocr_official_api/cli.html>
- <https://www.paddleocr.ai/latest/version3.x/integrations/mcp_server.html>

## Local Setup

From the repository root:

```bash
python3 -m venv .venv-paddleocr-mcp
. .venv-paddleocr-mcp/bin/activate
pip install -r mcp/paddleocr/requirements.txt
cp mcp/paddleocr/.env.example mcp/paddleocr/.env
chmod 600 mcp/paddleocr/.env
```

Then edit `mcp/paddleocr/.env`:

```text
PADDLEOCR_ACCESS_TOKEN=your-access-token-here
```

## Codex MCP Config Example

Use absolute paths in your real config:

```toml
[mcp_servers.paddleocr]
type = "stdio"
command = "/absolute/path/to/paper2zh-chatgpt-web-skill/.venv-paddleocr-mcp/bin/python"
args = ["/absolute/path/to/paper2zh-chatgpt-web-skill/mcp/paddleocr/server.py"]
```

The server also accepts these environment variables:

- `PADDLEOCR_ACCESS_TOKEN`
- `PADDLEOCR_MCP_AISTUDIO_ACCESS_TOKEN`
- `AISTUDIO_ACCESS_TOKEN`
- `PADDLEOCR_JOB_URL`
- `PADDLEOCR_MODEL`
- `PADDLEOCR_MCP_ENV_FILE`

## Security

Never commit `.env` or paste your access token into public files. The repository ignores `.env` by default.
