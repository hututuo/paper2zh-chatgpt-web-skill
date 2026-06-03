# Configuration Prompt For An AI Agent

Use this prompt when you already have an AI Studio access token and want an AI coding agent to configure this project for local use.

```text
I have a Paper2ZH ChatGPT Web Skill repository with a bundled PaddleOCR MCP server at mcp/paddleocr/server.py.

Please configure it locally without committing my secret:

1. Create a Python virtual environment for the MCP server.
2. Install mcp/paddleocr/requirements.txt into that environment.
3. Create an external secrets folder at ~/.config/paper2zh-chatgpt-web-skill/secrets.
4. Create ~/.config/paper2zh-chatgpt-web-skill/secrets/paddleocr.env from mcp/paddleocr/.env.example.
5. Put my AI Studio token into PADDLEOCR_ACCESS_TOKEN in paddleocr.env.
6. Ensure the secrets folder and env file have private permissions.
7. Add or update the Codex MCP configuration so the paddleocr server runs with:
   - command: the virtual environment python
   - args: the absolute path to mcp/paddleocr/server.py
   - env.PADDLEOCR_MCP_ENV_FILE: the absolute path to the external paddleocr.env file
8. Do not print, commit, or expose the token.
9. Verify the MCP server starts and that the repository has no committed token.
```
