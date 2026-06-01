# Configuration Prompt For An AI Agent

Use this prompt when you already have an AI Studio access token and want an AI coding agent to configure this project for local use.

```text
I have a Paper2ZH ChatGPT Web Skill repository with a bundled PaddleOCR MCP server at mcp/paddleocr/server.py.

Please configure it locally without committing my secret:

1. Create a Python virtual environment for the MCP server.
2. Install mcp/paddleocr/requirements.txt into that environment.
3. Create mcp/paddleocr/.env from mcp/paddleocr/.env.example.
4. Put my AI Studio token into PADDLEOCR_ACCESS_TOKEN in .env.
5. Ensure .env is ignored by Git and has private permissions.
6. Add or update the Codex MCP configuration so the paddleocr server runs with:
   - command: the virtual environment python
   - args: the absolute path to mcp/paddleocr/server.py
7. Do not print, commit, or expose the token.
8. Verify the MCP server starts and that the repository has no committed token.
```

