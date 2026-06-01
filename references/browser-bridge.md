# Browser Bridge

Default mode is manual.

For each exchange manifest:

1. Open ChatGPT Web.
2. Upload the `input_json`, or paste its contents below the prompt.
3. Paste the matching `prompt_file`.
4. Ask ChatGPT to return or create exactly the expected output filename.
5. Download the JSON file into the opened `downloads/` folder.
6. Run `accept-downloads`.

The local pipeline does not depend on ChatGPT page selectors. Browser automation may be added later by reading only `*_exchange.json` files.
