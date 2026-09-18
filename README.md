# Grok Word Add-in

A native **Grok AI sidebar** (task pane) for Microsoft Word.

## Features

- Chat with Grok directly inside Word
- Read current selection or entire document
- Insert Grok replies into the document
- Rewrite / improve selected text
- Smart context awareness (selection change)
- Uses your own xAI API key

## Quick Start (Sideload)

### 1. Get an xAI API Key
Go to [https://console.x.ai](https://console.x.ai) and create an API key.

### 2. Sideload the add-in in Word

**Windows / Mac (Word desktop):**
1. Open Word
2. Go to **Insert → Add-ins → My Add-ins → Upload My Add-in**
3. Select the `manifest.xml` file from this repo
4. The Grok sidebar will appear

**Or use the shared folder method (recommended for development):**
1. Create a folder: `%USERPROFILE%\AppData\Local\Microsoft\Office\16.0\Wef` (Windows)
2. Put the entire project (or at least `manifest.xml` + `src`) there
3. Restart Word → Insert → My Add-ins → Shared Folder

### 3. Enter your API key
In the sidebar, paste your xAI API key and click Save. It is stored only in your browser/local storage.

## Project Structure

```
grok-word-addin/
├── manifest.xml          # Office Add-in manifest
├── src/
│   └── taskpane/
│       ├── taskpane.html
│       ├── taskpane.css
│       └── taskpane.js
├── assets/               # Icons (optional)
└── README.md
```

## Development

This is a pure static Office Add-in (no build step required for basic use).

For local development with live reload you can use:
```bash
npx office-addin-dev-certs install
npx http-server -p 3000 --cors
```
Then update the URLs in `manifest.xml` to `https://localhost:3000`.

## Disclaimer

- This is an unofficial community project.
- You are responsible for your own xAI API usage and costs.
- Never commit your API key.

## License

MIT
