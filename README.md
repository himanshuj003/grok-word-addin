# Grok Word Add-in

**Native Grok AI sidebar for Microsoft Word**

Live repo: https://github.com/himanshuj003/grok-word-addin

## Features

- Chat with Grok directly inside Word (task pane)
- Read current selection or entire document
- Insert Grok replies into the document with one click
- Rewrite / improve selected text
- Optional auto-context from current selection
- Uses **your own** xAI API key (you control cost)

## How to use (Sideload)

### 1. Get an xAI API Key
Go to [console.x.ai](https://console.x.ai) → create an API key.

### 2. Enable GitHub Pages (so the add-in can load)
1. Go to the repo → **Settings → Pages**
2. Source: **Deploy from a branch**
3. Branch: `main` / root
4. Save

After a minute the files will be available at:
`https://himanshuj003.github.io/grok-word-addin/`

### 3. Sideload in Word

**Easiest method (Desktop Word):**
1. Download `manifest.xml` from this repo
2. Open Word → **Insert → Add-ins → My Add-ins → Upload My Add-in**
3. Choose the `manifest.xml` file
4. Click the **Open Grok** button that appears on the Home tab

**Alternative (Shared Folder method – good for development):**
- Windows: put the project in  
  `%USERPROFILE%\AppData\Local\Microsoft\Office\16.0\Wef`
- Restart Word → Insert → My Add-ins → Shared Folder

### 4. Enter your API key
Click the ⚙️ icon in the sidebar → paste your key → Save.

## Project Structure

```
grok-word-addin/
├── manifest.xml
├── src/taskpane/
│   ├── taskpane.html
│   ├── taskpane.css
│   └── taskpane.js
├── assets/
└── README.md
```

## Development notes

- Pure static add-in (no build step needed)
- Model used: `grok-3` (change in `taskpane.js` if needed)
- API is OpenAI-compatible: `https://api.x.ai/v1/chat/completions`

## Next improvements we can add

- Streaming responses
- Better document context (paragraph / heading aware)
- Tone / style presets
- Local icons
- One-click “Improve whole document”
- Support for more Office hosts (Outlook, PowerPoint later)

## License

MIT – free to use and modify.

---
**You are responsible for your own xAI API usage and costs.**
Never commit your API key.
