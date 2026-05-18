# 🌿 FCG Grow

**AI-powered Kanban task board for entrepreneurs and small business owners.**  
Free and open source. Built by [Fairy Circle Garden](https://fairycirclegarden.com).

---

## What It Is

FCG Grow is a lightweight, single-file desktop Kanban board with a built-in AI assistant. No subscriptions, no cloud accounts, no data leaving your machine. Just download, run, and start managing your business operations.

Built for the way small business owners actually work — juggling grants, outreach, admin, finance, and everything in between.

---

## Features

- **AI Chat Panel** — Ask your board questions, get summaries, move tasks by chat input. Supports OpenRouter, Anthropic, and local Ollama models
- **Kanban Board** — Four columns: To Do, In Progress, Review, Done
- **Drag & Drop** — Pick up cards from any column and drop them anywhere
- **Category Badges** — Click any badge to reassign a task's category instantly
- **Filter & Search** — Filter by category or search in real time
- **Workspaces** — Multiple boards, one app
- **Dark UI** — Easy on the eyes, dark titlebars on Windows
- **Portable** — Runs as a `.pyw` script or compiled `.exe`, no installer needed

---

## Quick Start (No Python Required)

1. Download `FCGGrow.exe` from [Releases](../../releases)
2. Drop it in a folder with `fcggrow_help.html`
3. Double-click and go

> First run will create a `fcggrow_workspaces/` folder automatically next to the exe.

---

## Run From Source

**Requirements:** Python 3.10+

```bash
pip install -r requirements.txt
python FCGGrow.pyw
```

---

## AI Setup

FCG Grow supports three AI providers. Add your key in **Tools → Settings → AI Providers**:

| Provider | Notes |
|---|---|
| [OpenRouter](https://openrouter.ai) | Recommended — access to many free models |
| [Anthropic](https://console.anthropic.com) | Claude models directly |
| [Ollama](https://ollama.com) | 100% local, no API key needed |

---

## Build the Exe

```bash
pip install pyinstaller
python -m PyInstaller --onefile --windowed --icon=fcggrow.ico --name=FCGGrow FCGGrow.pyw
```

Or just double-click `build.bat` if you're on Windows.

---

## File Structure

```
FCGGrow/
├── FCGGrow.pyw           ← main app
├── fcggrow_help.html     ← help page
├── fcggrow.ico           ← app icon
├── build.bat             ← Windows build script
├── requirements.txt
└── fcggrow_workspaces/   ← auto-created, stores your tasks
```

---

## Built By

**Fairy Circle Garden** — Small business tools for entrepreneurs.  
Open source forever. Free to use, fork, and build on.

[fairycirclegarden.com](https://fairycirclegarden.com) · [FCG-Builds on GitHub](https://github.com/FCG-Builds)

---

## License

MIT — see [LICENSE](LICENSE)
