# Office Viewer for Antigravity

A lightweight, zero-dependency **Antigravity UI Sidecar + MCP plugin** for previewing local Microsoft Office documents (`.docx`, `.xlsx`, `.pptx`) directly inside the Antigravity side pane, with automatic local Office app detection and one-click external handoff.

## Highlights

- **Zero External Dependencies**: Built entirely with Python 3 standard library + self-contained HTML/CSS/JS—no `pip`, `npm`, or LibreOffice required, fast startup, 100% offline.
- **High-Fidelity Word (`.docx`) Preview**: Clean A4 paper canvas preserving headings, tables (including merged cells, column widths, shading, and callout boxes), rich text styles, embedded images, page breaks, and an auto-generated clickable **Outline TOC**.
- **Interactive Excel (`.xlsx`) Preview**: Multi-sheet tab switching, merged cells, Excel serial date formatting, sticky row/column headers, live row search filtering, and click-to-copy cell values.
- **PowerPoint (`.pptx`) Slide Viewer**: Slide-by-slide stage with keyboard arrow navigation (`←` / `→`), embedded images, tables, and speaker notes.
- **Live Agent Sync (MCP + SSE)**: When your agent generates, updates, or opens an Office document via `preview_office_file`, the side pane updates automatically in real time via Server-Sent Events.
- **Smart Desktop App Handoff (`⋯` Menu)**: Automatically detects locally installed desktop Office applications (Microsoft Word/Excel/PowerPoint, WPS Office, Apple Pages/Numbers/Keynote, LibreOffice, ONLYOFFICE) across macOS, Windows, and Linux for one-click opening, plus reveal-in-Finder/Explorer, copy path, and print/save to PDF.

## Prerequisites

- **Google Antigravity** (App or IDE with Sidecar UI support)
- **Python 3.8+** available as `python3` in `PATH`

## Installation

### Option 1: Install via Antigravity CLI

```bash
agy plugin install https://github.com/1349897540qq-droid/office-viewer
```

### Option 2: Manual Install

Clone or copy this repository into your global Antigravity plugins directory:

```bash
git clone https://github.com/1349897540qq-droid/office-viewer.git ~/.gemini/config/plugins/office-viewer
```

Antigravity discovers and enables the plugin automatically without requiring a full restart.

## Usage

1. Ask your agent to preview any local `.docx`, `.xlsx`, or `.pptx` file—or generate/edit an Office document as usual.
2. The agent automatically calls `preview_office_file` on `office-viewer_office_viewer` and surfaces a one-click **`Office Viewer`** sidecar pill in chat.
3. Inside the right-hand side pane:
   - Click **`☰`** to browse sibling Office files in your workspace or jump through document headings.
   - Use **`−` / `+`** to zoom and **`↻`** to reload after external edits.
   - Click **`⋯`** in the top-right corner to open the current document directly in Microsoft Word / Excel / PowerPoint, WPS Office, Pages / Numbers / Keynote, LibreOffice, or reveal it in Finder/Explorer.

## Architecture

```text
office-viewer/
├── plugin.json                     # Plugin manifest & marketplace metadata
├── mcp_config.json                 # MCP stdio server registration
├── assets/logo.svg                 # High-contrast square vector icon
├── lib/office_core.py              # Pure-Python OOXML parser & cross-platform app detector
├── bin/office_viewer_mcp.py        # MCP tool server (preview_office_file, list_office_files)
├── sidecars/office_viewer/
│   ├── sidecar.json                # Antigravity AUX_PANE UI sidecar manifest
│   ├── server.py                   # Local HTTP/SSE server bound to ANTIGRAVITY_SIDECAR_WEB_PORT
│   └── ui.html                     # Self-contained responsive preview UI
├── skills/office-viewer/SKILL.md   # Agent skill instructions
└── rules/AGENTS.md                 # Auto-preview & sidecar link policy
```

## License

MIT License — see [LICENSE](LICENSE).
