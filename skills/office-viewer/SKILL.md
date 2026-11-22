---
name: office-viewer
description: Preview local Microsoft Office documents (.docx, .xlsx, .pptx) interactively in the Antigravity side pane via the Office Viewer UI sidecar, with one-click handoff to locally installed Office apps (Microsoft Word/Excel/PowerPoint, WPS Office, Pages/Numbers/Keynote, LibreOffice). Trigger automatically whenever generating, editing, inspecting, or asked to preview/open any local .docx, .xlsx, or .pptx file.
---

# Office Viewer (Side Pane Sidecar)

The Office Viewer sidecar renders local `.docx`, `.xlsx`, and `.pptx` files in the Antigravity right-hand auxiliary pane with high-fidelity formatting, tables, embedded images, outlines, Excel sheets, PowerPoint slides, and a top-right `⋯` menu that automatically detects installed desktop Office apps for one-click opening.

## Standard Preview Workflow

Whenever the user asks to view/preview a local `.docx`, `.xlsx`, or `.pptx` file, OR right after generating/editing one:

1. Call **`preview_office_file`** on MCP server `office-viewer_office_viewer` with:
   - `path`: absolute path to the Office file
   - `directory`: workspace or parent folder path (populates the quick-switch file drawer)
2. **Always surface the sidecar pill link** in your chat response so the user can open/focus the right-hand side pane with one click:

```markdown
[Office Viewer](sidecar://office-viewer/office_viewer/ui.html)
```

3. If the side pane is already open, it automatically switches/refreshes to the newly loaded file within 1 second via live state sync. Clicking the top-right `⋯` menu inside the viewer lets the user open the file directly in Microsoft Word / Excel / PowerPoint / WPS Office / Pages / LibreOffice or reveal it in Finder.
