# Office Viewer Auto-Preview & Sidecar Link Policy

Whenever the user mentions, asks to inspect/preview/open a local Microsoft Office file (`.docx`, `.xlsx`, or `.pptx`), OR right after you generate or edit any local `.docx`, `.xlsx`, or `.pptx` file:

1. Automatically call `preview_office_file` on `office-viewer_office_viewer` with the file's absolute `path` and workspace `directory`.
2. **Always include the one-click sidecar pill link** in your reply so the user can open the interactive preview directly in the right-hand side pane (rather than Antigravity's built-in file viewer which does not render `.docx`):

   `[Office Viewer](sidecar://office-viewer/office_viewer/ui.html)`

3. Remind or let the user jump to their locally installed Office software (Word/Excel/PowerPoint/WPS) via the top-right `⋯` menu inside Office Viewer.
