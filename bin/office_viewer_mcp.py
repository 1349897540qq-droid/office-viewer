#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Lightweight MCP stdio server for Office Viewer plugin (v2.1.1).
Syncs active file & directory into shared Sidecar state.json so the right-hand
Office Viewer Sidecar UI updates immediately, with zero inline blank iframe.
"""

import argparse
import json
import os
import sys
import traceback

PLUGIN_ROOT_DEFAULT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(PLUGIN_ROOT_DEFAULT, "lib"))

import office_core  # noqa: E402


class OfficeViewerMCP:
    def __init__(self, data_dir=None):
        # Always use the shared sidecar state dir so MCP & Sidecar stay 100% synced
        self.store = office_core.SharedStateStore(office_core.default_shared_state_dir())

    def get_tools(self):
        return [
            {
                "name": "preview_office_file",
                "description": (
                    "Load a local Microsoft Office file (.docx, .xlsx, or .pptx) into the "
                    "Office Viewer right-hand sidecar panel, automatically detecting installed "
                    "desktop Office apps for handoff. After calling this tool, always include "
                    "[Office Viewer](sidecar://office-viewer/office_viewer/ui.html) in your response."
                ),
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "path": {
                            "type": "string",
                            "description": "Absolute path to a local .docx, .xlsx, or .pptx file.",
                        },
                        "directory": {
                            "type": "string",
                            "description": "Optional workspace or folder path to list sibling Office files in the drawer.",
                        },
                    },
                    "additionalProperties": False,
                },
            },
            {
                "name": "list_office_files",
                "description": "List all local .docx, .xlsx, and .pptx files in a directory (skipping temporary ~$ files).",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "directory": {
                            "type": "string",
                            "description": "Absolute directory path to scan.",
                        },
                        "max_depth": {
                            "type": "integer",
                            "description": "Maximum folder recursion depth (default: 3).",
                            "default": 3,
                        },
                    },
                    "required": ["directory"],
                    "additionalProperties": False,
                },
            },
        ]

    def handle(self, req):
        method = req.get("method")
        req_id = req.get("id")
        params = req.get("params") or {}
        if req_id is None:
            return None

        try:
            if method == "initialize":
                proto = params.get("protocolVersion") or "2025-03-26"
                return {
                    "jsonrpc": "2.0",
                    "id": req_id,
                    "result": {
                        "protocolVersion": proto,
                        "capabilities": {"tools": {"listChanged": False}},
                        "serverInfo": {"name": "office-viewer", "version": "2.1.1"},
                    },
                }

            if method == "ping":
                return {"jsonrpc": "2.0", "id": req_id, "result": {}}

            if method == "tools/list":
                return {"jsonrpc": "2.0", "id": req_id, "result": {"tools": self.get_tools()}}

            if method == "tools/call":
                name = params.get("name", "")
                args = params.get("arguments") or {}

                if name == "preview_office_file":
                    path = args.get("path") or ""
                    directory = args.get("directory") or ""
                    payload, summary = office_core.load_office_payload(
                        path, directory, self.store, allow_large=False
                    )
                    apps = [a["name"] for a in payload.get("availableApps", [])]
                    apps_str = ("；已识别本机软件：" + "、".join(apps)) if apps else ""
                    out_text = (
                        "%s%s。\n请通过侧边栏链接打开或聚焦预览：[Office Viewer](sidecar://office-viewer/office_viewer/ui.html)"
                        % (summary, apps_str)
                    )
                    return {
                        "jsonrpc": "2.0",
                        "id": req_id,
                        "result": {
                            "content": [{"type": "text", "text": out_text}],
                            "isError": payload.get("mode") == "error",
                        },
                    }

                if name == "list_office_files":
                    directory = args.get("directory") or self.store.reload().get("lastDirectory") or ""
                    max_depth = int(args.get("max_depth", 3))
                    files = office_core.scan_office_files(directory, max_depth=max_depth)
                    if directory and os.path.isdir(os.path.expanduser(directory)):
                        self.store.record(directory=directory)
                    lines = ["在 %s 中找到 %d 个 Office 文件：" % (directory, len(files))]
                    for f in files[:30]:
                        lines.append("- %s (%s, %s)" % (f["relPath"], f["sizeHuman"], f["mtimeHuman"]))
                    return {
                        "jsonrpc": "2.0",
                        "id": req_id,
                        "result": {"content": [{"type": "text", "text": "\n".join(lines)}]},
                    }

                return {
                    "jsonrpc": "2.0",
                    "id": req_id,
                    "error": {"code": -32601, "message": "Unknown tool: %s" % name},
                }

            return {
                "jsonrpc": "2.0",
                "id": req_id,
                "error": {"code": -32601, "message": "Method not found: %s" % method},
            }

        except Exception as e:
            sys.stderr.write("MCP error: %s\n%s\n" % (e, traceback.format_exc()))
            sys.stderr.flush()
            return {
                "jsonrpc": "2.0",
                "id": req_id,
                "result": {
                    "content": [{"type": "text", "text": "执行出错：%s" % str(e)}],
                    "isError": True,
                },
            }

    def run(self):
        for raw_line in sys.stdin:
            line = raw_line.strip()
            if not line:
                continue
            try:
                req = json.loads(line)
            except Exception:
                continue
            resp = self.handle(req)
            if resp is not None:
                sys.stdout.write(json.dumps(resp, ensure_ascii=False) + "\n")
                sys.stdout.flush()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--plugin-root", default=PLUGIN_ROOT_DEFAULT)
    parser.add_argument("--data-dir", default="")
    args = parser.parse_args()
    if args.plugin_root:
        lib_dir = os.path.join(args.plugin_root, "lib")
        if lib_dir not in sys.path:
            sys.path.insert(0, lib_dir)
    OfficeViewerMCP(args.data_dir).run()


if __name__ == "__main__":
    main()
