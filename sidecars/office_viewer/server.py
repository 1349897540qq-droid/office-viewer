#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Antigravity UI Sidecar HTTP Server for Office Viewer (v2.1).
Binds to 127.0.0.1:${ANTIGRAVITY_SIDECAR_WEB_PORT}, serves /ui.html & /preload.js,
supports HEAD/GET/POST/OPTIONS, SSE live sync, and one-click external Office launching.
"""

import json
import os
import re
import signal
import sys
import time
import urllib.parse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

SIDECAR_DIR = os.path.dirname(os.path.abspath(__file__))
PLUGIN_ROOT = os.path.dirname(os.path.dirname(SIDECAR_DIR))
sys.path.insert(0, os.path.join(PLUGIN_ROOT, "lib"))

import office_core  # noqa: E402

UI_HTML_PATH = os.path.join(SIDECAR_DIR, "ui.html")


def get_port():
    raw = os.environ.get("ANTIGRAVITY_SIDECAR_WEB_PORT", "").strip()
    if raw.isdigit():
        return int(raw)
    return 8765


def get_token():
    for k in (
        "ANTIGRAVITY_SIDECAR_UI_TOKEN",
        "ANTIGRAVITY_SIDECAR_AUTH_TOKEN",
        "ANTIGRAVITY_SIDECAR_WEB_TOKEN",
    ):
        v = os.environ.get(k, "").strip()
        if v:
            return v
    return ""


def load_preload_js():
    app_data = os.environ.get("ANTIGRAVITY_APP_DATA_DIR", "").strip()
    if not app_data:
        app_data = os.path.join(os.path.expanduser("~"), ".gemini", "antigravity")
    cand = os.path.join(app_data, "lib", "node_modules", "sidecar_sdk", "preload.js")
    try:
        with open(cand, "r", encoding="utf-8") as f:
            txt = f.read()
            if txt.strip():
                return txt
    except Exception:
        pass
    return "// Antigravity sidecar preload bridge stub\nwindow.__sidecarPreloadLoaded = true;\n"


def inject_token_into_html(html_str, token, conv_id=""):
    if not token:
        return html_str
    qs = "token=%s" % urllib.parse.quote(token)
    if conv_id:
        qs += "&conversationId=%s" % urllib.parse.quote(conv_id)

    def repl(m):
        attr, quote, url = m.group(1), m.group(2), m.group(3)
        if url.startswith("//"):
            return m.group(0)
        parts = url.split("#", 1)
        base = parts[0]
        frag = ("#" + parts[1]) if len(parts) > 1 else ""
        sep = "&" if "?" in base else "?"
        return '%s=%s%s%s%s%s%s' % (attr, quote, base, sep, qs, frag, quote)

    out = re.sub(r'\b(href|src)=(["\'])(/[^"\'\s>]*)\2', repl, html_str)
    # Also inject runtime token meta so JS works even if token came via header
    meta = '<meta name="ov-token" content="%s"/>' % token.replace('"', "")
    if "</head>" in out:
        out = out.replace("</head>", "  %s\n</head>" % meta, 1)
    return out


def normalize_file_param(val):
    if not val:
        return ""
    v = val.strip()
    if v.startswith("file://"):
        parsed = urllib.parse.urlparse(v)
        v = urllib.parse.unquote(parsed.path)
    return v


class SidecarHandler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"

    def log_message(self, fmt, *args):
        pass

    def _cors_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, HEAD, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, X-Sidecar-Token")

    def _check_token(self, query):
        expected = self.server.auth_token
        if not expected:
            return True
        hdr = (self.headers.get("X-Sidecar-Token") or "").strip()
        qtok = (query.get("token", [""])[0] or "").strip()
        return hdr == expected or qtok == expected

    def _send_bytes(self, body, content_type, status=200, head_only=False):
        self.send_response(status)
        self._cors_headers()
        self.send_header("Content-Type", content_type)
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        if not head_only and body:
            self.wfile.write(body)

    def _send_json(self, obj, status=200, head_only=False):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self._send_bytes(body, "application/json; charset=utf-8", status=status, head_only=head_only)

    def _send_html(self, html_str, status=200, head_only=False):
        body = html_str.encode("utf-8")
        self._send_bytes(body, "text/html; charset=utf-8", status=status, head_only=head_only)

    def do_OPTIONS(self):
        self.send_response(204)
        self._cors_headers()
        self.send_header("Content-Length", "0")
        self.end_headers()

    def do_HEAD(self):
        self._handle_get(head_only=True)

    def do_GET(self):
        self._handle_get(head_only=False)

    def _handle_get(self, head_only=False):
        parsed = urllib.parse.urlparse(self.path)
        pathname = parsed.path or "/"
        query = urllib.parse.parse_qs(parsed.query)

        if pathname in ("/", "/ui.html", "/index.html"):
            try:
                with open(UI_HTML_PATH, "r", encoding="utf-8") as f:
                    raw_html = f.read()
                tok = (query.get("token", [""])[0] or "").strip() or (self.headers.get("X-Sidecar-Token") or "").strip() or self.server.auth_token
                conv = (query.get("conversationId", [""])[0] or "").strip()
                rendered = inject_token_into_html(raw_html, tok, conv)
                self._send_html(rendered, 200, head_only=head_only)
            except Exception as e:
                self._send_html("<h1>Error loading ui.html: %s</h1>" % e, 500, head_only=head_only)
            return

        if pathname == "/preload.js" or pathname.endswith("/preload.js"):
            js = self.server.preload_js.encode("utf-8")
            self._send_bytes(js, "application/javascript; charset=utf-8", 200, head_only=head_only)
            return

        if pathname == "/healthz":
            self._send_json({"ok": True}, 200, head_only=head_only)
            return

        if pathname == "/api/state":
            st = self.server.store.reload()
            self._send_json({
                "version": int(st.get("version", 0)),
                "lastFile": st.get("lastFile", ""),
                "lastDirectory": st.get("lastDirectory", ""),
                "updatedAt": st.get("updatedAt", ""),
            }, 200, head_only=head_only)
            return

        if pathname == "/api/preview":
            file_param = ""
            for key in ("path", "file", "filepath", "filePath", "uri"):
                v = (query.get(key, [""])[0] or "").strip()
                if v:
                    file_param = normalize_file_param(v)
                    break
            dir_param = normalize_file_param((query.get("directory", [""])[0] or "").strip())
            self.server.store.reload()
            try:
                payload, _ = office_core.load_office_payload(
                    file_param, dir_param, self.server.store, allow_large=True, bump_version=False
                )
                payload["stateVersion"] = int(self.server.store.reload().get("version", 0))
                self._send_json(payload, 200, head_only=head_only)
            except Exception as e:
                self._send_json({"mode": "error", "error": str(e)}, 200, head_only=head_only)
            return

        if pathname == "/api/events":
            self.send_response(200)
            self._cors_headers()
            self.send_header("Content-Type", "text/event-stream; charset=utf-8")
            self.send_header("Cache-Control", "no-cache, no-transform")
            self.send_header("Connection", "keep-alive")
            self.end_headers()
            if head_only:
                return
            last_ver = -1
            try:
                for _ in range(3600):
                    st = self.server.store.reload()
                    ver = int(st.get("version", 0))
                    if ver != last_ver:
                        last_ver = ver
                        msg = json.dumps({"version": ver, "lastFile": st.get("lastFile", ""), "lastDirectory": st.get("lastDirectory", "")}, ensure_ascii=False)
                        self.wfile.write(("data: %s\n\n" % msg).encode("utf-8"))
                        self.wfile.flush()
                    time.sleep(0.8)
            except (BrokenPipeError, ConnectionResetError, OSError):
                return
            return

        self._send_json({"error": "Not Found"}, 404, head_only=head_only)

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        pathname = parsed.path or "/"
        query = urllib.parse.parse_qs(parsed.query)

        if not self._check_token(query):
            self._send_json({"error": "unauthorized"}, 401)
            return

        length = int(self.headers.get("Content-Length") or "0")
        raw = self.rfile.read(min(length, 2_000_000)) if length > 0 else b"{}"
        try:
            body = json.loads(raw.decode("utf-8") or "{}")
        except Exception:
            body = {}

        if pathname == "/api/open-external":
            try:
                msg = office_core.open_external_app(
                    body.get("path", ""),
                    mode=body.get("mode", "default"),
                    app_path=body.get("appPath", ""),
                )
                self._send_json({"ok": True, "message": msg}, 200)
            except Exception as e:
                self._send_json({"ok": False, "error": str(e)}, 400)
            return

        self._send_json({"error": "Not Found"}, 404)


def write_server_info(data_dir, host, port, token):
    try:
        os.makedirs(data_dir, exist_ok=True)
        info_path = os.path.join(data_dir, "server_info.json")
        with open(info_path, "w", encoding="utf-8") as f:
            json.dump({"host": host, "port": port, "token": token or None, "pid": os.getpid()}, f)
    except Exception:
        pass


def main():
    port = get_port()
    token = get_token()
    data_dir = office_core.default_shared_state_dir()
    store = office_core.SharedStateStore(data_dir)

    server = ThreadingHTTPServer(("127.0.0.1", port), SidecarHandler)
    server.daemon_threads = True
    server.auth_token = token
    server.store = store
    server.preload_js = load_preload_js()
    actual_port = server.server_address[1]
    write_server_info(data_dir, "127.0.0.1", actual_port, token)

    def shutdown_handler(_signum, _frame):
        raise KeyboardInterrupt()

    signal.signal(signal.SIGTERM, shutdown_handler)
    signal.signal(signal.SIGINT, shutdown_handler)

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
