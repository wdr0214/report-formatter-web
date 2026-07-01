#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import http.server
import socketserver
import webbrowser
from pathlib import Path


PORT = 8765
ROOT = Path(__file__).resolve().parent


class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(ROOT), **kwargs)


if __name__ == "__main__":
    with socketserver.TCPServer(("127.0.0.1", PORT), Handler) as httpd:
        url = f"http://127.0.0.1:{PORT}/"
        print(f"本地预览地址：{url}")
        try:
            webbrowser.open(url)
        except Exception:
            pass
        httpd.serve_forever()
