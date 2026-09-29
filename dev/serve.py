"""Local test server: serves the page wired to the in-browser backend.

    python dev/serve.py        then open http://localhost:8335/

/        the page talking to Code.gs running against a fake spreadsheet
/plain   the page exactly as it ships
"""
import http.server
import pathlib

ROOT = pathlib.Path(__file__).resolve().parent.parent
PORT = 8335
INJECT = (
    '<script src="dev/mock-gas.js"></script>\n'
    '<script src="apps-script/Code.gs"></script>\n'
    '<script src="dev/wire.js"></script>\n'
    '</head>'
)


class Handler(http.server.SimpleHTTPRequestHandler):
    extensions_map = {**http.server.SimpleHTTPRequestHandler.extensions_map, '.gs': 'text/javascript'}

    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(ROOT), **kwargs)

    def do_GET(self):
        path = self.path.split('?')[0]
        if path not in ('/', '/index.html', '/plain'):
            return super().do_GET()
        page = (ROOT / 'index.html').read_text(encoding='utf-8')
        if path != '/plain':
            page = page.replace('</head>', INJECT, 1)
        body = page.encode('utf-8')
        self.send_response(200)
        self.send_header('Content-Type', 'text/html; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Cache-Control', 'no-store')
        self.end_headers()
        self.wfile.write(body)

    def end_headers(self):
        self.send_header('Cache-Control', 'no-store')
        super().end_headers()


if __name__ == '__main__':
    print(f'http://localhost:{PORT}/')
    http.server.ThreadingHTTPServer(('127.0.0.1', PORT), Handler).serve_forever()
