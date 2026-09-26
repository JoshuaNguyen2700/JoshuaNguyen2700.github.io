"""Live preview server for the whole site.

Serves the site at http://localhost:8000 and reloads open pages whenever a file changes.
Pages listed in LIVE_BUILDS are built on the fly from source, so edits show without running
that tool's build script. If a live build has a samples/ folder with an export in it, that
file is loaded into the page after every reload.

Run:  python scripts/dev.py        (Ctrl+C to stop)
"""
import http.server
import os
import pathlib
import time
import urllib.parse

ROOT = pathlib.Path(__file__).resolve().parent.parent
PORT = 8000
IGNORED_DIRS = {'.git', '__pycache__'}
TYPES = {'.html': 'text/html; charset=utf-8', '.css': 'text/css', '.js': 'text/javascript',
         '.json': 'application/json', '.svg': 'image/svg+xml', '.png': 'image/png',
         '.jpg': 'image/jpeg', '.ico': 'image/x-icon', '.txt': 'text/plain; charset=utf-8'}

# URL -> how to build it: source file, (placeholder, file to insert there), optional samples folder
LIVE_BUILDS = {
    '/tools/sla/SLA-Dashboard.html': {
        'source': ROOT / 'tools' / 'sla' / 'src' / 'app.html',
        'insert': ('/*__SHEETJS__*/', ROOT / 'tools' / 'sla' / 'vendor' / 'xlsx.full.min.js'),
        'samples': ROOT / 'tools' / 'sla' / 'samples',
    },
}

# Added to every page: reload when a file changes.
RELOAD_SCRIPT = "<script>new EventSource('/__reload').onmessage = () => location.reload();</script>\n"
# Added to live builds with samples: load the newest sample export into the page.
SAMPLE_SCRIPT = """<script>
fetch('/__sample?page=' + encodeURIComponent(location.pathname)).then(r => {
  if (!r.ok || !window.__slaLoadFiles) return;
  const name = decodeURIComponent(r.headers.get('X-File-Name') || 'sample.xlsx');
  return r.blob().then(b => window.__slaLoadFiles([new File([b], name)]));
});
</script>
"""


def site_stamp():
    newest = 0
    for folder, dirs, files in os.walk(ROOT):
        dirs[:] = [d for d in dirs if d not in IGNORED_DIRS]
        for f in files:
            newest = max(newest, os.stat(os.path.join(folder, f)).st_mtime)
    return newest


def sample_file(page):
    build = LIVE_BUILDS.get(page)
    folder = build and build.get('samples')
    if not folder or not folder.is_dir():
        return None
    files = [p for p in folder.iterdir() if p.suffix.lower() in ('.xlsx', '.xlsm', '.xls', '.csv')]
    return max(files, key=lambda p: p.stat().st_mtime, default=None)


def inject(html, script):
    i = html.rfind('</body>')
    return html[:i] + script + html[i:] if i >= 0 else html + script


class Handler(http.server.BaseHTTPRequestHandler):
    def do_GET(self):
        url = urllib.parse.urlsplit(self.path)
        path = urllib.parse.unquote(url.path)
        if path == '/__reload':
            self.stream_reloads()
        elif path == '/__sample':
            page = urllib.parse.parse_qs(url.query).get('page', [''])[0]
            f = sample_file(page)
            if not f:
                self.send_error(404)
                return
            self.send_body(f.read_bytes(), 'application/octet-stream', {'X-File-Name': urllib.parse.quote(f.name)})
        elif path in LIVE_BUILDS:
            build = LIVE_BUILDS[path]
            html = build['source'].read_text(encoding='utf-8')
            if build.get('insert'):
                marker, lib = build['insert']
                html = html.replace(marker, lib.read_text(encoding='utf-8'))
            if build.get('samples'):
                html = inject(html, SAMPLE_SCRIPT)
            self.send_html(html)
        else:
            self.serve_static(path)

    def serve_static(self, path):
        target = (ROOT / path.lstrip('/')).resolve()
        if target != ROOT and ROOT not in target.parents:
            self.send_error(403)
            return
        if target.is_dir():
            if not path.endswith('/'):
                # Match GitHub Pages: folders redirect to a trailing slash so relative links work
                self.send_response(301)
                self.send_header('Location', path + '/')
                self.end_headers()
                return
            target = target / 'index.html'
        if not target.is_file():
            self.send_error(404)
            return
        if target.suffix == '.html':
            self.send_html(target.read_text(encoding='utf-8'))
        else:
            self.send_body(target.read_bytes(), TYPES.get(target.suffix.lower(), 'application/octet-stream'))

    def send_html(self, html):
        self.send_body(inject(html, RELOAD_SCRIPT).encode('utf-8'), TYPES['.html'])

    def send_body(self, body, ctype, extra=None):
        self.send_response(200)
        self.send_header('Content-Type', ctype)
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Cache-Control', 'no-store')
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    def stream_reloads(self):
        self.send_response(200)
        self.send_header('Content-Type', 'text/event-stream')
        self.send_header('Cache-Control', 'no-store')
        self.end_headers()
        last = site_stamp()
        try:
            while True:
                time.sleep(0.5)
                now = site_stamp()
                if now != last:
                    last = now
                    self.wfile.write(b'data: reload\n\n')
                else:
                    self.wfile.write(b': ping\n\n')  # keep-alive; also detects closed tabs
                self.wfile.flush()
        except (BrokenPipeError, ConnectionResetError, ConnectionAbortedError):
            pass

    def log_message(self, fmt, *args):
        if not self.path.startswith('/__reload'):
            super().log_message(fmt, *args)


if __name__ == '__main__':
    server = http.server.ThreadingHTTPServer(('127.0.0.1', PORT), Handler)
    server.daemon_threads = True
    print(f'Live preview: http://localhost:{PORT}  (Ctrl+C to stop)')
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
