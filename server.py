#!/usr/bin/env python3
from base64 import b64decode
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
import os
from pathlib import Path

ROOT = Path(__file__).resolve().parent
USERNAME = os.environ.get('MEETAR_DASHBOARD_USER', 'meetar')
PASSWORD = os.environ.get('MEETAR_DASHBOARD_PASSWORD')


class AuthHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(ROOT), **kwargs)

    def _authorized(self):
        if not PASSWORD:
            return True
        header = self.headers.get('Authorization', '')
        if not header.startswith('Basic '):
            return False
        try:
            raw = b64decode(header.split(' ', 1)[1]).decode('utf-8')
        except Exception:
            return False
        return raw == f'{USERNAME}:{PASSWORD}'

    def do_HEAD(self):
        if not self._authorized():
            self.send_response(HTTPStatus.UNAUTHORIZED)
            self.send_header('WWW-Authenticate', 'Basic realm="Meetar Dashboard", charset="UTF-8"')
            self.send_header('Cache-Control', 'no-store')
            self.end_headers()
            return
        return super().do_HEAD()

    def do_GET(self):
        if not self._authorized():
            self.send_response(HTTPStatus.UNAUTHORIZED)
            self.send_header('WWW-Authenticate', 'Basic realm="Meetar Dashboard", charset="UTF-8"')
            self.send_header('Cache-Control', 'no-store')
            self.send_header('Content-Type', 'text/plain; charset=utf-8')
            self.end_headers()
            self.wfile.write('Authentication required'.encode('utf-8'))
            return
        return super().do_GET()


def main():
    port = 8123
    server = ThreadingHTTPServer(('0.0.0.0', port), AuthHandler)
    print(f'http://127.0.0.1:{port}')
    print(f'http://192.168.16.138:{port}')
    if PASSWORD:
        print(f'username: {USERNAME}')
        print('password: set from MEETAR_DASHBOARD_PASSWORD')
    else:
        print('password: not set; local auth disabled')
    server.serve_forever()


if __name__ == '__main__':
    main()
