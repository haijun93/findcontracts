"""Start the app, verify its response, then open the correct browser address."""
from pathlib import Path
import secrets
import sys
import threading
import urllib.request
import webbrowser

APP_DIR = Path(__file__).resolve().parent
VERSION = '2026-10-08-launcher-v2'


def create_server(application, first_port=5050, attempts=10):
    from werkzeug.serving import make_server
    for port in range(first_port, first_port + attempts):
        try:
            return make_server('127.0.0.1', port, application, threaded=True)
        except (OSError, SystemExit):
            print(f'Port {port} is unavailable; trying the next port.', flush=True)
    raise RuntimeError('No free local port. Close previous app windows and retry.')


def verify_home(url, token):
    # Local readiness must not depend on a system HTTP proxy.
    opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
    with opener.open(url, timeout=5) as response:
        body = response.read().decode('utf-8')
        if response.status != 200 or response.headers.get('X-Findcontracts-Instance') != token:
            raise RuntimeError('The response is not from this app instance.')
        if '공공구매' not in body:
            raise RuntimeError('The homepage response is invalid.')


def main(open_browser=True):
    print(f'Findcontracts version: {VERSION}', flush=True)
    print(f'Application directory: {APP_DIR}', flush=True)
    missing = [name for name in ('app.py', 'index.html') if not (APP_DIR / name).is_file()]
    if missing:
        print('Missing files: ' + ', '.join(missing), flush=True)
        print('Extract the ENTIRE ZIP into a new folder, then run start.bat.', flush=True)
        return 1
    from app import app
    token = secrets.token_hex(16)

    @app.after_request
    def identify_instance(response):
        response.headers['X-Findcontracts-Instance'] = token
        response.headers['Cache-Control'] = 'no-store'
        return response

    server = create_server(app)
    url = f'http://127.0.0.1:{server.server_port}/'
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    try:
        verify_home(url, token)
        print(f'Homepage verified. Open: {url}', flush=True)
        print('Keep this window open. Press Ctrl+C to stop.', flush=True)
        if open_browser and not webbrowser.open(url):
            print('Browser did not open automatically. Open the address above.', flush=True)
        while thread.is_alive():
            thread.join(timeout=1)
    except KeyboardInterrupt:
        print('Stopping server.', flush=True)
    finally:
        server.shutdown()
        server.server_close()
    return 0


if __name__ == '__main__':
    try:
        sys.exit(main(open_browser="--no-browser" not in sys.argv))
    except Exception as error:
        print(f'Startup failed: {error}', flush=True)
        sys.exit(1)
