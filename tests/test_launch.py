import threading
import unittest
from werkzeug.serving import make_server
from flask import Flask
import app
import launch

class LauncherTests(unittest.TestCase):
    def test_busy_port_uses_next_and_serves_real_homepage(self):
        blocker = make_server('127.0.0.1', 0, Flask('other'))
        server = None
        try:
            server = launch.create_server(app.app, blocker.server_port)
            self.assertNotEqual(server.server_port, blocker.server_port)
            token = 'test-instance'
            def identified(environ, start_response):
                def start(status, headers, exc_info=None):
                    return start_response(status, headers + [('X-Findcontracts-Instance', token)], exc_info)
                return app.app(environ, start)
            server.app = identified
            thread = threading.Thread(target=server.serve_forever, daemon=True)
            thread.start()
            url = f'http://127.0.0.1:{server.server_port}/'
            launch.verify_home(url, token)
            with self.assertRaisesRegex(RuntimeError, 'instance'):
                launch.verify_home(url, 'wrong-instance')
            server.shutdown()
            thread.join(timeout=5)
        finally:
            blocker.server_close()
            if server:
                server.server_close()

    def test_missing_assets_fail_before_server_start(self):
        import tempfile
        from pathlib import Path
        from unittest.mock import patch
        with tempfile.TemporaryDirectory() as directory, patch.object(launch, 'APP_DIR', Path(directory)), patch.object(launch, 'create_server') as create:
            self.assertEqual(launch.main(), 1)
            create.assert_not_called()
