import os
import tempfile
import unittest
from unittest.mock import patch
import app

class HomeTests(unittest.TestCase):
    def test_home_and_index_alias(self):
        client = app.app.test_client()
        for path in ('/', '/index.html'):
            response = client.get(path)
            self.assertEqual(response.status_code, 200)
            self.assertIn('공공구매', response.get_data(as_text=True))
            response.close()

    def test_home_from_other_working_directory(self):
        previous = os.getcwd()
        with tempfile.TemporaryDirectory() as directory:
            try:
                os.chdir(directory)
                response = app.app.test_client().get('/')
                self.assertEqual(response.status_code, 200)
                response.close()
            finally:
                os.chdir(previous)

    def test_missing_index_has_actionable_message(self):
        with tempfile.TemporaryDirectory() as directory, patch.object(app, 'APP_DIR', directory):
            response = app.app.test_client().get('/')
            self.assertEqual(response.status_code, 503)
            self.assertIn('ZIP 전체', response.get_data(as_text=True))
