import importlib.util
from pathlib import Path
import sqlite3
import tempfile
import unittest


MODULE_PATH = Path(__file__).resolve().parents[1] / 'deploy' / 'bootstrap-from-existing.py'
SPEC = importlib.util.spec_from_file_location('portal_bootstrap', MODULE_PATH)
bootstrap = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(bootstrap)


class BootstrapFromExistingTests(unittest.TestCase):
    def test_reads_legacy_credentials_without_executing_application(self):
        source = '''
LDAP_CONFIG = {
    "uri": "ldap://directory.local",
    "base": "DC=example,DC=local",
    "bind_dn": "CN=service,DC=example,DC=local",
    "bind_password": "test-only-password",
    "user_attr": "sAMAccountName",
}
SMTP_PASSWORD = os.environ.get("SMTP_PASSWORD", "test-only-smtp")
'''
        with tempfile.TemporaryDirectory() as directory:
            app_path = Path(directory) / 'app.py'
            app_path.write_text(source, encoding='utf-8')
            values = bootstrap.legacy_code_settings(app_path)
        self.assertEqual(values['LDAP_URI'], 'ldap://directory.local')
        self.assertEqual(values['LDAP_BIND_PASSWORD'], 'test-only-password')
        self.assertEqual(values['SMTP_PASSWORD'], 'test-only-smtp')

    def test_sqlite_database_is_copied_with_backup_api(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / 'source.db'
            target = root / 'target.db'
            connection = sqlite3.connect(source)
            connection.execute('CREATE TABLE sample (value TEXT)')
            connection.execute('INSERT INTO sample VALUES (?)', ('preserved',))
            connection.commit()
            connection.close()

            original_project_dir = bootstrap.PROJECT_DIR
            bootstrap.PROJECT_DIR = root
            try:
                bootstrap.backup_sqlite(source, target)
            finally:
                bootstrap.PROJECT_DIR = original_project_dir

            copied = sqlite3.connect(target)
            try:
                value = copied.execute('SELECT value FROM sample').fetchone()[0]
            finally:
                copied.close()
            self.assertEqual(value, 'preserved')


if __name__ == '__main__':
    unittest.main()
