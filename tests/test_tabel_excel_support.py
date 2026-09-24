import unittest
from pathlib import Path

import app as portal_app


class TabelExcelSupportTests(unittest.TestCase):
    def test_legacy_xls_reader_is_installed(self):
        self.assertIsNotNone(
            portal_app.xlrd,
            'Файлы табелей .xls требуют установленную библиотеку xlrd',
        )

    def test_xlrd_is_pinned_in_production_requirements(self):
        requirements = Path('requirements.txt').read_text(encoding='utf-8-sig')
        self.assertIn('xlrd==2.0.2', requirements.splitlines())


if __name__ == '__main__':
    unittest.main()
