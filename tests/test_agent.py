import json
import tempfile
import unittest
from pathlib import Path

import agent


class InputResolutionTests(unittest.TestCase):
    def test_url_marketplace_wins_over_default(self):
        seeds = agent.resolve_seeds(
            [{'url': 'https://www.amazon.de/dp/B0D9VS4PV7'}], 'US'
        )
        self.assertEqual(seeds, [('B0D9VS4PV7', 'DE')])

    def test_marketplace_column_is_supported(self):
        seeds = agent.resolve_seeds(
            [{'ASIN': 'B0D9VS4PV7', 'Marketplace': 'UK'}], None
        )
        self.assertEqual(seeds, [('B0D9VS4PV7', 'UK')])

    def test_missing_marketplace_is_actionable(self):
        with self.assertRaisesRegex(ValueError, 'Marketplace missing'):
            agent.resolve_seeds([{'ASIN': 'B0D9VS4PV7'}], None)


class ConfigTests(unittest.TestCase):
    def test_config_accepts_project_defaults(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / 'config.json'
            path.write_text(json.dumps({
                'input_dir': 'xray helium10', 'output_dir': 'reports',
                'marketplace': 'US', 'headless': True,
                'max_pages': 2, 'max_products': 10, 'delay': 0,
            }), encoding='utf-8')
            self.assertEqual(agent.load_config(path)['marketplace'], 'US')


if __name__ == '__main__':
    unittest.main()
