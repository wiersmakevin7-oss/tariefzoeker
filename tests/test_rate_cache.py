import pickle
import unittest
from rate_models import RateRow
from test_batch_quotes import load_engine


class RateCacheTest(unittest.TestCase):
    def test_cached_books_survive_script_reload(self):
        engine = load_engine()
        paths = tuple(str(p) for p in engine.DEFAULT_WORKBOOKS if p.exists())
        self.assertTrue(paths)
        engine.load_default_rate_books.clear()
        rows = engine.load_default_rate_books(paths)
        self.assertTrue(rows)
        serialized = pickle.dumps(rows)
        reloaded = load_engine()
        self.assertIs(reloaded.RateRow, RateRow)
        self.assertEqual(pickle.loads(serialized), rows)
        self.assertEqual(reloaded.load_default_rate_books(paths), rows)
        self.assertEqual(RateRow.__module__, 'rate_models')

    def test_single_book_cache(self):
        engine = load_engine()
        path = str(engine.DEFAULT_WORKBOOKS[0])
        engine.load_default_rates.clear()
        rows = engine.load_default_rates(path)
        self.assertTrue(rows)
        self.assertEqual(pickle.loads(pickle.dumps(rows)), rows)
        self.assertEqual(engine.load_default_rates(path), rows)
