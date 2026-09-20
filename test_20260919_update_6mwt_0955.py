import importlib.util
import unittest
from pathlib import Path

import pandas as pd


MODULE_PATH = Path(__file__).resolve().parent / "20260919_Update_6MWT_0955.py"
SPEC = importlib.util.spec_from_file_location("update_6mwt_0955", MODULE_PATH)
update_6mwt_0955 = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
SPEC.loader.exec_module(update_6mwt_0955)


class Update6MWT0955Tests(unittest.TestCase):
    def test_ci95_interval_formats_mean_confidence_interval(self):
        values = pd.Series([200.0, 240.0, 280.0, 320.0])

        result = update_6mwt_0955.ci95_interval(values)

        self.assertEqual(result, "(209.4-310.6)")

    def test_ci95_interval_returns_nan_interval_when_sample_too_small(self):
        values = pd.Series([64.0])

        result = update_6mwt_0955.ci95_interval(values)

        self.assertEqual(result, "(nan-nan)")

    def test_summarize_series_uses_nan_interval_for_missing_values(self):
        result = update_6mwt_0955.summarize_series(pd.Series([None, float("nan")]))

        self.assertEqual(result["ci_95"], "(nan-nan)")
        self.assertEqual(result["iqr_range"], "(nan-nan)")
        self.assertEqual(result["valid_count"], 0)


if __name__ == "__main__":
    unittest.main()
