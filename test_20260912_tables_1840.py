import importlib.util
import unittest
from pathlib import Path

import pandas as pd


MODULE_PATH = Path(__file__).resolve().parent / "20260912_Tables_1840.py"
SPEC = importlib.util.spec_from_file_location("tables_1840", MODULE_PATH)
tables_1840 = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
SPEC.loader.exec_module(tables_1840)


class Tables1840Tests(unittest.TestCase):
    def test_build_compact_performance_table_formats_metrics(self):
        model_explainers = pd.DataFrame(
            [
                {
                    "Overall_Rank": 1,
                    "CV_R2": 0.54321,
                    "CV_MAE": 89.285,
                    "Weighted_OOF_R2": 0.51234,
                    "Weighted_OOF_MAE": 90.44,
                },
                {
                    "Overall_Rank": 2,
                    "CV_R2": None,
                    "CV_MAE": None,
                    "Weighted_OOF_R2": 0.40001,
                    "Weighted_OOF_MAE": None,
                },
            ]
        )
        panel_a = pd.DataFrame(
            [
                {
                    "Rank": 1,
                    "Acronym": "RESTORE",
                    "Publication name": "Model One",
                    "Input vars": 21,
                    "Best BalAcc [95% CI]": "80.0% [70.0, 90.0]",
                    "Worst BalAcc [95% CI]": "78.0% [68.0, 88.0]",
                },
                {
                    "Rank": 2,
                    "Acronym": "AIMS",
                    "Publication name": "Model Two",
                    "Input vars": 12,
                    "Best BalAcc [95% CI]": "75.0% [65.0, 85.0]",
                    "Worst BalAcc [95% CI]": "73.0% [63.0, 83.0]",
                },
            ]
        )

        result = tables_1840.build_compact_performance_table(model_explainers, panel_a)

        self.assertEqual(result.loc[0, "Part 1 CV R²"], "0.5432")
        self.assertEqual(result.loc[0, "Part 1 MAE (m)"], "89.28")
        self.assertEqual(result.loc[0, "Part 2 weighted OOF R²"], "0.5123")
        self.assertEqual(result.loc[0, "Part 2 weighted OOF MAE (m)"], "90.4")
        self.assertEqual(result.loc[1, "Part 1 CV R²"], "—")
        self.assertEqual(result.loc[1, "Part 2 weighted OOF MAE (m)"], "—")

    def test_build_compact_predictor_table_keeps_shared_predictors_only(self):
        model_explainers = pd.DataFrame(
            [
                {"Overall_Rank": 1, "Acronym": "RESTORE", "Input_vars": 21},
                {"Overall_Rank": 2, "Acronym": "RESTORE", "Input_vars": 21},
            ]
        )
        lasso_coefficients = pd.DataFrame(
            [
                {
                    "Overall_Rank": 1,
                    "Acronym": "RESTORE",
                    "Input_vars": 21,
                    "Model_label": "Model 5",
                    "Predictor_Display": "Age",
                    "bootstrap_coef_mean": -32.68,
                    "Selection_Freq_Pct": "100%",
                    "selection_frequency": 1.0,
                    "abs_bootstrap_coef_mean": 32.68,
                },
                {
                    "Overall_Rank": 2,
                    "Acronym": "RESTORE",
                    "Input_vars": 21,
                    "Model_label": "Model 9",
                    "Predictor_Display": "Age",
                    "bootstrap_coef_mean": -10.00,
                    "Selection_Freq_Pct": "90%",
                    "selection_frequency": 0.9,
                    "abs_bootstrap_coef_mean": 10.0,
                },
                {
                    "Overall_Rank": 1,
                    "Acronym": "RESTORE",
                    "Input_vars": 21,
                    "Model_label": "Model 5",
                    "Predictor_Display": "Unique",
                    "bootstrap_coef_mean": 5.0,
                    "Selection_Freq_Pct": "80%",
                    "selection_frequency": 0.8,
                    "abs_bootstrap_coef_mean": 5.0,
                },
            ]
        )

        result = tables_1840.build_compact_predictor_table(model_explainers, lasso_coefficients)

        self.assertEqual(result.columns.tolist(), ["Predictor", "1. RESTORE (n=21)", "2. RESTORE (n=21)"])
        self.assertEqual(result["Predictor"].tolist(), ["Age"])
        self.assertEqual(result.loc[0, "1. RESTORE (n=21)"], "-32.7; 100%")
        self.assertEqual(result.loc[0, "2. RESTORE (n=21)"], "-10.0; 90%")


if __name__ == "__main__":
    unittest.main()
