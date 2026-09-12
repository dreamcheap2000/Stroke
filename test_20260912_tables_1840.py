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
    def test_build_compact_performance_table_orders_clinically_and_formats_metrics(self):
        model_explainers = pd.DataFrame(
            [
                {
                    "Overall_Rank": 1,
                    "Model_label": "Model 5",
                    "Acronym": "RESTORE",
                    "Publication_Name": "Model One",
                    "Input_vars": 21,
                    "CV_R2": 0.54321,
                    "CV_MAE": 89.285,
                    "Weighted_OOF_R2": 0.51234,
                    "Weighted_OOF_MAE": 90.44,
                },
                {
                    "Overall_Rank": 4,
                    "Model_label": "Model 1",
                    "Acronym": "AIMS",
                    "Publication_Name": "Model Two",
                    "Input_vars": 12,
                    "CV_R2": None,
                    "CV_MAE": None,
                    "Weighted_OOF_R2": 0.50001,
                    "Weighted_OOF_MAE": 95.00,
                },
                {
                    "Overall_Rank": 5,
                    "Model_label": "Model 4",
                    "Acronym": "BEDSIDE",
                    "Publication_Name": "Model Three",
                    "Input_vars": 4,
                    "CV_R2": 0.49001,
                    "CV_MAE": 96.00,
                    "Weighted_OOF_R2": 0.48000,
                    "Weighted_OOF_MAE": 96.10,
                },
            ]
        )
        panel_a = pd.DataFrame(
            [
                {
                    "Rank": 1,
                    "Acronym": "RESTORE",
                    "Model": "Model 5",
                    "Publication name": "Model One",
                    "Input vars": 21,
                    "Best BalAcc [95% CI]": "80.0% [70.0, 90.0]",
                    "Worst BalAcc [95% CI]": "78.0% [68.0, 88.0]",
                },
                {
                    "Rank": 4,
                    "Acronym": "AIMS",
                    "Model": "Model 1",
                    "Publication name": "Model Two",
                    "Input vars": 12,
                    "Best BalAcc [95% CI]": "75.0% [65.0, 85.0]",
                    "Worst BalAcc [95% CI]": "73.0% [63.0, 83.0]",
                },
                {
                    "Rank": 5,
                    "Acronym": "BEDSIDE",
                    "Model": "Model 4",
                    "Publication name": "Model Three",
                    "Input vars": 4,
                    "Best BalAcc [95% CI]": "74.0% [64.0, 84.0]",
                    "Worst BalAcc [95% CI]": "72.0% [62.0, 82.0]",
                },
            ]
        )
        calibration = pd.DataFrame(
            [
                {"Model_label": "Model 5", "Acronym": "RESTORE", "Calibration_Intercept": 10.2, "Calibration_Slope": 0.9653},
                {"Model_label": "Model 1", "Acronym": "AIMS", "Calibration_Intercept": None, "Calibration_Slope": 0.9724},
                {"Model_label": "Model 4", "Acronym": "BEDSIDE", "Calibration_Intercept": 1.2, "Calibration_Slope": 0.9969},
            ]
        )

        result = tables_1840.build_compact_performance_table(model_explainers, panel_a, calibration)

        self.assertEqual(result["Acronym"].tolist(), ["BEDSIDE", "AIMS", "RESTORE"])
        self.assertEqual(result.loc[0, "Clinical role"], "Primary recommendation")
        self.assertEqual(result.loc[0, "Part 1 CV R²"], "0.4900")
        self.assertEqual(result.loc[1, "Part 1 CV R²"], "—")
        self.assertEqual(result.loc[2, "Part 2 weighted OOF R²"], "0.5123")
        self.assertEqual(result.loc[0, "Δ vs RESTORE R²"], "-0.0323")
        self.assertEqual(result.loc[0, "Δ vs RESTORE MAE (m)"], "+5.7")
        self.assertEqual(result.loc[1, "Calibration intercept (m)"], "—")
        self.assertEqual(result.loc[0, "Calibration slope"], "0.997")

    def test_build_compact_performance_table_handles_missing_restore_reference(self):
        model_explainers = pd.DataFrame(
            [
                {
                    "Overall_Rank": 6,
                    "Model_label": "Model 99",
                    "Acronym": "ZETA",
                    "Publication_Name": "Model Two",
                    "Input_vars": 12,
                    "CV_R2": 0.49001,
                    "CV_MAE": 95.00,
                    "Weighted_OOF_R2": 0.50001,
                    "Weighted_OOF_MAE": 95.00,
                }
            ]
        )
        panel_a = pd.DataFrame(
            [
                {
                    "Rank": 6,
                    "Acronym": "ZETA",
                    "Model": "Model 99",
                    "Publication name": "Model Two",
                    "Input vars": 12,
                    "Best BalAcc [95% CI]": "75.0% [65.0, 85.0]",
                    "Worst BalAcc [95% CI]": "73.0% [63.0, 83.0]",
                }
            ]
        )
        calibration = pd.DataFrame(
            [
                {"Model_label": "Model 99", "Calibration_Intercept": 0.5, "Calibration_Slope": 0.99},
            ]
        )

        result = tables_1840.build_compact_performance_table(model_explainers, panel_a, calibration)

        self.assertEqual(result.loc[0, "Clinical role"], "Additional comparison")
        self.assertEqual(result.loc[0, "Calibration intercept (m)"], "0.5")
        self.assertEqual(result.loc[0, "Δ vs RESTORE R²"], "—")
        self.assertEqual(result.loc[0, "Δ vs RESTORE MAE (m)"], "—")

    def test_build_compact_performance_table_uses_model_identifier_for_merge(self):
        model_explainers = pd.DataFrame(
            [
                {
                    "Overall_Rank": 1,
                    "Model_label": "Model 5",
                    "Acronym": "RESTORE",
                    "Publication_Name": "Expected Name",
                    "Input_vars": 21,
                    "CV_R2": 0.54321,
                    "CV_MAE": 89.285,
                    "Weighted_OOF_R2": 0.51234,
                    "Weighted_OOF_MAE": 90.44,
                }
            ]
        )
        panel_a = pd.DataFrame(
            [
                {
                    "Rank": 1,
                    "Acronym": "RESTORE",
                    "Model": "Model 5",
                    "Publication name": "Wrong Name",
                    "Input vars": 21,
                    "Best BalAcc [95% CI]": "80.0% [70.0, 90.0]",
                    "Worst BalAcc [95% CI]": "78.0% [68.0, 88.0]",
                }
            ]
        )
        calibration = pd.DataFrame(
            [
                {"Model_label": "Model 5", "Acronym": "RESTORE", "Calibration_Intercept": 1.0, "Calibration_Slope": 0.95},
            ]
        )

        result = tables_1840.build_compact_performance_table(model_explainers, panel_a, calibration)

        self.assertEqual(result.loc[0, "Part 1 CV R²"], "0.5432")
        self.assertEqual(result.loc[0, "Part 2 weighted OOF MAE (m)"], "90.4")
        self.assertEqual(result.loc[0, "Calibration intercept (m)"], "1.0")
        self.assertEqual(result.loc[0, "Best BalAcc [95% CI]"], "80.0% [70.0, 90.0]")

    def test_build_compact_predictor_table_keeps_shared_predictors_only_and_reorders_models(self):
        model_explainers = pd.DataFrame(
            [
                {"Overall_Rank": 1, "Acronym": "RESTORE", "Input_vars": 21},
                {"Overall_Rank": 4, "Acronym": "AIMS", "Input_vars": 12},
                {"Overall_Rank": 5, "Acronym": "BEDSIDE", "Input_vars": 4},
                {"Overall_Rank": 6, "Acronym": "ZETA", "Input_vars": 8},
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
                    "Overall_Rank": 4,
                    "Acronym": "AIMS",
                    "Input_vars": 12,
                    "Model_label": "Model 1",
                    "Predictor_Display": "Age",
                    "bootstrap_coef_mean": -10.00,
                    "Selection_Freq_Pct": "90%",
                    "selection_frequency": 0.9,
                    "abs_bootstrap_coef_mean": 10.0,
                },
                {
                    "Overall_Rank": 5,
                    "Acronym": "BEDSIDE",
                    "Input_vars": 4,
                    "Model_label": "Model 4",
                    "Predictor_Display": "Age",
                    "bootstrap_coef_mean": -8.50,
                    "Selection_Freq_Pct": "100%",
                    "selection_frequency": 1.0,
                    "abs_bootstrap_coef_mean": 8.5,
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
                {
                    "Overall_Rank": 6,
                    "Acronym": "ZETA",
                    "Input_vars": 8,
                    "Model_label": "Model 99",
                    "Predictor_Display": "Age",
                    "bootstrap_coef_mean": -6.0,
                    "Selection_Freq_Pct": "75%",
                    "selection_frequency": 0.75,
                    "abs_bootstrap_coef_mean": 6.0,
                },
            ]
        )

        result = tables_1840.build_compact_predictor_table(model_explainers, lasso_coefficients)

        self.assertEqual(
            result.columns.tolist(),
            ["Predictor", "BEDSIDE (n=4)", "AIMS (n=12)", "RESTORE (n=21)", "ZETA (n=8)"],
        )
        self.assertEqual(result["Predictor"].tolist(), ["Age"])
        self.assertEqual(result.loc[0, "BEDSIDE (n=4)"], "-8.5; 100%")
        self.assertEqual(result.loc[0, "AIMS (n=12)"], "-10.0; 90%")
        self.assertEqual(result.loc[0, "RESTORE (n=21)"], "-32.7; 100%")
        self.assertEqual(result.loc[0, "ZETA (n=8)"], "-6.0; 75%")


if __name__ == "__main__":
    unittest.main()
