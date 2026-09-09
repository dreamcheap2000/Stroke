#!/usr/bin/env python3
"""
20260909_Comprehensive.py
=========================
Post-processing companion to 20260904_Comprehensive_1018.py.

This script reuses the retained five-model comprehensive analysis and adds:
  1. Youden-optimal cutpoints for the Part 2 binary classifiers.
  2. Bootstrap percentile confidence intervals around the cutpoints.
  3. Scenario-specific sensitivity, specificity, balanced accuracy, and accuracy.
  4. Publication-style model names/acronyms and ranked summary tables.

Outputs:
    20260909_Comprehensive.docx
    20260909_Comprehensive.xlsx
    20260909_Table1_2016.docx
    20260909_Publication_Tables.md
    20260909_Publication_Tables.tex
"""

from __future__ import annotations

import importlib.util
import math
from pathlib import Path

import numpy as np
import pandas as pd
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from sklearn.base import clone
from sklearn.metrics import accuracy_score, balanced_accuracy_score, confusion_matrix, roc_curve
from sklearn.model_selection import StratifiedKFold, cross_val_predict, cross_validate

ROOT = Path(__file__).resolve().parent
SOURCE_SCRIPT = ROOT / "20260904_Comprehensive_1018.py"
SOURCE_XLSX = ROOT / "20260904_Comprehensive_1018.xlsx"
OUTPUT_DOCX = ROOT / "20260909_Comprehensive.docx"
OUTPUT_XLSX = ROOT / "20260909_Comprehensive.xlsx"
OUTPUT_TABLE1 = ROOT / "20260909_Table1_2016.docx"
OUTPUT_PUBLICATION_MD = ROOT / "20260909_Publication_Tables.md"
OUTPUT_PUBLICATION_TEX = ROOT / "20260909_Publication_Tables.tex"
OUTPUT_PANEL_A_CSV = ROOT / "20260909_Publication_Table1_PanelA.csv"
OUTPUT_PANEL_B_CSV = ROOT / "20260909_Publication_Table1_PanelB.csv"
OUTPUT_TABLE2_CSV = ROOT / "20260909_Publication_Table2.csv"

RANDOM_STATE = 42
CV_FOLDS = 5
N_BOOTSTRAP_THRESHOLD = 2000

MODEL_BRANDING = {
    "Model 5": {
        "acronym": "RESTORE",
        "publication_name": "REhabilitation STroke Outcome Recovery Estimator",
        "focus": "Recovery trajectory using baseline function, gait speed, and early T1-to-T2 improvement among LOS ≥21 days.",
    },
    "Model 8": {
        "acronym": "COMPASS",
        "publication_name": "COMPrehensive Post-Acute Stroke Score",
        "focus": "Broad post-acute clinical and functional profile with imputed gait speed, without NIHSS discharge items.",
    },
    "Model 2": {
        "acronym": "CASCADE",
        "publication_name": "Clinical And Stroke Complexity Assessment for Discharge Endurance",
        "focus": "Full clinical burden model integrating stroke topology, comorbidities, complications, NIHSS, and function.",
    },
    "Model 1": {
        "acronym": "AIMS",
        "publication_name": "Admission Impairment Mobility Score",
        "focus": "Admission functional core using demographics, bedside disability, and imputed gait speed.",
    },
    "Model 4": {
        "acronym": "BEDSIDE",
        "publication_name": "Balance-Enhanced Demographic Speed Index for Discharge Endurance",
        "focus": "Ultra-compact bedside model built from age, sex, balance, and gait speed.",
    },
}

SCENARIO_ORDER = ["Best", "Worst"]
SCENARIO_COLUMNS = {
    "Best": "6MWT_Best_Scenario",
    "Worst": "6MWT_Worst_Scenario",
}
SCENARIO_NOTES = {
    "Best": "Optimistic non-completer walking classification scenario.",
    "Worst": "Conservative non-completer walking classification scenario.",
}

HEADER_FILL = "1F4E78"
SUBHEADER_FILL = "5B9BD5"
ALT_ROW_FILL = "EEF4FA"


def load_source_module():
    spec = importlib.util.spec_from_file_location("comp20260904", SOURCE_SCRIPT)
    if spec is None or spec.loader is None:
        raise ImportError(f"Cannot load source script: {SOURCE_SCRIPT}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def fmt_num(value: float, digits: int = 3) -> str:
    return f"{value:.{digits}f}" if math.isfinite(value) else "N/A"


def fmt_ci(point: float, lo: float, hi: float, digits: int = 3) -> str:
    if not all(math.isfinite(v) for v in [point, lo, hi]):
        return "N/A"
    return f"{point:.{digits}f} [{lo:.{digits}f}, {hi:.{digits}f}]"


def fmt_pct_ci(point: float, lo: float, hi: float, digits: int = 1) -> str:
    if not all(math.isfinite(v) for v in [point, lo, hi]):
        return "N/A"
    scale = 100
    return f"{point * scale:.{digits}f}% [{lo * scale:.{digits}f}, {hi * scale:.{digits}f}]"


def short_binary_model(name: str) -> str:
    mapping = {
        "LogisticRegression(class_weight='balanced', solver='liblinear')": "Logistic regression",
        "RandomForestClassifier(n_estimators=300, class_weight='balanced')": "Random forest",
        "ExtraTreesClassifier(n_estimators=300, class_weight='balanced')": "Extra trees",
    }
    return mapping.get(name, name)


def shade_row(row, fill: str) -> None:
    tr = row._tr
    tr_pr = tr.get_or_add_trPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tr_pr.append(shd)


def set_cell_text(cell, text: str, *, bold: bool = False, size: int = 8, color: str | None = None,
                  align: WD_ALIGN_PARAGRAPH = WD_ALIGN_PARAGRAPH.CENTER) -> None:
    cell.text = str(text)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    for para in cell.paragraphs:
        para.alignment = align
        for run in para.runs:
            run.font.size = Pt(size)
            run.font.bold = bold
            if color is not None:
                run.font.color.rgb = RGBColor.from_string(color)


def add_styled_table(doc: Document, rows: list[list[str]], *, title: str | None = None,
                     font_size: int = 8, landscape: bool = False,
                     first_col_left: bool = True) -> None:
    if title:
        p = doc.add_paragraph()
        p.add_run(title).bold = True
    if not rows:
        return
    if landscape:
        ensure_landscape(doc)
    table = doc.add_table(rows=len(rows), cols=len(rows[0]))
    table.style = "Table Grid"
    table.autofit = True
    for i, row_values in enumerate(rows):
        row = table.rows[i]
        if i == 0:
            shade_row(row, HEADER_FILL)
        elif i == 1 and any(str(v).startswith("__SUBHEADER__") for v in row_values):
            shade_row(row, SUBHEADER_FILL)
        elif i % 2 == 1:
            shade_row(row, ALT_ROW_FILL)
        for j, value in enumerate(row_values):
            display_value = str(value).replace("__SUBHEADER__", "")
            align = WD_ALIGN_PARAGRAPH.LEFT if first_col_left and j == 0 else WD_ALIGN_PARAGRAPH.CENTER
            set_cell_text(
                table.cell(i, j),
                display_value,
                bold=i in (0, 1),
                size=font_size,
                color="FFFFFF" if i in (0, 1) else None,
                align=align,
            )
    doc.add_paragraph()


def ensure_landscape(doc: Document) -> None:
    section = doc.sections[-1]
    if section.orientation != WD_ORIENT.LANDSCAPE:
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width, section.page_height = section.page_height, section.page_width
    section.left_margin = Inches(0.4)
    section.right_margin = Inches(0.4)
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)


def dataframe_to_markdown(df: pd.DataFrame) -> str:
    headers = [str(col) for col in df.columns]
    lines = [
        "| " + " | ".join(headers) + " |",
        "| " + " | ".join(["---"] * len(headers)) + " |",
    ]
    for row in df.fillna("—").astype(str).values.tolist():
        lines.append("| " + " | ".join(row) + " |")
    return "\n".join(lines)


def dataframe_to_latex(df: pd.DataFrame) -> str:
    def esc(value: str) -> str:
        return (
            value.replace("\\", "\\textbackslash{}")
            .replace("&", "\\&")
            .replace("%", "\\%")
            .replace("_", "\\_")
            .replace("#", "\\#")
        )

    headers = [esc(str(col)) for col in df.columns]
    lines = [
        "\\begin{tabular}{" + "l" * len(headers) + "}",
        "\\hline",
        " & ".join(headers) + " \\\\",
        "\\hline",
    ]
    for row in df.fillna("—").astype(str).values.tolist():
        lines.append(" & ".join(esc(value) for value in row) + " \\\\")
    lines.extend(["\\hline", "\\end{tabular}"])
    return "\n".join(lines)


def evaluate_threshold(y_true: np.ndarray, scores: np.ndarray, threshold: float) -> dict[str, float]:
    y_pred = (scores >= threshold).astype(int)
    tn, fp, fn, tp = confusion_matrix(y_true, y_pred, labels=[0, 1]).ravel()
    sensitivity = tp / (tp + fn) if (tp + fn) else float("nan")
    specificity = tn / (tn + fp) if (tn + fp) else float("nan")
    bal_acc = balanced_accuracy_score(y_true, y_pred)
    accuracy = accuracy_score(y_true, y_pred)
    return {
        "threshold": float(threshold),
        "sensitivity": float(sensitivity),
        "specificity": float(specificity),
        "balanced_accuracy": float(bal_acc),
        "accuracy": float(accuracy),
        "tp": int(tp),
        "fp": int(fp),
        "tn": int(tn),
        "fn": int(fn),
    }


def youden_optimal_metrics(y_true: np.ndarray, scores: np.ndarray) -> dict[str, float]:
    fpr, tpr, thresholds = roc_curve(y_true, scores)
    finite = np.isfinite(thresholds)
    fpr = fpr[finite]
    tpr = tpr[finite]
    thresholds = thresholds[finite]
    if len(thresholds) == 0:
        raise ValueError("No finite ROC thresholds were produced.")
    youden = tpr - fpr
    specificity = 1 - fpr
    balance_gap = np.abs(tpr - specificity)
    order = np.lexsort((
        np.abs(thresholds - 0.5),
        balance_gap,
        -specificity,
        -tpr,
        -youden,
    ))
    best_idx = int(order[0])
    metrics = evaluate_threshold(y_true, scores, float(thresholds[best_idx]))
    metrics["youden_j"] = float(youden[best_idx])
    return metrics


def bootstrap_youden_summary(y_true: np.ndarray, scores: np.ndarray,
                             *, n_bootstrap: int = N_BOOTSTRAP_THRESHOLD,
                             seed: int = RANDOM_STATE) -> dict[str, float]:
    point = youden_optimal_metrics(y_true, scores)
    rng = np.random.default_rng(seed)
    stats = {
        "threshold": [],
        "sensitivity": [],
        "specificity": [],
        "balanced_accuracy": [],
        "accuracy": [],
        "youden_j": [],
    }

    for _ in range(n_bootstrap):
        idx = rng.integers(0, len(y_true), size=len(y_true))
        y_boot = y_true[idx]
        if np.unique(y_boot).size < 2:
            continue
        boot_metrics = youden_optimal_metrics(y_boot, scores[idx])
        for key in stats:
            stats[key].append(boot_metrics[key])

    def ci(values: list[float]) -> tuple[float, float, float]:
        if not values:
            return float("nan"), float("nan"), float("nan")
        arr = np.asarray(values, dtype=float)
        return float(arr.mean()), float(np.percentile(arr, 2.5)), float(np.percentile(arr, 97.5))

    out = point.copy()
    for key, values in stats.items():
        mean_v, lo_v, hi_v = ci(values)
        out[f"{key}_boot_mean"] = mean_v
        out[f"{key}_ci_lo"] = lo_v
        out[f"{key}_ci_hi"] = hi_v
    out["bootstrap_valid_resamples"] = len(stats["threshold"])
    out["n_positive"] = int(y_true.sum())
    out["n_negative"] = int((1 - y_true).sum())
    out["prevalence"] = float(y_true.mean())
    return out


def analyse_binary_models(source, df: pd.DataFrame, features: list[str], scenario_name: str) -> tuple[dict, pd.DataFrame, pd.DataFrame]:
    scenario_col = SCENARIO_COLUMNS[scenario_name]
    model_df = df[df[scenario_col].notna()].copy()
    valid_features = source._filter_existing(features, model_df)
    X = model_df[valid_features]
    y = model_df[scenario_col].astype(int).to_numpy()
    cv = StratifiedKFold(n_splits=CV_FOLDS, shuffle=True, random_state=RANDOM_STATE)
    cv_splits = list(cv.split(X, y))

    rows: list[dict] = []
    coefficient_rows: list[dict] = []
    candidates = source._binary_candidates(valid_features)
    candidates = {name: pipe for name, pipe in candidates.items() if "LogisticRegression" in name}
    if not candidates:
        raise ValueError("No logistic regression candidate was found in source._binary_candidates().")
    for name, pipe in candidates.items():
        scores = cross_validate(
            clone(pipe),
            X,
            y,
            cv=cv_splits,
            n_jobs=-1,
            scoring=["balanced_accuracy", "accuracy", "f1"],
        )
        oof_prob = cross_val_predict(
            clone(pipe),
            X,
            y,
            cv=cv_splits,
            method="predict_proba",
            n_jobs=-1,
        )[:, 1]
        youden = bootstrap_youden_summary(y, oof_prob)
        fitted_pipe = clone(pipe)
        fitted_pipe.fit(X, y)
        fitted_model = fitted_pipe.named_steps["model"]
        coefficient_rows.append({
            "Predictor": "(Intercept)",
            "Coefficient": float(fitted_model.intercept_[0]),
            "Abs_Coefficient": abs(float(fitted_model.intercept_[0])),
        })
        for predictor, coef in zip(valid_features, fitted_model.coef_[0]):
            coefficient_rows.append({
                "Predictor": predictor,
                "Coefficient": float(coef),
                "Abs_Coefficient": abs(float(coef)),
            })
        rows.append({
            "Binary_model": name,
            "Binary_model_short": short_binary_model(name),
            "Default_balanced_accuracy": float(np.mean(scores["test_balanced_accuracy"])),
            "Default_accuracy": float(np.mean(scores["test_accuracy"])),
            "Default_f1": float(np.mean(scores["test_f1"])),
            "OOF_probability_mean": float(np.mean(oof_prob)),
            "Youden_threshold": youden["threshold"],
            "Threshold_CI_lo": youden["threshold_ci_lo"],
            "Threshold_CI_hi": youden["threshold_ci_hi"],
            "Sensitivity": youden["sensitivity"],
            "Sensitivity_CI_lo": youden["sensitivity_ci_lo"],
            "Sensitivity_CI_hi": youden["sensitivity_ci_hi"],
            "Specificity": youden["specificity"],
            "Specificity_CI_lo": youden["specificity_ci_lo"],
            "Specificity_CI_hi": youden["specificity_ci_hi"],
            "Balanced_Accuracy": youden["balanced_accuracy"],
            "Balanced_Accuracy_CI_lo": youden["balanced_accuracy_ci_lo"],
            "Balanced_Accuracy_CI_hi": youden["balanced_accuracy_ci_hi"],
            "Accuracy": youden["accuracy"],
            "Accuracy_CI_lo": youden["accuracy_ci_lo"],
            "Accuracy_CI_hi": youden["accuracy_ci_hi"],
            "Youden_J": youden["youden_j"],
            "Youden_J_CI_lo": youden["youden_j_ci_lo"],
            "Youden_J_CI_hi": youden["youden_j_ci_hi"],
            "Bootstrap_valid_resamples": youden["bootstrap_valid_resamples"],
            "N_positive": youden["n_positive"],
            "N_negative": youden["n_negative"],
            "Prevalence": youden["prevalence"],
        })

    leaderboard = pd.DataFrame(rows).sort_values(
        ["Default_balanced_accuracy", "Default_accuracy", "Default_f1"],
        ascending=[False, False, False],
    ).reset_index(drop=True)
    selected = leaderboard.iloc[0].to_dict()
    coefficients = pd.DataFrame(coefficient_rows)
    return selected, leaderboard, coefficients


def build_ranked_model_frame(source, lasso_summary: pd.DataFrame, ipcw_summary: pd.DataFrame) -> pd.DataFrame:
    ipcw_unique = (
        ipcw_summary.groupby("Model", as_index=False)[["Weighted_OOF_R2", "Weighted_OOF_MAE", "Binary_OOF_Bal_Acc"]]
        .mean()
        .rename(columns={"Model": "Model_label"})
    )
    lasso = lasso_summary.copy().rename(columns={"Model": "Model_full_name"})
    lasso["Model_label"] = lasso["Model_full_name"].str.split(":", n=1).str[0]
    ranked = lasso.merge(ipcw_unique, on="Model_label", how="left")
    ranked = ranked.sort_values(
        ["Weighted_OOF_R2", "CV_R2", "Weighted_OOF_MAE", "CV_MAE"],
        ascending=[False, False, True, True],
    ).reset_index(drop=True)
    ranked["Overall_Rank"] = np.arange(1, len(ranked) + 1)
    ranked["Acronym"] = ranked["Model_label"].map(lambda s: MODEL_BRANDING[s]["acronym"])
    ranked["Publication_Name"] = ranked["Model_label"].map(lambda s: MODEL_BRANDING[s]["publication_name"])
    ranked["Model_Focus"] = ranked["Model_label"].map(lambda s: MODEL_BRANDING[s]["focus"])
    ranked["Predictor_List"] = ranked["Model_full_name"].map(
        lambda name: ", ".join(source.MODEL_SPECS[name])
    )
    return ranked


def build_part1_coefficients(source, ranked_models: pd.DataFrame) -> pd.DataFrame:
    df = pd.read_excel(source.INPUT_XLSX)
    df, _ = source.apply_manual_patient_corrections(df)
    df, _ = source.audit_and_clean_t1t2_data(df)

    retained_model_specs = {
        model_name: source.MODEL_SPECS[model_name]
        for model_name in source.RETAINED_MODEL_ORDER
        if model_name in source.MODEL_SPECS
    }
    all_candidate_cols = list({c for cols in retained_model_specs.values() for c in cols})
    for col in all_candidate_cols + ["6MWT4"] + source.NIHSS_IN:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    coefficient_frames: list[pd.DataFrame] = []
    for i, model_row in ranked_models.reset_index(drop=True).iterrows():
        model_full_name = model_row["Model_full_name"]
        candidate_features = retained_model_specs.get(model_full_name, source.MODEL_SPECS[model_full_name])
        valid_features = source._filter_existing(candidate_features, df)
        if not valid_features:
            continue
        is_t1t2_model = any(col in candidate_features for col in source.T1T2_IMPROVEMENT)
        result = source.fit_bootstrap_lasso(
            df,
            valid_features,
            model_full_name,
            model_seed=source.RANDOM_STATE + i,
            model5_restrict=is_t1t2_model,
        )
        importance = result["importance"].copy()
        importance.insert(0, "Overall_Rank", int(model_row["Overall_Rank"]))
        importance.insert(1, "Model_label", model_row["Model_label"])
        importance.insert(2, "Acronym", model_row["Acronym"])
        importance.insert(3, "Publication_Name", model_row["Publication_Name"])
        importance.insert(4, "Input_vars", int(model_row["Input_vars"]))
        importance.insert(5, "Model_full_name", model_full_name)
        coefficient_frames.append(importance)

    if not coefficient_frames:
        return pd.DataFrame(columns=[
            "Overall_Rank", "Model_label", "Acronym", "Publication_Name", "Input_vars", "Model_full_name",
            "predictor", "base_coef", "bootstrap_coef_mean", "bootstrap_coef_sd",
            "bootstrap_coef_ci_lo", "bootstrap_coef_ci_hi", "selection_frequency",
            "abs_bootstrap_coef_mean",
        ])

    part1_df = pd.concat(coefficient_frames, ignore_index=True)
    part1_df["Stable"] = part1_df["selection_frequency"] >= source.STABILITY_THRESHOLD
    part1_df["Selection_Freq_Pct"] = (part1_df["selection_frequency"] * 100).round().astype(int).astype(str) + "%"
    part1_df["Coef_95_CI"] = part1_df.apply(
        lambda row: fmt_ci(
            float(row["bootstrap_coef_mean"]),
            float(row["bootstrap_coef_ci_lo"]),
            float(row["bootstrap_coef_ci_hi"]),
            2,
        ),
        axis=1,
    )
    part1_df["Predictor_Display"] = part1_df["predictor"].astype(str)
    return part1_df


def build_publication_tables(
    ranked_models: pd.DataFrame,
    scenario_df: pd.DataFrame,
    part1_df: pd.DataFrame,
    *,
    stability_threshold: float,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    perf_rows: list[dict] = []
    for _, model_row in ranked_models.iterrows():
        row = {
            "Rank": int(model_row["Overall_Rank"]),
            "Acronym": model_row["Acronym"],
            "Model": model_row["Model_label"],
            "Publication name": model_row["Publication_Name"],
            "Input vars": int(model_row["Input_vars"]),
        }
        model_slice = scenario_df[scenario_df["Model_label"] == model_row["Model_label"]]
        for scenario_name, prefix in [("Best", "Best"), ("Worst", "Worst")]:
            scenario_match = model_slice[model_slice["Scenario"] == scenario_name].iloc[0]
            row[f"{prefix} Acc [95% CI]"] = fmt_pct_ci(
                float(scenario_match["Accuracy"]),
                float(scenario_match["Accuracy_CI_lo"]),
                float(scenario_match["Accuracy_CI_hi"]),
            )
            row[f"{prefix} BalAcc [95% CI]"] = fmt_pct_ci(
                float(scenario_match["Balanced_Accuracy"]),
                float(scenario_match["Balanced_Accuracy_CI_lo"]),
                float(scenario_match["Balanced_Accuracy_CI_hi"]),
            )
            row[f"{prefix} Se [95% CI]"] = fmt_pct_ci(
                float(scenario_match["Sensitivity"]),
                float(scenario_match["Sensitivity_CI_lo"]),
                float(scenario_match["Sensitivity_CI_hi"]),
            )
            row[f"{prefix} Sp [95% CI]"] = fmt_pct_ci(
                float(scenario_match["Specificity"]),
                float(scenario_match["Specificity_CI_lo"]),
                float(scenario_match["Specificity_CI_hi"]),
            )
        perf_rows.append(row)
    panel_a = pd.DataFrame(perf_rows)

    stable = part1_df[part1_df["selection_frequency"] >= stability_threshold].copy()
    stable["Model_Column"] = stable.apply(
        lambda row: f"{row['Acronym']} ({row['Model_label']}; {int(row['Input_vars'])} vars)",
        axis=1,
    )
    stable["Cell"] = stable.apply(
        lambda row: (
            f"{float(row['bootstrap_coef_mean']):+.2f} "
            f"({float(row['bootstrap_coef_sd']):.2f}) "
            f"[{float(row['bootstrap_coef_ci_lo']):.2f}, {float(row['bootstrap_coef_ci_hi']):.2f}]; "
            f"{row['Selection_Freq_Pct']}"
        ),
        axis=1,
    )
    predictor_order = (
        stable.groupby("Predictor_Display", as_index=False)
        .agg(
            Model_Count=("Model_label", "nunique"),
            Max_SF=("selection_frequency", "max"),
            Max_Abs=("abs_bootstrap_coef_mean", "max"),
        )
        .sort_values(["Model_Count", "Max_SF", "Max_Abs", "Predictor_Display"], ascending=[False, False, False, True])
    )
    column_order = [
        f"{row['Acronym']} ({row['Model_label']}; {int(row['Input_vars'])} vars)"
        for _, row in ranked_models.iterrows()
    ]
    panel_b = (
        stable.pivot_table(index="Predictor_Display", columns="Model_Column", values="Cell", aggfunc="first")
        .reindex(predictor_order["Predictor_Display"].tolist())
        .reindex(columns=column_order)
        .reset_index()
        .rename(columns={"Predictor_Display": "Predictor"})
        .fillna("—")
    )

    table2 = (
        stable[[
            "Overall_Rank", "Acronym", "Model_label", "Publication_Name", "Input_vars", "Predictor_Display",
            "base_coef", "bootstrap_coef_mean", "bootstrap_coef_sd", "bootstrap_coef_ci_lo",
            "bootstrap_coef_ci_hi", "Selection_Freq_Pct",
        ]]
        .rename(columns={
            "Overall_Rank": "Rank",
            "Publication_Name": "Publication name",
            "Input_vars": "Input vars",
            "Predictor_Display": "Predictor",
            "base_coef": "Full-fit Coef",
            "bootstrap_coef_mean": "Boot Mean Coef",
            "bootstrap_coef_sd": "Boot SD",
            "bootstrap_coef_ci_lo": "CI lo",
            "bootstrap_coef_ci_hi": "CI hi",
            "Selection_Freq_Pct": "Sel Freq",
        })
        .sort_values(["Rank", "Acronym", "Predictor"])
        .reset_index(drop=True)
    )
    table2["95% CI"] = table2.apply(
        lambda row: f"[{float(row['CI lo']):.2f}, {float(row['CI hi']):.2f}]",
        axis=1,
    )
    table2["Full-fit Coef"] = table2["Full-fit Coef"].map(lambda x: f"{float(x):+.2f}")
    table2["Boot Mean Coef"] = table2["Boot Mean Coef"].map(lambda x: f"{float(x):+.2f}")
    table2["Boot SD"] = table2["Boot SD"].map(lambda x: f"{float(x):.2f}")
    table2 = table2[[
        "Rank", "Acronym", "Model_label", "Publication name", "Input vars",
        "Predictor", "Full-fit Coef", "Boot Mean Coef", "Boot SD", "95% CI", "Sel Freq",
    ]]
    return panel_a, panel_b, table2


def build_outputs(source) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    predictions_df = pd.read_excel(SOURCE_XLSX, sheet_name="Predictions")
    lasso_summary = pd.read_excel(SOURCE_XLSX, sheet_name="LASSO_Summary")
    ipcw_summary = pd.read_excel(SOURCE_XLSX, sheet_name="IPCW_Summary")

    ranked_models = build_ranked_model_frame(source, lasso_summary, ipcw_summary)
    scenario_rows: list[dict] = []
    leaderboard_rows: list[pd.DataFrame] = []
    coefficient_frames: list[pd.DataFrame] = []

    for _, model_row in ranked_models.iterrows():
        model_full_name = model_row["Model_full_name"]
        features = source.MODEL_SPECS[model_full_name]
        for scenario_name in SCENARIO_ORDER:
            selected, leaderboard, coefficients = analyse_binary_models(source, predictions_df, features, scenario_name)
            leaderboard = leaderboard.copy()
            leaderboard.insert(0, "Scenario", scenario_name)
            leaderboard.insert(0, "Model_label", model_row["Model_label"])
            leaderboard.insert(0, "Acronym", model_row["Acronym"])
            leaderboard_rows.append(leaderboard)
            coefficients = coefficients.copy()
            coefficients.insert(0, "Scenario", scenario_name)
            coefficients.insert(0, "Model_label", model_row["Model_label"])
            coefficients.insert(0, "Acronym", model_row["Acronym"])
            coefficient_frames.append(coefficients)

            scenario_rows.append({
                "Overall_Rank": int(model_row["Overall_Rank"]),
                "Model_label": model_row["Model_label"],
                "Model_full_name": model_full_name,
                "Acronym": model_row["Acronym"],
                "Publication_Name": model_row["Publication_Name"],
                "Model_Focus": model_row["Model_Focus"],
                "Scenario": scenario_name,
                "Scenario_note": SCENARIO_NOTES[scenario_name],
                "Predictor_count": int(model_row["Input_vars"]),
                "Predictor_categories_summary": model_row["Predictor_categories_summary"],
                "LASSO_CV_R2": float(model_row["CV_R2"]),
                "LASSO_CV_MAE": float(model_row["CV_MAE"]),
                "IPCW_Weighted_OOF_R2": float(model_row["Weighted_OOF_R2"]),
                "IPCW_Weighted_OOF_MAE": float(model_row["Weighted_OOF_MAE"]),
                "Source_binary_OOF_bal_acc": float(
                    ipcw_summary.loc[
                        (ipcw_summary["Model"] == model_row["Model_label"])
                        & (ipcw_summary["Scenario"] == scenario_name),
                        "Binary_OOF_Bal_Acc",
                    ].iloc[0]
                ),
                **selected,
            })

    scenario_df = pd.DataFrame(scenario_rows).sort_values(
        ["Scenario", "Overall_Rank"], ascending=[True, True]
    ).reset_index(drop=True)
    leaderboard_df = pd.concat(leaderboard_rows, ignore_index=True)
    coefficients_df = pd.concat(coefficient_frames, ignore_index=True)
    part1_coefficients_df = build_part1_coefficients(source, ranked_models)
    explainers_df = ranked_models[[
        "Overall_Rank",
        "Model_label",
        "Acronym",
        "Publication_Name",
        "Model_Focus",
        "Predictor_categories_summary",
        "Predictor_List",
        "Input_vars",
        "CV_R2",
        "CV_MAE",
        "Weighted_OOF_R2",
        "Weighted_OOF_MAE",
    ]].copy()
    panel_a_df, panel_b_df, table2_df = build_publication_tables(
        ranked_models,
        scenario_df,
        part1_coefficients_df,
        stability_threshold=source.STABILITY_THRESHOLD,
    )
    return (
        ranked_models,
        scenario_df,
        leaderboard_df,
        explainers_df,
        coefficients_df,
        part1_coefficients_df,
        panel_a_df,
        panel_b_df,
        table2_df,
    )


def write_excel(ranked_models: pd.DataFrame, scenario_df: pd.DataFrame,
                leaderboard_df: pd.DataFrame, explainers_df: pd.DataFrame,
                coefficients_df: pd.DataFrame, part1_coefficients_df: pd.DataFrame,
                panel_a_df: pd.DataFrame, panel_b_df: pd.DataFrame,
                table2_df: pd.DataFrame) -> None:
    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        ranked_models.to_excel(writer, sheet_name="Model_Ranking", index=False)
        scenario_df.to_excel(writer, sheet_name="Scenario_Performance", index=False)
        leaderboard_df.to_excel(writer, sheet_name="Binary_Leaderboards", index=False)
        explainers_df.to_excel(writer, sheet_name="Model_Explainers", index=False)
        coefficients_df.to_excel(writer, sheet_name="Logistic_Coefficients", index=False)
        part1_coefficients_df.to_excel(writer, sheet_name="LASSO_Coefficients", index=False)
        panel_a_df.to_excel(writer, sheet_name="Pub_Table1_PanelA", index=False)
        panel_b_df.to_excel(writer, sheet_name="Pub_Table1_PanelB", index=False)
        table2_df.to_excel(writer, sheet_name="Pub_Table2", index=False)


def write_publication_files(panel_a_df: pd.DataFrame, panel_b_df: pd.DataFrame, table2_df: pd.DataFrame) -> None:
    panel_a_df.to_csv(OUTPUT_PANEL_A_CSV, index=False)
    panel_b_df.to_csv(OUTPUT_PANEL_B_CSV, index=False)
    table2_df.to_csv(OUTPUT_TABLE2_CSV, index=False)

    markdown_parts = [
        "## Table 1. Retained 20260909 models: binary performance and stable Part 1 predictors",
        "",
        "### Panel A. Part 2 IPCW binary classification performance",
        "",
        dataframe_to_markdown(panel_a_df),
        "",
        "### Panel B. Part 1 stable bootstrap LASSO coefficients (selection frequency ≥70%)",
        "",
        dataframe_to_markdown(panel_b_df),
        "",
        "## Table 2. Detailed stable Part 1 coefficients",
        "",
        dataframe_to_markdown(table2_df),
        "",
        "*β values are standardized bootstrap LASSO coefficients. In Panel B, cells are shown as Boot Mean Coef (Boot SD) [95% CI]; selection frequency.*",
        "*Acc = accuracy; BalAcc = balanced accuracy; Se = sensitivity; Sp = specificity; CI = 95% bootstrap confidence interval.*",
    ]
    OUTPUT_PUBLICATION_MD.write_text("\n".join(markdown_parts), encoding="utf-8")

    latex_parts = [
        "% Table 1, Panel A",
        dataframe_to_latex(panel_a_df),
        "",
        "% Table 1, Panel B",
        dataframe_to_latex(panel_b_df),
        "",
        "% Table 2",
        dataframe_to_latex(table2_df),
        "",
    ]
    OUTPUT_PUBLICATION_TEX.write_text("\n".join(latex_parts), encoding="utf-8")


def dataframe_to_rows(df: pd.DataFrame) -> list[list[str]]:
    rows = [list(df.columns)]
    rows.extend(df.fillna("—").astype(str).values.tolist())
    return rows


def write_table1_docx(panel_a_df: pd.DataFrame, panel_b_df: pd.DataFrame) -> None:
    doc = Document()
    ensure_landscape(doc)
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("Table 1. Retained 20260909 stroke PAC models: binary performance and stable Part 1 predictors")
    run.bold = True
    run.font.size = Pt(11)

    subtitle = doc.add_paragraph(
        "Panel A summarizes Part 2 IPCW binary classification performance across the retained 20260909 models. "
        "Panel B shows Part 1 stable bootstrap LASSO coefficients carried forward from the retained 20260904 models, "
        "using the same 20260909 ranks, model labels, publication names, and acronyms."
    )
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in subtitle.runs:
        run.font.size = Pt(8)

    add_styled_table(
        doc,
        dataframe_to_rows(panel_a_df),
        title="Panel A. Part 2 IPCW binary classification performance",
        font_size=7,
        landscape=True,
    )
    add_styled_table(
        doc,
        dataframe_to_rows(panel_b_df),
        title="Panel B. Part 1 stable bootstrap LASSO coefficients (selection frequency ≥70%)",
        font_size=7,
        landscape=True,
    )
    legend = doc.add_paragraph()
    legend.add_run("Abbreviations. ").bold = True
    legend.add_run(
        "PAC = post-acute care; IPCW = inverse probability of completion weighting; "
        "CI = percentile bootstrap 95% confidence interval; Se = sensitivity; Sp = specificity; "
        "BalAcc = balanced accuracy; NIHSS = National Institutes of Health Stroke Scale. "
        "Panel B cells are formatted as Boot Mean Coef (Boot SD) [95% CI]; selection frequency."
    )
    for run in legend.runs:
        run.font.size = Pt(8)
    doc.save(OUTPUT_TABLE1)


def write_comprehensive_docx(ranked_models: pd.DataFrame, scenario_df: pd.DataFrame,
                             coefficients_df: pd.DataFrame, part1_coefficients_df: pd.DataFrame,
                             panel_a_df: pd.DataFrame, panel_b_df: pd.DataFrame) -> None:
    doc = Document()
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("2026-09-09 Comprehensive Stroke PAC Model Update")
    run.bold = True
    run.font.size = Pt(12)

    intro = doc.add_paragraph(
        "This report extends the retained five-model 2026-09-04 comprehensive analysis. "
        "The cohort is a post-acute care stroke rehabilitation cohort initially hospitalized in a neurology ward, "
        "where stroke topology, comorbidities, complications, and NIHSS were recorded. After transfer to the PAC "
        "rehabilitation ward, functional assessments were recorded, including gait speed and the 6-minute walk test."
    )
    for run in intro.runs:
        run.font.size = Pt(9)

    methods = doc.add_paragraph(
        "Part 1 stable predictor summaries were re-derived from the retained 20260904 bootstrap LASSO models so that "
        "all tables inherit the 20260909 ranks, model labels, publication names, and acronyms. "
        "Part 2 logistic-regression classifiers were re-evaluated using out-of-fold predicted probabilities. "
        "For each retained model and each scenario (Best and Worst), a Youden-optimal cutpoint was then derived from "
        "the out-of-fold ROC curve, and bootstrap percentile confidence intervals around the cutpoint and its "
        "operating characteristics were estimated from 2,000 resamples."
    )
    for run in methods.runs:
        run.font.size = Pt(9)

    ranking_rows = [[
        "Rank", "Acronym", "Publication name", "Source model", "Input vars",
        "Step 1 CV R²", "Step 1 CV MAE", "Part 2 weighted OOF R²", "Part 2 weighted OOF MAE",
    ]]
    for _, row in ranked_models.iterrows():
        ranking_rows.append([
            str(int(row["Overall_Rank"])),
            row["Acronym"],
            row["Publication_Name"],
            row["Model_label"],
            str(int(row["Input_vars"])),
            fmt_num(row["CV_R2"], 4),
            fmt_num(row["CV_MAE"], 2),
            fmt_num(row["Weighted_OOF_R2"], 4),
            fmt_num(row["Weighted_OOF_MAE"], 1),
        ])
    add_styled_table(doc, ranking_rows, title="Overall ranking of the retained five models", font_size=8)
    add_styled_table(
        doc,
        dataframe_to_rows(panel_a_df),
        title="Publication-ready summary: Part 2 IPCW binary performance",
        font_size=7,
        landscape=True,
    )
    add_styled_table(
        doc,
        dataframe_to_rows(panel_b_df),
        title="Publication-ready summary: Part 1 stable bootstrap LASSO coefficients",
        font_size=7,
        landscape=True,
    )

    for scenario_name in SCENARIO_ORDER:
        scenario_rows = [[
            "Rank", "Acronym", "Cutpoint [95% CI]",
            "Sensitivity [95% CI]", "Specificity [95% CI]", "Balanced accuracy [95% CI]",
            "Accuracy [95% CI]", "Bootstrap n",
        ]]
        scenario_slice = scenario_df[scenario_df["Scenario"] == scenario_name]
        for _, row in scenario_slice.iterrows():
            scenario_rows.append([
                str(int(row["Overall_Rank"])),
                row["Acronym"],
                fmt_ci(row["Youden_threshold"], row["Threshold_CI_lo"], row["Threshold_CI_hi"], 3),
                fmt_ci(row["Sensitivity"], row["Sensitivity_CI_lo"], row["Sensitivity_CI_hi"], 3),
                fmt_ci(row["Specificity"], row["Specificity_CI_lo"], row["Specificity_CI_hi"], 3),
                fmt_ci(row["Balanced_Accuracy"], row["Balanced_Accuracy_CI_lo"], row["Balanced_Accuracy_CI_hi"], 3),
                fmt_ci(row["Accuracy"], row["Accuracy_CI_lo"], row["Accuracy_CI_hi"], 3),
                str(int(row["Bootstrap_valid_resamples"])),
            ])
        add_styled_table(
            doc,
            scenario_rows,
            title=f"{scenario_name} scenario binary operating characteristics",
            font_size=8,
            landscape=True,
        )

    for _, row in ranked_models.iterrows():
        section_head = doc.add_paragraph()
        section_head.add_run(
            f"Rank {int(row['Overall_Rank'])}. {row['Acronym']} — {row['Publication_Name']}"
        ).bold = True

        focus = doc.add_paragraph(f"Model focus: {row['Model_Focus']}")
        predictors = doc.add_paragraph(
            f"Predictors ({int(row['Input_vars'])}): {row['Predictor_categories_summary']}. "
            f"Full list: {row['Predictor_List']}"
        )
        performance = doc.add_paragraph(
            f"Step 1 performance: CV R² {fmt_num(row['CV_R2'], 4)} and CV MAE {fmt_num(row['CV_MAE'], 2)} m. "
            f"Part 2 weighted OOF regression performance: R² {fmt_num(row['Weighted_OOF_R2'], 4)} and "
            f"MAE {fmt_num(row['Weighted_OOF_MAE'], 1)} m."
        )
        for paragraph in [focus, predictors, performance]:
            for run in paragraph.runs:
                run.font.size = Pt(9)

        model_part1 = (
            part1_coefficients_df[
                (part1_coefficients_df["Model_label"] == row["Model_label"])
                & (part1_coefficients_df["Stable"])
            ]
            .sort_values(["selection_frequency", "abs_bootstrap_coef_mean"], ascending=[False, False])
            .reset_index(drop=True)
        )
        part1_rows = [[
            "Predictor", "Full-fit Coef", "Boot Mean Coef", "Boot SD", "95% CI", "Selection Freq",
        ]]
        for _, coef_row in model_part1.iterrows():
            part1_rows.append([
                coef_row["Predictor_Display"],
                f"{float(coef_row['base_coef']):+.4f}",
                f"{float(coef_row['bootstrap_coef_mean']):+.4f}",
                f"{float(coef_row['bootstrap_coef_sd']):.4f}",
                f"[{float(coef_row['bootstrap_coef_ci_lo']):+.4f}, {float(coef_row['bootstrap_coef_ci_hi']):+.4f}]",
                coef_row["Selection_Freq_Pct"],
            ])
        if len(part1_rows) == 1:
            part1_rows.append(["None met stability threshold", "—", "—", "—", "—", "—"])
        add_styled_table(
            doc,
            part1_rows,
            title="Part 1 stable bootstrap LASSO coefficients (selection frequency ≥70%)",
            font_size=8,
            landscape=True,
        )

        this_model = scenario_df[scenario_df["Model_label"] == row["Model_label"]]
        op_rows = [[
            "Scenario", "Cutpoint [95% CI]", "Se [95% CI]",
            "Sp [95% CI]", "BalAcc [95% CI]", "Acc [95% CI]",
        ]]
        for _, op in this_model.iterrows():
            op_rows.append([
                op["Scenario"],
                fmt_ci(op["Youden_threshold"], op["Threshold_CI_lo"], op["Threshold_CI_hi"], 3),
                fmt_ci(op["Sensitivity"], op["Sensitivity_CI_lo"], op["Sensitivity_CI_hi"], 3),
                fmt_ci(op["Specificity"], op["Specificity_CI_lo"], op["Specificity_CI_hi"], 3),
                fmt_ci(op["Balanced_Accuracy"], op["Balanced_Accuracy_CI_lo"], op["Balanced_Accuracy_CI_hi"], 3),
                fmt_ci(op["Accuracy"], op["Accuracy_CI_lo"], op["Accuracy_CI_hi"], 3),
            ])
        add_styled_table(doc, op_rows, font_size=8)

        model_coef = coefficients_df[coefficients_df["Model_label"] == row["Model_label"]].copy()
        coef_pivot = (
            model_coef
            .pivot_table(index="Predictor", columns="Scenario", values="Coefficient", aggfunc="first")
            .reset_index()
        )
        coef_abs = (
            model_coef.groupby("Predictor", as_index=False)["Abs_Coefficient"].max()
            .rename(columns={"Abs_Coefficient": "MaxAbs"})
        )
        coef_pivot = coef_pivot.merge(coef_abs, on="Predictor", how="left")
        coef_pivot = coef_pivot.sort_values("MaxAbs", ascending=False)
        coef_rows = [["Predictor", "Best coefficient", "Worst coefficient"]]
        for _, coef_row in coef_pivot.iterrows():
            coef_rows.append([
                coef_row["Predictor"],
                fmt_num(float(coef_row["Best"]), 4) if "Best" in coef_pivot.columns and math.isfinite(float(coef_row["Best"])) else "N/A",
                fmt_num(float(coef_row["Worst"]), 4) if "Worst" in coef_pivot.columns and math.isfinite(float(coef_row["Worst"])) else "N/A",
            ])
        add_styled_table(doc, coef_rows, title="Logistic regression coefficients (Best vs Worst)", font_size=8)

    reproducibility = doc.add_paragraph(
        "Reproducibility: run `python 20260909_Comprehensive.py` from the repository root clone to regenerate "
        "20260909_Comprehensive.docx, 20260909_Comprehensive.xlsx, 20260909_Table1_2016.docx, "
        "20260909_Publication_Tables.md, and 20260909_Publication_Tables.tex."
    )
    for run in reproducibility.runs:
        run.font.size = Pt(8)
    doc.save(OUTPUT_DOCX)


def main() -> None:
    if not SOURCE_SCRIPT.exists():
        raise FileNotFoundError(f"Missing source script: {SOURCE_SCRIPT}")
    if not SOURCE_XLSX.exists():
        raise FileNotFoundError(f"Missing source workbook: {SOURCE_XLSX}")

    source = load_source_module()
    (
        ranked_models,
        scenario_df,
        leaderboard_df,
        explainers_df,
        coefficients_df,
        part1_coefficients_df,
        panel_a_df,
        panel_b_df,
        table2_df,
    ) = build_outputs(source)
    write_excel(
        ranked_models,
        scenario_df,
        leaderboard_df,
        explainers_df,
        coefficients_df,
        part1_coefficients_df,
        panel_a_df,
        panel_b_df,
        table2_df,
    )
    write_publication_files(panel_a_df, panel_b_df, table2_df)
    write_table1_docx(panel_a_df, panel_b_df)
    write_comprehensive_docx(
        ranked_models,
        scenario_df,
        coefficients_df,
        part1_coefficients_df,
        panel_a_df,
        panel_b_df,
    )

    print(f"Saved: {OUTPUT_XLSX.name}")
    print(f"Saved: {OUTPUT_TABLE1.name}")
    print(f"Saved: {OUTPUT_DOCX.name}")
    print(f"Saved: {OUTPUT_PUBLICATION_MD.name}")
    print(f"Saved: {OUTPUT_PUBLICATION_TEX.name}")


if __name__ == "__main__":
    main()
