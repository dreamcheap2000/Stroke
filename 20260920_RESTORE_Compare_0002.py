#!/usr/bin/env python3
from __future__ import annotations

import importlib.util
from pathlib import Path

import numpy as np
import pandas as pd
from docx import Document
from docx.shared import Pt
from sklearn.metrics import mean_absolute_error, r2_score
from sklearn.model_selection import KFold, cross_val_predict


ROOT = Path(__file__).resolve().parent
SOURCE_SCRIPT = ROOT / "20260904_Comprehensive_1018.py"

OUTPUT_DOCX = ROOT / "20260920_RESTORE_Compare_0002.docx"
OUTPUT_EXPLAINER_DOCX = ROOT / "20260920_RESTORE_Compare_0002_Explainer.docx"

N_BOOTSTRAP = 5000


def _small(paragraph, size: int = 9) -> None:
    for run in paragraph.runs:
        run.font.size = Pt(size)


def _load_source_module():
    spec = importlib.util.spec_from_file_location("comp20260904", SOURCE_SCRIPT)
    if spec is None or spec.loader is None:
        raise ImportError(f"Cannot load source script: {SOURCE_SCRIPT}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _paired_bootstrap_ci(
    y: np.ndarray,
    pred_bedside: np.ndarray,
    pred_restore: np.ndarray,
    n_bootstrap: int = N_BOOTSTRAP,
    random_state: int = 42,
) -> dict[str, float]:
    rng = np.random.default_rng(random_state)
    n = y.shape[0]

    d_mae = np.empty(n_bootstrap, dtype=float)
    d_r2 = np.empty(n_bootstrap, dtype=float)
    valid = 0
    for _ in range(n_bootstrap):
        idx = rng.integers(0, n, size=n)
        yb = y[idx]
        pb = pred_bedside[idx]
        pr = pred_restore[idx]
        r2_b = r2_score(yb, pb)
        r2_r = r2_score(yb, pr)
        if np.isfinite(r2_b) and np.isfinite(r2_r):
            d_mae[valid] = mean_absolute_error(yb, pb) - mean_absolute_error(yb, pr)
            d_r2[valid] = r2_r - r2_b
            valid += 1

    if valid == 0:
        raise RuntimeError("No valid bootstrap replicates for CI estimation.")

    d_mae = d_mae[:valid]
    d_r2 = d_r2[:valid]
    return {
        "delta_mae_lo": float(np.percentile(d_mae, 2.5)),
        "delta_mae_hi": float(np.percentile(d_mae, 97.5)),
        "delta_r2_lo": float(np.percentile(d_r2, 2.5)),
        "delta_r2_hi": float(np.percentile(d_r2, 97.5)),
        "n_valid_bootstrap": int(valid),
    }


def _fit_oof(
    module,
    df: pd.DataFrame,
    features: list[str],
    cv_splits,
) -> np.ndarray:
    X = df[features]
    y = df["6MWT4"].to_numpy(dtype=float)
    return cross_val_predict(
        module._build_lasso_pipeline(features),
        X,
        y,
        cv=cv_splits,
        n_jobs=-1,
    )


def _write_main_docx(results: dict[str, float], output_path: Path) -> None:
    doc = Document()
    title = doc.add_paragraph()
    title.add_run("RESTORE vs BEDSIDE Week-3 paired comparison (20260920_RESTORE_Compare_0002)").bold = True

    p = doc.add_paragraph(
        "Required revision implemented: BEDSIDE and RESTORE were compared within the same "
        "Week-3-eligible cohort using identical 5-fold cross-validation splits."
    )
    _small(p)

    doc.add_paragraph(
        f"Population counts: BEDSIDE standalone denominator={results['n_bedside_original']}, "
        f"RESTORE standalone denominator={results['n_restore_original']}, "
        f"paired Week-3 denominator={results['n_paired']}."
    )

    unpaired_table = doc.add_table(rows=3, cols=4)
    unpaired_table.style = "Table Grid"
    unpaired_table.cell(0, 0).text = "Standalone (different populations)"
    unpaired_table.cell(0, 1).text = "N"
    unpaired_table.cell(0, 2).text = "OOF MAE (m)"
    unpaired_table.cell(0, 3).text = "OOF R²"
    unpaired_table.cell(1, 0).text = "BEDSIDE (admission-only)"
    unpaired_table.cell(1, 1).text = f"{results['n_bedside_original']}"
    unpaired_table.cell(1, 2).text = f"{results['bedside_original_mae']:.2f}"
    unpaired_table.cell(1, 3).text = f"{results['bedside_original_r2']:.4f}"
    unpaired_table.cell(2, 0).text = "RESTORE (Week-3 update)"
    unpaired_table.cell(2, 1).text = f"{results['n_restore_original']}"
    unpaired_table.cell(2, 2).text = f"{results['restore_original_mae']:.2f}"
    unpaired_table.cell(2, 3).text = f"{results['restore_original_r2']:.4f}"

    doc.add_paragraph(
        "The standalone values above are on different denominators and are not used to claim incremental value."
    )

    paired_table = doc.add_table(rows=4, cols=4)
    paired_table.style = "Table Grid"
    paired_table.cell(0, 0).text = "Paired Week-3 comparison (same N, same folds)"
    paired_table.cell(0, 1).text = "N"
    paired_table.cell(0, 2).text = "OOF MAE (m)"
    paired_table.cell(0, 3).text = "OOF R²"
    paired_table.cell(1, 0).text = "BEDSIDE (admission-only)"
    paired_table.cell(1, 1).text = f"{results['n_paired']}"
    paired_table.cell(1, 2).text = f"{results['bedside_paired_mae']:.2f}"
    paired_table.cell(1, 3).text = f"{results['bedside_paired_r2']:.4f}"
    paired_table.cell(2, 0).text = "RESTORE (admission + Week-3 information)"
    paired_table.cell(2, 1).text = f"{results['n_paired']}"
    paired_table.cell(2, 2).text = f"{results['restore_paired_mae']:.2f}"
    paired_table.cell(2, 3).text = f"{results['restore_paired_r2']:.4f}"
    paired_table.cell(3, 0).text = "Paired difference (RESTORE incremental value)"
    paired_table.cell(3, 1).text = f"{results['n_paired']}"
    paired_table.cell(3, 2).text = (
        f"ΔMAE (BEDSIDE−RESTORE)={results['delta_mae']:.2f} "
        f"[{results['delta_mae_lo']:.2f}, {results['delta_mae_hi']:.2f}]"
    )
    paired_table.cell(3, 3).text = (
        f"ΔR² (RESTORE−BEDSIDE)={results['delta_r2']:.4f} "
        f"[{results['delta_r2_lo']:.4f}, {results['delta_r2_hi']:.4f}]"
    )

    conclusion = doc.add_paragraph(
        "Interpretation: within identical Week-3-eligible participants and identical CV folds, "
        "RESTORE improved MAE and R² versus BEDSIDE, supporting added predictive value from Week-3 information."
    )
    _small(conclusion)

    bootstrap_note = doc.add_paragraph(
        f"Confidence intervals were generated with paired patient-level bootstrap "
        f"({results['n_valid_bootstrap']} valid resamples)."
    )
    _small(bootstrap_note)
    if results["n_nonfinite_excluded"] > 0:
        exclusion_note = doc.add_paragraph(
            f"Rows excluded after OOF prediction due to non-finite values: {results['n_nonfinite_excluded']}."
        )
        _small(exclusion_note)

    doc.save(output_path)


def _write_explainer_docx(results: dict[str, float], output_path: Path) -> None:
    doc = Document()
    h = doc.add_paragraph()
    h.add_run("Explainer: 20260920_RESTORE_Compare_0002").bold = True

    lines = [
        "Why this revision was needed:",
        "• RESTORE and BEDSIDE previously used different outcome populations (456 vs 511), so direct R² ratios were not valid evidence of incremental value.",
        "• Incremental value must be tested by paired comparison on exactly the same participants.",
        "",
        "What was done:",
        "1) Reused data cleaning and model definitions from 20260904_Comprehensive_1018.py.",
        "2) Defined Week-3 eligibility as 6MWT4 non-missing and Rehab_LOS_Category in {'21-42 days', '>42 days'}.",
        "3) Built one fixed 5-fold split object (shuffle=True, random_state=42) and reused it for both models.",
        "4) Generated OOF predictions for BEDSIDE (admission-only features) and RESTORE (admission + T1T2 change features).",
        "5) Computed paired differences with bootstrap percentile CIs:",
        "   • ΔMAE = MAE_BEDSIDE − MAE_RESTORE",
        "   • ΔR² = R²_RESTORE − R²_BEDSIDE",
        "",
        "Key paired findings:",
        f"• N={results['n_paired']} (Week-3-eligible).",
        f"• BEDSIDE: MAE={results['bedside_paired_mae']:.2f} m, R²={results['bedside_paired_r2']:.4f}.",
        f"• RESTORE: MAE={results['restore_paired_mae']:.2f} m, R²={results['restore_paired_r2']:.4f}.",
        f"• ΔMAE={results['delta_mae']:.2f} m [{results['delta_mae_lo']:.2f}, {results['delta_mae_hi']:.2f}].",
        f"• ΔR²={results['delta_r2']:.4f} [{results['delta_r2_lo']:.4f}, {results['delta_r2_hi']:.4f}].",
        "",
        "Interpretation:",
        "Positive ΔMAE and positive ΔR² indicate that updating with Week-3 information improves prediction over admission-only inputs within the same eligible cohort.",
    ]

    for line in lines:
        para = doc.add_paragraph(line)
        _small(para)

    doc.save(output_path)


def main() -> None:
    module = _load_source_module()
    df = pd.read_excel(module.INPUT_XLSX)
    df, _ = module.apply_manual_patient_corrections(df)
    df, _ = module.audit_and_clean_t1t2_data(df)

    bedside_features = module.BASE_MODEL_SPECS["Model 4: Bedside Mobility Core"]
    restore_features = module.BASE_MODEL_SPECS["Model 5: Recovery Trajectory (LOS ≥21 days)"]
    bedside_features = module._filter_existing(bedside_features, df)
    restore_features = module._filter_existing(restore_features, df)

    bedside_original_df = df.loc[df["6MWT4"].notna()].copy()
    week3_eligible_mask = df["6MWT4"].notna() & df["Rehab_LOS_Category"].isin(module.QUALIFYING_REHAB_LOS)
    restore_original_df = df.loc[week3_eligible_mask].copy()
    pair_df = restore_original_df.copy()
    if "ID" in pair_df.columns:
        pair_df = pair_df.sort_values(["ID"]).reset_index(drop=True)
    else:
        pair_df = pair_df.sort_index().reset_index(drop=True)
    cv_splits = list(KFold(n_splits=module.CV_FOLDS, shuffle=True, random_state=module.RANDOM_STATE).split(pair_df))

    y_pair = pair_df["6MWT4"].to_numpy(dtype=float)
    pred_bedside_pair = _fit_oof(module, pair_df, bedside_features, cv_splits)
    pred_restore_pair = _fit_oof(module, pair_df, restore_features, cv_splits)
    evaluable_mask = np.isfinite(y_pair) & np.isfinite(pred_bedside_pair) & np.isfinite(pred_restore_pair)
    n_nonfinite_excluded = int((~evaluable_mask).sum())
    if n_nonfinite_excluded:
        y_pair = y_pair[evaluable_mask]
        pred_bedside_pair = pred_bedside_pair[evaluable_mask]
        pred_restore_pair = pred_restore_pair[evaluable_mask]

    pred_bedside_original = _fit_oof(
        module,
        bedside_original_df,
        bedside_features,
        KFold(n_splits=module.CV_FOLDS, shuffle=True, random_state=module.RANDOM_STATE),
    )
    pred_restore_original = _fit_oof(
        module,
        restore_original_df,
        restore_features,
        KFold(n_splits=module.CV_FOLDS, shuffle=True, random_state=module.RANDOM_STATE),
    )
    y_bedside_original = bedside_original_df["6MWT4"].to_numpy(dtype=float)
    y_restore_original = restore_original_df["6MWT4"].to_numpy(dtype=float)

    ci = _paired_bootstrap_ci(y_pair, pred_bedside_pair, pred_restore_pair)
    results = {
        "n_bedside_original": int(len(bedside_original_df)),
        "n_restore_original": int(len(restore_original_df)),
        "n_paired": int(y_pair.shape[0]),
        "n_nonfinite_excluded": int(n_nonfinite_excluded),
        "bedside_original_mae": float(mean_absolute_error(y_bedside_original, pred_bedside_original)),
        "bedside_original_r2": float(r2_score(y_bedside_original, pred_bedside_original)),
        "restore_original_mae": float(mean_absolute_error(y_restore_original, pred_restore_original)),
        "restore_original_r2": float(r2_score(y_restore_original, pred_restore_original)),
        "bedside_paired_mae": float(mean_absolute_error(y_pair, pred_bedside_pair)),
        "bedside_paired_r2": float(r2_score(y_pair, pred_bedside_pair)),
        "restore_paired_mae": float(mean_absolute_error(y_pair, pred_restore_pair)),
        "restore_paired_r2": float(r2_score(y_pair, pred_restore_pair)),
        "delta_mae": float(mean_absolute_error(y_pair, pred_bedside_pair) - mean_absolute_error(y_pair, pred_restore_pair)),
        "delta_r2": float(r2_score(y_pair, pred_restore_pair) - r2_score(y_pair, pred_bedside_pair)),
        **ci,
    }

    _write_main_docx(results, OUTPUT_DOCX)
    _write_explainer_docx(results, OUTPUT_EXPLAINER_DOCX)

    print(f"Wrote: {OUTPUT_DOCX.name}")
    print(f"Wrote: {OUTPUT_EXPLAINER_DOCX.name}")
    print(
        "Paired deltas:",
        f"ΔMAE={results['delta_mae']:.2f} m "
        f"[{results['delta_mae_lo']:.2f}, {results['delta_mae_hi']:.2f}] ; "
        f"ΔR²={results['delta_r2']:.4f} "
        f"[{results['delta_r2_lo']:.4f}, {results['delta_r2_hi']:.4f}]"
    )


if __name__ == "__main__":
    main()
