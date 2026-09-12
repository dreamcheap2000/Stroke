from __future__ import annotations

import importlib.util
from pathlib import Path

import matplotlib
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import statsmodels.api as sm
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from sklearn.impute import SimpleImputer
from sklearn.linear_model import Ridge
from sklearn.model_selection import KFold
from statsmodels.stats.outliers_influence import variance_inflation_factor


matplotlib.use("Agg")

ROOT = Path(__file__).resolve().parent
SOURCE_XLSX = ROOT / "20260909_Comprehensive.xlsx"
SOURCE_COMPREHENSIVE_XLSX = ROOT / "20260904_Comprehensive_1018.xlsx"
SOURCE_SCRIPT = ROOT / "20260904_Comprehensive_1018.py"
SOURCE_DATA_XLSX = ROOT / "20260806_DeID.xlsx"

OUTPUT_DOCX = ROOT / "20260912_Tables_1840.docx"
OUTPUT_SUPPLEMENTARY = ROOT / "20260912_Tables_1840_Supplementary.docx"
OUTPUT_METHODS = ROOT / "20260912_Method_2039.docx"
OUTPUT_RESULTS = ROOT / "20260912_Result_2039.docx"
OUTPUT_CALIBRATION_PNG = ROOT / "20260912_Calibration_1840.png"

STABILITY_THRESHOLD = 0.70
SHARED_PREDICTOR_MIN_MODELS = 2
PRIMARY_PRESENTATION_ORDER = ["BEDSIDE", "AIMS", "RESTORE", "COMPASS", "CASCADE"]
CLINICAL_ROLE = {
    "BEDSIDE": "Primary recommendation",
    "AIMS": "Secondary recommendation",
    "RESTORE": "Reference trajectory model",
    "COMPASS": "Complexity-sensitive comparison",
    "CASCADE": "Complexity-sensitive comparison",
}
DEFAULT_CLINICAL_ROLE = "Additional comparison"
MCID_BENCHMARKS_M = [20.0, 34.4]
COLLINEARITY_VARIABLES = ["BBS1", "Gait_Speed_1", "Age", "FuglUE1"]


def load_source_module():
    spec = importlib.util.spec_from_file_location("comp20260904", SOURCE_SCRIPT)
    if spec is None or spec.loader is None:
        raise ImportError(f"Cannot load source script: {SOURCE_SCRIPT}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def order_key(acronym: str) -> int:
    try:
        return PRIMARY_PRESENTATION_ORDER.index(str(acronym))
    except ValueError:
        return len(PRIMARY_PRESENTATION_ORDER)


def sort_by_presentation(df: pd.DataFrame, column: str = "Acronym") -> pd.DataFrame:
    return df.sort_values([column], key=lambda s: s.map(order_key)).reset_index(drop=True)


def first_row(df: pd.DataFrame, acronym: str) -> pd.Series | None:
    subset = df.loc[df["Acronym"].eq(acronym)]
    if subset.empty:
        return None
    return subset.iloc[0]


def set_cell_shading(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def set_table_borders(table) -> None:
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    borders = tbl_pr.first_child_found_in("w:tblBorders")
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = f"w:{edge}"
        elem = borders.find(qn(tag))
        if elem is None:
            elem = OxmlElement(tag)
            borders.append(elem)
        elem.set(qn("w:val"), "single")
        elem.set(qn("w:sz"), "4")
        elem.set(qn("w:space"), "0")
        elem.set(qn("w:color"), "D9D9D9")


def ensure_landscape(doc: Document) -> None:
    section = doc.sections[-1]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)
    section.top_margin = Inches(0.55)
    section.bottom_margin = Inches(0.55)


def add_table(doc: Document, df: pd.DataFrame, *, title: str, note: str | None = None, font_size: int = 8) -> None:
    heading = doc.add_paragraph()
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = heading.add_run(title)
    run.bold = True
    run.font.size = Pt(10)

    if note:
        paragraph = doc.add_paragraph(note)
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for note_run in paragraph.runs:
            note_run.font.size = Pt(8)

    table = doc.add_table(rows=1, cols=len(df.columns))
    table.style = "Table Grid"
    table.autofit = True
    set_table_borders(table)

    header_cells = table.rows[0].cells
    for col_idx, col_name in enumerate(df.columns):
        cell = header_cells[col_idx]
        cell.text = str(col_name)
        set_cell_shading(cell, "1F4E78")
        for paragraph in cell.paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in cell.paragraphs[0].runs:
            run.bold = True
            run.font.size = Pt(font_size)
            run.font.color.rgb = RGBColor(255, 255, 255)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    for row_values in df.fillna("—").astype(str).itertuples(index=False, name=None):
        row_cells = table.add_row().cells
        for col_idx, value in enumerate(row_values):
            cell = row_cells[col_idx]
            cell.text = value
            if col_idx == 0:
                set_cell_shading(cell, "EAF2F8")
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER if col_idx else WD_ALIGN_PARAGRAPH.LEFT
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(font_size)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    doc.add_paragraph()


def fmt_or_dash(value, digits: int) -> str:
    if pd.isna(value):
        return "—"
    return f"{float(value):.{digits}f}"


def fmt_signed(value, digits: int) -> str:
    if pd.isna(value):
        return "—"
    return f"{float(value):+.{digits}f}"


def weighted_mean(values: np.ndarray, weights: np.ndarray) -> float:
    return float(np.average(values, weights=weights))


def build_compact_performance_table(
    model_explainers: pd.DataFrame,
    panel_a: pd.DataFrame,
    calibration_df: pd.DataFrame | None = None,
) -> pd.DataFrame:
    explainers = model_explainers.rename(columns={"Publication_Name": "Publication name"}).copy()
    merged = panel_a.merge(
        explainers[
            [
                "Overall_Rank",
                "Model_label",
                "Acronym",
                "Publication name",
                "Input_vars",
                "CV_R2",
                "CV_MAE",
                "Weighted_OOF_R2",
                "Weighted_OOF_MAE",
            ]
        ],
        left_on=["Acronym", "Publication name", "Input vars"],
        right_on=["Acronym", "Publication name", "Input_vars"],
        how="left",
        validate="one_to_one",
    )

    if calibration_df is not None and not calibration_df.empty:
        merged = merged.merge(
            calibration_df[["Model_label", "Calibration_Intercept", "Calibration_Slope"]],
            left_on="Model",
            right_on="Model_label",
            how="left",
            validate="one_to_one",
        )
    else:
        merged["Calibration_Intercept"] = np.nan
        merged["Calibration_Slope"] = np.nan

    best_row = first_row(merged, "RESTORE")
    merged["Clinical role"] = merged["Acronym"].map(CLINICAL_ROLE).fillna(DEFAULT_CLINICAL_ROLE)
    if best_row is None:
        merged["Δ vs RESTORE R²"] = np.nan
        merged["Δ vs RESTORE MAE (m)"] = np.nan
    else:
        merged["Δ vs RESTORE R²"] = merged["Weighted_OOF_R2"] - float(best_row["Weighted_OOF_R2"])
        merged["Δ vs RESTORE MAE (m)"] = merged["Weighted_OOF_MAE"] - float(best_row["Weighted_OOF_MAE"])
    merged = sort_by_presentation(merged)

    compact = merged[
        [
            "Clinical role",
            "Acronym",
            "Publication name",
            "Input vars",
            "CV_R2",
            "CV_MAE",
            "Weighted_OOF_R2",
            "Weighted_OOF_MAE",
            "Δ vs RESTORE R²",
            "Δ vs RESTORE MAE (m)",
            "Calibration_Intercept",
            "Calibration_Slope",
            "Best BalAcc [95% CI]",
            "Worst BalAcc [95% CI]",
        ]
    ].copy()
    compact.rename(
        columns={
            "CV_R2": "Part 1 CV R²",
            "CV_MAE": "Part 1 MAE (m)",
            "Weighted_OOF_R2": "Part 2 weighted OOF R²",
            "Weighted_OOF_MAE": "Part 2 weighted OOF MAE (m)",
            "Calibration_Intercept": "Calibration intercept (m)",
            "Calibration_Slope": "Calibration slope",
        },
        inplace=True,
    )
    metric_formats = {
        "Part 1 CV R²": 4,
        "Part 1 MAE (m)": 2,
        "Part 2 weighted OOF R²": 4,
        "Part 2 weighted OOF MAE (m)": 1,
        "Δ vs RESTORE R²": 4,
        "Δ vs RESTORE MAE (m)": 1,
        "Calibration intercept (m)": 1,
        "Calibration slope": 3,
    }
    for column_name, digits in metric_formats.items():
        formatter = fmt_signed if column_name.startswith("Δ ") else fmt_or_dash
        compact[column_name] = compact[column_name].map(lambda x, digits=digits, formatter=formatter: formatter(x, digits))
    return compact


def build_compact_predictor_table(model_explainers: pd.DataFrame, lasso_coefficients: pd.DataFrame) -> pd.DataFrame:
    stable = lasso_coefficients[lasso_coefficients["selection_frequency"] >= STABILITY_THRESHOLD].copy()
    predictor_summary = (
        stable.groupby("Predictor_Display", as_index=False)
        .agg(
            Model_Count=("Model_label", "nunique"),
            Max_SF=("selection_frequency", "max"),
            Max_Abs=("abs_bootstrap_coef_mean", "max"),
        )
    )
    keep_predictors = predictor_summary[predictor_summary["Model_Count"] >= SHARED_PREDICTOR_MIN_MODELS].copy()
    keep_predictors.sort_values(
        ["Model_Count", "Max_SF", "Max_Abs", "Predictor_Display"],
        ascending=[False, False, False, True],
        inplace=True,
    )

    stable = stable[stable["Predictor_Display"].isin(keep_predictors["Predictor_Display"])].copy()
    stable["Model_Column"] = (
        stable["Acronym"].astype(str)
        + " (n="
        + stable["Input_vars"].astype(int).astype(str)
        + ")"
    )
    stable["Cell"] = stable["bootstrap_coef_mean"].map(lambda value: f"{float(value):.1f}") + "; " + stable["Selection_Freq_Pct"].astype(str)

    order_frame = (
        model_explainers[["Acronym", "Input_vars"]]
        .drop_duplicates()
        .assign(_order=lambda df: df["Acronym"].map(order_key))
        .sort_values(["_order", "Input_vars"])
    )
    column_order = [
        f"{row.Acronym} (n={int(row.Input_vars)})"
        for row in order_frame.itertuples(index=False)
    ]
    compact = (
        stable.pivot_table(index="Predictor_Display", columns="Model_Column", values="Cell", aggfunc="first")
        .reindex(keep_predictors["Predictor_Display"].tolist())
        .reindex(columns=column_order)
        .reset_index()
        .rename(columns={"Predictor_Display": "Predictor"})
        .fillna("—")
    )
    return compact


def compute_continuous_calibration(source, model_explainers: pd.DataFrame) -> tuple[pd.DataFrame, dict[str, pd.DataFrame]]:
    predictions_df = pd.read_excel(SOURCE_COMPREHENSIVE_XLSX, sheet_name="Predictions")
    ipcw = source.compute_ipcw_weights_single_model(predictions_df, source.IPCW_COMPLETION_FEATURES)
    completer_mask = (
        predictions_df["PAC_Program_Completion"].astype("string").eq("Completed PAC program")
        & predictions_df["6MWT4"].notna()
    )
    df_comp = predictions_df.loc[completer_mask].copy()
    y = df_comp["6MWT4"].to_numpy(dtype=float)
    weights = ipcw["weights"].loc[completer_mask].fillna(1.0).to_numpy(dtype=float)

    rows: list[dict] = []
    decile_data: dict[str, pd.DataFrame] = {}
    explainers_by_model = model_explainers.set_index("Model_label")

    for model_name in source.RETAINED_MODEL_ORDER:
        model_label = model_name.split(":", 1)[0]
        explainer = explainers_by_model.loc[model_label]
        valid_features = source._filter_existing(source.MODEL_SPECS[model_name], predictions_df)
        X_raw = df_comp[valid_features]
        cv_splits = list(KFold(n_splits=source.CV_FOLDS, shuffle=True, random_state=source.RANDOM_STATE).split(X_raw))
        oof = np.zeros(len(df_comp), dtype=float)

        for tr, te in cv_splits:
            imputer = SimpleImputer(strategy="median")
            X_tr = imputer.fit_transform(X_raw.iloc[tr])
            X_te = imputer.transform(X_raw.iloc[te])
            model = Ridge(alpha=1.0)
            model.fit(X_tr, y[tr], sample_weight=weights[tr])
            oof[te] = np.maximum(0, model.predict(X_te))

        predicted_mean = weighted_mean(oof, weights)
        observed_mean = weighted_mean(y, weights)
        centered_pred = oof - predicted_mean
        slope_model = sm.WLS(y - observed_mean, centered_pred, weights=weights).fit()
        order = np.argsort(oof)
        bins = np.array_split(order, 10)
        decile_rows = []
        for idx, bin_idx in enumerate(bins, start=1):
            if len(bin_idx) == 0:
                continue
            decile_rows.append(
                {
                    "Decile": idx,
                    "Predicted": weighted_mean(oof[bin_idx], weights[bin_idx]),
                    "Observed": weighted_mean(y[bin_idx], weights[bin_idx]),
                }
            )
        decile_data[str(explainer["Acronym"])] = pd.DataFrame(decile_rows)
        rows.append(
            {
                "Model_label": model_label,
                "Acronym": str(explainer["Acronym"]),
                "Calibration_Intercept": float(observed_mean - predicted_mean),
                "Calibration_Slope": float(slope_model.params[0]),
                "Observed_mean_6MWT4": observed_mean,
                "Predicted_mean_6MWT4": predicted_mean,
                "Weighted_OOF_R2": float(explainer["Weighted_OOF_R2"]),
                "Weighted_OOF_MAE": float(explainer["Weighted_OOF_MAE"]),
            }
        )

    calibration_df = sort_by_presentation(pd.DataFrame(rows))
    return calibration_df, decile_data


def save_calibration_plot(decile_data: dict[str, pd.DataFrame]) -> None:
    fig, axes = plt.subplots(2, 3, figsize=(14, 8))
    axes_flat = axes.flatten()
    ordered_items = sorted(decile_data.items(), key=lambda item: order_key(item[0]))
    used_axes = axes_flat[: len(ordered_items)]

    all_values = []
    for df in decile_data.values():
        all_values.extend(df["Predicted"].tolist())
        all_values.extend(df["Observed"].tolist())
    limit = max(all_values) if all_values else 1.0
    limit = max(50.0, float(limit) * 1.05)

    for ax, (acronym, df) in zip(used_axes, ordered_items):
        ax.plot([0, limit], [0, limit], linestyle="--", color="gray", linewidth=1)
        ax.plot(df["Predicted"], df["Observed"], marker="o", color="#1F4E78")
        ax.set_title(acronym)
        ax.set_xlabel("Weighted mean predicted 6MWT4 (m)")
        ax.set_ylabel("Weighted mean observed 6MWT4 (m)")
        ax.set_xlim(0, limit)
        ax.set_ylim(0, limit)
        ax.grid(alpha=0.25)

    for ax in axes_flat[len(ordered_items):]:
        ax.axis("off")

    fig.suptitle("Internal weighted OOF calibration by clinically ordered retained model", fontsize=12)
    fig.tight_layout()
    fig.savefig(OUTPUT_CALIBRATION_PNG, dpi=200, bbox_inches="tight")
    plt.close(fig)


def build_parsimony_table(
    ranking_df: pd.DataFrame,
    model_explainers: pd.DataFrame,
    lasso_coefficients: pd.DataFrame,
    calibration_df: pd.DataFrame,
) -> pd.DataFrame:
    merged = ranking_df.copy()
    merged = merged.merge(
        calibration_df[["Acronym", "Calibration_Slope", "Calibration_Intercept"]],
        on="Acronym",
        how="left",
        validate="one_to_one",
    )
    restore_row = first_row(merged, "RESTORE")
    best_r2 = float(restore_row["Weighted_OOF_R2"]) if restore_row is not None else np.nan
    best_mae = float(restore_row["Weighted_OOF_MAE"]) if restore_row is not None else np.nan

    stable = lasso_coefficients[lasso_coefficients["selection_frequency"] >= STABILITY_THRESHOLD].copy()
    rows = []
    for row in merged.itertuples(index=False):
        stable_model = stable[stable["Acronym"] == row.Acronym].copy()
        other_predictors = stable.loc[stable["Acronym"] != row.Acronym, "Predictor_Display"]
        unique_count = int((~stable_model["Predictor_Display"].isin(other_predictors)).sum())
        rows.append(
            {
                "Clinical role": CLINICAL_ROLE.get(row.Acronym, DEFAULT_CLINICAL_ROLE),
                "Acronym": row.Acronym,
                "Original point-estimate rank": int(row.Overall_Rank),
                "Input vars": int(row.Input_vars),
                "Outcome observations": int(row.N_patients),
                "Observations/variable": f"{float(row.N_patients / row.Input_vars):.1f}",
                "Δ vs RESTORE R²": "—" if pd.isna(best_r2) else f"{float(row.Weighted_OOF_R2 - best_r2):+.4f}",
                "Δ vs RESTORE MAE (m)": "—" if pd.isna(best_mae) else f"{float(row.Weighted_OOF_MAE - best_mae):+.1f}",
                "Calibration intercept (m)": fmt_or_dash(row.Calibration_Intercept, 1),
                "Calibration slope": fmt_or_dash(row.Calibration_Slope, 3),
                "Stable predictors (≥70%)": int(len(stable_model)),
                "Stable predictors at 70-89%": int(((stable_model["selection_frequency"] >= 0.70) & (stable_model["selection_frequency"] < 0.90)).sum()),
                "Stable predictors at 90-99%": int(((stable_model["selection_frequency"] >= 0.90) & (stable_model["selection_frequency"] < 1.0)).sum()),
                "Stable predictors at 100%": int((stable_model["selection_frequency"] == 1.0).sum()),
                "Stable predictors unique to model": unique_count,
            }
        )
    parsimony_df = sort_by_presentation(pd.DataFrame(rows))
    return parsimony_df


def build_collinearity_outputs() -> tuple[pd.DataFrame, pd.DataFrame]:
    data = pd.read_excel(SOURCE_DATA_XLSX)[COLLINEARITY_VARIABLES].apply(pd.to_numeric, errors="coerce")
    corr = data.corr().round(3).reset_index().rename(columns={"index": "Predictor"})

    complete = data.dropna().reset_index(drop=True)
    X = sm.add_constant(complete)
    vif_rows = []
    for idx, column in enumerate(complete.columns, start=1):
        vif_rows.append(
            {
                "Predictor": column,
                "VIF": round(float(variance_inflation_factor(X.values, idx)), 3),
                "Complete-case n": int(len(complete)),
            }
        )
    return corr, pd.DataFrame(vif_rows)


def build_mcid_table(model_explainers: pd.DataFrame) -> pd.DataFrame:
    rows = []
    low_benchmark = min(MCID_BENCHMARKS_M)
    high_benchmark = max(MCID_BENCHMARKS_M)
    for row in model_explainers.sort_values(["Acronym"], key=lambda s: s.map(order_key)).itertuples(index=False):
        mae = float(row.Weighted_OOF_MAE)
        rows.append(
            {
                "Acronym": row.Acronym,
                "Weighted OOF MAE (m)": f"{mae:.1f}",
                f"MAE / {low_benchmark:.1f} m MCID": f"{mae / low_benchmark:.2f}",
                f"MAE / {high_benchmark:.1f} m MCID": f"{mae / high_benchmark:.2f}",
            }
        )
    return pd.DataFrame(rows)


def reorder_panel_b(panel_b: pd.DataFrame) -> pd.DataFrame:
    ordered_columns = ["Predictor"]
    for acronym in PRIMARY_PRESENTATION_ORDER:
        ordered_columns.extend([c for c in panel_b.columns if c.startswith(f"{acronym} ")])
    remaining = [c for c in panel_b.columns if c not in ordered_columns]
    return panel_b[ordered_columns + remaining].copy()


def reorder_table2(table2: pd.DataFrame) -> pd.DataFrame:
    return table2.sort_values(["Acronym", "Predictor"], key=lambda s: s.map(order_key) if s.name == "Acronym" else s).reset_index(drop=True)


def load_tables() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    with pd.ExcelFile(SOURCE_XLSX) as workbook:
        model_ranking = workbook.parse("Model_Ranking").sort_values("Overall_Rank")
        model_explainers = workbook.parse("Model_Explainers").sort_values("Overall_Rank")
        panel_a = workbook.parse("Pub_Table1_PanelA")
        panel_b = workbook.parse("Pub_Table1_PanelB")
        table2 = workbook.parse("Pub_Table2")
        lasso_coefficients = workbook.parse("LASSO_Coefficients")
    return model_ranking, model_explainers, panel_a, panel_b, table2, lasso_coefficients


def write_main_document(compact_performance: pd.DataFrame, compact_predictors: pd.DataFrame) -> None:
    doc = Document()
    ensure_landscape(doc)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("Publication Tables")
    run.bold = True
    run.font.size = Pt(12)

    intro = doc.add_paragraph(
        "Models are presented in clinical deployment order rather than point-estimate rank because the internal balanced-accuracy intervals are broadly overlapping and do not support a definitive superiority claim."
    )
    intro.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in intro.runs:
        run.font.size = Pt(8)

    intro2 = doc.add_paragraph(
        "BEDSIDE is the primary deployment recommendation and AIMS is the secondary recommendation because both preserve internal performance within about 0.06 weighted OOF R² and 6 m weighted OOF MAE of the larger retained models while using far fewer predictors."
    )
    intro2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in intro2.runs:
        run.font.size = Pt(8)

    add_table(
        doc,
        compact_performance,
        title="Table 1. Clinically ordered comparison of the 5 retained models",
        note="Rows are intentionally ordered by clinical parsimony. Calibration metrics are from weighted out-of-fold predictions among completers; balanced accuracy is still shown for the Best and Worst non-completer walking scenarios.",
        font_size=7,
    )
    add_table(
        doc,
        compact_predictors,
        title="Table 2. Stable predictors shared across retained models",
        note="Rows are stable predictors retained in at least 2 models. Cells show bootstrap mean standardized β and selection frequency, reordered to place the parsimonious deployment models first.",
        font_size=7,
    )

    legend = doc.add_paragraph()
    legend.add_run("Abbreviations. ").bold = True
    legend.add_run("BalAcc = balanced accuracy; MAE = mean absolute error; OOF = out of fold; SF = selection frequency.")
    for run in legend.runs:
        run.font.size = Pt(8)

    doc.save(OUTPUT_DOCX)


def write_supplementary_document(
    parsimony_df: pd.DataFrame,
    panel_a: pd.DataFrame,
    calibration_df: pd.DataFrame,
    corr_df: pd.DataFrame,
    vif_df: pd.DataFrame,
    mcid_df: pd.DataFrame,
    panel_b: pd.DataFrame,
    table2: pd.DataFrame,
) -> None:
    doc = Document()
    ensure_landscape(doc)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("Supplementary Tables")
    run.bold = True
    run.font.size = Pt(12)

    add_table(
        doc,
        parsimony_df,
        title="Supplementary Table S1. Parsimony, calibration, and coefficient-stability comparison across retained models",
        note="Original point-estimate rank is retained only as reference. The observation-per-variable summary highlights why the largest models should be interpreted cautiously.",
        font_size=7,
    )
    add_table(
        doc,
        sort_by_presentation(panel_a),
        title="Supplementary Table S2. Full Part 2 IPCW binary classification performance",
        note="Binary performance is preserved exactly, but ordered clinically rather than by point estimate.",
        font_size=7,
    )
    add_table(
        doc,
        calibration_df.assign(
            **{
                "Calibration_Intercept": lambda df: df["Calibration_Intercept"].map(lambda x: fmt_or_dash(x, 1)),
                "Calibration_Slope": lambda df: df["Calibration_Slope"].map(lambda x: fmt_or_dash(x, 3)),
                "Observed_mean_6MWT4": lambda df: df["Observed_mean_6MWT4"].map(lambda x: fmt_or_dash(x, 1)),
                "Predicted_mean_6MWT4": lambda df: df["Predicted_mean_6MWT4"].map(lambda x: fmt_or_dash(x, 1)),
                "Weighted_OOF_R2": lambda df: df["Weighted_OOF_R2"].map(lambda x: fmt_or_dash(x, 4)),
                "Weighted_OOF_MAE": lambda df: df["Weighted_OOF_MAE"].map(lambda x: fmt_or_dash(x, 1)),
            }
        )[
            [
                "Acronym",
                "Calibration_Intercept",
                "Calibration_Slope",
                "Observed_mean_6MWT4",
                "Predicted_mean_6MWT4",
                "Weighted_OOF_R2",
                "Weighted_OOF_MAE",
            ]
        ].rename(
            columns={
                "Calibration_Intercept": "Calibration intercept (m)",
                "Calibration_Slope": "Calibration slope",
                "Observed_mean_6MWT4": "Weighted mean observed 6MWT4 (m)",
                "Predicted_mean_6MWT4": "Weighted mean predicted 6MWT4 (m)",
                "Weighted_OOF_R2": "Weighted OOF R²",
                "Weighted_OOF_MAE": "Weighted OOF MAE (m)",
            }
        ),
        title="Supplementary Table S3. Internal weighted OOF calibration summary",
        note="Calibration intercept is the weighted mean observed-minus-predicted difference (ideal 0 m), and calibration slope is the centered weighted recalibration slope (ideal 1.0) among completers.",
        font_size=7,
    )

    figure_note = doc.add_paragraph(
        "Supplementary Figure S1. Decile-based internal weighted OOF calibration plot for each retained model."
    )
    figure_note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in figure_note.runs:
        run.font.size = Pt(8)
    doc.add_picture(str(OUTPUT_CALIBRATION_PNG), width=Inches(9.5))
    doc.add_paragraph()

    add_table(
        doc,
        mcid_df,
        title="Supplementary Table S4. Weighted OOF MAE relative to published 6MWT MCID benchmarks already used in the repository sensitivity analyses",
        note="All retained-model MAEs exceed both the 20.0 m and 34.4 m benchmarks, so distance-level prediction error remains clinically material despite internal discrimination.",
        font_size=7,
    )
    add_table(
        doc,
        corr_df,
        title="Supplementary Table S5a. Correlation structure of dominant stable predictors",
        note="Pairwise Pearson correlations are shown for BBS1, Gait_Speed_1, Age, and FuglUE1.",
        font_size=7,
    )
    add_table(
        doc,
        vif_df,
        title="Supplementary Table S5b. Variance inflation factors for dominant stable predictors",
        note="VIFs were calculated on complete cases for the same 4 predictors.",
        font_size=7,
    )
    add_table(
        doc,
        reorder_panel_b(panel_b),
        title="Supplementary Table S6. Full publication Table 1 Panel B stable coefficient matrix",
        note="Rows are predictors and columns are retained models. Cells are bootstrap mean coefficient (bootstrap SD) [95% CI]; selection frequency.",
        font_size=7,
    )
    add_table(
        doc,
        reorder_table2(table2),
        title="Supplementary Table S7. Full publication Table 2 detailed stable coefficients",
        note="This table retains the detailed coefficient rows for every stable predictor-model pairing from the original publication export.",
        font_size=7,
    )

    doc.save(OUTPUT_SUPPLEMENTARY)


def write_narrative_documents(
    model_explainers: pd.DataFrame,
    calibration_df: pd.DataFrame,
    parsimony_df: pd.DataFrame,
    corr_df: pd.DataFrame,
    vif_df: pd.DataFrame,
) -> None:
    restore_metrics = first_row(model_explainers, "RESTORE")
    bedside_metrics = first_row(model_explainers, "BEDSIDE")
    aims_metrics = first_row(model_explainers, "AIMS")
    cal_slope_min = calibration_df["Calibration_Slope"].min()
    cal_slope_max = calibration_df["Calibration_Slope"].max()
    cal_int_min = calibration_df["Calibration_Intercept"].min()
    cal_int_max = calibration_df["Calibration_Intercept"].max()
    corr_bbs_gs = float(corr_df.loc[corr_df["Predictor"].eq("BBS1"), "Gait_Speed_1"].iloc[0])
    vif_lookup = dict(zip(vif_df["Predictor"], vif_df["VIF"]))
    complex_rows = parsimony_df.loc[parsimony_df["Clinical role"].eq("Complexity-sensitive comparison")].copy()
    complexity_sentence = "Complexity-sensitive models were not available in the retained set."
    if not complex_rows.empty:
        complex_rows = sort_by_presentation(complex_rows)
        pieces = []
        for _, row in complex_rows.iterrows():
            pieces.append(
                f"{row['Acronym']} ({row['Observations/variable']} observations per variable; {int(row['Stable predictors at 70-89%'])} stable predictors at 70-89% selection frequency)"
            )
        unique_max = int(complex_rows["Stable predictors unique to model"].max())
        complexity_sentence = (
            "The larger comparison models were "
            + "; ".join(pieces)
            + f", and up to {unique_max} stable predictors unique to a single complex model, supporting the interpretation that added complexity mainly reflects coefficient instability."
        )
    mcid_low, mcid_high = MCID_BENCHMARKS_M

    if any(item is None for item in [restore_metrics, bedside_metrics, aims_metrics]):
        raise ValueError("RESTORE, BEDSIDE, and AIMS must be present to write the publication narrative outputs.")

    methods_doc = Document()
    methods_doc.add_paragraph(
        "For publication-facing tables, retained models were reframed as statistically comparable rather than definitively rank ordered. The clinically parsimonious presentation now leads with BEDSIDE (4 variables) and AIMS (12 variables), while the original point-estimate ordering is preserved only as a supplementary reference. Continuous-model calibration was added by refitting the retained IPCW-weighted Ridge models with the original 5-fold out-of-fold procedure among completers, then estimating calibration intercept and slope from weighted linear recalibration of observed 6MWT4 on out-of-fold predictions. Decile-based calibration plots were generated for each retained model. To address overfitting concerns, we report outcome observations per variable together with stable-predictor frequency bands and model-unique stable predictors for the larger COMPASS and CASCADE models. Clinical interpretability was strengthened by benchmarking weighted out-of-fold MAE against previously used stroke-relevant 6MWT MCID anchors (20.0 m and 34.4 m) and by adding correlation/VIF diagnostics for BBS1, Gait_Speed_1, Age, and FuglUE1. No temporally or geographically distinct external validation cohort was available, so the revised documents now state explicitly that internal cross-validation and weighted out-of-fold estimates do not substitute for external validation before final model selection."
    )
    methods_doc.save(OUTPUT_METHODS)

    results_doc = Document()
    results_doc.add_paragraph(
        f"Internal performance differences across retained models were modest enough that the publication-facing comparison now emphasizes clinical parsimony rather than point-estimate rank. BEDSIDE is presented as the primary deployment model and AIMS as the secondary option because, relative to RESTORE, their weighted out-of-fold R² differed by {float(bedside_metrics['Weighted_OOF_R2'] - restore_metrics['Weighted_OOF_R2']):+.4f} and {float(aims_metrics['Weighted_OOF_R2'] - restore_metrics['Weighted_OOF_R2']):+.4f}, while weighted out-of-fold MAE increased by only {float(bedside_metrics['Weighted_OOF_MAE'] - restore_metrics['Weighted_OOF_MAE']):+.1f} m and {float(aims_metrics['Weighted_OOF_MAE'] - restore_metrics['Weighted_OOF_MAE']):+.1f} m. Continuous calibration remained reasonably close to ideal across retained models (slopes {cal_slope_min:.3f}-{cal_slope_max:.3f}; intercepts {cal_int_min:.1f}-{cal_int_max:.1f} m), but weighted out-of-fold MAE still ranged from {float(restore_metrics['Weighted_OOF_MAE']):.1f} to {float(bedside_metrics['Weighted_OOF_MAE']):.1f} m, exceeding the {mcid_low:.1f}-{mcid_high:.1f} m 6MWT MCID benchmarks. {complexity_sentence} Collinearity diagnostics showed a strong physiologic correlation between BBS1 and Gait_Speed_1 (r={corr_bbs_gs:.3f}) but only moderate VIFs (BBS1 {vif_lookup['BBS1']:.2f}, Gait_Speed_1 {vif_lookup['Gait_Speed_1']:.2f}, Age {vif_lookup['Age']:.2f}, FuglUE1 {vif_lookup['FuglUE1']:.2f}). These findings remain based on internal validation only and should be treated as provisional until tested in a distinct external cohort."
    )
    results_doc.save(OUTPUT_RESULTS)


def main() -> None:
    source = load_source_module()
    model_ranking, model_explainers, panel_a, panel_b, table2, lasso_coefficients = load_tables()
    calibration_df, decile_data = compute_continuous_calibration(source, model_explainers)
    save_calibration_plot(decile_data)
    compact_performance = build_compact_performance_table(model_explainers, panel_a, calibration_df)
    compact_predictors = build_compact_predictor_table(model_explainers, lasso_coefficients)
    parsimony_df = build_parsimony_table(model_ranking, model_explainers, lasso_coefficients, calibration_df)
    corr_df, vif_df = build_collinearity_outputs()
    mcid_df = build_mcid_table(model_explainers)
    write_main_document(compact_performance, compact_predictors)
    write_supplementary_document(
        parsimony_df,
        panel_a,
        calibration_df,
        corr_df,
        vif_df,
        mcid_df,
        panel_b,
        table2,
    )
    write_narrative_documents(model_explainers, calibration_df, parsimony_df, corr_df, vif_df)
    print(f"Wrote {OUTPUT_DOCX.name}")
    print(f"Wrote {OUTPUT_SUPPLEMENTARY.name}")
    print(f"Wrote {OUTPUT_METHODS.name}")
    print(f"Wrote {OUTPUT_RESULTS.name}")
    print(f"Wrote {OUTPUT_CALIBRATION_PNG.name}")


if __name__ == "__main__":
    main()
