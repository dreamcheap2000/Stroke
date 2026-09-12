from pathlib import Path

import pandas as pd
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


ROOT = Path(__file__).resolve().parent
SOURCE_XLSX = ROOT / "20260909_Comprehensive.xlsx"
OUTPUT_DOCX = ROOT / "20260912_Tables_1840.docx"
OUTPUT_SUPPLEMENTARY = ROOT / "20260912_Tables_1840_Supplementary.docx"
STABILITY_THRESHOLD = 0.70
SHARED_PREDICTOR_MIN_MODELS = 2


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


def build_compact_performance_table(model_explainers: pd.DataFrame, panel_a: pd.DataFrame) -> pd.DataFrame:
    def fmt_or_dash(value, digits: int) -> str:
        if pd.isna(value):
            return "—"
        return f"{float(value):.{digits}f}"

    merged = panel_a.merge(
        model_explainers[
            [
                "Overall_Rank",
                "CV_R2",
                "CV_MAE",
                "Weighted_OOF_R2",
                "Weighted_OOF_MAE",
            ]
        ],
        left_on="Rank",
        right_on="Overall_Rank",
        how="left",
        validate="one_to_one",
    )
    compact = merged[
        [
            "Rank",
            "Acronym",
            "Publication name",
            "Input vars",
            "CV_R2",
            "CV_MAE",
            "Weighted_OOF_R2",
            "Weighted_OOF_MAE",
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
        },
        inplace=True,
    )
    compact["Part 1 CV R²"] = compact["Part 1 CV R²"].map(lambda x: fmt_or_dash(x, 4))
    compact["Part 1 MAE (m)"] = compact["Part 1 MAE (m)"].map(lambda x: fmt_or_dash(x, 2))
    compact["Part 2 weighted OOF R²"] = compact["Part 2 weighted OOF R²"].map(lambda x: fmt_or_dash(x, 4))
    compact["Part 2 weighted OOF MAE (m)"] = compact["Part 2 weighted OOF MAE (m)"].map(lambda x: fmt_or_dash(x, 1))
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
    stable["Model_Column"] = stable.apply(
        lambda row: f"{int(row['Overall_Rank'])}. {row['Acronym']} (n={int(row['Input_vars'])})",
        axis=1,
    )
    stable["Cell"] = stable.apply(
        lambda row: f"{float(row['bootstrap_coef_mean']):+.1f}; {row['Selection_Freq_Pct']}",
        axis=1,
    )

    column_order = [
        f"{int(row['Overall_Rank'])}. {row['Acronym']} (n={int(row['Input_vars'])})"
        for _, row in model_explainers.sort_values("Overall_Rank").iterrows()
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


def load_tables() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    model_explainers = pd.read_excel(SOURCE_XLSX, sheet_name="Model_Explainers").sort_values("Overall_Rank")
    panel_a = pd.read_excel(SOURCE_XLSX, sheet_name="Pub_Table1_PanelA")
    panel_b = pd.read_excel(SOURCE_XLSX, sheet_name="Pub_Table1_PanelB")
    table2 = pd.read_excel(SOURCE_XLSX, sheet_name="Pub_Table2")
    lasso_coefficients = pd.read_excel(SOURCE_XLSX, sheet_name="LASSO_Coefficients")
    return model_explainers, panel_a, panel_b, table2, lasso_coefficients


def write_main_document(compact_performance: pd.DataFrame, compact_predictors: pd.DataFrame) -> None:
    doc = Document()
    ensure_landscape(doc)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("Publication Tables")
    run.bold = True
    run.font.size = Pt(12)

    intro = doc.add_paragraph(
        "Compact journal tables are shown here; full model-by-model performance metrics and coefficient details are moved to supplementary tables."
    )
    intro.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in intro.runs:
        run.font.size = Pt(8)

    add_table(
        doc,
        compact_performance,
        title="Table 1. Compact performance summary of the 5 retained models",
        note="R² and MAE are restored here, while sensitivity, specificity, and other detailed operating characteristics are provided in the supplementary tables.",
        font_size=7,
    )
    add_table(
        doc,
        compact_predictors,
        title="Table 2. Stable predictors shared across retained models",
        note="Rows are stable predictors retained in at least 2 models. Cells show bootstrap mean standardized β and selection frequency. Predictors unique to a single model and full confidence-interval details are reported in the supplementary tables.",
        font_size=7,
    )

    legend = doc.add_paragraph()
    legend.add_run("Abbreviations. ").bold = True
    legend.add_run("BalAcc = balanced accuracy; MAE = mean absolute error; OOF = out of fold; SF = selection frequency.")
    for run in legend.runs:
        run.font.size = Pt(8)

    doc.save(OUTPUT_DOCX)


def write_supplementary_document(panel_a: pd.DataFrame, panel_b: pd.DataFrame, table2: pd.DataFrame) -> None:
    doc = Document()
    ensure_landscape(doc)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("Supplementary Tables")
    run.bold = True
    run.font.size = Pt(12)

    add_table(
        doc,
        panel_a,
        title="Supplementary Table S1. Full Part 2 IPCW binary classification performance",
        note="This table retains the full accuracy, balanced accuracy, sensitivity, and specificity results for the Best and Worst scenarios.",
        font_size=7,
    )
    add_table(
        doc,
        panel_b,
        title="Supplementary Table S2. Full publication Table 1 Panel B stable coefficient matrix",
        note="Rows are predictors and columns are retained models. Cells are bootstrap mean coefficient (bootstrap SD) [95% CI]; selection frequency.",
        font_size=7,
    )
    add_table(
        doc,
        table2,
        title="Supplementary Table S3. Full publication Table 2 detailed stable coefficients",
        note="This table retains the detailed coefficient rows for every stable predictor-model pairing from the original publication Table 2 export.",
        font_size=7,
    )

    doc.save(OUTPUT_SUPPLEMENTARY)


def main() -> None:
    model_explainers, panel_a, panel_b, table2, lasso_coefficients = load_tables()
    compact_performance = build_compact_performance_table(model_explainers, panel_a)
    compact_predictors = build_compact_predictor_table(model_explainers, lasso_coefficients)
    write_main_document(compact_performance, compact_predictors)
    write_supplementary_document(panel_a, panel_b, table2)
    print(f"Wrote {OUTPUT_DOCX.name}")
    print(f"Wrote {OUTPUT_SUPPLEMENTARY.name}")


if __name__ == "__main__":
    main()
