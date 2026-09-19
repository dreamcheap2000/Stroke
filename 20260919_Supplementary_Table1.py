#!/usr/bin/env python3
from pathlib import Path

import numpy as np
import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from scipy.stats import fisher_exact, mannwhitneyu, ttest_ind

ROOT = Path(__file__).resolve().parent
INPUT_XLSX = ROOT / "20260826_DeID.xlsx"
OUTPUT_DOCX = ROOT / "20260919_Supplementary_Table1.docx"

COMPLETION_COL = "PAC_Program_Completion"
COMPLETER_LABEL = "Completed PAC program"

IN_NIHSS_COLS = [
    "ConsIn", "AnswerIn", "OrderIn", "EOMIn", "VisualIn", "FaceIn", "LUIn", "RUIn",
    "LLIn", "RLIn", "CoordinateIn", "SensoryIn", "LanguageIn", "ArticulateIn", "NeglectIn",
]
OUT_NIHSS_COLS = [
    "ConsOut", "AnswerOut", "OrderOut", "EOMOut", "VisualOut", "FacialOut", "LUOut", "RUOut",
    "LLOut", "RLOut", "Coordinateout", "SensoryOut", "LanguageOut", "ArticulateOut", "NeglectOut",
]


def fmt_p(value: float) -> str:
    if pd.isna(value):
        return "—"
    if value < 0.001:
        return "<0.001"
    return f"{value:.3f}"


def fmt_mean_sd(series: pd.Series) -> str:
    x = pd.to_numeric(series, errors="coerce").dropna()
    if x.empty:
        return "—"
    return f"{x.mean():.1f} ± {x.std(ddof=0):.1f}"


def fmt_median_iqr(series: pd.Series) -> str:
    x = pd.to_numeric(series, errors="coerce").dropna()
    if x.empty:
        return "—"
    q1 = x.quantile(0.25)
    q3 = x.quantile(0.75)
    return f"{x.median():.1f} [{q1:.1f}, {q3:.1f}]"


def fmt_n_pct(binary_series: pd.Series) -> str:
    x = pd.to_numeric(binary_series, errors="coerce")
    denom = x.notna().sum()
    if denom == 0:
        return "—"
    num = int((x == 1).sum())
    return f"{num} ({100.0 * num / denom:.1f}%)"


def p_ttest(a: pd.Series, b: pd.Series) -> float:
    aa = pd.to_numeric(a, errors="coerce").dropna()
    bb = pd.to_numeric(b, errors="coerce").dropna()
    if aa.empty or bb.empty:
        return np.nan
    return float(ttest_ind(aa, bb, equal_var=False, nan_policy="omit").pvalue)


def p_mannwhitney(a: pd.Series, b: pd.Series) -> float:
    aa = pd.to_numeric(a, errors="coerce").dropna()
    bb = pd.to_numeric(b, errors="coerce").dropna()
    if aa.empty or bb.empty:
        return np.nan
    return float(mannwhitneyu(aa, bb, alternative="two-sided").pvalue)


def p_fisher(binary_a: pd.Series, binary_b: pd.Series) -> float:
    aa = pd.to_numeric(binary_a, errors="coerce")
    bb = pd.to_numeric(binary_b, errors="coerce")
    a1 = int((aa == 1).sum())
    a0 = int((aa == 0).sum())
    b1 = int((bb == 1).sum())
    b0 = int((bb == 0).sum())
    if (a1 + a0 == 0) or (b1 + b0 == 0):
        return np.nan
    return float(fisher_exact([[a1, a0], [b1, b0]], alternative="two-sided").pvalue)


def main() -> None:
    df = pd.read_excel(INPUT_XLSX)
    comp_mask = df[COMPLETION_COL].eq(COMPLETER_LABEL)

    # Derived binary variables
    df["Male"] = pd.to_numeric(df["Sex, F0 M1"], errors="coerce")
    df["IschemicStroke"] = np.where(pd.to_numeric(df["HemorrhageStroke"], errors="coerce").isna(), np.nan,
                                      (pd.to_numeric(df["HemorrhageStroke"], errors="coerce") == 0).astype(int))
    df["GIB_Present"] = np.where(pd.to_numeric(df["GIB"], errors="coerce").isna(), np.nan,
                                   (pd.to_numeric(df["GIB"], errors="coerce") > 0).astype(int))

    in_total = df[IN_NIHSS_COLS].apply(pd.to_numeric, errors="coerce").sum(axis=1, min_count=1)
    out_total = df[OUT_NIHSS_COLS].apply(pd.to_numeric, errors="coerce").sum(axis=1, min_count=1)

    comp = df.loc[comp_mask].copy()
    non = df.loc[~comp_mask].copy()

    rows = [
        (
            "Baseline Age (years, mean ± SD)",
            fmt_mean_sd(comp["Age"]),
            fmt_mean_sd(non["Age"]),
            fmt_p(p_ttest(comp["Age"], non["Age"])),
        ),
        (
            "Male Sex (n, %)",
            fmt_n_pct(comp["Male"]),
            fmt_n_pct(non["Male"]),
            fmt_p(p_fisher(comp["Male"], non["Male"])),
        ),
        (
            "Ischemic Stroke Etiology (n, %)",
            fmt_n_pct(comp["IschemicStroke"]),
            fmt_n_pct(non["IschemicStroke"]),
            fmt_p(p_fisher(comp["IschemicStroke"], non["IschemicStroke"])),
        ),
        (
            "Baseline Barthel Index (median, IQR)",
            fmt_median_iqr(comp["BI1"]),
            fmt_median_iqr(non["BI1"]),
            fmt_p(p_mannwhitney(comp["BI1"], non["BI1"])),
        ),
        (
            "Baseline Berg Balance Scale (mean ± SD)",
            fmt_mean_sd(comp["BBS1"]),
            fmt_mean_sd(non["BBS1"]),
            fmt_p(p_ttest(comp["BBS1"], non["BBS1"])),
        ),
        (
            "Acute Admission NIHSS (median, IQR)",
            fmt_median_iqr(in_total.loc[comp.index]),
            fmt_median_iqr(in_total.loc[non.index]),
            fmt_p(p_mannwhitney(in_total.loc[comp.index], in_total.loc[non.index])),
        ),
        (
            "Acute Discharge NIHSS (median, IQR)",
            fmt_median_iqr(out_total.loc[comp.index]),
            fmt_median_iqr(out_total.loc[non.index]),
            fmt_p(p_mannwhitney(out_total.loc[comp.index], out_total.loc[non.index])),
        ),
        (
            "Gastrointestinal Bleeding (GIB) (n, %)",
            fmt_n_pct(comp["GIB_Present"]),
            fmt_n_pct(non["GIB_Present"]),
            fmt_p(p_fisher(comp["GIB_Present"], non["GIB_Present"])),
        ),
    ]

    doc = Document()
    title = doc.add_paragraph("Supplementary Table 1. Baseline comparison of completers vs non-completers")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.runs[0].bold = True
    title.runs[0].font.size = Pt(11)

    n_comp = len(comp)
    n_non = len(non)
    table = doc.add_table(rows=1, cols=4)
    table.style = "Table Grid"

    headers = ["Characteristic", f"Completers (n={n_comp})", f"Non-completers (n={n_non})", "P value"]
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = h
        for run in cell.paragraphs[0].runs:
            run.bold = True
            run.font.size = Pt(9)

    for r in rows:
        cells = table.add_row().cells
        for i, val in enumerate(r):
            cells[i].text = str(val)
            for run in cells[i].paragraphs[0].runs:
                run.font.size = Pt(9)

    doc.add_paragraph(
        "Continuous mean±SD variables were compared with Welch t-test; median [IQR] variables with Mann–Whitney U test; categorical variables with Fisher's exact test. "
        "Ischemic stroke etiology was derived as HemorrhageStroke=0. Admission/discharge NIHSS were computed as sums of NIHSS item scores (In/Out components)."
    )

    doc.save(OUTPUT_DOCX)
    print(f"Saved: {OUTPUT_DOCX}")


if __name__ == "__main__":
    main()
