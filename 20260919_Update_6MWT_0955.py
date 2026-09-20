from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd


ROOT = Path(__file__).resolve().parent
SOURCE_TABLE_CSV = ROOT / "20260919_6MWT.csv"
SOURCE_PREDICTIONS_XLSX = ROOT / "20260904_Comprehensive_1018.xlsx"
SOURCE_PREDICTIONS_SHEET = "Predictions"

OUTPUT_CSV = ROOT / "20260919_Update_6MWT_0955.csv"
OUTPUT_COMPARISON_MD = ROOT / "20260919_Update_6MWT_0955_Comparison.md"

GROUP_COLS = ["Rehab_LOS_Category", "PAC_Program_Completion", "First_6MWT_TP"]
VALUE_COLS = ["Initial_6MWT_Distance", "6MWT_P1_Change", "6MWT_P2_Change", "6MWT_P3_Change"]

REHAB_ORDER = ["21-42 days", "<=21 days", ">42 days"]
PAC_ORDER = ["Completed PAC program", "Did not complete PAC program"]
FIRST_TP_ORDER = ["Never", "T1", "T2", "T3", "T4"]


def ci95_interval(values: pd.Series) -> str:
    n = len(values)
    if n < 2:
        return "(nan-nan)"
    mean = float(values.mean())
    half_width = float(1.96 * values.std(ddof=1) / np.sqrt(n))
    return f"({mean - half_width:.1f}-{mean + half_width:.1f})"


def iqr_range(values: pd.Series) -> str:
    if values.empty:
        return "(nan-nan)"
    q1 = float(values.quantile(0.25))
    q3 = float(values.quantile(0.75))
    return f"({q1:.1f}-{q3:.1f})"


def summarize_series(values: pd.Series) -> dict[str, object]:
    clean = values.dropna()
    if clean.empty:
        return {
            "mean": np.nan,
            "median": np.nan,
            "iqr_range": "(nan-nan)",
            "ci_95": "(nan-nan)",
            "valid_count": 0,
        }

    return {
        "mean": float(clean.mean()),
        "median": float(clean.median()),
        "iqr_range": iqr_range(clean),
        "ci_95": ci95_interval(clean),
        "valid_count": int(clean.shape[0]),
    }


def build_summary(df: pd.DataFrame) -> pd.DataFrame:
    grouped = df.groupby(GROUP_COLS, dropna=False, sort=False)

    rows: list[dict[str, object]] = []
    for keys, subset in grouped:
        row: dict[str, object] = {
            "Rehab_LOS_Category": keys[0],
            "PAC_Program_Completion": keys[1],
            "First_6MWT_TP": keys[2],
            ("ID", "count"): int(subset["ID"].shape[0]),
        }
        for col in VALUE_COLS:
            stats = summarize_series(subset[col])
            for metric, value in stats.items():
                row[(col, metric)] = value
        rows.append(row)

    summary = pd.DataFrame(rows)

    summary["Rehab_LOS_Category"] = pd.Categorical(summary["Rehab_LOS_Category"], categories=REHAB_ORDER, ordered=True)
    summary["PAC_Program_Completion"] = pd.Categorical(summary["PAC_Program_Completion"], categories=PAC_ORDER, ordered=True)
    summary["First_6MWT_TP"] = pd.Categorical(summary["First_6MWT_TP"], categories=FIRST_TP_ORDER, ordered=True)

    summary = summary.sort_values(GROUP_COLS).set_index(GROUP_COLS)

    ordered_columns = [("ID", "count")]
    for col in VALUE_COLS:
        ordered_columns.extend(
            [
                (col, "mean"),
                (col, "median"),
                (col, "iqr_range"),
                (col, "ci_95"),
                (col, "valid_count"),
            ]
        )

    summary = summary[ordered_columns]
    summary.columns = pd.MultiIndex.from_tuples(summary.columns)
    return summary


def markdown_table(df: pd.DataFrame) -> str:
    cols = list(df.columns)
    header = "| " + " | ".join(cols) + " |"
    sep = "| " + " | ".join(["---"] * len(cols)) + " |"
    body = ["| " + " | ".join(str(v) for v in row) + " |" for row in df.itertuples(index=False, name=None)]
    return "\n".join([header, sep, *body])


def write_summary_csv(summary: pd.DataFrame, output_path: Path) -> None:
    block_header = (
        ["", "", "", "ID"]
        + ["Initial_6MWT_Distance", "", "", "", ""]
        + ["6MWT_P1_Change", "", "", "", ""]
        + ["6MWT_P2_Change", "", "", "", ""]
        + ["6MWT_P3_Change", "", "", "", ""]
    )
    stat_header = (
        ["", "", "", "count"]
        + ["mean", "median", "iqr_range", "ci_95", "valid_count"]
        + ["mean", "median", "iqr_range", "ci_95", "valid_count"]
        + ["mean", "median", "iqr_range", "ci_95", "valid_count"]
        + ["mean", "median", "iqr_range", "ci_95", "valid_count"]
    )
    group_header = ["Rehab_LOS_Category", "PAC_Program_Completion", "First_6MWT_TP"] + [""] * (len(block_header) - 3)

    metric_cols = [("ID", "count")]
    for col in VALUE_COLS:
        metric_cols.extend(
            [
                (col, "mean"),
                (col, "median"),
                (col, "iqr_range"),
                (col, "ci_95"),
                (col, "valid_count"),
            ]
        )

    out = summary.reset_index()
    out.columns = [
        "Rehab_LOS_Category",
        "PAC_Program_Completion",
        "First_6MWT_TP",
        *metric_cols,
    ]

    rows: list[list[object]] = []
    prev_rehab: object = None
    prev_pac: object = None
    for row in out.itertuples(index=False, name=None):
        rehab, pac, tp = row[0], row[1], row[2]
        rehab_val = "" if rehab == prev_rehab else rehab
        pac_val = "" if (rehab == prev_rehab and pac == prev_pac) else pac
        values = [("" if pd.isna(v) else v) for v in row[3:]]
        rows.append([rehab_val, pac_val, tp, *values])
        prev_rehab, prev_pac = rehab, pac

    with output_path.open("w", encoding="utf-8", newline="") as handle:
        import csv

        writer = csv.writer(handle)
        writer.writerow(block_header)
        writer.writerow(stat_header)
        writer.writerow(group_header)
        writer.writerows(rows)


def main() -> None:
    old_table = pd.read_csv(SOURCE_TABLE_CSV, header=[0, 1], index_col=[0, 1, 2])
    old_index = old_table.index.to_frame(index=False)
    old_index["Rehab_LOS_Category"] = old_index["Rehab_LOS_Category"].ffill()
    old_index["PAC_Program_Completion"] = old_index["PAC_Program_Completion"].ffill()
    old_count_col = pd.to_numeric(old_table[("ID", "count")], errors="coerce").fillna(0).reset_index(drop=True)
    old_counts = int(
        old_count_col.loc[
            old_index["PAC_Program_Completion"].eq("Completed PAC program")
            & old_index["First_6MWT_TP"].ne("Never")
        ].sum()
    )

    predictions = pd.read_excel(SOURCE_PREDICTIONS_XLSX, sheet_name=SOURCE_PREDICTIONS_SHEET)

    model_ineligible_mask = (
        predictions["PAC_Program_Completion"].eq("Completed PAC program")
        & predictions["First_6MWT_TP"].ne("Never")
        & predictions["6MWT4"].isna()
    )

    filtered = predictions.loc[~model_ineligible_mask].copy()
    summary = build_summary(filtered)
    write_summary_csv(summary, OUTPUT_CSV)

    model_eligible_count = int(
        (
            predictions["PAC_Program_Completion"].eq("Completed PAC program")
            & predictions["First_6MWT_TP"].ne("Never")
            & predictions["6MWT4"].notna()
        ).sum()
    )

    summary_index = summary.index.to_frame(index=False)
    updated_count_col = pd.to_numeric(summary[("ID", "count")], errors="coerce").fillna(0).reset_index(drop=True)
    updated_count = int(
        updated_count_col.loc[
            summary_index["PAC_Program_Completion"].eq("Completed PAC program")
            & summary_index["First_6MWT_TP"].ne("Never")
        ].sum()
    )

    discrepant = predictions.loc[
        model_ineligible_mask,
        [
            "ID",
            "Rehab_LOS_Category",
            "PAC_Program_Completion",
            "First_6MWT_TP",
            "6MWT4",
            "Initial_6MWT_Distance",
            "6MWT_P1_Change",
            "6MWT_P2_Change",
            "6MWT_P3_Change",
        ],
    ].copy()
    discrepant["Missing_or_excluded_field"] = "6MWT4 missing (model outcome required for eligibility)"
    discrepant = discrepant.sort_values(["Rehab_LOS_Category", "First_6MWT_TP", "ID"]).reset_index(drop=True)

    discrepancy_summary = pd.DataFrame(
        [
            {
                "Metric": "Current 20260919_6MWT.csv PAC completers with First_6MWT_TP != Never",
                "Count": int(old_counts),
            },
            {
                "Metric": "Model-eligible PAC completers from 20260904_Comprehensive_1018.xlsx (6MWT4 non-missing)",
                "Count": model_eligible_count,
            },
            {
                "Metric": "Updated 20260919_Update_6MWT_0955.csv PAC completers with First_6MWT_TP != Never",
                "Count": updated_count,
            },
            {
                "Metric": "Difference (current table - model-eligible)",
                "Count": old_counts - model_eligible_count,
            },
        ]
    )

    lines = [
        "# 6MWT denominator comparison (20260919 update)",
        "",
        "## Dataset lineage used",
        "- `20260912_Tables_1840.py` points to `20260904_Comprehensive_1018.xlsx` (`Predictions` sheet) as the patient-level source for 6MWT4 model denominators.",
        "- This update uses that same patient-level lineage dataset.",
        "",
        "## Count comparison",
        markdown_table(discrepancy_summary),
        "",
        "## Why 517 vs 511",
        "The prior visible table count of 517 includes all PAC completers whose `First_6MWT_TP` is not `Never`, regardless of whether the model outcome (`6MWT4`) is present.",
        "The modeling denominator requires non-missing `6MWT4`. Exactly 6 PAC completers had an initial/first 6MWT time point but missing `6MWT4`, so they are excluded from model-eligible counts (511).",
        "",
        "## Discrepant records (n=6)",
        markdown_table(discrepant.fillna("")),
        "",
    ]
    OUTPUT_COMPARISON_MD.write_text("\n".join(lines), encoding="utf-8")


if __name__ == "__main__":
    main()
