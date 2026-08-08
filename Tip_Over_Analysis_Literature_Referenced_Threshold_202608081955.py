#!/usr/bin/env python3
"""Tip-over (tipping-point) MNAR sensitivity analysis for 6MWT4 using literature-referenced thresholds."""

import os
import sys
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from sklearn.experimental import enable_iterative_imputer  # noqa: F401
from sklearn.impute import IterativeImputer, SimpleImputer
from sklearn.linear_model import LogisticRegression, Ridge
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import StandardScaler


ROOT = Path(__file__).resolve().parent
ARTIFACT_CSV = ROOT / "Tip_Over_Analysis_artifacts_out.csv"


def get_covariates(df: pd.DataFrame) -> list[str]:
    demographics = ["Age", "Sex, F0 M1"]
    acute = [
        "Pneumonia", "UTI", "GIB", "Cellulitis", "StrokeInEvolution", "tPA", "IA", "tPAIA", "Neurology_LOS"
    ]
    stroke_chars = [
        "Dissection", "ACA", "Undetermined", "LVS", "LVO", "Side_Right", "Side_Left", "Side_Bilateral",
        "Loc_CortSub", "Loc_Subcortical", "Loc_Infratentorial"
    ]
    comorbidities = [
        "AF", "DM", "HTN", "Dyslipidemia", "CAD", "CKD", "RestrictiveLung", "GIUlcer", "LiverCirrhosis",
        "Hepatitis", "Parkinsonism", "Malignancy", "OldStroke", "Dementia", "Psychiatric", "Gout"
    ]
    nihss_out = [
        "ConsOut", "AnswerOut", "OrderOut", "EOMOut", "VisualOut", "FaceOut", "LUOut", "RUOut", "LLOut", "RLOut",
        "CoordinateOut", "SensoryOut", "LanguageOut", "ArticulateOut", "NeglectOut"
    ]
    func_t1 = [
        "MRS1", "BI1", "FOIS1", "MNA1", "EuroQoL5D1", "IADL1", "BBS1", "FuglUE1", "FuglSEN1", "CCAT1",
        "6MWT1", "Gait_Speed_1"
    ]
    covars = demographics + acute + stroke_chars + comorbidities + nihss_out + func_t1

    for c in covars:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    return [c for c in covars if c in df.columns]


def load_or_build_artifacts(xlsx_path: Path) -> tuple[pd.DataFrame, list[str]]:
    if ARTIFACT_CSV.exists():
        df = pd.read_csv(ARTIFACT_CSV)
        covars = [c for c in df.columns if c.startswith("cov_")]
        covars = [c.removeprefix("cov_") for c in covars]
        rename_map = {f"cov_{c}": c for c in covars}
        df = df.rename(columns=rename_map)
        return df, covars

    df = pd.read_excel(xlsx_path)
    df["outcome_6mwt4"] = pd.to_numeric(df.get("6MWT4", df.get("6mwt4", np.nan)), errors="coerce")
    completion_status = df["PAC_Program_Completion"]
    df["completer"] = np.where(
        completion_status == "Completed PAC program",
        1.0,
        np.where(completion_status == "Did not complete PAC program", 0.0, np.nan),
    )
    covars = get_covariates(df)

    mask_all = completion_status.notna()
    X_all = df.loc[mask_all, covars]
    y_comp = df.loc[mask_all, "completer"].astype(int).to_numpy()

    denom_model = Pipeline([
        ("imp", SimpleImputer(strategy="median")),
        ("sc", StandardScaler()),
        ("lr", LogisticRegression(max_iter=5000, solver="lbfgs", C=1.0)),
    ])
    denom_model.fit(X_all, y_comp)
    p_denom = np.clip(denom_model.predict_proba(X_all)[:, 1], 1e-4, 1 - 1e-4)

    if "Age" not in X_all.columns:
        raise KeyError("Required covariate 'Age' not found for numerator IPCW model.")

    numer_model = Pipeline([
        ("imp", SimpleImputer(strategy="median")),
        ("sc", StandardScaler()),
        ("lr", LogisticRegression(max_iter=1000, solver="lbfgs", C=1e6)),
    ])
    numer_model.fit(X_all[["Age"]], y_comp)
    p_numer = np.clip(numer_model.predict_proba(X_all[["Age"]])[:, 1], 1e-4, 1 - 1e-4)

    df_work = df.loc[mask_all].copy().reset_index(drop=True)
    df_work["p_denom"] = p_denom
    df_work["p_numer"] = p_numer

    df_work["ipcw_raw"] = np.where(df_work["completer"] == 1, df_work["p_numer"] / df_work["p_denom"], np.nan)
    ipcw_vals = df_work.loc[df_work["completer"] == 1, "ipcw_raw"]
    lo, hi = np.nanquantile(ipcw_vals, [0.01, 0.99])
    df_work["ipcw"] = np.where(df_work["completer"] == 1, np.clip(df_work["ipcw_raw"], lo, hi), np.nan)

    out_df = df_work[["outcome_6mwt4", "completer", "ipcw"] + covars].copy()
    out_df["missing_outcome"] = out_df["outcome_6mwt4"].isna().astype(int)

    save_df = out_df.copy()
    save_df = save_df.rename(columns={c: f"cov_{c}" for c in covars})
    save_df.to_csv(ARTIFACT_CSV, index=False)

    return out_df, covars


def tipping_point(delta: np.ndarray, m_obs: float, m_mis_0: float, p_mis: float) -> np.ndarray:
    return (1 - p_mis) * m_obs + p_mis * (m_mis_0 - delta)


def rubins_rules(q: np.ndarray, u: np.ndarray) -> tuple[float, float]:
    m = len(q)
    q_bar = float(np.mean(q))
    u_bar = float(np.mean(u))
    b = float(np.var(q, ddof=1)) if m > 1 else 0.0
    t = u_bar + (1 + 1 / m) * b
    return q_bar, np.sqrt(max(t, 0.0))


def delta_to_reach_target(target_mean: float, m_obs: float, m_mis_0: float, p_mis: float) -> float:
    return (((1 - p_mis) * m_obs) + (p_mis * m_mis_0) - target_mean) / p_mis


LITERATURE_THRESHOLDS = [
    {
        "threshold_type": "change_based_mcid_general_low",
        "threshold_value": 14.0,
        "citation_label": "Bohannon_Crouch_MCID_low",
        "target_mode": "baseline_minus_value",
    },
    {
        "threshold_type": "change_based_mcid_general_high",
        "threshold_value": 30.5,
        "citation_label": "Bohannon_Crouch_MCID_high",
        "target_mode": "baseline_minus_value",
    },
    {
        "threshold_type": "change_based_mcid_neuro_musculoskeletal",
        "threshold_value": 37.0,
        "citation_label": "Daynes_MetaAnalysis_MID_neuro_msk",
        "target_mode": "baseline_minus_value",
    },
    {
        "threshold_type": "change_based_mic_stroke_subacute_low",
        "threshold_value": 63.0,
        "citation_label": "Kubo_2022_MIC_low",
        "target_mode": "baseline_minus_value",
    },
    {
        "threshold_type": "change_based_mic_stroke_subacute_high",
        "threshold_value": 83.0,
        "citation_label": "Kubo_2022_MIC_high",
        "target_mode": "baseline_minus_value",
    },
    {
        "threshold_type": "absolute_distance_cutoff_independence_kubo",
        "threshold_value": 304.0,
        "citation_label": "Kubo_2022_FAC_independence_cutoff",
        "target_mode": "absolute_value",
    },
    {
        "threshold_type": "absolute_distance_cutoff_community_blennerhassett",
        "threshold_value": 367.0,
        "citation_label": "Blennerhassett_community_ambulation_cutoff",
        "target_mode": "absolute_value",
    },
    {
        "threshold_type": "absolute_distance_cutoff_fac3_to_fac4_lee",
        "threshold_value": 99.35,
        "citation_label": "Lee_FAC3_to_FAC4_cutoff",
        "target_mode": "absolute_value",
    },
]
PLAUSIBILITY_FAC2_FLOOR_M = 140.0


def resolve_target_mean(reference_mean: float, threshold: dict[str, float | str]) -> float:
    if threshold["target_mode"] == "baseline_minus_value":
        return reference_mean - float(threshold["threshold_value"])
    return float(threshold["threshold_value"])


def classify_plausibility(true_mean_at_tip: float) -> str:
    if true_mean_at_tip < 0:
        return "physically_impossible_below_0m"
    if true_mean_at_tip < PLAUSIBILITY_FAC2_FLOOR_M:
        return "clinically_implausible_below_fac2_mean"
    return "within_observed_range_clinically_plausible"


def main() -> None:
    xlsx_arg = sys.argv[1] if len(sys.argv) > 1 else os.getenv("STROKE_XLSX_PATH")
    xlsx_path = Path(xlsx_arg) if xlsx_arg else ROOT / "20260806_DeID.xlsx"
    if not xlsx_path.exists():
        raise FileNotFoundError(f"Cannot find input XLSX: {xlsx_path}")

    df, covars = load_or_build_artifacts(xlsx_path)
    df = df.reset_index(drop=True)

    n_total = len(df)
    n_completers = int((df["completer"] == 1).sum())
    n_non_completers = int((df["completer"] == 0).sum())
    n_observed_all = int(df["outcome_6mwt4"].notna().sum())
    n_observed_completers = int(((df["completer"] == 1) & df["outcome_6mwt4"].notna()).sum())

    comp_obs_mask = (df["completer"] == 1) & df["outcome_6mwt4"].notna()
    train_df = df.loc[comp_obs_mask].copy()

    imp = SimpleImputer(strategy="median")
    X_train = pd.DataFrame(imp.fit_transform(train_df[covars]), columns=covars)
    y_train = train_df["outcome_6mwt4"].to_numpy()
    w_train = train_df["ipcw"].to_numpy()

    base_model = Ridge(alpha=1.0)
    base_model.fit(X_train, y_train, sample_weight=w_train)

    miss_mask = df["outcome_6mwt4"].isna()
    X_missing = pd.DataFrame(imp.transform(df.loc[miss_mask, covars]), columns=covars)
    mar_pred_missing = base_model.predict(X_missing)

    m_obs = float(np.average(y_train, weights=w_train))
    m_mis_0_all = float(np.mean(mar_pred_missing))
    miss_noncomp_mask = miss_mask & (df["completer"] == 0)
    X_missing_noncomp = pd.DataFrame(imp.transform(df.loc[miss_noncomp_mask, covars]), columns=covars)
    m_mis_0_noncomp = float(np.mean(base_model.predict(X_missing_noncomp)))

    p_mis_noncomp = n_non_completers / n_total
    p_mis_all_missing = float(miss_mask.mean())

    deltas = np.arange(0, 151, 5)

    m_total_noncomp = tipping_point(deltas, m_obs, m_mis_0_noncomp, p_mis_noncomp)
    m_total_all_missing = tipping_point(deltas, m_obs, m_mis_0_all, p_mis_all_missing)

    mar_population_mean_noncomp = tipping_point(np.array([0.0]), m_obs, m_mis_0_noncomp, p_mis_noncomp)[0]
    mar_population_mean_all_missing = tipping_point(np.array([0.0]), m_obs, m_mis_0_all, p_mis_all_missing)[0]

    tip_rows = []
    for scenario_name, p_mis in [
        ("pmis_noncompleters_56_633", p_mis_noncomp),
        ("pmis_all_missing", p_mis_all_missing),
    ]:
        m_mis_0 = m_mis_0_noncomp if scenario_name == "pmis_noncompleters_56_633" else m_mis_0_all
        reference_mean = (
            mar_population_mean_noncomp
            if scenario_name == "pmis_noncompleters_56_633"
            else mar_population_mean_all_missing
        )
        for threshold in LITERATURE_THRESHOLDS:
            target_mean = resolve_target_mean(reference_mean, threshold)
            delta_tip = delta_to_reach_target(target_mean, m_obs, m_mis_0, p_mis)
            true_mean_at_tip = m_mis_0 - delta_tip
            tip_rows.append(
                {
                    "scenario": scenario_name,
                    "threshold_type": threshold["threshold_type"],
                    "threshold_value": round(float(threshold["threshold_value"]), 3),
                    "citation_label": threshold["citation_label"],
                    "delta_tip": round(float(delta_tip), 3),
                    "true_mean_at_tip": round(float(true_mean_at_tip), 3),
                    "plausibility_flag": classify_plausibility(float(true_mean_at_tip)),
                }
            )

    tip_df = pd.DataFrame(tip_rows)

    # Step 5: plot for primary scenario (non-completer fraction)
    fig, ax = plt.subplots(figsize=(12, 7))
    ax.plot(deltas, m_total_noncomp, marker="o", label="Mtotal(delta), p_mis=56/633")
    for threshold in LITERATURE_THRESHOLDS:
        target_mean = resolve_target_mean(mar_population_mean_noncomp, threshold)
        ax.axhline(y=target_mean, linestyle="--", linewidth=1)
        delta_cross = delta_to_reach_target(target_mean, m_obs, m_mis_0_noncomp, p_mis_noncomp)
        if 0 <= delta_cross <= 150:
            ax.scatter([delta_cross], [target_mean], s=40)
            ax.text(delta_cross + 1, target_mean + 1, f"{threshold['threshold_type']}: Δ={delta_cross:.1f}", fontsize=8)

    ax.set_xlabel("Assumed extra deficit among missing/non-completers (delta, meters)")
    ax.set_ylabel("Population mean 6MWT4 (meters)")
    ax.set_title("Tipping-point curve with literature-referenced thresholds")
    ax.grid(alpha=0.25)
    ax.legend()
    fig.tight_layout()
    plot_path = ROOT / "tipping_point_curve_Literature_Referenced_Threshold_202608081955.png"
    fig.savefig(plot_path, dpi=200)
    plt.close(fig)

    # Step 6: full pattern-mixture MI (m=50), shifting non-completer imputations
    m_imputations = 50
    mi_cols = covars + ["outcome_6mwt4"]
    mi_input = df[mi_cols].copy()
    noncomp_missing_mask = (df["completer"] == 0) & miss_mask
    noncomp_missing_idx = noncomp_missing_mask.to_numpy()
    if not mi_input.index.equals(df.index):
        raise ValueError("MI input and working dataframe indices must stay aligned.")

    imputed_outcomes = []
    for m in range(m_imputations):
        imputer = IterativeImputer(max_iter=15, sample_posterior=True, random_state=42 + m, initial_strategy="median")
        x_full = pd.DataFrame(imputer.fit_transform(mi_input), columns=mi_cols)
        imputed_outcomes.append(x_full["outcome_6mwt4"].to_numpy())

    x_model = pd.DataFrame(imp.transform(df[covars]), columns=covars)
    noncomp_idx = (df["completer"] == 0).to_numpy()

    mi_rows = []
    for delta in deltas:
        q_vals = []
        u_vals = []
        for y_imp in imputed_outcomes:
            y_shift = y_imp.copy()
            y_shift[noncomp_missing_idx] = y_shift[noncomp_missing_idx] - delta

            model_mi = Ridge(alpha=1.0)
            model_mi.fit(x_model, y_shift)
            pred_noncomp = model_mi.predict(x_model.loc[noncomp_idx])
            m_noncomp_m = float(np.mean(pred_noncomp))
            q_m = float((1 - p_mis_noncomp) * m_obs + p_mis_noncomp * m_noncomp_m)
            u_m = float(np.var(pred_noncomp, ddof=1) / max(pred_noncomp.shape[0], 1))
            q_vals.append(q_m)
            u_vals.append(u_m)

        pooled_mean, pooled_se = rubins_rules(np.array(q_vals), np.array(u_vals))
        mi_rows.append({"delta": delta, "pattern_mixture_mi_pooled_mean": pooled_mean, "pattern_mixture_mi_pooled_se": pooled_se})

    mi_df = pd.DataFrame(mi_rows)

    # Step 7: compare mean-shift vs full MI (primary p_mis scenario)
    compare_df = pd.DataFrame({
        "delta": deltas,
        "mean_shift_population_mean": m_total_noncomp,
    }).merge(mi_df, on="delta", how="left")
    compare_df["abs_diff_m"] = (compare_df["mean_shift_population_mean"] - compare_df["pattern_mixture_mi_pooled_mean"]).abs()
    compare_df["discrepancy_gt_5m"] = compare_df["abs_diff_m"] > 5

    # Save outputs
    counts_df = pd.DataFrame([
        {"metric": "N_total", "value": n_total},
        {"metric": "N_completers", "value": n_completers},
        {"metric": "N_non_completers", "value": n_non_completers},
        {"metric": "N_observed_6MWT4_all", "value": n_observed_all},
        {"metric": "N_observed_6MWT4_completers", "value": n_observed_completers},
        {"metric": "M_obs_weighted", "value": m_obs},
        {"metric": "M_mis0_MAR_predicted_all_missing", "value": m_mis_0_all},
        {"metric": "M_mis0_MAR_predicted_noncompleters", "value": m_mis_0_noncomp},
        {"metric": "p_mis_noncompleters", "value": p_mis_noncomp},
        {"metric": "p_mis_all_missing", "value": p_mis_all_missing},
    ])

    counts_path = ROOT / "Tip_Over_Analysis_counts_Literature_Referenced_Threshold_202608081955.csv"
    tip_path = ROOT / "tipping_point_by_threshold_Literature_Referenced_Threshold_202608081955.csv"
    compare_path = ROOT / "Tip_Over_Analysis_results_Literature_Referenced_Threshold_202608081955.csv"

    counts_df.to_csv(counts_path, index=False)
    tip_df.to_csv(tip_path, index=False)
    compare_df.to_csv(compare_path, index=False)

    # Step 9 summary markdown
    primary_rows = tip_df[tip_df["scenario"] == "pmis_noncompleters_56_633"]
    n_missing_total = n_total - n_observed_all
    nonnegative_primary = primary_rows[primary_rows["delta_tip"] >= 0]
    already_crossed_count = int((primary_rows["delta_tip"] < 0).sum())
    min_delta = float(nonnegative_primary["delta_tip"].min()) if not nonnegative_primary.empty else float("nan")
    max_delta = float(nonnegative_primary["delta_tip"].max()) if not nonnegative_primary.empty else float("nan")
    plausible_primary = nonnegative_primary[
        nonnegative_primary["plausibility_flag"] == "within_observed_range_clinically_plausible"
    ]
    plausible_text = (
        f" The smallest literature-anchored shift that remains within the observed plausibility range is **{float(plausible_primary['delta_tip'].min()):.1f} meters**."
        if not plausible_primary.empty
        else " No non-negative threshold crossings remain within the observed plausibility range."
    )
    crossed_label = "threshold is" if already_crossed_count == 1 else "thresholds are"

    md_lines = [
        "# Tip-Over (Tipping-Point) Analysis Summary",
        "",
        "## Cohort counts",
        f"- Total N: **{n_total}**",
        f"- Completers: **{n_completers}**",
        f"- Non-completers: **{n_non_completers}**",
        f"- Observed 6MWT4 (all): **{n_observed_all}**",
        f"- Observed 6MWT4 among completers: **{n_observed_completers}**",
        "",
        "## Base MAR model quantities",
        f"- Weighted mean observed 6MWT4 among completers (M_obs): **{m_obs:.3f} m**",
        f"- Mean MAR-imputed 6MWT4 among missing outcomes (M_mis_0, all missing): **{m_mis_0_all:.3f} m**",
        f"- Mean MAR-imputed 6MWT4 among non-completer missing outcomes (M_mis_0, non-completers): **{m_mis_0_noncomp:.3f} m**",
        "",
        "## Literature-referenced tipping thresholds",
        "See `tipping_point_by_threshold_Literature_Referenced_Threshold.csv` for side-by-side results for:",
        f"- p_mis = {n_non_completers}/{n_total} (non-completers only)",
        f"- p_mis = {n_missing_total}/{n_total} equivalent to all missing-outcome fraction in this dataset",
        "",
        "## Mean-shift vs full pattern-mixture MI",
        "See `Tip_Over_Analysis_results_Literature_Referenced_Threshold.csv` for delta-grid comparisons and discrepancy flags (`>5m`).",
        "",
        "## Plot",
        "- `tipping_point_curve_Literature_Referenced_Threshold.png`",
        "",
        "## Interpretation",
        (
            f"Under the primary scenario, **{already_crossed_count}** literature {crossed_label} already crossed at the MAR estimate; "
            f"the remaining thresholds require non-completers' true 6MWT4 to average between **{min_delta:.1f}** and **{max_delta:.1f} meters** lower than their MAR-imputed value."
            f"{plausible_text}"
        ),
    ]

    (ROOT / "Tip_Over_Analysis_summary_Literature_Referenced_Threshold_202608081955.md").write_text("\n".join(md_lines), encoding="utf-8")

    print("Tip-over analysis complete.")
    print(f"N={n_total}, completers={n_completers}, non-completers={n_non_completers}, observed-outcome={n_observed_completers}")
    print(f"M_obs={m_obs:.3f}, M_mis_0_all={m_mis_0_all:.3f}, M_mis_0_noncomp={m_mis_0_noncomp:.3f}")
    print(f"Saved: {counts_path.name}, {tip_path.name}, {compare_path.name}, tipping_point_curve_Literature_Referenced_Threshold.png, Tip_Over_Analysis_summary_Literature_Referenced_Threshold.md")


if __name__ == "__main__":
    main()
