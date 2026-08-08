# Tip-Over (Tipping-Point) Analysis Summary

## Cohort counts
- Total N: **633**
- Completers: **577**
- Non-completers: **56**
- Observed 6MWT4 (all): **511**
- Observed 6MWT4 among completers: **511**

## Base MAR model quantities
- Weighted mean observed 6MWT4 among completers (M_obs): **319.917 m**
- Mean MAR-imputed 6MWT4 among missing outcomes (M_mis_0, all missing): **202.613 m**
- Mean MAR-imputed 6MWT4 among non-completer missing outcomes (M_mis_0, non-completers): **244.404 m**

## Delta-tip tables
See `Tip_Over_Analysis_delta_tip_table.csv` for side-by-side results for:
- p_mis = 56/633 (non-completers only)
- p_mis = 122/633 equivalent to all missing-outcome fraction in this dataset

## Mean-shift vs full pattern-mixture MI
See `Tip_Over_Analysis_results.csv` for delta-grid comparisons and discrepancy flags (`>5m`).

## Plot
- `tipping_point_curve.png`

## Interpretation
The primary estimate is robust to MNAR bias unless non-completers' true 6MWT4 is on average at least **1223.5 meters** lower than their MAR-imputed value.