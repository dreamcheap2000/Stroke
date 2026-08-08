# Tip-Over (Tipping-Point) Analysis Summary

## Cohort counts
- Total N: **633**
- Completers: **577**
- Non-completers: **56**
- Observed 6MWT4 (all): **511**
- Observed 6MWT4 among completers: **511**

## Base MAR model quantities
- Weighted mean observed 6MWT4 among completers (M_obs): **319.613 m**
- Mean MAR-imputed 6MWT4 among missing outcomes (M_mis_0, all missing): **233.708 m**
- Mean MAR-imputed 6MWT4 among non-completer missing outcomes (M_mis_0, non-completers): **253.147 m**

## Literature-referenced tipping thresholds
See `tipping_point_by_threshold_Literature_Referenced_Threshold.csv` for side-by-side results for:
- p_mis = 56/633 (non-completers only)
- p_mis = 122/633 equivalent to all missing-outcome fraction in this dataset

## Mean-shift vs full pattern-mixture MI
See `Tip_Over_Analysis_results_Literature_Referenced_Threshold.csv` for delta-grid comparisons and discrepancy flags (`>5m`).

## Plot
- `tipping_point_curve_Literature_Referenced_Threshold.png`

## Interpretation
Under the primary scenario, **1** literature threshold is already crossed at the MAR estimate; the remaining thresholds require non-completers' true 6MWT4 to average between **110.0** and **2423.3 meters** lower than their MAR-imputed value. The smallest literature-anchored shift that remains within the observed plausibility range is **110.0 meters**.