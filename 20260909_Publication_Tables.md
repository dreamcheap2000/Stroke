## Table 1. Retained 20260909 models: binary performance and stable Part 1 predictors

### Panel A. Part 2 IPCW binary classification performance

| Rank | Acronym | Model | Publication name | Input vars | Best Acc [95% CI] | Best BalAcc [95% CI] | Best Se [95% CI] | Best Sp [95% CI] | Worst Acc [95% CI] | Worst BalAcc [95% CI] | Worst Se [95% CI] | Worst Sp [95% CI] |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| 1 | RESTORE | Model 5 | REhabilitation STroke Outcome Recovery Estimator | 21 | 82.5% [71.6, 88.3] | 80.6% [76.5, 86.4] | 82.9% [70.0, 89.6] | 78.3% [70.0, 93.3] | 86.6% [75.0, 89.6] | 82.8% [78.9, 87.1] | 88.1% [72.4, 91.3] | 77.4% [70.4, 90.7] |
| 2 | COMPASS | Model 8 | COMPrehensive Post-Acute Stroke Score | 44 | 68.7% [63.0, 82.6] | 77.5% [73.8, 82.9] | 66.7% [60.0, 83.4] | 88.3% [71.4, 96.7] | 75.8% [64.3, 81.2] | 78.3% [74.8, 82.7] | 74.8% [59.3, 81.8] | 81.7% [73.3, 94.4] |
| 3 | CASCADE | Model 2 | Clinical And Stroke Complexity Assessment for Discharge Endurance | 59 | 72.4% [63.7, 80.1] | 78.8% [74.6, 83.8] | 70.9% [60.8, 80.0] | 86.7% [75.5, 96.4] | 72.7% [61.0, 77.3] | 74.2% [70.5, 78.8] | 72.0% [56.1, 76.9] | 76.3% [68.5, 92.1] |
| 4 | AIMS | Model 1 | Admission Impairment Mobility Score | 12 | 82.0% [69.2, 85.0] | 81.8% [77.4, 86.8] | 82.0% [67.0, 85.1] | 81.7% [73.8, 94.6] | 77.3% [72.8, 86.4] | 81.3% [78.1, 85.5] | 75.6% [69.8, 87.6] | 87.1% [74.7, 94.7] |
| 5 | BEDSIDE | Model 4 | Balance-Enhanced Demographic Speed Index for Discharge Endurance | 4 | 76.8% [73.5, 85.8] | 82.0% [77.8, 86.8] | 75.6% [72.0, 86.0] | 88.3% [75.6, 95.2] | 80.9% [74.4, 85.8] | 83.0% [79.6, 87.1] | 80.0% [71.6, 86.4] | 86.0% [78.3, 94.4] |

### Panel B. Part 1 stable bootstrap LASSO coefficients (selection frequency ≥70%)

| Predictor | RESTORE (Model 5; 21 vars) | COMPASS (Model 8; 44 vars) | CASCADE (Model 2; 59 vars) | AIMS (Model 1; 12 vars) | BEDSIDE (Model 4; 4 vars) |
| --- | --- | --- | --- | --- | --- |
| BBS1 | +77.72 (10.98) [57.21, 98.78]; 100% | +62.54 (10.00) [42.40, 81.97]; 100% | +63.94 (9.52) [45.64, 81.84]; 100% | +61.23 (9.53) [42.55, 79.89]; 100% | +68.60 (8.64) [52.39, 86.17]; 100% |
| Age | -32.68 (6.55) [-45.10, -20.00]; 100% | -33.85 (6.37) [-45.78, -21.18]; 100% | -33.53 (6.22) [-45.19, -20.87]; 100% | -39.64 (6.32) [-52.01, -27.24]; 100% | -35.10 (5.92) [-46.15, -23.68]; 100% |
| Gait_Speed_1_Imputed | +35.60 (9.03) [18.88, 54.17]; 100% | +30.26 (8.63) [13.79, 47.14]; 100% | +28.54 (8.50) [12.01, 45.78]; 100% | +33.06 (8.46) [16.66, 49.35]; 100% | +38.29 (8.29) [20.56, 53.03]; 100% |
| Sex, F0 M1 | +16.86 (5.19) [7.29, 27.10]; 100% | +15.90 (5.16) [6.10, 25.81]; 100% | +15.85 (5.14) [6.07, 25.75]; 100% | +15.51 (5.24) [5.19, 25.25]; 100% | +17.59 (5.35) [7.43, 27.98]; 100% |
| FuglUE1 | +27.72 (8.43) [11.02, 43.70]; 100% | +21.26 (6.67) [7.80, 34.22]; 100% | +17.06 (7.30) [2.63, 31.61]; 99% | +23.74 (7.18) [10.07, 37.45]; 100% | — |
| EuroQoL5D1 | +24.16 (8.19) [9.05, 41.29]; 100% | +8.18 (5.35) [0.00, 18.52]; 92% | +8.49 (5.58) [0.00, 19.64]; 91% | +10.67 (5.68) [0.00, 21.48]; 95% | — |
| MNA1 | -10.94 (6.00) [-22.38, 0.00]; 96% | -14.68 (5.77) [-26.00, -3.16]; 99% | -14.96 (5.65) [-25.87, -3.05]; 99% | -11.89 (5.42) [-22.08, 0.00]; 97% | — |
| MRS1 | -16.46 (7.94) [-31.61, -0.92]; 98% | -9.56 (6.27) [-21.49, 0.00]; 91% | -9.21 (6.49) [-22.46, 0.00]; 89% | -6.42 (6.08) [-19.72, 0.54]; 82% | — |
| IADL1 | +0.46 (5.92) [-10.61, 13.35]; 77% | -5.48 (5.09) [-16.65, 0.00]; 78% | -4.48 (4.91) [-15.97, 0.25]; 71% | -5.10 (5.20) [-16.64, 1.53]; 77% | — |
| GIB | — | +13.70 (4.31) [5.26, 22.12]; 100% | +14.84 (5.10) [5.10, 24.19]; 100% | — | — |
| OldStroke | — | -14.49 (5.38) [-24.60, -3.41]; 100% | -14.82 (5.27) [-25.78, -4.86]; 100% | — | — |
| Dyslipidemia | — | +9.62 (5.41) [0.00, 20.17]; 96% | +9.11 (5.17) [0.00, 19.70]; 96% | — | — |
| Loc_Subcortical | — | -11.13 (5.77) [-21.97, 0.00]; 96% | -9.32 (5.50) [-20.20, 0.00]; 94% | — | — |
| CKD | — | -6.79 (4.21) [-14.90, 0.00]; 93% | -6.77 (4.67) [-16.54, 0.00]; 89% | — | — |
| Dementia | — | -6.01 (3.84) [-13.66, 0.00]; 93% | -4.39 (3.63) [-12.39, 0.00]; 84% | — | — |
| ACA | — | +6.88 (5.31) [0.00, 18.22]; 87% | +8.23 (5.87) [0.00, 21.71]; 91% | — | — |
| Side_Right | — | +10.17 (6.37) [0.00, 22.32]; 89% | +11.21 (7.93) [0.00, 26.91]; 87% | — | — |
| Undetermined | — | -5.60 (4.25) [-14.26, 0.00]; 87% | -5.68 (4.14) [-14.66, 0.00]; 88% | — | — |
| CAD | — | +7.21 (5.63) [0.00, 18.35]; 87% | +5.30 (4.62) [0.00, 15.14]; 80% | — | — |
| Cellulitis | — | -4.85 (4.03) [-13.65, 0.00]; 83% | -5.37 (4.21) [-14.37, 0.00]; 86% | — | — |
| LVS | — | +5.55 (4.95) [-0.68, 15.83]; 84% | +4.85 (4.66) [-0.22, 15.12]; 80% | — | — |
| BBS_T1T2_Change | +32.51 (7.95) [17.14, 48.24]; 100% | — | — | — | — |
| MRS_T1T2_Change | -24.47 (8.79) [-41.33, -8.29]; 99% | — | — | — | — |
| IADL_T1T2_Change | +14.76 (6.54) [1.88, 27.87]; 99% | — | — | — | — |
| EOMOut | — | — | +14.31 (6.43) [1.74, 26.68]; 99% | — | — |
| FuglUE_T1T2_Change | +13.52 (6.23) [0.24, 25.45]; 98% | — | — | — | — |
| EuroQoL5D_T1T2_Change | +14.08 (7.54) [0.00, 29.30]; 96% | — | — | — | — |
| MNA_T1T2_Change | +8.10 (5.07) [-0.18, 18.34]; 95% | — | — | — | — |
| BI_T1T2_Change | -10.27 (8.06) [-27.60, 0.28]; 90% | — | — | — | — |
| FuglSEN_T1T2_Change | -5.89 (5.36) [-17.09, 3.03]; 89% | — | — | — | — |
| FuglSEN1 | -4.67 (5.58) [-16.48, 4.58]; 83% | — | — | — | — |
| LLOut | — | — | -8.56 (7.15) [-23.20, 0.00]; 80% | — | — |
| NeglectOut | — | — | -1.03 (6.75) [-14.16, 13.09]; 79% | — | — |
| FOIS_T1T2_Change | +3.50 (6.86) [-8.42, 19.73]; 76% | — | — | — | — |
| Gout | — | +3.83 (4.77) [-3.68, 14.75]; 76% | — | — | — |
| LanguageOut | — | — | +4.90 (4.95) [0.00, 16.35]; 74% | — | — |
| FOIS1 | — | — | +5.30 (5.49) [0.00, 18.28]; 73% | — | — |
| RLOut | — | — | -6.84 (7.27) [-23.95, 0.00]; 72% | — | — |
| Psychiatric | — | -2.31 (2.66) [-8.77, 0.43]; 70% | — | — | — |

## Table 2. Detailed stable Part 1 coefficients

| Rank | Acronym | Model_label | Publication name | Input vars | Predictor | Full-fit Coef | Boot Mean Coef | Boot SD | 95% CI | Sel Freq |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| 1 | RESTORE | Model 5 | REhabilitation STroke Outcome Recovery Estimator | 21 | Age | -32.46 | -32.68 | 6.55 | [-45.10, -20.00] | 100% |
| 1 | RESTORE | Model 5 | REhabilitation STroke Outcome Recovery Estimator | 21 | BBS1 | +78.31 | +77.72 | 10.98 | [57.21, 98.78] | 100% |
| 1 | RESTORE | Model 5 | REhabilitation STroke Outcome Recovery Estimator | 21 | BBS_T1T2_Change | +33.20 | +32.51 | 7.95 | [17.14, 48.24] | 100% |
| 1 | RESTORE | Model 5 | REhabilitation STroke Outcome Recovery Estimator | 21 | BI_T1T2_Change | -9.54 | -10.27 | 8.06 | [-27.60, 0.28] | 90% |
| 1 | RESTORE | Model 5 | REhabilitation STroke Outcome Recovery Estimator | 21 | EuroQoL5D1 | +25.19 | +24.16 | 8.19 | [9.05, 41.29] | 100% |
| 1 | RESTORE | Model 5 | REhabilitation STroke Outcome Recovery Estimator | 21 | EuroQoL5D_T1T2_Change | +15.40 | +14.08 | 7.54 | [0.00, 29.30] | 96% |
| 1 | RESTORE | Model 5 | REhabilitation STroke Outcome Recovery Estimator | 21 | FOIS_T1T2_Change | +1.42 | +3.50 | 6.86 | [-8.42, 19.73] | 76% |
| 1 | RESTORE | Model 5 | REhabilitation STroke Outcome Recovery Estimator | 21 | FuglSEN1 | -4.51 | -4.67 | 5.58 | [-16.48, 4.58] | 83% |
| 1 | RESTORE | Model 5 | REhabilitation STroke Outcome Recovery Estimator | 21 | FuglSEN_T1T2_Change | -5.05 | -5.89 | 5.36 | [-17.09, 3.03] | 89% |
| 1 | RESTORE | Model 5 | REhabilitation STroke Outcome Recovery Estimator | 21 | FuglUE1 | +27.50 | +27.72 | 8.43 | [11.02, 43.70] | 100% |
| 1 | RESTORE | Model 5 | REhabilitation STroke Outcome Recovery Estimator | 21 | FuglUE_T1T2_Change | +14.24 | +13.52 | 6.23 | [0.24, 25.45] | 98% |
| 1 | RESTORE | Model 5 | REhabilitation STroke Outcome Recovery Estimator | 21 | Gait_Speed_1_Imputed | +35.32 | +35.60 | 9.03 | [18.88, 54.17] | 100% |
| 1 | RESTORE | Model 5 | REhabilitation STroke Outcome Recovery Estimator | 21 | IADL1 | -0.00 | +0.46 | 5.92 | [-10.61, 13.35] | 77% |
| 1 | RESTORE | Model 5 | REhabilitation STroke Outcome Recovery Estimator | 21 | IADL_T1T2_Change | +14.97 | +14.76 | 6.54 | [1.88, 27.87] | 99% |
| 1 | RESTORE | Model 5 | REhabilitation STroke Outcome Recovery Estimator | 21 | MNA1 | -11.52 | -10.94 | 6.00 | [-22.38, 0.00] | 96% |
| 1 | RESTORE | Model 5 | REhabilitation STroke Outcome Recovery Estimator | 21 | MNA_T1T2_Change | +8.33 | +8.10 | 5.07 | [-0.18, 18.34] | 95% |
| 1 | RESTORE | Model 5 | REhabilitation STroke Outcome Recovery Estimator | 21 | MRS1 | -16.46 | -16.46 | 7.94 | [-31.61, -0.92] | 98% |
| 1 | RESTORE | Model 5 | REhabilitation STroke Outcome Recovery Estimator | 21 | MRS_T1T2_Change | -25.02 | -24.47 | 8.79 | [-41.33, -8.29] | 99% |
| 1 | RESTORE | Model 5 | REhabilitation STroke Outcome Recovery Estimator | 21 | Sex, F0 M1 | +17.22 | +16.86 | 5.19 | [7.29, 27.10] | 100% |
| 2 | COMPASS | Model 8 | COMPrehensive Post-Acute Stroke Score | 44 | ACA | +6.30 | +6.88 | 5.31 | [0.00, 18.22] | 87% |
| 2 | COMPASS | Model 8 | COMPrehensive Post-Acute Stroke Score | 44 | Age | -33.34 | -33.85 | 6.37 | [-45.78, -21.18] | 100% |
| 2 | COMPASS | Model 8 | COMPrehensive Post-Acute Stroke Score | 44 | BBS1 | +63.93 | +62.54 | 10.00 | [42.40, 81.97] | 100% |
| 2 | COMPASS | Model 8 | COMPrehensive Post-Acute Stroke Score | 44 | CAD | +7.50 | +7.21 | 5.63 | [0.00, 18.35] | 87% |
| 2 | COMPASS | Model 8 | COMPrehensive Post-Acute Stroke Score | 44 | CKD | -6.34 | -6.79 | 4.21 | [-14.90, 0.00] | 93% |
| 2 | COMPASS | Model 8 | COMPrehensive Post-Acute Stroke Score | 44 | Cellulitis | -5.02 | -4.85 | 4.03 | [-13.65, 0.00] | 83% |
| 2 | COMPASS | Model 8 | COMPrehensive Post-Acute Stroke Score | 44 | Dementia | -6.22 | -6.01 | 3.84 | [-13.66, 0.00] | 93% |
| 2 | COMPASS | Model 8 | COMPrehensive Post-Acute Stroke Score | 44 | Dyslipidemia | +9.87 | +9.62 | 5.41 | [0.00, 20.17] | 96% |
| 2 | COMPASS | Model 8 | COMPrehensive Post-Acute Stroke Score | 44 | EuroQoL5D1 | +7.73 | +8.18 | 5.35 | [0.00, 18.52] | 92% |
| 2 | COMPASS | Model 8 | COMPrehensive Post-Acute Stroke Score | 44 | FuglUE1 | +21.23 | +21.26 | 6.67 | [7.80, 34.22] | 100% |
| 2 | COMPASS | Model 8 | COMPrehensive Post-Acute Stroke Score | 44 | GIB | +13.34 | +13.70 | 4.31 | [5.26, 22.12] | 100% |
| 2 | COMPASS | Model 8 | COMPrehensive Post-Acute Stroke Score | 44 | Gait_Speed_1_Imputed | +30.25 | +30.26 | 8.63 | [13.79, 47.14] | 100% |
| 2 | COMPASS | Model 8 | COMPrehensive Post-Acute Stroke Score | 44 | Gout | +2.92 | +3.83 | 4.77 | [-3.68, 14.75] | 76% |
| 2 | COMPASS | Model 8 | COMPrehensive Post-Acute Stroke Score | 44 | IADL1 | -3.99 | -5.48 | 5.09 | [-16.65, 0.00] | 78% |
| 2 | COMPASS | Model 8 | COMPrehensive Post-Acute Stroke Score | 44 | LVS | +4.40 | +5.55 | 4.95 | [-0.68, 15.83] | 84% |
| 2 | COMPASS | Model 8 | COMPrehensive Post-Acute Stroke Score | 44 | Loc_Subcortical | -13.89 | -11.13 | 5.77 | [-21.97, 0.00] | 96% |
| 2 | COMPASS | Model 8 | COMPrehensive Post-Acute Stroke Score | 44 | MNA1 | -14.17 | -14.68 | 5.77 | [-26.00, -3.16] | 99% |
| 2 | COMPASS | Model 8 | COMPrehensive Post-Acute Stroke Score | 44 | MRS1 | -8.55 | -9.56 | 6.27 | [-21.49, 0.00] | 91% |
| 2 | COMPASS | Model 8 | COMPrehensive Post-Acute Stroke Score | 44 | OldStroke | -14.42 | -14.49 | 5.38 | [-24.60, -3.41] | 100% |
| 2 | COMPASS | Model 8 | COMPrehensive Post-Acute Stroke Score | 44 | Psychiatric | -1.99 | -2.31 | 2.66 | [-8.77, 0.43] | 70% |
| 2 | COMPASS | Model 8 | COMPrehensive Post-Acute Stroke Score | 44 | Sex, F0 M1 | +15.95 | +15.90 | 5.16 | [6.10, 25.81] | 100% |
| 2 | COMPASS | Model 8 | COMPrehensive Post-Acute Stroke Score | 44 | Side_Right | +11.89 | +10.17 | 6.37 | [0.00, 22.32] | 89% |
| 2 | COMPASS | Model 8 | COMPrehensive Post-Acute Stroke Score | 44 | Undetermined | -4.99 | -5.60 | 4.25 | [-14.26, 0.00] | 87% |
| 3 | CASCADE | Model 2 | Clinical And Stroke Complexity Assessment for Discharge Endurance | 59 | ACA | +5.84 | +8.23 | 5.87 | [0.00, 21.71] | 91% |
| 3 | CASCADE | Model 2 | Clinical And Stroke Complexity Assessment for Discharge Endurance | 59 | Age | -31.32 | -33.53 | 6.22 | [-45.19, -20.87] | 100% |
| 3 | CASCADE | Model 2 | Clinical And Stroke Complexity Assessment for Discharge Endurance | 59 | BBS1 | +64.33 | +63.94 | 9.52 | [45.64, 81.84] | 100% |
| 3 | CASCADE | Model 2 | Clinical And Stroke Complexity Assessment for Discharge Endurance | 59 | CAD | +3.14 | +5.30 | 4.62 | [0.00, 15.14] | 80% |
| 3 | CASCADE | Model 2 | Clinical And Stroke Complexity Assessment for Discharge Endurance | 59 | CKD | -3.85 | -6.77 | 4.67 | [-16.54, 0.00] | 89% |
| 3 | CASCADE | Model 2 | Clinical And Stroke Complexity Assessment for Discharge Endurance | 59 | Cellulitis | -3.33 | -5.37 | 4.21 | [-14.37, 0.00] | 86% |
| 3 | CASCADE | Model 2 | Clinical And Stroke Complexity Assessment for Discharge Endurance | 59 | Dementia | -4.05 | -4.39 | 3.63 | [-12.39, 0.00] | 84% |
| 3 | CASCADE | Model 2 | Clinical And Stroke Complexity Assessment for Discharge Endurance | 59 | Dyslipidemia | +8.38 | +9.11 | 5.17 | [0.00, 19.70] | 96% |
| 3 | CASCADE | Model 2 | Clinical And Stroke Complexity Assessment for Discharge Endurance | 59 | EOMOut | +12.49 | +14.31 | 6.43 | [1.74, 26.68] | 99% |
| 3 | CASCADE | Model 2 | Clinical And Stroke Complexity Assessment for Discharge Endurance | 59 | EuroQoL5D1 | +4.72 | +8.49 | 5.58 | [0.00, 19.64] | 91% |
| 3 | CASCADE | Model 2 | Clinical And Stroke Complexity Assessment for Discharge Endurance | 59 | FOIS1 | +0.00 | +5.30 | 5.49 | [0.00, 18.28] | 73% |
| 3 | CASCADE | Model 2 | Clinical And Stroke Complexity Assessment for Discharge Endurance | 59 | FuglUE1 | +19.12 | +17.06 | 7.30 | [2.63, 31.61] | 99% |
| 3 | CASCADE | Model 2 | Clinical And Stroke Complexity Assessment for Discharge Endurance | 59 | GIB | +11.44 | +14.84 | 5.10 | [5.10, 24.19] | 100% |
| 3 | CASCADE | Model 2 | Clinical And Stroke Complexity Assessment for Discharge Endurance | 59 | Gait_Speed_1_Imputed | +29.71 | +28.54 | 8.50 | [12.01, 45.78] | 100% |
| 3 | CASCADE | Model 2 | Clinical And Stroke Complexity Assessment for Discharge Endurance | 59 | IADL1 | -0.00 | -4.48 | 4.91 | [-15.97, 0.25] | 71% |
| 3 | CASCADE | Model 2 | Clinical And Stroke Complexity Assessment for Discharge Endurance | 59 | LLOut | -4.81 | -8.56 | 7.15 | [-23.20, 0.00] | 80% |
| 3 | CASCADE | Model 2 | Clinical And Stroke Complexity Assessment for Discharge Endurance | 59 | LVS | +1.88 | +4.85 | 4.66 | [-0.22, 15.12] | 80% |
| 3 | CASCADE | Model 2 | Clinical And Stroke Complexity Assessment for Discharge Endurance | 59 | LanguageOut | +0.00 | +4.90 | 4.95 | [0.00, 16.35] | 74% |
| 3 | CASCADE | Model 2 | Clinical And Stroke Complexity Assessment for Discharge Endurance | 59 | Loc_Subcortical | -12.44 | -9.32 | 5.50 | [-20.20, 0.00] | 94% |
| 3 | CASCADE | Model 2 | Clinical And Stroke Complexity Assessment for Discharge Endurance | 59 | MNA1 | -10.89 | -14.96 | 5.65 | [-25.87, -3.05] | 99% |
| 3 | CASCADE | Model 2 | Clinical And Stroke Complexity Assessment for Discharge Endurance | 59 | MRS1 | -5.80 | -9.21 | 6.49 | [-22.46, 0.00] | 89% |
| 3 | CASCADE | Model 2 | Clinical And Stroke Complexity Assessment for Discharge Endurance | 59 | NeglectOut | -0.00 | -1.03 | 6.75 | [-14.16, 13.09] | 79% |
| 3 | CASCADE | Model 2 | Clinical And Stroke Complexity Assessment for Discharge Endurance | 59 | OldStroke | -12.84 | -14.82 | 5.27 | [-25.78, -4.86] | 100% |
| 3 | CASCADE | Model 2 | Clinical And Stroke Complexity Assessment for Discharge Endurance | 59 | RLOut | -1.43 | -6.84 | 7.27 | [-23.95, 0.00] | 72% |
| 3 | CASCADE | Model 2 | Clinical And Stroke Complexity Assessment for Discharge Endurance | 59 | Sex, F0 M1 | +15.69 | +15.85 | 5.14 | [6.07, 25.75] | 100% |
| 3 | CASCADE | Model 2 | Clinical And Stroke Complexity Assessment for Discharge Endurance | 59 | Side_Right | +9.76 | +11.21 | 7.93 | [0.00, 26.91] | 87% |
| 3 | CASCADE | Model 2 | Clinical And Stroke Complexity Assessment for Discharge Endurance | 59 | Undetermined | -3.33 | -5.68 | 4.14 | [-14.66, 0.00] | 88% |
| 4 | AIMS | Model 1 | Admission Impairment Mobility Score | 12 | Age | -39.17 | -39.64 | 6.32 | [-52.01, -27.24] | 100% |
| 4 | AIMS | Model 1 | Admission Impairment Mobility Score | 12 | BBS1 | +62.42 | +61.23 | 9.53 | [42.55, 79.89] | 100% |
| 4 | AIMS | Model 1 | Admission Impairment Mobility Score | 12 | EuroQoL5D1 | +10.47 | +10.67 | 5.68 | [0.00, 21.48] | 95% |
| 4 | AIMS | Model 1 | Admission Impairment Mobility Score | 12 | FuglUE1 | +23.96 | +23.74 | 7.18 | [10.07, 37.45] | 100% |
| 4 | AIMS | Model 1 | Admission Impairment Mobility Score | 12 | Gait_Speed_1_Imputed | +33.06 | +33.06 | 8.46 | [16.66, 49.35] | 100% |
| 4 | AIMS | Model 1 | Admission Impairment Mobility Score | 12 | IADL1 | -3.46 | -5.10 | 5.20 | [-16.64, 1.53] | 77% |
| 4 | AIMS | Model 1 | Admission Impairment Mobility Score | 12 | MNA1 | -11.45 | -11.89 | 5.42 | [-22.08, 0.00] | 97% |
| 4 | AIMS | Model 1 | Admission Impairment Mobility Score | 12 | MRS1 | -6.19 | -6.42 | 6.08 | [-19.72, 0.54] | 82% |
| 4 | AIMS | Model 1 | Admission Impairment Mobility Score | 12 | Sex, F0 M1 | +15.69 | +15.51 | 5.24 | [5.19, 25.25] | 100% |
| 5 | BEDSIDE | Model 4 | Balance-Enhanced Demographic Speed Index for Discharge Endurance | 4 | Age | -35.40 | -35.10 | 5.92 | [-46.15, -23.68] | 100% |
| 5 | BEDSIDE | Model 4 | Balance-Enhanced Demographic Speed Index for Discharge Endurance | 4 | BBS1 | +69.04 | +68.60 | 8.64 | [52.39, 86.17] | 100% |
| 5 | BEDSIDE | Model 4 | Balance-Enhanced Demographic Speed Index for Discharge Endurance | 4 | Gait_Speed_1_Imputed | +37.95 | +38.29 | 8.29 | [20.56, 53.03] | 100% |
| 5 | BEDSIDE | Model 4 | Balance-Enhanced Demographic Speed Index for Discharge Endurance | 4 | Sex, F0 M1 | +17.58 | +17.59 | 5.35 | [7.43, 27.98] | 100% |

*β values are standardized bootstrap LASSO coefficients. In Panel B, cells are shown as Boot Mean Coef (Boot SD) [95% CI]; selection frequency.*
*Acc = accuracy; BalAcc = balanced accuracy; Se = sensitivity; Sp = specificity; CI = 95% bootstrap confidence interval.*