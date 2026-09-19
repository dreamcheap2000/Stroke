# 6MWT denominator comparison (20260919 update)

## Dataset lineage used
- `20260912_Tables_1840.py` points to `20260904_Comprehensive_1018.xlsx` (`Predictions` sheet) as the patient-level source for 6MWT4 model denominators.
- This update uses that same patient-level lineage dataset.

## Count comparison
| Metric | Count |
| --- | --- |
| Current 20260919_6MWT.csv PAC completers with First_6MWT_TP != Never | 517 |
| Model-eligible PAC completers from 20260904_Comprehensive_1018.xlsx (6MWT4 non-missing) | 511 |
| Updated 20260919_Update_6MWT_0955.csv PAC completers with First_6MWT_TP != Never | 511 |
| Difference (current table - model-eligible) | 6 |

## Why 517 vs 511
The prior visible table count of 517 includes all PAC completers whose `First_6MWT_TP` is not `Never`, regardless of whether the model outcome (`6MWT4`) is present.
The modeling denominator requires non-missing `6MWT4`. Exactly 6 PAC completers had an initial/first 6MWT time point but missing `6MWT4`, so they are excluded from model-eligible counts (511).

## Discrepant records (n=6)
| ID | Rehab_LOS_Category | PAC_Program_Completion | First_6MWT_TP | 6MWT4 | Initial_6MWT_Distance | 6MWT_P1_Change | 6MWT_P2_Change | 6MWT_P3_Change | Missing_or_excluded_field |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| 208 | 21-42 days | Completed PAC program | T1 |  | 349.8 |  |  |  | 6MWT4 missing (model outcome required for eligibility) |
| 257 | 21-42 days | Completed PAC program | T1 |  | 294.0 | 58.0 |  |  | 6MWT4 missing (model outcome required for eligibility) |
| 112 | >42 days | Completed PAC program | T1 |  | 82.0 | 18.0 | 5.0 |  | 6MWT4 missing (model outcome required for eligibility) |
| 107 | >42 days | Completed PAC program | T2 |  | 174.0 | 8.0 |  |  | 6MWT4 missing (model outcome required for eligibility) |
| 458 | >42 days | Completed PAC program | T3 |  | 230.0 |  |  |  | 6MWT4 missing (model outcome required for eligibility) |
| 550 | >42 days | Completed PAC program | T3 |  | 88.0 |  |  |  | 6MWT4 missing (model outcome required for eligibility) |
