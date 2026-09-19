# Stroke: retained files and generation map

## Retained dataset
- `20260806_DeID.xlsx`
  - Source dataset used for all retained analysis outputs (directly or through derived artifacts).

## Requested retained outputs
- `tipping_point_curve_202608081955.png`
- `Tip_Over_Analysis_results_202608081955.csv`
- `Tip_Over_Analysis_artifacts_out_202608081955.csv`
- `IPCW_summary_out_202608081955.csv`
- `IPCW_smd_out_202608081955.csv`
- `IPCW_Table_202608090845.docx`

## Code and provenance (which dataset produced which file)

### A) IPCW summary outputs
- Script: `IPCW_Out_202608081955.py`
- Input dataset: `20260806_DeID.xlsx`
- Produces:
  - `IPCW_summary_out_202608081955.csv`
  - `IPCW_smd_out_202608081955.csv`

Run:
```bash
python IPCW_Out_202608081955.py 20260806_DeID.xlsx
```

### B) Tip-over results and figure (timestamped set)
- Script: `Tip_Over_Analysis_202608081955.py`
- Input dataset: `20260806_DeID.xlsx`
- Produces:
  - `Tip_Over_Analysis_results_202608081955.csv`
  - `tipping_point_curve_202608081955.png`

Run:
```bash
python Tip_Over_Analysis_202608081955.py 20260806_DeID.xlsx
```

### C) Tip-over IPCW artifact file used downstream
- Script: `Tip_Over_Analysis_out_202608081955.py`
- Input dataset: `20260806_DeID.xlsx`
- Produces:
  - `Tip_Over_Analysis_artifacts_out_202608081955.csv`

Run:
```bash
python Tip_Over_Analysis_out_202608081955.py 20260806_DeID.xlsx
```

### D) IPCW table document
- Script: `IPCW_Table_202608090845.py`
- Input data: `Tip_Over_Analysis_artifacts_out_202608081955.csv`
  - (This artifact is derived from `20260806_DeID.xlsx` via script C.)
- Produces:
  - `IPCW_Table_202608090845.docx`

Run:
```bash
python IPCW_Table_202608090845.py
```

### E) Binary 6MWT scenario dataset and model reports
- Script: `Binary_1038_20260826.py`
- Input dataset: `20260806_DeID.xlsx`
- Produces:
  - `20260826_DeID.xlsx`
  - `20260826_Binary_1038.docx`
  - `20260826_Binary_1038_Code.docx`

Run:
```bash
python Binary_1038_20260826.py
```

### F) Tiered IPCW 6MWT4 extrapolation outputs
- Script: `20260828_5_Tiers_1201.py`
- Input dataset: `20260826_DeID.xlsx`
  - (This dataset is derived from `20260806_DeID.xlsx` via script E.)
- Produces:
  - `20260828_DeID_IPCW_Predicted_6MWT.xlsx`
  - `20260828_5_Tiers_1201.docx`

Run:
```bash
python 20260828_5_Tiers_1201.py
```

### G) 20260919 IPCW/SMD refresh and downstream synchronized outputs
- Newest patient-level input used for refresh: `20260826_DeID.xlsx`
  - Chosen because it is the latest DeID patient-level cohort file used by active IPCW/comprehensive modeling scripts.
  - Newer workbooks (e.g., `20260903_*`, `20260904_*`, `20260909_*`) are model-output workbooks with derived prediction/summary sheets, not raw patient-level source inputs.
- IPCW rerun script: `IPCW_Out_20260919.py`
- Input dataset: `20260826_DeID.xlsx`
- Produces:
  - `IPCW_summary_out_20260919.csv`
  - `IPCW_smd_out_20260919.csv`

Run:
```bash
python IPCW_Out_20260919.py 20260826_DeID.xlsx
```

- Downstream synchronized outputs regenerated from the same input chain and saved with `20260919` naming:
  - `20260919_DeID_IPCW_Predicted_6MWT.xlsx`
  - `20260919_5_Tiers_1201.docx`
  - `20260919_LASSO_1153.xlsx`
  - `20260919_LASSO_1153.docx`
  - `20260919_Comprehensive_1327.xlsx`
  - `20260919_Comprehensive_1327.docx`
  - `20260919_Comprehensive_1018.xlsx`
  - `20260919_Comprehensive_1018.docx`
