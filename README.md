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
