# DP Automation v1.2 — SPSS + Excel Zero-Mapping

A standalone Streamlit prototype for automating market-research DP preparation.

## Inputs

### SPSS `.sav`
Best input because the app can read variable labels and value labels directly.

### Excel `.xlsx`
The app reads respondent data from `Raw_Data`, `Data`, `Raw Data`, or otherwise the first sheet.

For full metadata support, add a sheet named `SPSS_Metadata_Reference` or `Metadata` with these columns:
- `Variable`
- `Variable Label`
- `Value Labels` using syntax such as `1=Male; 2=Female`

The metadata sheet is optional. If absent, the app still works from column names and observed data patterns, but value-label/semantic inference is necessarily weaker.

## Zero-mapping workflow
1. Upload `.sav` or `.xlsx`.
2. Automatic variable mapping and question-type detection.
3. Review only exceptions / low-confidence rows in Streamlit.
4. Optional recode by editing `new_value` directly in the screen.
5. Generate final `.sav`, `.xlsx`, `.csv`, metadata, audit log, mapping plan and QC report.

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

On Windows you can also run `start_app.bat`.

## Synthetic test SAV
This package includes `create_sample_sav.py`. After dependencies are installed, run:
```bash
python create_sample_sav.py
```
It creates `Sample_Sawtooth_DP_Input.sav` with single-code, rating, multi-response, grid, open-end and DK examples.
