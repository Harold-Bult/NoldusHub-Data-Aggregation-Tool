# NoldusHub Data Aggregation Tool

Streamlit app that aggregates and resamples NoldusHub physiological/behavioral
CSV exports (EDA, PPG, eye tracking, facial expressions, events, ...) into a
single table, with per-signal completeness overview and `.xlsx` export.

## Requirements

- Python 3.x
- `streamlit`, `pandas`, `numpy`, `scipy`, `openpyxl`

```bash
pip install streamlit pandas numpy scipy openpyxl
```

## Files to add before running (not in this repo)

These come with your existing NoldusHub tool installation and must be placed
next to `app.py`:

- `NH_datsegmentation.cfg` — signal/AOI/header configuration (required)
- `static/noldus.ico`
- `static/logo_noldus.png`

Session data (the folders with the CSVs) are **not** stored here: they are
selected at runtime via the "Select data root folder" button.

## How to run

```bash
streamlit run app.py
```

Then: "Select data root folder" -> pick the folder containing the session
subfolders (each with its CSVs) -> the table shows the completeness % per
signal -> click a row to preview a resampled session, or "Create aggregate
xlsx" to export.

## Fixes in this version

1. Robust timestamp parsing (`_parse_iso_flexible`): handles the ISO 8601 `T`
   separator and 6/7-digit second fractions (.NET format), in both the header
   and the data rows.
2. Dynamic `PROGPATH` (no hard-coded path); works on Windows and Linux.
3. Session detection via the full `os.walk` path (a folder is a session if it
   contains signal CSVs).
4. Robustness: missing durations, all-NaN columns, read-only numpy array in
   `NH_resample` (`to_numpy(copy=True)`), timezone alignment of timestamps.

> Note: the `.xlsx` export is inherently slow (openpyxl). This is expected.
