# Stability Report Generator

An automated data pipeline for immunoassay stability studies in regulated laboratory environments. Replaces a manual process that took 10–12 hours per timepoint (processing + authentication) with a one-click application that reads raw instrument output files and generates complete formatted reports in minutes.

Built as a portfolio project at Fujirebio Diagnostics alongside an M.S. in Data Science.

---

## The Problem

Stability studies for diagnostic assays run over 27 months across 8 timepoints. At each timepoint, 18 raw instrument files need to be parsed, consolidated across 2 instruments × 3 runs × 3 lots, evaluated against acceptance criteria, formatted into reports, and authenticated. The manual process took **5–6 hours of processing + 5–6 hours of data authentication per timepoint**, across 8 timepoints and 2 concurrent studies — roughly **160+ hours per study cycle**.

---

## The Solution

A fully automated pipeline that:

1. Reads raw Raven instrument output files directly from a structured folder
2. Parses filenames to extract instrument, run number, lot, and timepoint
3. Loads acceptance criteria from a per-study `config.xlsx` — no code changes needed for new assays
4. Evaluates every run against specs with full rerun and supersession logic
5. Generates formatted Excel reports with pass/fail colour coding and individual replicate grids
6. Produces a trends report with embedded concentration-over-time charts
7. Exports Tableau-ready CSVs for optional dashboard analysis
8. Packages everything into a standalone Windows `.exe` — no Python required for end users

**Validated against authenticated manual summaries — 100% match across 126 run-sample records and 1,512 individual replicate values.**

---

## Key Features

**Config-driven architecture** — each study folder contains a `config.xlsx` defining the assay name, sample names, spec limits, CV limit, and minimum rep count. The pipeline adapts automatically to any assay without code changes.

**Rerun handling** — original runs preserved as superseded, reruns detected by filename suffix (`_Rerun1`, `_Rerun2`), best-2-of-3 logic for three-version scenarios.

**Robust edge cases** — partial rep counts, below-curve results excluded from calculations, unrecognized sample names flagged with exact file and mismatch details, corrupt files skipped gracefully.

**Flexible outputs** — per-timepoint Excel report, combined trends report (regenerated on every run), and a `Tableau/` subfolder with three CSVs: RunSummary, GrandMeans, Replicates.

---

## Results

| Metric | Value |
|---|---|
| Processing time reduction | ~95% |
| Hours saved per study cycle | 160+ |
| Raw files processed per study | 144 (18 per timepoint × 8 timepoints) |
| Individual replicates evaluated per timepoint | 1,512 |
| Validation accuracy vs authenticated summaries | 100% |
| Assays supported | Any — config-driven |

---

## Technical Stack

**Python** · **pandas** · **openpyxl** · **matplotlib** · **tkinter** · **PyInstaller** · **regex**

---

## Project Structure

```
stability-report-generator/
│
├── stability_report_v2.py      # Core pipeline
├── stability_report_app.py     # GUI wrapper (tkinter)
├── config_template.xlsx        # Config template — copy to each study folder
├── README.md
│
├── example_data/
│   └── TestAssay_Stab11/
│       ├── config.xlsx
│       ├── TP1/
│       │   ├── A4_Run1_Lot1_TP1.xlsx
│       │   ├── A4_Run1_Lot2_TP1.xlsx
│       │   ├── A4_Run1_Validity_TP1.xlsx
│       │   └── ... (18 files per timepoint)
│       └── TP2/ ... TP8/
│
└── example_output/
    ├── TestAssay_Stab11_TP1_Report.xlsx
    └── TestAssay_Stab11_Trends_Report.xlsx
```

---

## How to Run

### Python (recommended for development)

```bash
pip install pandas openpyxl matplotlib
python stability_report_v2.py path/to/study/folder
```

### GUI application

```bash
python stability_report_app.py
```

Point the GUI at a study folder containing TP subfolders and click **Generate All Reports**.

### Build the Windows .exe

```bash
pip install pyinstaller
pyinstaller --onedir --windowed --name "StabilityReportGenerator" \
    --add-data "stability_report_v2.py;." \
    --exclude-module PySide6 --exclude-module PyQt5 \
    --exclude-module scipy --exclude-module IPython \
    stability_report_app.py
```

---

## Folder and File Naming Convention

```
{AssayName}_{StudyNumber}/        e.g.  HE4_Stab11/
    config.xlsx
    TP1/
        {Instrument}_{Run}_{Lot}_{Timepoint}.xlsx
        e.g.  A4_Run1_Lot1_TP1.xlsx
              A4_Run1_Lot2_TP1.xlsx
              A4_Run1_Validity_TP1.xlsx
    TP2/ ... TP8/
```

Reruns use a suffix: `A4_Run1_Lot1_TP1_Rerun1.xlsx`, `_Rerun2.xlsx`

---

## Config File

The `config.xlsx` in each study folder has two sheets:

- **Study Info** — assay name, study number, CV limit, minimum valid reps
- **Specs** — one row per sample per lot with individual rep limits, mean limits, and validity limits. Supports mid-study lot changes via an `Effective From` timepoint column.

See `config_template.xlsx` for a pre-filled example.

---

## Report Structure

### Per-timepoint report
- **Summary sheet** — run-level results for all lots, between-run RefCon validity, grand mean acceptance
- **Per-sample sheets** — individual replicate grids (green/red by spec), stats rows (SD, Mean, %CV, Valid?, Accepted), run summary table

### Trends report
- **Overview sheet** — pass/fail summary across all timepoints
- **Per-sample sheets** — grand mean and %CV table across timepoints, one chart per feasibility lot showing concentration trend with spec limit lines and green/red pass/fail data points

---

## Background

This project began as a direct replacement for a fully manual stability study workflow in a regulated diagnostics environment. All outputs were cross-validated against authenticated manual summaries before being put into use. The pipeline handles multiple concurrent studies across different assays with no modifications to the codebase.

The regulated environment context (FDA stability studies, data authentication requirements) shaped several design decisions — particularly around error transparency, the rerun supersession model, and the validation approach.

---

## Author

Evan Lutz
Process Design Scientist → Data Science (M.S. Eastern University, expected Dec 2026)
[LinkedIn](https://www.linkedin.com/in/evan-lutz-2589b6167/) · [GitHub](https://github.com/elutz98)
