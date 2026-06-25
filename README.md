# EMEA DG Curators + Designers Dashboard

Self-contained HTML dashboard for EMEA Demand Generation — Curators and Designers team performance.

## Files

| File | Description |
|------|-------------|
| `EMEA Curators+Designers Dashboard v1.html` | The dashboard — open directly in any browser, no server needed |
| `build_dashboard.py` | Python script that reads the source Excel files and regenerates the embedded JSON |

## How to update the data

1. Drop updated Excel files in `01 - Source Data/` (same folder structure)
2. Run the build script:
   ```bash
   pip install pandas openpyxl
   python build_dashboard.py
   ```
3. The script outputs `/tmp/dash_v4.json`
4. Run the aggregation step (see script comments) to produce the final HTML

## Data sources required

- `01 - Source Data/DG Plans/` — MU and LOB DG Action Plan Excel files
- `01 - Source Data/Outreach/Outreach Reports CUR+DES.xlsx`
- `01 - Source Data/Pipeline/Global Cloud Pipeline Data Analyzer CUR+DES.xlsx`
- `01 - Source Data/Curators/CURATORS X MSLIST REQUEST CUR+DES.xlsx`
- `01 - Source Data/Designers/DESIGNERS CREATIVE TEAM REQUEST REVIEW.xlsx`

> **Note:** Source data files are not tracked in this repo (size + confidentiality).
