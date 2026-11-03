# Nevis take-home: US RIA market analysis (SEC Form ADV, 2021–2026)

**Notebook:** `nevis_ria_analysis.ipynb` (opens in Google Colab or Jupyter; Runtime → Run all).
**Outputs:** `outputs/nevis_ria_targets.xlsx` (ranked lists by segment and the Nevis target list) and `outputs/ria_firm_year_table.csv` (analysis-ready firm-year table).

## Data
SEC monthly "Information About Registered Investment Advisers" roster (Form ADV Part 1A):
- **September 2021 to September 2026** for metrics (AUM, clients, employees, advisors), giving 1-, 3- and 5-year windows exactly 12 months apart, after annual-amendment season.
- **November 2026** for firm identity (name, address, website), because Item 1 must be amended promptly.

The notebook downloads the files from sec.gov on first run (set `SEC_USER_AGENT` if SEC asks for identification) and caches them in `data/`.

## Structure
0 Setup · 1 Sourcing rationale · 2 Field mapping + website data-quality fix · 3 QC · 4 Firm table · 5 Segments · 6 Method · 7 Fastest growth · 8 Capacity · 9 Under-staffing · 10 Nevis target list · 11 Assumptions · 12 How AI was used

`nevis_ria_analysis.py` is the same notebook in plain-text (jupytext) form for easy diffing.

## Run it yourself
- **Colab:** open the notebook from this repo in Google Colab and choose Runtime → Run all.
- **Locally:** `pip install -r requirements.txt` then `jupyter notebook nevis_ria_analysis.ipynb`.

Raw SEC files (~75 MB) are not committed; the notebook downloads them from sec.gov on first run.

## Approach in one paragraph
Six September snapshots (2021–2026) give clean 12-month growth intervals after annual-amendment season. Firms are linked across years by CRD#. Growth is CAGR with segment-scaled minimum starting AUM; acquisitions, re-registrations and advisor-count reclassifications are flagged and ranked separately. Rankings are shown for all RIAs (the literal answer) and for **advisor-led wealth managers** (Nevis's ICP), then combined into a prioritised target list per segment.

## Note on GitHub's language bar
GitHub reports ~94% "Jupyter Notebook" because it weighs files by size: the `.ipynb` stores the saved outputs (charts as images, result tables). `nevis_ria_analysis.py` is the identical code as plain text for easy reading and diffing.
