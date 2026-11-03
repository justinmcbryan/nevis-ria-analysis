# Nevis take-home: US RIA market analysis (SEC Form ADV, 2021–2026)

**Notebook:** `nevis_ria_analysis.ipynb` (opens in Google Colab or Jupyter; Runtime → Run all).
**Outputs:** `outputs/nevis_ria_targets.xlsx` (ranked lists by segment and the Nevis target list) and `outputs/ria_firm_year_table.csv` (analysis-ready firm-year table).

## Data
SEC monthly "Information About Registered Investment Advisers" roster (Form ADV Part 1A):
- **September 2021 to September 2026** snapshots (= **FY2020–FY2025** annual-filing data) for metrics (AUM, clients, employees, advisors), giving 1-, 3- and 5-year windows exactly 12 months apart, after annual-amendment season.
- **November 2026** for firm identity (name, address, website), because Item 1 must be amended promptly.
- **SEC bulk Form ADV archive** (all filings Nov 2011 – Dec 2024, ~1.1 GB) for Schedule D detail: custodians (5.K(3)), owners and control persons (Schedules A/B), successions (Section 4) and full website lists (1.I). The notebook downloads it on first run and caches three small extracts in `data/reference/`.

The notebook downloads the files from sec.gov on first run (set `SEC_USER_AGENT` if SEC asks for identification) and caches them in `data/`.

## Structure
1. **Executive summary**: the questions being answered and why they matter
2. **Q1**: RIAs by AUM segment (FY2025), and the target RIAs Nevis would sell to
3. **Q2a**: fastest-growing firms (1/3/5 years)
4. **Q2b**: capacity-strained firms
5. **Q2c**: under-staffed firms
6. **Extra: consolidated findings**: prioritised Nevis target list for firms above \$10B
7. **Extra: additional supporting data**: custodians, potential PE consolidation
The 22 automated data checks run inside the pipeline cell (critical checks stop the notebook before any result is shown); what each check proves and its result are documented separately.

All code is collapsed in Colab so the notebook reads as a report. A single collapsed cell just before Q1 runs the whole pipeline (data loading, analysis engine and all 22 data checks) before any result is shown; click **Show code** on any cell to see it.

`nevis_ria_analysis.py` is the same notebook in plain-text (jupytext) form for easy diffing.

## Run it yourself
- **Colab:** open the notebook from this repo in Google Colab and choose Runtime → Run all.
- **Locally:** `pip install -r requirements.txt` then `jupyter notebook nevis_ria_analysis.ipynb`.

Raw SEC files (~75 MB of roster snapshots + ~1.1 GB bulk archive) are not committed; the notebook downloads them from sec.gov on first run (about 5 minutes in Colab) and caches them.

## Approach in one paragraph
Six September snapshots (2021–2026) give clean 12-month growth intervals after annual-amendment season. Firms are linked across years by CRD#. Growth is CAGR with segment-scaled minimum starting AUM; acquisitions, re-registrations and advisor-count reclassifications are flagged and ranked separately. The presentation focuses on **target RIAs** (the RIAs Nevis sells to: US-based, mainly serving individuals, offering financial planning): asset managers, fund advisers and robo platforms are excluded, while wirehouses and broker-dealers are kept and tagged. All-RIA rankings are in the Excel file.

## Note on GitHub's language bar
GitHub reports ~94% "Jupyter Notebook" because it weighs files by size: the `.ipynb` stores the saved outputs (charts as images, result tables). `nevis_ria_analysis.py` is the identical code as plain text for easy reading and diffing.
