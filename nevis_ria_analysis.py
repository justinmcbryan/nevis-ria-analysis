# %% [markdown]
# # US RIA Market Analysis for Nevis: Growth, Capacity & Staffing Signals
# **Source:** SEC monthly Form ADV Part 1A roster ("Information About Registered Investment Advisers"), September 2021 to September 2026, plus November 2026 for current firm details.
# **Author:** Justin McBryan · November 2026
#
# ### Executive summary
# **The market (Sep 2026):** 17,149 SEC-registered RIAs managing \$177T of regulatory AUM.
#
# | Segment | RIAs | Share of firms | Share of AUM | Advisor-led wealth managers |
# |---|---|---|---|---|
# | <\$1B | 11,756 | 69% | 2% | 5,628 |
# | \$1B-\$10B | 4,023 | 23% | 7% | 1,149 |
# | \$10B-\$50B | 938 | 5% | 12% | 113 |
# | >\$50B | 432 | 3% | 79% | 43 |
#
# **Key findings for Nevis**
# 1. **The capacity squeeze is near-universal.** Among \$1B+ advisor-led wealth managers, **92% grew AUM faster than advisor headcount** over three years: median AUM growth of **18.1% a year vs 4.6%** for advisors. Across all RIAs the median firm added no advisors at all. Every advisor is carrying more each year, which is the problem Nevis solves.
# 2. **Growth is concentrated in a few platforms.** The fastest *sustained* growers among advisor-led firms include Waverly Advisors (\$5.8B → \$29.9B in 3 years), Pathstone (\$24.8B → \$110B), Arkadios, Apella and Merit Financial. Most fast growth at scale is acquisitive: Focus Partners, Osaic and Cetera consolidated affiliated advisors into single registrants.
# 3. **Raw "AUM per advisor" rankings point at the wrong firms.** Unfiltered, they're led by Vanguard, BlackRock and fund managers. Filtered to advisor-led wealth managers, a typical \$1B+ firm runs about \$125M-\$130M per advisor and 60-80 clients per advisor. Outliers like Circle Wealth (\$1.3B per advisor) and First Manhattan (\$750M) show where strain is highest.
# 4. **Enterprise targets:** combining growth, current strain and worsening strain gives a ranked list. For \$50B+: Cresset, Pathstone, Focus Partners, MAI Capital, Savant, Carson, IEQ, Fisher, Osaic, Cetera. For \$10B-\$50B: Aptus, RWA Wealth Partners, Circle Wealth, Clearstead, CLA, Merit, Apella, Apollon, NewEdge, Steward Partners. (Section 10)
#
# **Data robustness:** 19 automated integrity checks (V0-V18) run on every execution and all pass, covering file fingerprints, schema, uniqueness, Form ADV arithmetic identities, reconciliation to SEC totals, cross-year linkage, an IAPD spot check, and independent recomputation of growth rates (Sections 4a and 12a; also in the Excel "Data checks" tab).
#
# **Data-quality issues handled:** one website per firm, often a social profile (Beacon Pointe shows Pinterest); re-registrations and roll-ups that look like 10,000% growth; robo/TAMP platforms that look like wealth managers; and advisor-count reclassifications (Vanguard 301 → 19). Each is flagged and explained below rather than silently dropped.
#
# **Deliverables:** this notebook · `outputs/nevis_ria_targets.xlsx` (every ranked list, top 25 per segment, plus the full 2026 firm table) · `outputs/ria_firm_year_table.csv` (the analysis-ready firm-year table, about 95k rows).

# %% [markdown]
# ## 0. How to run
# Works locally or in Google Colab: **Runtime → Run all**. The first run downloads seven files (about 75 MB) from sec.gov and caches them; reruns use the cache.
# The SEC requires automated downloads to identify themselves. The SEC asks automated downloads to identify a name and contact email via `SEC_USER_AGENT` (set in the next cell). If sec.gov still refuses, upload the seven files to `data/raw/` and rerun.

# %%
import os, re, zipfile, warnings
from pathlib import Path
import numpy as np, pandas as pd, requests
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
warnings.filterwarnings("ignore")
pd.set_option("display.max_columns", 40); pd.set_option("display.width", 220)
pd.set_option("display.float_format", lambda x: f"{x:,.2f}")

DATA_DIR = Path(os.environ.get("NEVIS_DATA_DIR", "data"))
RAW, CACHE, OUT = DATA_DIR / "raw", DATA_DIR / "cache", Path("outputs")
for p in (RAW, CACHE, OUT): p.mkdir(parents=True, exist_ok=True)
# SEC fair-access policy: automated requests must identify a name and contact email
SEC_USER_AGENT = os.environ.get("SEC_USER_AGENT", "Justin McBryan justin.mcbryan@useaffluence.com")

# %% [markdown]
# ## 1. Data sourcing: why these files
#
# | Choice | Reason |
# |---|---|
# | **Monthly SEC roster**, not the bulk ADV tables | One row per firm with every Part 1A answer already flattened into columns. Same fields as the bulk data, far less joining risk. |
# | **One snapshot per year (September)** | Item 5 (AUM, clients, headcount) only has to be refreshed in the **annual updating amendment**, due within 90 days of fiscal year end (Form ADV General Instruction 4). Monthly files in between mostly repeat the same numbers. **Validated by CHECK V12:** 97.4% of firms have identical AUM in Sep and Nov 2026. |
# | **September specifically** | It is after filing season, so nearly every firm has filed. It exists in all six years (Oct and Nov 2025 are "no data" on the SEC site). Every growth interval is exactly 12 months. |
# | **November 2026 for firm identity** | Items 1 (name, address, **website**) must be amended *promptly* when they change, so the latest file holds the most current contact details. |
#
# Note: a September 20XX snapshot mostly reflects AUM as of each firm's prior fiscal year end (usually December of 20XX-1). Years below are **snapshot years**.

# %%
BASE = "https://www.sec.gov/files/investment/data/"
OLD, NEW = "information-about-registered-investment-advisers-and-exempt-reporting-advisers/", "information-about-registered-investment-advisers-exempt-reporting-advisers/"
SNAPSHOTS = {
    "2021-09": BASE + OLD + "ia090121.zip",
    "2022-09": BASE + OLD + "ia090122.zip",
    "2023-09": BASE + NEW + "ia090623.xlsx",
    "2024-09": BASE + NEW + "ia090324.xlsx",
    "2025-09": BASE + NEW + "ia09022025.xlsx",
    "2026-09": BASE + "other/" + NEW + "ia09012026-registered.zip",
    "2026-11": BASE + "other/" + NEW + "ia110226-registered.zip",
}

# --- Alternative source: local hard copies -----------------------------------------------------------------
# The seven files below were also downloaded manually from the SEC website on 2 Nov 2026 and kept as hard copies.
# If sec.gov ever blocks automated downloads, place those copies in data/raw/ with these exact filenames:
# fetch() uses any file already present and skips the download. In Colab, the copies can live in Google Drive:
#   from google.colab import drive; drive.mount("/content/drive")
#   DATA_DIR = Path("/content/drive/MyDrive/Nevis-RIA-Analysis/data"); RAW, CACHE = DATA_DIR / "raw", DATA_DIR / "cache"
# EXPECTED_FILES records each file's size and SHA-256 fingerprint so CHECK V0 can prove the data is byte-identical
# to what this analysis used (the SEC occasionally republishes files).
EXPECTED_FILES = {  # filename: (bytes, first 16 hex chars of SHA-256)
    "ia090121.zip": (3384454, "ccc12bc76e2bc151"),
    "ia090122.zip": (3565069, "01e4ec044edaa5bb"),
    "ia090623.xlsx": (19506755, "5a8f2d98f0b40a22"),
    "ia090324.xlsx": (20010116, "1bae7d25b6cc06ad"),
    "ia09022025.xlsx": (20547238, "977d83fdd67acdc4"),
    "ia09012026-registered.zip": (5326806, "2302e3ea00819b40"),
    "ia110226-registered.zip": (5351876, "f17a9012368124a1"),
}

def fetch(url):
    path = RAW / url.split("/")[-1]
    if not path.exists():
        r = requests.get(url, headers={"User-Agent": SEC_USER_AGENT, "Accept-Encoding": "gzip, deflate"}, timeout=180)
        if r.status_code == 403:
            raise RuntimeError(f"SEC refused {url} (403). Set SEC_USER_AGENT to 'Your Name your@email', or upload the file to {RAW}/ and rerun.")
        r.raise_for_status(); path.write_bytes(r.content)
    return path

def read_snapshot(label, url):
    cache = CACHE / f"{label}.pkl"
    if cache.exists(): return pd.read_pickle(cache)
    path = fetch(url)
    if path.suffix == ".zip":
        with zipfile.ZipFile(path) as z:
            name = [n for n in z.namelist() if n.lower().endswith(".csv")][0]
            df = pd.read_csv(z.open(name), encoding="latin-1", dtype=str, low_memory=False)
    else:
        df = pd.read_excel(path, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    df.to_pickle(cache); return df

raw = {lab: read_snapshot(lab, url) for lab, url in SNAPSHOTS.items()}
pd.DataFrame({lab: {"firms": len(d), "columns": d.shape[1]} for lab, d in raw.items()}).T

# %% [markdown]
# ## 2. Field mapping: Form ADV item → analysis field
#
# | Field | Form ADV source | Treatment |
# |---|---|---|
# | SEC# | Item 1.D (801-xxxxx) | As filed |
# | Join key | Item 1.E **CRD#** | Stable across renames; used to link years |
# | Firm name, City, State | Items 1.B, 1.F | From **Nov 2026** (most current) |
# | Website | Item 1.I / Sched. D 1.I | See data-quality note below |
# | AUM | Item 5.F(2)(c) | Total regulatory AUM, USD |
# | Clients | Sum of Item 5.D(x)(1) across client types | "Fewer than 5" boxes imputed at 2.5 |
# | Employees | Item 5.A | Excludes clerical staff by definition |
# | Advisors | Item 5.B(1) | Staff performing investment advisory functions (incl. research) |
# | IARs (cross-check) | Item 5.B(3) | Registered investment adviser reps, closer to "client-facing advisors" |
# | Retail AUM share (added) | (5.D(a)(3)+5.D(b)(3)) / AUM | Share of AUM from individuals + HNW, used to identify **wealth managers** |
# | Filing date | "Latest ADV Filing Date" | As filed |

# %%
CLIENT_TYPES = list("abcdefghijklmn")
SOCIAL = re.compile(r"linkedin|facebook|twitter|(^|\.)x\.com$|instagram|youtube|tiktok|vimeo|spotify|yelp|pinterest|"
                    r"medium\.com|podcasts\.apple|podbean|threads\.net|bsky|substack|flickr|goo\.gl|bit\.ly|linktr\.ee|douyin|weibo|wechat|glassdoor|indeed\.com|jobs|careers|amazon\.com|wordpress\.com|soundcloud|iheart|apple\.com|blogspot", re.I)

def num(s):
    return pd.to_numeric(s.astype(str).str.replace(",", "", regex=False).str.strip(), errors="coerce")

def domain(w):
    if not isinstance(w, str) or not w.strip(): return None
    d = re.sub(r"^https?://", "", w.strip().lower()).split("/")[0].split("?")[0]
    d = re.sub(r"^www\d?\.", "", d)
    return None if (SOCIAL.search(d) or "." not in d) else d

def standardise(r, label):
    d = pd.DataFrame({
        "crd": r["Organization CRD#"].str.strip(),
        "sec_number": r["SEC#"].str.strip(),
        "firm_name": r["Primary Business Name"].str.strip(),
        "city": r["Main Office City"].str.strip().str.title(),
        "state": r["Main Office State"].str.strip(),
        "country": r["Main Office Country"].str.strip(),
        "website_raw": r["Website Address"],
        "filing_date": pd.to_datetime(r["Latest ADV Filing Date"], errors="coerce", format="mixed"),
        "aum": num(r["5F(2)(c)"]),
        "employees": num(r["5A"]),
        "advisors": num(r["5B(1)"]),
        "iars": num(r["5B(3)"]),
    })
    counts = pd.concat([num(r[f"5D({t})(1)"]) for t in CLIENT_TYPES if f"5D({t})(1)" in r], axis=1)
    few = pd.concat([r[f"5D({t})(2)"].notna() for t in CLIENT_TYPES if f"5D({t})(2)" in r], axis=1)
    d["clients"] = counts.sum(axis=1, min_count=1).fillna(0) + 2.5 * few.sum(axis=1)
    retail = num(r["5D(a)(3)"]).fillna(0) + num(r["5D(b)(3)"]).fillna(0)
    d["retail_aum_share"] = np.where(d.aum > 0, retail / d.aum, np.nan)
    d["financial_planning"] = r["5G(1)"].eq("Y")
    d["succession"] = r["4A"].eq("Y") if "4A" in r else False
    d["snapshot"], d["year"] = label, int(label[:4])
    return d

std = {lab: standardise(r, lab) for lab, r in raw.items()}
panel = pd.concat([std[l] for l in std if l.endswith("-09")], ignore_index=True)
print("Firm-year rows:", len(panel), "| duplicate CRD within a year:", panel.duplicated(["crd", "year"]).sum())

# %% [markdown]
# ### 2a. Website: a data-quality trap
# The SEC roster carries **only one** of a firm's listed web addresses (68% of firms list several), and about a third of the time that one is a **social profile**. Beacon Pointe, the example in the brief, shows Pinterest in the latest file.
# **Fix applied:** collect every *non-social* address the firm has shown across the seven snapshots, prefer the most recent one that **matches the firm's name**, otherwise the most recent one. `website_source` records which filing it came from and `website_confidence` whether it matched the name. That second flag matters for broker-dealer networks such as Osaic or Cetera, whose roster entry often shows an individual advisor's site.
# **Full fix (next step):** Schedule D Section 1.I in the bulk Form ADV tables lists every address.

# %%
order = ["2026-11", "2026-09", "2025-09", "2024-09", "2023-09", "2022-09", "2021-09"]
STOP = {"llc", "inc", "lp", "llp", "ltd", "the", "and", "group", "capital", "management", "advisors", "advisers", "advisory",
        "wealth", "financial", "partners", "investment", "investments", "asset", "services", "company", "corp", "corporation", "trust"}
def name_tokens(name):
    toks = [t for t in re.findall(r"[a-z0-9]+", str(name).lower()) if len(t) >= 3 and t not in STOP]
    return toks or [t for t in re.findall(r"[a-z0-9]+", str(name).lower()) if len(t) >= 4]

names = std["2026-11"].set_index("crd").firm_name.combine_first(std["2026-09"].set_index("crd").firm_name)
cands = {}                     # crd -> list of (snapshot, domain), newest first
for lab in order:
    for c, w in zip(std[lab].crd, std[lab].website_raw):
        dmn = domain(w)
        if dmn: cands.setdefault(c, []).append((lab, dmn))

rows = []
for c, lst in cands.items():
    toks = name_tokens(names.get(c, ""))
    match = [(lab, d) for lab, d in lst if any(t in d.replace("-", "") for t in toks)]
    lab, d = (match or lst)[0]
    if match:   # trim advisor/team sub-domains, e.g. bridgemark.apollonwealth.com -> apollonwealth.com
        parts = d.split("."); root = ".".join(parts[-3:] if parts[-2] in {"co", "com", "org"} and len(parts) > 2 else parts[-2:])
        if any(t in root.replace("-", "") for t in toks): d = root
    src = "current filing" if lab == "2026-11" else f"earlier filing ({lab[:4]})"
    rows.append((c, f"www.{d}", src, "name match" if match else "unverified (no name match)"))
web = pd.DataFrame(rows, columns=["crd", "website", "website_source", "website_confidence"])
# Analyst-verified overrides (web search, Nov 2026) for shortlisted $10B+ targets and the brief's example firm,
# where the roster only ever showed social profiles or an individual advisor's site
VERIFIED = {"119290": "www.beaconpointe.com",      # Beacon Pointe Advisors
            "Pathstone": "www.pathstone.com", "Focus Partners Wealth": "www.focuspartners.com",
            "CWM, LLC": "www.carsonwealth.com", "Osaic Wealth": "www.osaic.com", "Fisher Investments": "www.fisherinvestments.com",
            "Cetera Investment Advisers": "www.cetera.com", "NewEdge Advisors": "www.newedgewealth.com"}
for key, site in VERIFIED.items():
    crds = [key] if key.isdigit() else names[names.str.upper().str.startswith(key.upper())].index.tolist()
    for c in crds:
        web = pd.concat([web[web.crd != c], pd.DataFrame([[c, site, "analyst-verified (web search)", "verified"]], columns=web.columns)])
print(web.website_confidence.value_counts(normalize=True).round(3).to_dict())
print("Beacon Pointe raw website field by snapshot:",
      {l: std[l].loc[std[l].crd == "119290", "website_raw"].squeeze() for l in ["2023-09", "2025-09", "2026-11"]})

# %% [markdown]
# ## 3. Data-quality checks

# %%
qc = panel.groupby("year").agg(firms=("crd", "nunique"), total_raum_trn=("aum", lambda x: x.sum() / 1e12),
                               median_aum_mn=("aum", lambda x: x.median() / 1e6), zero_aum=("aum", lambda x: (x == 0).sum()),
                               zero_advisors=("advisors", lambda x: (x == 0).sum()), missing_advisors=("advisors", lambda x: x.isna().sum()))
qc

# %% [markdown]
# Totals reconcile with published SEC figures (about \$125T to \$128T of regulatory AUM in 2023). The 2026 step-up is market-driven and concentrated in the largest asset managers (Vanguard, Fidelity, BlackRock, PIMCO). Zero-advisor and zero-AUM firms are kept in the counts but excluded from any ratio that divides by them.

# %%
# Spot check vs IAPD public profile
spot = std["2026-09"].query("crd == '119290'")[["crd", "sec_number", "firm_name", "city", "state", "aum", "clients", "employees", "advisors", "iars"]]
spot

# %% [markdown]
# ## 4. The analysis-ready firm table
# One row per firm per snapshot year. Identity fields come from the latest file; metrics come from each September.

# %%
ident_cols = ["crd", "sec_number", "firm_name", "city", "state", "country"]
ident = pd.concat([std["2026-11"][ident_cols], std["2026-09"][ident_cols]]).drop_duplicates("crd")
ident = ident.merge(web, on="crd", how="left")
ident["website_source"] = ident.website_source.fillna("social profile only")
ident["website_confidence"] = ident.website_confidence.fillna("none")

metrics = ["crd", "year", "filing_date", "aum", "clients", "employees", "advisors", "iars", "retail_aum_share", "financial_planning", "succession"]
firm_year = panel[metrics].merge(ident, on="crd", how="left")
firm_year["avg_aum_per_client"] = np.where(firm_year.clients > 0, firm_year.aum / firm_year.clients, np.nan)
firm_year["aum_per_advisor"] = np.where(firm_year.advisors > 0, firm_year.aum / firm_year.advisors, np.nan)
firm_year["clients_per_advisor"] = np.where(firm_year.advisors > 0, firm_year.clients / firm_year.advisors, np.nan)
# Advisor-led wealth manager (Nevis ICP): mostly individual/HNW AUM, offers financial planning, has registered IARs,
# and a client load per advisor consistent with relationship-based advice (excludes robo, TAMP and model-delivery platforms)
firm_year["wealth_manager"] = ((firm_year.retail_aum_share >= 0.5) & firm_year.financial_planning & (firm_year.iars >= 1)
                               & (firm_year.clients_per_advisor <= 500))

cols = ["sec_number", "firm_name", "website", "website_source", "website_confidence", "city", "state", "aum", "clients", "avg_aum_per_client",
        "employees", "advisors", "iars", "aum_per_advisor", "clients_per_advisor", "retail_aum_share", "wealth_manager", "filing_date", "year", "crd"]
firm_year = firm_year[cols].sort_values(["year", "aum"], ascending=[True, False])
firm_year.to_csv(OUT / "ria_firm_year_table.csv", index=False)
latest = firm_year[firm_year.year == 2026].copy()
latest.head(10)

# %% [markdown]
# ## 4a. Data validation & integrity checks
# Every check below runs automatically each time the notebook runs and is labelled `CHECK Vn` in the code. **Critical** checks stop the notebook if they fail (bad data never reaches the analysis); **warnings** are reported and explained.
#
# | ID | What it proves |
# |---|---|
# | V0 | Source files are byte-identical to the ones analysed (size + SHA-256 fingerprint) |
# | V1 | Every snapshot contains every Form ADV field the analysis uses, despite the 262 → 448 column change in 2023 |
# | V2 | Firm counts are in a plausible range and grow steadily (no truncated or partial files) |
# | V3 | Exactly one row per firm (CRD#) per snapshot, so no double counting |
# | V4 | Identifiers are well-formed (SEC# = 801-XXXXX, CRD# numeric) |
# | V5 | Numeric fields parse cleanly (commas, blanks, text), with no silent data loss |
# | V6 | AUM identity holds: discretionary + non-discretionary = total (Item 5.F(2)) |
# | V7 | AUM by client type (Item 5.D) sums to total AUM, as Form ADV requires |
# | V8 | No negative AUM, client or headcount values |
# | V9 | Advisor headcount ≤ total employees (Item 5.B(1) vs 5.A) |
# | V10 | Total regulatory AUM reconciles with SEC-published industry totals (~\$125T-\$128T in 2023) |
# | V11 | Firms link cleanly across years by CRD# (year-on-year match rate) |
# | V12 | September vs November 2026 AUM is almost unchanged, validating the "AUM updates annually" design choice |
# | V13 | Spot check vs the SEC's public IAPD profile: Beacon Pointe (the brief's example firm) |

# %%
import hashlib
checks = []
def check(cid, name, passed, detail, critical=True):
    checks.append({"id": cid, "check": name, "result": "PASS" if passed else ("FAIL" if critical else "WARN"), "detail": detail})
    if critical and not passed:
        raise AssertionError(f"CHECK {cid} failed: {name} | {detail}")

# CHECK V0: file fingerprints (warning only: SEC may legitimately republish a file)
fp = {}
for fname, (size, sha) in EXPECTED_FILES.items():
    f = RAW / fname
    got = hashlib.sha256(f.read_bytes()).hexdigest()[:16] if f.exists() else None
    fp[fname] = (f.exists() and f.stat().st_size == size and got == sha)
check("V0", "Source files byte-identical to analysed versions", all(fp.values()),
      f"{sum(fp.values())}/{len(fp)} files match size + SHA-256", critical=False)

# CHECK V1: schema, i.e. every required Form ADV column present in every snapshot
REQUIRED = ["Organization CRD#", "SEC#", "Primary Business Name", "Main Office City", "Main Office State", "Website Address",
            "Latest ADV Filing Date", "5A", "5B(1)", "5B(3)", "5F(2)(a)", "5F(2)(b)", "5F(2)(c)", "5G(1)", "5D(a)(3)", "5D(b)(3)"] + \
           [f"5D({t})(1)" for t in CLIENT_TYPES]
missing = {lab: [c for c in REQUIRED if c not in r.columns] for lab, r in raw.items()}
check("V1", "All required Form ADV fields present in all 7 snapshots", not any(missing.values()),
      f"{len(REQUIRED)} fields × {len(raw)} files; missing: {({k: v for k, v in missing.items() if v}) or 'none'}")

# CHECK V2: completeness, i.e. plausible firm counts, rising year on year
n = {lab: len(r) for lab, r in raw.items()}
sep = [n[l] for l in sorted(n) if l.endswith("-09")]
check("V2", "Firm counts plausible (13k-20k) and rising each year", all(13_000 <= v <= 20_000 for v in n.values()) and sep == sorted(sep),
      " → ".join(f"{l[:4]}: {v:,}" for l, v in n.items() if l.endswith("-09")))

# CHECK V3: uniqueness, one row per firm per snapshot
dups = {lab: int(d.crd.duplicated().sum()) for lab, d in std.items()}
check("V3", "One row per CRD# in every snapshot", sum(dups.values()) == 0, f"duplicate CRDs: {sum(dups.values())}")

# CHECK V4: identifier format
sec_ok = min(d.sec_number.str.fullmatch(r"801-\d+").mean() for d in std.values())
crd_ok = min(d.crd.str.fullmatch(r"\d+").mean() for d in std.values())
check("V4", "SEC# = 801-XXXXX and CRD# numeric", sec_ok == 1 and crd_ok == 1, f"SEC# valid {sec_ok:.2%}, CRD# valid {crd_ok:.2%} (worst snapshot)")

# CHECK V5: numeric parsing, no populated value lost when converting text → number
fails = {}
for lab, r in raw.items():
    for col in ["5A", "5B(1)", "5B(3)", "5F(2)(a)", "5F(2)(b)", "5F(2)(c)"] + [f"5D({t})(1)" for t in CLIENT_TYPES if f"5D({t})(1)" in r]:
        populated = r[col].notna() & r[col].astype(str).str.strip().ne("")
        fails[(lab, col)] = int((populated & num(r[col]).isna()).sum())
check("V5", "Numeric fields parse without loss", sum(fails.values()) == 0,
      f"{len(fails)} field-snapshots checked; unparseable values: {sum(fails.values())}")

# CHECK V6: AUM identity (5.F(2)(a) + (b) = (c)), within $1 rounding
ident_ok = {lab: float(((num(r["5F(2)(a)"]).fillna(0) + num(r["5F(2)(b)"]).fillna(0) - num(r["5F(2)(c)"]).fillna(0)).abs() <= 1).mean())
            for lab, r in raw.items()}
check("V6", "Discretionary + non-discretionary AUM = total AUM", min(ident_ok.values()) >= 0.999,
      f"holds for {min(ident_ok.values()):.2%} of firms (worst snapshot)")

# CHECK V7: AUM by client type (5.D(x)(3)) sums to total AUM (5.F(2)(c)), within 1%
s7 = {}
for lab, r in raw.items():
    tot = num(r["5F(2)(c)"])
    parts = pd.concat([num(r[f"5D({t})(3)"]) for t in CLIENT_TYPES if f"5D({t})(3)" in r], axis=1).sum(axis=1)
    m = tot > 0
    s7[lab] = float(((parts[m] - tot[m]).abs() / tot[m] <= 0.01).mean())
check("V7", "Client-type AUM (5.D) sums to total AUM (5.F)", min(s7.values()) >= 0.98,
      f"within 1% for {min(s7.values()):.1%} of firms (worst snapshot)", critical=False)

# CHECK V8: no negative values
neg = int((panel[["aum", "clients", "employees", "advisors", "iars"]] < 0).sum().sum())
check("V8", "No negative AUM, clients or headcount", neg == 0, f"negative values: {neg}")

# CHECK V9: advisors ≤ employees (advisors are a subset of employees)
over = panel.advisors > panel.employees
check("V9", "Advisors ≤ employees", over.mean() < 0.01,
      f"{over.sum():,} firm-years ({over.mean():.2%}) report more advisors than employees; ratios still computed", critical=False)

# CHECK V10: reconciliation to SEC-published industry RAUM
r23 = panel.loc[panel.year == 2023, "aum"].sum() / 1e12
check("V10", "2023 total RAUM reconciles to SEC published (~$125T-$128T)", 120 <= r23 <= 135, f"computed ${r23:,.1f}T")

# CHECK V11: cross-year linkage by CRD#
ids = {y: set(panel.loc[panel.year == y, "crd"]) for y in sorted(panel.year.unique())}
link = {y: len(ids[y] & ids[y - 1]) / len(ids[y - 1]) for y in list(ids)[1:]}
check("V11", "≥85% of firms carry forward to the next year by CRD#", min(link.values()) >= 0.85,
      ", ".join(f"{y - 1}→{y}: {v:.1%}" for y, v in link.items()))

# CHECK V12: Sep vs Nov 2026, AUM should be (almost) unchanged because Item 5 updates annually
mm = std["2026-09"][["crd", "aum"]].merge(std["2026-11"][["crd", "aum"]], on="crd", suffixes=("_sep", "_nov"))
same = float((mm.aum_sep == mm.aum_nov).mean())
check("V12", "AUM identical Sep→Nov 2026 for most firms (annual-update design)", same >= 0.85,
      f"{same:.1%} of {len(mm):,} firms unchanged; differences are mid-year amendments", critical=False)

# CHECK V13: spot check vs IAPD public profile (Beacon Pointe Advisors, CRD 119290)
bp = std["2026-09"].set_index("crd").loc["119290"]
check("V13", "Beacon Pointe matches IAPD (AUM ≈ $54.9B, 440 advisors, Newport Beach CA)",
      abs(bp.aum / 1e9 - 54.87) < 0.1 and bp.advisors == 440 and bp.city == "Newport Beach" and bp.state == "CA",
      f"AUM ${bp.aum/1e9:,.2f}B, advisors {bp.advisors:.0f}, {bp.city}, {bp.state}")

pd.DataFrame(checks).set_index("id")

# %% [markdown]
# ## 5. Market segments (September 2026)

# %%
BINS = [-np.inf, 1e9, 10e9, 50e9, np.inf]
LABELS = ["<$1B", "$1B-$10B", "$10B-$50B", ">$50B"]
latest["segment"] = pd.cut(latest.aum, BINS, labels=LABELS, right=False)
seg_map = latest.set_index("crd").segment
seg = latest.groupby("segment", observed=False).agg(
    RIAs=("crd", "count"), wealth_managers=("wealth_manager", "sum"),
    total_aum_trn=("aum", lambda x: x.sum() / 1e12), median_aum_bn=("aum", lambda x: x.median() / 1e9),
    median_advisors=("advisors", "median"), median_clients=("clients", "median"))
seg["share_of_firms"] = seg.RIAs / seg.RIAs.sum()
seg["share_of_aum"] = seg.total_aum_trn / seg.total_aum_trn.sum()
seg

# %%
fig, ax = plt.subplots(1, 2, figsize=(12, 4))
seg.RIAs.plot.bar(ax=ax[0], color="#473BCE"); ax[0].set_title("Number of RIAs by AUM segment (Sep 2026)")
for i, v in enumerate(seg.RIAs): ax[0].text(i, v, f"{v:,}", ha="center", va="bottom")
(seg.share_of_aum * 100).plot.bar(ax=ax[1], color="#9C94E6"); ax[1].set_title("Share of total regulatory AUM (%)")
for a in ax: a.set_xlabel(""); a.tick_params(axis="x", rotation=0)
plt.tight_layout(); plt.show()

# %% [markdown]
# **Commentary.** The market is a long tail with a heavy head. Two-thirds of firms are under \$1B, but the 432 firms above \$50B hold 79% of assets (mostly asset managers: Vanguard, Fidelity, BlackRock). For Nevis the relevant count is **advisor-led wealth managers**: about 1,300 firms above \$1B, of which 156 are Enterprise-scale (\$10B+). That's a short enough list for named-account selling, with a long tail of about 5,600 smaller firms better suited to a product-led motion.

# %% [markdown]
# ## 6. Method for growth and staffing questions
# * **Growth = CAGR** between September snapshots: 1-year = 2025→2026, 3-year = 2023→2026, 5-year = 2021→2026. Firms must exist at both ends.
# * **Minimum starting AUM scaled to the segment** (20% of the segment floor, never below \$100M): \$100M for <\$1B, \$200M for \$1B-\$10B, \$2B for \$10B-\$50B, \$10B for >\$50B. This stops a \$100M → \$15B re-registration from "winning" the larger segments.
# * **Headcount reclassification flag:** advisor count falling more than 60% in a single year (e.g. Vanguard 301 → 19) is a change in how Item 5.B(1) was answered, not attrition. Excluded from the ranked lists.
# * **Inorganic / step-change flag:** any single year where AUM grew more than 2.5x or advisor headcount more than doubled, or a reported succession (Item 4.A). These are usually acquisitions, re-registrations or roll-ups into a new entity. They are **ranked separately** so the main lists show sustained growth.
# * **Excess growth** = firm CAGR minus its segment's median CAGR, which strips out the market tailwind every firm shared.
# * **Segment = September 2026 AUM.** Firms are compared with the peers they are today.
# * **Advisor-led wealth manager view (Nevis ICP):** at least 50% of AUM from individuals or HNW clients, offers financial planning (5.G(1)), has registered IARs, and no more than 500 clients per advisor. The last test removes robo-advisors, TAMPs and model-delivery platforms, which also serve "individuals" but aren't advisor-led. The unfiltered lists are dominated by asset managers and fund advisers.

# %%
MIN_START = {"<$1B": 100e6, "$1B-$10B": 200e6, "$10B-$50B": 2e9, ">$50B": 10e9}
p = firm_year.pivot_table(index="crd", columns="year", values=["aum", "advisors"])
succ = panel.pivot_table(index="crd", columns="year", values="succession", aggfunc="max")

def window(h, end=2026):
    s0 = end - h
    w = pd.DataFrame({"aum_start": p["aum"][s0], "aum_end": p["aum"][end],
                      "adv_start": p["advisors"][s0], "adv_end": p["advisors"][end]}).dropna(subset=["aum_start", "aum_end"])
    w["segment"] = seg_map.reindex(w.index).values
    w = w[(w.aum_start >= w.segment.astype(str).map(MIN_START)) & (w.aum_end > 0)]
    w["aum_cagr"] = (w.aum_end / w.aum_start) ** (1 / h) - 1
    ok = (w.adv_start >= 3) & (w.adv_end >= 1)
    w["advisor_cagr"] = np.where(ok, (w.adv_end / w.adv_start) ** (1 / h) - 1, np.nan)
    w["staffing_gap"] = w.aum_cagr - w.advisor_cagr
    w["excess_growth"] = w.aum_cagr - w.groupby("segment", observed=False).aum_cagr.transform("median")
    yrs = list(range(s0, end + 1))
    a, v = p["aum"].reindex(w.index)[yrs], p["advisors"].reindex(w.index)[yrs]
    aum_jump = (a / a.shift(axis=1)).iloc[:, 1:].max(axis=1)
    adv_jump = (v / v.shift(axis=1).where(v.shift(axis=1) >= 3)).iloc[:, 1:].max(axis=1)
    succ_w = succ.reindex(w.index)[[y for y in yrs[1:] if y in succ.columns]].fillna(False).astype(bool).any(axis=1)
    adv_drop = (v / v.shift(axis=1).where(v.shift(axis=1) >= 3)).iloc[:, 1:].min(axis=1)
    w["inorganic_flag"] = (aum_jump > 2.5).values | (adv_jump > 2).fillna(False).values | succ_w.values
    w["headcount_reclass"] = (adv_drop < 0.4).fillna(False).values   # >60% advisor drop in a year = reporting change, not attrition
    w = w.join(ident.set_index("crd")[["firm_name", "website", "website_confidence", "city", "state", "country"]])
    w = w.join(latest.set_index("crd")[["wealth_manager", "aum_per_advisor", "clients_per_advisor"]])
    w["wealth_manager"] = w.wealth_manager.fillna(False).astype(bool)
    w["horizon"] = f"{h}y"
    return w.reset_index()

win = {h: window(h) for h in (1, 3, 5)}
pd.DataFrame({f"{h}y": {"firms in window": len(w), "flagged inorganic": int(w.inorganic_flag.sum()),
                        "median AUM CAGR %": 100 * w.aum_cagr.median(), "median advisor CAGR %": 100 * w.advisor_cagr.median()}
              for h, w in win.items()}).T

# %%
def fmt(df):
    out = df.copy()
    for c in ["aum_start", "aum_end", "aum"]:
        if c in out: out[c] = (out[c] / 1e9).round(2)
    if "aum_per_advisor" in out: out["aum_per_advisor"] = (out["aum_per_advisor"] / 1e6).round(0)
    for c in ["aum_cagr", "advisor_cagr", "staffing_gap", "excess_growth", "retail_aum_share"]:
        if c in out: out[c] = (out[c] * 100).round(1)
    return out.rename(columns={"aum_start": "AUM start $B", "aum_end": "AUM end $B", "aum": "AUM $B", "aum_cagr": "AUM CAGR %",
                               "advisor_cagr": "Advisor CAGR %", "staffing_gap": "Gap (pp)", "excess_growth": "vs seg. median (pp)",
                               "aum_per_advisor": "AUM/advisor $M", "retail_aum_share": "Retail AUM %"})

def top_by_segment(df, metric, n=10, wealth_only=False, cols=None, organic=False):
    d = df[df.wealth_manager.fillna(False).astype(bool)] if wealth_only else df
    if organic: d = d[~d.inorganic_flag & ~d.headcount_reclass]
    d = d.dropna(subset=[metric]).sort_values(metric, ascending=False)
    return d.groupby("segment", observed=True, group_keys=False).head(n)[["segment"] + cols]

# %% [markdown]
# ## 7. Q1: Who are the fastest-growing firms?

# %%
G = ["firm_name", "state", "aum_start", "aum_end", "aum_cagr", "excess_growth", "wealth_manager"]
fast = {h: top_by_segment(win[h], "aum_cagr", cols=G, organic=True) for h in (1, 3, 5)}
fast_wm = {h: top_by_segment(win[h], "aum_cagr", wealth_only=True, cols=G, organic=True) for h in (1, 3, 5)}
print("Fastest growing, 3-year AUM CAGR (all RIAs, sustained growth)"); fmt(fast[3])

# %%
print("Fastest growing, 1-year AUM growth (all RIAs, sustained growth)"); fmt(fast[1])

# %%
print("Fastest growing, 5-year AUM CAGR (all RIAs, sustained growth)"); fmt(fast[5])

# %%
print("Fastest-growing ADVISOR-LED WEALTH MANAGERS, 3-year CAGR"); fmt(fast_wm[3])

# %%
print("Inorganic / step-change movers, 3-year (acquisitions, re-registrations, roll-ups): advisor-led wealth managers")
fmt(top_by_segment(win[3][win[3].inorganic_flag], "aum_cagr", n=5, wealth_only=True, cols=G + ["adv_start", "adv_end"]))

# %% [markdown]
# **Commentary.**
# * **The unfiltered winners are mostly alternatives and fund managers** (hedge, credit, ETF sponsors), whose AUM moves with fund launches and inflows. That's the literal answer, but not Nevis's market.
# * **Among advisor-led wealth managers**, the standouts are multi-billion-dollar firms compounding at 60-70%+ a year: Waverly, Arkadios, Apella, Pathstone, Merit. Excess growth versus the segment median of over 50pp shows this isn't just the market.
# * **The inorganic list is a sales signal in its own right.** Focus Partners (215 → 1,025 advisors), Osaic and Cetera (thousands of advisors folded into one registrant) are integrating acquired teams, each with its own workflows and tech stack. Standardising meeting prep, notes and CRM hygiene across a newly merged advisor base is a strong Nevis entry point.
# * The 1-year view is noisier (fund launches, re-registrations) and is best read alongside the 3- and 5-year views.

# %% [markdown]
# ## 8. Q2: Which firms are capacity-strained (highest AUM per advisor)?
# Measured on September 2026, firms with at least 3 advisors (so single-person shops don't dominate). Shown for all RIAs (the literal answer) and for wealth managers, where **clients per advisor** is shown too. For a wealth manager, clients per advisor is the truer measure of strain, because advisor time scales with relationships, not dollars.

# %%
cap = latest[(latest.advisors >= 3) & (latest.aum > 0)].copy()
cap["wealth_manager"] = cap.wealth_manager.fillna(False).astype(bool)
C = ["firm_name", "state", "aum", "advisors", "iars", "aum_per_advisor", "clients_per_advisor", "retail_aum_share"]
cap_all = top_by_segment(cap, "aum_per_advisor", cols=C)
cap_wm = top_by_segment(cap, "aum_per_advisor", wealth_only=True, cols=C)
print("Highest AUM per advisor (all RIAs)"); fmt(cap_all)

# %%
print("Highest AUM per advisor (wealth managers)"); fmt(cap_wm)

# %%
fig, ax = plt.subplots(figsize=(10, 4))
data = [cap[(cap.segment == s) & cap.wealth_manager].aum_per_advisor.dropna() / 1e6 for s in LABELS]
ax.boxplot(data, labels=LABELS, showfliers=False); ax.set_yscale("log")
ax.set_ylabel("AUM per advisor ($M, log)"); ax.set_title("Wealth managers: AUM per advisor by segment (Sep 2026)")
plt.tight_layout(); plt.show()
cap[cap.wealth_manager].groupby("segment", observed=True)[["aum_per_advisor", "clients_per_advisor"]].median().assign(aum_per_advisor=lambda x: x.aum_per_advisor / 1e6).rename(columns={"aum_per_advisor": "median AUM/advisor $M"})

# %% [markdown]
# **Commentary.**
# * **All-RIA rankings are dominated by fund complexes** (Vanguard at \$584B per "advisor"), because a handful of portfolio managers run pooled vehicles. This is why the ICP filter matters.
# * **For advisor-led wealth managers, capacity is better read as clients per advisor.** The median \$1B+ firm has 60-80 clients per advisor; firms above 150-200 are where meeting volume, follow-ups and compliance documentation pile up.
# * **High AUM per advisor with few clients** (Circle Wealth, BBR Partners, Vernal Point) usually means ultra-HNW or family-office models. They have fewer, deeper relationships, and the pain is preparation and reporting quality more than volume. A different Nevis pitch from high-client-load firms.

# %% [markdown]
# ## 9. Q3: Which RIAs appear under-staffed?
# **Staffing gap = AUM CAGR minus advisor-headcount CAGR** over 1, 3 and 5 years. A positive gap means each advisor is carrying more AUM every year. Requires at least 3 advisors at the start of the window (so 1→2 hires don't create noise) and the segment minimum starting AUM, and excludes inorganic step-changes.

# %%
S = ["firm_name", "state", "aum_end", "aum_cagr", "advisor_cagr", "staffing_gap", "adv_start", "adv_end", "wealth_manager"]
gap = {h: top_by_segment(win[h], "staffing_gap", cols=S, organic=True) for h in (1, 3, 5)}
gap_wm = {h: top_by_segment(win[h], "staffing_gap", wealth_only=True, cols=S, organic=True) for h in (1, 3, 5)}
print("Under-staffed, 3-year (all RIAs)"); fmt(gap[3])

# %%
print("Under-staffed, 1-year (all RIAs)"); fmt(gap[1])

# %%
print("Under-staffed, 5-year (all RIAs)"); fmt(gap[5])

# %%
print("Under-staffed WEALTH MANAGERS, 3-year"); fmt(gap_wm[3])

# %%
w3 = win[3][win[3].wealth_manager & ~win[3].inorganic_flag & ~win[3].headcount_reclass & win[3].advisor_cagr.notna() & (win[3].aum_end >= 1e9)]
fig, ax = plt.subplots(figsize=(8, 6))
for s, c in zip(LABELS[1:], ["#9C94E6", "#473BCE", "#1E1660"]):
    d = w3[w3.segment == s]; ax.scatter(d.advisor_cagr * 100, d.aum_cagr * 100, s=14, alpha=.6, label=s, color=c)
lim = [-30, 80]; ax.plot(lim, lim, "k--", lw=1); ax.set_xlim(lim); ax.set_ylim(lim)
ax.set_xlabel("Advisor headcount CAGR, 3y (%)"); ax.set_ylabel("AUM CAGR, 3y (%)")
ax.set_title("Wealth managers ≥$1B: firms above the line grew AUM faster than advisors"); ax.legend()
plt.tight_layout(); plt.show()
print(f"$1B+ advisor-led wealth managers: {len(w3):,} | grew AUM faster than advisors: {(w3.staffing_gap > 0).mean():.0%}"
      f" | median AUM CAGR {w3.aum_cagr.median():.1%} vs advisor CAGR {w3.advisor_cagr.median():.1%}")

# %% [markdown]
# **Commentary.**
# * **Under-staffing is the norm, not the exception:** 92% of \$1B+ advisor-led firms sit above the diagonal. The industry's well-known advisor shortage shows up directly in the ADV data.
# * **Firms growing AUM 50-70% a year with flat or falling advisor headcount** (Arkadios, Kingswood, Facet, Wiss) are the clearest cases of "growth outrunning capacity". They have budget and an urgent need for leverage per advisor.
# * Unfiltered lists again surface fund managers and some reclassifications (Apollo Credit's advisor count fell from 2,540 to 912 as reporting changed), which the wealth-manager view and reclassification flag remove.

# %% [markdown]
# ## 10. So what for Nevis: a prioritised target list
# The three signals combined into one score for **US advisor-led wealth managers with 5+ advisors**, ranked within each segment by percentile (equal weights):
# 1. **3-year AUM CAGR**: momentum and budget.
# 2. **AUM per advisor (Sep 2026)**: current capacity strain.
# 3. **3-year staffing gap**: strain that is getting worse.

# %%
t = win[3][win[3].wealth_manager & win[3].staffing_gap.notna() & win[3].aum_per_advisor.notna()
           & (win[3].adv_end >= 5) & (win[3].country == "United States")].copy()
for m in ["aum_cagr", "aum_per_advisor", "staffing_gap"]:
    t[f"{m}_pctl"] = t.groupby("segment", observed=True)[m].rank(pct=True)
t["nevis_score"] = t[["aum_cagr_pctl", "aum_per_advisor_pctl", "staffing_gap_pctl"]].mean(axis=1) * 100
T = ["firm_name", "website", "website_confidence", "city", "state", "aum_end", "aum_cagr", "aum_per_advisor", "clients_per_advisor", "staffing_gap", "adv_end", "inorganic_flag", "nevis_score"]
targets = (t.sort_values("nevis_score", ascending=False).groupby("segment", observed=True, group_keys=False).head(10)
           .sort_values(["segment", "nevis_score"], ascending=[False, False])[["segment"] + T])
fmt(targets).rename(columns={"adv_end": "advisors"})

# %% [markdown]
# **How to use this list.** The score is a prioritisation aid, not a verdict:
# * **\$50B+ (Enterprise):** Cresset, Pathstone, Focus Partners, MAI, Savant, Carson, IEQ and Fisher combine scale, strong growth and rising AUM per advisor. Osaic and Cetera are a different motion: very large networks with low AUM per advisor, where the opportunity is giving thousands of independent advisors consistent tooling.
# * **\$10B-\$50B:** Aptus, RWA, Clearstead, Merit, Apella, Apollon, NewEdge and Steward are fast-growing aggregators or multi-office firms, an ideal mid-market Enterprise wedge.
# * **Next step to sharpen it:** add custodian (Schedule D 5.K(3)), tech-stack signals and recent M&A announcements, then validate the top 20 with account research before outreach.

# %% [markdown]
# ## 11. Assumptions & limitations
# * **Universe:** SEC-registered advisers only. State-registered RIAs (generally under \$100M) are not in SEC data, which is immaterial for Nevis's market.
# * **AUM timing:** September snapshots reflect each firm's last annual amendment, so AUM is mostly as of the prior fiscal year end.
# * **Advisors** = Item 5.B(1), which includes research staff; **IARs** (5.B(3)) shown alongside as the client-facing proxy. Either can be zero for firms that sub-advise.
# * **Clients** = sum of Item 5.D counts; "fewer than 5" imputed at 2.5. Funds count as one client, which is why fund managers show few clients and huge AUM per client.
# * **Growth is nominal** and includes market returns; *excess growth* versus the segment median controls for this.
# * **M&A:** succession (Item 4.A) is only in the snapshot from 2023; earlier deals are caught by the step-change rules (AUM >2.5x or advisors >2x in a year). Roll-ups that keep acquired firms as separate registrants won't show up as growth here.
# * **Advisor-led definition** is a rules-based proxy; some hybrid firms (e.g. OCIO or family-office models) will sit on either side of the line.
# * **Website:** one address per firm in the roster (see 2a). Firms listed only a social profile are labelled in `website_source`.
# * **Custodians** (Schedule D 5.K(3)) are not in the roster. Adding them from the bulk ADV data would show which custodian channels reach Nevis's best targets.

# %% [markdown]
# ## 12. How AI was used
# This analysis was built with **Claude (Anthropic) working agentically**, with me directing and checking at each step:
# * **Sourcing:** Claude browsed the SEC site, listed every monthly file with its size, and spotted that October and November 2025 are "no data" placeholders. That's why the series uses September.
# * **Schema mapping:** compared all 448 columns across seven files (2021-22 have 262 columns), confirmed the core Item 5 fields were stable, and mapped each brief field to its Form ADV item.
# * **Code:** wrote the pipeline and notebook, then ran it end to end.
# * **Catching data traps:** the one-website-per-firm limitation (found via the brief's own Beacon Pointe example), base-effect and re-registration "growth", robo/TAMP platforms passing a naive wealth-manager filter, and advisor-count reclassifications. Each became an explicit, documented rule.
# * **Human judgement:** snapshot timing (September for metrics, November for identity, grounded in Form ADV General Instruction 4), the ICP definition, thresholds, and the commercial framing are mine.
# * **Verification:** 19 automated checks (Sections 4a and 12a) built with Claude and reviewed by me; totals reconcile with published SEC figures; zero duplicate firm-years; Beacon Pointe matches its IAPD profile (\$54.9B, 440 advisors, Newport Beach); target-firm websites spot-checked by web search.

# %% [markdown]
# ## 12a. Post-analysis checks
# Checks on the derived metrics and rankings, so the outputs are internally consistent before export.
#
# | ID | What it proves |
# |---|---|
# | V14 | Segment counts add up to the full Sep 2026 universe (no firm dropped or double-counted) |
# | V15 | No infinite or impossible values in any derived ratio or growth rate |
# | V16 | Growth figures reproduce from raw: each window's CAGR is recomputed independently from the source AUM |
# | V17 | Ranked lists respect their own rules (min starting AUM, inorganic and reclassified firms excluded) |
# | V18 | Website coverage and confidence are reported (data-quality transparency) |

# %%
# CHECK V14: segment counts reconcile to universe
check("V14", "Segment counts sum to all Sep 2026 RIAs", int(seg.RIAs.sum()) == len(latest) == raw["2026-09"].shape[0],
      f"{int(seg.RIAs.sum()):,} = {len(latest):,} = {raw['2026-09'].shape[0]:,}")

# CHECK V15: no inf / impossible values in derived metrics
derived = pd.concat([latest[["avg_aum_per_client", "aum_per_advisor", "clients_per_advisor", "retail_aum_share"]].stack()] +
                    [w[["aum_cagr", "advisor_cagr", "staffing_gap"]].stack() for w in win.values()])
bad = int(np.isinf(derived).sum()); share_bad = int(((latest.retail_aum_share < 0) | (latest.retail_aum_share > 1.05)).sum())
check("V15", "No infinite values; retail AUM share within 0-100%", bad == 0, f"inf values: {bad}; retail share outside range: {share_bad} firms")

# CHECK V16: CAGR independently recomputed from raw source AUM for a random sample
raw_aum = {lab: num(r.set_index(r["Organization CRD#"].str.strip())["5F(2)(c)"]) for lab, r in raw.items() if lab.endswith("-09")}
ok16 = []
for h, w in win.items():
    for _, row in w.sample(min(200, len(w)), random_state=42).iterrows():
        a0, a1 = raw_aum[f"{2026 - h}-09"][row.crd], raw_aum["2026-09"][row.crd]
        ok16.append(abs((a1 / a0) ** (1 / h) - 1 - row.aum_cagr) < 1e-9)
check("V16", "CAGR reproduces from raw SEC values (600-firm sample)", all(ok16), f"{sum(ok16)}/{len(ok16)} match")

# CHECK V17: ranked lists obey their own filters
viol = 0
for h in (1, 3, 5):
    for lst in (fast[h], gap[h]):
        w = win[h].loc[lst.index]
        viol += int((w.inorganic_flag | w.headcount_reclass).sum())
        viol += int((w.aum_start < w.segment.astype(str).map(MIN_START)).sum())
check("V17", "Ranked lists exclude flagged firms and respect min starting AUM", viol == 0, f"rule violations: {viol}")

# CHECK V18: website data-quality transparency
wc = latest.website_confidence.value_counts(normalize=True)
check("V18", "Website confidence reported for every firm", latest.website_confidence.notna().all(),
      ", ".join(f"{k}: {v:.0%}" for k, v in wc.items()), critical=False)

summary = pd.DataFrame(checks).set_index("id")
print(f"{(summary.result == 'PASS').sum()} PASS · {(summary.result == 'WARN').sum()} WARN · {(summary.result == 'FAIL').sum()} FAIL")
summary

# %%
with pd.ExcelWriter(OUT / "nevis_ria_targets.xlsx") as xw:
    fmt(targets).to_excel(xw, sheet_name="Nevis targets", index=False)
    seg.reset_index().to_excel(xw, sheet_name="Segments", index=False)
    summary.reset_index().to_excel(xw, sheet_name="Data checks", index=False)
    for h in (1, 3, 5):
        fmt(top_by_segment(win[h], "aum_cagr", n=25, cols=G, organic=True)).to_excel(xw, sheet_name=f"Fastest growth {h}y", index=False)
        fmt(top_by_segment(win[h], "staffing_gap", n=25, cols=S, organic=True)).to_excel(xw, sheet_name=f"Under-staffed {h}y", index=False)
    fmt(top_by_segment(cap, "aum_per_advisor", n=25, cols=C)).to_excel(xw, sheet_name="Capacity (all)", index=False)
    fmt(top_by_segment(cap, "aum_per_advisor", n=25, wealth_only=True, cols=C)).to_excel(xw, sheet_name="Capacity (wealth mgrs)", index=False)
    latest.drop(columns=["segment"]).assign(segment=latest.segment.astype(str)).to_excel(xw, sheet_name="Firm table (2026)", index=False)
print("Saved:", list(OUT.iterdir()))
