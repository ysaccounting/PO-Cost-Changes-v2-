# PO Cost Changes (v2 — new DB export)

Flask web app that turns the **new TicketVault DB export** of PO cost changes
into QBO-ready Bills/Expenses workbooks. It's a rebuild of the original
`PO_Cost_Changes` app for the richer source data: same outputs, simpler
inputs.

## What changed from v1

| | v1 (old) | v2 (this app) |
|---|---|---|
| Source | `PO_Cost_Changes.xlsm` export (two-row header) | New DB export — single-header sheet named `Sheet` |
| Purchase Details upload | **Required** (used to exclude matched PO #s) | **Removed entirely** |
| Row exclusion | PO match *and* `Remove = X` | **`Remove = X`**, or **created same day as the cost change** |
| `Total Adjustment` | a source column | **derived = `TicketCostTotal − InitialTicketCostTotal`** (End − Start) |
| Adjustment date | local date in the export | **`AdjustedDateTimeUTC` converted UTC → US Central**, then date only |
| Outputs | Combined + per-company + zip | **same** |

Everything else — company mapping, the team-rename rule, the TC-vs-open
offset, the Bills/Expenses ledger logic, the Summary tab, and all workbook
styling — is unchanged from the proven v1 pipeline.

> **Note on `PurchaseDetailMatchFound`:** the new export ships this flag
> (plus `MatchedPurchaseOrderID` / `MatchedPurchaseDetailID`). Per spec it is
> **ignored** — every row is processed regardless of match status.

## Layout

```
app.py                ← Flask app (routes + background job)
processor.py          ← pipeline + new-export reader/normalizer + process_files()
mapping.py            ← loads Master_Mapping_List.xlsx (unchanged)
teams.py              ← major-league teams accessor (used by vendor_rules)
vendors.py            ← TC-vs-open offset rule (third tab of Vendors_Open.xlsx)
vendor_rules.py       ← full vendor-rename pipeline + order-number cleaning
                        (ported from the Purchase Details app)
templates/
  index.html          ← drag-and-drop UI (single upload zone)
data/
  Master_Mapping_List.xlsx
  major_league_teams.xlsx
  Vendors_Open.xlsx
docs/
  Section1.m          ← original Power Query M, kept for reference
tests/
  test_transform.py
samples/              ← (local only, gitignored) real export files for testing
```

## Local dev

```bash
pip install -r requirements.txt
python app.py                # http://localhost:5000/
pytest                       # 13 tests; the reconciliation test runs if samples/ exists
```

Or with gunicorn (matches Railway):
```bash
gunicorn app:app --workers 1 --worker-class gthread --threads 4 --timeout 120
```

## How it works

1. User drops one or more new-export `.xlsx` / `.xlsm` / `.csv` files.
2. `POST /upload` saves them and kicks off a background thread running
   `process_files()`, which computes the data + per-company stats and pickles
   the intermediate DataFrames (it does **not** build the output files yet).
3. UI polls `GET /status/<job_id>` until `done` or `error`.
4. On done, the UI opens a **"Select companies"** modal (checkbox list of every
   QBO company, with Select all). The user picks which companies to include.
5. `POST /configure/<job_id>` with the selected companies runs
   `build_filtered_outputs()`, which builds the combined workbook, per-company
   Expenses files, and per-company bills files for just those companies, and
   writes them to disk.
6. The UI then shows a date-range badge (US Central), a summary table, a
   "Download All (.zip)" button, and grids of per-company download buttons.

In the combined workbook, **Source Data and Excluded are always the full
upload**; the Combined, Summary, Bills, and per-company tabs reflect only the
selected companies.

## New-export columns used

The reader maps these and drops the rest (seat-level detail, IDs, venue, etc.):

| New column | Internal name |
|---|---|
| `CompanyName` | Company (→ mapped to QBO) |
| `PurchaseOrderID` | PO # |
| `Vendor` | Vendor (+ team-rename rule) |
| `PerformerName` | Team/Performer |
| `InitialTicketCostTotal` | Ticket Cost Total Start |
| `TicketCostTotal` | Ticket Cost Total End |
| `IsCancelled` | Cancelled (`True`→`Yes`) |
| `UpdateUser` | User |
| `AdjustedDateTimeUTC` | Adjustment Date (UTC → US Central, date only) |
| `Remove` *(manual)* | row exclusion when value is `X` |

`Total Adjustment` is derived (`End − Start`). A positive adjustment becomes a
**Bill**; a negative one becomes an offsetting **Expense** debit/credit pair.

The expense pair's offset (credit) line depends on the **final (renamed)
vendor**: if that vendor appears on the **third tab** of
`data/Vendors_Open.xlsx` ("Consolidated (without company)", the single
`Account` column), the offset Category is **`<Vendor> (TC)`**; otherwise it is
**`Due from Vendors - Open`**. Matching is case-insensitive. (The offset leg
shows on the per-company tabs / per-company files; the combined workbook's
`Combined` tab is a single-row-per-event ledger that only carries the
`Inventory Asset` leg.)

### Aggregation keys

Rows are summed into one output row when they share **Company, Adjustment Date,
Vendor, Team/Performer, AccountEmail, and ExtPONumber**. (Adding `AccountEmail`
and `ExtPONumber` to the key set makes the grouping finer than v1's — it does
not change totals, only how rows are split.)

### Memo / Description format

Both the **Memo** and **Description** columns on the Bills/Expenses ledgers read:

```
Performer / account email / ext PO # / Cost Changes (Company)
```

A blank ext PO (or email) is **omitted entirely** rather than left as an empty
segment — e.g. `Hamilton / x@y.com / Cost Changes (YSA)` (no ext PO), or
`Chicago Cubs / Cost Changes (YSA)` (no email or ext PO). All three memo
columns (Memo, Memo2, Team/Performer) end with a **`(PO created date …)`**
suffix listing every distinct PO-created date in the aggregated group, e.g.
`… / Cost Changes (YSA) (PO created date 03/09/2026)` or, when a group spans
multiple created dates, `(PO created date 12/17/2025, 03/20/2026)`. The suffix
is omitted when the group has no created date. `ExtPONumber` is read as text so
long numeric PO numbers keep full precision. `CreatedDate` is shown on the
Source Data / Excluded tabs as a US-Central date with no timestamp.

## Row exclusion

A row is excluded from the output ledgers (Combined / Bills / Expenses /
per-company) for either of two reasons, and excluded rows are listed on the
**Excluded** tab of the combined workbook (Source Data still shows everything):

1. **`Remove = X`** — a manual flag the accounting team adds in a `Remove`
   column (case-insensitive, value `X`).
2. **Same created/adjusted day** — the PO's `CreatedDate` and
   `AdjustedDateTimeUTC` fall on the **same US-Central calendar day**. Both
   timestamps are UTC and are compared after conversion to Central. A blank
   `CreatedDate` is never a match (the row is kept).

The reconciliation invariant holds across the split:
`Source Data total − Excluded total = Combined total`.

## Company labels & per-company files

Each company has two names, mirroring the Purchase Details app:

- a **short sheet/file/UI label** (e.g. `GK`, `Ticket Guy`, `Chase`, `Waxler`), and
- the **Company-column value** written into the data (e.g. `YSKG`, `Ticket Guy`,
  `Chase (Jacks)`, `YSW (Waxler)`).

Only these four companies are renamed in-data; every other company keeps its
raw TicketVault name in the Company column (e.g. `YS-Seatgeek2`). The short
labels are defined in `FILE_LABELS`, the Company-value renames in
`COMPANY_VALUE_RENAMES`, and the QBO→Company-value map in `DISPLAY_NAMES`.

The downloadable **individual company files lead with their data sheet
(Expenses or Bills), then carry a `Summary` pivot (same Company › Vendor ›
Description layout as the combined file, scoped to that company), then
`Source Data` and `Excluded` tabs** — also scoped to that company. So each file
is self-contained: the company's pivot, its own input rows, and exactly which
of them were removed. (The Excluded tab is colored red when that company had
nothing excluded.) Expenses and bills ship as separate files; a company with no
negative adjustments gets no expenses file, and one with no positive
adjustments gets no bills file.

## Per-company "Bills" files (QBO import format)

In addition to the combined workbook and the full per-company workbooks, the
app produces a **separate bills file per company** in the Purchase Details
output layout, for direct QBO import. One file per QBO company that has at
least one positive adjustment. Columns:

`Company · Bill No. · PO Created · Account · Vendor · Memo2 · Team/Performer · Memo · Total Cost · Seasons`

- **Company** — the raw/original company (e.g. `YS-Seatgeek2`).
- **Bill No.** — random 8-digit integer, one per row.
- **PO Created** — the Adjustment Date (the filename's date).
- **Account** — `Inventory Asset`.
- **Memo2** — `Performer / email / ext PO #` (no company, blanks omitted).
- **Team/Performer** and **Memo** — the full `… / Cost Changes (Company)` memo.
- **Total Cost** — the positive adjustment.
- **Seasons** — left blank for manual entry.

Only positive adjustments (the Bills-tab items) appear here; negatives remain
in the Expenses output. These files download individually from the UI and are
bundled under a `Bills/` folder in the "Download All" zip. The combined
workbook's **Bills tab uses this same PD layout**, and the per-company files'
row counts and totals tie out exactly to it.

## Vendor renaming (`vendor_rules.py`)

Ported in full from the Purchase Details app. Applied per-row after company
mapping (so the company-gated rules can see the raw company) and before the
first groupby (so renamed vendors drive aggregation and the memo). The order
is significant and mirrors the source app:

1. Pre-map fixes: `Box Office → Default Vendor`; `Live Nation Flex → Concert
   Seasons`; `Broadway Groups → Broadway Seasons`; and for YSA companies only,
   `Live Nation → Concert Seasons`.
2. `Ticketmaster AM` / `Ballpark` → the **Team/Performer** if it's a known
   major-league team, otherwise the **VenueName**.
3. `VENDOR_REPLACEMENTS` substring list (~60 entries: `AXS → Veritix`, the
   `Live Nation <venue>` normalizations, MLS team fixups, etc.).
4. `Sports Extras` → VenueName.
5. Ticket Guy box office: a `Default Vendor` row → `BROADWAY_VENUES[venue]`
   (or `Box Office - New World Stages`).
6. `Concert Seasons` → `CONCERT_SEASONS_MAP[venue]`.
7. `Broadway Seasons` → collapse Team/Performer to `Various`, then
   `BROADWAY_SEASONS_MAP[venue]`.
8. `Tickets.com` → Team/Performer if MLB, else VenueName.
9. Proper-case the Vendor, then fix `76ers` / `49ers` casing.

Major-league team membership comes from `data/major_league_teams.xlsx` (the
same file the rest of the app uses; verified identical to the source app's
hardcoded set). The large venue/vendor lookup tables are kept inline in
`vendor_rules.py`, as in the source app, so the two can be diffed.

> Like the source app, step 9's blanket title-casing produces a few quirks
> (e.g. `Toronto FC → Toronto Fc`); only `76ers`/`49ers` are corrected. This
> is faithful to the Purchase Details behavior.

## Order-number cleaning

`ExtPONumber` is blanked (matching the Purchase Details `clean_ext_po`) when:

- the value is a **UUID**, or
- it's an **all-numeric string of 19+ digits**, or
- the row's **Vendor is `Concert Seasons`** (always), or
- the row's **Vendor is `Ticketmaster AM`** *and* the order number is a
  **15+-digit number**.

This runs **before** vendor renaming (as in the source app, where
`clean_ext_po(df_raw)` precedes `build_all_query`), so the Concert Seasons /
Ticketmaster AM check sees the **raw** vendor names — a qualifying
`Ticketmaster AM` order number is blanked even though the vendor later resolves
to a team or venue. Other order numbers (e.g. `N4Q-KZGM9MZ`, or short numeric
ones on Ticketmaster AM) are kept.

## Endpoints

| Method | Path | Returns |
|--------|------|---------|
| GET | `/` | HTML drag-and-drop UI |
| POST | `/upload` | `{job_id}` — kicks off background processing |
| GET | `/status/<job_id>` | `{status, date_range, all_companies, stats, excluded, ignored_companies, ...}` |
| POST | `/configure/<job_id>` | `{selected_companies}` → builds the chosen companies' files; returns `{status, companies, bills_companies, ...}` |
| GET | `/download/<job_id>/combined` | combined multi-tab xlsx |
| GET | `/download/<job_id>/company/<name>` | one company's Expenses xlsx |
| GET | `/download/<job_id>/bills/<name>` | one company's PD-format bills xlsx |
| GET | `/download/<job_id>/all` | zip of combined + selected per-company files |

## Timezone

`AdjustedDateTimeUTC` is treated as UTC and converted to **US Central**
(`America/Chicago`, DST-aware) before the date is taken. This means the
date-range badge can differ from the export's filename date — e.g. a row at
`2026-06-01T04:51 UTC` lands on `2026-05-31` Central. The `tzdata` package is
pinned in `requirements.txt` so the zone resolves inside slim containers.

To change the timezone, edit `LOCAL_TZ` in `processor.py`.

## Company mapping

`data/Master_Mapping_List.xlsx` is the source of truth: only rows whose
`CompanyName` maps to a QBO company appear in output. All 15 company names seen
in the sample data map cleanly. Rows from companies not in the master are
skipped and surfaced in the `ignored_companies` field for visibility. To update
the mapping, edit the master file and redeploy (or point `MASTER_MAPPING_PATH`
at a different file).

## Deploy to Railway

```bash
git init && git add . && git commit -m "Initial commit"
# Push to GitHub, then in Railway: New Project → Deploy from GitHub
```

`railway.json` pins the start command and healthcheck. `Dockerfile` and
`Procfile` are also provided.

## Carried-over operational notes

- The accounting-team **`Remove = X`** reminder and the
  vendor-categorization reminder are kept on the UI. Rows are also excluded
  automatically when the PO was created the same Central day as the cost
  change (see "Row exclusion" above).
