# ABB Report Specification v1.0

### Single source of truth for the ABB Energy Efficiency Appraisal (EEA) Report Production Engine

**Status**: Draft for review. Prepared 2026-07-23 by systematic reverse-engineering of all existing report-related assets in the `ABB-Appraisal-Tool` repository (branch `feature/standard-report-implementation`). No code was written or modified in the production of this document. Every claim is either directly observed in an artifact (cited by path, and by `file:line` for code) or explicitly marked **Requires Verification**. Nothing in this document is invented.

**Sources studied**:
- 3 approved, human-delivered "gold standard" reports and their raw input data (`tests/eea-report-24312_method_A.zip`, `tests/eea-report-24314_method_B.zip`, `tests/eea-report-24418_NEMA_Var.zip`)
- 11 Word template files (4 in production under `src/report_templates/`, 7 archived/experimental under `legacy/archive/report_generation/report_templates/`)
- 4 production report-generation scripts (`src/generate_report_{standard,cee,poland,xlatam}.py`, 4,301 lines total) and 4 in-progress refactor files (`legacy/archive/report_generation/`, 2,238 lines)
- The Saving Calculations Excel engine (`src/fill_saving_calculations.py`, `src/excel_clone_poc.py`, 3 Excel templates)
- The Flask application (`src/app.py`, `src/templates/index.html`) and its legacy counterpart, and git history

---

## 0. Executive Summary — the critical finding

**The reports the current production code generates do not have the same page structure as the approved gold-standard reports.**

The three sample bundles in `tests/` contain real, approved, ABB-branded reports (`MSA EEA Report_XXXX.docx`, built on the `ABB Technical Project Document.dotx` master, marked `LifecycleStatus=Approved` in their custom XML). Their page skeleton is:

> Cover → Background → Summary → Business Case → Sustainability → Energy savings (Top 10) → Recommendation details (Top 10) → Appendix: Energy savings (All Assets) → Notes on Payback/NPV → Recommendation details (All Assets) → Input details (All Assets) → *[Operating hours (All Assets) — Method B fleets only]*

The four production Word templates the live app actually fills (`src/report_templates/ea_report_template_{standard,cee,poland,xlatam}.docx`, run by `src/generate_report_*.py`) have a **different, older skeleton**:

> Cover → Background → Summary (single combined KPI table + 2 charts) → Energy savings (Top 10) → Application Details (Top 10 motor specs) → Appendix: Application Details (All Assets) + Details of Recommendation (All Assets) + Calculation Methodology + NPV Methodology text

The production templates have **no "Business Case" section** (no payback donut, no NPV/IRR donuts, no payback-bucket bar chart, no sensitivity chart) and **no "Sustainability" section** (no CO2/vehicle-equivalent KPI block as its own page) as distinct sections — some of that content is folded into a single denser Summary table instead. The gold-standard reports also present the Top-10 table with **13 columns in a different order**, a separate **"Recommendation details"** table (distinct from the current pipeline's "Application Details"), and an appendix **"Input details"** table that has no equivalent in the current templates at all. Column sets, section order, and even which KPIs get their own chart differ in specific, catalogued ways (see §4–§7).

**Implication**: this is not a bug-fix or formatting task. Before any code is written, the team must confirm which structure is the actual current target — it is possible the sample `.docx` templates under `src/report_templates/` are simply stale (last touched for an India/"ION EXCHANGE" case, never updated to the current "MSA EEA Report" master) while the *approved* report format has since moved on to the `ABB Technical Project Document.dotx`-based layout captured in `tests/`. **This document specifies the report structure as observed in the three approved gold-standard samples** (§4–§7), since the project objective states the approved manual report is the gold standard the engine must match. Every place the current code/templates diverge from that structure is flagged inline. **Requires Verification** with the business owner: is the `tests/` structure the one true current target, are both formats still in active use for different customers, or is a third, still-newer format now standard? This single question should be resolved before Milestone 1 begins.

---

## 1. Overall Report Architecture

### 1.1 What the report is

An ABB Energy Efficiency Appraisal (EEA) report is a Word document (built on ABB's `ABB Technical Project Document.dotx` master template) that presents the business case for replacing/upgrading a customer's motor and drive fleet with ABB energy-efficient equivalents. It is produced per customer/site, from two pieces of raw input (an assessment-tool export of engineering nameplate data, and a pre-computed financial/engineering results export), merged through an Excel "Saving Calculations" workbook, and rendered into a fixed, ABB-branded page sequence covering: legal/background boilerplate, headline savings KPIs, the financial business case (NPV/IRR/payback), sustainability impact (CO2, EV-equivalents), a ranked shortlist of the best assets ("Top 10"), and a full appendix listing every assessed asset.

### 1.2 Report "axes" of variation (confirmed, not assumed)

Two **independent** axes of variation exist, and must not be conflated:

1. **Region/locale axis** — determines currency, default tariff/CO2/tax/discount assumptions, and region-specific financial columns. Confirmed variants in the current codebase: **Standard** (default/global), **CEE** (Central & Eastern Europe — adds a government-subsidy-adjusted investment/payback), **Poland** (adds White Certificate incentive-scheme earnings/payback), **X-LATAM** (Latin America — omits a "Take-Back" equipment-recycling column). The gold-standard sample set additionally demonstrates a **NEMA/US** locale (USD, horsepower instead of kW, "Motor Efficiency %" instead of IE-class, an added "Enclosure" field).
2. **Calculation-methodology axis** — how motor duty-cycle/load is captured and therefore how energy savings are computed. Confirmed variants: **Method A** (single average-load/average-frequency point) and **Method B** (a 9-band flow-percentage operating-hours histogram, which adds an extra "Operating hours – All Assets" appendix table). These two methods were observed applied to the *same* underlying fleet in the sample data, confirming the axis is genuinely orthogonal to region.

A third, UI-level "row-selection" variant exists in the current/legacy codebase but is **not wired into the live app**: an archived **"Top 10" vs "Complete/All Assets"** report flavor (`legacy/archive/report_generation/report_templates/ea_report_template_standard_top 10.docx` / `_all_assets.docx`), where the only difference is how many ranked assets populate the tables. In the gold-standard samples this distinction is not a separate report flavor at all — both a Top-10 view *and* an All-Assets appendix are present in every single report, one after the other (§4).

### 1.3 High-level component architecture (current system, as-built)

```
┌───────────────┐    ┌──────────────────────┐    ┌───────────────────────┐    ┌──────────────────────┐
│  Assessment    │    │  Saving Calculations  │    │   Business Objects     │    │   Renderer /          │
│  (raw export)  │───▶│  (Excel engine)       │───▶│   (Python dicts, one   │───▶│   Word Document       │
│  2 xlsx files  │    │  fill_saving_         │    │   per asset + KPI      │    │   OOXML text/table    │
│  per customer  │    │  calculations.py +    │    │   scalars)             │    │   surgery on a        │
│                │    │  excel_clone_poc.py   │    │                        │    │   pre-authored .docx  │
└───────────────┘    └──────────────────────┘    └───────────────────────┘    └──────────────────────┘
```

This is the conceptual pipeline; §9 documents exactly how (and how imperfectly) it is wired together in the real code — notably, the "Business Objects" stage is **re-derived independently inside the renderer**, not passed forward from the Excel stage, which is one of the most important architectural facts for the rebuild (§9.3).

### 1.4 Report "identity" metadata (observed in gold-standard `.docx` custom XML)

Every approved report's `customXml/item1.xml` carries ABB-schema metadata: `LifecycleStatus=Approved`, `SecurityLevel=Public`, `DocumentRevisionId=A`. **Requires Verification**: whether the rebuilt engine is expected to also stamp this metadata (it is plausible this is a manual, human sign-off step applied after generation, not something the generator itself should assert).

---

## 2. Report Execution Flow (current system)

*(This section documents the system as it exists today, to ground the roadmap in §11 — it describes mechanism, not the target design.)*

### 2.1 Entry point and routes

`src/app.py` (Flask, `Procfile`: `gunicorn src.app:app`) exposes three routes:

- **`GET /`** — serves the upload UI (`src/templates/index.html`).
- **`GET /defaults/<key>`** — returns default economic assumptions (currency, tariff, CO2, tax, discount) for a bundled region key, used to pre-fill the UI.
- **`POST /generate`** — the main single-report flow (see 2.2).
- **`POST /generate_multi`** — a multi-sheet Excel-only flow (no Word report path exists for this mode).

### 2.2 `/generate` flow, step by step

1. A fresh, **never-cleaned-up** OS temp directory is created per request (`tempfile.mkdtemp()`).
2. The Excel template is resolved: a bundled region key (`standard`/`france`/`poland`) selects one of the 3 shipped `.xlsx` templates, or the user uploads a custom one.
3. One or more assessment-export ZIPs are uploaded and validated.
4. Optional economic-assumption overrides (currency/tariff/CO2/tax/discount) are parsed from the form.
5. `run_fill()` (`src/fill_saving_calculations.py`) populates the Saving Calculations workbook from the ZIPs.
6. If report generation is requested (customer + plant + `gen_report=1`), the app selects a report-generation script by a `report_type` form field (`standard`/`cee`/`poland`/`xlatam`), defaulting to whatever the chosen region template implies, and invokes it as a **child process**: `subprocess.run([sys.executable, script_path, excel_out, customer, plant, rpt_date], cwd=BASE_DIR, timeout=180)`.
7. The child script writes its `.docx` to **its own script directory** (`src/`), not the request's temp dir; `app.py` locates it by globbing a sanitized-filename pattern, zips it together with the filled Excel workbook, deletes the loose `.docx`, and returns the zip.
8. All exceptions are caught at the route level and returned as a raw `str(e)` JSON 500 — no structured error taxonomy.

### 2.3 Region/report-type selection mechanism

Two independently-settable form fields choose (a) which Excel template + default assumptions (`template_key`) and (b) which Word generator/template (`report_type`), with "auto" meaning "inherit from `template_key`". There is a **naming inconsistency**: the bundled key `"france"` maps to report type `"cee"` — the same region, two different strings, hardcoded in two separate dictionaries in `src/app.py` with no shared schema. `xlatam` has no bundled Excel template at all — it is reachable only via a custom template upload plus an explicit dropdown choice.

### 2.4 legacy/ vs src/ relationship

Confirmed via commit history (`b887b4b Sync production code from legacy`, `f713a35 Sync latest legacy features to production`, `c76a308 Sync report template path fixes to production`) and byte-level diffing: **`legacy/` is the active development/staging tree; `src/` is the manually-synced deployed production tree** (Procfile points at `src.app:app`). At the current HEAD, `legacy/` is mid-refactor: the report-generation scripts and templates have been git-moved into `legacy/archive/report_generation/`, and `legacy/app.py`/`legacy/templates/index.html` have had the entire Word-report-generation feature (customer/plant fields, report-type dropdown, subprocess call) stripped out of their working tree — evidence that a rebuild has already begun by archiving the old code path out of active development, while `src/` (production) continues running the old, unchanged pipeline described above.

---

## 3. Rendering Pipeline (current system mechanics)

### 3.1 No object model, no templating library

None of the 4 production generator scripts import `python-docx`, `docxtpl`, `jinja2`, or `docx-mailmerge`. Every `.docx` is opened as a raw zip archive; `word/document.xml` is decoded to a plain string; every mutation (scalar substitution, table-row cloning, chart-data patching) is performed via regex/string-offset surgery on that string (and on chart/footer XML strings); the result is re-zipped. **This "OOXML-as-text" choice is the root architectural fact from which nearly every fragility in the current pipeline derives** — it appears identically across all four region scripts.

### 3.2 Two competing token/placeholder schemes exist in this codebase

**Scheme A — "Sample-Value Substitution + Positional Row Index"** (the scheme actually in production, used by all 4 live templates):
- The `.docx` is not a token template — it is a fully-rendered **sample report for one specific historical customer** (India / "ION EXCHANGE", currency INR). The literal sample text (`'ION EXCHANGE'`, `'8.00 INR/kWh'`, `'11 / 18'`, etc.) *is* the placeholder address: `repl(xml, old_literal, new_computed)` does an exact string replace.
- Table rows are located by **flat ordinal position**: every `<w:tr>` in the whole document is scanned into one list, and rows are addressed as `tr[N]` using hardcoded integer offsets (`get_tr_pos()`, `fill_row()`). Because rows are conditionally added/removed (padding, appendix trimming), the code carries manually-tracked offset variables (`irr_adj`, `npv_adj`, `NT`) to keep the arithmetic correct — a single upstream template edit silently misaligns every table below it, with no runtime error in the common case.
- Charts are **native embedded Word/OOXML charts** (`word/charts/chart{1,2,3}.xml`), patched by rewriting `<c:numRef>` blocks to self-contained `<c:numLit>` literal-value blocks, keyed to hardcoded sample cache values (e.g. literal `192`, `6`, `'2nd Qtr'`) that must be found verbatim in the chart XML for the patch to apply.
- **Confirmed real defect**: the footer report-date patch in the Standard/Poland/X-LATAM scripts does a literal `.replace('2025-09-19', new_date)`; the current templates' footers actually contain later dates (`2026-03-25`/`2026-04-08`), so this substitution **silently no-ops** for those three regions — the footer date is never updated. CEE uses a regex instead and is unaffected.

**Scheme B — `{{TOKEN}}` + `{{ROW_TEMPLATE}}` anchor + heading-anchored, header-name-matched tables** (a newer, not-yet-production engine, found only in `legacy/archive/report_generation/report_templates/ea_report_template_standard_all_assets.docx` / `..._top 10.docx`, built by a one-shot migration tool `legacy/archive/report_generation/tools/prepare_templates.py` whose own docstring states the intent: *"Replaces Spain-specific data values with `{{TOKEN}}` placeholders so that the report generator can do simple string substitution instead of fragile XML-node-position arithmetic."*):
- Scalar tokens are literal `{{NAME}}` text (e.g. `{{CUSTOMER}}`, `{{PLANT}}`, `{{NPV_COUNT}}`), consumed by bespoke regex code — not a real templating engine, but a disciplined convention.
- Table row cloning uses a single explicit **anchor marker** `{{ROW_TEMPLATE}}` placed in the row meant to be cloned once per asset; the row is located by scanning for that literal substring, not by index.
- Table location uses **heading-anchor text search** (e.g. anchor `'Energy savings with ABB premium efficiency solutions'`) followed by **header-row column-name matching** — columns are matched by name, not fixed index, which is structurally far more robust than Scheme A.
- A pre-flight `validate_required_tokens()` check fails fast if required tokens are missing — a real, testable contract that Scheme A entirely lacks.
- This scheme currently only exists for the Standard region (Top-10 and All-Assets sub-variants); it has not been extended to CEE/Poland/X-LATAM.

**Implication for the rebuild**: Scheme B's approach (heading-anchor lookup + `{{ROW_TEMPLATE}}` marker + named-column field maps + `{{TOKEN}}` scalars, with fail-fast validation) is the more defensible target already partially prototyped in this repo — but even it still manipulates raw XML strings rather than a real document object model, and does not yet cover every region.

### 3.3 Rendering pipeline steps (as currently implemented, Scheme A)

1. Parse CLI args (Excel path, customer, plant, date, data source).
2. Load the filled Saving Calculations workbook with `openpyxl(data_only=True)` — cached values only, formulas not evaluated (§9.2 explains why this matters).
3. Locate the calculations sheet by tolerant prefix/typo matching; read KPI scalars from **hardcoded fixed cell coordinates**, with zero header-text validation.
4. Read per-asset rows starting at a hardcoded row into a flat list of dicts (the "Business Objects").
5. Apply Top-N selection/padding/NPV-fallback business rules (§8).
6. Open the target `.docx` as a zip; decode `document.xml`.
7. Scalar substitution via literal-text `repl()`.
8. Table population via positional `<w:tr>` cloning/filling (`fill_row`, `nth_wt` for split-run cells).
9. Appendix trimming (delete Application-Details/Details-of-Recommendation tables) when the asset count is at/below a threshold (10).
10. Chart patching (`numRef`→`numLit` literal injection into the 3 embedded charts, plus rewriting one embedded chart's own mini Excel workbook).
11. Finalize (strip explicit row heights so Word auto-sizes rows; set `updateFields` in `settings.xml`) and re-zip to the output path.

---

## 4. Page-by-Page Breakdown

*Basis: the three approved gold-standard reports (`tests/eea-report-24312_method_A.zip`, `-24314_method_B.zip`, `-24418_NEMA_Var.zip`), which share one OOXML skeleton family (`ABB Technical Project Document.dotx` master). Page/section boundaries below are derived from Word section-break (`w:sectPr`) metadata; exact rendered page counts within "continuous" runs cannot be confirmed without a live Word/PDF render — flagged per-section as **Requires Verification** where relevant. Where the current production templates (`src/report_templates/*.docx`) differ, this is called out under "Current-pipeline delta."*

Reading order (confirmed identical in all 3 samples): **Cover → Background → Summary → Business Case → Sustainability → Energy savings (Top 10) → Recommendation details (Top 10) → Appendix: Energy savings (All Assets) → Notes on Payback/NPV → Recommendation details (All Assets) → Input details (All Assets) → [Operating hours (All Assets), Method B only]**.

### 4.1 Cover Page

- **Purpose**: brand the report and give the reader the report's headline identity at a glance before any content.
- **Static elements**: ABB red accent glyph; "ABB MOTION SERVICES" kicker; title "ABB Energy Efficiency Appraisal" / subtitle "Energy Savings Potential"; full-bleed background photograph (forest scene + white line-art icon chain: sensor → control panel/VFD → motor), landscape orientation; ABB logo in the first-page header only.
- **Dynamic elements**: a 7-row × 2-column customer-info table (see §6 for each scalar).
- **Data source**: assessment metadata (customer name, site, asset count) + economic assumption inputs (tariff, CO2 intensity, tax rate) that also drive the Saving Calculations workbook.
- **Business rules**: none beyond simple pass-through formatting.
- **Formatting rules**: labels use font "ABBvoice Medium" 8.5pt; table has only top/bottom hairline borders (no verticals); background image is absolutely positioned, page-relative.
- **Header/Footer**: first page uses a distinct title-page header (logo only) and empty footer; every subsequent page uses the default header (`"ABB MOTION SERVICES ENERGY EFFICIENCY APPRAISAL / <MONTH> <YEAR>"`, a literal text run generated at build time, not a Word field) and default footer (right-aligned `PAGE` field, centered `"© Copyright <YEAR> ABB. All rights reserved."`, also literal text, not a Word `YEAR` field).
- **Current-pipeline delta**: the production Standard/X-LATAM templates instead use a static cover photo with no line-art overlay described by the researching agent as differing stylistically; the exact current cover table shape (9×2 vs the gold-standard's 7×2) also differs — **Requires Verification** exact field-for-field diff, not fully enumerated.

### 4.2 Background Page

- **Purpose**: legal/contextual boilerplate — what the report is, its disclaimer, and terms & conditions.
- **Static elements**: three fixed paragraphs, **100% identical text across all 3 gold-standard samples**: (1) a "Background" paragraph explaining the report's purpose and limitations; (2) a "Disclaimer" paragraph capping ABB's liability at "USD100 (US Dollar one hundred)", with an embedded hyperlink; (3) a "Terms & Conditions" hyperlink to ABB's published Energy Appraisal T&Cs (`https://search.abb.com/library/Download.aspx?DocumentID=4MWA000145...`).
- **Dynamic elements**: none — this entire page is static boilerplate.
- **Data source**: none (hardcoded legal text, not derived from assessment data).
- **Business rules**: none.
- **Formatting rules**: `Heading 1` for the section title; hyperlinks styled blue/underlined.

### 4.3 Summary Page

- **Purpose**: headline financial/energy KPIs at a glance, with two supporting charts.
- **Static elements**: section heading "Summary"; chart titles; decorative shield+motor icon.
- **Dynamic elements**: a 1×3 layout — Col 0: two stacked KPI tiles (annual energy cost savings [all NPV-positive]; annual cost savings, Top 10); Col 1: bar chart "Annual Energy Consumption" (Current Fleet vs Upgraded Fleet, GWh); Col 2: icon + pie chart (NPV-positive fraction, no legend) + "{n}/{N} motors NPV positive" text.
- **Data source**: Business Objects KPI scalars (§6, §9).
- **Business rules**: number-abbreviation formatting rule (values ≥ ~1,000 shown as `"k"`-suffixed rounded thousands, e.g. `"EUR 116k"`; values under ~1,000 shown with one decimal place, e.g. `"USD 547.2"`) — **Requires Verification** exact threshold/algorithm; not confirmable from output alone.
- **Formatting rules**: currency symbol + space + value, no locale-specific thousands/decimal convention observed.
- **Current-pipeline delta**: the production templates combine most of this KPI content (plus some Business-Case-page content) into a single denser 6×4 KPI table alongside 2 (not necessarily the same 2) charts, rather than this dedicated 1×3 Summary layout — **Requires Verification** exact 1:1 field mapping.

### 4.4 Business Case Page

- **Purpose**: the financial investment case — payback, NPV, investment cost, IRR — with supporting distribution and sensitivity charts.
- **Static elements**: section content only (no separate "Business Case" heading was independently confirmed vs. being visually grouped with Summary — **Requires Verification** exact heading text/placement); chart titles as listed below.
- **Dynamic elements**: a 2×3 grid — Row 0: donut chart "Payback time - Top 10" (center label `"X.0 yrs"`); bar chart "# of NPV positive assets by payback time" (4 buckets: `<3 / 3-5 / 5-7 / >7 yrs`); stacked KPI tiles "Net Present Value" + "Net Present Value - Top 10". Row 1: stacked KPI tiles "Investment cost" + "Investment cost - Top 10"; bar chart "Payback time sensitivity to electricity price variation" (`-20%/-10%/0%/10%/20%` buckets, 0% bar highlighted); donut chart "Internal rate of return - Top 10" (center label `"X.X%"`, IRR is **only ever shown for the Top-10 subset** — no whole-fleet IRR figure exists anywhere in the report).
- **Data source**: Business Objects KPI scalars + the payback-bucket histogram + the sensitivity grid.
- **Business rules** (verified by cross-checking chart values against table totals):
  - The sensitivity chart's 0% bar always equals the NPV-positive aggregate payback from the All-Assets appendix totals row.
  - The payback-bucket bar chart counts only the NPV-positive subset, bucketed by individual asset payback.
  - The "Payback time - Top 10" donut does **not** match either the Top-10 table's aggregate payback or the NPV-positive aggregate payback — it is closest to (but not exactly reproducible as) an investment-weighted average of the 10 shortest-payback NPV-positive assets. **Requires Verification** with the business/calculation owner — exact formula not confidently reverse-engineerable from output data alone.
- **Formatting rules**: currency symbol + space + value for KPI tiles; percentage with 1 decimal for IRR.
- **Current-pipeline delta**: **this entire section, as a distinct page with 5 charts, has no confirmed equivalent in the current production templates**, which were only confirmed to have 2 charts (bar + pie) total in the Summary section. This is the single largest structural gap identified between gold-standard and current pipeline.

### 4.5 Sustainability Page

- **Purpose**: environmental-impact framing of the savings (CO2 avoided, EV-equivalents).
- **Static elements**: section heading "Sustainability"; lightning-bolt/car icon; caption "Savings in kWh equal the annual operation of Battery Electric Vehicles"; footnote "Assuming annual battery operation of 3.500kWh annually"; a trailing note (still part of this section in the XML) "Note: Only considering the NPV positive assets."
- **Dynamic elements**: a 1×3 layout — Col 0: two stacked KPI tiles ("{X}kt CO2 — Annual CO2 avoided", "{Y}kt CO2 — Annual CO2 avoided - Top 10"); Col 1: icon + "{N} Vehicles" text; Col 2: empty in all 3 samples (no content observed).
- **Data source**: Business Objects CO2/energy-savings scalars.
- **Business rules**: displayed CO2 figure = `round(raw_kg_total / 1000)`; "Vehicles" figure = `round(NPV-positive Energy Savings kWh / 3500)`, consistent with the footnote's stated 3,500 kWh/year assumption (this constant also appears, undocumented, throughout the current codebase — §8).
- **Formatting/labeling finding — flag for spec decision, not a guess**: the suffix used is `"kt"` (kilotonnes = 1,000,000 kg), but the displayed number is numerically `raw_kg / 1000`, i.e. it is numerically equal to **tonnes**, not kilotonnes, in all 3 samples and in both the overall and Top-10 variants — a systematic, consistent template behavior, not a one-off typo. **Requires Verification** with the business owner whether the correct unit label for the rebuild should be `"t CO2"` rather than `"kt CO2"`.
- **Minor observed quirk**: singular counts are not grammatically pluralized (e.g. "1 Vehicles") — static template text, not dynamically adjusted; decide in the rebuild whether to fix.
- **Current-pipeline delta**: no confirmed dedicated "Sustainability" page exists in current production templates as a standalone section — CO2/vehicle content, if present, would need to be located within the combined Summary KPI table. **Requires Verification**.

### 4.6 Energy Savings — Top 10 (landscape)

- **Purpose**: the ranked shortlist of the best-performing assets by energy savings, restricted to financially-viable (NPV-positive) assets.
- **Static elements**: section heading "Energy savings with ABB premium efficiency solutions – Top 10"; footnote "*Data is listed as total annual figures."
- **Dynamic elements**: see Table spec §5.1.
- **Data source**: Business Objects (per-asset dicts).
- **Business rules**: see §8.1 (Top-10 selection).
- **Formatting rules**: landscape orientation; totals block distinguished visually from data rows.
- **Current-pipeline delta**: current production Top-10 table has 12 columns and an extra hardcoded "Take-Back" checkmark column not present in the gold-standard's 13-column layout (§5.1 vs §5, current-pipeline table).

### 4.7 Recommendation Details — Top 10 (landscape)

- **Purpose**: the specific ABB equipment recommendation for each of the Top-10 assets.
- **Static elements**: section heading "Recommendation details – Top 10 Assets".
- **Dynamic elements**: see Table spec §5.2.
- **Data source**: Business Objects (recommended-equipment fields).
- **Business rules**: same row order/asset sequence as the Energy Savings Top-10 table (verified identical across all 3 samples).
- **Current-pipeline delta**: the current pipeline's nearest equivalent is called "Application Details" and shows motor **specification/nameplate** data (poles, voltage, frequency, current, power factor, etc.), not the recommended-equipment narrative fields this gold-standard table shows. These appear to be two conceptually different tables that the rebuild must reconcile — **Requires Verification**: does the current pipeline's "Application Details" table need to be renamed/repurposed, or does the gold standard need both an "Application Details" (nameplate) *and* a "Recommendation details" (proposed equipment) table? The gold-standard sample's structure suggests the latter — nameplate data appears separately in the Appendix's "Input details" table (§4.10) — so the rebuild likely needs **three** distinct per-asset table concepts (Recommendation, Input/nameplate, Energy-savings-financials), not two.

### 4.8 Appendix — Energy Savings, All Assets (landscape)

- **Purpose**: the complete, unranked fleet-wide version of §4.6, for full transparency/audit.
- **Dynamic elements**: see Table spec §5.3.
- **Business rules**: see §8.4 (sort order, all-assets scope).
- **Current-pipeline delta**: corresponds to the current pipeline's "Details of Recommendation" / "Application Details" all-assets appendix tables, but with a different column set and sort order — see §5.3 vs current-pipeline table for the exact diff.

### 4.9 Notes on Payback / NPV (landscape, static text)

- **Purpose**: disclose the exact business assumptions underpinning the payback and NPV figures, for audit/credibility.
- **Static elements** (100% identical across all 3 samples): a paragraph explaining ABB's payback methodology accounts for real-world system inefficiencies rather than idealized "cube law" calculations; a paragraph stating **"NPV is calculated assuming a discount rate of 6.5%, a useful life of 20 years and depreciation of the assets during 10 years, a constant electricity price and country corporate tax rate."**
- **Business rules**: these are hard, load-bearing business-rule constants (discount rate 6.5%, useful life 20 years, depreciation period 10 years) that any NPV re-implementation must match exactly — and which are corroborated independently by the Excel engine's own NPV formula (§9.2) and the Python fallback formula in every current generator script (§8.2), all of which use a 20-year horizon and (inconsistently — see §8.6) a discount-rate variable.
- **Current-pipeline delta**: the current templates' equivalent section is called "NPV Methodology" — **Requires Verification** whether the wording matches verbatim.

### 4.10 Appendix — Recommendation Details, All Assets (landscape)

- Same 9-column shape and per-asset content as §4.7, but scoped to the full fleet, in the same payback-ascending sort order as §4.8.

### 4.11 Appendix — Input Details, All Assets (landscape)

- **Purpose**: full nameplate/engineering input data for every asset, for audit/traceability back to the raw assessment.
- **Dynamic elements**: see Table spec §5.4. **Column set varies by calculation-method and locale variant** (17 columns for Method A / IEC-kW; 16 for Method B, dropping the average-load/frequency column; 18 for the NEMA/US variant, replacing IE-class with numeric Motor-Efficiency-%, adding an Enclosure column).
- **Data source**: near-direct passthrough of the raw "Input assets" spreadsheet, re-sorted into the same payback-ascending order as the rest of the appendix and re-formatted (rounding, thousands separators).
- **Current-pipeline delta**: **no confirmed equivalent exists in the current production pipeline** — this appears to be new/additional content relative to what the current generator scripts produce.

### 4.12 Appendix — Operating Hours, All Assets (landscape, Method-B fleets only)

- **Purpose**: disclose the full duty-cycle histogram backing the Method-B energy-savings calculation, for audit.
- **Dynamic elements**: 13 columns — `# | Customer Equipment Id | Driven load | Annual Running Hours [h] | Hours at Flow 20% | 30% | 40% | 50% | 60% | 70% | 80% | 90% | 100%`.
- **Business rules**: this table's presence/absence is itself a business rule — it appears if and only if the underlying assessment data was captured using the flow-percentage/Method-B methodology (confirmed: present in the 24314/Method-B sample, absent in the 24312/Method-A and 24418/NEMA samples).
- **Current-pipeline delta**: no confirmed equivalent in the current pipeline (the current codebase does not appear to have a Method A/B distinction as a first-class concept at all — **Requires Verification**).

---

## 5. Table Specifications

*Each table below is specified against the gold-standard reports; where the current-pipeline table differs materially, it is documented as its own separate entry so both are available to the implementation team.*

### 5.1 "Energy savings with ABB premium efficiency solutions – Top 10" (gold standard)
- **Section**: §4.6. **Purpose**: ranked shortlist of best assets by energy savings, NPV-positive only.
- **Header columns, in order**: `# | Customer/Equipment ID | Existing Energy Cons.(kWh) | Energy Cost(EUR/USD) | CO2 Consumption(kg) | Optimized Energy Cons.(kWh) | Energy Savings(kWh) | Energy Cost Savings(EUR/USD) | Investment(EUR/USD) | Energy Savings(%) | Payback(years) | NPV(EUR/USD) | Avoided CO2(kg)`.
- **Data source**: Business Objects (per-asset financial fields).
- **Sorting rule**: Energy Savings (kWh) descending.
- **Filtering rule**: NPV-positive assets only.
- **Dynamic row rule**: `min(10, count of NPV-positive assets)` rows — **not padded** to 10 when fewer qualify (verified: exactly 1 row shown for the 24418 sample, which had only 1 NPV-positive asset).
- **Formatting rules**: thousands-comma grouping, no decimals for kWh/kg/currency columns; negative values in this table were not observed (NPV-positive filter guarantees no negatives in the financial columns shown; the currency-vs-percent columns use plain formatting).
- **Validation rule**: row count must never exceed `min(10, NPV-positive count)`.
- **Totals rows** (2, appended below data): (1) `"Top 10 energy savers"` — aggregate of displayed rows; (2) `"NPV positive assets ({N})"` — aggregate of **all** NPV-positive assets (superset).

### 5.1a Current-pipeline "Energy Savings – Top 10" table (production)
- 12 or 13 columns depending on region (Standard: 12, adds a hardcoded "Take-Back" checkmark column with no data-driven source at all; CEE: 13-14, adds gross+net investment/payback; Poland: 13, adds White-Certificate earnings/payback; X-LATAM: 11, omits Take-Back with no substitute).
- **Sort/filter/cap rule**: identical across all 4 region scripts — `sorted(npv_positive, key=lambda a: (payback ascending, -energy_cost_savings as tiebreak))[:10]`, then **padded** with NPV-negative assets (by descending cost savings) up to 10 total if fewer than 10 NPV-positive assets exist (Poland pads only up to `min(10, NA)`). **This padding behavior is the opposite of the gold-standard's "don't pad" rule (§5.1)** — a concrete, verified discrepancy the rebuild must resolve deliberately, not by accident.
- **Sort key discrepancy**: current pipeline sorts by **payback ascending**; gold-standard sorts by **energy savings descending**. These produce different row orders and are not interchangeable — **flag for explicit business decision**.

### 5.2 "Recommendation details – Top 10 Assets" (gold standard)
- **Section**: §4.7. **Purpose**: the specific proposed ABB equipment per shortlisted asset.
- **Header columns, in order**: `# | Customer Equipment Id | IE | Driven load | Flow control method | Existing Connection | Output(kW or Hp) | Recommended ESS motor | Recommended ESS drive`.
- **Data source**: Business Objects (recommended-equipment fields, sourced from the assessment export's `Recommended ESS motor`/`Recommended ESS drive` fields).
- **Sorting rule**: identical row order to §5.1 (same asset sequence).
- **Formatting rules**: `IE` = efficiency class label (values observed: `"Not Known"`, or an IE-class string; one sample showed a bare numeric `"4"` — **Requires Verification** whether this is a valid encoding or a data-quality gap). "Recommended ESS motor" is a single pipe-delimited concatenated string (`{Model} | {Power} {unit} | [IE-class |] Shaft: {size} | {Voltage} V | {Frequency} Hz | {Motor family name}`). "Recommended ESS drive" can be legitimately empty for DOL-connected assets with no drive upgrade recommended.
- **Validation rule**: row count must equal the Top-10 table's row count exactly, same asset IDs in the same order.

### 5.3 "Energy savings with ABB premium efficiency solutions – All Assets" (gold standard, Appendix)
- **Section**: §4.8. **Purpose**: full-fleet audit version of §5.1.
- **Header columns, in order**: `# | Customer/Equipment ID | Existing Energy Cons.(kWh) | Energy Cost | CO2 Consumption(kg) | Optimized Energy Cons.(kWh) | Energy Savings(kWh) | Energy Cost Savings | Avoided CO2(kg) | Investment | Payback(years) | NPV` (note: drops the `%` savings column present in the Top-10 table; moves `Avoided CO2` earlier in column order relative to §5.1).
- **Sorting rule**: **Payback ascending, globally across the whole fleet** — rigorously verified against raw input serial-number order (does not match) and confirmed correct in every sample including edge cases (negative-payback/negative-savings rows sorted correctly to the tail).
- **Filtering rule**: none — every assessed asset appears.
- **Formatting rules — inconsistency flagged**: negative NPV values are rendered in **parentheses** (accounting style, e.g. `(1 749)`); negative energy-savings/cost-savings values in the same table use a **leading minus sign** instead (e.g. `-360`). This is an inconsistent-but-observed convention within the same table — **Requires Verification** whether intentional; the rebuild should pick one convention deliberately.
- **Totals row**: 1 row, `"NPV positive assets ({N})"`.

### 5.3a Current-pipeline "Details of Recommendation" / "Application Details" all-assets tables (production)
- 12 columns each, sourced from `assets_by_num` (ascending asset number, **not** payback-sorted) — this is a **verified sort-order discrepancy** vs. the gold standard's payback-ascending rule. Deleted entirely (both tables) when total asset count ≤ 10.

### 5.4 "Input details – All Assets" (gold standard, Appendix)
- **Section**: §4.11. **Purpose**: full nameplate/engineering audit trail.
- **Header columns (Method A / IEC-kW variant, 17 cols)**: `# | Customer Equipment Id | IE | Driven load | Flow control method | Existing Connection | Output(kW) | Poles | Rated voltage(V) | Rated frequency(Hz) | Rated speed(RPM) | Current(A) | Power factor | Running time(hours) | Avg. Load (%) / Frequency (Hz) | Keep Shaft Height Fixed | Downsizing Option`.
- **Method B variant (16 cols)**: identical, minus `Avg. Load (%) / Frequency (Hz)`.
- **NEMA/US variant (18 cols)**: `IE` replaced by `Motor Efficiency` (numeric %), `Output` in `Hp`, plus a trailing `Enclosure` column.
- **Data source**: direct passthrough of the raw "Input assets" spreadsheet (verified 1:1 field mapping, see §9.1), re-sorted and re-formatted.
- **Sorting rule**: same payback-ascending order as the rest of the appendix.
- **Validation rule**: column set must be selected based on which methodology/locale variant the source data represents — this is itself a business rule (§8.7).

### 5.5 "Operating hours – All Assets" (gold standard, Appendix, Method-B only)
- **Section**: §4.12. **Header columns**: `# | Customer Equipment Id | Driven load | Annual Running Hours [h] | Hours at Flow 20% | 30% | 40% | 50% | 60% | 70% | 80% | 90% | 100%`.
- **Data source**: direct passthrough of the Method-B assessment export's flow-band columns.
- **Validation rule**: the 9 "Hours at Flow X%" values for a row must sum to that row's "Annual Running Hours" value (not independently verified against every row, but true by construction of the underlying data — **Requires Verification** whether the engine should assert this at render time).

### 5.6 Cover-page customer info table (gold standard)
See §6 for each individual scalar; structurally a 7-row × 2-column key/value table, no headers, top/bottom hairline borders only.

### 5.7 Background info table (gold standard)
3-row × 2-column table: `Background` / `Disclaimer` / `Terms & Conditions` — 100% static text, no data-driven cells (see §4.2).

---

## 6. Scalar Value Specifications

*Format: label as it appears in the report → source → formatting rule → business rule → validation, based on the gold-standard reports. Current-pipeline equivalents noted where confirmed to differ.*

| Scalar (label as printed) | Source | Formatting | Business rule | Validation |
|---|---|---|---|---|
| `Customer:` | Assessment metadata | Free text, observed pattern `Country_CompanyName` | Pass-through | Non-empty required |
| `Date of Report:` | Report generation timestamp / user input | ISO `YYYY-MM-DD` | Should reflect the actual generation date — **Requires Verification**: current pipeline has a confirmed defect where this substitution silently no-ops for 3 of 4 regions (§3.2) | Must be a valid date |
| `Site address:` | Assessment metadata | Free text, `Country, Postcode City` | Pass-through | Sample data showed placeholder-looking values (e.g. "0000 Netherlands") — **Requires Verification** these are test artifacts, not a formatting defect |
| `# Of Assets:` | Count of rows in the assessment export | Integer | Pass-through count | Must equal actual row count in Business Objects |
| `Cost of electricity:` | Economic assumption input | `{rate} {CCY}/kWh`, raw decimal, no fixed precision observed | User-supplied or region default | Must be > 0 |
| `Carbon intensity:` | Economic assumption input | `{rate} kg/kWh` | User-supplied or region default | Must be > 0 |
| `Tax Rate:` | Economic assumption input | Integer % | User-supplied or region default | 0-100 |
| Annual energy cost savings (Summary KPI) | Sum of energy-cost-savings over NPV-positive assets | `k`-abbreviated ≥1000, else 1 decimal (§4.3 rule, **Requires Verification** exact threshold) | NPV-positive scope only | Must reconcile with Appendix totals row |
| Annual cost savings - Top 10 (Summary KPI) | Sum over Top-10 subset | same as above | Top-10 scope only | Must reconcile with Top-10 totals row |
| `{n}/{N} motors NPV positive` | Count of NPV-positive assets / total assets | `int/int` | NPV > 0 test | n ≤ N always |
| Net Present Value / NPV - Top 10 (Business Case KPI) | Business Objects NPV field(s) | Currency + space + value | Aggregate over NPV-positive / Top-10 scope respectively | See §8.2 for the underlying NPV formula |
| Investment cost / Investment cost - Top 10 | Business Objects investment field(s) | Currency + space + value | Aggregate over NPV-positive / Top-10 scope | — |
| Payback time - Top 10 (donut center label) | Derived | `"X.0 yrs"` | **Requires Verification** — does not match either simple aggregate (§4.4) | Flag for business-owner confirmation before re-implementing |
| Internal rate of return - Top 10 (donut center label) | Business Objects IRR, Top-10 scope only | `"X.X%"` | Newton-Raphson IRR (see §8.3), shown for Top-10 only, never whole-fleet | Convergence must be validated (0 < r < acceptance band) |
| Annual CO2 avoided / - Top 10 (kt CO2) | Business Objects CO2-avoided field(s) | `round(raw_kg/1000)` + `"kt CO2"` suffix | **Unit-label flag**: numerically equals tonnes, not kilotonnes (§4.5) — confirm intended label before rebuild | — |
| `{N} Vehicles` | `round(NPV-positive Energy Savings kWh / 3500)` | Integer + `"Vehicles"` (not pluralization-aware) | Hardcoded 3,500 kWh/vehicle-year constant, stated in the accompanying footnote | — |

**Current-pipeline scalar handling delta**: the production code reads KPI scalars from hardcoded fixed `(row, col)` Excel coordinates with **zero header-text validation** (`src/generate_report_standard.py:145-165` and identically in the other 3 region scripts) — if the Saving Calculations workbook's summary block ever shifts by one row, every KPI scalar silently reads the wrong cell with no error. This is a load-bearing fragility the rebuild must not repeat (§10).

---

## 7. Chart Specifications

*All charts in the gold-standard reports are pre-rendered raster PNG images (`word/media/*.bin`, true PNGs saved with a `.bin` extension) — there are no native OOXML chart objects (`word/charts/chartN.xml`) in the gold-standard files, in contrast to the current production templates, which use native, editable OOXML charts patched via `numRef`→`numLit` injection (§3.2). This is itself a significant, confirmed structural difference the rebuild must decide on deliberately: native editable charts (current pipeline's approach, more "Word-native" and user-editable post-delivery) vs. flattened images (gold-standard's approach, guarantees pixel-perfect fidelity regardless of the recipient's Word/font environment, at the cost of post-delivery editability). Requires Verification with the business owner which is actually wanted for the rebuild — the objective's phrase "visually ... identical to manually prepared ABB reports" suggests the flattened-image approach used in the actual approved samples is the real target.*

| Chart | Section | Purpose | Data source | Series | Axis | Labels | Formatting | Validation |
|---|---|---|---|---|---|---|---|---|
| "Annual Energy Consumption" (bar) | Summary | Compare pre/post-upgrade energy use | Business Objects: total consumption before/after | "Current Fleet" (black), "Upgraded Fleet" (red) | Y: GWh (unit auto-scales by magnitude per current-pipeline logic — **Requires Verification** whether gold-standard uses the same auto-scaling rule) | Bar values | 2-series bar | Series values must sum consistently with the Summary KPI tiles |
| NPV-positive fraction (pie, no legend) | Summary | Visualize the NPV-positive proportion of the fleet | Business Objects: NPV-positive count / total | 2 wedges (dark/light gray) | n/a | none observed | Grayscale, no data labels | Wedge proportion must equal the `{n}/{N}` scalar |
| "Payback time - Top 10" (donut) | Business Case | Headline payback figure | Business Objects (see §6 flag) | 1 value, center label | n/a | `"X.0 yrs"` center label | Donut | **Requires Verification** exact formula (§4.4) |
| "# of NPV positive assets by payback time" (bar) | Business Case | Distribution of paybacks across NPV-positive assets | Business Objects: per-asset payback, bucketed | 4 buckets: `<3 / 3-5 / 5-7 / >7 yrs` | Y: count | Bucket counts, color-coded red/black/gray/light-gray | Bucketed bar | Bucket counts must sum to total NPV-positive count |
| "Net Present Value" / "- Top 10" KPI tiles | Business Case | Headline investment-case value | Business Objects NPV | n/a | n/a | Currency value | KPI tile, not a chart | Must reconcile with appendix totals |
| "Payback time sensitivity to electricity price variation" (bar) | Business Case | Show payback's dependence on tariff assumption | Business Objects: sensitivity grid (§8.5) | 5 bars: `-20%/-10%/0%/10%/20%` | Y: payback years | `"X.X yrs"` per bar, 0% bar highlighted (red/baseline) | Bar, 0% bar = current tariff baseline | 0% bar value must equal the NPV-positive aggregate payback exactly (verified rule) |
| "Internal rate of return - Top 10" (donut) | Business Case | Headline IRR | Business Objects IRR (Top-10 scope only) | 1 value, center label | n/a | `"X.X%"` | Donut | IRR solver must converge (§8.3) |
| Sustainability icon + vehicle count | Sustainability | Environmental-impact framing | Business Objects savings kWh | n/a | n/a | `"{N} Vehicles"` | Icon, not a data chart | See BEV constant, §6/§8 |

**Current-pipeline chart delta**: only 2 charts (1 bar, 1 pie) are confirmed present, both in the combined Summary section; the Business Case section's 4 charts (payback donut, payback-bucket bar, sensitivity bar, IRR donut) have **no confirmed current-pipeline equivalent** — this is the single largest content gap between what the current code produces and the gold standard (also flagged in §0 and §4.4).

---

## 8. Business Rules (consolidated, with sources)

### 8.1 Top-10 / ranked-shortlist selection
- **Gold standard** (§5.1): filter to NPV-positive assets, sort by **Energy Savings (kWh) descending**, take `min(10, NPV-positive count)` — do **not** pad to 10 when fewer qualify (verified: 24418 shows exactly 1 row when only 1 asset is NPV-positive).
- **Current pipeline** (all 4 region scripts, `src/generate_report_*.py`, identical logic): filter to NPV-positive, sort by **(payback ascending, energy-cost-savings descending as tiebreak)**, take `[:10]`, then **pad** remaining slots with NPV-negative assets sorted by descending cost savings, up to 10 total (Poland pads only to `min(10, NA)`). If **zero** NPV-positive assets exist at all ("NO_NPV_MODE"), Top-10 becomes all assets sorted by descending cost savings, `[:10]`.
- **These are two materially different, non-interchangeable rules** (different sort key, opposite padding behavior) — this must be resolved as an explicit business decision, not silently inherited from either side.

### 8.2 NPV calculation
- **Documented business assumption** (gold-standard report text, §4.9, and corroborated independently by the Excel engine formula and every current generator script's fallback): discount rate 6.5%, useful life 20 years, straight-line depreciation over 10 years, constant electricity price, constant country corporate tax rate.
- **Excel formula** (`Saving_Calculations` sheet, all 3 region workbooks): `NPV($discount_rate, cashflow_years_1_to_20) + year_0_cashflow`, where year-0 cashflow = `-Investment`, years 1-10 cashflow = `after-tax savings + straight-line depreciation tax shield`, years 11-20 = `after-tax savings only` (no more depreciation shield).
- **Python fallback formula** (used whenever the Excel cached NPV cell is empty — which is the common case, see §9.2): `after_tax = energy_cost_savings * (1 - tax_rate); pv_factor = (1-(1+r)^-20)/r; npv = after_tax * pv_factor - investment` — a simplified annuity approximation of the Excel's year-by-year formula, **not an exact reimplementation** (it omits the depreciation-shield year-11-20 cutoff nuance) — **Requires Verification** whether this approximation is deliberately acceptable or a latent numerical-accuracy bug.
- **Confirmed bug, present identically in all 4 region scripts**: two conflicting hardcoded tax-rate fallback defaults exist in the same file (`0.349` at one call site, `0.25` at the NPV-formula call site) — **Requires Verification** whether intentional.

### 8.3 IRR calculation
- Newton-Raphson solver, 200 iterations, convergence tolerance `1e-9`, initial guess `annual_cashflow/investment`, result accepted only if `0 < r < 10` (0-1000%) — identical across all 4 region scripts. Shown for the Top-10 subset only; no whole-fleet IRR figure exists anywhere in the gold-standard report.

### 8.4 All-assets appendix sort order
- **Gold standard**: payback ascending, globally, across the entire fleet (verified rigorously, including edge cases with negative/undefined payback sorted to the tail).
- **Current pipeline**: sorts by ascending asset number (`assets_by_num`), **not** payback — a confirmed discrepancy.

### 8.5 Sensitivity analysis
- Hardcoded tariff-delta grid `[-20%, -10%, 0%, +10%, +20%]`, identical across all 4 region scripts; payback recomputed at each delta as `investment / (savings * (1+delta))` (Poland's variant additionally folds in White Certificate earnings). The gold-standard's 0% bar is confirmed to equal the NPV-positive aggregate payback exactly.

### 8.6 Appendix inclusion/trigger threshold
- **Current pipeline**: an asset-count threshold of 10 controls whether the "Application Details"/"Details of Recommendation" all-assets tables are built at all — `NA ≤ 10` removes them entirely (Standard/Poland/X-LATAM) or `NA > 10` triggers cross-template borrowing from the Standard template (CEE, see §10). This threshold and mechanism are **entirely current-pipeline-specific** — the gold-standard reports always include the full appendix regardless of asset count (all 3 samples, with fleet sizes 28/28/26, include it; **Requires Verification** what happens for very small fleets, since no sample with ≤10 assets was available to confirm the gold-standard's behavior at that boundary).

### 8.7 Calculation-methodology / locale selection
- Which "Input details" column set (§5.4) and whether an "Operating hours" appendix table (§4.12/§5.5) appears is determined by which methodology the source assessment data was captured with (Method A single-load-point vs. Method B flow-band histogram) and which locale variant (IEC/kW/IE-class vs. NEMA/Hp/Motor-Efficiency-%) the input data represents. This selection is **auto-detected from column headers present in the raw "Input assets" spreadsheet** in the current Excel-fill script (`is_hp_variant` check, `src/fill_saving_calculations.py`) — a good pattern to carry forward, since it means the *data* (not a manual UI flag) determines the report variant.

### 8.8 Region-specific financial rules
- **CEE**: net investment = gross investment − government subsidy; both gross and net payback shown side-by-side.
- **Poland**: White Certificate earnings = `energy_savings_kWh × (1/11,630 kWh-per-toe) × WC_rate_PLN_per_toe`; White-Certificate-adjusted payback = `investment / (energy_cost_savings + WC_earnings)`; WC earnings also folded into the region's IRR and sensitivity calculations.
- **X-LATAM**: omits the "Take-Back" equipment-recycling column present in the Standard region's Top-10 table (achieved via a structurally different template with one fewer column, not a runtime conditional).

### 8.9 Take-back / circularity assumptions (Excel engine only, not yet surfaced in any Word report table observed)
- Baked into the Region and France Excel templates' "Five-Year Plan" block: 100% of assets assumed taken back (`$BK$3=1`), 95% of that material assumed recirculated (`$BH$4=0.95`), 0.5 tons average material per asset (`$BH$5=0.5`). These are not exposed as user-overridable assumptions anywhere in the current UI/CLI — **Requires Verification** whether the rebuild needs to expose them.

### 8.10 Scope-3 emissions
- Excel formula: `Scope 3 Emissions = 80 kg CO2 × Output(kW)` per asset (constant `80` stored in template cell `D9`, all 3 region workbooks) — **Requires Verification** whether/where this figure surfaces in the Word report; no Scope-3 scalar or table was independently confirmed in the gold-standard sample catalog above.

---

## 9. Data Flow

```
[Assessment export — customer/site ZIP, from the ABB EEA assessment tool]
   ├── "assessment data-NNNNN.xlsx"   (fixed-position columns: Sr.No, Equip ID,
   │      Existing Energy Cons (kWh), Energy Cost, CO2 Cons, Optimized Energy Cons,
   │      Energy Savings (kWh), Cost Savings, Savings %, CO2 Reduction, Investment,
   │      Payback, NPV, Recommended ESS motor/drive [+ Hours-at-Flow columns if Method B])
   └── "Input assets-NNNNN.xlsx"      (header-name-keyed columns: Serial#, Equip ID,
          Driven Load, Dol Vsd, Flow Control, Shaft Height, Rated Power [kW|Hp], Poles,
          Running Hours, IE Eff Class|Motor Efficiency[%], Voltage/Frequency/Speed/
          Current/PF, Avg Loading[%]/Avg Frequency[Hz] [or Hours-at-Flow bands])
        │
        ▼  fill_saving_calculations.py: read_assessment_data() + read_input_assets()
           — auto-detects Method A / Method B / NEMA-HP variant from column headers present
        │
        ▼  excel_clone_poc.py: clone_saving_calculations()  (multi-sheet workbook path only)
           — OOXML-level sheet clone preserving native Excel Table + chart fidelity
           (openpyxl's own copy_worksheet() does not reliably preserve these)
        │
        ▼  fill_saving_calculations.py: writes raw INPUT-block values via openpyxl into the
           Saving_Calculations sheet (assumption cells: currency, tariff, CO2, discount, tax;
           per-asset input cells: #, Equip ID, Application, energy/savings/investment,
           motor nameplate fields)
        │
        ▼  [SAVED .xlsx — native Excel FORMULAS remain as text; their cached results are
           largely None/stale, because nothing in the pipeline ever triggers a recalculation
           (no LibreOffice/Excel COM call anywhere in the codebase). Formula-driven columns
           (Annual Energy Cost, CO2 Cons, Cost Savings, Savings %, Payback, NPV, IRR,
           Scope-3 Emissions, CEE/White-Certificate columns, Five-Year-Plan pivot, all
           Top-10/20/30 & KPI summary rollups) are therefore usually empty at this point.]
        │
        ▼  generate_report_{standard,cee,poland,xlatam}.py  (invoked as a SEPARATE CHILD
           PROCESS by app.py, given only the .xlsx file path as its input — not an
           in-process function call, and not given any of fill_saving_calculations.py's
           already-parsed Python data)
           — openpyxl.load_workbook(data_only=True): reads the (mostly empty) cached values
           — for every empty cached cell, a PARALLEL, INDEPENDENTLY-WRITTEN Python
             reimplementation of the Excel formula logic recomputes NPV/payback/CO2/IRR/
             White-Certificate math directly from the raw input cells (§8.2)
           — produces "Business Objects": a flat list of per-asset dicts + ~20 KPI scalar
             variables + a sensitivity-grid list
        │
        ▼  Renderer: OOXML text-run / table-row / chart patching of a pre-authored
           ea_report_template_{region}.docx (Scheme A — literal sample-value substitution
           and positional row/run indexing, §3.2)
        │
        ▼  Final Word .docx report, bundled together with the filled .xlsx into a zip
           by app.py and returned to the user
```

### 9.1 Verified field-level mappings (Assessment/Input → Saving Calculations)

| Source field | Destination (Saving Calculations column) | Transformation |
|---|---|---|
| assessment `Existing Energy Cons. (kWh)` | `Annual Energy. Cons (kWh)` | direct copy |
| assessment `Energy Savings (kWh)` | `Annual Energy Savings, kWh` | direct copy |
| assessment `Investment (EUR)` | `Investment` | direct copy |
| assessment `Recommended ESS motor`/`drive` | `Recommended ESS motor` / `ESS connection` | direct copy |
| Input assets `Driven Load`/`Dol Vsd`/`Flow Control` | `Application`/`Dol Vsd`/`Flow Control Method` | direct copy |
| Input assets `Rated Power [kW]` or `[Hp]` | `Output (kW/HP)` | direct copy, **no unit conversion between kW and Hp** — written as-is regardless of source unit system |
| Input assets `IE Eff Class` (or forced blank if HP-variant) | `Ie Eff Class` | direct copy; `"Not known"` substituted only when the field is genuinely absent |
| Input assets `Hours at Flow 20-100%` (Method B) | `Average Flow (%)` | **computed**: weighted average `Σ(flow% × hours) / Σ(hours)` over non-zero bands |
| Input assets `Shaft Height [mm]` | *(no matching column in any current template)* | **dead field** — written by the fill script but silently dropped; **Requires Verification** whether a "Shaft height" column exists in some other/future template |

### 9.2 The critical formula-evaluation gap

Because `fill_saving_calculations.py` writes cell values with `openpyxl` and **never triggers Excel/LibreOffice to recalculate the workbook**, essentially every formula-driven cell (NPV, IRR, payback, CO2 savings, the 20-year cash-flow vector, Scope-3 emissions, region-specific incentive columns, all KPI/Top-N rollups) is `None` when the report generator reads the file back with `data_only=True`. **The Excel formulas are therefore, in practice, not the actual source of truth for the numbers that appear in the report** — a second, independently-maintained Python reimplementation of the same financial math inside each `generate_report_*.py` is what actually produces the numbers almost all of the time. This is the single most consequential architectural finding for the rebuild: **any change to the "official" Excel calculation logic must currently be manually re-implemented in up to 4 separate Python files, or the two will silently diverge** depending on whether a given cell happens to have a cached value. A rebuilt engine must decide, deliberately, whether Excel or Python is the one true calculation authority, and eliminate the parallel-implementation risk.

### 9.3 The "Business Objects" stage is not actually passed forward

Contrary to the idealized `Assessment → Saving Calculations → Business Objects → Renderer` pipeline description, the real system has **no in-process Business Objects handoff at all**. `fill_saving_calculations.py` and `generate_report_*.py` communicate exclusively through the filled `.xlsx` file on disk, crossing a `subprocess.run()` process boundary; the report generator does its own from-scratch Excel read and its own from-scratch business-object construction, independent of (and structurally different from) whatever in-memory Python objects `fill_saving_calculations.py` built moments earlier. A rebuild should treat "Business Objects" as a real, shared, in-process (or at minimum, versioned-schema) data structure — not an implicit contract re-derived twice.

---

## 10. Dependency Map & Technical Debt

### 10.1 Module dependency diagram (current production system)

```
Browser (src/templates/index.html)
        │ multipart/form-data POST
        ▼
src/app.py  (Flask: /, /defaults/<key>, /generate, /generate_multi)
        │                                   │
   imports (Python)                subprocess.run() — CLI boundary, not an import
        ▼                                   ▼
src/fill_saving_calculations.py     src/generate_report_{standard,cee,poland,xlatam}.py
   run_fill()                          (re-reads the .xlsx from scratch, independent
   run_fill_multi()  ← DEAD, unused      Business-Object construction, §9.3)
   generate_multi_workbook()                    │ reads
        │ imports                               ▼
        ▼                              src/report_templates/*.docx
src/excel_clone_poc.py                 (Scheme A: sample-value substitution + positional
   clone_saving_calculations()          row/run indexing)
        │ reads/writes
        ▼
src/excel_templates/*.xlsx
   (Region / France-CEE / Poland variants)

Output artifacts land in: tempfile.mkdtemp() (per-request, NEVER cleaned up)
Report .docx transiently lands in: src/ itself (BASE_DIR), then is deleted after zipping
```

### 10.2 Coupling issues (app-architecture level)
- **Filesystem is the only IPC channel** between the Excel-fill stage and the report-generation stage — a full round-trip through disk, with no shared in-process data model (§9.3).
- **`subprocess.run` writes into the shared, non-request-scoped `src/` directory**, not the request's own temp dir — a genuine concurrency hazard: two simultaneous requests for the same customer/plant name can race on the same output filename (glob/delete collision).
- **Hardcoded region dictionaries** (`BUNDLED`, `REPORT_SCRIPT_MAP` in `src/app.py`) with an internal **naming inconsistency** (`"france"` key vs. `"cee"` report-type value) and no shared config schema. Adding a region today means hand-editing 4 separate places (dict entry, `.xlsx`, `.docx`, new `generate_report_X.py`) with no validation tying them together.
- **`excel_clone_poc.py` is production-load-bearing code mislabeled as a proof-of-concept** in its own docstrings and stale path constants — `app.py`'s multi-sheet route depends on it in production despite every internal signal suggesting it is experimental.
- **`run_fill_multi()` is dead code** — fully implemented and imported, never called (superseded by `generate_multi_workbook()` per its own docstring, but never removed).

### 10.3 Weak areas / operational technical debt
- Temp directories are never cleaned up on any code path (success or failure) — an unbounded disk-fill and customer-data-retention concern.
- Fully synchronous, long-running requests with a hardcoded 180s subprocess timeout inside a Flask/gunicorn sync worker — no job queue, no real progress reporting (the UI's progress bar is a fixed CSS animation, not tied to actual server state).
- No authentication/authorization, no rate limiting, 500MB upload limit — a resource-abuse/DoS surface if exposed without a gateway.
- No structured error taxonomy — every failure returns a raw exception string as JSON, potentially leaking internal details.
- No automated test coverage for the Flask app or the fill/report pipeline; only ad hoc standalone scripts under `legacy/` that print PASS/FAIL and are not wired into any CI/test runner.

### 10.4 Report-generation code technical debt (root cause and cross-cutting list)
**Root cause, one sentence**: all four production report generators treat a Word document as an opaque string blob addressed by literal text content and ordinal element position, rather than as a structured document tree — nearly every fragility below traces back to that single architectural choice (§3.1).

Cross-cutting issues, present identically in all 4 region scripts:
1. Positional `<w:tr>`/`<w:t>` indexing with manually-tracked cumulative offset variables (`irr_adj`, `npv_adj`, `NT`) — a template edit anywhere upstream silently misaligns every table below it, usually with no runtime error.
2. Scalar substitution keyed to literal historical sample text, not real tokens — any Word re-save that reflows text across XML runs silently breaks a field with zero diagnostic.
3. Hardcoded `(row, col)` KPI cell coordinates with no header-based validation, in contrast to the (partially) header-driven asset-table columns.
4. A pervasive silent-failure culture: bare/broad `except: pass`, silent truncation on run-count mismatch, silent no-ops when an anchor string isn't found, chart patches that just skip if a hardcoded literal isn't present.
5. Chart data is patched via hardcoded literal cache values (e.g. `192`, `6`, `'2nd Qtr'`, and in CEE's case, three near-duplicate float-literal fallbacks for the same number) — brittle to any template re-save.
6. No XML well-formedness validation before the final zip write — a bug in any regex/slice step can silently produce a `document.xml` Word refuses to open, while the script itself exits 0.
7. Monolithic, module-scope-only execution — none of the four files expose unit-testable functions.
8. The exact same non-critical bugs (the tax-rate 0.349-vs-0.25 inconsistency; a CO2-intensity 0.54-vs-0.054 inconsistency in Poland; a dead `SKIP_APP_DET` variable; duplicate helper-function definitions within one file) are duplicated identically across all four files — direct evidence of copy-paste-and-diverge lineage, and proof that any bugfix today must be manually ported up to four (or five, including the in-progress refactor) times.

**Duplication estimate**: roughly 80-90% of the code by volume and logic is identical or near-identical across the four region scripts. The genuine, load-bearing regional deltas are narrow: CEE's net-vs-gross investment/payback and its cross-template appendix-borrowing from the Standard template (a hardcoded, invisible cross-file coupling); Poland's White-Certificate math and its dual scalar-substitution mechanism (`repl()` for single-run text, `nth_wt()` for Word-fragmented multi-run text); X-LATAM's simple absence of the Take-Back column via a shorter template and value list; and per-region default constants.

### 10.5 Where the in-progress refactor (`legacy/archive/report_generation/`) already helps, and where it doesn't
Genuine progress: `report_table_utils.py` consolidates the two repeating asset tables into one shared, testable function that resolves columns by **header text** (not fixed index) and **fails fast** (hard `sys.exit`) if required headers or `{{TOKEN}}`s are missing — both real, valuable improvements over every production script. Not yet addressed even in this branch: the underlying mechanism is still 100% regex-over-raw-XML-string manipulation with no real OOXML object model; the KPI/NPV/IRR/sensitivity financial-math layer remains imperative and un-abstracted; chart patching is unchanged hardcoded-literal-cache surgery; and the refactor currently only covers the Standard region (Top-10/All-Assets pair), not CEE/Poland/X-LATAM — and it has inherited, rather than fixed, the tax-rate and CO2-intensity constant inconsistencies from production.

---

## 11. Implementation Roadmap

*Ordering follows the milestones specified in the project brief. Each milestone's scope has been adjusted to reflect what was actually found in this research (e.g., "Sustainability" and a "Business Case" milestone are pulled forward because the gold-standard structure treats them as first-class sections the current pipeline does not fully implement).*

**Milestone 0 — Resolve the structural ambiguity (blocking, not in the original brief but required before Milestone 1)**
- **Objective**: get explicit business-owner confirmation of which report structure (§0) is the actual current target, and freeze one gold-standard sample bundle per region/methodology combination as the canonical acceptance fixture.
- **Dependencies**: none — this is the first thing to do.
- **Success criteria**: a written, signed-off mapping of "for region X and methodology Y, the target page/table/chart structure is exactly ______", backed by an agreed set of fixture files.
- **Risk level**: **High** if skipped — every subsequent milestone risks being built against the wrong target.

**Milestone 1 — Cover Page**
- **Objective**: render the cover page (branding, background photo, customer-info table) exactly matching the gold standard's 7-row scalar table and layout (§4.1).
- **Dependencies**: Milestone 0 (confirmed target structure); a real OOXML object model or equivalent rendering approach chosen (this is also the first place the "native chart vs. flattened image" decision, §7, and the "Scheme A vs Scheme B token" decision, §3.2, must be made, since they affect every later milestone).
- **Success criteria**: generated cover page is visually indistinguishable from the gold-standard sample for at least one real customer case, across all locale variants (currency/date format).
- **Risk level**: Low (self-contained, no cross-table dependencies) — but the *decisions* made here (templating mechanism, chart strategy) carry high downstream risk if wrong.

**Milestone 2 — Summary**
- **Objective**: render the Summary section's KPI tiles + 2 charts (§4.3) matching gold-standard values and the (still-unconfirmed) rounding/abbreviation rule.
- **Dependencies**: Milestone 1; Business Objects layer must supply the annual-savings and NPV-positive-count scalars.
- **Success criteria**: KPI values reconcile exactly with the appendix totals rows (§9); chart data matches computed aggregates.
- **Risk level**: Medium — the `"k"`-abbreviation rounding rule is flagged Requires Verification and must be resolved here, not assumed.

**Milestone 3 — Sustainability**
- **Objective**: render the Sustainability section (CO2 avoided, vehicle-equivalent) as its own section (§4.5), resolving the "kt vs t" unit-label question first.
- **Dependencies**: Milestone 2 (shares the CO2/energy-savings Business Objects); a business decision on the CO2 unit label (§4.5, §6).
- **Success criteria**: unit label decision documented and implemented; vehicle-equivalent constant (3,500 kWh) implemented and cited from this spec, not re-derived ad hoc.
- **Risk level**: Low, contingent on the unit-label decision being made explicitly rather than silently copied from either source.

**Milestone 4 — Energy Savings Table (Top 10 and All Assets)**
- **Objective**: implement the Top-10 selection rule and the All-Assets appendix, resolving the two confirmed rule conflicts: sort key (energy-savings-descending vs. payback-ascending) and padding behavior (no-pad vs. pad-to-10) (§8.1, §8.4).
- **Dependencies**: Milestone 0's structural decision must explicitly cover this table, since it is the single largest confirmed logic discrepancy between gold standard and current pipeline.
- **Success criteria**: row order, row count, and all 12/13 column values reconcile exactly against the gold-standard sample for at least one Method-A and one Method-B fixture.
- **Risk level**: **High** — this table encodes the most consequential, currently-unresolved business-rule conflict in the whole project.

**Milestone 5 — Application/Recommendation/Input Details**
- **Objective**: resolve and implement the three related but distinct per-asset table concepts identified in §4.7 (Recommendation details vs. current pipeline's "Application Details" vs. the gold-standard's separate "Input details" nameplate table), plus the Method-B-only Operating Hours table (§4.12).
- **Dependencies**: Milestone 0's structural decision; Milestone 4 (shares sort order and asset-selection scope).
- **Success criteria**: all three/four table concepts are unambiguously named and scoped in the spec and implementation, each reconciling column-for-column against the gold standard.
- **Risk level**: Medium-High — the current pipeline's single "Application Details" table maps ambiguously onto what the gold standard treats as (at least) two separate tables; this needs explicit design, not a 1:1 port.

**Milestone 6 — Appendix (assembly, thresholds, cross-region variants)**
- **Objective**: implement the appendix-inclusion logic — resolve whether the current pipeline's ≤10-asset threshold-based deletion (§8.6) should be carried forward at all, since the gold standard always includes the full appendix; implement CEE/Poland/X-LATAM regional deltas (§8.8) as configuration rather than forked code, per the recommendation in §10.5.
- **Dependencies**: Milestones 4-5 (the appendix reuses those tables' rendering logic at full-fleet scope).
- **Success criteria**: no cross-template borrowing hack (§10.5's CEE finding) remains in the rebuilt engine; region deltas expressed as data/config, not four separately-maintained files.
- **Risk level**: Medium — mostly a design/consolidation risk rather than a numerical-accuracy risk, given the underlying math is already well understood from this research.

**Milestone 7 — Charts**
- **Objective**: implement the 2 (Summary) + 5 (Business Case) charts (§7), resolving the native-OOXML-chart vs. flattened-image decision explicitly.
- **Dependencies**: Milestone 2 (Summary charts) and Milestone 6 (Business Case charts depend on fully-resolved NPV/IRR/sensitivity math, §8.2-8.5); the templating-mechanism decision from Milestone 1.
- **Success criteria**: every chart's data reconciles exactly with the corresponding scalar/table figures (the cross-checks already verified manually in this research, e.g. the sensitivity 0%-bar-equals-aggregate-payback rule, should become automated tests).
- **Risk level**: Medium — the "Payback time - Top 10" donut's exact formula is still unconfirmed (§4.4) and must be resolved with the business owner before this milestone can be marked complete, not approximated.

**Milestone 8 — Validation**
- **Objective**: build an automated acceptance-test harness that renders each frozen gold-standard fixture (Milestone 0) through the new engine and diffs page-by-page/table-by-table/scalar-by-scalar against the approved original, plus unit tests for the previously-untested financial math (NPV/IRR/sensitivity) and the Excel-fill field mappings (§9.1).
- **Dependencies**: all prior milestones; also depends on deciding whether Excel formulas or Python calculations are the one true authority (§9.2) — this milestone cannot be considered complete while that ambiguity remains unresolved, since it would be validating against a moving/dual target.
- **Success criteria**: 100% reconciliation (exact page/table/scalar/chart match, not "close enough") against every frozen fixture across at least Standard/CEE/Poland/X-LATAM × Method-A/Method-B; CI-wired regression tests exist so future changes cannot silently reintroduce the divergence patterns catalogued in §10.4.
- **Risk level**: Low in execution once prior milestones are solid, but **High** as a gating function — this is the milestone that actually proves (or disproves) the project's stated objective of "visually and numerically identical to manually prepared ABB reports."

---

## 12. Consolidated "Requires Verification" register

1. **[Critical, blocking]** Which report structure is the actual current target — the `tests/*.zip` "MSA EEA Report" (`ABB Technical Project Document.dotx`) structure, the current `src/report_templates/*.docx` "EA Report" structure, or a third, newer format? (§0, Milestone 0)
2. Exact rendered page count / page breaks inside "continuous" Word section runs — only confirmable by an actual Word/PDF render, not from OOXML section metadata alone (§4).
3. Exact number-abbreviation/rounding algorithm behind the `"k"`-suffix KPI formatting (§4.3, §6).
4. Exact formula behind the "Payback time - Top 10" donut chart — does not match either simple aggregate reproducible from the data (§4.4, §7).
5. Whether `"kt CO2"` is an intentional shorthand or should be corrected to `"t CO2"` (tonnes) given the displayed numbers are numerically tonnes, not kilotonnes (§4.5, §6).
6. Whether the `IE` column's bare numeric value (e.g. `"4"`) observed in one gold-standard sample is a valid encoding or a data-quality gap (§5.2).
7. Whether negative-number notation (parentheses for NPV vs. leading minus for savings columns, observed in the same appendix table) is intentional or should be unified (§5.3).
8. Whether native OOXML charts (current pipeline) or flattened raster images (gold standard) is the intended chart strategy for the rebuild (§7).
9. Whether Excel formulas or the Python fallback reimplementation should be the one true calculation authority going forward, given they currently silently diverge (§9.2).
10. Whether a "Shaft Height" destination column exists in some other/future Saving Calculations template — the current fill script writes this field but no shipped template consumes it (§9.1).
11. Whether the take-back/circularity assumptions baked into the Excel templates (100% take-back, 95% material recirculated, 0.5 tons/asset) need to become user-overridable in the rebuilt engine (§8.9).
12. Where (if anywhere) the Scope-3 emissions figure (80 kg CO2/kW output) is meant to surface in the Word report — no corresponding scalar/table was found in the gold-standard catalog (§8.10).
13. What the current pipeline's "Application Details" table is actually meant to become relative to the gold standard's separate "Recommendation details" and "Input details" tables — one table mapping to two/three concepts (§4.7, Milestone 5).
14. Whether the two conflicting hardcoded tax-rate defaults (0.349 vs 0.25) and the CO2-intensity 0.54-vs-0.054 inconsistency (Poland) found identically across all four production scripts (and inherited by the in-progress refactor) are intentional or latent bugs (§8.2, §10.4).
15. Whether the gold-standard reports' apparent placeholder-looking site-address values are test-data artifacts or reveal a real formatting issue (§4.1, §6).
16. What the current appendix ≤10-asset inclusion threshold should become, given the gold-standard samples (all with 26-28 assets) always show the full appendix regardless of count, and no small-fleet gold-standard sample exists to confirm behavior at that boundary (§8.6).

---

*End of ABB Report Specification v1.0.*
