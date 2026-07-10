# Kitchen Docs Builder

Programmatic rebuild of Brad & Lisa Kitchen portfolio documents from the live GitHub dashboard.

## Setup
```bash
npm install
```

## Current refresh workflow (as of July 2026)

Pull fresh `index.html` from GitHub, extract data to `data_clean.json`, then:

**Programmatic docs (regenerate from scratch — bulletproof, instant):**
```bash
node build_detail.js        # Kitchen_Pillars_Detail_Dollars.docx
node build_threepillars.js  # Kitchen_ThreePillars.html
```

**Patch-based docs (preserve hand-written content, update figures only):**
- `Kitchen_Action_Items.docx` — log new trades, update income summary
- `Kitchen_Complete_Reference.docx` — update figures, preserve narrative/contacts

**SKIPPED (not regenerated — Brad revising which he needs):**
- Kitchen_Summary_Dollars.docx (build_summary.js still works if needed)
- Kitchen_Pillars_Condensed_v2.docx (build_condensed.js still works if needed)
- Kitchen_Advisor_Presentation.html

## Income model (Option 1 — all-in income)
- `computeAllInIncome()` in lib.js computes P2/P3 real dividends (~$9,830/yr)
- "Total All-In Income" = P1 + Cap One + P2/P3 divs (~$91,700/yr)
- P2/P3 dividends are harvested to cash and redeployed into Pillar 1
- Footnote wording: "incl. $X/yr P2/P3 divs" (not "+") — already baked into total
- Pillar 3 projections use total return rates; actual compounding tracks below-avg
  (6.6%) since dividends are redirected out by design

## skip_div tickers (bogus div field — 401k funds store balance not per-share div)
GPool, VGIntl, BREQI1, WTSCER, Cash, SWVXX, FDRXX

## Data source
- Live dashboard: bradk-md.github.io/household-portfolio
- Raw: raw.githubusercontent.com/bradk-md/household-portfolio/main/index.html
- Data between /*DATA_START*/ and /*DATA_END*/ markers
- Brad's live Schwab quote is authoritative for CEF NAV discounts (web data is stale)
- data_clean.json is regenerated each session — never commit (always stale by next)
