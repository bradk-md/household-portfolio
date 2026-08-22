# Kitchen Docs Builder

Fully programmatic rebuild of the Brad & Lisa Kitchen portfolio documents from the live GitHub dashboard. As of August 2026, **every document regenerates from the dashboard on each refresh — nothing is hand-patched, so nothing can drift.**

## Setup
```bash
npm install
```

## Refresh workflow

Pull fresh `index.html` from GitHub, extract the data block to `data_clean.json` (at `/home/claude/rebuild/data_clean.json`, the path `lib.js` reads), then run the four builders:

```bash
node build_complete_reference.js   # "Kitchen Complete Reference.docx"
node build_action_items.js         # "Kitchen Action Items.docx"
node build_detail.js               # Kitchen_Pillars_Detail_Dollars.docx
node build_threepillars.js         # Kitchen_ThreePillars.html
```

All four read live data and rebuild from scratch every time. The per-account holdings tables, pillar assignments, income figures, and projections are always current with the dashboard.

## Documents

**Kitchen Complete Reference.docx** (`build_complete_reference.js`)
Combined document. The **data half** (portfolio summary, per-account holdings for all six Pillar 1 accounts including cash/money-market positions, Pillar 2/3 holdings, income summary, PIMCO suite, crypto) is generated from `data_clean.json`. The **narrative half** (tax/IRMAA/RMD strategy, Roth conversions, action plan, contacts, notes for Lisa) is read from `narrative.json` and appended after a page break. Pillar assignments come straight from the data — e.g. TPL and SPCX sit physically in the Joint account but are Pillar 2, so they render under Pillar 2 by category. The Joint account carries a tax note (MDXBX = federal + Maryland-state-exempt muni interest; EPD K-1 = return-of-capital, tax-deferred; only ORCL ordinary taxable).

**Kitchen Action Items.docx** (`build_action_items.js`)
Task lists (Completed / Outstanding-Urgent / Outstanding-Planned) read from `action_items.json`, plus a live Income Summary pulled from `data_clean.json`. Supports dynamic placeholders — e.g. `{SCHD_REMAINING}` computes shares left to a target from live holdings.

**Kitchen_Pillars_Detail_Dollars.docx** (`build_detail.js`) — full position detail.
**Kitchen_ThreePillars.html** (`build_threepillars.js`) — three-pillar overview.

## Hand-maintained source files

Two JSON files hold content that isn't dashboard data. They are edited deliberately (by Brad, or by Claude on Brad's request) — never auto-generated:
- **narrative.json** — the strategy sections of Complete Reference. Figures that need Brad's verification are marked `[VERIFY: $X?]` and render in red until confirmed.
- **action_items.json** — the task lists. Mark items done, add new actions, adjust targets here.

## Income model (Total All-In Income)
- `computeAllInIncome()` in `lib.js` computes real P2/P3 dividends (~$9,830/yr).
- "Total All-In Income" = Pillar 1 income + Cap One interest + P2/P3 dividends (~$92K/yr).
- P2/P3 dividends are harvested to cash and redeployed into Pillar 1.
- Cap One APY is read from the data (currently 3.00%), never hardcoded.

## Projection formulas
The dashboard JS defines the growth assumptions; the Complete Reference generator mirrors them.
- **Pillar 1**: `Year N = p1Inc*(1.03)^N + capOne interest`. Dividends grow 3%/yr (distribution raises); Cap One interest added once/yr, never compounded.
- **Pillar 2**: `p2Val*(1+g)^N`, g ∈ {5% / 7% / 9%} total-return — the 401k reinvests and compounds untouched until RMDs begin 2035.
- **Pillar 3**: PRICE-ONLY. Dividends are swept out to Pillar 1, so Pillar 3's own growth is price appreciation only. Rate = total-return minus the live P3 dividend yield (`p3Inc/p3Val`, ~2.43%): avg ≈ 6.1%, conservative ≈ 4.2%. This avoids double-counting the swept dividends.

## skip_div tickers (401k funds store balance, not a per-share div)
GPool, VGIntl, BREQI1, WTSCER, Cash, SWVXX, FDRXX — excluded from P2/P3 income only; Pillar 1 counts all income including money-market interest.

## Data source
- Live dashboard: bradk-md.github.io/household-portfolio
- Raw: raw.githubusercontent.com/bradk-md/household-portfolio/main/index.html
- Data between `/*DATA_START*/` and `/*DATA_END*/` markers.
- The dashboard is the source of truth. `data_clean.json` is regenerated each session — never commit it (stale by the next refresh).
- After a push, the GitHub raw CDN can serve a stale copy for ~1 min — verify via the Contents API if a just-pushed file looks old.

## Snapshots
Each refresh writes a dated snapshot to `docs/snapshots/data-YYYY-MM-DD.json` — used as the diff baseline for the next refresh (detects trades and income changes).
