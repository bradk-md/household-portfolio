# Kitchen Docs Builder

Programmatic rebuild of the Brad & Lisa Kitchen portfolio documents from the live GitHub dashboard data.

## Setup

```bash
npm install
```

## Usage

```bash
node build_summary.js      # Kitchen_Summary_Dollars.docx
node build_condensed.js    # Kitchen_Pillars_Condensed_v2.docx
node build_detail.js       # Kitchen_Pillars_Detail_Dollars.docx
```

Or all at once:
```bash
npm run build-all
```

## How it works

1. `lib.js` — shared helpers: data loader, formatters (fmtMoney, fmtShares, fmtPct), projection math, position value/income calculators
2. Each build script pulls `data_clean.json` (generated from the live GitHub index.html), computes all figures live, and writes a validated .docx

## Weekly refresh workflow

1. Upload updated dashboard to GitHub
2. In a new Claude session: "refresh the docs" — Claude pulls fresh data, runs these scripts, patches the other 4 docs
3. Build scripts are at: `bradk-md/household-portfolio/docs/`

## Documents NOT covered by scripts (patched manually each refresh)
- `Kitchen_Complete_Reference.docx` — contains personal info/contacts/narrative; structure preserved, figures patched
- `Kitchen_Action_Items.docx` — historical log; Income Summary section patched, new trades logged
- `Kitchen_Advisor_Presentation.html` — per-position table patched from live data
- `Kitchen_ThreePillars.html` — summary tiles/projections patched from live data

## Data file
`data_clean.json` is regenerated each session from the live GitHub dashboard — never commit this file, it will always be stale by the next session.
