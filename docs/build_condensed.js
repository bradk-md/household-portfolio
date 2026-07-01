const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, AlignmentType,
        WidthType, ShadingType, BorderStyle, VerticalAlign } = require('docx');
const fs = require('fs');
const { loadData, computeSummary, fmtMoney, fmtShares, fmtPct, grow, posVal, posInc, posYield } = require('./lib.js');
const { computeAllInIncome } = require('./lib.js');

const data = loadData();
const s = computeSummary(data);
const ai = computeAllInIncome(data);
const allInInc = s.combinedInc + ai.p2p3Inc;
const allInMo = allInInc / 12;

const NAVY = '1F3864', TEAL = '0D6B52', PURPLE = '534AB7', BLUE = '185FA5', GRAY = '5A5E6B', LIGHT = 'F5F5F3', GOLD='B8960C';
const border = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
const borders = { top: border, bottom: border, left: border, right: border };

function cell(text, opts = {}) {
  const { bold = false, color = '000000', fill = null, width = 1500, align = AlignmentType.LEFT, size = 17, italics=false } = opts;
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: fill ? { fill, type: ShadingType.CLEAR } : undefined,
    margins: { top: 60, bottom: 60, left: 110, right: 110 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: align,
      children: [new TextRun({ text: String(text), bold, color, size, font: 'Calibri', italics })]
    })]
  });
}
function headerRow(cells, widths, fill = 'EEEEEE') {
  return new TableRow({ children: cells.map((t, i) => cell(t, { bold: true, fill, width: widths[i], size: 16 })) });
}
function dataRow(cells, widths, opts = []) {
  return new TableRow({ children: cells.map((t, i) => cell(t, { width: widths[i], ...(opts[i] || {}) })) });
}
function h1(text, color=NAVY) {
  return new Paragraph({ children: [new TextRun({ text, bold: true, size: 26, color, font: 'Calibri' })], spacing: { before: 240, after: 120 } });
}
function h2(text, color=NAVY, size=20) {
  return new Paragraph({ children: [new TextRun({ text, bold: true, size, color, font: 'Calibri' })], spacing: { before: 160, after: 100 } });
}
function spacer() {
  return new Paragraph({ children: [new TextRun({ text: '', size: 8 })], spacing: { after: 80 } });
}
function table(widths, rows) {
  return new Table({ width: { size: widths.reduce((a,b)=>a+b,0), type: WidthType.DXA }, columnWidths: widths, rows });
}

const children = [];

// Title
children.push(new Paragraph({ children: [new TextRun({ text: 'Brad & Lisa Kitchen — Three Pillars', bold: true, size: 32, color: NAVY, font: 'Calibri' })], spacing: { after: 60 } }));
children.push(new Paragraph({ children: [new TextRun({ text: `Condensed portfolio summary · ${s.asOf} · Dividend rates verified against official sources`, size: 18, color: GRAY, font: 'Calibri' })], spacing: { after: 240 } }));

// Top summary table (4 pillar tiles as a table)
const sumWidths = [2340, 2340, 2340, 2340];
const sumRows = [
  headerRow(['Pillar 1 — Income', 'Pillar 2 — Growth', 'Pillar 3 — Health/LTC', 'Grand Total'], sumWidths, 'EEEEEE'),
  dataRow([
    fmtMoney(s.p1Val), fmtMoney(s.p2Val), fmtMoney(s.p3Val), fmtMoney(s.grandTotal)
  ], sumWidths, [{bold:true,color:TEAL,size:20},{bold:true,color:PURPLE,size:20},{bold:true,color:BLUE,size:20},{bold:true,color:NAVY,size:20}]),
  dataRow([
    `${fmtMoney(s.combinedInc)}/yr · ${fmtMoney(s.combinedMo)}/mo`, 'Brad 401k + Lisa Roth', 'Pure growth · 8 positions', `+${fmtMoney(s.capOne.balance)} Capital One`
  ], sumWidths, [{size:15,color:GRAY},{size:15,color:GRAY},{size:15,color:GRAY},{size:15,color:GRAY}]),
];
children.push(table(sumWidths, sumRows));

// ===== PILLAR 1 =====
const p1Positions = data.holdings.filter(p => p.pillar === 1);
const totalP1Count = p1Positions.filter(p => !['Cash','SWVXX','FDRXX'].includes(p.ticker)).length;
children.push(h1('Pillar 1 Income Portfolio', TEAL));
children.push(new Paragraph({ children: [new TextRun({ text: `${fmtMoney(s.combinedInc)}/yr · ${fmtMoney(s.combinedMo)}/mo · ${totalP1Count} positions across 6 account types · rates verified`, size: 17, color: GRAY, font: 'Calibri' })], spacing: { after: 120 } }));
children.push(new Paragraph({ children: [new TextRun({ text: `Total Asset Value `, size: 18, bold:true, font:'Calibri' }), new TextRun({ text: fmtMoney(s.p1Val), size: 22, bold:true, color: TEAL, font:'Calibri' })], spacing: { after: 120 } }));

const posWidths = [1400, 1300, 1100, 1500, 900, 3160];
const acctOrder = ['Brad IRA', 'Lisa IRA', 'Brad Roth IRA', 'Lisa HSA', 'Brad HSA', 'Joint'];
const acctDisplay = { 'Brad IRA': 'Brad IRA', 'Lisa IRA': 'Lisa IRA', 'Brad Roth IRA': 'Brad Roth', 'Lisa HSA': 'Lisa HSA', 'Brad HSA': 'Brad HSA', 'Joint': 'Joint' };
const skipTickers = ['Cash', 'SWVXX', 'FDRXX'];
const notesMap = {
  'O': 'Net lease REIT — monthly payer', 'PTY': 'PIMCO corp credit CEF', 'ET': 'Midstream energy MLP',
  'ARCC': 'BDC — middle market lending', 'SPG': 'Mall REIT — premier', 'IRM': 'Data/storage REIT',
  'MAIN': 'BDC — LMM equity+debt', 'OKE': 'Midstream pipeline', 'DLR': 'Data center REIT',
  'LADR': 'Commercial mortgage REIT', 'STAG': 'Industrial REIT — quarterly', 'IIPR': 'Cannabis REIT',
  'PBA': 'Canadian pipeline',
  'UTF': 'Infrastructure CEF', 'BME': 'Healthcare sciences CEF', 'SCHD': 'Dividend equity ETF · +563sh remaining pending',
  'PDO': 'PIMCO dynamic multi-sector', 'WPC': 'Net lease REIT', 'VICI': 'Gaming/hospitality REIT',
  'DFP': 'Preferred securities CEF', 'PLD': 'Industrial REIT', 'SCHZ': 'US aggregate bond ETF',
  'BLOK': 'Blockchain ETF', 'DIS': 'Walt Disney — semi-annual', 'NLOP': 'Special divs only — not forecast',
  'PFFA': 'Preferred infra ETF', 'BIZD': 'BDC basket ETF', 'GOF': 'Guggenheim strategic opps',
  'THQ': 'Healthcare opps CEF', 'AMLP': 'MLP ETF', 'AGNC': 'Agency mREIT',
  'BSTZ': 'BlackRock sci/tech CEF', 'JPC': 'Nuveen preferred & income', 'JRI': 'Nuveen real asset income',
  'ABR': 'Commercial mortgage REIT', 'JEPI': 'JPMorgan equity premium', 'BTO': 'Financial sector CEF',
  'UTG': 'Utility income CEF', 'BCX': 'Resources & commodities CEF',
  'GHY': 'PGIM global high yield', 'JEPQ': 'JPMorgan Nasdaq premium', 'PDI': 'PIMCO dynamic income',
  'EPD': 'Enterprise Products Partners (K-1)', 'MDXBX': 'T.Rowe MD Tax-Free Bond Fund',
  'ORCL': "Oracle Corp — Brad's employer", 'TPL': 'Texas Pacific Land',
};

for (const acct of acctOrder) {
  let positions = data.holdings.filter(p => p.acct === acct && p.pillar === 1 && !skipTickers.includes(p.ticker));
  positions = positions.sort((a,b) => a.ticker.localeCompare(b.ticker));
  if (positions.length === 0) continue;
  const rows = [headerRow(['Account', 'Position', 'Shares', 'Annual', 'Yield', 'Notes'], posWidths)];
  let acctInc = 0, acctVal = 0;
  positions.forEach((p, i) => {
    const val = posVal(p), inc = posInc(p), yld = posYield(p);
    acctInc += inc; acctVal += val;
    rows.push(dataRow([
      i === 0 ? acctDisplay[acct] : '',
      p.ticker, fmtShares(p.shares) + 'sh', `${fmtMoney(inc)}/yr`, fmtPct(yld), notesMap[p.ticker] || p.notes || ''
    ], posWidths, [{bold:true}]));
  });
  rows.push(dataRow([
    '', `${acctDisplay[acct]} Subtotal`, `${positions.length} positions`, `${fmtMoney(acctInc)}/yr`, '', `${fmtMoney(acctVal)} asset value`
  ], posWidths, [{},{bold:true,fill:LIGHT},{bold:true,fill:LIGHT},{bold:true,fill:LIGHT},{fill:LIGHT},{fill:LIGHT}]));
  children.push(table(posWidths, rows));
  children.push(spacer());
}

children.push(new Paragraph({
  children: [
    new TextRun({ text: `INCOME TOTAL  `, bold: true, size: 19, color: NAVY, font: 'Calibri' }),
    new TextRun({ text: `${totalP1Count} positions · ${fmtMoney(s.p1Inc)}/yr · ${fmtMoney(s.p1Inc/12)}/mo · ${fmtMoney(s.p1Val)} total assets · rates verified ${s.asOf}`, size: 17, font: 'Calibri' }),
  ], spacing: { before: 80, after: 60 }
}));
children.push(new Paragraph({
  children: [
    new TextRun({ text: `TOTAL ALL-IN INCOME  `, bold: true, size: 19, color: TEAL, font: 'Calibri' }),
    new TextRun({ text: `${fmtMoney(allInInc)}/yr · ${fmtMoney(allInMo)}/mo`, bold: true, size: 19, color: TEAL, font: 'Calibri' }),
    new TextRun({ text: `  (adds ~${fmtMoney(ai.p2p3Inc)}/yr P2/P3 dividends harvested → redeployed into Pillar 1)`, size: 15, color: GRAY, font: 'Calibri' }),
  ], spacing: { after: 60 }
}));
children.push(new Paragraph({
  children: [new TextRun({ text: `Pending trades: SCHD +425sh remaining · THW ~1,050sh (Lisa IRA) · BCX 250sh (Lisa IRA) · THQ (Lisa IRA — wait for premium to narrow)`, size: 15, italics: true, color: GRAY, font: 'Calibri' })],
  spacing: { after: 60 }
}));
children.push(new Paragraph({
  children: [new TextRun({ text: `Pillar 3 note: Growth projections use total return rates. Since P3 dividends (~${fmtMoney(ai.p3Inc)}/yr) are redirected to income rather than reinvested, actual compounding tracks closer to the below-average (6.6%) scenario by design.`, size: 15, italics: true, color: GRAY, font: 'Calibri' })],
  spacing: { after: 200 }
}));

// ===== PILLAR 2 =====
children.push(h1('Pillar 2 Growth — 401k', PURPLE));
children.push(new Paragraph({ children: [new TextRun({ text: 'Brad 401k · Growth Pool + Vanguard Total International · pure growth compounding', size: 17, color: GRAY, font: 'Calibri' })], spacing: { after: 80 } }));
children.push(new Paragraph({ children: [new TextRun({ text: `Total Asset Value `, size: 18, bold:true, font:'Calibri' }), new TextRun({ text: fmtMoney(s.p2Val), size: 22, bold:true, color: PURPLE, font:'Calibri' })], spacing: { after: 120 } }));

const p2Widths = [2300, 1500, 1600, 1900, 1960];
const p2Rows = [headerRow(['Fund', 'Units', 'Value', '7% Yr 10', '9% Yr 10'], p2Widths)];
const brad401k = data.holdings.filter(p => p.acct === 'Brad 401k');
let p2Total = 0;
const p2Names = { 'GPool': 'Growth Pool Fund', 'VGIntl': 'Vanguard Total Intl' };
for (const p of brad401k) {
  const val = posVal(p);
  p2Total += val;
  p2Rows.push(dataRow([p2Names[p.ticker] || p.name, `${fmtShares(p.shares)}u`, fmtMoney(val), fmtMoney(grow(val,0.07,10)), fmtMoney(grow(val,0.09,10))], p2Widths));
}
p2Rows.push(dataRow(['401K TOTAL', '', fmtMoney(p2Total), fmtMoney(grow(p2Total,0.07,10)), fmtMoney(grow(p2Total,0.09,10))], p2Widths, [{bold:true,fill:LIGHT},{fill:LIGHT},{bold:true,fill:LIGHT},{bold:true,fill:LIGHT},{bold:true,fill:LIGHT}]));
children.push(table(p2Widths, p2Rows));
children.push(new Paragraph({ children: [new TextRun({ text: `No withdrawals assumed. 7% = average · 9% = above average. Brad contributing ~$28K/yr. Total portfolio projects to ~${fmtMoney(s.investTotal+grow(s.p2Val,0.07,10)-s.p2Val,0)} by Year 10.`, size: 15, italics: true, color: GRAY, font: 'Calibri' })], spacing: { before: 100, after: 200 } }));

// ===== PILLAR 3 =====
children.push(h1('Pillar 3 Health / LTC Self-Insurance', BLUE));
children.push(new Paragraph({ children: [new TextRun({ text: 'Pure growth · no income counted · both average LTC scenarios covered TODAY', size: 17, color: GRAY, font: 'Calibri' })], spacing: { after: 80 } }));
children.push(new Paragraph({ children: [new TextRun({ text: `Total Asset Value `, size: 18, bold:true, font:'Calibri' }), new TextRun({ text: fmtMoney(s.p3Val), size: 22, bold:true, color: BLUE, font:'Calibri' })], spacing: { after: 120 } }));

const p3Widths = [1100, 1500, 1300, 1300, 900, 2960];
const p3Rows = [headerRow(['Position', 'Account', 'Shares', 'Value', '%', 'Role'], p3Widths)];
const p3RoleMap = {
  'FDGRX': 'US large growth — active', 'VUG': 'US large growth — passive', 'VTV': 'US large value — passive',
  'SPDW': 'Intl developed large/mid', 'SCHM': 'US mid cap — passive',
  'SCHA': 'US small cap', 'SCHC': 'Intl small cap — 50% sell pending', 'SCHE': 'Emerging markets — 50% sell after Jun 24',
};
const p3Holdings = data.holdings.filter(p => p.pillar === 3);
for (const p of p3Holdings) {
  const val = posVal(p);
  p3Rows.push(dataRow([p.ticker, p.acct, fmtShares(p.shares) + (p.ticker==='FDGRX'?'u':'sh'), fmtMoney(val), fmtPct(val/s.p3Val*100), p3RoleMap[p.ticker] || ''], p3Widths));
}
p3Rows.push(dataRow(['LTC TOTAL', '2 accounts', '', fmtMoney(s.p3Val), '100%', '24% large · 36% sm/mid · 30% intl · 9% EM'], p3Widths, [{bold:true,fill:LIGHT},{fill:LIGHT},{fill:LIGHT},{bold:true,fill:LIGHT},{bold:true,fill:LIGHT},{fill:LIGHT}]));
children.push(table(p3Widths, p3Rows));
children.push(spacer());

const p3pWidths = [1560, 2200, 2200, 2200, 1200];
const p3pRows = [headerRow(['Year', 'Below avg (6.6%)', 'Average (8.5%)', 'Above avg (10.7%)', ''], p3pWidths)];
for (const y of [0,5,10,15]) {
  const below = grow(s.p3Val, 0.066, y), avg = grow(s.p3Val, 0.085, y), above = grow(s.p3Val, 0.107, y);
  p3pRows.push(dataRow([y===0?'Today':`Year ${y}`, fmtMoney(below), fmtMoney(avg), fmtMoney(above), y>=10?'✓':''], p3pWidths));
}
children.push(table(p3pWidths, p3pRows));
children.push(new Paragraph({ children: [new TextRun({ text: `Below avg = difficult decade · Average = long-run norm · Above avg = strong decade · Allocation: 24% large cap · 36% small/mid · 30% international · 9% emerging markets`, size: 15, italics: true, color: GRAY, font: 'Calibri' })], spacing: { before: 100, after: 80 } }));
children.push(new Paragraph({ children: [new TextRun({ text: `Pending: SCHC 50% sell anytime · SCHE 50% sell after Jun 24 · Insurance cancellations free $7,575/yr when complete`, size: 15, italics: true, color: GRAY, font: 'Calibri' })] }));

const doc = new Document({
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 } } },
    children
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('/home/claude/rebuild/Kitchen_Pillars_Condensed_v2.docx', buffer);
  console.log('Written successfully');
});
