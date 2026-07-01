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

const NAVY = '1F3864', TEAL = '0D6B52', PURPLE = '534AB7', BLUE = '185FA5', GRAY = '5A5E6B', LIGHT = 'F5F5F3';
const border = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
const borders = { top: border, bottom: border, left: border, right: border };

function cell(text, opts = {}) {
  const { bold = false, color = '000000', fill = null, width = 1500, align = AlignmentType.LEFT, size = 17, italics=false } = opts;
  return new TableCell({
    borders, width: { size: width, type: WidthType.DXA },
    shading: fill ? { fill, type: ShadingType.CLEAR } : undefined,
    margins: { top: 60, bottom: 60, left: 110, right: 110 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ alignment: align, children: [new TextRun({ text: String(text), bold, color, size, font: 'Calibri', italics })] })]
  });
}
function headerRow(cells, widths, fill='EEEEEE') { return new TableRow({ children: cells.map((t,i)=>cell(t,{bold:true,fill,width:widths[i],size:16})) }); }
function dataRow(cells, widths, opts=[]) { return new TableRow({ children: cells.map((t,i)=>cell(t,{width:widths[i],...(opts[i]||{})})) }); }
function h1(text, color=NAVY) { return new Paragraph({ children: [new TextRun({ text, bold: true, size: 26, color, font: 'Calibri' })], spacing: { before: 220, after: 100 } }); }
function spacer() { return new Paragraph({ children: [new TextRun({ text: '', size: 8 })], spacing: { after: 80 } }); }
function table(widths, rows) { return new Table({ width: { size: widths.reduce((a,b)=>a+b,0), type: WidthType.DXA }, columnWidths: widths, rows }); }

const children = [];
children.push(new Paragraph({ children: [new TextRun({ text: 'Brad & Lisa Kitchen', bold: true, size: 32, color: NAVY, font: 'Calibri' })], spacing: { after: 60 } }));
children.push(new Paragraph({ children: [new TextRun({ text: `Three Pillars — Complete Holdings · ${s.asOf} · Confidential · For Brad & Lisa`, size: 18, color: GRAY, font: 'Calibri' })], spacing: { after: 200 } }));

const skipTickers = ['Cash','SWVXX','FDRXX'];
children.push(h1('Pillar 1 Income Portfolio', TEAL));
children.push(new Paragraph({ children: [new TextRun({ text: `${fmtMoney(s.p1Inc)}/yr · ${fmtMoney(s.p1Inc/12)}/mo · 6 accounts · ${fmtMoney(s.p1Val)} total assets`, size: 17, color: GRAY, font: 'Calibri' })], spacing: { after: 120 } }));

const notesMap = {
  'O': 'Net lease REIT — monthly payer', 'PTY': 'PIMCO credit CEF — monthly', 'ET': 'Midstream energy MLP',
  'ARCC': 'BDC — middle market lending', 'SPG': 'Mall REIT — premier', 'IRM': 'Data/document storage REIT',
  'MAIN': 'BDC — LMM equity+debt', 'OKE': 'Natural gas pipeline', 'DLR': 'Data center REIT',
  'LADR': 'Commercial mortgage REIT', 'STAG': 'Industrial REIT — quarterly', 'IIPR': 'Cannabis REIT',
  'PBA': 'Canadian pipeline',
  'UTF': 'Infrastructure CEF — monthly', 'BME': 'Healthcare sciences CEF', 'SCHD': '+563sh remaining pending — additional income when complete',
  'PDO': 'PIMCO multi-sector CEF — monthly', 'WPC': 'Net lease REIT — diversified', 'VICI': 'Gaming & hospitality REIT',
  'DFP': 'Preferred securities CEF', 'PLD': 'Industrial REIT', 'SCHZ': 'Bond ETF — monthly',
  'BLOK': 'Blockchain ETF — annual Dec dist', 'DIS': 'Semi-annual payer', 'NLOP': 'Special distributions only',
  'PFFA': 'Infrastructure preferred ETF', 'BIZD': 'BDC basket ETF', 'GOF': 'Multi-sector CEF — monthly',
  'THQ': 'Healthcare opps CEF', 'AMLP': 'MLP ETF — no K-1', 'AGNC': 'Agency mREIT — monthly',
  'BSTZ': 'Science & tech CEF — monthly', 'JPC': 'Preferred securities CEF', 'JRI': 'Real asset income CEF',
  'ABR': 'Commercial mortgage REIT', 'JEPI': 'Covered call income ETF', 'BTO': 'Financial sector CEF',
  'UTG': 'Utility income CEF — monthly', 'BCX': 'Resources & commodities CEF',
  'GHY': 'Global high yield', 'JEPQ': 'Covered call Nasdaq ETF', 'PDI': 'PIMCO flagship — monthly',
  'EPD': 'Midstream MLP (K-1) — correctly placed in taxable', 'MDXBX': 'Maryland muni bond — tax-exempt interest',
  'ORCL': "$0.50/qtr · Brad is Oracle employee", 'TPL': '$0.60/qtr',
};

const acctOrder = [
  ['Brad IRA', 'Brad IRA — Traditional', 'Traditional IRA · taxable on withdrawal · Schwab'],
  ['Lisa IRA', 'Lisa IRA — Traditional', 'Traditional IRA · taxable on withdrawal · Schwab'],
  ['Brad Roth IRA', 'Brad Roth IRA — TAX FREE', 'Roth IRA · ALL income tax-free forever · Fidelity'],
  ['Lisa HSA', 'Lisa HSA — Triple Tax-Free', 'HSA · Triple tax-free · Fidelity'],
  ['Brad HSA', 'Brad HSA — Triple Tax-Free', 'HSA · Triple tax-free · Fidelity'],
  ['Joint', 'Joint Brokerage — Taxable', 'Joint account · long-term capital gains rates · Schwab · trust structure in place'],
];

const dWidths = [2700, 1400, 1500, 1100, 2660];
for (const [acct, label, footNote] of acctOrder) {
  let positions = data.holdings.filter(p => p.acct === acct && p.pillar === 1 && !skipTickers.includes(p.ticker));
  if (positions.length === 0) continue;
  children.push(new Paragraph({ children: [new TextRun({ text: label, bold: true, size: 19, color: NAVY, font: 'Calibri' })], spacing: { before: 160, after: 60 } }));
  const rows = [headerRow(['Position', 'Shares', 'Annual Income', 'Yield', 'Notes'], dWidths)];
  let acctInc = 0, acctVal = 0;
  for (const p of positions) {
    const val = posVal(p), inc = posInc(p), yld = posYield(p);
    acctInc += inc; acctVal += val;
    rows.push(dataRow([`${p.ticker} — ${p.name}`, fmtShares(p.shares)+'sh', `${fmtMoney(inc)}/yr`, fmtPct(yld), notesMap[p.ticker]||p.notes||''], dWidths));
  }
  const yld = acctVal ? acctInc/acctVal*100 : 0;
  rows.push(dataRow([`${label} Total`, '', `${fmtMoney(acctInc)}/yr`, fmtPct(yld), footNote], dWidths, [{bold:true,fill:LIGHT},{fill:LIGHT},{bold:true,fill:LIGHT},{bold:true,fill:LIGHT},{fill:LIGHT,size:15,italics:true}]));
  children.push(table(dWidths, rows));
}

children.push(spacer());
const totWidths = [3000, 1700, 1700, 3960];
const totRows = [
  headerRow(['INCOME TOTAL', 'Annual', 'Monthly', '% of Target / Notes'], totWidths),
  dataRow(['Portfolio income (all accounts)', `${fmtMoney(s.p1Inc)}/yr`, `${fmtMoney(s.p1Inc/12)}/mo`, `${fmtPct(s.p1Inc/84000*100)} of $84K target`], totWidths),
  dataRow(['Capital One interest (3.20%)', `${fmtMoney(s.capOneInc)}/yr`, `${fmtMoney(s.capOneInc/12)}/mo`, 'Sacred principal · interest reinvested in joint'], totWidths),
  dataRow(['Combined Income (P1 + Cap One)', `${fmtMoney(s.combinedInc)}/yr`, `${fmtMoney(s.combinedMo)}/mo`, `${fmtPct(s.combinedInc/84000*100)} of $84K target`], totWidths),
  dataRow(['P2/P3 dividends (redirected to income)', `+${fmtMoney(ai.p2p3Inc)}/yr`, `+${fmtMoney(ai.p2p3Inc/12)}/mo`, 'Harvested to cash · redeployed into Pillar 1 positions'], totWidths),
  dataRow(['TOTAL ALL-IN INCOME', `${fmtMoney(allInInc)}/yr`, `${fmtMoney(allInMo)}/mo`, 'All household investment income · P1+P2+P3+CapOne'], totWidths, [{bold:true,fill:'E6F4EF',color:TEAL},{bold:true,fill:'E6F4EF'},{bold:true,fill:'E6F4EF'},{fill:'E6F4EF'}]),
];
children.push(table(totWidths, totRows));

children.push(new Paragraph({ children: [new TextRun({ text: `Income note: "Combined Income" = Pillar 1 + Capital One. "Total All-In Income" (${fmtMoney(allInInc)}/yr · ${fmtMoney(allInMo)}/mo) adds ~${fmtMoney(ai.p2p3Inc)}/yr in Pillar 2/3 dividends harvested to cash and redeployed into Pillar 1. Pillar 3 projections use total return rates but actual compounding tracks closer to below-average (6.6%) since dividends are redirected out by design.`, size: 15, color: GRAY, italics: true, font: 'Calibri' })], spacing: { before: 80, after: 200 } }));
children.push(h1('Pillar 2 Growth — 401k', PURPLE));
children.push(new Paragraph({ children: [new TextRun({ text: `Brad 401k · Fidelity · pure growth · no distributions · ${fmtMoney(s.p2Val)}`, size: 17, color: GRAY, font: 'Calibri' })], spacing: { after: 100 } }));
children.push(new Paragraph({ children: [new TextRun({ text: `Total Value `, bold:true, size:18, font:'Calibri' }), new TextRun({ text: fmtMoney(s.p2Val), bold:true, size:22, color:PURPLE, font:'Calibri' })], spacing: { after: 120 } }));

const p2Widths = [2400, 1500, 1700, 1700, 1960];
const p2Rows = [headerRow(['Fund', 'Units', 'Price', 'Value', '% of 401k'], p2Widths)];
const brad401k = data.holdings.filter(p => p.acct === 'Brad 401k');
const p2Names = { 'GPool': 'Growth Pool Fund', 'VGIntl': 'Vanguard Total Intl' };
let p2Total = 0;
for (const p of brad401k) { p2Total += posVal(p); }
for (const p of brad401k) {
  const val = posVal(p);
  p2Rows.push(dataRow([p2Names[p.ticker]||p.name, `${fmtShares(p.shares)}u`, fmtMoney(p.price,2), fmtMoney(val), fmtPct(val/p2Total*100)], p2Widths));
}
p2Rows.push(dataRow(['401K TOTAL', '', '', fmtMoney(p2Total), '100%'], p2Widths, [{bold:true,fill:LIGHT},{fill:LIGHT},{fill:LIGHT},{bold:true,fill:LIGHT},{bold:true,fill:LIGHT}]));
children.push(table(p2Widths, p2Rows));
children.push(new Paragraph({ children: [new TextRun({ text: `Brad contributing ~$28K/yr pre-tax.`, size: 15, italics: true, color: GRAY, font: 'Calibri' })], spacing: { before: 80, after: 100 } }));

children.push(new Paragraph({ children: [new TextRun({ text: '401k growth projections', bold: true, size: 19, color: NAVY, font: 'Calibri' })], spacing: { before: 100, after: 60 } }));
const gWidths = [2500, 1715, 1715, 1715, 1715];
const gRows = [headerRow(['Scenario', 'Now', 'Year 5', 'Year 10', 'Year 15'], gWidths)];
gRows.push(dataRow(['7% CAGR — average', fmtMoney(p2Total), fmtMoney(grow(p2Total,0.07,5)), fmtMoney(grow(p2Total,0.07,10)), fmtMoney(grow(p2Total,0.07,15))], gWidths));
gRows.push(dataRow(['9% CAGR — above average', fmtMoney(p2Total), fmtMoney(grow(p2Total,0.09,5)), fmtMoney(grow(p2Total,0.09,10)), fmtMoney(grow(p2Total,0.09,15))], gWidths));
children.push(table(gWidths, gRows));
children.push(new Paragraph({ children: [new TextRun({ text: `No withdrawals planned until RMDs at age 73 (Brad: 2035). Brad contributing ~$28,000/yr pre-tax. Combined with Pillar 1 income and Pillar 3 LTC, total investable portfolio currently stands at ${fmtMoney(s.investTotal,0)}.`, size: 15, italics: true, color: GRAY, font: 'Calibri' })], spacing: { before: 80, after: 200 } }));

children.push(h1('Pillar 3 Health / LTC Self-Insurance', BLUE));
children.push(new Paragraph({ children: [new TextRun({ text: 'Pure growth · no income counted · both avg LTC scenarios covered TODAY', size: 17, color: GRAY, font: 'Calibri' })], spacing: { after: 80 } }));
children.push(new Paragraph({ children: [new TextRun({ text: `Total Value `, bold:true, size:18, font:'Calibri' }), new TextRun({ text: fmtMoney(s.p3Val), bold:true, size:22, color:BLUE, font:'Calibri' })], spacing: { after: 60 } }));
children.push(new Paragraph({ children: [new TextRun({ text: 'Virtual account — positions held inside Brad IRA and Lisa IRA but designated as the self-funded LTC reserve. No income is counted from any of these positions. All growth compounds untouched.', size: 15, color: GRAY, font: 'Calibri' })], spacing: { after: 120 } }));

const p3RoleMap = {
  'FDGRX': 'US large cap growth — active managed', 'VUG': 'US large cap growth — passive · 0.03% fee',
  'VTV': 'US large cap value — passive · 0.03% fee', 'SPDW': 'International developed large/mid · 46 countries',
  'SCHM': 'US mid cap passive', 'SCHA': 'US small cap — passive', 'SCHC': 'Intl small cap developed — FULLY KEPT',
  'SCHE': 'Emerging markets — 50% sell after Jun 24 ex-div',
};
const ltcWidths = [3000, 1300, 1600, 1300, 2160];
for (const [acctKey, acctLabel] of [['Brad IRA (LTC)', 'Brad IRA'], ['Lisa IRA (LTC)', 'Lisa IRA']]) {
  const positions = data.holdings.filter(p => p.pillar===3 && p.acct === acctKey);
  if (positions.length===0) continue;
  children.push(new Paragraph({ children: [new TextRun({ text: `${acctLabel} — LTC Positions`, bold: true, size: 18, color: NAVY, font: 'Calibri' })], spacing: { before: 120, after: 60 } }));
  const rows = [headerRow(['Position', 'Shares', 'Value', '% of LTC', 'Role'], ltcWidths)];
  let subVal = 0;
  for (const p of positions) {
    const val = posVal(p); subVal += val;
    rows.push(dataRow([`${p.ticker} — ${p.name}`, fmtShares(p.shares)+(p.ticker==='FDGRX'?'u':'sh'), fmtMoney(val), fmtPct(val/s.p3Val*100), p3RoleMap[p.ticker]||''], ltcWidths));
  }
  rows.push(dataRow([`${acctLabel} LTC Subtotal`, `${positions.length} positions`, fmtMoney(subVal), fmtPct(subVal/s.p3Val*100), `${acctLabel} · Schwab`], ltcWidths, [{bold:true,fill:LIGHT},{fill:LIGHT},{bold:true,fill:LIGHT},{bold:true,fill:LIGHT},{fill:LIGHT}]));
  children.push(table(ltcWidths, rows));
}

children.push(spacer());
const ltcTotWidths = [3300, 1700, 2160, 2200];
children.push(table(ltcTotWidths, [
  headerRow(['LTC TOTAL', 'Value', '% of LTC', 'Allocation'], ltcTotWidths),
  dataRow(['All 8 positions · Brad IRA + Lisa IRA', fmtMoney(s.p3Val), '100%', '24% large cap · 36% small/mid · 30% intl · 9% EM'], ltcTotWidths, [{bold:true,fill:LIGHT},{bold:true,fill:LIGHT},{bold:true,fill:LIGHT},{fill:LIGHT,size:15}]),
]));

children.push(new Paragraph({ children: [new TextRun({ text: 'LTC growth projections — three scenarios', bold: true, size: 19, color: NAVY, font: 'Calibri' })], spacing: { before: 160, after: 60 } }));
const lgWidths = [2400, 1600, 1600, 1600, 2160];
const lgRows = [headerRow(['Scenario', 'Now', 'Year 5', 'Year 10', 'Year 15 / Coverage'], lgWidths)];
lgRows.push(dataRow(['Below avg — 6.6%/yr', fmtMoney(s.p3Val), fmtMoney(grow(s.p3Val,0.066,5)), fmtMoney(grow(s.p3Val,0.066,10)), `${fmtMoney(grow(s.p3Val,0.066,15))} · Both extended ($800K) covered`], lgWidths));
lgRows.push(dataRow(['Average — 8.5%/yr', fmtMoney(s.p3Val), fmtMoney(grow(s.p3Val,0.085,5)), fmtMoney(grow(s.p3Val,0.085,10)), `${fmtMoney(grow(s.p3Val,0.085,15))} · Fully covered \u2713`], lgWidths));
lgRows.push(dataRow(['Above avg — 10.7%/yr', fmtMoney(s.p3Val), fmtMoney(grow(s.p3Val,0.107,5)), fmtMoney(grow(s.p3Val,0.107,10)), `${fmtMoney(grow(s.p3Val,0.107,15))} · Worst case covered`], lgWidths));
children.push(table(lgWidths, lgRows));
children.push(new Paragraph({ children: [new TextRun({ text: `LTC cost benchmarks: ~$100K/yr national median nursing home. One person average (2.5 yrs = $250K) covered TODAY. Both people average (4 yrs = $400K) covered by Year 2. Both people extended (8 yrs = $800K) covered at Year 10 average case.`, size: 15, italics: true, color: GRAY, font: 'Calibri' })], spacing: { before: 100, after: 80 } }));
children.push(new Paragraph({ children: [new TextRun({ text: `Pending: SCHE — sell 50% after June 24 ex-div and redeploy to income (+$2,136/yr). SCHC — FULLY KEPT in LTC (confirmed decision — provides unique intl small cap exposure).`, size: 15, italics: true, color: GRAY, font: 'Calibri' })], spacing: { after: 80 } }));
children.push(new Paragraph({ children: [new TextRun({ text: `${s.asOf} · Prices from Schwab & Fidelity · Dividend rates verified · Not financial advice`, size: 14, color: GRAY, font: 'Calibri' })] }));

const doc = new Document({
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 } } },
    children
  }]
});
Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('/home/claude/rebuild/Kitchen_Pillars_Detail_Dollars.docx', buffer);
  console.log('Written successfully');
});
