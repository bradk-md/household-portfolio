const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, AlignmentType,
        WidthType, ShadingType, BorderStyle, VerticalAlign } = require('docx');
const fs = require('fs');
const { loadData, computeSummary, fmtMoney, fmtPct, grow } = require('./lib.js');
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
  const { bold = false, color = '000000', fill = null, width = 1500, align = AlignmentType.LEFT, size = 18 } = opts;
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: fill ? { fill, type: ShadingType.CLEAR } : undefined,
    margins: { top: 70, bottom: 70, left: 130, right: 130 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: align,
      children: [new TextRun({ text: String(text), bold, color, size, font: 'Calibri' })]
    })]
  });
}

function headerRow(cells, widths, fill = 'EEEEEE') {
  return new TableRow({
    children: cells.map((t, i) => cell(t, { bold: true, fill, width: widths[i], size: 17 }))
  });
}

function dataRow(cells, widths, opts = []) {
  return new TableRow({
    children: cells.map((t, i) => cell(t, { width: widths[i], ...(opts[i] || {}) }))
  });
}

const children = [];

// Title
children.push(new Paragraph({
  children: [new TextRun({ text: 'Brad & Lisa Kitchen', bold: true, size: 32, color: NAVY, font: 'Calibri' })],
  spacing: { after: 60 }
}));
children.push(new Paragraph({
  children: [new TextRun({ text: `Investment Portfolio Summary · ${s.asOf} · Confidential · For Brad & Lisa`, size: 18, color: GRAY, font: 'Calibri' })],
  spacing: { after: 240 }
}));

// Investment Portfolio Overview table
children.push(new Paragraph({
  children: [new TextRun({ text: 'Investment Portfolio Overview', bold: true, size: 24, color: NAVY, font: 'Calibri' })],
  spacing: { before: 120, after: 120 }
}));

const ovWidths = [2800, 1600, 1600, 3360];
const ovRows = [
  headerRow(['Category', 'Value', '% of Total', 'Notes'], ovWidths),
  dataRow(['Pillar 1 — Income', fmtMoney(s.p1Val), fmtPct(s.p1Val/s.grandTotal*100), `${fmtMoney(s.combinedInc)}/yr · ${fmtMoney(s.combinedMo)}/mo`], ovWidths, [{color:TEAL,bold:true}]),
  dataRow(['Pillar 2 — Growth (401k)', fmtMoney(s.p2Val), fmtPct(s.p2Val/s.grandTotal*100), 'Pure growth — no distributions'], ovWidths, [{color:PURPLE,bold:true}]),
  dataRow(['Pillar 3 — Health/LTC', fmtMoney(s.p3Val), fmtPct(s.p3Val/s.grandTotal*100), 'Pure growth — self-insured LTC fund'], ovWidths, [{color:BLUE,bold:true}]),
  dataRow(['Lisa Roth IRA', fmtMoney(s.acctsP1['Lisa Roth IRA'] ? s.acctsP1['Lisa Roth IRA'].val : 0), fmtPct((s.acctsP1['Lisa Roth IRA']?s.acctsP1['Lisa Roth IRA'].val:0)/s.grandTotal*100), 'Building via annual Q4 conversions'], ovWidths),
  dataRow(['Investment Portfolio', fmtMoney(s.investTotal), fmtPct(s.investTotal/s.grandTotal*100), 'Schwab · Fidelity'], ovWidths, [{bold:true,fill:LIGHT},{bold:true,fill:LIGHT},{bold:true,fill:LIGHT},{fill:LIGHT}]),
  dataRow(['Capital One Savings', fmtMoney(s.capOne.balance), fmtPct(s.capOne.balance/s.grandTotal*100), 'Sacred emergency reserve — principal never touched'], ovWidths),
  dataRow(['Crypto (iTrust + Crypto.com)', '~' + fmtMoney(s.crypto), fmtPct(s.crypto/s.grandTotal*100), 'Speculative — approximate · market prices vary'], ovWidths),
  dataRow(['GRAND TOTAL', fmtMoney(s.grandTotal), '100%', s.asOf], ovWidths, [{bold:true,fill:'E6F4EF',color:NAVY},{bold:true,fill:'E6F4EF',color:NAVY},{bold:true,fill:'E6F4EF'},{fill:'E6F4EF'}]),
];
children.push(new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: ovWidths, rows: ovRows }));

// Pillar 1 Income by account
children.push(new Paragraph({ children: [new TextRun({ text: '', size: 8 })], spacing: { after: 200 } }));
children.push(new Paragraph({
  children: [new TextRun({ text: 'Pillar 1 — Income', bold: true, size: 24, color: TEAL, font: 'Calibri' })],
  spacing: { before: 120, after: 120 }
}));

const acctLabels = {
  'Brad IRA': 'Brad IRA — Traditional',
  'Lisa IRA': 'Lisa IRA — Traditional',
  'Brad Roth IRA': 'Brad Roth — TAX-FREE',
  'Lisa HSA': 'Lisa HSA — Triple Tax-Free',
  'Brad HSA': 'Brad HSA — Triple Tax-Free',
  'Joint': 'Joint Brokerage — Taxable',
};
const incWidths = [2900, 1700, 1700, 3060];
const incRows = [headerRow(['Account', 'Annual', 'Monthly', 'Notes'], incWidths)];
let acctTotalInc = 0, acctTotalMo = 0;
for (const [acct, label] of Object.entries(acctLabels)) {
  const a = s.acctsP1[acct] || { val: 0, inc: 0, count: 0 };
  const yieldPct = a.val ? (a.inc / a.val * 100) : 0;
  const positionsLabel = acct === 'Brad HSA' ? `${fmtPct(yieldPct)} yield · PDI` : `${fmtPct(yieldPct)} yield · ${a.count} position${a.count===1?'':'s'}`;
  incRows.push(dataRow([label, `${fmtMoney(a.inc)}/yr`, `${fmtMoney(a.inc/12)}/mo`, positionsLabel], incWidths));
  acctTotalInc += a.inc; acctTotalMo += a.inc/12;
}
incRows.push(dataRow(['Portfolio Income', `${fmtMoney(s.p1Inc)}/yr`, `${fmtMoney(s.p1Inc/12)}/mo`, '6 accounts · verified rates'], incWidths, [{bold:true,fill:LIGHT},{bold:true,fill:LIGHT},{bold:true,fill:LIGHT},{fill:LIGHT}]));
incRows.push(dataRow(['Capital One interest (3.20%)', `${fmtMoney(s.capOneInc)}/yr`, `${fmtMoney(s.capOneInc/12)}/mo`, 'Transfers to joint for reinvestment'], incWidths));
incRows.push(dataRow(['Combined Income (P1 + Cap One)', `${fmtMoney(s.combinedInc)}/yr`, `${fmtMoney(s.combinedMo)}/mo`, 'Portfolio income + Capital One interest'], incWidths));
incRows.push(dataRow(['P2/P3 dividends (redirected to income)', `+${fmtMoney(ai.p2p3Inc)}/yr`, `+${fmtMoney(ai.p2p3Inc/12)}/mo`, 'Harvested to cash · redeployed into Pillar 1 positions'], incWidths));
incRows.push(dataRow(['TOTAL ALL-IN INCOME', `${fmtMoney(allInInc)}/yr`, `${fmtMoney(allInMo)}/mo`, 'All household investment income sources combined'], incWidths, [{bold:true,fill:'E6F4EF',color:TEAL},{bold:true,fill:'E6F4EF',color:TEAL},{bold:true,fill:'E6F4EF'},{fill:'E6F4EF'}]));
children.push(new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: incWidths, rows: incRows }));

// Income trajectory
children.push(new Paragraph({ children: [new TextRun({ text: '', size: 8 })], spacing: { after: 200 } }));
children.push(new Paragraph({
  children: [new TextRun({ text: 'Income trajectory — 3% natural annual dividend growth', bold: true, size: 22, color: NAVY, font: 'Calibri' })],
  spacing: { before: 120, after: 120 }
}));
const trajWidths = [1400, 2200, 2200, 3560];
const trajRows = [headerRow(['Year', 'Annual Income', 'Monthly', 'Notes'], trajWidths)];
const trajNotes = {
  0: 'Confirmed current rates',
  3: '',
  5: '',
  10: '3% annual dividend growth assumed',
};
for (const y of [0, 3, 5, 10]) {
  const inc = grow(s.combinedInc, 0.03, y);
  trajRows.push(dataRow([y === 0 ? 'Now' : `Year ${y}`, `${fmtMoney(inc)}/yr`, `${fmtMoney(inc/12)}/mo`, trajNotes[y]], trajWidths));
}
children.push(new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: trajWidths, rows: trajRows }));

// Pillar 2 + 3 projections
children.push(new Paragraph({ children: [new TextRun({ text: '', size: 8 })], spacing: { after: 200 } }));
children.push(new Paragraph({
  children: [new TextRun({ text: 'Pillar 2 (Growth) + Pillar 3 (Health/LTC) — Projections', bold: true, size: 22, color: NAVY, font: 'Calibri' })],
  spacing: { before: 120, after: 120 }
}));
const pWidths = [2400, 1700, 1700, 1700, 1860];
const pRows = [headerRow(['', 'Now', 'Year 5', 'Year 10', 'Year 15'], pWidths)];
pRows.push(dataRow(['Pillar 2 — 401k (7% avg)', fmtMoney(s.p2Val), fmtMoney(grow(s.p2Val,0.07,5)), fmtMoney(grow(s.p2Val,0.07,10)), fmtMoney(grow(s.p2Val,0.07,15))], pWidths));
pRows.push(dataRow(['Pillar 2 — 401k (9% above avg)', fmtMoney(s.p2Val), fmtMoney(grow(s.p2Val,0.09,5)), fmtMoney(grow(s.p2Val,0.09,10)), fmtMoney(grow(s.p2Val,0.09,15))], pWidths));
pRows.push(dataRow(['Pillar 3 — LTC (6.6% below avg)', fmtMoney(s.p3Val), fmtMoney(grow(s.p3Val,0.066,5)), fmtMoney(grow(s.p3Val,0.066,10)), fmtMoney(grow(s.p3Val,0.066,15))], pWidths));
pRows.push(dataRow(['Pillar 3 — LTC (8.5% average)', fmtMoney(s.p3Val), fmtMoney(grow(s.p3Val,0.085,5)), fmtMoney(grow(s.p3Val,0.085,10)), fmtMoney(grow(s.p3Val,0.085,15))], pWidths));
pRows.push(dataRow(['Pillar 3 — LTC (10.7% above avg)', fmtMoney(s.p3Val), fmtMoney(grow(s.p3Val,0.107,5)), fmtMoney(grow(s.p3Val,0.107,10)), fmtMoney(grow(s.p3Val,0.107,15))], pWidths));
children.push(new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: pWidths, rows: pRows }));

// Combined trajectory
children.push(new Paragraph({ children: [new TextRun({ text: '', size: 8 })], spacing: { after: 200 } }));
children.push(new Paragraph({
  children: [new TextRun({ text: 'Combined Portfolio Trajectory — Average Case', bold: true, size: 22, color: NAVY, font: 'Calibri' })],
  spacing: { before: 120, after: 120 }
}));
const cWidths = [1600, 2100, 1900, 1900, 1860];
const cRows = [headerRow(['', 'Pillar 1 Assets', 'Pillar 2 (401k)', 'Pillar 3 (LTC)', 'Income/yr'], cWidths)];
cRows.push(dataRow(['Today', fmtMoney(s.investTotal), fmtMoney(s.p2Val), fmtMoney(s.p3Val), `${fmtMoney(s.combinedInc)}/yr`], cWidths));
cRows.push(dataRow(['Year 5', '~' + fmtMoney(s.p1Val,0), fmtMoney(grow(s.p2Val,0.07,5)), fmtMoney(grow(s.p3Val,0.085,5)), `${fmtMoney(grow(s.combinedInc,0.03,5))}/yr`], cWidths));
cRows.push(dataRow(['Year 10', '~' + fmtMoney(s.p1Val,0), fmtMoney(grow(s.p2Val,0.07,10)), fmtMoney(grow(s.p3Val,0.085,10)), `${fmtMoney(grow(s.combinedInc,0.03,10))}/yr`], cWidths));
children.push(new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: cWidths, rows: cRows }));

// Footnotes
children.push(new Paragraph({ children: [new TextRun({ text: '', size: 8 })], spacing: { after: 80 } }));
children.push(new Paragraph({ children: [new TextRun({ text: `Income note: "Combined Income" (${fmtMoney(s.combinedInc)}/yr) = Pillar 1 + Capital One. "Total All-In Income" (${fmtMoney(allInInc)}/yr) adds ~${fmtMoney(ai.p2p3Inc)}/yr in Pillar 2/3 dividends harvested to cash and redeployed into Pillar 1 income positions.`, size: 15, color: GRAY, italics: true, font: 'Calibri' })], spacing: { after: 60 } }));
children.push(new Paragraph({ children: [new TextRun({ text: `Pillar 3 projection note: Growth scenarios (6.6%/8.5%/10.7%) assume total return. Since Pillar 3 dividends (~${fmtMoney(ai.p3Inc)}/yr) are redirected to income rather than reinvested, actual Pillar 3 compounding tracks closer to the below-average (6.6%) scenario by design — P3 dividends accelerate Pillar 1 income instead.`, size: 15, color: GRAY, italics: true, font: 'Calibri' })], spacing: { after: 100 } }));
// Footer note
children.push(new Paragraph({ children: [new TextRun({ text: '', size: 8 })], spacing: { after: 200 } }));
children.push(new Paragraph({
  children: [new TextRun({
    text: `Capital One 3.20% APY · $${s.capOne.balance.toLocaleString()} principal sacred (never touched) · Interest transfers to joint for reinvestment · Crypto ~${fmtMoney(s.crypto)} speculative not included in pillar totals · Dividend rates verified ${s.asOf} · Not financial advice`,
    size: 15, color: GRAY, italics: true, font: 'Calibri'
  })]
}));

const doc = new Document({
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 } } },
    children
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('/home/claude/rebuild/Kitchen_Summary_Dollars.docx', buffer);
  console.log('Written successfully');
});
