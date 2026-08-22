// Stage 1: DATA SECTIONS of the Complete Reference, generated fresh from the dashboard.
// Produces per-account holdings tables (correct pillar assignments), pillar summaries,
// income summary, PIMCO suite, Pillar 2/3 holdings, crypto. NO hand-maintained content here.
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, AlignmentType,
        WidthType, ShadingType, BorderStyle, VerticalAlign } = require('docx');
const fs = require('fs');
const { loadData, computeSummary, computeAllInIncome, fmtMoney, fmtShares, fmtPct,
        posVal, posInc, posYield } = require('./lib.js');

const data = loadData();
const s = computeSummary(data);
const ai = computeAllInIncome(data);

const NAVY='1F3864', TEAL='0D6B52', PURPLE='534AB7', BLUE='185FA5', GRAY='5A5E6B', LIGHT='F5F5F3', GOLD='B8860B';
const border = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
const borders = { top: border, bottom: border, left: border, right: border };

function cell(text, opts = {}) {
  const { bold=false, color='000000', fill=null, width=1500, align=AlignmentType.LEFT, size=17, italics=false } = opts;
  return new TableCell({
    borders, width: { size: width, type: WidthType.DXA },
    shading: fill ? { fill, type: ShadingType.CLEAR } : undefined,
    margins: { top: 55, bottom: 55, left: 110, right: 110 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ alignment: align, children: [new TextRun({ text: String(text), bold, color, size, font: 'Calibri', italics })] })]
  });
}
const R = AlignmentType.RIGHT, L = AlignmentType.LEFT;
function headerRow(cells, widths, fill='EEEEEE') { return new TableRow({ children: cells.map((t,i)=>cell(t,{bold:true,fill,width:widths[i],size:16, align: i===0?L:R})) }); }
function row(cells, widths, opts=[]) { return new TableRow({ children: cells.map((t,i)=>cell(t,{width:widths[i], align:i===0?L:R, ...(opts[i]||{})})) }); }
function table(widths, rows) { return new Table({ width:{size:widths.reduce((a,b)=>a+b,0),type:WidthType.DXA}, columnWidths:widths, rows }); }
function h1(text, color=NAVY) { return new Paragraph({ children:[new TextRun({text,bold:true,size:26,color,font:'Calibri'})], spacing:{before:240,after:100} }); }
function h2(text, color=NAVY) { return new Paragraph({ children:[new TextRun({text,bold:true,size:21,color,font:'Calibri'})], spacing:{before:160,after:80} }); }
function sub(text) { return new Paragraph({ children:[new TextRun({text,size:17,color:GRAY,font:'Calibri'})], spacing:{after:120} }); }
function spacer(h=80) { return new Paragraph({ children:[new TextRun({text:'',size:8})], spacing:{after:h} }); }

const notesMap = {
  'O':'Net lease REIT — monthly','PTY':'PIMCO credit CEF — monthly','ET':'Midstream MLP (K-1) · UBTI watch','ARCC':'BDC — middle market',
  'SPG':'Mall REIT — premier','IRM':'Storage REIT','MAIN':'BDC — LMM','OKE':'Nat gas pipeline','DLR':'Data center REIT',
  'LADR':'Commercial mortgage REIT','STAG':'Industrial REIT — quarterly','IIPR':'Cannabis REIT','PBA':'Canadian pipeline (1099)',
  'UTF':'Infrastructure CEF — monthly','BME':'Healthcare CEF','SCHD':'Dividend ETF','PDO':'PIMCO multi-sector — monthly',
  'WPC':'Net lease REIT','VICI':'Gaming REIT','DFP':'Preferred CEF','PLD':'Industrial REIT','SCHZ':'Bond ETF — monthly',
  'BLOK':'Blockchain ETF — annual','DIS':'Semi-annual','NLOP':'Special distributions only','PFFA':'Preferred ETF',
  'BIZD':'BDC basket ETF','GOF':'Multi-sector CEF — monthly','THQ':'Healthcare CEF','AMLP':'MLP ETF — no K-1','AGNC':'Agency mREIT — monthly',
  'BSTZ':'Sci/tech CEF — monthly','JPC':'Preferred CEF','JRI':'Real asset CEF','ABR':'Commercial mortgage REIT','JEPI':'Covered call ETF',
  'BTO':'Financial CEF','UTG':'Utility CEF — monthly','BCX':'Resources CEF','GHY':'Global high yield CEF','JEPQ':'Covered call Nasdaq ETF',
  'THQ':'Healthcare opportunities CEF','THW':'Global healthcare CEF — monthly',
  'PDI':'PIMCO flagship — monthly','EPD':'Midstream MLP (K-1) — ROC distributions, tax-deferred','MDXBX':'MD muni bond — federally tax-exempt interest','ORCL':'Oracle — Brad employer','TPL':'Texas Pacific Land',
  'SPCX':'SPAC ETF','FDGRX':'Fidelity Growth','VUG':'Vanguard Growth','VTV':'Vanguard Value','SPDW':'Developed mkts','SCHM':'Mid-cap',
  'SCHA':'Small-cap','SCHE':'Emerging mkts','SCHC':'Intl small-cap','GPool':'401k stable value','VGIntl':'Vanguard Total Intl',
  'BREQI1':'401k real estate','WTSCER':'401k TIPS',
};

// Account display order and tax treatment labels
const P1_ACCTS = [
  ['Brad IRA','Traditional — tax-deferred'],
  ['Lisa IRA','Traditional — tax-deferred'],
  ['Brad Roth IRA','Roth — tax-free'],
  ['Lisa HSA','HSA — triple tax-free'],
  ['Brad HSA','HSA — triple tax-free'],
  ['Joint','Taxable account — but see per-holding tax treatment below'],
];

const children = [];

// ---- Title ----
children.push(new Paragraph({ children:[new TextRun({text:'Brad & Lisa Kitchen — Complete Reference',bold:true,size:32,color:NAVY,font:'Calibri'})], spacing:{after:60} }));
children.push(new Paragraph({ children:[new TextRun({text:`Portfolio data as of ${s.asOf} · Confidential · Dashboard is source of truth`,size:18,color:GRAY,font:'Calibri'})], spacing:{after:80} }));
children.push(new Paragraph({ children:[new TextRun({text:'This document is regenerated from the live dashboard on every refresh — the data sections below never go stale.',size:16,color:GRAY,italics:true,font:'Calibri'})], spacing:{after:160} }));

// ---- Investment Portfolio Summary ----
children.push(h1('Investment Portfolio Summary', NAVY));
const grandTotal = s.grandTotal;
const w1 = [3200, 2200, 2200, 1600];
children.push(table(w1, [
  headerRow(['Pillar / Category','Value','Income/yr','% of total'], w1),
  row(['Pillar 1 — Income', fmtMoney(s.p1Val), fmtMoney(s.p1Inc), fmtPct(s.p1Val/s.investTotal*100)], w1, [{bold:true,color:TEAL}]),
  row(['Pillar 2 — Growth (401k)', fmtMoney(s.p2Val), '—', fmtPct(s.p2Val/s.investTotal*100)], w1, [{bold:true,color:PURPLE}]),
  row(['Pillar 3 — Health/LTC', fmtMoney(s.p3Val), fmtMoney(ai.p3Inc)+' (swept→P1)', fmtPct(s.p3Val/s.investTotal*100)], w1, [{bold:true,color:BLUE}]),
  row(['Three Pillars Subtotal', fmtMoney(s.investTotal), fmtMoney(s.p1Inc), '100%'], w1, [{bold:true}]),
  row(['Cap One Savings', fmtMoney(s.capOne.balance), fmtMoney(s.capOneInc)+` (${fmtPct(s.capOne.apy*100)})`, '—'], w1),
  row(['Crypto (speculative)', fmtMoney(s.crypto), '—', '—'], w1),
  row(['Grand Total', fmtMoney(grandTotal), '', ''], w1, [{bold:true,fill:LIGHT,color:NAVY}]),
]));

// ---- Income Summary ----
children.push(h1('Income Summary', TEAL));
const allInInc = s.combinedInc + ai.p2p3Inc;
const w2 = [4400, 2400, 2400];
children.push(table(w2, [
  headerRow(['Income Source','Annual','Monthly'], w2),
  row(['Pillar 1 portfolio income', fmtMoney(s.p1Inc), fmtMoney(s.p1Inc/12)], w2),
  row(['Cap One interest', fmtMoney(s.capOneInc), fmtMoney(s.capOneInc/12)], w2),
  row(['Combined (Pillar 1 + Cap One)', fmtMoney(s.combinedInc), fmtMoney(s.combinedMo)], w2, [{bold:true,color:TEAL},{bold:true,color:TEAL},{bold:true,color:TEAL}]),
  row(['Pillar 2/3 divs (swept to P1)', fmtMoney(ai.p2p3Inc), fmtMoney(ai.p2p3Inc/12)], w2),
  row(['Total All-In Income', fmtMoney(allInInc), fmtMoney(allInInc/12)], w2, [{bold:true,fill:LIGHT},{bold:true,fill:LIGHT},{bold:true,fill:LIGHT}]),
]));

// ---- Pillar 1 per-account holdings tables ----
children.push(h1('Pillar 1 — Income Portfolio (by account)', TEAL));
const wA = [1500, 3400, 900, 1100, 1300, 1000];
for (const [acct, taxlabel] of P1_ACCTS) {
  const holds = data.holdings.filter(p => p.acct === acct && p.pillar === 1)
                             .sort((a,b)=>posInc(b)-posInc(a));
  if (!holds.length) continue;
  const acctVal = holds.reduce((t,p)=>t+posVal(p),0);
  const acctInc = holds.reduce((t,p)=>t+posInc(p),0);
  children.push(h2(`${acct} — ${fmtMoney(acctVal)} · ${fmtMoney(acctInc)}/yr`, NAVY));
  children.push(sub(taxlabel));
  const rows = [ headerRow(['Ticker','Name / Note','Shares','Price','Value','Inc/yr'], wA) ];
  for (const p of holds) {
    rows.push(row([p.ticker, notesMap[p.ticker]||p.name||'', fmtShares(p.shares), fmtMoney(p.price,2), fmtMoney(posVal(p)), fmtMoney(posInc(p))], wA));
  }
  rows.push(row([`${acct} Total`,'','', '', fmtMoney(acctVal), fmtMoney(acctInc)], wA, [{bold:true,fill:LIGHT},{},{},{},{bold:true,fill:LIGHT},{bold:true,fill:LIGHT}]));
  children.push(table(wA, rows));
  if (acct === 'Joint') {
    children.push(new Paragraph({ children:[new TextRun({ text:'Tax note: though this is a taxable account, MDXBX pays federally tax-exempt municipal-bond interest, and EPD (a K-1 MLP) distributes largely return-of-capital that is tax-deferred until sale. Only the ORCL dividends are ordinary taxable income here. This is why the Joint account contributes little to MAGI — see the retirement tax section below.', size:15, italics:true, color:GRAY, font:'Calibri' })], spacing:{ before:40, after:120 } }));
  }
}

// ---- Pillar 2 holdings ----
children.push(h1('Pillar 2 — Growth / 401k', PURPLE));
const p2 = data.holdings.filter(p=>p.pillar===2).sort((a,b)=>posVal(b)-posVal(a));
const wP2 = [1500, 3600, 1100, 1300, 1400];
const p2rows = [ headerRow(['Ticker','Name / Note','Shares','Price','Value'], wP2) ];
for (const p of p2) p2rows.push(row([p.ticker, notesMap[p.ticker]||p.name||'', fmtShares(p.shares), fmtMoney(p.price,2), fmtMoney(posVal(p))], wP2));
p2rows.push(row(['Pillar 2 Total','','','', fmtMoney(s.p2Val)], wP2, [{bold:true,fill:LIGHT},{},{},{},{bold:true,fill:LIGHT}]));
children.push(table(wP2, p2rows));

// ---- Pillar 3 holdings ----
children.push(h1('Pillar 3 — Health / LTC Self-Insurance', BLUE));
children.push(sub('Dividends are swept to Pillar 1 — projections use price-only growth.'));
const p3 = data.holdings.filter(p=>p.pillar===3).sort((a,b)=>posVal(b)-posVal(a));
const wP3 = [1500, 3200, 1000, 1200, 1300, 1100];
const p3rows = [ headerRow(['Ticker','Name / Note','Shares','Price','Value','Div/yr'], wP3) ];
for (const p of p3) p3rows.push(row([p.ticker, notesMap[p.ticker]||p.name||'', fmtShares(p.shares), fmtMoney(p.price,2), fmtMoney(posVal(p)), fmtMoney(posInc(p))], wP3));
p3rows.push(row(['Pillar 3 Total','','','', fmtMoney(s.p3Val), fmtMoney(ai.p3Inc)], wP3, [{bold:true,fill:LIGHT},{},{},{},{bold:true,fill:LIGHT},{bold:true,fill:LIGHT}]));
children.push(table(wP3, p3rows));

// ---- PIMCO Suite Summary (data-driven) ----
const pimco = ['PDI','PDO','PTY','PDX'];
const pimcoHoldings = data.holdings.filter(p=>pimco.includes(p.ticker));
if (pimcoHoldings.length) {
  const pimcoInc = pimcoHoldings.reduce((t,p)=>t+posInc(p),0);
  const pimcoVal = pimcoHoldings.reduce((t,p)=>t+posVal(p),0);
  children.push(h1(`PIMCO Suite — ${fmtMoney(pimcoInc)}/yr (${fmtMoney(pimcoInc/12)}/mo)`, NAVY));
  const wPi = [1400, 2600, 1400, 1600];
  const pirows = [ headerRow(['Ticker','Account','Value','Inc/yr'], wPi) ];
  for (const p of pimcoHoldings.sort((a,b)=>posInc(b)-posInc(a)))
    pirows.push(row([p.ticker, p.acct, fmtMoney(posVal(p)), fmtMoney(posInc(p))], wPi));
  pirows.push(row(['PIMCO Total','', fmtMoney(pimcoVal), fmtMoney(pimcoInc)], wPi, [{bold:true,fill:LIGHT},{},{bold:true,fill:LIGHT},{bold:true,fill:LIGHT}]));
  children.push(table(wPi, pirows));
}

// ---- Crypto ----
children.push(h1('Crypto Holdings (speculative)', GOLD));
children.push(sub(`Total ~${fmtMoney(s.crypto)} · not part of retirement income plan`));

// ═══════════════════════════════════════════════════════════════
// NARRATIVE RENDERER — appends hand-maintained narrative.json below the data
// ═══════════════════════════════════════════════════════════════
function renderNarrative(children) {
  if (!fs.existsSync('./narrative.json')) return;
  const nar = JSON.parse(fs.readFileSync('./narrative.json','utf8'));
  children.push(new Paragraph({ children:[new TextRun({text:'',size:8})], spacing:{after:200}, pageBreakBefore:true }));
  for (const b of nar.blocks) {
    if (b.type === 'h1') children.push(h1(b.text, NAVY));
    else if (b.type === 'h2') children.push(h2(b.text, NAVY));
    else if (b.type === 'p') {
      const hasVerify = b.text.includes('[VERIFY');
      children.push(new Paragraph({ children:[new TextRun({text:b.text, size:17, bold:b.bold||false,
        color: hasVerify ? 'C00000' : '000000', font:'Calibri'})], spacing:{after:120} }));
    }
    else if (b.type === 'table') {
      const rows = [ headerRow(b.header, b.widths) ];
      for (const r of b.rows) {
        const cells = r.map((t,i)=> {
          const hasVerify = String(t).includes('[VERIFY');
          return cell(t, { width:b.widths[i], align: i===0?L:L, size:15,
                           color: hasVerify ? 'C00000' : '000000' });
        });
        rows.push(new TableRow({ children: cells }));
      }
      children.push(table(b.widths, rows));
      children.push(spacer(60));
    }
  }
}
renderNarrative(children);

const doc = new Document({ sections:[{ properties:{ page:{ margin:{ top:720,bottom:720,left:720,right:720 } } }, children }] });
Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('Kitchen Complete Reference.docx', buf);
  console.log('✓ Combined Complete Reference written');
});
