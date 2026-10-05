// Programmatic Action Items: renders hand-maintained action_items.json +
// pulls the live Income Summary from the dashboard data (data_clean.json via lib.js).
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, AlignmentType,
        WidthType, ShadingType, BorderStyle, VerticalAlign, HeadingLevel } = require('docx');
const fs = require('fs');
const { loadData, computeSummary, computeAllInIncome, fmtMoney } = require('./lib.js');

const data = loadData();
const s = computeSummary(data);
const ai = computeAllInIncome(data);
const A = JSON.parse(fs.readFileSync('./action_items.json','utf8'));

// Dynamic placeholders computed from live holdings (e.g. shares remaining to a target)
function sharesOf(ticker) { return data.holdings.filter(p=>p.ticker===ticker).reduce((t,p)=>t+p.shares,0); }
function sharesOfAcct(ticker, acct) { return data.holdings.filter(p=>p.ticker===ticker && p.acct===acct).reduce((t,p)=>t+p.shares,0); }
const dyn = A._dynamic || {};
const SCHD_REMAINING = dyn.SCHD_target ? Math.max(0, dyn.SCHD_target - sharesOf('SCHD')) : 0;
const BCX_REMAINING = dyn.BCX_target ? Math.max(0, dyn.BCX_target - sharesOfAcct('BCX', dyn.BCX_target_acct || 'Lisa IRA')) : 0;
const THW_REMAINING = dyn.THW_target ? Math.max(0, dyn.THW_target - sharesOf('THW')) : 0;
function subst(text) {
  return String(text)
    .replace('{SCHD_REMAINING}', SCHD_REMAINING.toLocaleString())
    .replace('{BCX_REMAINING}', BCX_REMAINING.toLocaleString())
    .replace('{THW_REMAINING}', THW_REMAINING.toLocaleString());
}

const NAVY='1F3864', TEAL='0D6B52', GRAY='5A5E6B', LIGHT='F5F5F3', CAT='E8ECF3';
const border = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
const borders = { top: border, bottom: border, left: border, right: border };
const L = AlignmentType.LEFT;

function cell(text, opts = {}) {
  const { bold=false, color='000000', fill=null, width=1500, size=16, italics=false, span=null } = opts;
  const c = new TableCell({
    borders, width: { size: width, type: WidthType.DXA },
    shading: fill ? { fill, type: ShadingType.CLEAR } : undefined,
    margins: { top: 50, bottom: 50, left: 100, right: 100 },
    verticalAlign: VerticalAlign.CENTER,
    columnSpan: span || undefined,
    children: [new Paragraph({ alignment: L, children: [new TextRun({ text: String(text), bold, color, size, font: 'Calibri', italics })] })]
  });
  return c;
}
function table(widths, rows) { return new Table({ width:{size:widths.reduce((a,b)=>a+b,0),type:WidthType.DXA}, columnWidths:widths, rows }); }
function h1(text,color=NAVY){ return new Paragraph({children:[new TextRun({text,bold:true,size:28,color,font:'Calibri'})],spacing:{before:220,after:100}}); }
function h2(text,color=NAVY){ return new Paragraph({children:[new TextRun({text,bold:true,size:22,color,font:'Calibri'})],spacing:{before:180,after:80}}); }
// Same look as h2, but carries Word's built-in Heading1 style so the paragraph has an
// outline level — required for Word's collapse-heading feature. A post-processing pass
// (below) then marks this specific paragraph collapsed-by-default via the w15:collapsed
// extension, since the docx library has no API for it.
function collapsibleH2(text,color=NAVY){ return new Paragraph({heading:HeadingLevel.HEADING_1, children:[new TextRun({text,bold:true,size:22,color,font:'Calibri'})],spacing:{before:180,after:80}}); }
function para(text,opts={}){ return new Paragraph({children:[new TextRun({text,size:opts.size||17,color:opts.color||'000000',bold:opts.bold||false,italics:opts.italics||false,font:'Calibri'})],spacing:{after:opts.after||100}}); }

const children = [];
children.push(new Paragraph({ children:[new TextRun({text:'Brad & Lisa Kitchen',bold:true,size:32,color:NAVY,font:'Calibri'})], spacing:{after:40} }));
children.push(h1(A.title, NAVY));
children.push(para(`As of ${s.asOf}  ·  ${A.subtitle}`, {color:GRAY, size:18, after:60}));
children.push(para(`Status Key:  ${A.statusKey}`, {color:GRAY, size:15, italics:true, after:160}));

// ---- Live Income Summary (from dashboard) ----
children.push(h2('Income Summary', TEAL));
const wI = [4600, 2200, 2200, 1600];
const allInInc = s.combinedInc + ai.p2p3Inc;
function hdr(cells, widths) { return new TableRow({ children: cells.map((t,i)=>cell(t,{bold:true,fill:'EEEEEE',width:widths[i],size:15})) }); }
function drow(cells, widths, opts=[]) { return new TableRow({ children: cells.map((t,i)=>cell(t,{width:widths[i],size:15,...(opts[i]||{})})) }); }
children.push(table(wI, [
  hdr(['Income','Annual','Monthly','Status'], wI),
  drow(['Current combined income (portfolio + Cap One)', fmtMoney(s.combinedInc), fmtMoney(s.combinedMo), 'Confirmed ✅'], wI, [{bold:true,color:TEAL},{bold:true,color:TEAL},{bold:true,color:TEAL},{}]),
  drow(['Pillar 2/3 divs swept to Pillar 1', fmtMoney(ai.p2p3Inc), fmtMoney(ai.p2p3Inc/12), 'Reinvested'], wI),
  drow(['Total All-In Income', fmtMoney(allInInc), fmtMoney(allInInc/12), ''], wI, [{bold:true,fill:LIGHT},{bold:true,fill:LIGHT},{bold:true,fill:LIGHT},{fill:LIGHT}]),
]));
children.push(para(A.debtLine, {color:GRAY, size:15, italics:true, after:40}));

// ---- Action item sections ----
const wT = [1300, 4200, 3200, 1800];
for (const sec of A.sections) {
  children.push(sec.header === 'Completed' ? collapsibleH2(sec.header, NAVY) : h2(sec.header, NAVY));
  const rows = [ hdr(['Status','Action Item','Account / Notes','Target Date'], wT) ];
  for (const it of sec.items) {
    if (it.category) {
      rows.push(new TableRow({ children:[ cell(it.category, {bold:true, fill:CAT, width:wT.reduce((a,b)=>a+b,0), size:15, span:4}) ] }));
    } else {
      rows.push(drow([it.status||'', subst(it.action||''), subst(it.notes||''), it.date||''], wT));
    }
  }
  children.push(table(wT, rows));
}

// ---- HTML version (same data; "Completed" is a native <details>, collapsed by default) ----
function esc(t){ return String(t).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
function buildHtml() {
  const secHtml = A.sections.map(sec => {
    const collapsible = sec.header === 'Completed';
    const count = sec.items.filter(i=>!i.category).length;
    let rows = '';
    for (const it of sec.items) {
      if (it.category) rows += `<tr class="cat"><td colspan="4">${esc(it.category)}</td></tr>`;
      else rows += `<tr><td class="st">${esc(it.status||'')}</td><td>${esc(subst(it.action||''))}</td><td class="nt">${esc(subst(it.notes||''))}</td><td class="dt">${esc(it.date||'')}</td></tr>`;
    }
    const table = `<div class="tw"><table><thead><tr><th>Status</th><th>Action Item</th><th>Account / Notes</th><th>Target Date</th></tr></thead><tbody>${rows}</tbody></table></div>`;
    return collapsible
      ? `<details class="sec"><summary><span class="chev"></span>${esc(sec.header)} <span class="cnt">${count} items — click to expand</span></summary>${table}</details>`
      : `<section class="sec"><h2>${esc(sec.header)}</h2>${table}</section>`;
  }).join('\n');
  const html = `<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Brad &amp; Lisa Kitchen — Action Items</title>
<link href="https://fonts.googleapis.com/css2?family=DM+Serif+Display&family=DM+Sans:wght@300;400;500;600&display=swap" rel="stylesheet">
<style>
:root{--navy:#1F3864;--teal:#0D6B52;--gray:#5A5E6B;--border:#E2E4E9;--light:#F5F5F3;--cat:#E8ECF3;}
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:'DM Sans',system-ui,sans-serif;background:#FAFAF8;color:#1A1C23;padding:2rem 1rem;}
.page{max-width:1000px;margin:0 auto;}
.hdr{border-bottom:2px solid var(--navy);padding-bottom:1rem;margin-bottom:1.5rem;}
.brand{font-family:'DM Serif Display',serif;font-size:1.9rem;color:var(--navy);}
.title{font-family:'DM Serif Display',serif;font-size:1.2rem;color:var(--navy);margin-top:.2rem;}
.sub{font-size:.8rem;color:var(--gray);margin-top:.3rem;}
.key{font-size:.75rem;color:var(--gray);font-style:italic;margin-top:.4rem;}
h2{font-family:'DM Serif Display',serif;font-size:1.25rem;color:var(--navy);margin:1.6rem 0 .6rem;}
h2.teal{color:var(--teal);}
.tw{overflow-x:auto;background:#fff;border:1px solid var(--border);border-radius:8px;}
table{width:100%;border-collapse:collapse;font-size:.82rem;}
th{background:#EEE;text-align:left;font-size:.72rem;text-transform:uppercase;letter-spacing:.05em;padding:7px 10px;}
td{padding:7px 10px;border-top:1px solid var(--border);vertical-align:top;}
td.st,td.dt{white-space:nowrap;}
td.nt{color:var(--gray);}
tr.cat td{background:var(--cat);font-weight:600;color:var(--navy);font-size:.78rem;}
.inc td:nth-child(n+2){font-variant-numeric:tabular-nums;}
.inc tr.hl td{font-weight:700;color:var(--teal);}
.inc tr.tot td{font-weight:700;background:var(--light);}
.debt{font-size:.78rem;color:var(--gray);font-style:italic;margin:.6rem 0 0;}
details.sec{margin-top:1.6rem;}
details.sec>summary{cursor:pointer;list-style:none;font-family:'DM Serif Display',serif;font-size:1.25rem;color:var(--navy);padding:.5rem 0;display:flex;align-items:baseline;gap:.5rem;}
details.sec>summary::-webkit-details-marker{display:none;}
.chev{display:inline-block;width:.55em;height:.55em;border-right:2px solid var(--navy);border-bottom:2px solid var(--navy);transform:rotate(-45deg);transition:transform .15s;margin-right:.2rem;}
details[open]>summary .chev{transform:rotate(45deg);}
.cnt{font-family:'DM Sans',sans-serif;font-size:.75rem;color:var(--gray);}
details[open] .cnt{display:none;}
</style></head><body><div class="page">
<div class="hdr"><div class="brand">Brad &amp; Lisa Kitchen</div><div class="title">${esc(A.title)}</div>
<div class="sub">As of ${esc(s.asOf)} · ${esc(A.subtitle)}</div><div class="key">Status Key: ${esc(A.statusKey)}</div></div>
<section><h2 class="teal">Income Summary</h2><div class="tw"><table class="inc"><thead><tr><th>Income</th><th>Annual</th><th>Monthly</th><th>Status</th></tr></thead><tbody>
<tr class="hl"><td>Current combined income (portfolio + Cap One)</td><td>${fmtMoney(s.combinedInc)}</td><td>${fmtMoney(s.combinedMo)}</td><td>Confirmed ✅</td></tr>
<tr><td>Pillar 2/3 divs swept to Pillar 1</td><td>${fmtMoney(ai.p2p3Inc)}</td><td>${fmtMoney(ai.p2p3Inc/12)}</td><td>Reinvested</td></tr>
<tr class="tot"><td>Total All-In Income</td><td>${fmtMoney(allInInc)}</td><td>${fmtMoney(allInInc/12)}</td><td></td></tr>
</tbody></table></div><p class="debt">${esc(A.debtLine)}</p></section>
${secHtml}
</div></body></html>`;
  fs.writeFileSync('Kitchen Action Items.html', html);
}
buildHtml();

const doc = new Document({ sections:[{ properties:{ page:{ margin:{ top:720,bottom:720,left:720,right:720 } } }, children }] });
Packer.toBuffer(doc).then(async buf => {
  const outPath = 'Kitchen Action Items.docx';
  fs.writeFileSync(outPath, buf);
  await markCompletedCollapsed(outPath);
  console.log('✓ Action Items written');
});

// Post-process the generated .docx: mark the "Completed" Heading1 paragraph collapsed-by-
// default (Word's w15:collapsed extension — not exposed by the docx library). Word desktop
// (2013+) honors this; other viewers (Google Docs, Word Online, LibreOffice) typically just
// show the section expanded and ignore the hint, which is a safe fallback.
async function markCompletedCollapsed(path) {
  const JSZip = require('jszip');
  const buf = fs.readFileSync(path);
  const zip = await JSZip.loadAsync(buf);
  const docXmlPath = 'word/document.xml';
  let xml = await zip.file(docXmlPath).async('string');
  // Find the Heading1 paragraph whose text is exactly "Completed" and inject
  // <w:outlineLvl w:val="0"/> (direct formatting — makes Word treat this specific
  // paragraph as an outline-level-0 heading even though the generated Heading1 style
  // itself doesn't define one) plus <w15:collapsed w15:val="1"/>, appended at the end
  // of its <w:pPr> (CT_PPr requires outlineLvl after spacing/ind/jc etc.; the w15
  // extension element goes after that, before </w:pPr>).
  const re = /(<w:p\b[^>]*>\s*<w:pPr>\s*<w:pStyle w:val="Heading1"\s*\/>.*?)(<\/w:pPr>)(.*?Completed.*?<\/w:p>)/s;
  const match = xml.match(re);
  if (!match) {
    console.warn('⚠ Could not find Completed Heading1 paragraph to mark collapsed — section will open expanded.');
    return;
  }
  const patched = match[1] + '<w:outlineLvl w:val="0"/><w15:collapsed w15:val="1"/>' + match[2] + match[3];
  xml = xml.slice(0, match.index) + patched + xml.slice(match.index + match[0].length);
  zip.file(docXmlPath, xml);
  const newBuf = await zip.generateAsync({ type: 'nodebuffer' });
  fs.writeFileSync(path, newBuf);
}
