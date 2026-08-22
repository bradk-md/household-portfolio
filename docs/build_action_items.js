// Programmatic Action Items: renders hand-maintained action_items.json +
// pulls the live Income Summary from the dashboard data (data_clean.json via lib.js).
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, AlignmentType,
        WidthType, ShadingType, BorderStyle, VerticalAlign } = require('docx');
const fs = require('fs');
const { loadData, computeSummary, computeAllInIncome, fmtMoney } = require('./lib.js');

const data = loadData();
const s = computeSummary(data);
const ai = computeAllInIncome(data);
const A = JSON.parse(fs.readFileSync('./action_items.json','utf8'));

// Dynamic placeholders computed from live holdings (e.g. shares remaining to a target)
function sharesOf(ticker) { return data.holdings.filter(p=>p.ticker===ticker).reduce((t,p)=>t+p.shares,0); }
const dyn = A._dynamic || {};
const SCHD_REMAINING = dyn.SCHD_target ? Math.max(0, dyn.SCHD_target - sharesOf('SCHD')) : 0;
function subst(text) {
  return String(text).replace('{SCHD_REMAINING}', SCHD_REMAINING.toLocaleString());
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
  children.push(h2(sec.header, NAVY));
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

const doc = new Document({ sections:[{ properties:{ page:{ margin:{ top:720,bottom:720,left:720,right:720 } } }, children }] });
Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('Kitchen_Action_Items.docx', buf);
  console.log('✓ Action Items written');
});
