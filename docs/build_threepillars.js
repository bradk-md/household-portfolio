const fs = require('fs');
const { loadData, computeSummary, computeAllInIncome, fmtMoney, fmtShares, fmtPct, grow, posVal, posInc } = require('./lib.js');

const data = loadData();
const s = computeSummary(data);
const ai = computeAllInIncome(data);
const allInInc = s.combinedInc + ai.p2p3Inc;
const skipDiv = new Set(['GPool','VGIntl','BREQI1','WTSCER','Cash','SWVXX','FDRXX']);

// Helpers
const P1 = s.acctsP1;
const capBal = s.capOne.balance;
const capInc = s.capOneInc;
const capAPY = s.capOne.apy;

// Build ticker list string for an account (pillar 1)
function tickerList(acct, pillar = 1) {
  return data.holdings
    .filter(p => p.acct === acct && p.pillar === pillar && !skipDiv.has(p.ticker))
    .map(p => p.ticker)
    .join(' · ');
}

// Pillar 2 accounts
let p2Total = 0, brad401k = 0, lisaRoth401k = 0;
data.holdings.filter(p => p.pillar === 2).forEach(p => {
  const v = posVal(p);
  p2Total += v;
  if (p.acct === 'Brad 401k') brad401k += v;
  if (p.acct === 'Lisa 401k Roth' || p.acct === 'Lisa Roth 401k') lisaRoth401k += v;
});

// Pillar 3 accounts (LTC portions)
let p3Total = 0, bradLTC = 0, lisaLTC = 0;
const bradLTCpos = [], lisaLTCpos = [];
data.holdings.filter(p => p.pillar === 3).forEach(p => {
  const v = posVal(p);
  p3Total += v;
  if (p.acct.startsWith('Brad IRA')) { bradLTC += v; bradLTCpos.push(p); }
  else if (p.acct.startsWith('Lisa IRA')) { lisaLTC += v; lisaLTCpos.push(p); }
});

function ltcNote(positions) {
  return positions.map(p => `${p.ticker} ${fmtShares(p.shares)}${p.ticker==='FDGRX'?'u':'sh'}`).join(' · ');
}

// Pillar 2 fund detail note
function p2Note(acct) {
  const funds = data.holdings.filter(p => p.acct === acct && p.pillar === 2);
  return funds.map(p => `${p.ticker} ${fmtShares(p.shares)}u @ ${fmtMoney(p.price,2)}`).join(' · ');
}

const grandTotal = s.grandTotal;
const p1pct = fmtPct(s.p1Val/grandTotal*100);
const p2pct = fmtPct(p2Total/grandTotal*100);
const p3pct = fmtPct(p3Total/grandTotal*100);

const acctMeta = {
  'Brad IRA':      { label:'Brad IRA — Traditional · Schwab' },
  'Lisa IRA':      { label:'Lisa IRA — Traditional · Schwab' },
  'Brad Roth IRA': { label:'Brad Roth IRA — Tax-Free · Fidelity' },
  'Lisa HSA':      { label:'Lisa HSA — Triple Tax-Free · Fidelity' },
  'Brad HSA':      { label:'Brad HSA — Triple Tax-Free · Fidelity' },
  'Joint':         { label:'Joint Brokerage — Taxable · Schwab' },
};
const acctOrder = ['Brad IRA','Lisa IRA','Brad Roth IRA','Lisa HSA','Brad HSA','Joint'];

function acctCard(acct) {
  const a = P1[acct]; if (!a) return '';
  return `    <div class="acct-card">
      <div class="acct-name">${acctMeta[acct].label}</div>
      <div class="acct-val">${fmtMoney(a.val)}</div>
      <div class="acct-inc">${fmtMoney(a.inc)}/yr · ${fmtMoney(a.inc/12)}/mo</div>
      <div class="acct-note">${tickerList(acct)}</div>
    </div>`;
}

// Income trajectory (3% growth on combined)
function traj(base, years, rate=0.03) { return grow(base, rate, years); }

const html = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Brad & Lisa Kitchen — Three Pillars</title>
<link href="https://fonts.googleapis.com/css2?family=DM+Serif+Display&family=DM+Sans:wght@300;400;500;600&family=DM+Mono&display=swap" rel="stylesheet">
<style>
:root{--navy:#1F3864;--teal:#0D6B52;--blue:#185FA5;--purple:#534AB7;--gold:#B8960C;--gray:#5A5E6B;--border:#E2E4E9;--light:#F5F5F3;}
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:'DM Sans',sans-serif;background:#FAFAF8;color:#1A1C23;padding:2rem;}
.page{max-width:960px;margin:0 auto;}
.hdr{border-bottom:2px solid var(--navy);padding-bottom:1rem;margin-bottom:1.75rem;}
.title{font-family:'DM Serif Display',serif;font-size:2rem;color:var(--navy);}
.sub{font-size:0.85rem;color:var(--gray);margin-top:0.3rem;}
.date{font-family:'DM Mono',monospace;font-size:0.68rem;color:var(--gray);text-align:right;margin-top:0.5rem;}
.summary{display:grid;grid-template-columns:repeat(3,1fr);gap:1rem;margin-bottom:2rem;}
.sum-card{background:#fff;border:1px solid var(--border);border-radius:10px;padding:1.1rem 1.2rem;}
.sum-label{font-size:0.65rem;font-weight:600;text-transform:uppercase;letter-spacing:0.09em;color:var(--gray);margin-bottom:0.5rem;}
.sum-val{font-family:'DM Serif Display',serif;font-size:1.6rem;color:var(--navy);line-height:1.1;}
.sum-sub{font-size:0.75rem;color:var(--gray);margin-top:0.35rem;}
.sum-teal  {border-top:3px solid var(--teal);}
.sum-navy  {border-top:3px solid var(--navy);}
.sum-blue  {border-top:3px solid var(--blue);}
.sum-purple{border-top:3px solid var(--purple);}
.sum-gold  {border-top:3px solid var(--gold);}
.pillar{margin-bottom:2.5rem;}
.pillar-hdr{display:flex;align-items:baseline;gap:1rem;margin-bottom:1rem;}
.pillar-hdr h2{font-family:'DM Serif Display',serif;font-size:1.4rem;color:var(--navy);}
.badge{font-size:0.68rem;font-weight:600;padding:3px 12px;border-radius:20px;letter-spacing:0.06em;}
.badge-teal  {background:#E6F4EF;color:var(--teal);}
.badge-purple{background:#EEEDFE;color:var(--purple);}
.badge-blue  {background:#E6EEF7;color:var(--blue);}
.acct-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:0.85rem;margin-bottom:1rem;}
.acct-card{background:#fff;border:1px solid var(--border);border-radius:8px;padding:0.9rem;}
.acct-name{font-size:0.72rem;font-weight:600;color:var(--gray);text-transform:uppercase;letter-spacing:0.07em;margin-bottom:0.4rem;}
.acct-val{font-family:'DM Mono',monospace;font-size:1rem;font-weight:700;color:var(--navy);}
.acct-inc{font-size:0.78rem;color:var(--teal);font-weight:600;margin-top:0.2rem;}
.acct-note{font-size:0.68rem;color:var(--gray);margin-top:0.3rem;line-height:1.4;}
.subtotal{background:var(--light);border:1px solid var(--border);border-radius:8px;padding:0.75rem 1.1rem;
          display:flex;justify-content:space-between;align-items:center;margin-bottom:0.75rem;}
.subtotal-label{font-family:'DM Serif Display',serif;font-size:1rem;color:var(--navy);}
.subtotal-val{font-family:'DM Mono',monospace;font-weight:700;font-size:1rem;color:var(--teal);}
.proj-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:0.85rem;margin-top:0.75rem;}
.proj-card{background:#fff;border:1px solid var(--border);border-radius:8px;overflow:hidden;}
.proj-hdr-teal  {background:var(--teal);  color:#fff;padding:7px 12px;font-size:0.68rem;font-weight:700;text-transform:uppercase;letter-spacing:0.07em;}
.proj-hdr-purple{background:var(--purple);color:#fff;padding:7px 12px;font-size:0.68rem;font-weight:700;text-transform:uppercase;letter-spacing:0.07em;}
.proj-hdr-blue  {background:var(--blue);  color:#fff;padding:7px 12px;font-size:0.68rem;font-weight:700;text-transform:uppercase;letter-spacing:0.07em;}
.proj-row{display:flex;justify-content:space-between;padding:5px 12px;border-bottom:1px solid var(--border);font-size:11.5px;}
.proj-row:last-child{border-bottom:none;}
.proj-yr{color:var(--gray);}
.proj-val{font-family:'DM Mono',monospace;font-weight:600;color:var(--navy);}
.pending{background:#FEF8E4;border:1px solid #D4A820;border-radius:8px;padding:0.7rem 1rem;font-size:12px;margin-bottom:1rem;}
.pending strong{color:#8A6A00;}
.footer{font-size:0.7rem;color:var(--gray);text-align:center;margin-top:2rem;padding-top:1rem;border-top:1px solid var(--border);}
</style>
</head>
<body>
<div class="page">

<div class="hdr">
  <div class="title">Brad &amp; Lisa Kitchen &nbsp;<span style="font-size:1rem;font-weight:400;font-family:'DM Sans',sans-serif;color:var(--gray)">Household Portfolio</span></div>
  <div class="sub">Self-Directed Three-Pillar Retirement Strategy &nbsp;·&nbsp; Brad Kitchen (Oracle Corp) &amp; Lisa Kitchen (Reed Integration Services)</div>
  <div class="date">As of ${s.asOf} &nbsp;·&nbsp; Source: Schwab &amp; Fidelity exports &nbsp;·&nbsp; Dividend rates: verified</div>
</div>

<!-- SUMMARY -->
<div class="summary">
  <div class="sum-card sum-navy">
    <div class="sum-label">Total Investment Portfolio</div>
    <div class="sum-val">${fmtMoney(grandTotal)}</div>
    <div class="sum-sub">Pillar 1 + 2 + 3 + Capital One + Crypto</div>
  </div>
  <div class="sum-card sum-teal">
    <div class="sum-label">Pillar 1 — Combined Income</div>
    <div class="sum-val">${fmtMoney(s.combinedInc)}/yr</div>
    <div class="sum-sub">${fmtMoney(s.combinedMo)}/mo &nbsp;·&nbsp; incl. Capital One ${fmtMoney(capInc)}/yr</div>
  </div>
  <div class="sum-card sum-teal" style="border:2px solid #0D6B52">
    <div class="sum-label">Total All-In Income</div>
    <div class="sum-val">${fmtMoney(allInInc)}/yr</div>
    <div class="sum-sub">${fmtMoney(allInInc/12)}/mo &nbsp;·&nbsp; incl. ${fmtMoney(ai.p2p3Inc)}/yr P2/P3 divs harvested → redeployed to P1</div>
  </div>
  <div class="sum-card sum-teal">
    <div class="sum-label">Pillar 1 — Income Holdings</div>
    <div class="sum-val">${fmtMoney(s.p1Val)}</div>
    <div class="sum-sub">6 accounts · ${Object.values(P1).reduce((a,b)=>a+b.count,0)} positions · no principal sold</div>
  </div>
  <div class="sum-card sum-purple">
    <div class="sum-label">Pillar 2 — Growth (401k)</div>
    <div class="sum-val">${fmtMoney(p2Total)}</div>
    <div class="sum-sub">Brad Traditional + Lisa Roth · pure growth · ${p2pct}</div>
  </div>
  <div class="sum-card sum-blue">
    <div class="sum-label">Pillar 3 — Health/LTC Reserve</div>
    <div class="sum-val">${fmtMoney(p3Total)}</div>
    <div class="sum-sub">Self-funded · virtual IRA reserve · ${p3pct}</div>
  </div>
</div>

<!-- PILLAR 1 -->
<div class="pillar">
  <div class="pillar-hdr">
    <h2>Pillar 1 — Income</h2>
    <span class="badge badge-teal">6 accounts &nbsp;·&nbsp; ${fmtMoney(s.p1Val)} &nbsp;·&nbsp; ${fmtMoney(s.p1Inc)}/yr portfolio &nbsp;·&nbsp; ${fmtMoney(s.combinedInc)}/yr combined</span>
  </div>
  <div class="acct-grid">
${acctOrder.map(acctCard).join('\n')}
  </div>
  <div class="subtotal">
    <span class="subtotal-label">Pillar 1 Total &nbsp;+&nbsp; Capital One ${fmtMoney(capBal)}</span>
    <span class="subtotal-val">${fmtMoney(s.p1Inc)}/yr portfolio &nbsp;+&nbsp; ${fmtMoney(capInc)}/yr Cap One = ${fmtMoney(s.combinedInc)}/yr combined &nbsp;·&nbsp; ${fmtMoney(s.combinedMo)}/mo</span>
  </div>
  <div class="proj-grid">
    <div class="proj-card">
      <div class="proj-hdr-teal">Income Trajectory — 3% growth</div>
      <div class="proj-row"><span class="proj-yr">Now</span><span class="proj-val">${fmtMoney(s.combinedInc)}/yr</span></div>
      <div class="proj-row"><span class="proj-yr">Year 3</span><span class="proj-val">${fmtMoney(traj(s.combinedInc,3))}/yr</span></div>
      <div class="proj-row"><span class="proj-yr">Year 5</span><span class="proj-val">${fmtMoney(traj(s.combinedInc,5))}/yr</span></div>
      <div class="proj-row"><span class="proj-yr">Year 10</span><span class="proj-val">${fmtMoney(traj(s.combinedInc,10))}/yr</span></div>
      <div class="proj-row"><span class="proj-yr">Year 15</span><span class="proj-val">${fmtMoney(traj(s.combinedInc,15))}/yr</span></div>
    </div>
    <div class="proj-card">
      <div class="proj-hdr-teal">Monthly Income Trajectory</div>
      <div class="proj-row"><span class="proj-yr">Now</span><span class="proj-val">${fmtMoney(s.combinedMo)}/mo</span></div>
      <div class="proj-row"><span class="proj-yr">Year 3</span><span class="proj-val">${fmtMoney(traj(s.combinedInc,3)/12)}/mo</span></div>
      <div class="proj-row"><span class="proj-yr">Year 5</span><span class="proj-val">${fmtMoney(traj(s.combinedInc,5)/12)}/mo</span></div>
      <div class="proj-row"><span class="proj-yr">Year 10</span><span class="proj-val">${fmtMoney(traj(s.combinedInc,10)/12)}/mo</span></div>
      <div class="proj-row"><span class="proj-yr">Year 15</span><span class="proj-val">${fmtMoney(traj(s.combinedInc,15)/12)}/mo</span></div>
    </div>
    <div class="proj-card">
      <div class="proj-hdr-teal">Capital One (${fmtPct(capAPY*100)} APY)</div>
      <div class="proj-row"><span class="proj-yr">Balance (sacred)</span><span class="proj-val">${fmtMoney(capBal)}</span></div>
      <div class="proj-row"><span class="proj-yr">Annual income</span><span class="proj-val">${fmtMoney(capInc)}/yr</span></div>
      <div class="proj-row"><span class="proj-yr">Sweep to</span><span class="proj-val">Joint → MDXBX</span></div>
      <div class="proj-row"><span class="proj-yr">Principal</span><span class="proj-val">Never touched</span></div>
    </div>
  </div>
</div>

<!-- PILLAR 2 -->
<div class="pillar">
  <div class="pillar-hdr">
    <h2>Pillar 2 — Growth</h2>
    <span class="badge badge-purple">Brad 401k (Oracle/Fidelity) + Lisa Roth 401k (Reed Integration) &nbsp;·&nbsp; ${fmtMoney(p2Total)} &nbsp;·&nbsp; pure growth · ${p2pct} of portfolio</span>
  </div>
  <div class="acct-grid" style="grid-template-columns:repeat(2,1fr)">
    <div class="acct-card">
      <div class="acct-name">Brad 401k — Traditional · Oracle / Fidelity</div>
      <div class="acct-val">${fmtMoney(brad401k)}</div>
      <div class="acct-inc" style="color:var(--purple)">Pre-tax · RMDs begin 2035</div>
      <div class="acct-note">${p2Note('Brad 401k')} &nbsp;·&nbsp; Contributing ~$28K/yr</div>
    </div>
    <div class="acct-card">
      <div class="acct-name">Lisa Roth 401k — Reed Integration / Fidelity</div>
      <div class="acct-val">${fmtMoney(lisaRoth401k)}</div>
      <div class="acct-inc" style="color:var(--purple)">Roth · tax-free · no RMD requirement</div>
      <div class="acct-note">${p2Note('Lisa 401k Roth')} &nbsp;·&nbsp; Contributing ~$17K/yr</div>
    </div>
  </div>
  <div class="subtotal">
    <span class="subtotal-label">Pillar 2 Total — Brad &amp; Lisa 401k</span>
    <span class="subtotal-val" style="color:var(--purple)">${fmtMoney(p2Total)} &nbsp;·&nbsp; compounding untouched until retirement</span>
  </div>
  <div class="proj-grid" style="grid-template-columns:repeat(2,1fr)">
    <div class="proj-card">
      <div class="proj-hdr-purple">Growth Projections</div>
      <div class="proj-row"><span class="proj-yr">Now</span><span class="proj-val">${fmtMoney(p2Total)}</span></div>
      <div class="proj-row"><span class="proj-yr">Year 5 · conservative (5%)</span><span class="proj-val">${fmtMoney(grow(p2Total,0.05,5))}</span></div>
      <div class="proj-row"><span class="proj-yr">Year 5 · avg (7%)</span><span class="proj-val">${fmtMoney(grow(p2Total,0.07,5))}</span></div>
      <div class="proj-row"><span class="proj-yr">Year 10 · conservative (5%)</span><span class="proj-val">${fmtMoney(grow(p2Total,0.05,10))}</span></div>
      <div class="proj-row"><span class="proj-yr">Year 10 · avg (7%)</span><span class="proj-val">${fmtMoney(grow(p2Total,0.07,10))}</span></div>
    </div>
    <div class="proj-card">
      <div class="proj-hdr-purple">Key Notes</div>
      <div class="proj-row"><span class="proj-yr">Brad RMDs begin</span><span class="proj-val">2035 (age 73)</span></div>
      <div class="proj-row"><span class="proj-yr">Lisa RMDs</span><span class="proj-val">None required (Roth)</span></div>
      <div class="proj-row"><span class="proj-yr">Brad contributions</span><span class="proj-val">~$28K/yr pre-tax</span></div>
      <div class="proj-row"><span class="proj-yr">Lisa contributions</span><span class="proj-val">~$17K/yr Roth</span></div>
      <div class="proj-row"><span class="proj-yr">Both cease at</span><span class="proj-val">Retirement</span></div>
    </div>
  </div>
</div>

<!-- PILLAR 3 -->
<div class="pillar">
  <div class="pillar-hdr">
    <h2>Pillar 3 — Health/LTC</h2>
    <span class="badge badge-blue">Virtual IRA reserve &nbsp;·&nbsp; ${fmtMoney(p3Total)} &nbsp;·&nbsp; ${bradLTCpos.length + lisaLTCpos.length} positions · 2 IRAs · pure growth · ${p3pct} of portfolio</span>
  </div>
  <div class="acct-grid" style="grid-template-columns:repeat(2,1fr)">
    <div class="acct-card">
      <div class="acct-name">Brad IRA (LTC portion) · Schwab</div>
      <div class="acct-val">${fmtMoney(bradLTC)}</div>
      <div class="acct-inc" style="color:var(--blue)">${fmtPct(bradLTC/p3Total*100)} of pillar &nbsp;·&nbsp; US &amp; intl growth</div>
      <div class="acct-note">${ltcNote(bradLTCpos)}</div>
    </div>
    <div class="acct-card">
      <div class="acct-name">Lisa IRA (LTC portion) · Schwab</div>
      <div class="acct-val">${fmtMoney(lisaLTC)}</div>
      <div class="acct-inc" style="color:var(--blue)">${fmtPct(lisaLTC/p3Total*100)} of pillar &nbsp;·&nbsp; small cap &amp; intl</div>
      <div class="acct-note">${ltcNote(lisaLTCpos)}</div>
    </div>
  </div>
  <div class="subtotal">
    <span class="subtotal-label">Pillar 3 Total — Health/LTC Reserve</span>
    <span class="subtotal-val" style="color:var(--blue)">${fmtMoney(p3Total)} &nbsp;·&nbsp; permanently walled off · never counted as income</span>
  </div>
  <div class="proj-grid">
    <div class="proj-card">
      <div class="proj-hdr-blue">Below Average — 6.6%</div>
      <div class="proj-row"><span class="proj-yr">Now</span><span class="proj-val">${fmtMoney(p3Total)}</span></div>
      <div class="proj-row"><span class="proj-yr">Year 5</span><span class="proj-val">${fmtMoney(grow(p3Total,0.066,5))}</span></div>
      <div class="proj-row"><span class="proj-yr">Year 10</span><span class="proj-val">${fmtMoney(grow(p3Total,0.066,10))} ✓</span></div>
      <div class="proj-row"><span class="proj-yr">Year 15</span><span class="proj-val">${fmtMoney(grow(p3Total,0.066,15))}</span></div>
    </div>
    <div class="proj-card">
      <div class="proj-hdr-blue">Average — 8.5%</div>
      <div class="proj-row"><span class="proj-yr">Now</span><span class="proj-val">${fmtMoney(p3Total)}</span></div>
      <div class="proj-row"><span class="proj-yr">Year 5</span><span class="proj-val">${fmtMoney(grow(p3Total,0.085,5))}</span></div>
      <div class="proj-row"><span class="proj-yr">Year 10</span><span class="proj-val">${fmtMoney(grow(p3Total,0.085,10))} ✓</span></div>
      <div class="proj-row"><span class="proj-yr">Year 15</span><span class="proj-val">${fmtMoney(grow(p3Total,0.085,15))}</span></div>
    </div>
    <div class="proj-card">
      <div class="proj-hdr-blue">Above Average — 10.7%</div>
      <div class="proj-row"><span class="proj-yr">Now</span><span class="proj-val">${fmtMoney(p3Total)}</span></div>
      <div class="proj-row"><span class="proj-yr">Year 5</span><span class="proj-val">${fmtMoney(grow(p3Total,0.107,5))}</span></div>
      <div class="proj-row"><span class="proj-yr">Year 10</span><span class="proj-val">${fmtMoney(grow(p3Total,0.107,10))} ✓</span></div>
      <div class="proj-row"><span class="proj-yr">Year 15</span><span class="proj-val">${fmtMoney(grow(p3Total,0.107,15))}</span></div>
    </div>
  </div>
  <div style="font-size:0.7rem;color:var(--gray);margin-top:0.75rem;font-style:italic;">
    Note: Growth projections use total return rates. Since Pillar 3 dividends (~${fmtMoney(ai.p3Inc)}/yr) are harvested to cash and redeployed into Pillar 1 income rather than reinvested, actual Pillar 3 compounding tracks closer to the below-average (6.6%) scenario by design.
  </div>
</div>

<div class="footer">
  Brad &amp; Lisa Kitchen &nbsp;·&nbsp; Three-Pillar Self-Directed Retirement Strategy &nbsp;·&nbsp; ${s.asOf} &nbsp;·&nbsp; All data sourced from Schwab &amp; Fidelity exports
</div>

</div>
</body>
</html>`;

fs.writeFileSync('/home/claude/rebuild/Kitchen_ThreePillars.html', html);
console.log('Written successfully');
