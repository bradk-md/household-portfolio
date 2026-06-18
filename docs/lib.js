const fs = require('fs');

function loadData() {
  const raw = JSON.parse(fs.readFileSync('/home/claude/rebuild/data_clean.json', 'utf8'));
  return raw;
}

function fmtMoney(n, decimals = 0) {
  const sign = n < 0 ? '-' : '';
  const abs = Math.abs(n);
  return sign + '$' + abs.toLocaleString('en-US', { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
}

function fmtShares(n) {
  if (Number.isInteger(n)) return n.toLocaleString('en-US');
  // Trim to reasonable precision, drop trailing zeros
  let s = n.toFixed(4);
  s = s.replace(/0+$/, '').replace(/\.$/, '');
  const parts = s.split('.');
  parts[0] = Number(parts[0]).toLocaleString('en-US');
  return parts.join('.');
}

function fmtPct(n, decimals = 1) {
  return n.toFixed(decimals) + '%';
}

function grow(value, rate, years) {
  return value * Math.pow(1 + rate, years);
}

// Compute summary aggregates fresh from holdings (single source of truth)
function computeSummary(data) {
  const h = data.holdings;
  const capOne = { balance: 60000, apy: 0.032 };
  const crypto = data.summary.crypto;
  const capOneInc = capOne.balance * capOne.apy;

  let p1Val = 0, p1Inc = 0, p1Cost = 0, p2Val = 0, p3Val = 0;
  const acctsP1 = {};
  for (const pos of h) {
    const val = pos.shares * pos.price;
    const inc = pos.shares * pos.div;
    const cost = pos.shares * (pos.costBasis !== undefined && pos.costBasis !== null && pos.costBasis !== '-' ? pos.costBasis : pos.price);
    if (pos.pillar === 1) {
      p1Val += val; p1Inc += inc; p1Cost += cost;
      if (!acctsP1[pos.acct]) acctsP1[pos.acct] = { val: 0, inc: 0, count: 0 };
      acctsP1[pos.acct].val += val;
      acctsP1[pos.acct].inc += inc;
      acctsP1[pos.acct].count += 1;
    } else if (pos.pillar === 2) {
      p2Val += val;
    } else if (pos.pillar === 3) {
      p3Val += val;
    }
  }
  const investTotal = p1Val + p2Val + p3Val;
  const grandTotal = investTotal + capOne.balance + crypto;
  const combinedInc = p1Inc + capOneInc;

  return {
    asOf: data.summary.asOf,
    p1Val, p1Inc, p1Cost, p2Val, p3Val,
    investTotal, grandTotal, capOne, capOneInc, crypto, combinedInc,
    combinedMo: combinedInc / 12,
    acctsP1
  };
}

function getHoldingsByAcct(data, acct, pillar = 1) {
  return data.holdings.filter(p => p.acct === acct && p.pillar === pillar);
}

function posVal(p) { return p.shares * p.price; }
function posInc(p) { return p.shares * p.div; }
function posYield(p) { const v = posVal(p); return v ? (posInc(p) / v * 100) : 0; }

module.exports = { loadData, fmtMoney, fmtShares, fmtPct, grow, computeSummary, getHoldingsByAcct, posVal, posInc, posYield };
