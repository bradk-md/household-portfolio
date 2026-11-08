/* divgrowth.js — shared dividend-growth math for the Scorecard page and the dashboard P1 tile.
   Data: divgrowth.json (per-ticker expected rate + per-share rate 12 months ago; updated by Claude at refreshes).
   Weights: live dashboard holdings. Share count only sets each holding's WEIGHT; growth is per share.

   Per holding:
     now    = live dashboard div (method 'rate')  or  stored trailing-12-month figure (method 'ttm')
     actual = now / yago - 1        (per-share dividend growth over the last 12 months)
     status = cut (actual < 0) · on track (actual >= expected - 1pt) · behind · n/a
   Portfolio:
     expected = Σ(income × expected) / Σ(income)                      — "mix rate", moves when you buy/sell
     actual   = Σ(shares × now) / Σ(shares × yago) - 1                — same shares, now vs a year ago
*/
var DG = (function () {
  var EXCLUDE = ['Cash', 'SWVXX', 'FDRXX', 'SPAXX', 'SNSXX', 'SNVXX', 'GPool', 'VGIntl', 'BREQI1', 'WTSCER'];
  var TOL = 0.01; // on-track tolerance: within 1 point of expected

  function rows(holdings, dg, pillars) {
    var T = (dg && dg.tickers) || {}, agg = {};
    (holdings || []).forEach(function (h) {
      var t = h.ticker;
      if (EXCLUDE.indexOf(t) >= 0 || !((h.shares || 0) > 0)) return;
      if (!((h.div || 0) > 0) && !T[t]) return;            // zero payers only if we track them (e.g. a suspension)
      if (pillars && pillars.indexOf(h.pillar) < 0) return;
      var a = agg[t] || (agg[t] = { ticker: t, name: h.name || '', shares: 0, div: h.div || 0, pillars: [] });
      a.shares += h.shares;
      if (a.pillars.indexOf(h.pillar) < 0) a.pillars.push(h.pillar);
    });
    return Object.keys(agg).map(function (k) {
      var a = agg[k], t = T[k] || {};
      a.inc = a.shares * a.div;
      a.exp = (t.exp === undefined || t.exp === null) ? null : t.exp;
      a.method = t.method || 'none';
      a.flag = t.flag || '';
      a.note = t.note || '';
      a.yago = (t.method === 'rate' || t.method === 'ttm') ? t.yago : null;
      a.now = t.method === 'rate' ? a.div : (t.method === 'ttm' ? t.now : null);
      a.act = (a.yago > 0 && a.now !== null && a.now !== undefined) ? a.now / a.yago - 1 : null;
      if (a.method === 'excl') a.status = 'excluded';
      else if (a.act === null) a.status = 'n/a';
      else if (a.act < -0.0005) a.status = (a.flag === 'variable') ? 'down' : 'cut';
      else if (a.exp === null) a.status = 'no target';
      else if (a.act >= a.exp - TOL) a.status = 'on track';
      else a.status = 'behind';
      return a;
    });
  }

  function summary(rs) {
    var tot = 0, wi = 0, we = 0, nn = 0, ny = 0, ia = 0;
    rs.forEach(function (r) {
      tot += r.inc;
      if (r.exp !== null && r.inc > 0) { wi += r.inc; we += r.inc * r.exp; }
      if (r.act !== null) { nn += r.shares * r.now; ny += r.shares * r.yago; ia += r.inc; }
    });
    rs.forEach(function (r) { r.contrib = (r.act !== null && ny > 0) ? r.shares * (r.now - r.yago) / ny : null; });
    return {
      expected: wi ? we / wi : null, actual: ny ? nn / ny - 1 : null,
      dollars: nn - ny, income: tot, expCov: tot ? wi / tot : 0, actCov: tot ? ia / tot : 0
    };
  }

  function load() {
    return fetch('divgrowth.json?nc=' + Date.now()).then(function (r) { return r.ok ? r.json() : null; });
  }

  return { rows: rows, summary: summary, load: load, TOL: TOL };
})();
