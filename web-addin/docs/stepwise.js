/*
 * Stepwise regression engine. Pure JavaScript with no Office dependency, so it
 * runs in the task pane and under Node for tests.
 *
 * Methods:   1 = forward selection, 2 = backward elimination, 3 = stepwise
 * Criteria:  1 = T-Stat  (enter when |t| >= enter, remove when |t| < remove)
 *            2 = P-value (enter when p < enter,    remove when p > remove)
 * Entry and removal use the partial t-test on the candidate's coefficient,
 * which is equivalent to the partial F-test.
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    module.exports = factory();
  } else {
    root.Stepwise = factory();
  }
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  var METHOD = { FORWARD: 1, BACKWARD: 2, STEPWISE: 3 };
  var CRIT = { T: 1, P: 2 };
  var SING_TOL = 1e-10;
  var MAX_STEPS = 1000;

  // ---------------------------------------------------------------------------
  // Distributions
  // ---------------------------------------------------------------------------

  function logGamma(x) {
    // Lanczos approximation, g = 7, n = 9
    var c = [0.99999999999980993, 676.5203681218851, -1259.1392167224028,
      771.32342877765313, -176.61502916214059, 12.507343278686905,
      -0.13857109526572012, 9.9843695780195716e-6, 1.5056327351493116e-7];
    if (x < 0.5) {
      return Math.log(Math.PI / Math.sin(Math.PI * x)) - logGamma(1 - x);
    }
    x -= 1;
    var a = c[0];
    var t = x + 7.5;
    for (var i = 1; i < 9; i++) a += c[i] / (x + i);
    return 0.5 * Math.log(2 * Math.PI) + (x + 0.5) * Math.log(t) - t + Math.log(a);
  }

  // Continued fraction for the incomplete beta function (Numerical Recipes betacf)
  function betaContinuedFraction(x, a, b) {
    var FPMIN = 1e-300;
    var qab = a + b, qap = a + 1, qam = a - 1;
    var c = 1, d = 1 - (qab * x) / qap;
    if (Math.abs(d) < FPMIN) d = FPMIN;
    d = 1 / d;
    var h = d;
    for (var m = 1; m <= 300; m++) {
      var m2 = 2 * m;
      var aa = (m * (b - m) * x) / ((qam + m2) * (a + m2));
      d = 1 + aa * d; if (Math.abs(d) < FPMIN) d = FPMIN;
      c = 1 + aa / c; if (Math.abs(c) < FPMIN) c = FPMIN;
      d = 1 / d; h *= d * c;
      aa = (-(a + m) * (qab + m) * x) / ((a + m2) * (qap + m2));
      d = 1 + aa * d; if (Math.abs(d) < FPMIN) d = FPMIN;
      c = 1 + aa / c; if (Math.abs(c) < FPMIN) c = FPMIN;
      d = 1 / d;
      var del = d * c;
      h *= del;
      if (Math.abs(del - 1) < 3e-16) break;
    }
    return h;
  }

  // Regularized incomplete beta I_x(a, b)
  function incompleteBeta(x, a, b) {
    if (x <= 0) return 0;
    if (x >= 1) return 1;
    var lbt = logGamma(a + b) - logGamma(a) - logGamma(b) + a * Math.log(x) + b * Math.log(1 - x);
    var bt = Math.exp(lbt);
    if (x < (a + 1) / (a + b + 2)) return (bt * betaContinuedFraction(x, a, b)) / a;
    return 1 - (bt * betaContinuedFraction(1 - x, b, a)) / b;
  }

  // Two-tailed p-value of a t statistic (Excel T.DIST.2T)
  function tDist2T(t, df) {
    if (!isFinite(t)) return 0;
    return incompleteBeta(df / (df + t * t), df / 2, 0.5);
  }

  // Right-tail p-value of an F statistic (Excel F.DIST.RT)
  function fDistRT(f, d1, d2) {
    if (!isFinite(f)) return 0;
    if (f <= 0) return 1;
    return incompleteBeta(d2 / (d2 + d1 * f), d2 / 2, d1 / 2);
  }

  // Two-tailed critical t value (Excel T.INV.2T)
  function tInv2T(alpha, df) {
    var lo = 0, hi = 1;
    while (tDist2T(hi, df) > alpha) hi *= 2;
    for (var i = 0; i < 200; i++) {
      var mid = (lo + hi) / 2;
      if (tDist2T(mid, df) > alpha) lo = mid; else hi = mid;
    }
    return (lo + hi) / 2;
  }

  function tStat(coef, se) {
    if (se > 0) return coef / se;
    return coef === 0 ? 0 : (coef < 0 ? -Infinity : Infinity);
  }

  function pValue(coef, se, df) {
    return tDist2T(Math.abs(tStat(coef, se)), df);
  }

  // ---------------------------------------------------------------------------
  // Data
  // ---------------------------------------------------------------------------

  function isNumber(v) {
    return typeof v === "number" && isFinite(v);
  }

  /*
   * yColumn: array of cell values for Y.
   * xColumns: array of { values: [...], name: "..." }, one per candidate column.
   * hasLabels: first value of each column is its label.
   * Returns { ok, error, data, dropped }.
   */
  function buildData(yColumn, xColumns, hasLabels, yFallbackName) {
    var nRows = yColumn.length;
    var first = hasLabels ? 1 : 0;
    var k = xColumns.length;
    if (k === 0) return { ok: false, error: "Select at least one X column." };
    for (var j = 0; j < k; j++) {
      if (xColumns[j].values.length !== nRows) {
        return { ok: false, error: "Every X column must cover the same rows as Y." };
      }
    }

    function label(v, fallback) {
      if (v === null || v === undefined) return fallback;
      var s = String(v).trim();
      return s.length ? s : fallback;
    }

    var names = xColumns.map(function (c) {
      return hasLabels ? label(c.values[0], c.name) : c.name;
    });
    var yName = hasLabels ? label(yColumn[0], yFallbackName || "Y") : (yFallbackName || "Y");

    var y = [], x = [];
    for (var r = first; r < nRows; r++) {
      if (!isNumber(yColumn[r])) continue;
      var row = [];
      var ok = true;
      for (j = 0; j < k; j++) {
        var v = xColumns[j].values[r];
        if (!isNumber(v)) { ok = false; break; }
        row.push(v);
      }
      if (!ok) continue;
      y.push(yColumn[r]);
      x.push(row);
    }
    var n = y.length;
    var dropped = (nRows - first) - n;
    if (n < 3) {
      return { ok: false, error: "Fewer than 3 complete numeric rows. Check the ranges and the labels setting." };
    }

    var yBar = 0, xBar = new Array(k).fill(0);
    for (var i = 0; i < n; i++) {
      yBar += y[i];
      for (j = 0; j < k; j++) xBar[j] += x[i][j];
    }
    yBar /= n;
    for (j = 0; j < k; j++) xBar[j] /= n;

    var syy = 0;
    var cxy = new Array(k).fill(0);
    var cxx = [];
    for (j = 0; j < k; j++) cxx.push(new Array(k).fill(0));
    for (i = 0; i < n; i++) {
      var dy = y[i] - yBar;
      syy += dy * dy;
      for (j = 0; j < k; j++) {
        var dj = x[i][j] - xBar[j];
        cxy[j] += dj * dy;
        for (var c = j; c < k; c++) cxx[j][c] += dj * (x[i][c] - xBar[c]);
      }
    }
    for (j = 0; j < k; j++) for (c = j + 1; c < k; c++) cxx[c][j] = cxx[j][c];

    if (syy <= 0) {
      return { ok: false, error: "Y is constant across the usable rows, so there is nothing to explain." };
    }
    return {
      ok: true,
      dropped: dropped,
      data: { n: n, k: k, y: y, x: x, yBar: yBar, xBar: xBar, syy: syy, cxx: cxx, cxy: cxy, names: names, yName: yName }
    };
  }

  // ---------------------------------------------------------------------------
  // Least squares
  // ---------------------------------------------------------------------------

  // Gauss-Jordan inverse of a symmetric positive semi-definite matrix, no pivoting.
  // Each pivot is 1 - R^2 of that column on the earlier ones times its diagonal,
  // so the tolerance check rejects constant and collinear predictors.
  function invertSPD(a) {
    var m = a.length;
    a = a.map(function (row) { return row.slice(); });
    var inv = [];
    var d = [];
    for (var i = 0; i < m; i++) {
      inv.push(new Array(m).fill(0));
      inv[i][i] = 1;
      d.push(a[i][i]);
    }
    for (var c = 0; c < m; c++) {
      var piv = a[c][c];
      if (d[c] <= 0 || piv <= SING_TOL * d[c]) return null;
      for (var j = 0; j < m; j++) { a[c][j] /= piv; inv[c][j] /= piv; }
      for (var r = 0; r < m; r++) {
        if (r === c) continue;
        var f = a[r][c];
        if (f === 0) continue;
        for (j = 0; j < m; j++) { a[r][j] -= f * a[c][j]; inv[r][j] -= f * inv[c][j]; }
      }
    }
    return inv;
  }

  // Fits Y on the given column indexes. Returns null when not estimable.
  function fit(ds, vars) {
    var p = vars.length;
    var dfe = ds.n - p - 1;
    if (dfe < 1) return null;
    if (p === 0) {
      return { vars: [], coef: [ds.yBar], se: [Math.sqrt(ds.syy / dfe / ds.n)], sse: ds.syy, dfe: dfe };
    }
    var a = vars.map(function (vi) { return vars.map(function (vj) { return ds.cxx[vi][vj]; }); });
    var inv = invertSPD(a);
    if (!inv) return null;

    var b = [];
    for (var i = 0; i < p; i++) {
      var s = 0;
      for (var j = 0; j < p; j++) s += inv[i][j] * ds.cxy[vars[j]];
      b.push(s);
    }
    var b0 = ds.yBar;
    for (i = 0; i < p; i++) b0 -= b[i] * ds.xBar[vars[i]];

    var sse = 0;
    for (var r = 0; r < ds.n; r++) {
      var e = ds.y[r] - b0;
      for (j = 0; j < p; j++) e -= b[j] * ds.x[r][vars[j]];
      sse += e * e;
    }
    var mse = sse / dfe;
    var v0 = 1 / ds.n;
    for (i = 0; i < p; i++) for (j = 0; j < p; j++) v0 += ds.xBar[vars[i]] * inv[i][j] * ds.xBar[vars[j]];
    var se = [Math.sqrt(mse * v0)];
    for (i = 0; i < p; i++) se.push(Math.sqrt(mse * inv[i][i]));
    return { vars: vars.slice(), coef: [b0].concat(b), se: se, sse: sse, dfe: dfe };
  }

  function fitCurrent(ds, inModel) {
    var vars = [];
    for (var j = 0; j < ds.k; j++) if (inModel[j]) vars.push(j);
    return fit(ds, vars);
  }

  // ---------------------------------------------------------------------------
  // Selection
  // ---------------------------------------------------------------------------

  function defaultThresholds(crit) {
    return crit === CRIT.T ? { enter: 2, remove: 2 } : { enter: 0.05, remove: 0.1 };
  }

  // Returns an error message, or "" when the thresholds are usable
  function checkThresholds(method, crit, enter, remove) {
    if (crit === CRIT.T) {
      if (!(enter > 0) || !(remove > 0)) return "|T-Stat| thresholds must be greater than 0.";
      if (method === METHOD.STEPWISE && enter < remove) {
        return "|T-Stat| to enter must be at least |T-Stat| to remove, otherwise a variable can cycle in and out.";
      }
    } else {
      if (!(enter > 0 && enter < 1) || !(remove > 0 && remove < 1)) return "P-value thresholds must be between 0 and 1.";
      if (method === METHOD.STEPWISE && enter > remove) {
        return "P-value to enter must not exceed p-value to remove, otherwise a variable can cycle in and out.";
      }
    }
    return "";
  }

  function stepSummary(ds, action, name, t, p, f) {
    var k = f.vars.length;
    return {
      action: action,
      variable: name,
      tStat: t,
      pValue: p,
      terms: k,
      r2: 1 - f.sse / ds.syy,
      adjR2: 1 - (f.sse / f.dfe) / (ds.syy / (ds.n - 1)),
      rmse: Math.sqrt(f.sse / f.dfe),
      aic: f.sse > 0 ? ds.n * Math.log(f.sse / ds.n) + 2 * (k + 1) : null,
      bic: f.sse > 0 ? ds.n * Math.log(f.sse / ds.n) + (k + 1) * Math.log(ds.n) : null
    };
  }

  /*
   * Runs the selection. Returns { ok, error, inModel, steps, hitLimit }.
   */
  function select(ds, method, crit, enter, remove) {
    var inModel = new Array(ds.k).fill(method === METHOD.BACKWARD);
    var steps = [];

    function enterOK(t, p) { return crit === CRIT.T ? t >= enter : p < enter; }
    function removeOK(t, p) { return crit === CRIT.T ? t < remove : p > remove; }

    var start = fitCurrent(ds, inModel);
    if (method === METHOD.BACKWARD) {
      if (ds.n < ds.k + 2) {
        return { ok: false, error: "Backward elimination needs at least " + (ds.k + 2) + " complete rows for " +
          ds.k + " predictors; only " + ds.n + " are usable. Use forward or stepwise." };
      }
      if (!start) {
        return { ok: false, error: "The model with every predictor cannot be estimated: some predictors are " +
          "constant or perfectly collinear. Remove the redundant columns or use forward or stepwise." };
      }
      steps.push(stepSummary(ds, "Start", "(all candidates)", null, null, start));
    } else {
      steps.push(stepSummary(ds, "Start", "(intercept only)", null, null, start));
    }

    function tryEnter() {
      var best = -1, bestT = -1, bestFit = null;
      for (var j = 0; j < ds.k; j++) {
        if (inModel[j]) continue;
        inModel[j] = true;
        var f = fitCurrent(ds, inModel);
        inModel[j] = false;
        if (!f) continue;
        var pos = f.vars.indexOf(j) + 1;
        // Candidates share the same residual df, so the largest |t| has the smallest p-value
        var t = Math.abs(tStat(f.coef[pos], f.se[pos]));
        if (t > bestT) { bestT = t; best = j; bestFit = f; }
      }
      if (best < 0) return false;
      var bp = bestFit.vars.indexOf(best) + 1;
      var pv = pValue(bestFit.coef[bp], bestFit.se[bp], bestFit.dfe);
      if (!enterOK(bestT, pv)) return false;
      inModel[best] = true;
      steps.push(stepSummary(ds, "Entered", ds.names[best], tStat(bestFit.coef[bp], bestFit.se[bp]), pv, bestFit));
      return true;
    }

    function tryRemove() {
      var f = fitCurrent(ds, inModel);
      if (!f || f.vars.length === 0) return false;
      var worst = 0, worstT = Infinity;
      for (var i = 1; i <= f.vars.length; i++) {
        var t = Math.abs(tStat(f.coef[i], f.se[i]));
        if (t < worstT) { worstT = t; worst = i; }
      }
      var pv = pValue(f.coef[worst], f.se[worst], f.dfe);
      if (!removeOK(worstT, pv)) return false;
      var removed = f.vars[worst - 1];
      inModel[removed] = false;
      steps.push(stepSummary(ds, "Removed", ds.names[removed], tStat(f.coef[worst], f.se[worst]), pv,
        fitCurrent(ds, inModel)));
      return true;
    }

    var guard = 0, changed;
    do {
      changed = false;
      if (method !== METHOD.BACKWARD && tryEnter()) changed = true;
      if (method !== METHOD.FORWARD) {
        while (tryRemove()) {
          changed = true;
          if (++guard > MAX_STEPS) break;
        }
      }
      guard++;
    } while (changed && guard <= MAX_STEPS);

    return { ok: true, inModel: inModel, steps: steps, hitLimit: guard > MAX_STEPS };
  }

  // ---------------------------------------------------------------------------
  // Report content (values only; formatting lives in the Excel writer)
  // ---------------------------------------------------------------------------

  function finalModel(ds, inModel) {
    var f = fitCurrent(ds, inModel);
    var k = f.vars.length;
    var ssr = ds.syy - f.sse;
    var tCrit = tInv2T(0.05, f.dfe);
    var fStat = k > 0 && f.sse > 0 ? (ssr / k) / (f.sse / f.dfe) : null;

    var coefficients = [];
    for (var i = 0; i <= k; i++) {
      var b = f.coef[i], s = f.se[i];
      coefficients.push({
        term: i === 0 ? "Intercept" : ds.names[f.vars[i - 1]],
        coef: b, se: s, tStat: tStat(b, s), pValue: pValue(b, s, f.dfe),
        lower: b - tCrit * s, upper: b + tCrit * s
      });
    }

    var excluded = [];
    for (var j = 0; j < ds.k; j++) {
      if (inModel[j]) continue;
      inModel[j] = true;
      var trial = fitCurrent(ds, inModel);
      inModel[j] = false;
      if (trial) {
        var pos = trial.vars.indexOf(j) + 1;
        excluded.push({ variable: ds.names[j], tStat: tStat(trial.coef[pos], trial.se[pos]),
          pValue: pValue(trial.coef[pos], trial.se[pos], trial.dfe), note: "" });
      } else {
        excluded.push({ variable: ds.names[j], tStat: null, pValue: null,
          note: ds.cxx[j][j] <= 0 ? "Constant column" : "Collinear with terms in the model or too few rows" });
      }
    }

    return {
      stats: {
        multipleR: Math.sqrt(Math.max(0, 1 - f.sse / ds.syy)),
        r2: 1 - f.sse / ds.syy,
        adjR2: 1 - (f.sse / f.dfe) / (ds.syy / (ds.n - 1)),
        rmse: Math.sqrt(f.sse / f.dfe),
        n: ds.n
      },
      anova: {
        regression: { df: k, ss: ssr, ms: k > 0 ? ssr / k : null, f: fStat,
          sigF: fStat === null ? null : fDistRT(fStat, k, f.dfe) },
        residual: { df: f.dfe, ss: f.sse, ms: f.sse / f.dfe },
        total: { df: ds.n - 1, ss: ds.syy }
      },
      coefficients: coefficients,
      excluded: excluded
    };
  }

  /*
   * One call for the task pane: validates, builds data, selects and summarizes.
   * opts: { yColumn, xColumns, hasLabels, yName, method, criterion, enter, remove }
   */
  function run(opts) {
    var method = opts.method, crit = opts.criterion;
    var d = defaultThresholds(crit);
    var enter = opts.enter == null ? d.enter : opts.enter;
    var remove = opts.remove == null ? d.remove : opts.remove;
    var msg = checkThresholds(method, crit, enter, remove);
    if (msg) return { ok: false, error: msg };

    var built = buildData(opts.yColumn, opts.xColumns, opts.hasLabels, opts.yName);
    if (!built.ok) return built;
    var ds = built.data;
    var sel = select(ds, method, crit, enter, remove);
    if (!sel.ok) return sel;
    return {
      ok: true,
      settings: { yName: ds.yName, method: method, criterion: crit, enter: enter, remove: remove,
        n: ds.n, dropped: built.dropped },
      steps: sel.steps,
      hitLimit: sel.hitLimit,
      model: finalModel(ds, sel.inModel),
      maxSteps: MAX_STEPS
    };
  }

  return {
    METHOD: METHOD,
    CRIT: CRIT,
    run: run,
    buildData: buildData,
    select: select,
    fit: fit,
    finalModel: finalModel,
    defaultThresholds: defaultThresholds,
    checkThresholds: checkThresholds,
    tDist2T: tDist2T,
    fDistRT: fDistRT,
    tInv2T: tInv2T
  };
});
