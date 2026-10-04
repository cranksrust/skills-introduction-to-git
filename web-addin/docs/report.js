/*
 * Builds the report layout from a Stepwise.run() result and writes it to a new
 * worksheet with the Excel JavaScript API. Layout and labels follow Excel's
 * Data Analysis > Regression report, with Standard Error shown as RMSE.
 *
 * buildLayout() is pure (no Office calls) so tests can check the content.
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    module.exports = factory();
  } else {
    root.StepwiseReport = factory();
  }
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  var COLORS = { orange: "#BF5700", tint: "#F8E5D6", grey: "#595959", rule: "#BFBFBF", white: "#FFFFFF" };
  var CREDIT = "HOUMBA Class of '28";
  var FMT = {
    p: '[<0.0001]"<0.0001";0.0000',
    t: "0.00",
    r2: "0.0000",
    num: "#,##0.00",
    coef: "#,##0.0000",
    int: "0",
    text: "@",
    general: "General"
  };
  var LAST_COL = 11; // A:K

  function colLetter(index) { // 0-based
    var s = "";
    index += 1;
    while (index > 0) {
      var m = (index - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      index = Math.floor((index - 1) / 26);
    }
    return s;
  }

  function addr(r, c1, c2) { // 0-based row and columns
    return colLetter(c1) + (r + 1) + (c2 === undefined || c2 === c1 ? "" : ":" + colLetter(c2) + (r + 1));
  }

  function blockAddr(r1, c1, r2, c2) {
    return colLetter(c1) + (r1 + 1) + ":" + colLetter(c2) + (r2 + 1);
  }

  /*
   * Returns { values, formats, ops, rows, cols } where values and formats are
   * rows x LAST_COL grids and ops is a list of formatting operations.
   */
  function buildLayout(result, ranges) {
    var values = [], formats = [], ops = [];
    var r = 0;

    function ensureRow(i) {
      while (values.length <= i) {
        values.push(new Array(LAST_COL).fill(""));
        formats.push(new Array(LAST_COL).fill(FMT.general));
      }
    }
    function set(row, col, v, fmt) {
      ensureRow(row);
      if (v === null || v === undefined || (typeof v === "number" && !isFinite(v))) {
        values[row][col] = "";
        return;
      }
      values[row][col] = v;
      formats[row][col] = typeof v === "string" ? FMT.text : (fmt || FMT.general);
    }
    function section(title) {
      set(r, 0, title);
      ops.push({ type: "section", address: addr(r, 0) });
      r++;
    }
    function pair(label, v, fmt) {
      set(r, 0, label);
      set(r, 1, v, fmt);
      ops.push({ type: "label", address: addr(r, 0) });
      ops.push({ type: "align", address: addr(r, 1), horizontal: "Left" });
      r++;
    }
    // fmts: one per column; "@" marks a left-aligned column
    function header(labels, fmts) {
      labels.forEach(function (h, i) { set(r, i, h); });
      ops.push({ type: "header", address: addr(r, 0, labels.length - 1) });
      fmts.forEach(function (f, i) {
        ops.push({ type: "align", address: addr(r, i), horizontal: f === FMT.text ? "Left" : "Right" });
      });
      r++;
      return r;
    }
    function dataRow(vals, fmts) {
      vals.forEach(function (v, i) { set(r, i, v, fmts[i] === FMT.text ? undefined : fmts[i]); });
      r++;
    }
    function alignTextColumns(top, bottom, fmts) {
      if (bottom < top) return;
      fmts.forEach(function (f, i) {
        if (f === FMT.text) ops.push({ type: "align", address: blockAddr(top, i, bottom, i), horizontal: "Left" });
      });
    }

    var s = result.settings;
    var methodName = s.method === 1 ? "Forward selection" : s.method === 2 ? "Backward elimination" : "Stepwise";
    var isT = s.criterion === 1;

    // Title bar
    set(0, 0, "  Stepwise Regression Output");
    ensureRow(0);
    ops.push({ type: "title", address: addr(0, 0, LAST_COL - 1) });
    r = 2;

    section("Settings");
    pair("Dependent variable", s.yName);
    pair("Method", methodName);
    pair("Criterion", isT ? "T-Stat" : "P-value");
    if (s.method !== 2) pair(isT ? "|T-Stat| to enter" : "P-value to enter", s.enter, FMT.t);
    if (s.method !== 1) pair(isT ? "|T-Stat| to remove" : "P-value to remove", s.remove, FMT.t);
    pair("Observations used", s.n, FMT.int);
    pair("Rows excluded (blank or non-numeric)", s.dropped, FMT.int);
    pair("Y range", ranges.y);
    pair("X range", ranges.x);
    if (result.hitLimit) {
      pair("Warning", "Stopped after " + result.maxSteps + " iterations; the procedure was cycling.");
      ops.push({ type: "warning", address: addr(r - 1, 1) });
    }
    r++;

    // Selection steps
    section("Selection steps");
    var stepFmts = [FMT.text, FMT.text, FMT.text, FMT.t, FMT.p, FMT.int, FMT.r2, FMT.r2, FMT.num, FMT.num, FMT.num];
    var top = header(["Step", "Action", "Variable", "T-Stat", "P-value", "Terms in model", "R Square",
      "Adjusted R Square", "RMSE", "AIC", "BIC"], stepFmts);
    result.steps.forEach(function (st, i) {
      dataRow([i, st.action, st.variable, st.tStat, st.pValue, st.terms, st.r2, st.adjR2, st.rmse, st.aic, st.bic],
        stepFmts);
    });
    alignTextColumns(top, r - 1, stepFmts);
    r++;

    // Final model: Regression Statistics
    var m = result.model;
    section("Final model");
    header(["Regression Statistics", ""], [FMT.text, FMT.text]);
    pair("Multiple R", m.stats.multipleR, FMT.r2);
    pair("R Square", m.stats.r2, FMT.r2);
    pair("Adjusted R Square", m.stats.adjR2, FMT.r2);
    pair("RMSE", m.stats.rmse, FMT.num);
    pair("Observations", m.stats.n, FMT.int);
    r++;

    // ANOVA
    section("ANOVA");
    var anovaFmts = [FMT.text, FMT.int, FMT.num, FMT.num, FMT.num, FMT.p];
    top = header(["", "df", "SS", "MS", "F", "Significance F"], anovaFmts);
    var a = m.anova;
    dataRow(["Regression", a.regression.df, a.regression.ss, a.regression.ms, a.regression.f, a.regression.sigF], anovaFmts);
    dataRow(["Residual", a.residual.df, a.residual.ss, a.residual.ms, null, null], anovaFmts);
    ops.push({ type: "ruleTop", address: addr(r, 0, 5) });
    dataRow(["Total", a.total.df, a.total.ss, null, null, null], anovaFmts);
    alignTextColumns(top, r - 1, anovaFmts);
    r++;

    // Coefficients
    var coefFmts = [FMT.text, FMT.coef, FMT.coef, FMT.t, FMT.p, FMT.coef, FMT.coef];
    top = header(["", "Coefficients", "Standard Error", "T-Stat", "P-value", "Lower 95%", "Upper 95%"], coefFmts);
    m.coefficients.forEach(function (c) {
      dataRow([c.term, c.coef, c.se, c.tStat, c.pValue, c.lower, c.upper], coefFmts);
    });
    alignTextColumns(top, r - 1, coefFmts);

    // Variables not in the final model
    if (m.excluded.length) {
      r++;
      section("Variables not in the final model");
      var exFmts = [FMT.text, FMT.t, FMT.p, FMT.text];
      top = header(["Variable", "T-Stat if added", "P-value if added", "Note"], exFmts);
      m.excluded.forEach(function (e) {
        dataRow([e.variable, e.tStat, e.pValue, e.note || null], exFmts);
      });
      alignTextColumns(top, r - 1, exFmts);
    }

    // Footer
    r++;
    set(r, 0, "Stepwise Regression add-in  |  " + CREDIT);
    ops.push({ type: "footerRule", address: addr(r, 0, LAST_COL - 1) });
    ops.push({ type: "footer", address: addr(r, 0) });
    ensureRow(r);

    return { values: values, formats: formats, ops: ops, rows: values.length, cols: LAST_COL };
  }

  function uniqueName(existing, base) {
    var lower = existing.map(function (n) { return n.toLowerCase(); });
    if (lower.indexOf(base.toLowerCase()) < 0) return base;
    for (var i = 2; ; i++) {
      var name = base + " (" + i + ")";
      if (lower.indexOf(name.toLowerCase()) < 0) return name;
    }
  }

  function applyOp(sheet, op) {
    var rg = sheet.getRange(op.address);
    var f = rg.format;
    switch (op.type) {
      case "title":
        f.fill.color = COLORS.orange;
        f.font.color = COLORS.white;
        f.font.bold = true;
        f.font.size = 14;
        f.verticalAlignment = "Center";
        f.rowHeight = 26;
        break;
      case "section":
        f.font.bold = true;
        f.font.size = 11;
        f.font.color = COLORS.orange;
        break;
      case "label":
        f.font.color = COLORS.grey;
        break;
      case "align":
        f.horizontalAlignment = op.horizontal;
        break;
      case "header":
        f.font.bold = true;
        f.fill.color = COLORS.tint;
        f.wrapText = true;
        f.verticalAlignment = "Bottom";
        var b = f.borders.getItem("EdgeBottom");
        b.style = "Continuous";
        b.weight = "Medium";
        b.color = COLORS.orange;
        break;
      case "ruleTop":
        var t = f.borders.getItem("EdgeTop");
        t.style = "Continuous";
        t.color = COLORS.rule;
        break;
      case "footerRule":
        var fr = f.borders.getItem("EdgeTop");
        fr.style = "Continuous";
        fr.weight = "Thin";
        fr.color = COLORS.orange;
        break;
      case "footer":
        f.font.italic = true;
        f.font.size = 8;
        f.font.color = COLORS.grey;
        break;
      case "warning":
        f.font.color = "#C00000";
        break;
    }
  }

  /*
   * Writes the report to a new sheet. context is an Excel.RequestContext.
   * Returns the new sheet's name.
   */
  async function writeReport(context, result, ranges) {
    var layout = buildLayout(result, ranges);
    var sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();

    var name = uniqueName(sheets.items.map(function (s) { return s.name; }), "Stepwise");
    var sheet = sheets.add(name);

    var all = sheet.getRange();
    all.format.font.name = "Calibri";
    all.format.font.size = 10;

    var block = sheet.getRange(blockAddr(0, 0, layout.rows - 1, layout.cols - 1));
    block.numberFormat = layout.formats; // before values, so text stays text
    block.values = layout.values;

    layout.ops.forEach(function (op) { applyOp(sheet, op); });

    // Fixed widths (points) so long range addresses overflow instead of widening column B
    sheet.getRange("A:A").format.columnWidth = 175;
    sheet.getRange("B:B").format.columnWidth = 75;
    sheet.getRange("C:C").format.columnWidth = 100;
    sheet.getRange("D:K").format.columnWidth = 68;
    sheet.showGridlines = false;
    sheet.activate();
    sheet.getRange("A1").select();
    await context.sync();
    return name;
  }

  return { buildLayout: buildLayout, writeReport: writeReport, uniqueName: uniqueName, colLetter: colLetter, FMT: FMT };
});
