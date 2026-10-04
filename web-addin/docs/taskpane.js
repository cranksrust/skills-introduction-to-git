/* global Office, Excel, Stepwise, StepwiseReport */
(function () {
  "use strict";

  var STORE_KEY = "stepwise-settings";
  var el = {};

  // ---------------------------------------------------------------------------
  // Address parsing
  // ---------------------------------------------------------------------------

  // Splits "Sheet1!A1:A5, 'My Sheet'!C1:C5" on commas outside quotes
  function splitAreas(text) {
    var parts = [], cur = "", inQuote = false;
    for (var i = 0; i < text.length; i++) {
      var ch = text[i];
      if (ch === "'") inQuote = !inQuote;
      if (ch === "," && !inQuote) { parts.push(cur); cur = ""; continue; }
      cur += ch;
    }
    parts.push(cur);
    return parts.map(function (p) { return p.trim(); }).filter(Boolean);
  }

  function parseArea(text) {
    var bang = text.lastIndexOf("!");
    if (bang < 0) return { sheet: null, local: text.replace(/\$/g, "") };
    var sheet = text.slice(0, bang);
    if (sheet.charAt(0) === "'" && sheet.charAt(sheet.length - 1) === "'") {
      sheet = sheet.slice(1, -1).replace(/''/g, "'");
    }
    return { sheet: sheet, local: text.slice(bang + 1).replace(/\$/g, "") };
  }

  // ---------------------------------------------------------------------------
  // Reading ranges
  // ---------------------------------------------------------------------------

  // Queues loads for every area in an address list; call context.sync() after
  function queueAreas(context, text) {
    return splitAreas(text).map(function (part) {
      var a = parseArea(part);
      var sheet = a.sheet ? context.workbook.worksheets.getItem(a.sheet) : context.workbook.worksheets.getActiveWorksheet();
      // Clip whole-column selections such as A:A to the sheet's used range
      var rg = sheet.getRange(a.local).getIntersectionOrNullObject(sheet.getUsedRange());
      rg.load("values, rowIndex, rowCount, columnIndex, columnCount, address");
      sheet.load("name");
      return { text: part, range: rg, sheet: sheet };
    });
  }

  function displayAddress(area) {
    var a = area.range.address;
    return area.sheet.name + "!" + a.slice(a.lastIndexOf("!") + 1);
  }

  function column(values, c) {
    return values.map(function (row) { return row[c]; });
  }

  function overlaps(a, b) {
    if (a.sheet.name !== b.sheet.name) return false;
    var ra = a.range, rb = b.range;
    var colsOverlap = ra.columnIndex < rb.columnIndex + rb.columnCount && rb.columnIndex < ra.columnIndex + ra.columnCount;
    var rowsOverlap = ra.rowIndex < rb.rowIndex + rb.rowCount && rb.rowIndex < ra.rowIndex + ra.rowCount;
    return colsOverlap && rowsOverlap;
  }

  function fail(msg) { var e = new Error(msg); e.userFacing = true; throw e; }

  // ---------------------------------------------------------------------------
  // UI
  // ---------------------------------------------------------------------------

  function setStatus(msg, kind) {
    el.status.textContent = msg || "";
    el.status.className = kind || "";
  }

  function numberOrNull(input) {
    var v = parseFloat(input.value);
    return isNaN(v) ? null : v;
  }

  function updateThresholdUI(resetValues) {
    var isT = el.criterion.value === "1";
    var method = el.method.value;
    el.enterLabel.textContent = isT ? "|T-Stat| to enter" : "P-value to enter";
    el.removeLabel.textContent = isT ? "|T-Stat| to remove" : "P-value to remove";
    if (resetValues) {
      var d = Stepwise.defaultThresholds(isT ? 1 : 2);
      el.enter.value = d.enter;
      el.remove.value = d.remove;
    }
    el.enter.disabled = method === "2";
    el.remove.disabled = method === "1";
  }

  function saveSettings() {
    try {
      localStorage.setItem(STORE_KEY, JSON.stringify({
        yRange: el.yRange.value, xRange: el.xRange.value, hasLabels: el.hasLabels.checked,
        method: el.method.value, criterion: el.criterion.value, enter: el.enter.value, remove: el.remove.value
      }));
    } catch (e) { /* storage unavailable */ }
  }

  function loadSettings() {
    try {
      var s = JSON.parse(localStorage.getItem(STORE_KEY) || "null");
      if (!s) return;
      el.yRange.value = s.yRange || "";
      el.xRange.value = s.xRange || "";
      el.hasLabels.checked = s.hasLabels !== false;
      el.method.value = s.method || "3";
      el.criterion.value = s.criterion || "1";
      el.enter.value = s.enter;
      el.remove.value = s.remove;
    } catch (e) { /* storage unavailable */ }
  }

  async function useSelection(targetId) {
    try {
      await Excel.run(async function (context) {
        var text;
        if (Office.context.requirements.isSetSupported("ExcelApi", "1.9")) {
          var areas = context.workbook.getSelectedRanges();
          areas.load("address");
          await context.sync();
          text = areas.address;
        } else {
          var rg = context.workbook.getSelectedRange();
          rg.load("address");
          await context.sync();
          text = rg.address;
        }
        el[targetId].value = text.split(",").join(", ");
      });
      setStatus("");
    } catch (e) {
      setStatus("Could not read the selection: " + e.message, "error");
    }
  }

  async function run() {
    var yText = el.yRange.value.trim();
    var xText = el.xRange.value.trim();
    if (!yText || !xText) { setStatus("Enter or select both a Y range and an X range.", "error"); return; }
    saveSettings();
    el.run.disabled = true;
    setStatus("Running...");

    try {
      var sheetName = await Excel.run(async function (context) {
        var yAreas = queueAreas(context, yText);
        var xAreas = queueAreas(context, xText);
        await context.sync();

        if (yAreas.length !== 1) fail("Y must be a single column.");
        var y = yAreas[0];
        if (y.range.isNullObject) fail("The Y range is empty.");
        if (y.range.columnCount !== 1) fail("Y must be a single column.");

        var xColumns = [];
        xAreas.forEach(function (a) {
          if (a.range.isNullObject) fail("The X range " + a.text + " is empty.");
          if (a.range.rowIndex !== y.range.rowIndex || a.range.rowCount !== y.range.rowCount) {
            fail("Every X column must start and end on the same rows as Y (" + displayAddress(y) + ").");
          }
          if (overlaps(a, y)) fail("The Y column must not be part of the X selection.");
          for (var c = 0; c < a.range.columnCount; c++) {
            xColumns.push({
              values: column(a.range.values, c),
              name: "Col " + StepwiseReport.colLetter(a.range.columnIndex + c)
            });
          }
        });

        var result = Stepwise.run({
          yColumn: column(y.range.values, 0),
          xColumns: xColumns,
          hasLabels: el.hasLabels.checked,
          yName: "Col " + StepwiseReport.colLetter(y.range.columnIndex),
          method: parseInt(el.method.value, 10),
          criterion: parseInt(el.criterion.value, 10),
          enter: numberOrNull(el.enter),
          remove: numberOrNull(el.remove)
        });
        if (!result.ok) fail(result.error);

        return StepwiseReport.writeReport(context, result, {
          y: displayAddress(y),
          x: xAreas.map(displayAddress).join(", ")
        });
      });
      setStatus("Report written to sheet \"" + sheetName + "\".", "ok");
    } catch (e) {
      setStatus(e.userFacing ? e.message : "Something went wrong: " + e.message, "error");
    } finally {
      el.run.disabled = false;
    }
  }

  function init() {
    ["yRange", "xRange", "hasLabels", "method", "criterion", "enter", "remove", "enterLabel",
      "removeLabel", "run", "status"].forEach(function (id) { el[id] = document.getElementById(id); });

    loadSettings();
    updateThresholdUI(false);

    el.criterion.addEventListener("change", function () { updateThresholdUI(true); });
    el.method.addEventListener("change", function () { updateThresholdUI(false); });
    el.run.addEventListener("click", run);
    document.querySelectorAll("[data-pick]").forEach(function (btn) {
      btn.addEventListener("click", function () { useSelection(btn.getAttribute("data-pick")); });
    });
  }

  Office.onReady(function (info) {
    if (info.host === Office.HostType.Excel) {
      init();
    } else {
      document.getElementById("status").textContent = "Open this add-in from Excel.";
    }
  });
})();
