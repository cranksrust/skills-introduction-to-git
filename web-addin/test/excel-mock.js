/*
 * Minimal, strict mock of the Excel JavaScript API surface the add-in uses.
 * Setting an unknown property throws, so a misspelled API name fails tests.
 * Works in Node (module.exports) and in the browser (window.ExcelMock).
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.ExcelMock = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  var MAX_ROWS = 1048576, MAX_COLS = 16384;

  function colIndex(letters) {
    var n = 0;
    for (var i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
    return n - 1;
  }
  function colLetters(i) {
    var s = ""; i += 1;
    while (i > 0) { var m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); }
    return s;
  }

  // Parses A1, A1:B2, A:A, D:K, 3:3 into 0-based bounds
  function parseA1(a) {
    a = a.replace(/\$/g, "").toUpperCase();
    var parts = a.split(":");
    if (parts.length === 1) parts.push(parts[0]);
    function one(p, isEnd) {
      var m = /^([A-Z]*)(\d*)$/.exec(p);
      if (!m) throw new Error("InvalidArgument: bad address " + a);
      var c = m[1] ? colIndex(m[1]) : (isEnd ? MAX_COLS - 1 : 0);
      var r = m[2] ? parseInt(m[2], 10) - 1 : (isEnd ? MAX_ROWS - 1 : 0);
      return { r: r, c: c };
    }
    var s = one(parts[0], false), e = one(parts[1], true);
    return { r1: s.r, c1: s.c, r2: e.r, c2: e.c };
  }

  function strict(target, allowed, label) {
    return new Proxy(target, {
      set: function (t, k, v) {
        if (allowed.indexOf(k) < 0) throw new Error("Mock: unknown property " + label + "." + String(k));
        t[k] = v;
        return true;
      },
      get: function (t, k) {
        if (k in t || typeof k === "symbol") return t[k];
        throw new Error("Mock: unknown member " + label + "." + String(k));
      }
    });
  }

  function Sheet(wb, name) {
    this.wb = wb;
    this._name = name;
    this.cells = {};      // "r,c" -> value
    this.formats = {};    // "r,c" -> numberFormat
    this.style = {};      // "r,c" -> { key: value }
    this.colWidths = {};
    this.rowHeights = {};
    this.gridlines = true;
    this.active = false;
  }
  Sheet.prototype.cellStyle = function (r, c) {
    var k = r + "," + c;
    return this.style[k] || (this.style[k] = {});
  };
  Sheet.prototype.usedBounds = function () {
    var keys = Object.keys(this.cells).filter(function (k) { return this.cells[k] !== ""; }, this);
    if (!keys.length) return null;
    var b = { r1: Infinity, c1: Infinity, r2: -1, c2: -1 };
    keys.forEach(function (k) {
      var p = k.split(",").map(Number);
      b.r1 = Math.min(b.r1, p[0]); b.c1 = Math.min(b.c1, p[1]);
      b.r2 = Math.max(b.r2, p[0]); b.c2 = Math.max(b.c2, p[1]);
    });
    return b;
  };
  Sheet.prototype.setData = function (topLeft, rows) {
    var s = parseA1(topLeft);
    rows.forEach(function (row, i) {
      row.forEach(function (v, j) { this.cells[(s.r1 + i) + "," + (s.c1 + j)] = v; }, this);
    }, this);
  };

  function makeSheetProxy(sheet) {
    var api = {
      get name() { return sheet._name; },
      set name(v) { sheet._name = v; },
      get showGridlines() { return sheet.gridlines; },
      set showGridlines(v) { sheet.gridlines = v; },
      load: function () { return api; },
      activate: function () { sheet.wb.sheets.forEach(function (s) { s.active = false; }); sheet.active = true; },
      getRange: function (address) {
        return makeRange(sheet, address === undefined ? { r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 } : parseA1(address));
      },
      getUsedRange: function () {
        var b = sheet.usedBounds();
        return b ? makeRange(sheet, b) : makeRange(sheet, null);
      }
    };
    return strict(api, ["name", "showGridlines"], "Worksheet");
  }

  function borderProxy(sheet, b, edge) {
    var store = {};
    function apply(key, v) {
      for (var r = b.r1; r <= b.r2; r++) for (var c = b.c1; c <= b.c2; c++) {
        var onEdge = (edge === "EdgeBottom" && r === b.r2) || (edge === "EdgeTop" && r === b.r1) ||
          (edge === "EdgeLeft" && c === b.c1) || (edge === "EdgeRight" && c === b.c2);
        if (onEdge) sheet.cellStyle(r, c)["border." + edge + "." + key] = v;
      }
      store[key] = v;
    }
    var api = {};
    ["style", "weight", "color"].forEach(function (key) {
      Object.defineProperty(api, key, { get: function () { return store[key]; }, set: function (v) { apply(key, v); }, enumerable: true });
    });
    return strict(api, ["style", "weight", "color"], "RangeBorder");
  }

  function makeRange(sheet, b) {
    function each(fn) {
      if (!b) return;
      var maxR = Math.min(b.r2, 5000), maxC = Math.min(b.c2, 200);
      for (var r = b.r1; r <= maxR; r++) for (var c = b.c1; c <= maxC; c++) fn(r, c);
    }
    function styleSetter(key) {
      return function (v) { each(function (r, c) { sheet.cellStyle(r, c)[key] = v; }); };
    }
    function grid(fn) {
      var out = [];
      for (var r = b.r1; r <= b.r2; r++) {
        var row = [];
        for (var c = b.c1; c <= b.c2; c++) row.push(fn(r, c));
        out.push(row);
      }
      return out;
    }
    function checkDims(v, label) {
      if (!Array.isArray(v) || v.length !== b.r2 - b.r1 + 1 || v.some(function (row) { return row.length !== b.c2 - b.c1 + 1; })) {
        throw new Error("InvalidArgument: " + label + " dimensions do not match the range");
      }
    }

    var font = {}, fill = {};
    ["name", "size", "bold", "italic", "color"].forEach(function (k) {
      Object.defineProperty(font, k, { set: styleSetter("font." + k), get: function () { return undefined; }, enumerable: true });
    });
    Object.defineProperty(fill, "color", { set: styleSetter("fill.color"), get: function () { return undefined; }, enumerable: true });

    var format = {
      font: strict(font, ["name", "size", "bold", "italic", "color"], "RangeFont"),
      fill: strict(fill, ["color"], "RangeFill"),
      borders: { getItem: function (edge) { return borderProxy(sheet, b, edge); } },
      autofitColumns: function () {}
    };
    ["horizontalAlignment", "verticalAlignment", "wrapText"].forEach(function (k) {
      Object.defineProperty(format, k, { set: styleSetter(k), get: function () { return undefined; }, enumerable: true });
    });
    Object.defineProperty(format, "columnWidth", {
      set: function (v) { for (var c = b.c1; c <= b.c2; c++) sheet.colWidths[c] = v; }, get: function () { return undefined; }, enumerable: true
    });
    Object.defineProperty(format, "rowHeight", {
      set: function (v) { for (var r = b.r1; r <= Math.min(b.r2, 5000); r++) sheet.rowHeights[r] = v; }, get: function () { return undefined; }, enumerable: true
    });

    var api = {
      get isNullObject() { return !b; },
      get values() { return grid(function (r, c) { var v = sheet.cells[r + "," + c]; return v === undefined ? "" : v; }); },
      set values(v) {
        checkDims(v, "values");
        v.forEach(function (row, i) { row.forEach(function (x, j) { sheet.cells[(b.r1 + i) + "," + (b.c1 + j)] = x; }); });
      },
      set numberFormat(v) {
        checkDims(v, "numberFormat");
        v.forEach(function (row, i) { row.forEach(function (x, j) { sheet.formats[(b.r1 + i) + "," + (b.c1 + j)] = x; }); });
      },
      get numberFormat() { return grid(function (r, c) { return sheet.formats[r + "," + c] || "General"; }); },
      get rowIndex() { return b.r1; },
      get rowCount() { return b.r2 - b.r1 + 1; },
      get columnIndex() { return b.c1; },
      get columnCount() { return b.c2 - b.c1 + 1; },
      get address() {
        var q = /^[A-Za-z0-9_]+$/.test(sheet._name) ? sheet._name : "'" + sheet._name.replace(/'/g, "''") + "'";
        var a = colLetters(b.c1) + (b.r1 + 1);
        if (b.r2 !== b.r1 || b.c2 !== b.c1) a += ":" + colLetters(b.c2) + (b.r2 + 1);
        return q + "!" + a;
      },
      format: strict(format, ["horizontalAlignment", "verticalAlignment", "wrapText", "columnWidth", "rowHeight"], "RangeFormat"),
      load: function () { return api; },
      select: function () { sheet.wb.selection = api; },
      getIntersectionOrNullObject: function (other) {
        var o = other._bounds;
        if (!b || !o) return makeRange(sheet, null);
        var i = { r1: Math.max(b.r1, o.r1), c1: Math.max(b.c1, o.c1), r2: Math.min(b.r2, o.r2), c2: Math.min(b.c2, o.c2) };
        if (i.r1 > i.r2 || i.c1 > i.c2) return makeRange(sheet, null);
        return makeRange(sheet, i);
      },
      _bounds: b
    };
    return strict(api, ["values", "numberFormat"], "Range");
  }

  function Workbook() {
    this.sheets = [];
    this.selection = null;
    this.selectedAddress = null;
    var self = this;
    var worksheets = {
      get items() { return self.sheets.map(makeSheetProxy); },
      load: function () { return worksheets; },
      add: function (name) {
        if (self.sheets.some(function (s) { return s._name.toLowerCase() === name.toLowerCase(); })) {
          throw new Error("ItemAlreadyExists: " + name);
        }
        var s = new Sheet(self, name);
        self.sheets.push(s);
        return makeSheetProxy(s);
      },
      getItem: function (name) {
        var s = self.sheets.filter(function (x) { return x._name === name; })[0];
        if (!s) throw new Error("ItemNotFound: " + name);
        return makeSheetProxy(s);
      },
      getActiveWorksheet: function () {
        return makeSheetProxy(self.sheets.filter(function (s) { return s.active; })[0] || self.sheets[0]);
      }
    };
    this.api = {
      worksheets: worksheets,
      getSelectedRanges: function () { return { address: self.selectedAddress, load: function () { return this; } }; },
      getSelectedRange: function () { return { address: self.selectedAddress.split(",")[0], load: function () { return this; } }; }
    };
  }
  Workbook.prototype.addSheet = function (name) {
    var s = new Sheet(this, name);
    this.sheets.push(s);
    if (this.sheets.length === 1) s.active = true;
    return s;
  };
  Workbook.prototype.sheet = function (name) {
    return this.sheets.filter(function (s) { return s._name === name; })[0];
  };
  Workbook.prototype.context = function () {
    var self = this;
    return { workbook: self.api, sync: function () { return Promise.resolve(); } };
  };

  return { Workbook: Workbook, parseA1: parseA1, colLetters: colLetters };
});
