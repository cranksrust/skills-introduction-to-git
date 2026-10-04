// Run: node --test test/*.test.js
const test = require("node:test");
const assert = require("node:assert");
const fs = require("node:fs");
const path = require("node:path");
const S = require("../docs/stepwise.js");
const R = require("../docs/report.js");
const { Workbook } = require("./excel-mock.js");

const fixtures = JSON.parse(fs.readFileSync(path.join(__dirname, "fixtures.json"), "utf8"));

// Column x2 is an exact combination of x0 and x1; backward needs it left out
function sampleResult(method = 3, crit = 1) {
  const d = fixtures.datasets[0];
  const cols = d.cols.map((values, j) => ({ values: [d.names[j]].concat(values), name: "Col " + j }))
    .filter((c, j) => method !== 2 || j !== 2);
  return S.run({
    yColumn: ["Y"].concat(d.y),
    xColumns: cols,
    hasLabels: true, yName: "Col A", method, criterion: crit
  });
}

function findRow(sheet, text, col) {
  for (const k of Object.keys(sheet.cells)) {
    const [r, c] = k.split(",").map(Number);
    if (sheet.cells[k] === text && (col === undefined || c === col)) return r;
  }
  return -1;
}

test("report writes the expected sections and values", async () => {
  const wb = new Workbook();
  wb.addSheet("Data");
  const result = sampleResult();
  assert.ok(result.ok, result.error);
  const name = await R.writeReport(wb.context(), result, { y: "Data!A1:A61", x: "Data!B1:G61" });
  assert.strictEqual(name, "Stepwise");
  const sh = wb.sheet("Stepwise");

  assert.strictEqual(sh.cells["0,0"], "  Stepwise Regression Output");
  for (const label of ["Settings", "Selection steps", "Final model", "Regression Statistics", "ANOVA",
    "Coefficients", "Standard Error", "T-Stat", "RMSE", "Intercept"]) {
    assert.ok(findRow(sh, label) >= 0, "missing " + label);
  }
  // Final model numbers match the engine
  const rmseRow = findRow(sh, "RMSE", 0);
  const rmseCells = Object.keys(sh.cells).filter(k => k.startsWith(rmseRow + ",") && typeof sh.cells[k] === "number");
  assert.ok(Math.abs(sh.cells[rmseCells[0]] - result.model.stats.rmse) < 1e-12);
  const icRow = findRow(sh, "Intercept");
  assert.ok(Math.abs(sh.cells[icRow + ",1"] - result.model.coefficients[0].coef) < 1e-12);
  assert.strictEqual(sh.formats[icRow + ",4"], '[<0.0001]"<0.0001";0.0000');
  // Footer credit
  assert.ok(Object.values(sh.cells).includes("Stepwise Regression add-in  |  HOUMBA Class of '28"));
  // Styling
  assert.strictEqual(sh.cellStyle(0, 10)["fill.color"], "#BF5700");
  assert.strictEqual(sh.cellStyle(findRow(sh, "Settings"), 0)["font.color"], "#BF5700");
  assert.strictEqual(sh.gridlines, false);
  assert.ok(sh.active);
  // Text cells are stored as text
  assert.strictEqual(sh.formats[findRow(sh, "Method") + ",1"], "@");
});

test("second report gets a unique sheet name", async () => {
  const wb = new Workbook();
  wb.addSheet("Data");
  wb.addSheet("Stepwise");
  const name = await R.writeReport(wb.context(), sampleResult(1, 2), { y: "a", x: "b" });
  assert.strictEqual(name, "Stepwise (2)");
});

test("backward report lists only the remove threshold", () => {
  const result = sampleResult(2, 1);
  assert.ok(result.ok, result.error);
  const layout = R.buildLayout(result, { y: "a", x: "b" });
  const flat = layout.values.flat();
  assert.ok(flat.includes("|T-Stat| to remove"));
  assert.ok(!flat.includes("|T-Stat| to enter"));
  assert.ok(flat.includes("(all candidates)"));
});
