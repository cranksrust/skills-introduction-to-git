// Run: node --test test/
const test = require("node:test");
const assert = require("node:assert");
const fs = require("node:fs");
const path = require("node:path");
const S = require("../docs/stepwise.js");

const fixtures = JSON.parse(fs.readFileSync(path.join(__dirname, "fixtures.json"), "utf8"));

function close(a, b, rel = 1e-8) {
  if (a === null || b === null) return a === b;
  return Math.abs(a - b) <= rel * Math.max(1, Math.abs(b));
}

function runCase(d, method, crit) {
  return S.run({
    yColumn: d.y,
    xColumns: d.cols.map((values, j) => ({ values, name: d.names[j] })),
    hasLabels: false, yName: "y", method, criterion: crit
  });
}

function checkDataset(d, label) {
  for (const key of Object.keys(d.expected)) {
    const [method, crit] = key.split("-").map(Number);
    const exp = d.expected[key];
    test(`${label} method ${method} criterion ${crit}`, () => {
      const res = runCase(d, method, crit);
      if (exp.error) { assert.strictEqual(res.ok, false); return; }
      assert.ok(res.ok, res.error);
      assert.deepStrictEqual(res.model.coefficients.slice(1).map(c => c.term), exp.chosen);
      res.model.coefficients.forEach((c, i) => {
        assert.ok(close(c.coef, exp.coef[i]), `coef ${i}: ${c.coef} vs ${exp.coef[i]}`);
        assert.ok(close(c.se, exp.se[i]), `se ${i}`);
      });
      assert.strictEqual(res.steps.length, exp.steps.length);
      res.steps.forEach((s, i) => {
        assert.strictEqual(s.action === "Start" ? "Start" : s.variable, exp.steps[i][0] === "Start" ? "Start" : exp.steps[i][1]);
        if (exp.steps[i][3] !== null) assert.ok(close(s.pValue, exp.steps[i][3], 1e-6), `step p ${i}`);
        if (exp.steps[i][2] !== null) assert.ok(close(s.tStat, exp.steps[i][2], 1e-8), `step t ${i}`);
      });
    });
  }
}

fixtures.datasets.forEach((d, i) => checkDataset(d, `synthetic ${i}`));

test("distributions match scipy", () => {
  for (const c of fixtures.dist) {
    if (c.kind === "t2") assert.ok(close(S.tDist2T(c.x, c.df), c.p, 1e-9), `t2 ${c.x} ${c.df}`);
    if (c.kind === "f") assert.ok(close(S.fDistRT(c.x, c.d1, c.d2), c.p, 1e-9), `f ${c.x}`);
    if (c.kind === "tinv") assert.ok(close(S.tInv2T(0.05, c.df), c.t, 1e-9), `tinv ${c.df}`);
  }
});

test("labels, blanks and text rows are handled", () => {
  const y = ["SALES", 1, 2, "", 4, 5, 7, 8];
  const x = ["AD", 1, 2, 3, "n/a", 5, 6.5, 8.2];
  const res = S.run({ yColumn: y, xColumns: [{ values: x, name: "Col B" }], hasLabels: true,
    yName: "Col A", method: 1, criterion: 1 });
  assert.ok(res.ok, res.error);
  assert.strictEqual(res.settings.yName, "SALES");
  assert.strictEqual(res.settings.n, 5);
  assert.strictEqual(res.settings.dropped, 2);
  assert.strictEqual(res.model.coefficients[1].term, "AD");
});

test("threshold validation", () => {
  assert.match(S.checkThresholds(3, 1, 1.5, 2), /cycle/);
  assert.match(S.checkThresholds(3, 2, 0.2, 0.1), /cycle/);
  assert.strictEqual(S.checkThresholds(3, 1, 2, 2), "");
});

const hoyda = path.join(__dirname, "hoyda.json");
if (fs.existsSync(hoyda)) {
  const d = JSON.parse(fs.readFileSync(hoyda, "utf8"));
  checkDataset(d, "hoyda");
  test("hoyda matches the workbook and report values", () => {
    const res = runCase(d, 3, 1);
    assert.deepStrictEqual(res.model.coefficients.map(c => c.term), ["Intercept", "SENIORITY", "SENIORITY2", "ROW4"]);
    assert.ok(close(res.model.stats.r2, 0.582199, 1e-5));
    assert.ok(close(res.model.stats.rmse, 39.879, 1e-4));
    assert.ok(close(res.model.anova.regression.f, 18.5798, 1e-4));
  });
}
