/*
 * End-to-end check of the task pane in headless Chromium. Office.js is replaced
 * by a stub backed by test/excel-mock.js, so the real taskpane.html, taskpane.js,
 * stepwise.js and report.js run unchanged.
 *
 * Run: npm run test:e2e   (optional: SCREENSHOT=pane.png)
 */
const http = require("node:http");
const fs = require("node:fs");
const path = require("node:path");
const assert = require("node:assert");
const { chromium } = require("playwright");

const ROOT = path.join(__dirname, "..");
const TYPES = { ".html": "text/html", ".js": "text/javascript", ".css": "text/css", ".png": "image/png", ".json": "application/json" };

function serve() {
  return new Promise(resolve => {
    const server = http.createServer((req, res) => {
      const file = path.join(ROOT, decodeURIComponent(req.url.split("?")[0]));
      if (!file.startsWith(ROOT) || !fs.existsSync(file) || fs.statSync(file).isDirectory()) {
        res.writeHead(404); res.end(); return;
      }
      res.writeHead(200, { "Content-Type": TYPES[path.extname(file)] || "application/octet-stream" });
      fs.createReadStream(file).pipe(res);
    });
    server.listen(0, () => resolve(server));
  });
}

// Office.js stand-in: an Excel host whose workbook is an ExcelMock
function officeStub(data) {
  return fs.readFileSync(path.join(__dirname, "excel-mock.js"), "utf8") + `
    (function () {
      var data = ${JSON.stringify(data)};
      var wb = new ExcelMock.Workbook();
      var sh = wb.addSheet(data.sheet);
      sh.setData(data.topLeft, data.rows);
      window.__wb = wb;
      window.Office = {
        HostType: { Excel: "Excel" },
        context: { requirements: { isSetSupported: function () { return true; } } },
        onReady: function (cb) { setTimeout(function () { cb({ host: "Excel" }); }, 0); return Promise.resolve({ host: "Excel" }); }
      };
      window.Excel = { run: function (fn) { return fn(wb.context()); } };
    })();`;
}

function loadData() {
  const hoyda = path.join(__dirname, "hoyda.json");
  const d = fs.existsSync(hoyda)
    ? JSON.parse(fs.readFileSync(hoyda, "utf8"))
    : JSON.parse(fs.readFileSync(path.join(__dirname, "fixtures.json"), "utf8")).datasets[1];
  const header = ["SALARY"].concat(d.names);
  const rows = [header].concat(d.y.map((y, i) => [y].concat(d.cols.map(c => c[i]))));
  return { d, sheet: "Reg - 5 preds", topLeft: "A4", rows, last: 4 + d.y.length, isHoyda: fs.existsSync(hoyda) };
}

(async () => {
  const data = loadData();
  const server = await serve();
  const base = "http://localhost:" + server.address().port;
  let browser;
  try {
    browser = await chromium.launch();
  } catch (e) {
    browser = await chromium.launch({ executablePath: "/opt/pw-browsers/chromium" });
  }
  const page = await browser.newPage({ viewport: { width: 340, height: 720 } });
  const errors = [];
  page.on("pageerror", e => errors.push(e.message));
  await page.route("https://appsforoffice.microsoft.com/**", route =>
    route.fulfill({ contentType: "text/javascript", body: officeStub(data) }));

  await page.goto(base + "/docs/taskpane.html");
  await page.waitForFunction(() => window.Office && document.getElementById("run"));

  const lastCol = String.fromCharCode(65 + data.d.names.length);
  await page.evaluate(a => { window.__wb.selectedAddress = a; }, `'Reg - 5 preds'!A4:A${data.last}`);
  await page.click('[data-pick="yRange"]');
  await page.evaluate(a => { window.__wb.selectedAddress = a; }, `'Reg - 5 preds'!B4:${lastCol}${data.last}`);
  await page.click('[data-pick="xRange"]');
  assert.strictEqual(await page.inputValue("#yRange"), `'Reg - 5 preds'!A4:A${data.last}`);

  // Criterion switch resets thresholds and labels
  await page.selectOption("#criterion", "2");
  assert.strictEqual(await page.inputValue("#enter"), "0.05");
  assert.strictEqual(await page.textContent("#enterLabel"), "P-value to enter");
  await page.selectOption("#criterion", "1");
  assert.strictEqual(await page.inputValue("#remove"), "2");

  await page.click("#run");
  await page.waitForFunction(() => /Report written|wrong|must|Fewer|cannot/.test(document.getElementById("status").textContent));
  const status = await page.textContent("#status");
  assert.match(status, /Report written to sheet "Stepwise"/, status);

  const report = await page.evaluate(() => {
    const sh = window.__wb.sheet("Stepwise");
    return { cells: sh.cells, gridlines: sh.gridlines };
  });
  const values = Object.values(report.cells);
  assert.ok(values.includes("Reg - 5 preds!A4:A" + data.last), "Y range shown");
  assert.strictEqual(report.gridlines, false);
  if (data.isHoyda) {
    for (const v of ["SENIORITY", "SENIORITY2", "ROW4"]) assert.ok(values.includes(v), v);
    const r2 = values.find(v => typeof v === "number" && Math.abs(v - 0.582199) < 1e-5);
    assert.ok(r2 !== undefined, "final R Square 0.5822 present");
  }

  // A bad range gives a readable message
  await page.fill("#xRange", `'Reg - 5 preds'!B5:${lastCol}${data.last}`);
  await page.click("#run");
  await page.waitForFunction(() => document.getElementById("status").className === "error");
  assert.match(await page.textContent("#status"), /same rows as Y/);

  if (process.env.SCREENSHOT) {
    await page.fill("#xRange", `'Reg - 5 preds'!B4:${lastCol}${data.last}`);
    await page.click("#run");
    await page.waitForFunction(() => document.getElementById("status").className === "ok");
    await page.screenshot({ path: process.env.SCREENSHOT });
  }
  assert.deepStrictEqual(errors, []);
  await browser.close();
  server.close();
  console.log("e2e passed" + (data.isHoyda ? " (Hoyda data)" : " (synthetic data)"));
})().catch(e => { console.error(e); process.exit(1); });
