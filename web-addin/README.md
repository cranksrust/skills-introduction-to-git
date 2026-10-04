# Stepwise Regression for Excel

An Office web add-in that runs forward selection, backward elimination and stepwise regression in Excel for Windows, Mac and the web. It writes a report in the layout of Excel's Data Analysis > Regression tool, with Standard Error shown as RMSE. It uses no VBA, so students never touch macro security or the Trust Center.

## How it works

The add-in is a small static website. Excel loads `docs/taskpane.html` in a side panel, and everything runs in the browser inside Excel: no server, no data leaves the workbook. GitHub Pages hosts it for free.

| Path | Purpose |
|---|---|
| `docs/` | The published site: task pane, engine, report writer, icons, install/help/privacy pages |
| `docs/stepwise.js` | Regression engine and t/F distributions, no Office dependency |
| `docs/report.js` | Builds the report layout and writes it with the Excel JavaScript API |
| `docs/taskpane.*` | Side panel UI |
| `manifest.template.xml` | Add-in manifest with a `{{BASE_URL}}` placeholder |
| `build-manifest.js` | Writes `docs/manifest.xml` for the site's address |
| `test/` | Engine, report and browser end-to-end tests |

## One-time setup

1. **Repository and Pages.** Push this folder to the root of a public GitHub repository. In the repository, go to **Settings > Pages**, set **Source** to **Deploy from a branch**, branch **main**, folder **/docs**, and save. The site appears at `https://<github-user>.github.io/<repo>/` within a minute or two.
2. **Manifest.** Run `node build-manifest.js https://<github-user>.github.io/<repo>`, then commit and push `docs/manifest.xml`.
3. **Validate.** Run `npx office-addin-manifest validate docs/manifest.xml`. It should report the manifest as valid.
4. **Try it.** Open `https://<github-user>.github.io/<repo>/`, download `manifest.xml`, and follow the Excel for the web steps on that page.

## Publishing to the Office Add-ins store (AppSource)

1. **Publisher account.** Create a Microsoft Partner Center account and enroll in the Microsoft 365 and Copilot program. Use an account the program will keep, since the listing lives there. Verification can take several days.
2. **New offer.** In Partner Center, create a new **Office add-in** offer named **Stepwise Regression** and upload `docs/manifest.xml`.
3. **Listing details:**
   - Short and long description (start from the manifest description and the install page).
   - Logo: `docs/assets/icon-300.png`.
   - Screenshots at 1366x768: the task pane next to a finished report works well.
   - Privacy policy URL: `https://<github-user>.github.io/<repo>/privacy.html`
   - Support URL: `https://<github-user>.github.io/<repo>/support.html`
   - Category: Data Analytics.
4. **Notes for certification.** Tell the testers it needs no sign-in, and give them a short data set and steps, for example: "Put numbers in A1:D20 with headers in row 1, open Stepwise Regression from the Home tab, use A1:A20 as Y and B1:D20 as X, click Run regression."
5. **Submit.** Review usually takes several business days. Fix anything the report flags and resubmit.

Code changes in `docs/` go live as soon as they are pushed, with no store review. Changes to the manifest (name, icons, buttons, permissions, version) need a new submission, and the `<Version>` number must go up.

## Development

```
npm install            # only needed for the browser end-to-end test
npm test               # engine and report tests (Node 18+)
npm run test:e2e       # task pane in headless Chromium against a mock workbook
```

`test/make_fixtures.py` regenerates `test/fixtures.json` from the Python reference implementation in the VBA project (`excel-stepwise-regression/tests`) and scipy.

## Handoff

The add-in keeps working only while the GitHub repository and Pages site stay up and the Partner Center account stays active. Before the maintainer leaves, transfer the repository (**Settings > Transfer ownership**) and the Partner Center offer to the next owner.
