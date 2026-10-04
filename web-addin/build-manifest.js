// Writes docs/manifest.xml with the site's public address filled in.
// Usage: node build-manifest.js https://<github-user>.github.io/<repo>
const fs = require("node:fs");
const path = require("node:path");

const base = (process.argv[2] || "").replace(/\/+$/, "");
if (!/^https:\/\/[^\s]+$/.test(base)) {
  console.error("Usage: node build-manifest.js https://<github-user>.github.io/<repo>");
  process.exit(1);
}
const template = fs.readFileSync(path.join(__dirname, "manifest.template.xml"), "utf8");
fs.writeFileSync(path.join(__dirname, "docs", "manifest.xml"), template.split("{{BASE_URL}}").join(base));
console.log("Wrote docs/manifest.xml for " + base);
