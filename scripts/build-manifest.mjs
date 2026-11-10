// Writes dist/manifest.xml with the hosting URL filled in.
// BASE_URL comes from the environment (set by the GitHub Pages workflow); defaults to local dev.
import { readFileSync, writeFileSync, mkdirSync } from "node:fs";

const pkg = JSON.parse(readFileSync("package.json", "utf8"));
const base = (process.env.BASE_URL || "https://localhost:3000").replace(/\/+$/, "");
const origin = new URL(base).origin;
const [maj, min, pat] = pkg.version.split(".").map((n) => parseInt(n, 10) || 0);
const version = `${maj}.${min}.${pat}.${process.env.BUILD_NUMBER || 0}`;

const xml = readFileSync("manifest.template.xml", "utf8")
  .replaceAll("{{BASE_URL}}", base)
  .replaceAll("{{ORIGIN}}", origin)
  .replaceAll("{{VERSION}}", version);
if (/\{\{[A-Z_]+\}\}/.test(xml)) throw new Error("Unreplaced placeholder in manifest");
mkdirSync("dist", { recursive: true });
writeFileSync("dist/manifest.xml", xml);
writeFileSync("dist/.nojekyll", "");
console.log(`manifest.xml -> ${base} (v${version})`);
