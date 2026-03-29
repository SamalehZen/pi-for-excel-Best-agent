import { promises as fs } from "node:fs";
import path from "node:path";

const ROOT = process.cwd();

const LANDING_FILE = path.join(ROOT, "public", "index.html");
const INSTALL_DOC_FILE = path.join(ROOT, "docs", "install.md");

const LANDING_REQUIRED_SNIPPETS = [
  "npx pi-for-excel-proxy",
  "curl -fsSL https://piforexcel.com/proxy | sh",
  "Built-in multi-provider web search and page fetch",
  "Extensions can call LLMs, fetch HTTP data, persist storage, and register tools through the built-in bridge.",
];

const LANDING_FORBIDDEN_SNIPPETS = [
  "git clone https://github.com/SamalehZen/pi-for-excel-Best-agent.git ~/.pi-for-excel-Best-agent",
  "mkcert -install && mkcert localhost",
  "cd ~/.pi-for-excel && npm run proxy:https",
];

const INSTALL_DOC_REQUIRED_SNIPPETS = [
  "npx pi-for-excel-proxy",
  "curl -fsSL https://piforexcel.com/proxy | sh",
];

function rel(filePath) {
  return path.relative(ROOT, filePath);
}

function collectMissing(source, snippets) {
  return snippets.filter((snippet) => !source.includes(snippet));
}

function collectPresent(source, snippets) {
  return snippets.filter((snippet) => source.includes(snippet));
}

async function main() {
  const [landingSource, installSource] = await Promise.all([
    fs.readFile(LANDING_FILE, "utf8"),
    fs.readFile(INSTALL_DOC_FILE, "utf8"),
  ]);

  const landingMissing = collectMissing(landingSource, LANDING_REQUIRED_SNIPPETS);
  const landingForbiddenPresent = collectPresent(landingSource, LANDING_FORBIDDEN_SNIPPETS);
  const installMissing = collectMissing(installSource, INSTALL_DOC_REQUIRED_SNIPPETS);

  const hasErrors =
    landingMissing.length > 0 ||
    landingForbiddenPresent.length > 0 ||
    installMissing.length > 0;

  if (hasErrors) {
    console.error("\n✗ Landing/install connect copy drift detected.\n");

    if (landingMissing.length > 0) {
      console.error(`${rel(LANDING_FILE)} is missing required snippets:`);
      for (const snippet of landingMissing) {
        console.error(`  - ${snippet}`);
      }
      console.error("");
    }

    if (landingForbiddenPresent.length > 0) {
      console.error(`${rel(LANDING_FILE)} still contains forbidden legacy snippets:`);
      for (const snippet of landingForbiddenPresent) {
        console.error(`  - ${snippet}`);
      }
      console.error("");
    }

    if (installMissing.length > 0) {
      console.error(`${rel(INSTALL_DOC_FILE)} is missing required snippets:`);
      for (const snippet of installMissing) {
        console.error(`  - ${snippet}`);
      }
      console.error("");
    }

    process.exitCode = 1;
    return;
  }

  console.log("✓ Landing/install connect copy check passed.");
}

void main();
