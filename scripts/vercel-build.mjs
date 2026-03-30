/**
 * Custom Vercel build script — Build Output API v3.
 *
 * Vercel's Vite framework detection treats the project as static-only and
 * ignores the api/ directory for edge functions.  Using framework:null
 * doesn't reliably discover api/ functions either in CLI v50+.
 *
 * This script produces the Build Output API v3 directory structure
 * (.vercel/output/) so the deployment is fully deterministic:
 *
 *   .vercel/output/
 *     config.json          — routing (rewrites, headers)
 *     static/              — files from dist/
 *     functions/
 *       api/gemini/
 *         [...path].func/  — edge function (Gemini proxy)
 */

import { execSync } from "node:child_process";
import { cpSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";

const OUTPUT = ".vercel/output";

// ── 1. Vite build ───────────────────────────────────────────────────────────
console.log("[vercel-build] Running vite build…");
execSync("npx vite build", { stdio: "inherit" });

// ── 2. Create Build Output API directory structure ──────────────────────────
mkdirSync(`${OUTPUT}/static`, { recursive: true });
const funcDir = `${OUTPUT}/functions/api/gemini/[...path].func`;
mkdirSync(funcDir, { recursive: true });

// ── 3. Copy static files from dist/ ────────────────────────────────────────
console.log("[vercel-build] Copying static files…");
cpSync("dist", `${OUTPUT}/static`, { recursive: true });

// ── 4. Bundle edge function ─────────────────────────────────────────────────
console.log("[vercel-build] Bundling edge function…");
execSync(
  [
    "npx esbuild",
    '"api/gemini/[...path].ts"',
    `--outfile="${funcDir}/index.js"`,
    "--bundle",
    "--format=esm",
    "--target=es2022",
    "--platform=browser",
  ].join(" "),
  { stdio: "inherit" },
);

// ── 5. Edge function config ─────────────────────────────────────────────────
writeFileSync(
  `${funcDir}/.vc-config.json`,
  JSON.stringify({ runtime: "edge", entrypoint: "index.js" }, null, 2) + "\n",
);

// ── 6. Build Output API config (routing) ────────────────────────────────────
// Convert vercel.json rewrites + headers to Build Output API routes.
const vercelJson = JSON.parse(readFileSync("vercel.json", "utf-8"));
const routes = [];

// Headers → route with "headers" + "continue: true"
for (const h of vercelJson.headers ?? []) {
  const headers = {};
  for (const { key, value } of h.headers) {
    headers[key] = value;
  }
  routes.push({ src: h.source, headers, continue: true });
}

// Rewrites → route with "dest"
for (const r of vercelJson.rewrites ?? []) {
  routes.push({ src: r.source, dest: r.destination });
}

const config = { version: 3, routes };
writeFileSync(`${OUTPUT}/config.json`, JSON.stringify(config, null, 2) + "\n");

console.log("[vercel-build] Build Output API structure ready.");
console.log(`[vercel-build]   static:    ${OUTPUT}/static/`);
console.log(`[vercel-build]   functions: ${funcDir}/`);
