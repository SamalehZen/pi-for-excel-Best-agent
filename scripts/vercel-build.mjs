/**
 * Custom Vercel build script — Build Output API v3.
 *
 * Vercel's Vite framework detection treats the project as static-only and
 * ignores the api/ directory for edge functions.  The Build Output API gives
 * deterministic control over the deployment structure.
 *
 *   .vercel/output/
 *     config.json          — routing (rewrites, headers, function route)
 *     static/              — files from dist/
 *     functions/
 *       api/
 *         gemini-proxy.func/  — edge function (Gemini API proxy)
 */

import { execSync } from "node:child_process";
import { cpSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";

const OUTPUT = ".vercel/output";

// ── 1. Vite build ───────────────────────────────────────────────────────────
console.log("[vercel-build] Running vite build…");
execSync("npx vite build", { stdio: "inherit" });

// ── 2. Create Build Output API directory structure ──────────────────────────
mkdirSync(`${OUTPUT}/static`, { recursive: true });
const funcDir = `${OUTPUT}/functions/api/gemini-proxy.func`;
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
//
// Route processing order (Build Output API v3):
//   1. Header rules (continue: true — apply headers, keep matching)
//   2. Gemini proxy rewrite: /api/gemini/* → /api/gemini-proxy?proxyPath=...
//      (placed BEFORE filesystem so the rewritten URL is resolved by step 3)
//   3. { handle: "filesystem" } — resolve static files AND functions
//   4. { handle: "miss" }       — enter "miss" phase for unmatched paths
//   5. Other rewrites (proxy, oauth callbacks)
const vercelJson = JSON.parse(readFileSync("vercel.json", "utf-8"));
const routes = [];

// Phase 1: Headers
for (const h of vercelJson.headers ?? []) {
  const headers = {};
  for (const { key, value } of h.headers) {
    headers[key] = value;
  }
  routes.push({ src: h.source, headers, continue: true });
}

// Phase 2: Gemini proxy rewrite — BEFORE filesystem so the rewritten URL
// (/api/gemini-proxy) is resolved by the filesystem phase.
routes.push({ src: "/api/gemini/(.+)", dest: "/api/gemini-proxy?proxyPath=$1" });
routes.push({ src: "/api/gemini/?$", dest: "/api/gemini-proxy" });

// Phase 3: Filesystem lookup (static files + functions)
routes.push({ handle: "filesystem" });

// Phase 4: Miss — routes below only run when filesystem has no match.
routes.push({ handle: "miss" });

// Phase 5: Other rewrites (check: true required after handle:miss)
for (const r of vercelJson.rewrites ?? []) {
  routes.push({ src: r.source, dest: r.destination, check: true });
}

const config = { version: 3, routes };
writeFileSync(`${OUTPUT}/config.json`, JSON.stringify(config, null, 2) + "\n");

console.log("[vercel-build] Build Output API structure ready.");
console.log(`[vercel-build]   static:    ${OUTPUT}/static/`);
console.log(`[vercel-build]   functions: ${funcDir}/`);
