/**
 * Stub for Node.js 'node:zlib' module.
 *
 * just-bash's browser bundle imports gzipSync / gunzipSync / constants from
 * node:zlib for its gzip, gunzip, zcat, and rg commands. These Node-only APIs
 * don't exist in the browser, so we provide lightweight stubs.
 *
 * The gzip/gunzip/zcat commands will fail gracefully at runtime with a clear
 * error message instead of crashing the entire app at import time.
 */

function notSupported(name: string): never {
  throw new Error(`${name} is not available in the browser environment.`);
}

export function gzipSync(_data: unknown, _options?: unknown): never {
  notSupported("gzipSync");
}

export function gunzipSync(_data: unknown, _options?: unknown): never {
  notSupported("gunzipSync");
}

export function deflateSync(_data: unknown, _options?: unknown): never {
  notSupported("deflateSync");
}

export function inflateSync(_data: unknown, _options?: unknown): never {
  notSupported("inflateSync");
}

export const constants = {
  Z_NO_COMPRESSION: 0,
  Z_BEST_SPEED: 1,
  Z_BEST_COMPRESSION: 9,
  Z_DEFAULT_COMPRESSION: -1,
  Z_DEFAULT_STRATEGY: 0,
  Z_FILTERED: 1,
  Z_HUFFMAN_ONLY: 2,
  Z_RLE: 3,
  Z_FIXED: 4,
} as const;

export default {
  gzipSync,
  gunzipSync,
  deflateSync,
  inflateSync,
  constants,
};
