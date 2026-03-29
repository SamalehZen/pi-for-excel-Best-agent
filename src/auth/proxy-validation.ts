/**
 * Proxy URL validation for Office taskpanes.
 *
 * Office add-ins are served over HTTPS. Some Office webviews (notably WKWebView on macOS)
 * will block calls to an HTTP proxy from an HTTPS taskpane (mixed content), surfacing as
 * "Load failed" / "Connection error".
 */

export const DEFAULT_LOCAL_PROXY_URL = "https://localhost:3003";

/**
 * Target URL used for proxy reachability probes.
 *
 * Must stay inside scripts/cors-proxy-server.mjs DEFAULT_ALLOWED_TARGET_HOSTS,
 * otherwise the helper will return 403 and OAuth preflight checks will fail.
 */
export const PROXY_REACHABILITY_TARGET_URL = "https://github.com";

export const PROXY_HELPER_DOCS_URL =
  "https://github.com/SamalehZen/pi-for-excel-Best-agent/blob/main/docs/install.md#oauth-logins-and-cors-proxy";

export function normalizeProxyUrl(url: string): string {
  return url.trim().replace(/\/+$/, "");
}

/**
 * Resolve a user-configured proxy URL with sane defaults.
 */
export function resolveConfiguredProxyUrl(rawUrl: unknown): string {
  const trimmed = typeof rawUrl === "string" ? rawUrl.trim() : "";
  const candidate = trimmed.length > 0 ? trimmed : DEFAULT_LOCAL_PROXY_URL;
  return normalizeProxyUrl(candidate);
}

/**
 * Probe whether a proxy URL is reachable and can forward to the allowlisted
 * reachability target.
 */
export async function probeProxyReachability(
  proxyUrl: string,
  timeoutMs: number = 1500,
): Promise<boolean> {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), timeoutMs);

  try {
    const url = `${normalizeProxyUrl(proxyUrl)}/?url=${encodeURIComponent(PROXY_REACHABILITY_TARGET_URL)}`;
    const resp = await fetch(url, { signal: controller.signal });
    return resp.ok;
  } catch {
    return false;
  } finally {
    clearTimeout(timeout);
  }
}

export function validateOfficeProxyUrl(url: string): string {
  const normalized = normalizeProxyUrl(url);

  if (!/^https?:\/\//i.test(normalized)) {
    throw new Error(
      `Invalid Proxy URL: "${url}". Expected a full URL like ${DEFAULT_LOCAL_PROXY_URL}`,
    );
  }

  let parsed: URL;
  try {
    parsed = new URL(normalized);
  } catch {
    throw new Error(
      `Invalid Proxy URL: "${url}". Expected a full URL like ${DEFAULT_LOCAL_PROXY_URL}`,
    );
  }

  // Mixed content guardrail: HTTPS taskpane -> HTTP proxy.
  // This tends to fail in Office webviews (macOS), so fail fast with guidance.
  if (typeof window !== "undefined" && window.location?.protocol === "https:" && parsed.protocol === "http:") {
    throw new Error(
      `Proxy URL is HTTP (${normalized}) but the add-in is served over HTTPS. Office webviews may block this as mixed content. ` +
        `Use ${DEFAULT_LOCAL_PROXY_URL} and run a local HTTPS proxy. See ${PROXY_HELPER_DOCS_URL}.`,
    );
  }

  return normalized;
}

function isLoopbackHostname(hostname: string): boolean {
  const h = hostname.toLowerCase();
  if (h === "localhost") return true;
  if (h === "::1" || h === "0:0:0:0:0:0:0:1") return true;
  if (h.startsWith("127.")) return true;
  if (h.startsWith("::ffff:127.")) return true;
  return false;
}

/**
 * Returns true if the proxy URL points at a loopback/localhost address.
 * Useful for warning users when they configure a remote proxy.
 */
export function isLoopbackProxyUrl(url: string): boolean {
  const normalized = normalizeProxyUrl(url);
  try {
    const parsed = new URL(normalized);
    return isLoopbackHostname(parsed.hostname);
  } catch {
    return false;
  }
}
