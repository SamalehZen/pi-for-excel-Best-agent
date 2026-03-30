/**
 * SaaS mode detection.
 *
 * When the add-in runs on a known production host, it operates in "SaaS mode":
 * - Google Gemini is auto-configured as the default provider
 * - API calls are routed through the server-side proxy (/api/gemini/)
 * - Users don't need to enter an API key or run a local proxy
 */

/** Production hostnames where SaaS mode is active. */
const SAAS_HOSTS: ReadonlySet<string> = new Set([
  "hyperexcel.vercel.app",
]);

/**
 * Returns true when the app is running in SaaS mode (production URL).
 * Always returns false during local development (localhost).
 */
export function isSaasMode(): boolean {
  if (typeof window === "undefined") return false;
  return SAAS_HOSTS.has(window.location.hostname);
}

/** Base path for the Gemini API proxy on the server. */
export const SAAS_GEMINI_PROXY_PATH = "/api/gemini";

/** The provider name used in SaaS mode. */
export const SAAS_PROVIDER = "google";

/** Placeholder API key stored in ProviderKeysStore for SaaS mode. */
export const SAAS_PLACEHOLDER_KEY = "saas-managed";
