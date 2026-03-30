/**
 * Ambient type declarations for Vercel Edge Runtime.
 *
 * Edge Runtime provides standard Web APIs (covered by "DOM" lib)
 * plus `process.env` for environment variables.
 */
declare const process: { env: Record<string, string | undefined> };
