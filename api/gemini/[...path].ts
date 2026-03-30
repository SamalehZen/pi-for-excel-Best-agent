/**
 * Vercel Edge Function — Google Gemini API proxy.
 *
 * Catches all requests under /api/gemini/* and forwards them to
 * generativelanguage.googleapis.com with the server-side API key injected.
 *
 * This enables SaaS mode: users don't need their own API keys.
 */

export const config = {
  runtime: 'edge',
};

const GEMINI_BASE = 'https://generativelanguage.googleapis.com';

// Production origin. Localhost allowed for development.
const ALLOWED_ORIGINS = new Set([
  'https://hyperexcel.vercel.app',
]);

function isAllowedOrigin(origin: string | null): boolean {
  if (!origin) return false;
  if (ALLOWED_ORIGINS.has(origin)) return true;
  // Allow localhost for development
  try {
    const url = new URL(origin);
    return url.hostname === 'localhost' || url.hostname === '127.0.0.1';
  } catch {
    return false;
  }
}

function corsHeaders(origin: string | null): Record<string, string> {
  const headers: Record<string, string> = {
    'Access-Control-Allow-Methods': 'GET,POST,OPTIONS',
    'Access-Control-Expose-Headers': '*',
    'Access-Control-Max-Age': '86400',
  };
  if (origin && isAllowedOrigin(origin)) {
    headers['Access-Control-Allow-Origin'] = origin;
    headers['Vary'] = 'Origin';
  }
  return headers;
}

/** Build a JSON error response with correct CORS and Content-Type headers. */
function jsonError(
  body: Record<string, string>,
  status: number,
  origin: string | null,
): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: {
      'Content-Type': 'application/json',
      ...corsHeaders(origin),
    },
  });
}

export default async function handler(request: Request): Promise<Response> {
  const origin = request.headers.get('origin');

  // CORS preflight
  if (request.method === 'OPTIONS') {
    return new Response(null, {
      status: 204,
      headers: {
        ...corsHeaders(origin),
        'Access-Control-Allow-Headers': request.headers.get('access-control-request-headers') || '*',
      },
    });
  }

  // Only allow GET and POST
  if (request.method !== 'GET' && request.method !== 'POST') {
    return jsonError({ error: 'Method not allowed' }, 405, origin);
  }

  // Origin check — reject spoofed origins AND origin-less POST requests.
  // Browsers always send Origin on cross-origin POSTs (and modern browsers
  // on same-origin POSTs too), so a missing Origin on POST means a
  // non-browser client trying to use the proxy as an open relay.
  if (origin && !isAllowedOrigin(origin)) {
    return jsonError({ error: 'Forbidden' }, 403, origin);
  }
  if (!origin && request.method !== 'GET') {
    return jsonError({ error: 'Forbidden: Origin header required' }, 403, null);
  }

  // Validate API key is configured
  const geminiApiKey = process.env.GEMINI_API_KEY;
  if (!geminiApiKey) {
    return jsonError(
      { error: 'Server configuration error: GEMINI_API_KEY not set' },
      500,
      origin,
    );
  }

  // Extract the proxied sub-path.
  // In Build Output API mode, the config.json rewrite passes the original
  // path as ?proxyPath=... since the function is at /api/gemini-proxy
  // (no catch-all brackets in the directory name).
  // Fallback: parse from pathname for local dev / direct invocation.
  const url = new URL(request.url);
  const pathAfterPrefix =
    url.searchParams.get('proxyPath') ||
    url.pathname.replace(/^\/api\/gemini(?:-proxy)?\/?/, '');

  if (!pathAfterPrefix) {
    return jsonError(
      { error: 'Missing API path. Example: /api/gemini/v1beta/models' },
      400,
      origin,
    );
  }

  // Build target URL
  const targetUrl = new URL(`${GEMINI_BASE}/${pathAfterPrefix}`);

  // Copy query params but REPLACE the API key
  for (const [key, value] of url.searchParams) {
    if (key !== 'key') {
      targetUrl.searchParams.set(key, value);
    }
  }
  targetUrl.searchParams.set('key', geminiApiKey);

  // Build outbound headers — strip browser-specific headers
  const outboundHeaders = new Headers();
  const contentType = request.headers.get('content-type');
  if (contentType) {
    outboundHeaders.set('Content-Type', contentType);
  }
  // Forward x-goog-api-client if present (telemetry)
  const apiClient = request.headers.get('x-goog-api-client');
  if (apiClient) {
    outboundHeaders.set('x-goog-api-client', apiClient);
  }

  try {
    const response = await fetch(targetUrl.toString(), {
      method: request.method,
      headers: outboundHeaders,
      body: request.method === 'POST' ? request.body : undefined,
    });

    // Stream the response back with CORS headers
    const responseHeaders = new Headers();
    // Preserve content-type (important for SSE streaming: text/event-stream)
    const respContentType = response.headers.get('content-type');
    if (respContentType) {
      responseHeaders.set('Content-Type', respContentType);
    }
    // Preserve cache-control
    const cacheControl = response.headers.get('cache-control');
    if (cacheControl) {
      responseHeaders.set('Cache-Control', cacheControl);
    }
    // Add CORS headers
    const cors = corsHeaders(origin);
    for (const [key, value] of Object.entries(cors)) {
      responseHeaders.set(key, value);
    }
    // No caching for API responses
    if (!cacheControl) {
      responseHeaders.set('Cache-Control', 'no-store');
    }

    return new Response(response.body, {
      status: response.status,
      headers: responseHeaders,
    });
  } catch (error: unknown) {
    const message = error instanceof Error ? error.message : 'Unknown proxy error';
    return jsonError({ error: `Proxy error: ${message}` }, 502, origin);
  }
}
