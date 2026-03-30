/**
 * Vercel Edge Function — Google Gemini API proxy.
 *
 * We use an exact function route (`/api/gemini-proxy`) plus a Vercel rewrite
 * from `/api/gemini/*` to avoid relying on dynamic catch-all API filenames.
 * The rewrite passes the original suffix as `?proxyPath=...`.
 *
 * This enables SaaS mode: users don't need their own API keys.
 */

export const config = {
  runtime: 'edge',
};

const GEMINI_BASE = 'https://generativelanguage.googleapis.com';

const ALLOWED_ORIGINS = new Set([
  'https://hyperexcel.vercel.app',
]);

function isAllowedOrigin(origin: string | null): boolean {
  if (!origin) return false;
  if (ALLOWED_ORIGINS.has(origin)) return true;
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

function normalizeGeminiPath(path: string): string {
  const trimmed = path.replace(/^\/+/, '');
  if (/^v\d(?:alpha|beta)?\//i.test(trimmed)) {
    return trimmed;
  }
  return `v1beta/${trimmed}`;
}

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

  if (request.method === 'OPTIONS') {
    return new Response(null, {
      status: 204,
      headers: {
        ...corsHeaders(origin),
        'Access-Control-Allow-Headers':
          request.headers.get('access-control-request-headers') || '*',
      },
    });
  }

  if (request.method !== 'GET' && request.method !== 'POST') {
    return jsonError({ error: 'Method not allowed' }, 405, origin);
  }

  if (origin && !isAllowedOrigin(origin)) {
    return jsonError({ error: 'Forbidden' }, 403, origin);
  }
  if (!origin && request.method !== 'GET') {
    return jsonError({ error: 'Forbidden: Origin header required' }, 403, null);
  }

  const geminiApiKey = process.env.GEMINI_API_KEY;
  if (!geminiApiKey) {
    return jsonError(
      { error: 'Server configuration error: GEMINI_API_KEY not set' },
      500,
      origin,
    );
  }

  const url = new URL(request.url);
  const pathAfterPrefix =
    url.searchParams.get('proxyPath') ||
    url.searchParams.get('path') ||
    url.pathname.replace(/^\/api\/gemini(?:-proxy)?\/?/, '');

  if (!pathAfterPrefix) {
    return jsonError(
      { error: 'Missing API path. Example: /api/gemini/v1beta/models' },
      400,
      origin,
    );
  }

  const normalizedPath = normalizeGeminiPath(pathAfterPrefix);
  const targetUrl = new URL(`${GEMINI_BASE}/${normalizedPath}`);
  for (const [key, value] of url.searchParams) {
    if (key !== 'key' && key !== 'proxyPath' && key !== 'path') {
      targetUrl.searchParams.set(key, value);
    }
  }
  targetUrl.searchParams.set('key', geminiApiKey);

  const outboundHeaders = new Headers();
  const contentType = request.headers.get('content-type');
  if (contentType) {
    outboundHeaders.set('Content-Type', contentType);
  }
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

    const responseHeaders = new Headers();
    const respContentType = response.headers.get('content-type');
    if (respContentType) {
      responseHeaders.set('Content-Type', respContentType);
    }
    const cacheControl = response.headers.get('cache-control');
    if (cacheControl) {
      responseHeaders.set('Cache-Control', cacheControl);
    }
    const cors = corsHeaders(origin);
    for (const [key, value] of Object.entries(cors)) {
      responseHeaders.set(key, value);
    }
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
