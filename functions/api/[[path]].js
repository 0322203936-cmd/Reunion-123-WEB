// Cloudflare Pages Function: keep the existing Node backend/API during the
// migration. The browser keeps using relative /api URLs; Cloudflare forwards
// only those requests to the validated backend and serves the static UI itself.

const DEFAULT_ORIGIN = 'https://reunion-123.onrender.com';

function backendOrigin(env) {
  return String(env.RENDER_API_ORIGIN || DEFAULT_ORIGIN).replace(/\/$/, '');
}

export async function onRequest(context) {
  const incoming = new URL(context.request.url);
  const target = new URL(`${backendOrigin(context.env)}${incoming.pathname}${incoming.search}`);

  const headers = new Headers(context.request.headers);
  headers.delete('host');
  headers.set('x-forwarded-host', incoming.host);
  headers.set('x-forwarded-proto', incoming.protocol.replace(':', ''));

  const init = {
    method: context.request.method,
    headers,
    redirect: 'manual'
  };
  if (!['GET', 'HEAD'].includes(context.request.method)) {
    init.body = context.request.body;
  }

  const response = await fetch(target, init);
  const out = new Response(response.body, response);
  out.headers.set('cache-control', 'no-store');
  return out;
}
