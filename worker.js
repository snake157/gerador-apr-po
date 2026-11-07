// Cloudflare Workers: /api/... vai para o carteiro (functions/api/[acao].js); o resto é o site estático.
// O endereço principal (/) já abre o gerador direto, sem passar pela página "Abrir o Gerador" (index.html).
import { onRequestPost, onRequestOptions } from './functions/api/[acao].js';

const GERADOR = '/Gerador%20de%20APR%20e%20PO.html';
const MANUAL = '/Manual%20do%20Gerador.html';   // aberto pelo endereço curto /Manual (07/11/2026)

async function arquivo(env, request, caminho) {
  const base = new URL(request.url);
  let r = await env.ASSETS.fetch(new Request(new URL(caminho, base), request));
  // o Cloudflare pode redirecionar "x.html" para "x": segue aqui dentro, sem mudar o endereço na barra
  if (r.status >= 300 && r.status < 400 && r.headers.get('Location')) r = await env.ASSETS.fetch(new Request(new URL(r.headers.get('Location'), base), request));
  return r;
}

export default {
  async fetch(request, env) {
    const caminho = new URL(request.url).pathname;
    const m = caminho.match(/^\/api\/([a-z]+)\/?$/);
    if (m) {
      if (request.method === 'OPTIONS') return onRequestOptions();
      if (request.method === 'POST') return onRequestPost({ request, env, params: { acao: m[1] } });
      return new Response('Use POST', { status: 405 });
    }
    if (caminho === '/' || caminho === '/index.html' || caminho === '/index') return arquivo(env, request, GERADOR);
    if (/^\/manual$/i.test(caminho)) return arquivo(env, request, MANUAL);
    if (/^\/(manual\/|Manual%20do%20Gerador(\.html)?)$/i.test(caminho)) return Response.redirect(new URL('/Manual' + new URL(request.url).hash, request.url), 301);
    return env.ASSETS.fetch(request);
  },
};
