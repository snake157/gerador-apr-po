// Cloudflare Workers: /api/... vai para o carteiro (functions/api/[acao].js); o resto é o site estático.
import { onRequestPost, onRequestOptions } from './functions/api/[acao].js';

export default {
  async fetch(request, env) {
    const m = new URL(request.url).pathname.match(/^\/api\/([a-z]+)\/?$/);
    if (!m) return env.ASSETS.fetch(request);
    if (request.method === 'OPTIONS') return onRequestOptions();
    if (request.method === 'POST') return onRequestPost({ request, env, params: { acao: m[1] } });
    return new Response('Use POST', { status: 405 });
  },
};
