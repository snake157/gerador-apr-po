// "Carteiro" do gerador (Cloudflare Pages Functions): recebe pictogramas e perigos novos enviados por qualquer
// pessoa, guarda como PENDENTE no ramo "envios" do GitHub e, quando o aprovador aprova, grava no ramo main
// (o Cloudflare publica de novo sozinho, em ~1 minuto, e todos passam a ver).
//
// Endereços (todos POST, com JSON):
//   /api/enviar     {autor, tipo, ...}  envio aberto, sem senha (no máximo 10 por hora de cada computador)
//                   tipo "perigo" | "epi": {nome, arquivo (base64 do original), nomeArquivo, mini (data URL), obs}
//                   tipo "perigo-novo":    {perigo, risco, obs}
//   /api/entrar     {senha}                         -> {ok}   (confere a senha do aprovador)
//   /api/contar     {senha}                         -> {total}
//   /api/pendentes  {senha}                         -> {envios:[...]}
//   /api/aprovar    {senha, id, perigo?, risco?, pd?, obs?}   (perigo novo: o aprovador corrige e escolhe o PD)
//   /api/recusar    {senha, id}
//
// Configurar no Cloudflare (Settings > Variables and Secrets), como "Secret":
//   GITHUB_TOKEN  chave do GitHub que só mexe neste repositório (Contents: Read and write)
//   SENHA_ADMIN   senha de quem aprova
'use strict';

const REPO = 'snake157/gerador-apr-po';
const MAIN = 'main';
const RAMO = 'envios';
const PASTA = { perigo: 'Pictogramas/Perigos', epi: 'Pictogramas/EPIs' };
const LISTA = { perigo: 'perigos', epi: 'epis1' };
const BIB = 'Pictogramas/biblioteca.js';
const TABELA_XLSX = 'recursos/Perigos e Riscos com PD.xlsx';
const TABELA_JS = 'recursos/perigos_riscos.js';
const MAX = 2 * 1024 * 1024;          // imagem original: até 2 MB
const POR_HORA = 10;                  // envios por hora de cada computador

const CORS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};
const resposta = (obj, status = 200) =>
  new Response(JSON.stringify(obj), { status, headers: { 'Content-Type': 'application/json; charset=utf-8', ...CORS } });
const erro = (msg, status = 400) => resposta({ erro: msg }, status);
const limpar = (s, max) => String(s ?? '').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').trim().slice(0, max);
const norm = s => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().replace(/[^A-Z0-9]+/g, ' ').trim();

// ---------- GitHub ----------
async function gh(env, caminho, opcoes = {}) {
  const r = await fetch(`https://api.github.com/repos/${REPO}/${caminho}`, {
    method: opcoes.method || 'GET',
    body: opcoes.body,
    headers: {
      Authorization: `Bearer ${env.GITHUB_TOKEN}`,
      Accept: opcoes.raw ? 'application/vnd.github.raw' : 'application/vnd.github+json',
      'User-Agent': 'gerador-apr-po',
      'X-GitHub-Api-Version': '2022-11-28',
      ...(opcoes.body ? { 'Content-Type': 'application/json' } : {}),
    },
  });
  if (opcoes.aceitar404 && r.status === 404) return null;
  if (opcoes.aceitar422 && r.status === 422) return { conflito: true };
  if (!r.ok) throw new Error(`GitHub ${r.status}: ${(await r.text()).slice(0, 200)}`);
  if (opcoes.bytes) return new Uint8Array(await r.arrayBuffer());
  return opcoes.raw ? r.text() : r.json();
}
const caminhoUrl = c => c.split('/').map(encodeURIComponent).join('/');
const lerArquivo = (env, caminho, ref, bytes) =>
  gh(env, `contents/${caminhoUrl(caminho)}?ref=${ref}`, { raw: true, bytes });

// o ramo "envios" guarda só os pendentes (sem o resto do site); é criado na primeira vez
async function garantirRamo(env) {
  if (await gh(env, `git/ref/heads/${RAMO}`, { aceitar404: true })) return;
  const blob = await gh(env, 'git/blobs', { method: 'POST', body: JSON.stringify({ content: 'Envios pendentes do gerador (não apagar).\n', encoding: 'utf-8' }) });
  const tree = await gh(env, 'git/trees', { method: 'POST', body: JSON.stringify({ tree: [{ path: 'LEIA-ME.txt', mode: '100644', type: 'blob', sha: blob.sha }] }) });
  const commit = await gh(env, 'git/commits', { method: 'POST', body: JSON.stringify({ message: 'Ramo de envios pendentes', tree: tree.sha, parents: [] }) });
  await gh(env, 'git/refs', { method: 'POST', body: JSON.stringify({ ref: `refs/heads/${RAMO}`, sha: commit.sha }) });
}
const listarPendentes = async env =>
  ((await gh(env, `contents/pendentes?ref=${RAMO}`, { aceitar404: true })) || []).filter(f => f.name.endsWith('.json'));

async function lerPendente(env, id) {
  if (!/^[\w-]+$/.test(id || '')) throw new Error('envio inválido');
  const meta = await gh(env, `contents/pendentes/${id}.json?ref=${RAMO}`, { aceitar404: true });
  if (!meta) return null;
  return { meta, envio: JSON.parse(await lerArquivo(env, `pendentes/${id}.json`, RAMO)) };
}
const apagarPendente = (env, id, sha, motivo) =>
  gh(env, `contents/pendentes/${id}.json`, { method: 'DELETE', body: JSON.stringify({ message: motivo, sha, branch: RAMO }) });

// ---------- biblioteca.js: troca ou acrescenta um pictograma sem montar o arquivo inteiro de novo ----------
function colocarNaBiblioteca(texto, lista, nome, img) {
  const item = `{"nome":${JSON.stringify(nome)},"img":${JSON.stringify(img)}}`;
  const ini = texto.indexOf(`"${lista}":[`);
  if (ini < 0) throw new Error(`a biblioteca não tem a lista ${lista}`);
  const abre = ini + lista.length + 4;
  const fim = texto.indexOf(']', abre);   // imagens são base64 e nomes não têm "]": o primeiro "]" fecha a lista
  const antes = `{"nome":${JSON.stringify(nome)},"img":"`;
  const ja = texto.indexOf(antes, abre);
  if (ja >= 0 && ja < fim) {
    const fecha = texto.indexOf('"}', ja + antes.length) + 2;
    return { texto: texto.slice(0, ja) + item + texto.slice(fecha), substituiu: true };
  }
  return { texto: texto.slice(0, fim) + (fim === abre ? '' : ',') + item + texto.slice(fim), substituiu: false };
}

// ---------- tabela de perigos: acrescenta uma linha no perigos_riscos.js ----------
function colocarNaTabelaJs(texto, linha) {
  const fim = texto.lastIndexOf('];');
  if (fim < 0 || !texto.includes('window.TABELA_PERIGOS = [')) throw new Error('perigos_riscos.js com formato inesperado');
  const vazio = /\[\s*$/.test(texto.slice(0, fim));
  return texto.slice(0, fim) + (vazio ? '' : ',') + JSON.stringify(linha) + texto.slice(fim);
}
const perigosDaTabela = texto => {
  const ini = texto.indexOf('window.TABELA_PERIGOS = ') + 24;
  return JSON.parse(texto.slice(ini, texto.lastIndexOf(';')));
};

// ---------- planilha .xlsx: lê e grava o zip, acrescenta uma linha na 1ª aba ----------
const CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
const crc32 = b => { let c = 0xFFFFFFFF; for (let i = 0; i < b.length; i++) c = CRC[(c ^ b[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; };

async function lerZip(bytes) {
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let e = bytes.byteLength - 22;
  while (e >= 0 && dv.getUint32(e, true) !== 0x06054b50) e--;
  if (e < 0) throw new Error('a planilha de perigos está danificada');
  const n = dv.getUint16(e + 10, true);
  let p = dv.getUint32(e + 16, true);
  const dec = new TextDecoder(), arqs = [];
  for (let i = 0; i < n; i++) {
    const metodo = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
    const nl = dv.getUint16(p + 28, true), el = dv.getUint16(p + 30, true), cl = dv.getUint16(p + 32, true), lo = dv.getUint32(p + 42, true);
    const nome = dec.decode(bytes.subarray(p + 46, p + 46 + nl));
    const lnl = dv.getUint16(lo + 26, true), lel = dv.getUint16(lo + 28, true);
    let dados = bytes.subarray(lo + 30 + lnl + lel, lo + 30 + lnl + lel + csize);
    if (metodo === 8) dados = new Uint8Array(await new Response(new Blob([dados]).stream().pipeThrough(new DecompressionStream('deflate-raw'))).arrayBuffer());
    else if (metodo !== 0) throw new Error('compressão desconhecida na planilha');
    arqs.push({ nome, dados });
    p += 46 + nl + el + cl;
  }
  return arqs;
}
function gravarZip(arqs) {   // sem compressão, como o Exportar Excel do gerador
  const enc = new TextEncoder(), partes = [], central = [];
  let pos = 0;
  for (const a of arqs) {
    const nome = enc.encode(a.nome), d = a.dados, crc = crc32(d);
    const h = new DataView(new ArrayBuffer(30));
    h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true); h.setUint16(12, 0x21, true);
    h.setUint32(14, crc, true); h.setUint32(18, d.length, true); h.setUint32(22, d.length, true); h.setUint16(26, nome.length, true);
    partes.push(new Uint8Array(h.buffer), nome, d);
    const c = new DataView(new ArrayBuffer(46));
    c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(8, 0x0800, true);
    c.setUint16(14, 0x21, true); c.setUint32(16, crc, true); c.setUint32(20, d.length, true); c.setUint32(24, d.length, true);
    c.setUint16(28, nome.length, true); c.setUint32(42, pos, true);
    central.push(new Uint8Array(c.buffer), nome);
    pos += 30 + nome.length + d.length;
  }
  const tam = central.reduce((s, x) => s + x.length, 0);
  const fim = new DataView(new ArrayBuffer(22));
  fim.setUint32(0, 0x06054b50, true); fim.setUint16(8, arqs.length, true); fim.setUint16(10, arqs.length, true);
  fim.setUint32(12, tam, true); fim.setUint32(16, pos, true);
  const tudo = [...partes, ...central, new Uint8Array(fim.buffer)];
  const out = new Uint8Array(tudo.reduce((s, x) => s + x.length, 0));
  let o = 0; for (const x of tudo) { out.set(x, o); o += x.length; }
  return out;
}
const escXml = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
// copia a última linha da planilha (mesma altura e estilos) com os dados do perigo novo
function colocarNaPlanilha(xml, l) {
  const fimDados = xml.lastIndexOf('</sheetData>');
  const ini = xml.lastIndexOf('<row ', fimDados);
  if (fimDados < 0 || ini < 0) throw new Error('a planilha de perigos não tem linhas');
  const ultima = xml.slice(ini, fimDados);
  const n = Number((ultima.match(/^<row [^>]*\br="(\d+)"/) || [])[1]);
  if (!n) throw new Error('a planilha de perigos tem um formato inesperado');
  const nova = n + 1;
  const estilo = col => (ultima.match(new RegExp(`<c r="${col}${n}"[^>]*?\\ss="(\\d+)"`)) || [])[1];
  const attrs = (ultima.match(/^<row ([^>]*?)\/?>/) || [, ''])[1].replace(/\br="\d+"/, `r="${nova}"`).replace(/\s*spans="[^"]*"/, '');
  const sa = col => estilo(col) ? ` s="${estilo(col)}"` : '';
  const txt = (col, v) => `<c r="${col}${nova}"${sa(col)} t="inlineStr"><is><t xml:space="preserve">${escXml(v)}</t></is></c>`;
  const linha = `<row ${attrs}>` + txt('A', l.perigo) + txt('B', l.risco) + `<c r="C${nova}"${sa('C')}><v>${l.pd}</v></c>` +
    (l.obs ? txt('D', l.obs) : `<c r="D${nova}"${sa('D')}/>`) + '</row>';
  return (xml.slice(0, fimDados) + linha + xml.slice(fimDados)).replace(/<dimension ref="([A-Z]+\d+):([A-Z]+)\d+"\/>/, `<dimension ref="$1:$2${nova}"/>`);
}
async function planilhaComPerigo(bytes, linha) {
  const arqs = await lerZip(bytes);
  const aba = arqs.find(a => a.nome === 'xl/worksheets/sheet1.xml');
  if (!aba) throw new Error('a planilha de perigos não tem a primeira aba');
  aba.dados = new TextEncoder().encode(colocarNaPlanilha(new TextDecoder().decode(aba.dados), linha));
  return gravarZip(arqs);
}
const base64 = bytes => { let s = ''; for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000)); return btoa(s); };

// grava vários arquivos no main num commit só; se outro envio entrou no meio, monta de novo e tenta outra vez
async function gravarNoMain(env, mensagem, montar) {
  for (let tentativa = 0; tentativa < 3; tentativa++) {
    const ref = await gh(env, `git/ref/heads/${MAIN}`);
    const pai = await gh(env, `git/commits/${ref.object.sha}`);
    const lista = await montar(ref.object.sha);
    if (!lista.length) return null;   // nada a gravar (ex.: perigo que já estava na tabela)
    const tree = [];
    for (const a of lista) {
      const blob = await gh(env, 'git/blobs', { method: 'POST', body: JSON.stringify(a.base64 ? { content: a.base64, encoding: 'base64' } : { content: a.texto, encoding: 'utf-8' }) });
      tree.push({ path: a.caminho, mode: '100644', type: 'blob', sha: blob.sha });
    }
    const novaTree = await gh(env, 'git/trees', { method: 'POST', body: JSON.stringify({ base_tree: pai.tree.sha, tree }) });
    const commit = await gh(env, 'git/commits', { method: 'POST', body: JSON.stringify({ message: mensagem, tree: novaTree.sha, parents: [ref.object.sha] }) });
    const r = await gh(env, `git/refs/heads/${MAIN}`, { method: 'PATCH', body: JSON.stringify({ sha: commit.sha, force: false }), aceitar422: true });
    if (!r.conflito) return commit.sha;
  }
  throw new Error('o GitHub estava ocupado; tente aprovar de novo');
}

// ---------- ações ----------
async function enviar(env, request, d) {
  const tipo = d.tipo;
  if (!['perigo', 'epi', 'perigo-novo'].includes(tipo)) return erro('tipo inválido');
  const autor = limpar(d.autor, 60);
  if (!autor) return erro('escreva o seu nome');
  const envio = { tipo, autor, data: new Date().toISOString(), obs: limpar(d.obs, 500) };
  if (tipo === 'perigo-novo') {
    envio.perigo = limpar(d.perigo, 200); envio.risco = limpar(d.risco, 400);
    if (!envio.perigo || !envio.risco) return erro('preencha o PERIGO e o RISCO');
    envio.nome = envio.perigo;
  } else {
    envio.nome = limpar(d.nome, 150);
    if (!envio.nome) return erro('falta o nome do pictograma');
    if (/[\\/:*?"<>|[\]]/.test(envio.nome)) return erro('o nome não pode ter \\ / : * ? " < > | [ ]');
    envio.arquivo = String(d.arquivo || '');
    if (!/^[A-Za-z0-9+/=]+$/.test(envio.arquivo) || envio.arquivo.length > MAX * 1.37) return erro('a imagem está vazia ou tem mais de 2 MB');
    envio.mini = String(d.mini || '');
    if (!/^data:image\/(webp|png|jpeg);base64,[A-Za-z0-9+/=]+$/.test(envio.mini) || envio.mini.length > 400000) return erro('falta a miniatura do pictograma');
    envio.nomeArquivo = limpar(d.nomeArquivo, 200);
  }
  await garantirRamo(env);
  // limite por computador: o endereço vira um código (não fica guardado como endereço)
  const ip = request.headers.get('CF-Connecting-IP') || 'local';
  const dig = new Uint8Array(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(ip + REPO)));
  const quem = [...dig.slice(0, 4)].map(b => b.toString(16).padStart(2, '0')).join('');
  const agora = Date.now();
  const recentes = (await listarPendentes(env)).filter(f => {
    const m = f.name.match(/^(\d{14})-([0-9a-f]{8})-/);
    if (!m || m[2] !== quem) return false;
    const t = Date.UTC(+m[1].slice(0, 4), +m[1].slice(4, 6) - 1, +m[1].slice(6, 8), +m[1].slice(8, 10), +m[1].slice(10, 12), +m[1].slice(12, 14));
    return agora - t < 3600e3;
  });
  if (recentes.length >= POR_HORA) return erro(`limite de ${POR_HORA} envios por hora atingido; tente mais tarde`, 429);
  envio.id = new Date(agora).toISOString().replace(/\D/g, '').slice(0, 14) + '-' + quem + '-' + crypto.randomUUID().slice(0, 6);
  const json = new TextEncoder().encode(JSON.stringify(envio));
  await gh(env, `contents/pendentes/${envio.id}.json`, { method: 'PUT', body: JSON.stringify({ message: `Envio pendente: ${envio.nome} (${autor})`, content: base64(json), branch: RAMO }) });
  return resposta({ ok: true, id: envio.id });
}

async function pendentes(env) {
  const envios = [];
  for (const f of (await listarPendentes(env)).sort((a, b) => a.name < b.name ? -1 : 1)) {
    try {
      const e = JSON.parse(await lerArquivo(env, `pendentes/${f.name}`, RAMO));
      delete e.arquivo;   // o original não precisa ir para a tela (a miniatura basta)
      envios.push(e);
    } catch (x) { /* arquivo estragado: fica de fora */ }
  }
  return resposta({ envios });
}

async function aprovar(env, d) {
  const p = await lerPendente(env, d.id);
  if (!p) return erro('esse envio já foi aprovado ou recusado', 404);
  const e = p.envio;
  let substituiu = false, linha = null;
  if (e.tipo === 'perigo-novo') {
    linha = { perigo: limpar(d.perigo ?? e.perigo, 200), risco: limpar(d.risco ?? e.risco, 400), pd: Number(d.pd), obs: limpar(d.obs ?? '', 500) };
    if (!linha.perigo || !linha.risco) return erro('preencha o PERIGO e o RISCO');
    if (![1, 2, 3, 4, 5].includes(linha.pd)) return erro('escolha o PD (1 a 5)');
  }
  let ja = false;
  const sha = await gravarNoMain(env,
    e.tipo === 'perigo-novo' ? `Perigo novo "${linha.perigo}" na tabela (enviado por ${e.autor})`
      : `Pictograma de ${e.tipo === 'perigo' ? 'perigo' : 'EPI'} "${e.nome}" (enviado por ${e.autor})`,
    async base => {
      if (e.tipo === 'perigo-novo') {
        const js = await lerArquivo(env, TABELA_JS, base);
        if (perigosDaTabela(js).some(r => norm(r.perigo) === norm(linha.perigo))) { ja = true; return []; }
        const xlsx = await lerArquivo(env, TABELA_XLSX, base, true);
        return [{ caminho: TABELA_JS, texto: colocarNaTabelaJs(js, linha) }, { caminho: TABELA_XLSX, base64: base64(await planilhaComPerigo(xlsx, linha)) }];
      }
      const r = colocarNaBiblioteca(await lerArquivo(env, BIB, base), LISTA[e.tipo], e.nome, e.mini);
      substituiu = r.substituiu;
      const ext = ((e.nomeArquivo || '').match(/\.(png|jpe?g)$/i) || ['.png'])[0].toLowerCase();
      return [{ caminho: BIB, texto: r.texto }, { caminho: `${PASTA[e.tipo]}/${e.nome}${ext}`, base64: e.arquivo }];
    });
  if (ja) return erro(`o perigo "${linha.perigo}" já está na tabela`, 409);
  await apagarPendente(env, e.id, p.meta.sha, `Aprovado: ${e.nome}`);
  return resposta({ ok: true, commit: sha, substituiu, linha, tipo: e.tipo, nome: e.nome, mini: e.mini });
}

async function recusar(env, d) {
  const p = await lerPendente(env, d.id);
  if (!p) return erro('esse envio já foi aprovado ou recusado', 404);
  await apagarPendente(env, p.envio.id, p.meta.sha, `Recusado: ${p.envio.nome}`);
  return resposta({ ok: true });
}

export async function onRequestOptions() { return new Response(null, { headers: CORS }); }

export async function onRequestPost({ request, env, params }) {
  try {
    if (!env.GITHUB_TOKEN || !env.SENHA_ADMIN) return erro('o envio ainda não foi configurado no Cloudflare (falta GITHUB_TOKEN ou SENHA_ADMIN)', 503);
    let d;
    try { d = await request.json(); } catch (x) { return erro('pedido inválido'); }
    const acao = params.acao;
    if (acao === 'enviar') return await enviar(env, request, d);
    if (String(d.senha || '') !== env.SENHA_ADMIN) return erro('senha de aprovador incorreta', 401);
    if (acao === 'entrar') return resposta({ ok: true });
    if (acao === 'contar') return resposta({ total: (await listarPendentes(env)).length });
    if (acao === 'pendentes') return await pendentes(env);
    if (acao === 'aprovar') return await aprovar(env, d);
    if (acao === 'recusar') return await recusar(env, d);
    return erro('endereço desconhecido', 404);
  } catch (x) {
    return erro('não consegui terminar: ' + (x.message || x), 500);
  }
}

