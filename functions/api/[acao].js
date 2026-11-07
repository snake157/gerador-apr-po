// "Carteiro" do gerador (Cloudflare Pages Functions): recebe pictogramas e a tabela de perigos enviados pelos
// usuários, guarda como PENDENTE no ramo "envios" do GitHub e, quando o aprovador aprova, grava no ramo main
// (o Cloudflare publica de novo sozinho, em ~1 minuto, e todos passam a ver).
//
// Endereços (todos POST, com JSON):
//   /api/entrar     {senha}                              -> {nome, aprovador}
//   /api/enviar     {senha, tipo, nome, arquivo, nomeArquivo, mini?, js?}
//                   tipo = "perigo" | "epi" | "tabela"; arquivo = base64 do original;
//                   mini = data URL da miniatura (pictogramas); js = texto do perigos_riscos.js (tabela)
//   /api/pendentes  {senha (aprovador)}                  -> {envios:[...]}
//   /api/aprovar    {senha (aprovador), id}
//   /api/recusar    {senha (aprovador), id}
//
// Configurar no Cloudflare (Settings > Variables and Secrets), todas como "Secret":
//   GITHUB_TOKEN  chave do GitHub que só mexe no repositório (Contents: Read and write)
//   SENHAS        uma pessoa por linha (ou separadas por ;) no formato  Nome:senha
//   SENHA_ADMIN   senha de quem aprova (também pode enviar)
'use strict';

const REPO = 'snake157/gerador-apr-po';
const MAIN = 'main';
const RAMO = 'envios';
const PASTA = { perigo: 'Pictogramas/Perigos', epi: 'Pictogramas/EPIs' };
const LISTA = { perigo: 'perigos', epi: 'epis1' };
const TABELA_XLSX = 'recursos/Perigos e Riscos com PD.xlsx';
const TABELA_JS = 'recursos/perigos_riscos.js';
const MAX = 2 * 1024 * 1024;

const CORS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};
const resposta = (obj, status = 200) =>
  new Response(JSON.stringify(obj), { status, headers: { 'Content-Type': 'application/json; charset=utf-8', ...CORS } });
const erro = (msg, status = 400) => resposta({ erro: msg }, status);

// ---------- senhas ----------
function quem(env, senha) {
  senha = String(senha || '');
  if (!senha) return null;
  if (env.SENHA_ADMIN && senha === env.SENHA_ADMIN) return { nome: 'Aprovador', aprovador: true };
  for (const linha of String(env.SENHAS || '').split(/[\n;]+/)) {
    const i = linha.lastIndexOf(':');
    if (i > 0 && linha.slice(i + 1).trim() === senha) return { nome: linha.slice(0, i).trim(), aprovador: false };
  }
  return null;
}

// ---------- GitHub ----------
async function gh(env, caminho, opcoes = {}) {
  const r = await fetch(`https://api.github.com/repos/${REPO}/${caminho}`, {
    ...opcoes,
    headers: {
      Authorization: `Bearer ${env.GITHUB_TOKEN}`,
      Accept: opcoes.raw ? 'application/vnd.github.raw' : 'application/vnd.github+json',
      'User-Agent': 'gerador-apr-po',
      'X-GitHub-Api-Version': '2022-11-28',
      ...(opcoes.body ? { 'Content-Type': 'application/json' } : {}),
    },
  });
  if (opcoes.aceitar404 && r.status === 404) return null;
  if (!r.ok) throw new Error(`GitHub ${r.status}: ${(await r.text()).slice(0, 200)}`);
  return opcoes.raw ? r.text() : r.json();
}
const caminhoUrl = c => c.split('/').map(encodeURIComponent).join('/');

// o ramo "envios" guarda só os pendentes (sem o resto do site); é criado na primeira vez
async function garantirRamo(env) {
  if (await gh(env, `git/ref/heads/${RAMO}`, { aceitar404: true })) return;
  const blob = await gh(env, 'git/blobs', { method: 'POST', body: JSON.stringify({ content: 'Envios pendentes do gerador (não apagar).\n', encoding: 'utf-8' }) });
  const tree = await gh(env, 'git/trees', { method: 'POST', body: JSON.stringify({ tree: [{ path: 'LEIA-ME.txt', mode: '100644', type: 'blob', sha: blob.sha }] }) });
  const commit = await gh(env, 'git/commits', { method: 'POST', body: JSON.stringify({ message: 'Ramo de envios pendentes', tree: tree.sha, parents: [] }) });
  await gh(env, 'git/refs', { method: 'POST', body: JSON.stringify({ ref: `refs/heads/${RAMO}`, sha: commit.sha }) });
}

async function lerPendente(env, id) {
  if (!/^[\w-]+$/.test(id || '')) throw new Error('envio inválido');
  const meta = await gh(env, `contents/pendentes/${id}.json?ref=${RAMO}`, { aceitar404: true });
  if (!meta) return null;
  const texto = await gh(env, `contents/pendentes/${id}.json?ref=${RAMO}`, { raw: true });
  return { meta, envio: JSON.parse(texto) };
}
async function apagarPendente(env, id, sha, motivo) {
  await gh(env, `contents/pendentes/${id}.json`, { method: 'DELETE', body: JSON.stringify({ message: motivo, sha, branch: RAMO }) });
}

// ---------- biblioteca.js: troca ou acrescenta um pictograma sem montar o arquivo inteiro de novo ----------
function colocarNaBiblioteca(texto, lista, nome, img) {
  const item = `{"nome":${JSON.stringify(nome)},"img":${JSON.stringify(img)}}`;
  const ini = texto.indexOf(`"${lista}":[`);
  if (ini < 0) throw new Error(`a biblioteca não tem a lista ${lista}`);
  const abre = ini + lista.length + 3;
  const fim = texto.indexOf(']', abre);   // as imagens são base64 (sem "]"), então o primeiro "]" fecha a lista
  const antes = `{"nome":${JSON.stringify(nome)},"img":"`;
  const ja = texto.indexOf(antes, abre);
  if (ja >= 0 && ja < fim) {   // já existe com esse nome: substitui
    const fecha = texto.indexOf('"}', ja + antes.length) + 2;
    return { texto: texto.slice(0, ja) + item + texto.slice(fecha), substituiu: true };
  }
  const vazio = fim === abre;
  return { texto: texto.slice(0, fim) + (vazio ? '' : ',') + item + texto.slice(fim), substituiu: false };
}

// grava vários arquivos de uma vez no main (um commit só); tenta de novo se outro envio entrou no meio
async function gravarNoMain(env, arquivos, mensagem, montar) {
  for (let tentativa = 0; tentativa < 3; tentativa++) {
    const ref = await gh(env, `git/ref/heads/${MAIN}`);
    const pai = await gh(env, `git/commits/${ref.object.sha}`);
    const lista = await montar(ref.object.sha, arquivos);
    const tree = [];
    for (const a of lista) {
      const blob = await gh(env, 'git/blobs', { method: 'POST', body: JSON.stringify(a.base64 ? { content: a.base64, encoding: 'base64' } : { content: a.texto, encoding: 'utf-8' }) });
      tree.push({ path: a.caminho, mode: '100644', type: 'blob', sha: blob.sha });
    }
    const novaTree = await gh(env, 'git/trees', { method: 'POST', body: JSON.stringify({ base_tree: pai.tree.sha, tree }) });
    const commit = await gh(env, 'git/commits', { method: 'POST', body: JSON.stringify({ message: mensagem, tree: novaTree.sha, parents: [ref.object.sha] }) });
    const r = await fetch(`https://api.github.com/repos/${REPO}/git/refs/heads/${MAIN}`, {
      method: 'PATCH',
      headers: { Authorization: `Bearer ${env.GITHUB_TOKEN}`, Accept: 'application/vnd.github+json', 'User-Agent': 'gerador-apr-po', 'Content-Type': 'application/json' },
      body: JSON.stringify({ sha: commit.sha, force: false }),
    });
    if (r.ok) return commit.sha;
    if (r.status !== 422) throw new Error(`GitHub ${r.status}: ${(await r.text()).slice(0, 200)}`);
  }
  throw new Error('o GitHub estava ocupado; tente aprovar de novo');
}

// ---------- ações ----------
async function enviar(env, pessoa, d) {
  const tipo = d.tipo;
  if (!['perigo', 'epi', 'tabela'].includes(tipo)) return erro('tipo inválido');
  const nome = String(d.nome || '').trim();
  if (tipo !== 'tabela' && !nome) return erro('falta o nome do pictograma');
  if (/[\\/:*?"<>|\[\]]/.test(nome)) return erro('o nome não pode ter \\ / : * ? " < > | [ ]');
  const arquivo = String(d.arquivo || '');
  if (!arquivo || arquivo.length > MAX * 1.37) return erro('arquivo vazio ou maior que 2 MB');
  if (tipo !== 'tabela' && !/^data:image\/(webp|png|jpeg);base64,/.test(d.mini || '')) return erro('falta a miniatura do pictograma');
  if (tipo === 'tabela' && !String(d.js || '').startsWith('// Tabela de Perigos')) return erro('a tabela não foi lida');
  await garantirRamo(env);
  const id = new Date().toISOString().replace(/\D/g, '').slice(0, 14) + '-' + crypto.randomUUID().slice(0, 8);
  const envio = {
    id, tipo, nome, autor: pessoa.nome, data: new Date().toISOString(),
    nomeArquivo: String(d.nomeArquivo || '').slice(0, 200), arquivo, mini: d.mini || '', js: d.js || '', obs: String(d.obs || '').slice(0, 500),
  };
  const conteudo = btoa(unescape(encodeURIComponent(JSON.stringify(envio))));
  await gh(env, `contents/pendentes/${id}.json`, { method: 'PUT', body: JSON.stringify({ message: `Envio pendente: ${tipo} ${nome} (${pessoa.nome})`, content: conteudo, branch: RAMO }) });
  return resposta({ ok: true, id });
}

async function pendentes(env) {
  const lista = await gh(env, `contents/pendentes?ref=${RAMO}`, { aceitar404: true });
  if (!lista) return resposta({ envios: [] });
  const envios = [];
  for (const f of lista.filter(f => f.name.endsWith('.json')).sort((a, b) => a.name < b.name ? -1 : 1)) {
    try {
      const e = JSON.parse(await gh(env, `contents/pendentes/${caminhoUrl(f.name)}?ref=${RAMO}`, { raw: true }));
      envios.push({ id: e.id, tipo: e.tipo, nome: e.nome, autor: e.autor, data: e.data, nomeArquivo: e.nomeArquivo, mini: e.mini, obs: e.obs, tabela: e.tipo === 'tabela' ? e.js : undefined });
    } catch (x) { /* arquivo estragado: fica de fora da lista */ }
  }
  return resposta({ envios });
}

async function aprovar(env, d) {
  const p = await lerPendente(env, d.id);
  if (!p) return erro('esse envio já foi aprovado ou recusado', 404);
  const e = p.envio;
  let substituiu = false;
  const sha = await gravarNoMain(env, null, e.tipo === 'tabela'
      ? `Tabela de Perigos atualizada (enviada por ${e.autor})`
      : `Pictograma de ${e.tipo === 'perigo' ? 'perigo' : 'EPI'} "${e.nome}" (enviado por ${e.autor})`,
    async base => {
      if (e.tipo === 'tabela') return [{ caminho: TABELA_XLSX, base64: e.arquivo }, { caminho: TABELA_JS, texto: e.js }];
      const bib = await gh(env, `contents/Pictogramas/biblioteca.js?ref=${base}`, { raw: true });
      const r = colocarNaBiblioteca(bib, LISTA[e.tipo], e.nome, e.mini);
      substituiu = r.substituiu;
      const ext = (e.nomeArquivo.match(/\.(png|jpe?g)$/i) || ['.png'])[0].toLowerCase();
      return [{ caminho: 'Pictogramas/biblioteca.js', texto: r.texto }, { caminho: `${PASTA[e.tipo]}/${e.nome}${ext}`, base64: e.arquivo }];
    });
  await apagarPendente(env, e.id, p.meta.sha, `Aprovado: ${e.tipo} ${e.nome}`);
  return resposta({ ok: true, commit: sha, substituiu });
}

async function recusar(env, d) {
  const p = await lerPendente(env, d.id);
  if (!p) return erro('esse envio já foi aprovado ou recusado', 404);
  await apagarPendente(env, p.envio.id, p.meta.sha, `Recusado: ${p.envio.tipo} ${p.envio.nome}`);
  return resposta({ ok: true });
}

export async function onRequestOptions() { return new Response(null, { headers: CORS }); }

export async function onRequestPost({ request, env, params }) {
  try {
    if (!env.GITHUB_TOKEN) return erro('o carteiro ainda não foi configurado (falta GITHUB_TOKEN)', 500);
    let d;
    try { d = await request.json(); } catch (x) { return erro('pedido inválido'); }
    const pessoa = quem(env, d.senha);
    if (!pessoa) return erro('senha incorreta', 401);
    const acao = params.acao;
    if (acao === 'entrar') return resposta(pessoa);
    if (acao === 'enviar') return await enviar(env, pessoa, d);
    if (!pessoa.aprovador) return erro('só o aprovador pode fazer isso', 403);
    if (acao === 'pendentes') return await pendentes(env);
    if (acao === 'aprovar') return await aprovar(env, d);
    if (acao === 'recusar') return await recusar(env, d);
    return erro('endereço desconhecido', 404);
  } catch (x) {
    return erro('não consegui terminar: ' + (x.message || x), 500);
  }
}
