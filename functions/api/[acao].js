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
//                   pictograma: {senha, id, tipo?, nome?, modo?: "substituir" | "novo"}
//   /api/recusar    {senha, id}
//   /api/administrar {senha, tabela?: [{perigo, risco, pd}], renomeadosTabela?: {antigo: novo},
//                    pictos?: [{lista: "perigos"|"epis1", acao: "novo"|"editar"|"excluir", nome, novoNome?, mini?, arquivo?, nomeArquivo?}]}
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

// ---------- tabela de perigos (perigos_riscos.js) ----------
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
// refaz as linhas da 1ª aba a partir da tabela (Perigo, Risco, PD), com os estilos do cabeçalho e da 1ª linha de dados;
// a antiga coluna D (Observação) sai (07/11/2026). Linha que já existia mantém a altura; nova ganha altura pelo tamanho do texto.
function planilhaXml(xml, linhas) {
  const ini = xml.indexOf('<sheetData>'), fim = xml.indexOf('</sheetData>');
  if (ini < 0 || fim < 0) throw new Error('a planilha de perigos tem um formato inesperado');
  const rows = [...xml.slice(ini, fim).matchAll(/<row [^>]*?(?:\/>|>[\s\S]*?<\/row>)/g)].map(m => m[0]);
  const cab = rows.find(r => /\br="1"/.test(r)), mod = rows.find(r => /\br="2"/.test(r));
  if (!cab || !mod) throw new Error('a planilha de perigos não tem cabeçalho');
  const txtCel = c => { const m = c.match(/<t[^>]*>([\s\S]*?)<\/t>/); return m ? m[1].replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(+n)).replace(/&amp;/g, '&') : ''; };
  const alturas = {};
  for (const r of rows) { const a = (r.match(/<c r="A\d+"[^>]*>[\s\S]*?<\/c>/) || [''])[0]; const ht = (r.match(/\bht="([\d.]+)"/) || [])[1]; if (a && ht) alturas[txtCel(a).trim()] = ht; }
  const est = col => (mod.match(new RegExp(`<c r="${col}2"[^>]*?\\ss="(\\d+)"`)) || [])[1];
  const sa = col => est(col) ? ` s="${est(col)}"` : '';
  const semD = cab.replace(/<c r="D1"[^>]*?(?:\/>|>[\s\S]*?<\/c>)/, '');
  const extra = (mod.match(/^<row ([^>]*?)>/) || [, ''])[1].replace(/\s*\br="\d+"/, '').replace(/\s*\bht="[^"]*"/, '').replace(/\s*\bcustomHeight="[^"]*"/, '').replace(/\s*\bspans="[^"]*"/, '');
  const txt = (col, n, v) => `<c r="${col}${n}"${sa(col)} t="inlineStr"><is><t xml:space="preserve">${escXml(v)}</t></is></c>`;
  const corpo = linhas.map((l, k) => {
    const n = k + 2;
    const ht = alturas[l.perigo] || String(Math.max(28.5, 15 * Math.max(Math.ceil(l.perigo.length / 24), Math.ceil(l.risco.length / 30)) + 6));
    return `<row r="${n}" ht="${ht}" customHeight="1"${extra ? ' ' + extra : ''}>` + txt('A', n, l.perigo) + txt('B', n, l.risco) + `<c r="C${n}"${sa('C')}><v>${l.pd}</v></c></row>`;
  }).join('');
  return xml.slice(0, ini) + '<sheetData>' + semD + corpo + xml.slice(fim);
}
function arrumarXml(xml, linhas) {
  return planilhaXml(xml, linhas).replace(/<col [^>]*\bmin="4"[^>]*\/>/, '').replace(/<dimension ref="[^"]*"\/>/, `<dimension ref="A1:C${linhas.length + 1}"/>`);
}
async function planilhaDaTabela(bytes, linhas) {
  const arqs = await lerZip(bytes);
  const aba = arqs.find(a => a.nome === 'xl/worksheets/sheet1.xml');
  if (!aba) throw new Error('a planilha de perigos não tem a primeira aba');
  aba.dados = new TextEncoder().encode(arrumarXml(new TextDecoder().decode(aba.dados), linhas));
  return gravarZip(arqs);
}
const textoTabelaJs = linhas => '// Tabela de Perigos e Riscos da APR (perigo, risco / consequência e potencial de dano).\n' +
  `// Gravada pelo gerador online (${new Date().toLocaleDateString('pt-BR', { timeZone: 'America/Sao_Paulo' })}): Menu > ADMINISTRADOR ou APROVAR ENVIOS.\n` +
  'window.TABELA_PERIGOS = ' + JSON.stringify(linhas.map(l => ({ perigo: l.perigo, risco: l.risco, pd: l.pd }))) + ';\n';
// nomes antigos -> novos (APRs e POs antigos passam para o nome novo ao abrir); também atualiza quem apontava para o nome antigo
function colocarRenomeados(texto, variavel, mapa) {
  if (!Object.keys(mapa).length) return texto;
  return texto.replace(/\s*$/, '\n') + `// ${new Date().toISOString().slice(0, 10)}: renomeados no ADMINISTRADOR\n` +
    `(function (m, n) { for (const k in m) if (n[m[k]]) m[k] = n[m[k]]; Object.assign(m, n); })(window.${variavel} = window.${variavel} || {}, ${JSON.stringify(mapa, null, 0)});\n`;
}

const base64 = bytes => { let s = ''; for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000)); return btoa(s); };

// acha o item {"nome":...,"img":...} de uma lista da biblioteca; devolve [início, fim] no texto, ou null
function itemNaBiblioteca(texto, lista, nome) {
  const ini = texto.indexOf(`"${lista}":[`);
  if (ini < 0) throw new Error(`a biblioteca não tem a lista ${lista}`);
  const abre = ini + lista.length + 4, fim = texto.indexOf(']', abre);
  const antes = `{"nome":${JSON.stringify(nome)},"img":"`;
  const ja = texto.indexOf(antes, abre);
  if (ja < 0 || ja > fim) return null;
  return [ja, texto.indexOf('"}', ja + antes.length) + 2];
}
function tirarDaBiblioteca(texto, lista, nome) {
  const p = itemNaBiblioteca(texto, lista, nome);
  if (!p) return texto;
  let [a, b] = p;
  if (texto[b] === ',') b++; else if (texto[a - 1] === ',') a--;
  return texto.slice(0, a) + texto.slice(b);
}
function trocarNaBiblioteca(texto, lista, nome, novoNome, img) {
  const p = itemNaBiblioteca(texto, lista, nome);
  if (!p) throw new Error(`o pictograma "${nome}" não está na biblioteca`);
  const atual = texto.slice(p[0], p[1]);
  const imgAtual = atual.slice(atual.indexOf('"img":') + 6, -1);
  return texto.slice(0, p[0]) + `{"nome":${JSON.stringify(novoNome)},"img":${img ? JSON.stringify(img) : imgAtual}}` + texto.slice(p[1]);
}
// perigos_riscos.js + planilha refeitos a partir da tabela
async function arquivosDaTabela(env, base, linhas) {
  const xlsx = await lerArquivo(env, TABELA_XLSX, base, true);
  return [{ caminho: TABELA_JS, texto: textoTabelaJs(linhas) }, { caminho: TABELA_XLSX, base64: base64(await planilhaDaTabela(xlsx, linhas)) }];
}

// grava vários arquivos no main num commit só; se outro envio entrou no meio, monta de novo e tenta outra vez
async function gravarNoMain(env, mensagem, montar) {
  for (let tentativa = 0; tentativa < 3; tentativa++) {
    const ref = await gh(env, `git/ref/heads/${MAIN}`);
    const pai = await gh(env, `git/commits/${ref.object.sha}`);
    const lista = await montar(ref.object.sha);
    if (!lista.length) return null;   // nada a gravar (ex.: perigo que já estava na tabela)
    const tree = [];
    for (const a of lista) {
      if (a.apagar) { tree.push({ path: a.caminho, mode: '100644', type: 'blob', sha: null }); continue; }
      if (a.sha) { tree.push({ path: a.caminho, mode: '100644', type: 'blob', sha: a.sha }); continue; }   // renomear: mesmo conteúdo
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
    linha = { perigo: limpar(d.perigo ?? e.perigo, 200), risco: limpar(d.risco ?? e.risco, 400), pd: Number(d.pd) };
    if (!linha.perigo || !linha.risco) return erro('preencha o PERIGO e o RISCO');
    if (![1, 2, 3, 4, 5].includes(linha.pd)) return erro('escolha o PD (1 a 5)');
  }
  // pictograma: o aprovador pode trocar o tipo e o nome, e escolher SUBSTITUIR ou NOVO (entra como "Nome (2)")
  let tipo = e.tipo, nome = e.nome, modo = d.modo === 'novo' ? 'novo' : 'substituir', semNome = false;
  if (e.tipo !== 'perigo-novo') {
    if (d.tipo === 'perigo' || d.tipo === 'epi') tipo = d.tipo;
    if (d.nome !== undefined) nome = limpar(d.nome, 150);
    if (!nome) return erro('falta o nome do pictograma');
    if (/[\\/:*?"<>|[\]]/.test(nome)) return erro('o nome não pode ter \\ / : * ? " < > | [ ]');
  }
  let ja = false;
  const sha = await gravarNoMain(env,
    e.tipo === 'perigo-novo' ? `Perigo novo "${linha.perigo}" na tabela (enviado por ${e.autor})`
      : `Pictograma de ${tipo === 'perigo' ? 'perigo' : 'EPI'} "${nome}" (enviado por ${e.autor})`,
    async base => {
      if (e.tipo === 'perigo-novo') {
        const tab = perigosDaTabela(await lerArquivo(env, TABELA_JS, base));
        if (tab.some(r => norm(r.perigo) === norm(linha.perigo))) { ja = true; return []; }
        const nova = [...tab, linha];
        return arquivosDaTabela(env, base, nova);
      }
      if (tipo === 'perigo') {
        const tab = perigosDaTabela(await lerArquivo(env, TABELA_JS, base));
        if (!tab.some(r => r.perigo === nome)) { semNome = true; return []; }
      }
      const bib = await lerArquivo(env, BIB, base);
      const existe = n => bib.includes(`{"nome":${JSON.stringify(n)},"img":"`);
      let final = nome;
      if (modo === 'novo' && existe(nome)) { let k = 2; while (existe(`${nome} (${k})`)) k++; final = `${nome} (${k})`; }
      const r = colocarNaBiblioteca(bib, LISTA[tipo], final, e.mini);
      substituiu = r.substituiu; nome = final;
      const ext = ((e.nomeArquivo || '').match(/\.(png|jpe?g)$/i) || ['.png'])[0].toLowerCase();
      return [{ caminho: BIB, texto: r.texto }, { caminho: `${PASTA[tipo]}/${final}${ext}`, base64: e.arquivo }];
    });
  if (ja) return erro(`o perigo "${linha.perigo}" já está na tabela`, 409);
  if (semNome) return erro(`"${nome}" não está na tabela de perigos; escolha um nome da lista`, 400);
  await apagarPendente(env, e.id, p.meta.sha, `Aprovado: ${nome}`);
  return resposta({ ok: true, commit: sha, substituiu, linha, tipo: e.tipo === 'perigo-novo' ? e.tipo : tipo, nome, mini: e.mini });
}

// ADMINISTRADOR: grava de uma vez a tabela inteira e as mudanças nos pictogramas (incluir, editar, excluir)
const falha = msg => Object.assign(new Error(msg), { status: 400 });
const PASTA_LISTA = { perigos: 'Pictogramas/Perigos', epis1: 'Pictogramas/EPIs' };
const NOMES_ANTIGOS = 'recursos/perigos_nomes_antigos.js';
async function administrar(env, d) {
  const tabela = Array.isArray(d.tabela) ? d.tabela.map(l => ({ perigo: limpar(l.perigo, 200), risco: limpar(l.risco, 400), pd: Number(l.pd) })) : null;
  if (tabela) {
    if (!tabela.length) return erro('a tabela de perigos não pode ficar vazia');
    const vistos = new Set();
    for (const l of tabela) {
      if (!l.perigo || !l.risco) return erro('há perigo sem nome ou sem risco');
      if (![1, 2, 3, 4, 5].includes(l.pd)) return erro(`escolha o PD de "${l.perigo}"`);
      if (vistos.has(norm(l.perigo))) return erro(`o perigo "${l.perigo}" aparece duas vezes`);
      vistos.add(norm(l.perigo));
    }
  }
  const pictos = Array.isArray(d.pictos) ? d.pictos : [];
  const renTabela = d.renomeadosTabela && typeof d.renomeadosTabela === 'object' ? d.renomeadosTabela : {};
  for (const o of pictos) {
    if (!PASTA_LISTA[o.lista] || !['excluir', 'editar', 'novo'].includes(o.acao)) return erro('alteração de pictograma inválida');
    for (const n of [o.nome, o.novoNome]) if (n !== undefined && (!limpar(n, 150) || /[\\/:*?"<>|[\]]/.test(n))) return erro(`nome de pictograma inválido: "${n}"`);
    if (o.arquivo && (!/^[A-Za-z0-9+/=]+$/.test(o.arquivo) || o.arquivo.length > MAX * 1.37)) return erro(`a imagem de "${o.novoNome || o.nome}" tem mais de 2 MB`);
    if (o.mini && !/^data:image\/(webp|png|jpeg);base64,[A-Za-z0-9+/=]+$/.test(o.mini)) return erro('miniatura inválida');
    if (o.acao === 'novo' && (!o.arquivo || !o.mini)) return erro(`falta a imagem de "${o.nome}"`);
  }
  if (!tabela && !pictos.length && !Object.keys(renTabela).length) return erro('nenhuma alteração para salvar');
  const resumo = [tabela ? 'tabela de perigos' : '', pictos.length ? `${pictos.length} pictograma(s)` : ''].filter(Boolean).join(' e ');
  const sha = await gravarNoMain(env, `ADMINISTRADOR: ${resumo}`, async base => {
    const out = [];
    if (tabela) out.push(...await arquivosDaTabela(env, base, tabela));
    const renPer = {}, renEpi = {};
    if (pictos.length) {
      let bib = await lerArquivo(env, BIB, base);
      const dirs = {};
      for (const l of Object.keys(PASTA_LISTA)) dirs[l] = (await gh(env, `contents/${caminhoUrl(PASTA_LISTA[l])}?ref=${base}`, { aceitar404: true })) || [];
      const arquivosDe = (l, nome) => dirs[l].filter(f => f.type !== 'dir' && f.name.replace(/\.[^.]+$/, '') === nome);
      const existe = (l, nome) => !!itemNaBiblioteca(bib, l, nome);
      const ext = o => ((o.nomeArquivo || '').match(/\.(png|jpe?g)$/i) || ['.png'])[0].toLowerCase();
      for (const o of pictos) {
        const nome = limpar(o.nome, 150), pasta = PASTA_LISTA[o.lista];
        if (o.acao === 'excluir') {
          bib = tirarDaBiblioteca(bib, o.lista, nome);
          for (const f of arquivosDe(o.lista, nome)) out.push({ caminho: `${pasta}/${f.name}`, apagar: true });
        } else if (o.acao === 'novo') {
          if (existe(o.lista, nome)) throw falha(`já existe um pictograma "${nome}"`);
          bib = colocarNaBiblioteca(bib, o.lista, nome, o.mini).texto;
          out.push({ caminho: `${pasta}/${nome}${ext(o)}`, base64: o.arquivo });
        } else {
          const novo = limpar(o.novoNome ?? nome, 150);
          if (novo !== nome && existe(o.lista, novo)) throw falha(`já existe um pictograma "${novo}"`);
          bib = trocarNaBiblioteca(bib, o.lista, nome, novo, o.mini || null);
          const velhos = arquivosDe(o.lista, nome);
          if (o.arquivo) {
            for (const f of velhos) out.push({ caminho: `${pasta}/${f.name}`, apagar: true });
            out.push({ caminho: `${pasta}/${novo}${ext(o)}`, base64: o.arquivo });
          } else if (novo !== nome) {
            for (const f of velhos) { out.push({ caminho: `${pasta}/${novo}${f.name.slice(nome.length)}`, sha: f.sha }); out.push({ caminho: `${pasta}/${f.name}`, apagar: true }); }
          }
          if (novo !== nome) (o.lista === 'perigos' ? renPer : renEpi)[nome] = novo;
        }
      }
      out.push({ caminho: BIB, texto: bib });
    }
    const renTab = {};
    for (const [a, b] of Object.entries(renTabela)) if (limpar(a, 200) && limpar(b, 200) && a !== b) renTab[limpar(a, 200)] = limpar(b, 200);
    if (Object.keys(renTab).length + Object.keys(renPer).length + Object.keys(renEpi).length) {
      let t = await lerArquivo(env, NOMES_ANTIGOS, base);
      t = colocarRenomeados(t, 'PERIGOS_RENOMEADOS', renTab);
      t = colocarRenomeados(t, 'PERIGOS_NOMES_ANTIGOS', renPer);
      t = colocarRenomeados(t, 'EPIS_NOMES_ANTIGOS', renEpi);
      out.push({ caminho: NOMES_ANTIGOS, texto: t });
    }
    return out;
  });
  return resposta({ ok: true, commit: sha });
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
    if (acao === 'administrar') return await administrar(env, d);
    return erro('endereço desconhecido', 404);
  } catch (x) {
    return x.status ? erro(x.message, x.status) : erro('não consegui terminar: ' + (x.message || x), 500);
  }
}

