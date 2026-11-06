// Abrir um PO em Excel (.xlsx) no gerador.
// 1) Excel exportado pelo próprio gerador: os dados completos estão guardados dentro dele (customXml) e voltam iguais.
// 2) PO no modelo B.00.183 (inclusive os montados à mão): as colunas são achadas pelos títulos (Nº, Nome etapa, Fotos,
//    Pontos de atenção, Descrição, Perigos, EPI's) e cada imagem vai para a etapa e a coluna onde está o seu centro.
'use strict';

const IMPORTAR_XL = (() => {
  const {lerZip, xml, todos, norm, esc, papel, pictogramas, frase, assinatura, correlacao, codigoRev} = IMPORTAR.ajuda;
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const EMU_CM = 360000;

  // ---------------- pasta de trabalho ----------------
  const colNum = s => { let n = 0; for (const ch of s) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };
  const refRC = ref => { const m = /^([A-Z]+)(\d+)$/.exec(ref); return m ? {r: +m[2] - 1, c: colNum(m[1])} : null; };
  const filho = (el, nome) => el && [...el.children].find(n => n.localName === nome);
  const caminho = (base, alvo) => {
    if (alvo.startsWith('/')) return alvo.slice(1);
    const p = base.split('/'); p.pop();
    for (const s of alvo.split('/')) { if (s === '..') p.pop(); else if (s !== '.') p.push(s); }
    return p.join('/');
  };
  async function relsDe(z, parte) {
    const i = parte.lastIndexOf('/');
    const t = await z.texto(parte.slice(0, i) + '/_rels/' + parte.slice(i + 1) + '.rels');
    const out = {};
    if (t) for (const r of todos(xml(t), 'Relationship')) out[r.getAttribute('Id')] = {alvo: caminho(parte, r.getAttribute('Target')), tipo: r.getAttribute('Type') || ''};
    return out;
  }
  // texto rico (sharedStrings / inlineStr) -> {texto, html}
  function rico(si) {
    const runs = [...si.children].filter(n => n.localName === 'r');
    if (!runs.length) { const t = filho(si, 't'); const s = t ? t.textContent : ''; return {texto: s, html: esc(s)}; }
    let texto = '', html = '';
    for (const r of runs) {
      const t = filho(r, 't'); if (!t) continue;
      const s = t.textContent; texto += s;
      const pr = filho(r, 'rPr');
      const tem = n => { const e = filho(pr, n); return !!e && e.getAttribute('val') !== '0' && e.getAttribute('val') !== 'none'; };
      // cada linha fica com sua formatação (a descrição do gerador é uma linha por item)
      html += s.split('\n').map(p => { let h = esc(p); if (!p.trim()) return h; if (tem('b')) h = `<b>${h}</b>`; if (tem('i')) h = `<i>${h}</i>`; if (tem('u')) h = `<u>${h}</u>`; return h; }).join('\n');
    }
    return {texto, html};
  }

  async function lerPasta(z) {
    const wbParte = 'xl/workbook.xml';
    const wb = xml(await z.texto(wbParte));
    const wbRels = await relsDe(z, wbParte);
    const ss = [];
    const ssParte = Object.values(wbRels).find(r => /sharedStrings$/.test(r.tipo));
    if (ssParte) { const t = await z.texto(ssParte.alvo); if (t) for (const si of xml(t).documentElement.children) if (si.localName === 'si') ss.push(rico(si)); }
    const folhas = [];
    for (const sh of todos(wb, 'sheet')) {
      const rel = wbRels[sh.getAttributeNS(R, 'id') || sh.getAttribute('r:id')];
      if (!rel) continue;
      const t = await z.texto(rel.alvo);
      if (!t) continue;
      folhas.push(await lerFolha(z, sh.getAttribute('name'), rel.alvo, xml(t), ss));
    }
    return folhas;
  }

  async function lerFolha(z, nome, parte, doc, ss) {
    const cel = new Map();          // 'r,c' -> {texto, html, f, num}
    let maxR = 0, maxC = 0;
    for (const c of todos(doc, 'c')) {
      const p = refRC(c.getAttribute('r') || ''); if (!p) continue;
      const tipo = c.getAttribute('t'), v = filho(c, 'v'), f = filho(c, 'f');
      let o = null;
      if (tipo === 's' && v) o = ss[+v.textContent] || {texto: '', html: ''};
      else if (tipo === 'inlineStr') { const is = filho(c, 'is'); o = is ? rico(is) : {texto: '', html: ''}; }
      else if (v) o = {texto: v.textContent, html: esc(v.textContent), num: tipo !== 'str' && tipo !== 'e' && tipo !== 'b' ? +v.textContent : undefined};
      if (!o) continue;
      cel.set(p.r + ',' + p.c, Object.assign({}, o, {f: f ? f.textContent : ''}));
      maxR = Math.max(maxR, p.r); maxC = Math.max(maxC, p.c);
    }
    const merges = todos(doc, 'mergeCell').map(m => { const [a, b] = m.getAttribute('ref').split(':').map(refRC); return a && b ? {r0: a.r, c0: a.c, r1: b.r, c1: b.c} : null; }).filter(Boolean);
    // geometria (EMU) para achar a célula do centro de cada imagem
    const fmt = filho(filho(doc, 'worksheet') || doc.documentElement, 'sheetFormatPr');
    const larPadrao = +(fmt && (fmt.getAttribute('defaultColWidth') || fmt.getAttribute('baseColWidth'))) || 8.43;
    const altPadrao = +(fmt && fmt.getAttribute('defaultRowHeight')) || 15;
    const larg = [];
    for (const c of todos(doc, 'col')) for (let i = +c.getAttribute('min') - 1; i < +c.getAttribute('max') && i < 400; i++) larg[i] = c.getAttribute('hidden') === '1' ? 0 : +c.getAttribute('width');
    const alt = [];
    for (const r of todos(doc, 'row')) { const i = +r.getAttribute('r') - 1; alt[i] = r.getAttribute('hidden') === '1' ? 0 : (r.getAttribute('ht') ? +r.getAttribute('ht') : altPadrao); maxR = Math.max(maxR, i); }
    const colX = [0], rowY = [0];
    for (let i = 0; i < Math.max(maxC + 30, 60); i++) colX.push(colX[i] + Math.round(((larg[i] ?? larPadrao) * 7 + 5) * 9525));
    for (let i = 0; i < maxR + 60; i++) rowY.push(rowY[i] + Math.round((alt[i] ?? altPadrao) * 12700));
    const achar = (arr, v) => { let i = 0; while (i < arr.length - 2 && arr[i + 1] <= v) i++; return i; };
    const desenhos = await lerDesenhos(z, parte, doc, colX, rowY);
    for (const d of desenhos) { d.r = achar(rowY, (d.y0 + d.y1) / 2); d.c = achar(colX, (d.x0 + d.x1) / 2); }
    const folha = {nome, cel, merges, maxR, maxC, desenhos, colX, rowY};
    folha.txt = (r, c) => { const o = cel.get(r + ',' + c); return o ? o.texto.replace(/ /g, ' ').trim() : ''; };
    folha.mesclagem = (r, c) => merges.find(m => r >= m.r0 && r <= m.r1 && c >= m.c0 && c <= m.c1) || {r0: r, c0: c, r1: r, c1: c};
    return folha;
  }

  // imagens e formas da folha, com a caixa (EMU) de cada uma
  async function lerDesenhos(z, parte, doc, colX, rowY) {
    const dr = todos(doc, 'drawing')[0];
    if (!dr) return [];
    const rels = await relsDe(z, parte);
    const rel = rels[dr.getAttributeNS(R, 'id') || dr.getAttribute('r:id')];
    if (!rel) return [];
    const t = await z.texto(rel.alvo);
    if (!t) return [];
    const drRels = await relsDe(z, rel.alvo);
    const d = xml(t), out = [];
    const ponto = el => { const n = k => +(filho(el, k) || {}).textContent || 0; return {x: (colX[n('col')] || 0) + n('colOff'), y: (rowY[n('row')] || 0) + n('rowOff')}; };
    const coletar = (el, caixa, mapa) => {
      for (const n of el.children) {
        if (n.localName === 'grpSp') {
          // grupo: as posições dos filhos estão no sistema do grupo (chOff/chExt)
          const xf = todos(filho(n, 'grpSpPr') || n, 'xfrm')[0];
          const v = k => { const e = xf && filho(xf, k); return e ? {x: +(e.getAttribute('x') ?? e.getAttribute('cx')), y: +(e.getAttribute('y') ?? e.getAttribute('cy'))} : null; };
          const off = v('off'), ext = v('ext'), chOff = v('chOff'), chExt = v('chExt');
          const m = off && ext && chOff && chExt && chExt.x && chExt.y ? {sx: ext.x / chExt.x, sy: ext.y / chExt.y, ox: off.x - chOff.x * ext.x / chExt.x, oy: off.y - chOff.y * ext.y / chExt.y} : null;
          const comp = m && mapa ? {sx: m.sx * mapa.sx, sy: m.sy * mapa.sy, ox: mapa.ox + m.ox * mapa.sx, oy: mapa.oy + m.oy * mapa.sy} : (mapa || m);
          coletar(n, caixa, comp);
          continue;
        }
        if (n.localName !== 'pic' && n.localName !== 'sp' && n.localName !== 'cxnSp') continue;
        let bx = caixa;
        const xf = todos(n, 'xfrm')[0];
        if (mapa && xf) {
          const off = filho(xf, 'off'), ext = filho(xf, 'ext');
          if (off && ext) {
            const x = mapa.ox + +off.getAttribute('x') * mapa.sx, y = mapa.oy + +off.getAttribute('y') * mapa.sy;
            bx = {x0: x, y0: y, x1: x + +ext.getAttribute('cx') * mapa.sx, y1: y + +ext.getAttribute('cy') * mapa.sy};
          }
        }
        const pr = todos(n, 'cNvPr')[0];
        const it = Object.assign({}, bx, {tipo: n.localName === 'pic' ? 'img' : 'forma', nome: pr ? pr.getAttribute('name') || '' : ''});
        const geo = todos(n, 'prstGeom')[0]; it.geo = geo ? geo.getAttribute('prst') : '';
        it.texto = todos(n, 't').filter(x => x.namespaceURI === 'http://schemas.openxmlformats.org/drawingml/2006/main').map(x => x.textContent).join('').replace(/\s+/g, ' ').trim();
        const blip = todos(n, 'blip')[0];
        if (blip) {
          const r = drRels[blip.getAttributeNS(R, 'embed') || blip.getAttribute('r:embed')];
          it.alvo = r && r.alvo;
          const sr = todos(n, 'srcRect')[0];
          if (sr) it.corte = ['l', 't', 'r', 'b'].map(k => (+sr.getAttribute(k) || 0) / 100000);
          if (it.tipo !== 'img') it.tipo = 'img';
        }
        out.push(it);
      }
    };
    for (const a of d.documentElement.children) {
      let caixa = null;
      if (a.localName === 'twoCellAnchor') { const p0 = ponto(filho(a, 'from')), p1 = ponto(filho(a, 'to')); caixa = {x0: p0.x, y0: p0.y, x1: p1.x, y1: p1.y}; }
      else if (a.localName === 'oneCellAnchor') { const p0 = ponto(filho(a, 'from')), e = filho(a, 'ext'); caixa = {x0: p0.x, y0: p0.y, x1: p0.x + +e.getAttribute('cx'), y1: p0.y + +e.getAttribute('cy')}; }
      else if (a.localName === 'absoluteAnchor') { const p = filho(a, 'pos'), e = filho(a, 'ext'); caixa = {x0: +p.getAttribute('x'), y0: +p.getAttribute('y'), x1: +p.getAttribute('x') + +e.getAttribute('cx'), y1: +p.getAttribute('y') + +e.getAttribute('cy')}; }
      if (caixa) coletar(a, caixa, null);
    }
    return out;
  }

  // imagem da pasta -> dataURL (aplicando o recorte feito no Excel)
  function imagemDe(z, cache) {
    return async it => {
      if (!it.alvo) return null;
      const chave = it.alvo + '|' + (it.corte || []).join(',');
      if (cache[chave] !== undefined) return cache[chave];
      // EMF e WMF: desenhados pelo leitor próprio (recursos/emf.js); o que não der fica de fora com aviso
      if (/\.(emf|wmf)$/i.test(it.alvo)) {
        let url = null;
        if (typeof EMF !== 'undefined') {
          const d = await z.ler(it.alvo), emf = /\.emf$/i.test(it.alvo);
          try { url = d ? (emf ? EMF.paraPNG : EMF.wmfParaPNG)(d.buffer.slice(d.byteOffset, d.byteOffset + d.byteLength), 1600) : null; } catch (e) { url = null; }
          if (url && it.corte && it.corte.some(v => Math.abs(v) > 0.001)) url = await recortar(url, it.corte, 'image/png');
        }
        if (!url) cache.emf = (cache.emf || 0) + 1;
        return (cache[chave] = url);
      }
      const mime = {png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp'}[it.alvo.split('.').pop().toLowerCase()];
      const dados = mime && await z.ler(it.alvo);
      if (!dados) return (cache[chave] = null);
      let bin = ''; for (let i = 0; i < dados.length; i += 0x8000) bin += String.fromCharCode.apply(null, dados.subarray(i, i + 0x8000));
      let url = `data:${mime};base64,${btoa(bin)}`;
      if (it.corte && it.corte.some(v => Math.abs(v) > 0.001)) url = await recortar(url, it.corte, mime);
      return (cache[chave] = url);
    };
  }
  function recortar(url, [l, t, r, b], mime) {
    return new Promise(ok => {
      const im = new Image();
      im.onload = () => {
        const W = im.naturalWidth, H = im.naturalHeight;
        const sx = W * l, sy = H * t, sw = W * (1 - l - r), sh = H * (1 - t - b);
        if (sw < 2 || sh < 2) return ok(url);
        const k = Math.min(1, 1600 / Math.max(sw, sh));
        const c = document.createElement('canvas'); c.width = Math.round(sw * k); c.height = Math.round(sh * k);
        const g = c.getContext('2d');
        if (mime !== 'image/png') { g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height); }
        g.drawImage(im, sx, sy, sw, sh, 0, 0, c.width, c.height);
        ok(mime === 'image/png' ? c.toDataURL('image/png') : c.toDataURL('image/jpeg', 0.9));
      };
      im.onerror = () => ok(url);
      im.src = url;
    });
  }

  // ---------------- "impressão digital" do texto (Excel exportado e depois alterado) ----------------
  function digitalTextos(textos) {
    const s = norm(textos.join('')).replace(/[^A-Z0-9]/g, '');
    let h = 0x811c9dc5;
    for (let i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 0x01000193) >>> 0; }
    return s.length + '-' + h.toString(16);
  }
  // usado na exportação: texto fixo (não fórmula) de todas as folhas, na ordem
  function digitalFolhas(xmls) {
    return digitalTextos(xmls.map(x => todos(xml(x), 'c').filter(c => c.getAttribute('t') === 'inlineStr').map(c => todos(c, 't').map(t => t.textContent).join('')).join('')));
  }
  async function digitalPasta(z) {
    const wb = xml(await z.texto('xl/workbook.xml')), rels = await relsDe(z, 'xl/workbook.xml');
    const ss = [];
    const ssParte = Object.values(rels).find(r => /sharedStrings$/.test(r.tipo));
    if (ssParte) { const t = await z.texto(ssParte.alvo); if (t) for (const si of xml(t).documentElement.children) if (si.localName === 'si') ss.push(rico(si).texto); }
    const textos = [];
    for (const sh of todos(wb, 'sheet')) {
      const rel = rels[sh.getAttributeNS(R, 'id') || sh.getAttribute('r:id')];
      const t = rel && await z.texto(rel.alvo);
      if (!t) continue;
      textos.push(todos(xml(t), 'c').map(c => {
        if (c.getAttribute('t') === 's') { const v = filho(c, 'v'); return v ? ss[+v.textContent] || '' : ''; }
        if (c.getAttribute('t') === 'inlineStr') return todos(c, 't').map(x => x.textContent).join('');
        return '';
      }).join(''));
    }
    return digitalTextos(textos);
  }

  // ---------------- leitura do PO ----------------
  const ehRotuloSub = t => { const n = norm(t).replace(/[:.]$/, ''); return TIPOS_SUB.find(s => norm(s) === n || (n.length > 3 && norm(s).startsWith(n))); };
  const dataExcel = (o) => {
    if (!o) return '';
    if (o.f && /TODAY|HOJE|NOW|AGORA/i.test(o.f)) return '';
    if (o.num !== undefined && !isNaN(o.num) && o.num > 20000 && o.num < 80000) {
      const d = new Date(Date.UTC(1899, 11, 30) + Math.round(o.num) * 86400000);
      return d.toISOString().slice(0, 10);
    }
    const m = /(\d{1,2})\s*[\/.-]\s*(\d{1,2})\s*[\/.-]?\s*(\d{4}|\d{2}\b)/.exec(o.texto || '');
    if (m) return `${m[3].length === 2 ? '20' + m[3] : m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
    return '';
  };
  // linhas de títulos das colunas das etapas
  function mapaTitulos(f, r) {
    const mapa = {};
    for (let c = 0; c <= f.maxC; c++) {
      const t = f.txt(r, c); if (!t) continue;
      const p = papel(t); if (!p || mapa[p]) continue;
      const m = f.mesclagem(r, c);
      mapa[p] = [m.c0, m.c1];
    }
    if (!(mapa.nome && (mapa.fotos || mapa.desc) && (mapa.perigos || mapa.epis || mapa.desc))) return null;
    // descrição numa coluna sem título (ex.: "Máscara Rev.04"): a coluna vazia no título com texto longo nas linhas de baixo
    if (!mapa.desc) for (let c = 0; c <= f.maxC; c++) {
      const m = f.mesclagem(r, c);
      if (m.c0 !== c || f.txt(r, c) || papelDaColuna(mapa, c)) continue;
      let soma = 0, n = 0;
      for (let k = r + 1; k <= Math.min(f.maxR, r + 12); k++) { const t = f.txt(k, c); if (t) { soma += t.length; n++; } }
      if (n >= 2 && soma / n > 40) { mapa.desc = [m.c0, m.c1]; Object.defineProperty(mapa, 'descDeduzida', {value: true}); break; }
    }
    return mapa;
  }
  const papelDaColuna = (mapa, c) => Object.keys(mapa).find(p => c >= mapa[p][0] && c <= mapa[p][1]) || null;
  // valor à direita de um rótulo ("Elaborado por:" -> nome)
  function valorDe(f, linhas, rotulo) {
    for (const r of linhas) for (let c = 0; c <= f.maxC; c++) {
      const t = f.txt(r, c);
      if (!t || !norm(t).startsWith(norm(rotulo))) continue;
      const resto = t.slice(t.indexOf(':') + 1).trim();
      if (t.includes(':') && resto) return {texto: resto};
      const m = f.mesclagem(r, c);
      for (let c2 = m.c1 + 1; c2 <= Math.min(f.maxC, m.c1 + 4); c2++) {
        const o = f.cel.get(r + ',' + c2);
        if (o && String(o.texto).trim()) return o;
      }
    }
    return null;
  }

  async function importar(arquivo) {
    const buf = await arquivo.arrayBuffer();
    // .xlsx com senha para abrir fica criptografado, no formato do .xls antigo: não dá para ler
    const u8 = new Uint8Array(buf);
    if (u8[0] === 0xD0 && u8[1] === 0xCF) throw new Error(new TextDecoder('utf-16le').decode(u8.subarray(0, Math.min(u8.length, 1 << 20))).includes('EncryptedPackage')
      ? 'a planilha está protegida por senha. Abra no Excel, tire a senha (Arquivo > Informações > Proteger pasta de trabalho > Criptografar com senha, apagar a senha), salve e abra de novo.'
      : 'o arquivo está no formato do Excel antigo, mesmo com o nome .xlsx. Abra no Excel, use Salvar como > Pasta de Trabalho do Excel (.xlsx) e abra de novo.');
    const z = await lerZip(buf);
    // 1) Excel exportado pelo gerador
    let editado = false;
    for (const n of z.nomes.filter(n => /^customXml\/item\d+\.xml$/i.test(n))) {
      const t = await z.texto(n);
      if (!t || !t.includes('urn:gerador-po')) continue;
      const m = /<!\[CDATA\[([\s\S]*)\]\]>/.exec(t.replace(/\]\]\]\]><!\[CDATA\[>/g, ']]>'));
      if (!m) continue;
      const dig = /\btexto="([^"]+)"/.exec(t.slice(0, 300));
      if (dig && await digitalPasta(z) === dig[1]) return {doc: JSON.parse(m[1]), avisos: [], completo: true};
      editado = true;
    }
    // 2) PO no modelo do Excel
    const folhas = await lerPasta(z);
    const cache = {}, img = imagemDe(z, cache);
    const avisos = [];
    if (editado) avisos.push('Este Excel foi exportado pelo gerador e depois alterado no Excel; as informações foram lidas da planilha. Confira o resultado.');
    const novo = novoDoc();
    novo.etapas = []; novo.historico = []; novo.anexos = [];
    const c = novo.cab;
    c.tempoCiclo = 'N/A';
    let anotacoes = 0, pequenas = 0;

    // folhas com a tabela de etapas (a primeira tem o cabeçalho)
    const folhasPO = folhas.map(f => {
      for (let r = 0; r <= Math.min(f.maxR, 60); r++) { const mp = mapaTitulos(f, r); if (mp) return {f, r}; }
      return null;
    }).filter(Boolean);
    if (!folhasPO.length) throw new Error('não encontrei a tabela de etapas (Nº, Nome etapa, Fotos, Descrição, Perigos...)');
    if (folhasPO.some(({f, r}) => mapaTitulos(f, r).descDeduzida)) avisos.push('A coluna de descrição não tinha título na planilha; foi reconhecida pelo conteúdo. Confira.');

    // ---------- cabeçalho ----------
    const {f: f0, r: rTit} = folhasPO[0];
    let rBas = rTit;
    for (let r = 0; r < rTit; r++) for (let cc = 0; cc <= f0.maxC; cc++) if (/EPI.*BASICO/.test(norm(f0.txt(r, cc))) && rBas === rTit) rBas = r;
    const linhasCab = [...Array(rTit).keys()];
    const linhasTopo = [...Array(rBas).keys()];
    let cod = (() => { for (const r of linhasTopo) for (let cc = 0; cc <= f0.maxC; cc++) { const t = f0.txt(r, cc); if (/^([A-Z]{1,4}\.)?(\d{2}\.)?\d{2,4}(-\d{2})?$/.test(t)) return t; } return ''; })();
    if (cod) c.codigo = cod.replace(/^PO\.?\s*/i, '');
    // modelo antigo (JSL): célula "PO.MAN - 12" ou "PO.MAN 01  Rev. 00"
    let revCel = '';
    if (!cod) for (const r of linhasTopo) for (let cc = 0; cc <= f0.maxC && !cod; cc++) {
      const t = f0.txt(r, cc);
      if (!/^\s*PO\b/.test(t) || /PROCEDIMENTO/i.test(t)) continue;
      const cr = codigoRev(t);
      if (cr.codigo) { cod = t; c.codigo = cr.codigo; revCel = cr.revisao; }
    }
    const vRev0 = valorDe(f0, linhasTopo, 'Revisão n') || valorDe(f0, linhasTopo, 'Revisao') || valorDe(f0, linhasTopo, 'Rev');
    const vRev = vRev0 || (revCel ? {texto: revCel} : null);
    const vData = valorDe(f0, linhasTopo, 'Data');
    const vElab = valorDe(f0, linhasTopo, 'Elaborado por'), vAprov = valorDe(f0, linhasTopo, 'Aprovado por'), vTc = valorDe(f0, linhasTopo, 'Tempo de ciclo');
    const revPlanilha = vRev ? String(vRev.texto).replace(/\D/g, '') : '';
    const dataPlanilha = dataExcel(vData);
    if (vData && vData.f && /TODAY|HOJE/i.test(vData.f)) avisos.push('A data do cabeçalho era uma fórmula =HOJE(); foi usada a data do histórico de revisões.');
    const nomeProprio = s => (s = (s || '').replace(/\s+/g, ' ').trim()) && s === s.toUpperCase() ? s.toLowerCase().replace(/(^|\s)(\S)/g, (a, b, x) => b + x.toUpperCase()).replace(/\b(Da|De|Do|Das|Dos|E)\b/g, w => w.toLowerCase()) : s;
    c.elaborado = nomeProprio(vElab ? String(vElab.texto).trim() : '');
    c.aprovado = nomeProprio(vAprov ? String(vAprov.texto).replace(/\s*DATA\s*:[\s\S]*$/i, '').trim() : '');   // nome e "DATA: 30/07/2025" na mesma célula
    if (vTc && String(vTc.texto).trim()) c.tempoCiclo = /^N\.?\s*\/?\s*A\.?$/i.test(String(vTc.texto).trim()) ? 'N/A' : String(vTc.texto).trim();
    const ignorar = /^(PROCEDIMENTO|REVIS|DATA|ELABORADO|APROVADO|TEMPO DE CICLO|EPI|LEGENDA|FERRAMENTA|MATERIA|EQUIPAMENTO)/;
    const usados = new Set([cod, vRev && vRev.texto, vElab && vElab.texto, vAprov && vAprov.texto, vTc && vTc.texto].filter(Boolean).map(s => norm(String(s))));
    const cand = [];
    for (const r of linhasTopo) for (let cc = 0; cc <= f0.maxC; cc++) { const t = f0.txt(r, cc); if (t.length > 4 && !ignorar.test(norm(t)) && !usados.has(norm(t)) && isNaN(+t)) cand.push(t); }
    c.nome = frase((cand.sort((a, b) => b.length - a.length)[0] || '').replace(/\s+/g, ' ').trim());
    if (!c.nome) avisos.push('Não encontrei o nome do PO; preencha no cabeçalho.');
    if (!c.codigo || !cod) avisos.push('Não encontrei o código do PO; preencha no cabeçalho.');
    if (!c.elaborado) avisos.push('"Elaborado por" não encontrado.');
    if (!c.aprovado) avisos.push('"Aprovado por" não encontrado.');
    // EPI's básicos: imagens à direita do rótulo, até a legenda
    if (rBas < rTit) {
      let cBas = 0; for (let cc = 0; cc <= f0.maxC; cc++) if (/EPI.*BASICO/.test(norm(f0.txt(rBas, cc)))) { cBas = cc; break; }
      const mb = f0.mesclagem(rBas, cBas);
      let cFim = f0.maxC + 20;
      for (let cc = mb.c1 + 1; cc <= f0.maxC; cc++) if (f0.txt(rBas, cc)) { cFim = cc - 1; break; }
      const imgs = f0.desenhos.filter(d => d.tipo === 'img' && d.r >= mb.r0 && d.r <= mb.r1 && d.c > mb.c1 && d.c <= cFim).sort((a, b) => a.x0 - b.x0);
      const srcs = (await Promise.all(imgs.map(img))).filter(Boolean);
      if (srcs.length) {
        const itens = await pictogramas(srcs, 'epi', []);
        c.episBasicos = itens.map(it => it.nome ? {nome: it.nome} : {nome: '', img: it.img});
        const nao = itens.filter(it => !it.nome).length;
        if (nao) avisos.push(`EPI's básicos: ${nao} de ${itens.length} imagem(ns) não reconhecida(s) na biblioteca; ficaram com a imagem original.`);
      } else avisos.push("EPI's básicos não encontrados; ficou a lista padrão.");
    }
    // ferramentas, materiais e equipamentos escritos no cabeçalho (usados só se as etapas não tiverem sublinhas)
    const listaCab = {};
    for (const [k, rot] of [['ferramentas', 'FERRAMENTA'], ['materiais', 'MATERIA'], ['equipamentos', 'EQUIPAMENTO']]) {
      const v = valorDe(f0, linhasCab, rot);
      const s = v ? String(v.texto).trim() : '';
      if (s && !/^N\.?\s*\/?\s*A\.?$/i.test(s)) listaCab[k] = s;
    }

    // ---------- etapas ----------
    let cur = null;
    const fotosPend = [];       // {etapa, it}
    for (const {f, r: rIni} of folhasPO) {
      let mapa = null, rAnexo = Infinity;
      const dono = [];          // linha -> etapa
      for (let r = rIni; r <= f.maxR; r++) {
        const mp = mapaTitulos(f, r);
        if (mp) { mapa = mp; continue; }
        if (!mapa) continue;
        const linhaTxt = norm([...Array(f.maxC + 1).keys()].map(cc => f.txt(r, cc)).join(' '));
        if (/^ANEXO\s*\d/.test(linhaTxt)) { rAnexo = Math.min(rAnexo, r); break; }
        const inicio = (cc) => { const m = f.mesclagem(r, cc); return m.r0 === r && m.c0 === cc; };
        const celulas = p => { const out = []; if (!mapa[p]) return out; for (let cc = mapa[p][0]; cc <= mapa[p][1]; cc++) if (inicio(cc) && f.cel.has(r + ',' + cc)) out.push({cc, o: f.cel.get(r + ',' + cc)}); return out; };
        const descs = celulas('desc').filter(x => x.o.texto.trim());
        // sublinha: rótulo (Equipamento, Documentação...) na coluna da descrição e o valor ao lado
        const sub = descs.length && ehRotuloSub(descs[0].o.texto);
        if (sub) {
          const valor = descs.slice(1).map(x => x.o.texto.trim()).join(' ').trim();
          if (cur && valor && !/^N\.?\s*\/?\s*A\.?$/i.test(valor)) cur.sublinhas.push({tipo: sub, valor});
          if (cur) { cur._fechada = true; dono[r] = cur; }
          continue;
        }
        const num = (celulas('num')[0] || {o: {texto: ''}}).o.texto.replace(/\D+/g, ' ').trim().split(' ')[0] || '';
        const nome = celulas('nome').map(x => x.o.texto).join(' ').replace(/\s+/g, ' ').trim();
        if (!num && !nome && !descs.length) { if (cur) dono[r] = cur; continue; }
        const mesma = cur && (num || nome) && (!num || !cur._num || num === cur._num) && (!nome || !cur._nomeOrig || norm(nome) === norm(cur._nomeOrig));
        const outra = cur && ((num && cur._num && num !== cur._num) || (nome && cur._nomeOrig && norm(nome) !== norm(cur._nomeOrig)));
        if (!cur || outra || (cur._fechada && !mesma)) {
          cur = novaEtapa(); cur.pontos = []; cur._fotos = []; cur._legendas = []; cur._per = []; cur._epi = [];
          novo.etapas.push(cur);
        } else cur._fechada = false;
        if (num && !cur._num) cur._num = num;
        if (nome && !cur._nomeOrig) { cur._nomeOrig = nome; cur.nome = nome === nome.toUpperCase() ? frase(nome) : nome.charAt(0).toUpperCase() + nome.slice(1); }
        for (const d of descs) {
          const linhas = d.o.html.replace(/\r/g, '').split('\n').map(h => h.replace(/^(?:\s|&nbsp;)*(?:<[biu]>)?(?:\s)*(?:[•▪●◦\-–]|\d{1,2}\s*[.)-](?!\d)|\d{1,2}(?=[A-ZÀ-Ú]))\s*/, m => (m.match(/<[biu]>/) || [''])[0]).trim())
            .filter(h => h.replace(/<[^>]+>/g, '').trim());
          if (linhas.length) cur.descricao = [cur.descricao, ...linhas].filter(Boolean).join('\n');
        }
        // linhas mescladas a partir desta pertencem à etapa
        for (let cc = 0; cc <= f.maxC; cc++) { const m = f.mesclagem(r, cc); if (m.r0 === r) for (let k = r; k <= m.r1; k++) dono[k] = dono[k] || cur; }
        dono[r] = cur;
      }
      // imagens e formas: etapa pela linha do centro, papel pela coluna do centro
      const rotDono = r => { for (let k = r; k >= rIni; k--) if (dono[k]) return dono[k]; return null; };
      let mapaAt = null;
      const mapas = [];
      for (let r = rIni; r <= Math.min(f.maxR, rAnexo); r++) { const mp = mapaTitulos(f, r); if (mp) mapas.push({r, mp}); }
      for (const d of f.desenhos) {
        if (d.r <= rIni || d.r >= rAnexo) continue;
        mapaAt = (mapas.filter(m => m.r <= d.r).pop() || {}).mp;
        if (!mapaAt) continue;
        const e = rotDono(d.r);
        if (!e) continue;
        const p = papelDaColuna(mapaAt, d.c);
        const wCm = (d.x1 - d.x0) / EMU_CM, hCm = (d.y1 - d.y0) / EMU_CM;
        if (p === 'fotos') {
          if (d.tipo === 'img' && Math.min(wCm, hCm) >= 1.2) e._fotos.push(d);
          else if (d.tipo === 'forma' && /^FOTO\s*\.?\s*\d/i.test(d.texto)) e._legendas.push(d);
          else if (d.tipo === 'img') pequenas++;
          else anotacoes++;
        } else if (p === 'pontos') {
          const id = pontoDe(d);
          if (id) { if (!e.pontos.includes(id)) e.pontos.push(id); }
          else if (d.tipo === 'img') e._pontosImg = (e._pontosImg || []).concat(d);
        } else if (p === 'perigos' && d.tipo === 'img') e._per.push(d);
        else if (p === 'epis' && d.tipo === 'img') e._epi.push(d);
      }
      // anexos ("ANEXO n – título" e as imagens abaixo, até o próximo anexo)
      if (rAnexo < Infinity) {
        const titulos = [];
        for (let r = rAnexo; r <= f.maxR; r++) for (let cc = 0; cc <= f.maxC; cc++) { const m = /ANEXO\s*\d+\s*[–-]\s*(.+)/i.exec(f.txt(r, cc)); if (m) { titulos.push({r, t: m[1].trim()}); break; } }
        for (const [i, a] of titulos.entries()) {
          const fim = titulos[i + 1] ? titulos[i + 1].r : Infinity;
          const imgs = f.desenhos.filter(d => d.tipo === 'img' && d.r > a.r && d.r < fim).sort((x, y) => x.y0 - y.y0 || x.x0 - y.x0);
          const tit = a.t.replace(/\s+/g, ' ');
          const anexo = {titulo: tit === tit.toUpperCase() ? frase(tit) : tit, imagens: []};
          for (const d of imgs.slice(0, MAX_IMG_ANEXO)) { const src = await img(d); if (src) anexo.imagens.push({img: src, linha: anexo.imagens.length < MAX_POR_LINHA ? 1 : 2, aspecto: await aspectoDe(src)}); }
          novo.anexos.push(anexo);
          if (!anexo.imagens.length) avisos.push(`Anexo "${anexo.titulo}": sem imagem que o navegador consiga abrir; insira a imagem no anexo.`);
        }
      }
    }

    // páginas do modelo ainda não usadas: etapa "XXXX" com "[Escolher recurso]" e sem foto nem pictograma
    const marcador = s => !s || /^X{3,}$/i.test(s.trim()) || /^\[.*\]$/.test(s.trim());
    novo.etapas = novo.etapas.filter(e => !(marcador(e._nomeOrig) && !e._fotos.length && !e._per.length && !e._epi.length
      && e.descricao.split('\n').every(l => marcador(l.replace(/<[^>]+>/g, ''))) && e.sublinhas.every(s => marcador(s.valor))));
    for (const e of novo.etapas) e.descricao = e.descricao.split('\n').filter(l => !/^\[.*\]$/.test(l.replace(/<[^>]+>/g, '').trim())).join('\n');
    // fotos com as legendas ("FOTO 01 - ..."): cada legenda vai para a foto mais próxima acima dela
    for (const [i, e] of novo.etapas.entries()) {
      const fotos = e._fotos.sort((a, b) => a.y0 - b.y0 || a.x0 - b.x0);
      const leg = new Map();
      for (const l of e._legendas.sort((a, b) => a.y0 - b.y0)) {
        const cy = (l.y0 + l.y1) / 2, cx = (l.x0 + l.x1) / 2;
        let melhor = null, dist = Infinity;
        for (const fo of fotos) {
          if (leg.has(fo)) continue;
          const fy = (fo.y0 + fo.y1) / 2;
          if (fy > cy) continue;
          const dd = Math.abs(cy - fo.y1) + Math.abs(cx - (fo.x0 + fo.x1) / 2) * 0.5;
          if (dd < dist) { dist = dd; melhor = fo; }
        }
        if (melhor) leg.set(melhor, l.texto.replace(/^FOTO\s*\.?\s*\d+\s*[–.:-]?\s*/i, '').trim());
      }
      for (const fo of fotos) {
        const src = await img(fo);
        if (src) e.fotos.push({img: src, legenda: leg.get(fo) || '', aspecto: await aspectoDe(src)});
      }
      for (const [chave, lista, tipo] of [['_per', 'perigos', 'perigo'], ['_epi', 'epis', 'epi']]) {
        const srcs = (await Promise.all(e[chave].sort((a, b) => a.y0 - b.y0 || a.x0 - b.x0).map(img))).filter(Boolean);
        if (!srcs.length) continue;
        const itens = await pictogramas(srcs, tipo, []);
        e[lista].push(...itens.map(({conferir, novo: _n, ...it}) => it));
        const nao = itens.filter(it => !it.nome && it.novo).length;
        if (nao) avisos.push(`Etapa ${i + 1} (${e.nome || 'sem nome'}): ${nao} de ${itens.length} ${tipo === 'perigo' ? 'perigo(s)' : 'EPI(s)'} não reconhecido(s) na biblioteca; ficou a imagem original, sem nome.`);
      }
      // imagem não reconhecida na coluna dos pontos: no modelo, é o "+" verde de Segurança
      if (e._pontosImg) for (const d of e._pontosImg) { const id = (await pontoPorImagem(await img(d))) || 'seguranca'; if (!e.pontos.includes(id)) e.pontos.push(id); }
      e.pontos = PONTOS.map(p => p.id).filter(id => e.pontos.includes(id));
      if (!e.pontos.length) e.pontos = ['seguranca'];
      for (const k of ['_fotos', '_legendas', '_per', '_epi', '_pontosImg', '_num', '_nomeOrig', '_fechada']) delete e[k];
      if (!e.nome) avisos.push(`Etapa ${i + 1}: sem nome na planilha; preencha o nome da etapa.`);
    }
    if (!novo.etapas.length) {
      avisos.unshift('Não encontrei etapas preenchidas na planilha. O documento foi aberto só com o cabeçalho.');
      novo.etapas.push(novaEtapa());
    }
    if (!novo.etapas.some(e => e.sublinhas.length) && Object.keys(listaCab).length) { c.listasAuto = false; Object.assign(c, {ferramentas: 'N/A', materiais: 'N/A', equipamentos: 'N/A'}, listaCab); }
    if (anotacoes) avisos.push(`${anotacoes} marcação(ões) desenhada(s) sobre as fotos (setas, círculos, textos) não foram levadas; se forem importantes, marque na própria foto antes de inserir.`);
    if (pequenas) avisos.push(`${pequenas} imagem(ns) pequena(s) sobre as fotos (ícones, setas) não foram levadas.`);
    if (cache.emf) avisos.push(`${cache.emf} imagem(ns) em formato antigo (WMF/EMF) que não consegui converter ficaram de fora.`);

    // ---------- histórico de revisões ----------
    for (const f of folhas) {
      let rh = -1, mapa = null;
      for (let r = 0; r <= Math.min(f.maxR, 30) && !mapa; r++) {
        const m = {};
        for (let cc = 0; cc <= f.maxC; cc++) {
          const t = norm(f.txt(r, cc)); if (!t) continue;
          const mg = f.mesclagem(r, cc), faixa = [mg.c0, mg.c1];
          if (/^ITEM/.test(t)) m.item = faixa; else if (/DESCRI/.test(t)) m.descricao = faixa; else if (/FOTO/.test(t)) m.foto = faixa;
          else if (/^DATA/.test(t)) m.data = faixa; else if (/SOLICIT/.test(t)) m.solicitante = faixa; else if (/RESPONS/.test(t)) m.responsavel = faixa; else if (/OBSERV/.test(t)) m.obs = faixa;
        }
        if (m.descricao && m.data && (m.item || m.responsavel)) { mapa = m; rh = r; }
      }
      if (!mapa) continue;
      const valor = (r, k) => { if (!mapa[k]) return null; for (let cc = mapa[k][0]; cc <= mapa[k][1]; cc++) { const o = f.cel.get(r + ',' + cc); if (o && String(o.texto).trim()) return o; } return null; };
      for (let r = rh + 1; r <= f.maxR; r++) {
        const d = valor(r, 'descricao');
        if (!d) continue;
        const h = {descricao: String(d.texto).trim(), foto: '', data: dataExcel(valor(r, 'data')), solicitante: nomeProprio(String((valor(r, 'solicitante') || {texto: ''}).texto).trim()),
          responsavel: nomeProprio(String((valor(r, 'responsavel') || {texto: ''}).texto).trim()), obs: String((valor(r, 'obs') || {texto: ''}).texto).trim()};
        const ft = valor(r, 'foto'); if (ft) h.foto = String(ft.texto).trim();
        if (mapa.foto) {
          const im = f.desenhos.find(x => x.tipo === 'img' && x.r === r && x.c >= mapa.foto[0] && x.c <= mapa.foto[1]);
          const src = im && await img(im);
          if (src) h.fotoImg = {img: src, aspecto: await aspectoDe(src)};
        }
        novo.historico.push(h);
      }
      break;
    }
    // revisão e data: pelo histórico, quando ele bate com o cabeçalho
    const ultimo = novo.historico.length - 1;
    if (ultimo >= 0 && (!revPlanilha || +revPlanilha === ultimo)) c.revAuto = true;
    else {
      c.revAuto = false;
      c.revisao = revPlanilha ? revPlanilha.padStart(2, '0') : '00';
      c.data = dataPlanilha || (ultimo >= 0 ? novo.historico[ultimo].data : '');
      if (ultimo >= 0) avisos.push(`A revisão do cabeçalho (${c.revisao}) não bate com o histórico (último item ${ultimo}); ficou a do cabeçalho.`);
    }
    if (!novo.historico.length) {
      novo.historico.push({descricao: 'Importado do Excel: ' + arquivo.name, foto: '', data: dataPlanilha || '', solicitante: '', responsavel: '', obs: ''});
      avisos.push('Histórico de revisões não encontrado; foi criado um item com a importação.');
    }
    // lista de treinamento
    novo.treinamento.incluir = folhas.some(f => /ASSINATURA|TREINAMENTO/.test(norm(f.nome)));

    const nPer = novo.etapas.reduce((s, e) => s + e.perigos.length + e.epis.length, 0);
    if (nPer) avisos.push('Confira os nomes embaixo dos perigos e EPIs: no Excel eles fazem parte da imagem, então o gerador usou o nome do pictograma da biblioteca. Para trocar um pictograma, clique nele no editor.');
    return {doc: novo, avisos, completo: false};
  }

  // pontos de atenção: forma "+" verde = Segurança; nomes das imagens/formas; ou comparação com os ícones do gerador
  function pontoDe(d) {
    const n = norm(d.nome + ' ' + d.texto);
    if (/MEIO AMBIENTE|AMBIENT/.test(n)) return 'meio_ambiente';
    if (/ENERGIA/.test(n)) return 'energia';
    if (/QUALIDADE/.test(n) || d.geo === 'diamond') return 'qualidade';
    if (/SEGURANCA|^MAIS\b|PLUS|CRUZ/.test(n) || d.geo === 'plus' || d.geo === 'mathPlus') return 'seguranca';
    return null;
  }
  async function pontoPorImagem(src) {
    if (!src) return null;
    const a = await assinatura(src);
    if (!a) return null;
    let melhor = null, nota = 0.6;
    for (const p of PONTOS) {
      const ref = BIB.icones && BIB.icones[p.id];
      if (!ref) continue;
      const b = await assinatura(ref);
      const cc = b && correlacao(a, b);
      if (cc > nota) { nota = cc; melhor = p.id; }
    }
    return melhor;
  }

  return {importar, digitalFolhas};
})();
