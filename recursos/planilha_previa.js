// Desenha uma aba do Excel (.xlsx montado pelo gerador) como página da prévia: larguras, alturas, mesclagens,
// fontes, cores, bordas, alinhamentos e imagens da própria planilha, na escala da impressão do Excel.
// Usado na prévia da APR para as abas Ordem de Serviço, Controle de revisão e Matriz de Risco.
'use strict';

const PLANILHA = (() => {
  const xml = s => new DOMParser().parseFromString(s, 'application/xml');
  const tags = (el, nome) => el ? [...el.getElementsByTagNameNS('*', nome)] : [];
  const tag = (el, nome) => tags(el, nome)[0] || null;
  const filhos = (el, nome) => el ? [...el.children].filter(c => c.localName === nome) : [];
  const RID = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const MM_PT = 72 / 25.4;
  // o Excel imprime as colunas um pouco mais largas que na tela (medido no PDF da APR: 1,12)
  const FATOR_COL = 1.12;
  // abas em que o fator medido no PDF do Excel é outro (Ordem de Serviço: a tabela ocupa a largura toda, 0,5 cm dos lados)
  const FATOR_ABA = {'Ordem de Serviço': 1.147};

  // paleta "indexed" padrão do Excel
  const INDEXED = ['000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
    '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080', '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
    '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF', '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
    '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696', '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333'];
  function comTint(hex, t) {
    if (!t) return hex;
    let r = parseInt(hex.slice(0, 2), 16) / 255, g = parseInt(hex.slice(2, 4), 16) / 255, b = parseInt(hex.slice(4, 6), 16) / 255;
    const mx = Math.max(r, g, b), mn = Math.min(r, g, b);
    let h = 0, s = 0, l = (mx + mn) / 2;
    if (mx !== mn) {
      const d = mx - mn; s = l > 0.5 ? d / (2 - mx - mn) : d / (mx + mn);
      h = mx === r ? (g - b) / d + (g < b ? 6 : 0) : mx === g ? (b - r) / d + 2 : (r - g) / d + 4; h /= 6;
    }
    l = t < 0 ? l * (1 + t) : l * (1 - t) + t;
    const f = (p, q, x) => { if (x < 0) x += 1; if (x > 1) x -= 1; return x < 1 / 6 ? p + (q - p) * 6 * x : x < 1 / 2 ? q : x < 2 / 3 ? p + (q - p) * (2 / 3 - x) * 6 : p; };
    if (!s) r = g = b = l; else { const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q; r = f(p, q, h + 1 / 3); g = f(p, q, h); b = f(p, q, h - 1 / 3); }
    return [r, g, b].map(v => Math.round(v * 255).toString(16).padStart(2, '0')).join('').toUpperCase();
  }
  function lerTema(t) {
    if (!t) return [];
    const x = xml(t), cs = tag(x, 'clrScheme');
    const cor = n => { const e = filhos(cs, n)[0]; if (!e) return '000000'; const c = e.children[0]; return (c.getAttribute('lastClr') || c.getAttribute('val') || '000000').toUpperCase(); };
    // ordem do Excel: 0 = lt1, 1 = dk1, 2 = lt2, 3 = dk2, 4..9 = accent1..6
    return [cor('lt1'), cor('dk1'), cor('lt2'), cor('dk2'), cor('accent1'), cor('accent2'), cor('accent3'), cor('accent4'), cor('accent5'), cor('accent6'), cor('hlink'), cor('folHlink')];
  }
  function corDe(e, tema) {
    if (!e) return null;
    const t = parseFloat(e.getAttribute('tint') || '0');
    if (e.getAttribute('rgb')) return '#' + comTint(e.getAttribute('rgb').slice(-6).toUpperCase(), t);
    if (e.getAttribute('theme') !== null) return '#' + comTint(tema[+e.getAttribute('theme')] || '000000', t);
    if (e.getAttribute('indexed') !== null) { const i = +e.getAttribute('indexed'); return i === 64 ? null : '#' + (INDEXED[i] || '000000'); }
    return null;
  }
  function lerFonte(f, tema) {
    if (!f) return {};
    const v = n => { const e = filhos(f, n)[0]; return e ? (e.getAttribute('val') ?? true) : null; };
    const u = filhos(f, 'u')[0];
    return {b: !!filhos(f, 'b')[0] && v('b') !== '0', i: !!filhos(f, 'i')[0] && v('i') !== '0', u: !!u && (u.getAttribute('val') || 'single') !== 'none',
      sz: parseFloat(v('sz')) || 11, nome: v('name') || v('rFont') || 'Arial', cor: corDe(filhos(f, 'color')[0], tema)};
  }
  function lerEstilos(t, tema) {
    const x = xml(t);
    const fmts = {};
    for (const n of tags(x, 'numFmt')) fmts[n.getAttribute('numFmtId')] = n.getAttribute('formatCode');
    const fonts = filhos(tag(x, 'fonts'), 'font').map(f => lerFonte(f, tema));
    const fills = filhos(tag(x, 'fills'), 'fill').map(f => {
      const p = tag(f, 'patternFill');
      if (!p || !p.getAttribute('patternType') || p.getAttribute('patternType') === 'none') return null;
      return corDe(filhos(p, 'fgColor')[0], tema) || corDe(filhos(p, 'bgColor')[0], tema);
    });
    const borders = filhos(tag(x, 'borders'), 'border').map(b => {
      const lado = n => { const e = filhos(b, n)[0]; const st = e && e.getAttribute('style'); return st ? {st, cor: corDe(filhos(e, 'color')[0], tema) || '#000000'} : null; };
      return {l: lado('left'), r: lado('right'), t: lado('top'), b: lado('bottom')};
    });
    const xfs = filhos(tag(x, 'cellXfs'), 'xf').map(xf => {
      const al = filhos(xf, 'alignment')[0];
      return {font: fonts[+xf.getAttribute('fontId') || 0] || {}, fill: fills[+xf.getAttribute('fillId') || 0] || null, borda: borders[+xf.getAttribute('borderId') || 0] || {},
        h: al && al.getAttribute('horizontal') || 'general', v: al && al.getAttribute('vertical') || 'bottom', wrap: !!al && al.getAttribute('wrapText') === '1',
        rot: al ? +(al.getAttribute('textRotation') || 0) : 0, fmt: +xf.getAttribute('numFmtId') || 0, fmtCodigo: fmts[xf.getAttribute('numFmtId')] || ''};
    });
    return {xfs, fonts};
  }
  // texto compartilhado: lista de trechos {t, fonte} (texto rico) ou um trecho só
  function lerTextos(t, tema) {
    if (!t) return [];
    return filhos(xml(t).documentElement, 'si').map(si => {
      const rs = filhos(si, 'r');
      if (!rs.length) return [{t: tags(si, 't').filter(e => !e.closest('rPh')).map(e => e.textContent).join('')}];
      return rs.map(r => ({t: tags(r, 't').map(e => e.textContent).join(''), fonte: filhos(r, 'rPr')[0] ? lerFonte(filhos(r, 'rPr')[0], tema) : null}));
    });
  }
  const colNum = l => l.split('').reduce((s, c) => s * 26 + c.charCodeAt(0) - 64, 0);
  const ref = r => { const m = /^\$?([A-Z]+)\$?(\d+)$/.exec(r); return m ? {c: colNum(m[1]), r: +m[2]} : null; };
  function intervalo(s) { const [a, b] = s.split(':'); const x = ref(a), y = ref(b || a); return {c1: x.c, r1: x.r, c2: y.c, r2: y.r}; }
  function formatar(v, xf) {
    const n = Number(v);
    const data = [14, 15, 16, 17, 22].includes(xf.fmt) || /d+[\/.-]m+|m+[\/.-]d+|yy/i.test(xf.fmtCodigo);
    if (data && !isNaN(n)) { const d = new Date(Date.UTC(1899, 11, 30) + Math.round(n * 86400000)); return `${String(d.getUTCDate()).padStart(2, '0')}/${String(d.getUTCMonth() + 1).padStart(2, '0')}/${d.getUTCFullYear()}`; }
    return String(v);
  }
  // espessura das linhas já impressas (px), iguais às da tabela da APR: fina 1 px e média 2 px na escala de 57%
  const LARG_BORDA = {thin: 0.57, hair: 0.4, dotted: 0.57, dashed: 0.57, dashDot: 0.57, dashDotDot: 0.57, medium: 0.85, mediumDashed: 0.85, mediumDashDot: 0.85, mediumDashDotDot: 0.85, slantDashDot: 0.85, thick: 1.14, double: 1.14};
  const TIPO_BORDA = st => /dash/i.test(st) ? 'dashed' : st === 'dotted' || st === 'hair' ? 'dotted' : st === 'double' ? 'double' : 'solid';

  // arqs: {caminho: Uint8Array} do .xlsx; devolve o elemento da página (ou null)
  async function pagina(arqs, nomeFolha, imagens = {}) {
    const dec = new TextDecoder(), T = n => arqs[n] ? dec.decode(arqs[n]) : '';
    const tema = lerTema(T('xl/theme/theme1.xml'));
    const est = lerEstilos(T('xl/styles.xml'), tema);
    const ss = lerTextos(T('xl/sharedStrings.xml'), tema);
    const wb = xml(T('xl/workbook.xml'));
    const folhas = tags(wb, 'sheet');
    const k = folhas.findIndex(s => s.getAttribute('name') === nomeFolha);
    if (k < 0) return null;
    const rels = {};
    for (const r of tags(xml(T('xl/_rels/workbook.xml.rels')), 'Relationship')) rels[r.getAttribute('Id')] = r.getAttribute('Target');
    const caminho = 'xl/' + rels[folhas[k].getAttributeNS(RID, 'id') || folhas[k].getAttribute('r:id')].replace(/^\/?(xl\/)?/, '');
    const sh = xml(T(caminho));
    // área de impressão
    const dn = tags(wb, 'definedName').find(d => d.getAttribute('name') === '_xlnm.Print_Area' && d.getAttribute('localSheetId') === String(k));
    let area = dn ? intervalo(dn.textContent.split('!').pop().replace(/'/g, '')) : null;
    if (!area) { const dim = tag(sh, 'dimension'); area = intervalo(dim ? dim.getAttribute('ref') : 'A1:A1'); }
    // colunas e linhas
    const fmt = tag(sh, 'sheetFormatPr');
    const larguraPadrao = parseFloat(fmt && fmt.getAttribute('defaultColWidth')) || 8.43, alturaPadrao = parseFloat(fmt && fmt.getAttribute('defaultRowHeight')) || 15;
    const colW = {}, colOculta = {};
    for (const c of tags(sh, 'col')) for (let i = +c.getAttribute('min'); i <= Math.min(+c.getAttribute('max'), 200); i++) { colW[i] = parseFloat(c.getAttribute('width')); if (c.getAttribute('hidden') === '1') colOculta[i] = true; }
    const ptCol = i => colOculta[i] ? 0 : Math.trunc((256 * (colW[i] ?? larguraPadrao) + 18) / 256 * 7) * 0.75 * (FATOR_ABA[nomeFolha] || FATOR_COL);
    const rowH = {}, rowOculta = {}, celulas = {};
    for (const row of tags(sh, 'row')) {
      const r = +row.getAttribute('r');
      if (row.getAttribute('ht')) rowH[r] = parseFloat(row.getAttribute('ht'));
      if (row.getAttribute('hidden') === '1') rowOculta[r] = true;
      for (const c of filhos(row, 'c')) celulas[c.getAttribute('r')] = c;
    }
    const ptRow = r => rowOculta[r] ? 0 : (rowH[r] ?? alturaPadrao);
    const X = [0], Y = [0];
    for (let c = area.c1; c <= area.c2; c++) X.push(X[X.length - 1] + ptCol(c));
    for (let r = area.r1; r <= area.r2; r++) Y.push(Y[Y.length - 1] + ptRow(r));
    const xDe = c => X[c - area.c1], yDe = r => Y[r - area.r1];
    const W = X[X.length - 1], H = Y[Y.length - 1];
    // mesclagens
    const merges = tags(sh, 'mergeCell').map(m => intervalo(m.getAttribute('ref')));
    const mergeDe = {}, dentro = {};
    for (const m of merges) for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) { if (r === m.r1 && c === m.c1) mergeDe[r + ':' + c] = m; else dentro[r + ':' + c] = m; }
    // escala da impressão
    const ps = tag(sh, 'pageSetup'), po = tag(sh, 'printOptions'), pm = tag(sh, 'pageMargins'), spr = tag(sh, 'pageSetUpPr');
    const deitada = !ps || ps.getAttribute('orientation') !== 'portrait';
    const pgW = (deitada ? 297 : 210) * MM_PT, pgH = (deitada ? 210 : 297) * MM_PT;
    const mg = n => (parseFloat(pm && pm.getAttribute(n)) || 0.5) * 72;
    const uW = pgW - mg('left') - mg('right'), uH = pgH - mg('top') - mg('bottom');
    let esc = (parseFloat(ps && ps.getAttribute('scale')) || 100) / 100;
    if (spr && spr.getAttribute('fitToPage') === '1') {
      const fw = ps && ps.getAttribute('fitToWidth') !== null ? +ps.getAttribute('fitToWidth') : 1, fh = ps && ps.getAttribute('fitToHeight') !== null ? +ps.getAttribute('fitToHeight') : 1;
      esc = Math.min(fw ? uW / W : 9, fh ? uH / H : 9);
      esc = Math.floor(esc * 100) / 100;
    }
    esc = Math.min(esc, uW / W, 4);
    const ox = mg('left') + (po && po.getAttribute('horizontalCentered') === '1' ? (uW - W * esc) / 2 : 0);
    const oy = mg('top') + (po && po.getAttribute('verticalCentered') === '1' ? (uH - H * esc) / 2 : 0);

    const folha = el('div', {class: 'xl-folha', style: `width:${W}pt;height:${H}pt;zoom:${esc}`});
    const fundo = el('div', {class: 'xl-camada'}), linhas_ = el('div', {class: 'xl-camada'}), textos = el('div', {class: 'xl-camada'});
    folha.append(fundo, linhas_, textos);
    const estilo = c => est.xfs[+(c && c.getAttribute('s')) || 0] || est.xfs[0];
    // fundos e textos (por célula ou mesclagem)
    for (let r = area.r1; r <= area.r2; r++) for (let c = area.c1; c <= area.c2; c++) {
      if (dentro[r + ':' + c]) continue;
      const m = mergeDe[r + ':' + c] || {r1: r, c1: c, r2: r, c2: c};
      const r2 = Math.min(m.r2, area.r2), c2 = Math.min(m.c2, area.c2);
      const x = xDe(c), y = yDe(r), w = xDe(c2 + 1) - x, h = yDe(r2 + 1) - y;
      if (!w || !h) continue;
      const nomeRef = String.fromCharCode(...(n => { const s = []; while (n > 0) { s.unshift(65 + (n - 1) % 26); n = Math.floor((n - 1) / 26); } return s; })(c)) + r;
      const cel = celulas[nomeRef], xf = estilo(cel);
      if (xf.fill) fundo.append(el('div', {style: `left:${x}pt;top:${y}pt;width:${w}pt;height:${h}pt;background:${xf.fill}`}));
      if (!cel) continue;
      const t = cel.getAttribute('t'), v = filhos(cel, 'v')[0];
      let trechos = null, numero = false;
      if (t === 's' && v) trechos = ss[+v.textContent] || [];
      else if (t === 'inlineStr') { const is = filhos(cel, 'is')[0]; const rs = filhos(is, 'r'); trechos = rs.length ? rs.map(rr => ({t: tags(rr, 't').map(e => e.textContent).join(''), fonte: filhos(rr, 'rPr')[0] ? lerFonte(filhos(rr, 'rPr')[0], tema) : null})) : [{t: tags(is, 't').map(e => e.textContent).join('')}]; }
      else if (v && t !== 'e') { if (t === 'str' || t === 'b') trechos = [{t: v.textContent}]; else { trechos = [{t: formatar(v.textContent, xf)}]; numero = true; } }
      if (!trechos || !trechos.some(p => p.t)) continue;
      const f = xf.font;
      const hz = xf.h === 'general' ? (numero ? 'right' : 'left') : xf.h === 'centerContinuous' ? 'center' : xf.h === 'justify' || xf.h === 'distributed' ? 'justify' : xf.h;
      const vt = {top: 'flex-start', center: 'center', bottom: 'flex-end', justify: 'center', distributed: 'center'}[xf.v] || 'flex-end';
      const caixa = el('div', {class: 'xl-cel' + (xf.wrap || xf.h === 'justify' ? ' quebra' : ''), style: `left:${x}pt;top:${y}pt;width:${w}pt;height:${h}pt;align-items:${vt};${vt === 'flex-start' && xf.wrap ? 'padding-top:1.5pt;' : ''}` +
        `justify-content:${hz === 'right' ? 'flex-end' : hz === 'center' ? 'center' : 'flex-start'};text-align:${hz};font-family:'${f.nome || 'Arial'}',Arial,sans-serif;font-size:${f.sz || 11}pt;` +
        `font-weight:${f.b ? 700 : 400};font-style:${f.i ? 'italic' : 'normal'};text-decoration:${f.u ? 'underline' : 'none'};color:${f.cor || '#000'}`});
      const txt = el('div', {class: 'xl-txt' + (xf.rot === 90 ? ' rot90' : xf.rot === 180 ? ' rot270' : '')});
      for (const p of trechos) {
        const g = p.fonte;
        txt.append(g ? el('span', {style: `font-weight:${g.b ? 700 : 400};font-style:${g.i ? 'italic' : 'normal'};text-decoration:${g.u ? 'underline' : 'none'};font-size:${g.sz}pt;${g.cor ? 'color:' + g.cor : ''}`}, p.t) : p.t);
      }
      caixa.append(txt);
      textos.append(caixa);
    }
    // bordas: cada célula desenha as suas (sem as linhas de dentro das mesclagens); os pedaços seguidos viram uma linha só
    const segs = [];
    for (let r = area.r1; r <= area.r2; r++) for (let c = area.c1; c <= area.c2; c++) {
      const nomeRef = String.fromCharCode(...(n => { const s = []; while (n > 0) { s.unshift(65 + (n - 1) % 26); n = Math.floor((n - 1) / 26); } return s; })(c)) + r;
      const cel = celulas[nomeRef];
      if (!cel) continue;
      const b = estilo(cel).borda, m = dentro[r + ':' + c] || mergeDe[r + ':' + c];
      const x = xDe(c), y = yDe(r), w = xDe(c + 1) - x, h = yDe(r + 1) - y;
      if (!w || !h) continue;
      const lado = (bd, interno, hor, p, a0, a1) => { if (!bd || interno) return; segs.push({hor, p, a0, a1, lw: (LARG_BORDA[bd.st] || 0.57) / esc * 0.75, tipo: TIPO_BORDA(bd.st), cor: bd.cor}); };
      lado(b.t, m && r > m.r1, true, y, x, x + w);
      lado(b.b, m && r < m.r2, true, y + h, x, x + w);
      lado(b.l, m && c > m.c1, false, x, y, y + h);
      lado(b.r, m && c < m.c2, false, x + w, y, y + h);
    }
    const chave = s => [s.hor, s.p.toFixed(2), s.lw.toFixed(3), s.tipo, s.cor].join('|');
    const grupos = new Map();
    for (const s of segs) { const k = chave(s); if (!grupos.has(k)) grupos.set(k, []); grupos.get(k).push(s); }
    for (const lista of grupos.values()) {
      lista.sort((a, b) => a.a0 - b.a0);
      const unidos = [];
      for (const s of lista) { const u = unidos[unidos.length - 1]; if (u && s.a0 <= u.a1 + 0.5) u.a1 = Math.max(u.a1, s.a1); else unidos.push(Object.assign({}, s)); }
      for (const s of unidos) linhas_.append(el('div', {style: `border:0 ${s.tipo} ${s.cor};` + (s.hor
        ? `left:${s.a0}pt;top:${s.p - s.lw / 2}pt;width:${s.a1 - s.a0}pt;height:0;border-top-width:${s.lw}pt`
        : `left:${s.p - s.lw / 2}pt;top:${s.a0}pt;width:0;height:${s.a1 - s.a0}pt;border-left-width:${s.lw}pt`)}));
    }
    // imagens da aba (logotipo)
    const relSh = T(caminho.replace('worksheets/', 'worksheets/_rels/') + '.rels');
    const dr = tag(sh, 'drawing');
    if (dr && relSh) {
      const rmap = {}; for (const r of tags(xml(relSh), 'Relationship')) rmap[r.getAttribute('Id')] = r.getAttribute('Target');
      const dpath = 'xl/' + (rmap[dr.getAttributeNS(RID, 'id') || dr.getAttribute('r:id')] || '').replace(/^\.\.\//, '');
      const dxml = T(dpath), drel = T(dpath.replace('drawings/', 'drawings/_rels/') + '.rels');
      if (dxml) {
        const dm = {}; if (drel) for (const r of tags(xml(drel), 'Relationship')) dm[r.getAttribute('Id')] = r.getAttribute('Target');
        for (const a of [...xml(dxml).documentElement.children]) {
          const blip = tag(a, 'blip'); if (!blip) continue;
          const alvo = 'xl/' + (dm[blip.getAttributeNS(RID, 'embed') || blip.getAttribute('r:embed')] || '').replace(/^\.\.\//, '');
          if (!arqs[alvo]) continue;
          const fr = tag(a, 'from'), to = tag(a, 'to'), ext = tag(tag(a, 'xfrm'), 'ext');
          const pos = e => ({c: +tag(e, 'col').textContent + 1, co: +tag(e, 'colOff').textContent / 12700, r: +tag(e, 'row').textContent + 1, ro: +tag(e, 'rowOff').textContent / 12700});
          const p1 = pos(fr);
          if (p1.c < area.c1 || p1.r < area.r1) continue;
          const x1 = xDe(p1.c) + p1.co, y1 = yDe(p1.r) + p1.ro;
          let w = ext ? +ext.getAttribute('cx') / 12700 : 0, h = ext ? +ext.getAttribute('cy') / 12700 : 0;
          if (to && a.localName === 'twoCellAnchor' && a.getAttribute('editAs') !== 'oneCell') { const p2 = pos(to); w = xDe(Math.min(p2.c, area.c2 + 1)) + p2.co - x1; h = yDe(Math.min(p2.r, area.r2 + 1)) + p2.ro - y1; }
          if (!imagens[alvo]) imagens[alvo] = URL.createObjectURL(new Blob([arqs[alvo]], {type: /\.jpe?g$/i.test(alvo) ? 'image/jpeg' : 'image/png'}));
          textos.append(el('img', {src: imagens[alvo], style: `position:absolute;left:${x1}pt;top:${y1}pt;width:${w}pt;height:${h}pt`}));
        }
      }
    }
    return el('div', {class: 'pagina xl-pagina' + (deitada ? '' : ' retrato'), 'data-aba': nomeFolha},
      el('div', {class: 'xl-pos', style: `left:${ox}pt;top:${oy}pt`}, folha));
  }

  document.head.append(el('style', {}, `
.xl-pagina{padding:0!important;position:relative;overflow:hidden}
.xl-pagina.retrato{width:210mm;height:296.8mm}
.xl-pos{position:absolute}
.xl-folha{position:relative}
.xl-camada{position:absolute;inset:0}
.xl-camada>div,.xl-cel{position:absolute;box-sizing:border-box}
.xl-cel{display:flex;padding:0 2pt;line-height:1.12;white-space:pre;overflow:visible}
.xl-cel.quebra{white-space:pre-wrap;overflow:hidden}
.xl-cel.quebra .xl-txt{width:100%;overflow-wrap:break-word}
.xl-txt.rot90{writing-mode:vertical-rl;transform:rotate(180deg)}
.xl-txt.rot270{writing-mode:vertical-rl}
`));
  return {pagina};
})();
