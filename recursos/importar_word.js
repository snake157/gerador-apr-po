// Abrir um PO em Word (.docx) no gerador.
// 1) Word exportado pelo próprio gerador: os dados completos estão guardados dentro dele (customXml) e voltam iguais.
// 2) Outros modelos: a tabela de etapas é localizada pelos títulos das colunas (Nº, Etapa, Foto, Descrição/Segurança,
//    Perigos, EPI's...) e os campos são associados da melhor forma possível. O que não der para encaixar vira aviso.
'use strict';

const IMPORTAR = (() => {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

  // ---------------- zip ----------------
  async function lerZip(buf) {
    const dv = new DataView(buf);
    let e = buf.byteLength - 22;
    while (e >= 0 && dv.getUint32(e, true) !== 0x06054b50) e--;
    // o mesmo leitor abre .docx e .xlsx (05/11/2026: a mensagem falava ".docx" também para Excel)
    if (e < 0) throw new Error('o arquivo está danificado ou não é um arquivo do Word/Excel válido');
    const n = dv.getUint16(e + 10, true);
    let p = dv.getUint32(e + 16, true);
    const arqs = {}, dec = new TextDecoder();
    for (let i = 0; i < n; i++) {
      const metodo = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
      const nl = dv.getUint16(p + 28, true), el = dv.getUint16(p + 30, true), cl = dv.getUint16(p + 32, true), lo = dv.getUint32(p + 42, true);
      arqs[dec.decode(new Uint8Array(buf, p + 46, nl))] = {metodo, csize, lo};
      p += 46 + nl + el + cl;
    }
    async function ler(nome) {
      const a = arqs[nome];
      if (!a) return null;
      const nl = dv.getUint16(a.lo + 26, true), el = dv.getUint16(a.lo + 28, true);
      const dados = new Uint8Array(buf, a.lo + 30 + nl + el, a.csize);
      if (a.metodo === 0) return dados;
      const s = new Blob([dados]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
      return new Uint8Array(await new Response(s).arrayBuffer());
    }
    return {nomes: Object.keys(arqs), ler, texto: async nome => { const d = await ler(nome); return d ? dec.decode(d) : null; }};
  }

  // ---------------- utilidades ----------------
  const xml = s => new DOMParser().parseFromString(s, 'application/xml');
  const filhos = (el, nome) => [...el.childNodes].filter(n => n.nodeType === 1 && n.localName === nome);
  const todos = (el, nome) => [...el.getElementsByTagNameNS('*', nome)];
  const atrib = (el, nome) => el && (el.getAttributeNS(W, nome) || el.getAttribute('w:' + nome));
  const norm = s => (s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().replace(/\s+/g, ' ').trim();
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const CORES_REALCE = {yellow: '#FFFF00', green: '#00FF00', cyan: '#00FFFF', magenta: '#FF00FF', red: '#FF0000', blue: '#0000FF',
    darkYellow: '#808000', darkGreen: '#008000', darkCyan: '#008080', darkMagenta: '#800080', darkRed: '#800000', darkBlue: '#000080',
    lightGray: '#C0C0C0', darkGray: '#808080', black: '#000000'};

  // texto de um parágrafo (tabs e quebras viram espaço / nova linha)
  function textoPar(p) {
    let t = '';
    for (const n of p.getElementsByTagNameNS('*', '*')) {
      if (n.localName === 't' && n.namespaceURI === W) t += n.textContent;
      else if (n.localName === 'tab' && n.namespaceURI === W && n.parentNode.localName === 'r') t += ' ';
      else if ((n.localName === 'br' || n.localName === 'cr') && n.namespaceURI === W) t += '\n';
    }
    return t;
  }
  const parsDe = tc => todos(tc, 'p').filter(p => p.namespaceURI === W && !p.closest?.('txbxContent'));
  const textoCel = tc => parsDe(tc).map(textoPar).join('\n').replace(/ /g, ' ').trim();

  // parágrafo -> HTML com negrito/itálico/sublinhado/realce (o formato que a descrição do gerador usa)
  function htmlPar(p) {
    let h = '';
    // todos os trechos do parágrafo, também os de dentro de revisões (w:ins), campos e marcas; não os de caixas de texto
    const runs = todos(p, 'r').filter(r => r.namespaceURI === W && !r.closest?.('txbxContent') && !r.closest?.('Fallback') && !r.closest?.('del'));
    for (const r of runs) {
      let t = '';
      for (const n of r.childNodes) {
        if (n.nodeType !== 1) continue;
        if (n.localName === 't') t += n.textContent; else if (n.localName === 'tab') t += ' ';
        else if (n.localName === 'br' || n.localName === 'cr') t += '\n';     // quebra de linha dentro do parágrafo = item novo
      }
      if (!t) continue;
      const pr = filhos(r, 'rPr')[0];
      const tem = nome => { const e = pr && filhos(pr, nome)[0]; if (!e) return false; const v = atrib(e, 'val'); return v !== '0' && v !== 'false' && v !== 'none'; };
      let s = esc(t.replace(/ /g, ' '));
      const hl = pr && filhos(pr, 'highlight')[0], shd = pr && filhos(pr, 'shd')[0];
      const fundo = hl ? CORES_REALCE[atrib(hl, 'val')] : (shd && /^[0-9A-F]{6}$/i.test(atrib(shd, 'fill') || '') && atrib(shd, 'fill') !== 'FFFFFF' ? '#' + atrib(shd, 'fill') : null);
      // cada pedaço entre quebras de linha recebe a formatação, para as tags não atravessarem a quebra
      h += s.split('\n').map(seg => {
        if (!seg) return '';
        if (tem('b')) seg = '<b>' + seg + '</b>';
        if (tem('i')) seg = '<i>' + seg + '</i>';
        if (tem('u')) seg = '<u>' + seg + '</u>';
        if (fundo) seg = `<mark style="background-color:${fundo}">${seg}</mark>`;
        return seg;
      }).join('\n');
    }
    return h.trim();
  }

  // ---------------- tabelas com mesclagens ----------------
  // devolve as células com coluna inicial/final na grade e as continuações de mesclagem vertical
  function grade(tbl) {
    const linhas = filhos(tbl, 'tr');
    return linhas.map((tr, ri) => {
      let col = 0;
      const pr = filhos(tr, 'trPr')[0];
      const gb = pr && filhos(pr, 'gridBefore')[0];
      if (gb) col += +atrib(gb, 'val') || 0;
      return filhos(tr, 'tc').map(tc => {
        const tcPr = filhos(tc, 'tcPr')[0];
        const span = +(atrib(tcPr && filhos(tcPr, 'gridSpan')[0], 'val') || 1);
        const vm = tcPr && filhos(tcPr, 'vMerge')[0];
        const c = {tc, ri, c0: col, c1: col + span - 1, cont: !!vm && atrib(vm, 'val') !== 'restart'};
        col += span;
        return c;
      });
    });
  }
  const sobrepoe = (c, a, b) => c.c0 <= b && c.c1 >= a;

  // ---------------- imagens ----------------
  const VML = 'urn:schemas-microsoft-com:vml';
  // imagens na ordem do documento: as do Word novo (a:blip) e as do Word antigo (v:imagedata).
  // A cópia "Fallback" de um mc:AlternateContent é a mesma imagem repetida, então fica de fora.
  // recorte do Word: a:srcRect (1/100000) ou, no Word antigo, cropleft/croptop... do v:imagedata (fração ou "nnnf" = n/65536)
  function corteDe(n) {
    if (n.localName === 'blip') {
      const sr = n.parentNode && filhos(n.parentNode, 'srcRect')[0];
      if (!sr) return null;
      const c = ['l', 't', 'r', 'b'].map(k => (+(sr.getAttribute(k) || 0)) / 100000);
      return c.some(v => Math.abs(v) > 0.005) ? c : null;
    }
    const fr = s => { if (!s) return 0; s = String(s).trim(); return /f$/i.test(s) ? parseFloat(s) / 65536 : parseFloat(s) || 0; };
    const c = ['cropleft', 'croptop', 'cropright', 'cropbottom'].map(k => fr(n.getAttribute(k)));
    return c.some(v => Math.abs(v) > 0.005) ? c : null;
  }
  // tamanho da imagem na página, em cm (0 se não der para saber)
  function tamanhoDe(n) {
    if (n.localName === 'blip') {
      for (let p = n.parentNode; p && p.nodeType === 1; p = p.parentNode) {
        if ((p.localName === 'inline' || p.localName === 'anchor') && p.namespaceURI === WP) {
          const ext = filhos(p, 'extent')[0];
          return ext ? [+ext.getAttribute('cx') / 360000, +ext.getAttribute('cy') / 360000] : [0, 0];
        }
      }
      return [0, 0];
    }
    const st = (n.parentNode && n.parentNode.getAttribute && n.parentNode.getAttribute('style')) || '';
    const med = k => { const m = new RegExp('(?:^|;)\\s*' + k + ':\\s*([\\d.]+)(pt|in|cm|mm|px)?', 'i').exec(st); if (!m) return 0; return +m[1] * ({pt: 0.03528, in: 2.54, cm: 1, mm: 0.1, px: 0.02646}[(m[2] || 'px').toLowerCase()]); };
    return [med('width'), med('height')];
  }
  function refsImagem(no, ctx) {
    const out = [];
    for (const n of no.getElementsByTagNameNS('*', '*')) {
      let id = null;
      if (n.localName === 'blip' && n.namespaceURI === A) id = n.getAttributeNS(R, 'embed') || n.getAttribute('r:embed');
      else if (n.localName === 'imagedata' && n.namespaceURI === VML) id = n.getAttributeNS(R, 'id') || n.getAttribute('r:id');
      else continue;
      if (n.closest?.('Fallback')) continue;
      if (id && ctx.rels[id]) out.push({alvo: ctx.rels[id], corte: corteDe(n), tam: tamanhoDe(n)});
    }
    return out;
  }
  // imagens com o recorte do Word aplicado (ex.: um quadro com vários EPIs recortado para mostrar só um)
  async function urlImagem({alvo, corte}, ctx) {
    const url = await ctx.dataUrl(alvo);
    if (!url || !corte) return url;
    const chave = alvo + '|' + corte.join(',');
    if (!ctx.recortes) ctx.recortes = {};
    return ctx.recortes[chave] || (ctx.recortes[chave] = await recortar(url, corte));
  }
  async function imagensDe(tc, ctx) {
    const out = [];
    for (const ref of refsImagem(tc, ctx)) { const url = await urlImagem(ref, ctx); if (url) out.push(url); }
    return out;
  }

  const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
  // aplica o recorte do Word (frações da imagem cortadas à esquerda, em cima, à direita e embaixo)
  function recortar(src, [l, tp, r, b]) {
    return new Promise(ok => {
      const im = new Image();
      im.onload = () => {
        const W0 = im.naturalWidth, H0 = im.naturalHeight;
        const sx = Math.max(0, l) * W0, sy = Math.max(0, tp) * H0, sw = W0 * (1 - Math.max(0, l) - Math.max(0, r)), sh = H0 * (1 - Math.max(0, tp) - Math.max(0, b));
        if (sw < 4 || sh < 4) return ok(src);
        const c = document.createElement('canvas'); c.width = Math.round(sw); c.height = Math.round(sh);
        const g = c.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height);
        g.drawImage(im, sx, sy, sw, sh, 0, 0, c.width, c.height);
        ok(c.toDataURL('image/jpeg', 0.9));
      };
      im.onerror = () => ok(src);
      im.src = src;
    });
  }

  // ---------------- formas do Word sobre as fotos (05/11/2026) ----------------
  // Círculos, retângulos, setas, linhas e caixas de texto desenhados no Word por cima das fotos viram formas do
  // DESENHAR (editáveis no AJUSTAR FOTO). O Word não diz em qual foto a forma está: a posição de cada foto e de cada
  // forma na célula é refeita aqui (fotos em linha com o texto, uma ao lado da outra; fotos soltas pela posição
  // gravada) e a forma vai para a foto com que mais se sobrepõe. Foto solta, posição contada da página, forma girada
  // ou só em parte sobre a foto: a foto fica marcada "conferir formas". Formas atrás da foto (escondidas no Word) e
  // grupos (pictograma + nome) ficam de fora.
  const EMU_MM = 36000;
  const ESQUEMA = {bg1: 'FFFFFF', lt1: 'FFFFFF', bg2: 'E7E6E6', lt2: 'E7E6E6', tx1: '000000', dk1: '000000', tx2: '44546A', dk2: '44546A',
    accent1: '4472C4', accent2: 'ED7D31', accent3: 'A5A5A5', accent4: 'FFC000', accent5: '5B9BD5', accent6: '70AD47'};
  const PRST = {red: 'FF0000', yellow: 'FFFF00', white: 'FFFFFF', black: '000000', blue: '0000FF', green: '00FF00', orange: 'FFA500'};
  const filho = (el, nome) => el ? filhos(el, nome)[0] : undefined;
  const dentroDeCaixa = (n, p) => { const t = n.parentNode && n.parentNode.closest && n.parentNode.closest('txbxContent'); return !!(t && p.contains(t)); };
  function corDe(el) {
    if (!el) return null;
    if (filho(el, 'noFill')) return 'none';
    const c = todos(el, 'srgbClr')[0]; if (c) return c.getAttribute('val');
    const e = todos(el, 'schemeClr')[0]; if (e) return ESQUEMA[e.getAttribute('val')] || '4472C4';
    const y = todos(el, 'sysClr')[0]; if (y) return y.getAttribute('lastClr') || '000000';
    const p = todos(el, 'prstClr')[0]; if (p) return PRST[p.getAttribute('val')] || null;
    return null;
  }
  function lerForma(wsp) {
    const spPr = filho(wsp, 'spPr');
    if (!spPr) return null;
    const g = filho(spPr, 'prstGeom'), geo = g ? g.getAttribute('prst') : filho(spPr, 'custGeom') ? 'livre' : '';
    const xf = filho(spPr, 'xfrm');
    const rot = xf ? (+(xf.getAttribute('rot') || 0)) / 60000 : 0;
    const flipH = !!xf && xf.getAttribute('flipH') === '1', flipV = !!xf && xf.getAttribute('flipV') === '1';
    const ln = filho(spPr, 'ln'), estilo = filho(wsp, 'style');
    const lnRef = filho(estilo, 'lnRef'), fillRef = filho(estilo, 'fillRef');
    let linha = ln ? corDe(ln) : null;
    if (linha == null) linha = lnRef && lnRef.getAttribute('idx') !== '0' ? corDe(lnRef) || '000000' : 'none';
    const esp = ln && ln.getAttribute('w') ? +ln.getAttribute('w') / 12700 : .75;   // pt
    const ehLinha = /^(line|straightConnector1)$/.test(geo);
    let fundo = null;
    if (filho(spPr, 'noFill')) fundo = 'none';
    else for (const k of ['solidFill', 'gradFill', 'pattFill']) if (filho(spPr, k)) { fundo = corDe(filho(spPr, k)); break; }
    if (fundo == null) fundo = ehLinha ? 'none' : fillRef && fillRef.getAttribute('idx') !== '0' ? corDe(fillRef) || 'none' : 'none';
    const ponta = k => { const e = filho(ln, k); return !!(e && e.getAttribute('type') && e.getAttribute('type') !== 'none'); };
    const tx = todos(wsp, 'txbxContent')[0];
    const ps = tx ? todos(tx, 'p').filter(p => p.namespaceURI === W) : [];
    const texto = ps.map(textoPar).join('\n').replace(/\n+$/, '').trim();
    let tamPt = 11, corTxt = null, alinha = 'esq';
    if (texto) {
      const sz = todos(tx, 'sz').find(x => x.namespaceURI === W); if (sz && +atrib(sz, 'val')) tamPt = +atrib(sz, 'val') / 2;
      const c = todos(tx, 'color').find(x => x.namespaceURI === W && atrib(x, 'val') && atrib(x, 'val') !== 'auto'); if (c) corTxt = atrib(c, 'val');
      const jc = todos(tx, 'jc').find(x => x.namespaceURI === W);
      alinha = {center: 'centro', right: 'dir', end: 'dir', both: 'just', distribute: 'just'}[jc && atrib(jc, 'val')] || 'esq';
    }
    return {geo, rot, flipH, flipV, linha, esp, fundo, cabeca: ln ? ponta('headEnd') : false, cauda: ln ? ponta('tailEnd') : false, texto, tamPt, corTxt, alinha};
  }
  // posição (mm, a partir do canto da área útil da célula) de cada foto e de cada forma da célula
  function mapaCelula(tc, pars) {
    const w = todos(tc, 'tcW').find(x => x.namespaceURI === W && x.parentNode.parentNode === tc);
    const larg = w && atrib(w, 'type') === 'dxa' ? +atrib(w, 'w') / 1440 * 25.4 - 2 * 108 / 1440 * 25.4 : 80;
    const caixas = new Map(), formas = [];
    let yPar = 0;
    for (const p of pars) {
      const pPr = filho(p, 'pPr'), jcEl = filho(pPr, 'jc'), jc = jcEl ? atrib(jcEl, 'val') : 'left';
      const sp = filho(pPr, 'spacing'), y0 = yPar + (sp ? (+atrib(sp, 'before') || 0) / 1440 * 25.4 : 0);
      const linhas = []; let cur = [], x = 0;
      for (const n of p.getElementsByTagNameNS('*', '*')) {
        if (n.localName === 'inline' && n.namespaceURI === WP) {
          if (n.closest?.('Fallback') || dentroDeCaixa(n, p)) continue;
          const ext = filho(n, 'extent'); if (!ext) continue;
          const wi = +ext.getAttribute('cx') / EMU_MM, hi = +ext.getAttribute('cy') / EMU_MM;
          if (cur.length && x + wi > larg + .5) { linhas.push(cur); cur = []; x = 0; }
          cur.push({no: n, w: wi, h: hi, x}); x += wi;
        } else if (n.localName === 'br' && n.namespaceURI === W && n.parentNode.parentNode === p) { linhas.push(cur); cur = []; x = 0; }
        else if (n.localName === 't' && n.namespaceURI === W && n.parentNode.parentNode === p) x += n.textContent.length * 1.8;   // texto ao lado (aproximado)
      }
      if (cur.length) linhas.push(cur);
      let y = y0;
      for (const ln of linhas) {
        const tot = ln.reduce((s, i) => s + i.w, 0), hmax = Math.max(4.2, ...ln.map(i => i.h));
        const off = jc === 'center' ? (larg - tot) / 2 : jc === 'right' || jc === 'end' ? larg - tot : 0;
        for (const i of ln) caixas.set(i.no, {x: i.x + off, y: y + hmax - i.h, w: i.w, h: i.h, solta: false});
        y += hmax + .4;
      }
      if (!linhas.length) y += 4.6;
      for (const an of todos(p, 'anchor')) {
        if (an.namespaceURI !== WP || an.closest?.('Fallback') || dentroDeCaixa(an, p)) continue;
        const ph = filho(an, 'positionH'), pv = filho(an, 'positionV'), ext = filho(an, 'extent');
        if (!ph || !pv || !ext) continue;
        const wa = +ext.getAttribute('cx') / EMU_MM, ha = +ext.getAttribute('cy') / EMU_MM;
        const rh = ph.getAttribute('relativeFrom'), rv = pv.getAttribute('relativeFrom');
        const oh = filho(ph, 'posOffset'), ov = filho(pv, 'posOffset'), ah = filho(ph, 'align');
        let ax;
        if (oh) ax = +oh.textContent / EMU_MM;
        else if (ah) ax = ah.textContent === 'center' ? (larg - wa) / 2 : ah.textContent === 'right' ? larg - wa : 0;
        else continue;
        const ay = ov ? +ov.textContent / EMU_MM + (rv === 'paragraph' || rv === 'line' ? y0 : 0) : y0;
        const cx = {x: ax, y: ay, w: wa, h: ha, solta: true, incerto: !['column', 'character'].includes(rh) || !['paragraph', 'line'].includes(rv), atras: an.getAttribute('behindDoc') === '1'};
        const grupo = todos(an, 'wgp').length > 0;
        if (!grupo && todos(an, 'blip').some(b => b.namespaceURI === A)) { caixas.set(an, cx); continue; }
        if (grupo) { formas.push(Object.assign({grupo: true}, cx)); continue; }
        const wsp = todos(an, 'wsp')[0];
        const d = wsp && lerForma(wsp);
        if (d) formas.push(Object.assign(d, cx));
      }
      yPar = y;
    }
    return {caixas, formas};
  }
  // forma do Word -> forma do DESENHAR, com a posição em frações da caixa da foto (x pela largura, y pela altura) e
  // as espessuras em mm; a conversão para as unidades da foto (1000 = largura) é feita quando o formato da foto é conhecido
  function converterForma(fm, b) {
    const fx = x => (x - b.x) / b.w, fy = y => (y - b.y) / b.h;
    const cor = c => '#' + String(c).toUpperCase();
    const cx = fm.x + fm.w / 2, cy = fm.y + fm.h / 2;
    const giro = ((fm.rot % 360) + 360) % 360, a = giro * Math.PI / 180;
    const girar = (x, y) => [cx + (x - cx) * Math.cos(a) - (y - cy) * Math.sin(a), cy + (x - cx) * Math.sin(a) + (y - cy) * Math.cos(a)];
    const dePara = (x1, y1, x2, y2, extra) => {
      const [ax, ay] = girar(x1, y1), [bx, by] = girar(x2, y2);
      return Object.assign({x1: fx(ax), y1: fy(ay), x2: fx(bx), y2: fy(by)}, extra);
    };
    const t = fm.geo;
    if (fm.texto) {
      const fundo = fm.fundo !== 'none', tamMM = fm.tamPt * .3528;
      const c = fm.corTxt ? cor(fm.corTxt) : fm.linha !== 'none' ? cor(fm.linha) : '#000000';
      return {tipo: 'texto', x1: fx(fm.x), y1: fy(fm.y), x2: 0, y2: 0, cor: c, espMM: .5, tamMM, fundo, alinha: fm.alinha, texto: fm.texto,
        largMM: Math.max(tamMM * 2, fm.w - tamMM * .7)};
    }
    if (/^(ellipse|flowChartConnector|rect|roundRect|flowChartProcess|flowChartAlternateProcess)$/.test(t)) {
      const cheia = fm.fundo !== 'none';
      if (!cheia && fm.linha === 'none') return null;
      let w = fm.w, h = fm.h;
      if (Math.abs(giro - 90) < 20 || Math.abs(giro - 270) < 20) [w, h] = [h, w];
      return {tipo: /ellipse|Connector/.test(t) ? 'circ' : 'ret', x1: fx(cx - w / 2), y1: fy(cy - h / 2), x2: fx(cx + w / 2), y2: fy(cy + h / 2),
        cor: cor(cheia ? fm.fundo : fm.linha), espMM: fm.linha === 'none' ? .1 : fm.esp * .3528, cheia: cheia || undefined};
    }
    if (/^(line|straightConnector1)$/.test(t)) {
      if (fm.linha === 'none') return null;
      let x1 = fm.x, y1 = fm.y, x2 = fm.x + fm.w, y2 = fm.y + fm.h;
      if (fm.flipH) [x1, x2] = [x2, x1];
      if (fm.flipV) [y1, y2] = [y2, y1];
      if (fm.cabeca && !fm.cauda) { [x1, x2] = [x2, x1]; [y1, y2] = [y2, y1]; }   // ponta só no começo: vira para o fim
      return dePara(x1, y1, x2, y2, {tipo: fm.cabeca || fm.cauda ? 'seta' : 'linha', cor: cor(fm.linha), espMM: Math.max(.3, fm.esp * .3528)});
    }
    const BLOCO = {rightArrow: [1, 0], leftArrow: [-1, 0], upArrow: [0, -1], downArrow: [0, 1]};
    if (BLOCO[t]) {
      let [dx, dy] = BLOCO[t];
      if (fm.flipH) dx = -dx;
      if (fm.flipV) dy = -dy;
      const L = dx ? fm.w : fm.h, grossa = dx ? fm.h : fm.w;
      const c = fm.fundo !== 'none' ? fm.fundo : fm.linha;
      if (c === 'none') return null;
      return dePara(cx - dx * L / 2, cy - dy * L / 2, cx + dx * L / 2, cy + dy * L / 2, {tipo: 'seta', cor: cor(c), espMM: Math.max(.5, grossa * .4)});
    }
    return null;
  }
  // frações + mm -> unidades da foto (1000 = largura da foto; o y também é contado na largura)
  function formasNaFoto(lista, larguraMM, aspecto) {
    const k = 1000 / larguraMM, ky = 1000 / aspecto;
    return lista.map(c => {
      const f = Object.assign({}, c, {x1: c.x1 * 1000, y1: c.y1 * ky, x2: c.x2 * 1000, y2: c.y2 * ky, esp: c.espMM * k, tam: (c.tamMM || 4.2) * k});
      if (c.largMM) f.larg = c.largMM * k;
      for (const x of ['espMM', 'tamMM', 'largMM']) delete f[x];
      if (f.cheia === undefined) delete f.cheia;
      if (f.tipo !== 'texto') { f.fundo = true; f.alinha = 'esq'; }
      return f;
    });
  }
  // liga cada forma da célula à foto em que ela está
  function formasNasFotos(mapa, fotos, ctx) {
    const comCaixa = fotos.filter(f => f.caixa);
    let nTexto = 0;
    for (const fm of mapa.formas) {
      if (fm.atras) continue;   // atrás da foto: no Word ela nem aparece
      let melhor = null, area = 0;
      for (const ft of comCaixa) {
        const b = ft.caixa;
        const ix = Math.max(0, Math.min(fm.x + fm.w, b.x + b.w) - Math.max(fm.x, b.x)), iy = Math.max(0, Math.min(fm.y + fm.h, b.y + b.h) - Math.max(fm.y, b.y));
        const cx = fm.x + fm.w / 2, cy = fm.y + fm.h / 2;
        const ar = fm.w * fm.h > 0.01 ? ix * iy : (cx >= b.x && cx <= b.x + b.w && cy >= b.y && cy <= b.y + b.h ? 1 : 0);
        if (ar > area) { melhor = ft; area = ar; }
      }
      if (!melhor) continue;   // não está sobre nenhuma foto (ex.: seta apontando para o texto)
      const frac = fm.w * fm.h > 0.01 ? area / (fm.w * fm.h) : 1;
      const conv = !fm.grupo && frac >= .5 ? converterForma(fm, melhor.caixa) : null;
      if (!conv) { ctx.formasFora = (ctx.formasFora || 0) + 1; continue; }
      (melhor.formas = melhor.formas || []).push(conv);
      if (fm.incerto || melhor.caixa.solta || melhor.caixa.incerto || frac < .9 || (fm.rot % 90 && conv.tipo !== 'seta' && conv.tipo !== 'linha')) melhor.conferir = true;
      if (conv.tipo === 'texto') nTexto++;
    }
    return nTexto;
  }
  async function fotosDaCelula(tc, ctx) {
    const grupos = [];          // [{fotos: [{src, w, h, recorte}], legenda}]
    let pendentes = [];
    const pars = parsDe(tc);
    if (todos(tc, 'tbl').length) {       // fotos em tabela dentro da célula: uma por linha, na ordem
      for (const src of await imagensDe(tc, ctx)) grupos.push({fotos: [{src}], legenda: ''});
      const legs = textoCel(tc).split('\n').flatMap(s => s.split(/(?=FOTO\s*\d+)/i)).map(s => s.trim()).filter(s => /^FOTO\s*\d+/i.test(s));
      legs.forEach((l, k) => { if (grupos[k]) grupos[k].legenda = l; });
    } else {
      const mapa = mapaCelula(tc, pars);
      for (const p of pars) {
        const fotos = [];
        for (const dr of todos(p, 'drawing')) {
          const b = todos(dr, 'blip').find(x => x.namespaceURI === A);
          if (!b) continue;
          const alvo = ctx.rels[b.getAttributeNS(R, 'embed') || b.getAttribute('r:embed')];
          if (!alvo || dr.closest?.('Fallback')) continue;
          const src = await ctx.dataUrl(alvo);
          if (!src) continue;
          const ext = todos(dr, 'extent').find(x => x.namespaceURI === WP);
          const sr = todos(dr, 'srcRect')[0];
          const corte = sr ? ['l', 't', 'r', 'b'].map(k => (+(sr.getAttribute(k) || 0)) / 100000) : null;
          const recortada = corte && corte.some(v => Math.abs(v) > 0.005) ? await recortar(src, corte) : src;
          const no = filhos(dr, 'inline')[0] || filhos(dr, 'anchor')[0];
          fotos.push({src: recortada, w: ext ? +ext.getAttribute('cx') / 360000 : 0, h: ext ? +ext.getAttribute('cy') / 360000 : 0, caixa: no && mapa.caixas.get(no)});
        }
        if (fotos.length) { const g = {fotos, legenda: ''}; grupos.push(g); pendentes.push(g); }
        // legendas do parágrafo (pode haver duas lado a lado: "FOTO 21 – ...   FOTO 22 – ...")
        const legs = textoPar(p).split(/(?=FOTO\s*\d+)/i).map(s => s.trim()).filter(s => /^FOTO\s*\d+/i.test(s));
        for (const l of legs) {
          const g = pendentes.find(x => !x.legenda);
          if (g) g.legenda = l;
          // uma legenda para várias linhas de fotos (ex.: 2 × 2): as outras ficam sem legenda
          if (legs.length === 1) pendentes = [];
        }
        if (legs.length > 1) pendentes = [];
      }
      const nTexto = mapa.formas.length ? formasNasFotos(mapa, grupos.flatMap(g => g.fotos), ctx) : 0;
      // sem "FOTO nn": o texto das caixas de texto sobre as fotos ("ABRINDO BOBINA") vira a legenda, se houver um por foto
      if (grupos.length && grupos.every(g => !g.legenda)) {
        const caixas = [];
        for (const tx of todos(tc, 'txbxContent')) {
          if (tx.closest?.('Fallback') || tx.parentNode.closest?.('txbxContent')) continue;
          const s = todos(tx, 'p').map(textoPar).join(' ').replace(/\s+/g, ' ').trim();
          if (s && caixas[caixas.length - 1] !== s) caixas.push(s);
        }
        if (caixas.length === grupos.length) {
          caixas.forEach((s, k) => { grupos[k].legenda = s; });
          // virou legenda: a caixa de texto não vai também como forma
          for (const g of grupos) for (const f of g.fotos) if (f.formas) f.formas = f.formas.filter(c => c.tipo !== 'texto');
        } else if (caixas.length > nTexto) ctx.marcacoes = (ctx.marcacoes || 0) + caixas.length - nTexto;
      }
    }
    // imagens em outro formato (caixa de texto, desenho antigo...): volta ao jeito simples, uma foto por linha, sem perder nenhuma
    const todas = await imagensDe(tc, ctx);
    if (grupos.reduce((s, g) => s + g.fotos.length, 0) < todas.length) {
      const legendas = textoCel(tc).split('\n').map(s => s.trim()).filter(Boolean);
      return todas.map((src, k) => ({fotos: [{src}], legenda: legendas[k] || ''}));
    }
    return grupos;
  }

  // ---------------- reconhecimento de pictogramas pela imagem ----------------
  function carregar(src) { return new Promise(ok => { const i = new Image(); i.onload = () => ok(i); i.onerror = () => ok(null); i.src = src; }); }
  // assinatura: acha a área colorida do pictograma (triângulo amarelo, círculo azul...), ignorando o nome escrito
  // em preto embaixo; usa só o miolo (o símbolo) para não confundir pictogramas que têm a mesma moldura.
  async function assinatura(src, variantes) {
    const im = await carregar(src);
    if (!im) return null;
    const N = 128, c = document.createElement('canvas'), k = Math.min(1, N / Math.max(im.naturalWidth, im.naturalHeight));
    c.width = Math.max(1, Math.round(im.naturalWidth * k)); c.height = Math.max(1, Math.round(im.naturalHeight * k));
    const VV = window.IMPORTAR_VAR || {};
    const g = c.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height); g.drawImage(im, 0, 0, c.width, c.height);
    const d = g.getImageData(0, 0, c.width, c.height).data;
    let x0 = c.width, y0 = c.height, x1 = -1, y1 = -1, n = 0;
    for (let y = 0; y < c.height; y++) for (let x = 0; x < c.width; x++) {
      const i = (y * c.width + x) * 4, mx = Math.max(d[i], d[i + 1], d[i + 2]), mn = Math.min(d[i], d[i + 1], d[i + 2]);
      if (mx - mn > 70) { n++; if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y; }
    }
    if (n < 20) return null;
    const bw = x1 - x0 + 1, bh = y1 - y0 + 1;
    // a imagem do Word pode ter a moldura mais grossa/fina ou o símbolo deslocado: testa pequenas variações de recorte
    const escalas = variantes ? VV.esc || VAR_ESC : [1], desl = variantes ? VV.desl || VAR_DESL : [0];
    const S = VV.S || 32, s = document.createElement('canvas'); s.width = s.height = S;
    const gs = s.getContext('2d'); gs.imageSmoothingQuality = 'high';
    const out = [];
    for (const e of escalas) for (const dx of desl) for (const dy of desl) {
      const cw = bw * 0.6 * e, ch = bh * 0.58 * e, cx = x0 + bw * (0.5 + dx) - cw / 2, cy = y0 + bh * (0.59 + dy) - ch / 2;
      gs.fillStyle = '#fff'; gs.fillRect(0, 0, S, S);
      gs.drawImage(c, cx, cy, cw, ch, 0, 0, S, S);
      const p = gs.getImageData(0, 0, S, S).data, v = new Float32Array(S * S);
      let m = 0;
      for (let i = 0; i < S * S; i++) { v[i] = 0.299 * p[i * 4] + 0.587 * p[i * 4 + 1] + 0.114 * p[i * 4 + 2]; m += v[i]; }
      m /= S * S;
      let sd = 0; for (let i = 0; i < v.length; i++) sd += (v[i] - m) ** 2;
      sd = Math.sqrt(sd / v.length) || 1;
      for (let i = 0; i < v.length; i++) v[i] = (v[i] - m) / sd;
      out.push(v);
    }
    return variantes ? out : out[0];
  }
  // assinatura da imagem INTEIRA (símbolo + nome escrito), 48×48: distingue "Luva de vaqueta" de "Luva multitato",
  // que têm o mesmo desenho. Usada pelos pictogramas aprendidos/treinados (aprendidos.js).
  async function assinaturaInteira(src) {
    const im = await carregar(src); if (!im) return null;
    const c = document.createElement('canvas'), k = Math.min(1, 200 / Math.max(im.naturalWidth, im.naturalHeight));
    c.width = Math.max(1, Math.round(im.naturalWidth * k)); c.height = Math.max(1, Math.round(im.naturalHeight * k));
    const g = c.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height); g.drawImage(im, 0, 0, c.width, c.height);
    const d = g.getImageData(0, 0, c.width, c.height).data; let x0 = 1e9, y0 = 1e9, x1 = -1, y1 = -1;
    for (let y = 0; y < c.height; y++) for (let x = 0; x < c.width; x++) { const i = (y * c.width + x) * 4; if (d[i] + d[i + 1] + d[i + 2] < 600) { x0 = Math.min(x0, x); x1 = Math.max(x1, x); y0 = Math.min(y0, y); y1 = Math.max(y1, y); } }
    if (x1 < 0) return null;
    const S = 48, s = document.createElement('canvas'); s.width = s.height = S; const gs = s.getContext('2d');
    gs.fillStyle = '#fff'; gs.fillRect(0, 0, S, S); gs.drawImage(c, x0, y0, x1 - x0 + 1, y1 - y0 + 1, 0, 0, S, S);
    const p = gs.getImageData(0, 0, S, S).data, v = new Float32Array(S * S); let m = 0;
    for (let i = 0; i < S * S; i++) { v[i] = (p[i * 4] + p[i * 4 + 1] + p[i * 4 + 2]) / 3; m += v[i]; } m /= S * S;
    let sd = 0; for (const x of v) sd += (x - m) ** 2; sd = Math.sqrt(sd / v.length) || 1;
    for (let i = 0; i < v.length; i++) v[i] = (v[i] - m) / sd;
    return v;
  }
  const VAR_ESC = [0.95, 1, 1.05], VAR_DESL = [-0.03, 0, 0.03];
  const correlacao = (a, b) => { let s = 0; for (let i = 0; i < a.length; i++) s += a[i] * b[i]; return s / a.length; };
  const cacheBib = {};
  async function assinaturasBiblioteca(tipo) {
    const chave = tipo + JSON.stringify(window.IMPORTAR_VAR || {});
    if (cacheBib[chave]) return cacheBib[chave];
    const lista = tipo === 'perigo' ? BIB.perigos : listaEpis();
    const out = [];
    for (const it of lista) { const a = await assinatura(it.img); if (a) out.push({nome: it.nome, a}); }
    return (cacheBib[chave] = out);
  }
  const MINIMO = 0.68, EMPATE = 0.02;   // semelhança mínima com o pictograma da biblioteca; diferença mínima para o 2º colocado
  function nomesEquivalentes(a, b) {
    const pal = s => new Set(norm(s).replace(/(JPG|PNG|JPEG)$/, '').split(/[^A-Z0-9]+/).filter(w => w && !/^(DE|DA|DO|DAS|DOS|E|A|O|COM|EM|RISCO|NA|NO)$/.test(w)).map(w => w.replace(/S$/, '')));
    let [x, y] = [pal(a), pal(b)];
    return x.size > 0 && x.size === y.size && [...x].every(w => y.has(w));
  }
  // nome antigo (escrito na imagem ou embaixo dela) -> pictograma atual da biblioteca, pelas tabelas de nomes antigos
  let mapaAntigos = null;
  function nomePeloAntigo(rotulo, tipo) {
    if (!rotulo) return '';
    const k = t => norm(t).replace(/[^A-Z0-9]+/g, ' ').trim();
    if (!mapaAntigos) {
      mapaAntigos = {perigo: new Map(), epi: new Map()};
      for (const [a, b] of Object.entries(window.PERIGOS_NOMES_ANTIGOS || {})) mapaAntigos.perigo.set(k(a), b);
      for (const [a, b] of Object.entries(window.EPIS_NOMES_ANTIGOS || {})) mapaAntigos.epi.set(k(a), b);
    }
    const n = mapaAntigos[tipo] && mapaAntigos[tipo].get(k(rotulo));
    return n && (tipo === 'perigo' ? BIB.perigos : listaEpis()).some(it => it.nome === n) ? n : '';
  }
  async function reconhecer(src, tipo) {
    const vs = await assinatura(src, true);
    // 1º: imagens que você já trocou antes pelo pictograma certo, e as dos POs de treino
    // (pode voltar nome vazio: imagem conhecida que não existe na biblioteca, mas com o nome escrito nela)
    if (typeof APRENDIDOS !== 'undefined') {
      const ap = await APRENDIDOS.buscar(tipo, src, vs, () => assinaturaInteira(src));
      if (ap && ap.nome) return ap;
      // conhecida do treino, mas a biblioteca não tinha o pictograma naquela época: se agora tiver um com o nome escrito
      // na imagem ("Protetor solar", "Chuvas e raios"), usa esse
      if (ap) {
        const k = norm(ap.rotulo).replace(/[^A-Z0-9]+/g, ' ').trim();
        const iguais = k ? (tipo === 'perigo' ? BIB.perigos : listaEpis()).filter(it => norm(rotuloPadrao(it.nome)).replace(/[^A-Z0-9]+/g, ' ').trim() === k) : [];
        if (iguais.length === 1) return {nome: iguais[0].nome, rotulo: ap.rotulo, d: ap.d};
        // nome escrito na imagem antiga ("Queda de materiais", "Tombamento", "Chuvas e raios") que já tem pictograma novo
        // na tabela de nomes antigos (perigos_nomes_antigos.js): usa esse (07/11/2026, teste com os POs antigos)
        const antigo = nomePeloAntigo(ap.rotulo, tipo);
        if (antigo) return {nome: antigo, rotulo: ap.rotulo, d: ap.d};
        const lib = await reconhecerNaBiblioteca(vs, tipo, src);
        return lib && ap.rotulo && nomeBate(ap.rotulo, lib.nome) ? {nome: lib.nome, rotulo: ap.rotulo, d: lib.d} : ap;
      }
    }
    return reconhecerNaBiblioteca(vs, tipo, src);
  }
  async function reconhecerNaBiblioteca(vs, tipo, src) {
    if (!vs || !vs.length) return null;
    const ord = (await assinaturasBiblioteca(tipo)).map(b => ({b, c: Math.max(...vs.map(v => correlacao(v, b.a)))})).sort((x, y) => y.c - x.c);
    const [p1, p2] = ord;
    if (!p1) return null;
    // dois pictogramas da biblioteca quase iguais (ex.: luvas, uniformes) = não dá para saber qual é
    // ...a não ser que sejam o mesmo pictograma com nomes quase iguais ("material inflamavel" / "Material Inflamável")
    const empate = !!p2 && p1.c - p2.c < EMPATE && !nomesEquivalentes(p1.b.nome, p2.b.nome);
    if (window.IMPORTAR_DEBUG) window.IMPORTAR_DEBUG.push({tipo, nome: p1.b.nome, d: +p1.c.toFixed(3), seg: p2 ? p2.b.nome : '', d2: p2 ? +p2.c.toFixed(3) : 0, empate, top: ord.slice(0, 6).map(o => [o.b.nome, +o.c.toFixed(3)]), src: window.IMPORTAR_DEBUG_SRC ? src : undefined});
    return p1.c >= MINIMO && !empate ? {nome: p1.b.nome, d: p1.c} : null;
  }
  // palavras do nome (sem "de", "risco"...): "Queda de nível diferente" ~ "Risco de queda de diferente nível"
  function palavrasNome(s) {
    return new Set(norm(s).replace(/(JPG|PNG|JPEG|WEBP)$/, '').split(/[^A-Z0-9]+/).filter(w => w.length > 2 && !/^(DE|DA|DO|DAS|DOS|COM|EM|RISCO|NA|NO|PARA|POR|OU)$/.test(w)).map(w => w.replace(/S$/, '')));
  }
  function nomeBate(rotulo, nome) {
    const a = palavrasNome(rotulo), b = palavrasNome(nome);
    if (!a.size || !b.size) return false;
    const comum = [...a].filter(w => b.has(w)).length;
    return comum / Math.max(a.size, b.size) >= 0.6;
  }
  async function reconhecerComTexto(src, tipo, rotulo) {
    const r = await reconhecer(src, tipo);
    if (!rotulo) return r;
    const k = t => norm(t).replace(/[^A-Z0-9]+/g, ' ').trim(), kr = k(rotulo);
    const igualRotulo = nome => k(rotuloPadrao(nome)) === kr || k(nome) === kr;
    if (r && r.nome && igualRotulo(r.nome)) return r;
    const antigo = nomePeloAntigo(rotulo, tipo);
    if (antigo && (!r || !r.nome)) return {nome: antigo, d: r ? r.d : 0};
    const vs = await assinatura(src, true);
    if (!vs || !vs.length) return r;
    const ord = (await assinaturasBiblioteca(tipo)).map(b => ({b, c: Math.max(...vs.map(v => correlacao(v, b.a)))})).sort((x, y) => y.c - x.c);
    if (!ord.length) return r;
    // o nome escrito embaixo decide entre pictogramas parecidos ou iguais na biblioteca ("1 - Atropelamento" e
    // "2 - Atropelamento", luvas, uniformes): primeiro o nome exatamente igual, depois o que tem mais palavras em comum
    // (sem empate). Já reconhecido por outro nome: só troca se o do nome escrito for quase igual na imagem (07/11/2026)
    const perto = ord.filter(o => o.c >= 0.4 && o.c >= ord[0].c - (r ? 0.03 : 0.15)).slice(0, 12);
    const exato = perto.filter(o => igualRotulo(o.b.nome));
    if (exato.length) return {nome: exato[0].b.nome, rotulo: r && r.rotulo, d: exato[0].c};
    if (r) return r;
    const nota = o => { const a = palavrasNome(rotulo), b = palavrasNome(rotuloPadrao(o.b.nome)); if (!a.size || !b.size) return 0; const c = [...a].filter(w => b.has(w)).length; return c / Math.max(a.size, b.size); };
    const ns = perto.map(o => ({o, n: nota(o)})).filter(x => x.n >= 0.6).sort((x, y) => y.n - x.n || y.o.c - x.o.c);
    const ach = ns.length && (ns.length === 1 || ns[0].n > ns[1].n + 1e-9 || ns[0].o.c > ns[1].o.c + EMPATE) ? ns[0].o : null;
    return ach ? {nome: ach.b.nome, d: ach.c} : null;
  }
  // nome do pictograma -> nome exato da linha da planilha de perigos (vazio se o pictograma não é de um perigo da planilha)
  const nomePlanilha = nome => { const r = rotuloPadrao(nome); return (window.TABELA_PERIGOS || []).some(p => p.perigo === r) ? r : ''; };
  async function pictogramas(srcs, tipo, rotulosTexto) {
    const itens = [];
    for (const [i, src] of srcs.entries()) {
      const r = await reconhecerComTexto(src, tipo, rotulosTexto[i]);
      if (r && r.nome) itens.push({nome: r.nome, rotulo: r.rotulo || rotuloPadrao(r.nome), conferir: true});
      else itens.push({nome: '', rotulo: (r && r.rotulo) || '', img: src, conferir: true, novo: true});
    }
    // nomes escritos na célula (quando houver) completam os rótulos
    rotulosTexto.forEach((t, i) => { if (itens[i] && t) itens[i].rotulo = t; });
    // perigo reconhecido: o nome que aparece é o da planilha de perigos, não o texto antigo do Word (05/11/2026, regra do usuário)
    if (tipo === 'perigo') for (const it of itens) if (it.nome) { const p = nomePlanilha(it.nome); if (p) it.rotulo = p; }
    return itens;
  }

  // ---------------- leitura ----------------
  const singular = s => norm(s).replace(/AIS$/, 'AL').replace(/OES$/, 'AO').replace(/S$/, '');
  const ROTULOS_SUB = TIPOS_SUB.map(singular);
  function papel(texto) {
    const t = norm(texto);
    if (!t || t.length > 40) return null;
    if (/BASICO/.test(t)) return null;
    if (/^N[º°O.]?$|^N[º°]? ?ITEM$|^ITEM$/.test(t)) return 'num';
    if (/PONTOS/.test(t)) return 'pontos';
    if (/ETAPA/.test(t)) return 'nome';
    if (/FOTO/.test(t)) return 'fotos';
    if (/PERIGO|RISCO/.test(t)) return 'perigos';
    if (/EPI/.test(t)) return 'epis';
    if (/DESCRI|SEGURANCA|OPERACIONAL|PASSO|INSTRU/.test(t)) return 'desc';
    return null;
  }
  // linha de títulos das colunas: precisa ter pelo menos Etapa, Foto e Perigos/Descrição
  function ehLinhaTitulos(cels) {
    const ps = new Set(cels.map(c => papel(textoCel(c.tc))).filter(Boolean));
    return ps.has('nome') && (ps.has('fotos') || ps.has('desc')) && (ps.has('perigos') || ps.has('epis') || ps.has('desc'));
  }
  // coluna sem título (ou PO sem linha de títulos): o papel sai do conteúdo das linhas de etapa.
  // foto = imagens grandes; descrição = texto longo; perigos = a coluna com mais imagens pequenas; EPIs = a outra
  function perfilColunas(linhas, ctx) {
    const cols = new Map();
    for (const linha of linhas) for (const x of linha) {
      if (x.cont) continue;
      const p = cols.get(x.c0) || {c0: x.c0, c1: x.c1, txt: 0, img: 0, grande: 0, n: 0};
      p.n++; p.txt += textoCel(x.tc).length;
      for (const r of refsImagem(x.tc, ctx)) { p.img++; if (Math.max(r.tam[0], r.tam[1]) >= 3) p.grande++; }
      cols.set(x.c0, p);
    }
    return [...cols.values()].sort((a, b) => a.c0 - b.c0);
  }
  function completarMapa(mapa, titulos, seguintes, ctx) {
    const livres = titulos.filter(x => !x.cont && !textoCel(x.tc));
    if (!livres.length || (mapa.fotos && mapa.desc)) return [];
    const perf = perfilColunas(seguintes, ctx), novos = [];
    for (const x of livres) {
      const p = perf.find(q => sobrepoe(q, x.c0, x.c1));
      if (!p) continue;
      if (!mapa.fotos && p.grande && p.grande >= p.n / 2) { mapa.fotos = [x.c0, x.c1]; novos.push('fotos'); }
      else if (!mapa.desc && p.txt / p.n > 40) { mapa.desc = [x.c0, x.c1]; novos.push('descrição'); }
    }
    return novos;
  }
  const ehLinhaEtapaSolta = (linha, ctx) => linha.length >= 4 && /^\d{1,2}\.?$/.test(textoCel(linha[0].tc)) && textoCel(linha[1].tc).length > 2 && linha.some(x => refsImagem(x.tc, ctx).length);
  function deduzirMapa(seguintes, ctx) {
    const perf = perfilColunas(seguintes, ctx);
    if (perf.length < 4) return null;
    const mapa = {num: [perf[0].c0, perf[0].c1], nome: [perf[1].c0, perf[1].c1]};
    const resto = perf.slice(2);
    const desc = resto.slice().sort((a, b) => b.txt / b.n - a.txt / a.n)[0];
    if (desc && desc.txt / desc.n > 40) mapa.desc = [desc.c0, desc.c1];
    const comImg = resto.filter(p => p !== desc && p.img);
    const foto = comImg.filter(p => p.grande >= p.n / 2).sort((a, b) => b.grande - a.grande)[0];
    if (foto) mapa.fotos = [foto.c0, foto.c1];
    const peq = comImg.filter(p => p !== foto).sort((a, b) => b.img - a.img);
    if (peq[0]) mapa.perigos = [peq[0].c0, peq[0].c1];
    const epi = peq[1] || resto.find(p => p !== desc && p !== foto && p !== peq[0]);
    if (epi) mapa.epis = [epi.c0, epi.c1];
    return mapa.desc || mapa.fotos ? mapa : null;
  }
  // EPI escrito só com o nome (sem imagem): pictograma da biblioteca com o mesmo nome, se houver um só
  const PAL_EPI = s => palavrasNome(String(s).replace(/\bPFF\s*-?\s*(II|ll|2)\b/gi, 'PFF2'));
  // nomes que os POs e APRs antigos usam para EPIs da biblioteca (teste com os documentos antigos, 07/11/2026)
  const EPI_APELIDOS = [
    [/^CREME\b/, 'Creme Protetivo'],
    [/^RESPIRADOR (DESCARTAVEL )?PFF ?(2|II|LL)$/, 'Respirador Descartável PFF2'],
    [/REFLETIV/, 'Colete Refletivo'],
    [/^CINTO DE SEGURANCA$/, 'Cinto de segurança obrigatório'],
    [/^CINTO (DE SEGURANCA )?(TIPO )?PARA ?QUEDISTA$/, 'Cinto de segurança Paraquedista'],
    [/^OCULOS (DE )?(PROTECAO|SEGURANCA)$/, 'Óculos de Segurança'],
    [/^LUVA (DE MALHA )?(COM PALMA )?(EM )?PU$/, 'Luva de Malha com Palma em PU'],
    // pedido do usuário (07/11/2026): mesmo pictograma da luva de vaqueta / do protetor facial transparente
    [/^LUVA (DE )?(MULTI ?TATO|NEOPRENE|BORRACHA)$/, 'Luva de Vaqueta'],
    [/^PROTE[TR]OR FACIAL$/, 'Protetor Facial Transparente'],
  ];
  function epiPorNome(texto) {
    const tn = norm(texto).replace(/[^A-Z0-9]+/g, ' ').trim();
    const ap = EPI_APELIDOS.find(([re]) => re.test(tn));
    if (ap && listaEpis().some(it => it.nome === ap[1])) return ap[1];
    const t = PAL_EPI(texto);
    if (!t.size) return '';
    let melhor = [], nota = 0;
    for (const it of listaEpis()) {
      const b = PAL_EPI(rotuloPadrao(it.nome));
      if (!b.size) continue;
      const comum = [...b].filter(w => t.has(w)).length;
      // todas as palavras do nome da biblioteca estão no texto ("Capacete" em "Capacete de segurança"), ou todas as do
      // texto estão no nome da biblioteca ("Respirador PFF2" ~ "Respirador descartável PFF2"); empate = não dá para saber
      const s = comum === b.size ? 1 + comum / 10 : comum === t.size && comum / b.size >= 0.6 ? comum / b.size : 0;   // "Uniforme" sozinho não é "Uniforme NR10"
      if (s > nota + 1e-9) { nota = s; melhor = [it.nome]; } else if (s && Math.abs(s - nota) < 1e-9) melhor.push(it.nome);
    }
    return melhor.length === 1 ? melhor[0] : '';
  }
  const nomesEpiTexto = t => t.split(/[\n,;\/]/).map(s => s.replace(/^[\s'"´`.\-–•]+|[\s'"´`.;,]+$/g, '').replace(/\s+/g, ' ')).filter(s => s.length > 1 && !/^N\.?\s*\/?\s*A\.?$/i.test(s));

  // "impressão digital" do texto do Word: muda se alguém editar o documento no Word depois de exportado
  function digitalTexto(docXml) {
    const s = norm(todos(xml(docXml), 't').map(t => t.textContent).join('')).replace(/[^A-Z0-9]/g, '');
    let h = 0x811c9dc5;
    for (let i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 0x01000193) >>> 0; }
    return s.length + '-' + h.toString(16);
  }
  // "TEXTO EM MAIÚSCULAS 54A" -> "Texto em maiúsculas 54A"
  // siglas e números romanos ficam em maiúsculas ("TOTO II", "CAF", "GLP")
  const SIGLAS = /^(I{1,3}|IV|VI{0,3}|IX|X|CAF|VMSA|GLP|EPIS?|APR|PO|NR\d*|PU|PVC|PFF\d?|AMP|TI|ADM|CNC|ETE|ETA|VO|KM|ABS|EPC)$/;
  const frase = s => s.split(' ').map((w, i) => /\d/.test(w) || SIGLAS.test(w.replace(/[^\wÀ-ú]/g, '')) && w.length > 1 ? w : i ? w.toLowerCase() : w.charAt(0) + w.slice(1).toLowerCase()).join(' ');
  // "GUSTAVO PEREIRA" -> "Gustavo Pereira"; "WESKLEY M DA SILVA" -> "Weskley M da Silva"
  const nomePessoa = s => !s || s !== s.toUpperCase() || !/[A-Z]{3}/.test(s) ? s :
    s.toLowerCase().split(' ').map(w => /^(da|de|do|das|dos|e)$/.test(w) ? w : w.charAt(0).toUpperCase() + w.slice(1)).join(' ');
  // os textos guardados no Word exportado (título, etapas, descrições) ainda estão no documento?
  async function conferePalavras(z, dados) {
    const so = s => norm(s).replace(/[^A-Z0-9]/g, '');
    const doc = xml(await z.texto('word/document.xml'));
    const texto = so(todos(doc, 't').map(t => t.textContent).join(''));
    const linhas = [dados.cab && dados.cab.nome];
    for (const e of dados.etapas || []) {
      linhas.push(e.nome);
      for (const l of (e.descricao || '').split('\n')) linhas.push(new DOMParser().parseFromString(l, 'text/html').body.textContent);
    }
    return linhas.every(l => !so(l) || texto.includes(so(l)));
  }

  // código e revisão escritos no cabeçalho (também usado pelo importador de Excel)
  function codigoRev(texto) {
    const t = String(texto || '');
    // "PO" sozinho ("PORTA 4B" no título não é código); "PO.0100" = 100
    const m = /\bPO(?![A-Za-zÀ-ú])[\s.\-–]*(?:([A-Z]{1,4})[\s.\-–]+)?(\d{1,4}[A-Z]?)\b/.exec(t);
    if (!m) return {codigo: '', revisao: ''};
    const r = /\b(?:REV|VER)(?!IS)[\s.:º°nN]*(\d+)/i.exec(t.slice(m.index));
    return {codigo: (m[1] ? m[1] + '.' : '') + m[2].padStart(3, '0').replace(/^0+(?=\d{3})/, ''), revisao: r ? String(+r[1]).padStart(2, '0') : ''};
  }

  // ---------------- anexos ----------------
  const RX_ANEXO = /^\s*ANEXO\s*(\d{1,2}|[IVX]{1,5})(?![A-Z0-9])\s*[–—\-:.)]*\s*(.*)$/i;
  // seções que encerram o último anexo (treinamento, histórico de revisões...)
  const RX_FIM_ANEXO = /^(NOME COMPLETO|HISTORICO|CONTROLE DE REVIS|REGISTRO DE TREINAMENTO|LISTA DE PRESENCA|TREINAMENTO)/;
  // corpo do documento na ordem: cada linha de tabela (de 1º nível) e cada parágrafo fora de tabela vira um bloco
  function blocosCorpo(docXml) {
    const body = todos(docXml, 'body').find(b => b.namespaceURI === W);
    const out = [];
    const visitar = no => {
      for (const n of no.childNodes) {
        if (n.nodeType !== 1) continue;
        if (n.localName === 'tbl' && n.namespaceURI === W) for (const linha of grade(n)) out.push({linha, nos: linha.filter(x => !x.cont).map(x => x.tc)});
        else if (n.localName === 'p' && n.namespaceURI === W) out.push({nos: [n]});
        else if (/^(sdt|sdtContent|customXml|ins|smartTag)$/.test(n.localName)) visitar(n);
      }
    };
    if (body) visitar(body);
    for (const b of out) b.textos = b.nos.map(n => (n.localName === 'p' ? textoPar(n) : textoCel(n)).trim());
    return out;
  }
  // título de anexo no bloco: a faixa só com "ANEXO n – título" (não uma etapa que cita "conforme Anexo 1")
  function tituloAnexo(b) {
    const todo = b.textos.filter(Boolean).join('\n');
    if (!todo || todo.length > 200) return null;
    const cheias = b.textos.filter(Boolean).length;
    if (!/^ANEXO/.test(norm(todo)) && !(b.linha && cheias <= 3)) return null;
    for (const [k, t] of b.textos.entries()) {
      const linhas = t.split('\n').map(s => s.trim());
      for (const [j, l] of linhas.entries()) {
        const m = RX_ANEXO.exec(l);
        if (!m) continue;
        let tit = m[2].trim();
        if (!tit) tit = linhas.slice(j + 1).find(Boolean) || b.textos.slice(k + 1).find(s => s && !RX_ANEXO.test(s)) || '';
        tit = tit.replace(/\s+/g, ' ').replace(/[\s–—\-:.]+$/, '').trim();
        return {num: m[1], titulo: tit};
      }
    }
    return null;
  }
  // quebra de página no bloco: antes dele (pageBreakBefore, ou a quebra vem antes de qualquer conteúdo) ou depois
  function quebraDe(b) {
    let conteudo = false, antes = false, depois = false;
    for (const no of b.nos) for (const n of no.getElementsByTagNameNS('*', '*')) {
      if (n.namespaceURI === W && n.localName === 'pageBreakBefore' && atrib(n, 'val') !== '0' && atrib(n, 'val') !== 'false') { if (!conteudo) antes = true; }
      else if (n.namespaceURI === W && n.localName === 'br' && atrib(n, 'type') === 'page') { if (conteudo) depois = true; else antes = true; }
      else if (n.namespaceURI === W && n.localName === 'sectPr') depois = true;
      else if ((n.localName === 't' && n.namespaceURI === W && n.textContent.trim()) || n.localName === 'blip' || n.localName === 'imagedata') conteudo = true;
    }
    return {antes, depois};
  }
  // textos que são do cabeçalho do PO (repetido em cada página), não título de anexo
  const RX_ROTULO_CAB = /^(PROCEDIMENTO|ELABORADO|APROVADO|DOCUMENTA|REVIS|DATA\b|APR\b|EPI|PO\b|PO[\s.\-–]*[A-Z]{0,4}[\s.\-–]*\d|FOTO\s*\d|PAGINA|FOLHA)/;
  // Anexo = bloco com imagens DEPOIS da última etapa; não depende da palavra "ANEXO" nem do cabeçalho repetido.
  // Quebra de página (ou o cabeçalho do PO repetido na página) separa anexos; o título escrito acima das imagens é
  // o nome (sem título: "Anexo 1", "Anexo 2"...); páginas seguidas com o mesmo título viram um anexo só.
  // Ficam de fora: logotipo do cabeçalho, imagens muito pequenas, lista de treinamento e histórico de revisões.
  async function lerAnexos(docXml, ctx, avisos) {
    const blocos = blocosCorpo(docXml), out = [];
    // começo: logo depois da última etapa (e das linhas mescladas que ainda são dela); sem etapas, só vale "ANEXO n"
    let ini = 0;
    const soTitulados = !ctx.ultTr;
    if (ctx.ultTr) {
      ini = blocos.findIndex(b => b.linha && b.linha.length && b.linha[0].tc.parentNode === ctx.ultTr) + 1;
      while (ini > 0 && ini < blocos.length && blocos[ini].linha && blocos[ini].linha.some(x => x.cont)) ini++;
    }
    const logos = new Set();
    for (const tc of ctx.cabTcs || []) for (const src of await imagensDe(tc, ctx)) logos.add(src);
    const nomeCab = norm(ctx.cabNome);
    const ehCabecalho = (b, srcs) => srcs.some(s => logos.has(s)) || b.textos.some(t => /PROCEDIMENTO OPERACIONAL/.test(norm(t)) || (nomeCab.length > 5 && norm(t) === nomeCab));
    // páginas (trechos entre quebras) com o título e as imagens de cada uma
    const trechos = [];
    let atual = null, pular = false;
    const novo = () => { atual = {titulo: null, explicito: null, imgs: [], emf: ctx.emf}; trechos.push(atual); pular = false; };
    novo();
    for (let i = ini; i < blocos.length; i++) {
      const b = blocos[i];
      const refs = b.nos.flatMap(n => refsImagem(n, ctx));
      const srcs = [];
      for (const r of refs) {
        // imagem minúscula (marcador, linha, ícone) não é anexo
        if (r.tam[0] && r.tam[1] && Math.max(r.tam[0], r.tam[1]) < 1.5) continue;
        const s = await urlImagem(r, ctx);
        if (s) srcs.push(s);
      }
      const q = quebraDe(b);
      if (ehCabecalho(b, srcs)) {     // cabeçalho repetido = página nova; o logotipo e os rótulos dele não contam
        if (atual.imgs.length || atual.titulo) novo();
        continue;
      }
      if (q.antes && (atual.imgs.length || atual.titulo)) novo();
      const ex = tituloAnexo(b);
      if (ex && (atual.imgs.length || atual.titulo)) novo();
      const primeiro = norm(b.textos.find(Boolean) || '');
      if (!ex && RX_FIM_ANEXO.test(primeiro)) pular = true;     // treinamento, histórico de revisões...
      if (b.linha && ehLinhaTitulos(b.linha)) pular = true;
      if (!pular) {
        if (ex) { atual.titulo = ex.titulo; atual.explicito = ex.num; }
        // título: o primeiro texto curto acima das imagens (ou na mesma linha) que não é do cabeçalho
        else if (atual.titulo === null && !atual.imgs.length) {
          const t = b.textos.map(s => s.replace(/\s+/g, ' ').trim()).find(s => s.length >= 3 && s.length <= 150 && !RX_ROTULO_CAB.test(norm(s)) && norm(s) !== nomeCab && /[A-Za-zÀ-ú]{3}/.test(s));
          if (t) atual.titulo = t.replace(/[\s–—\-:.]+$/, '');
        }
        for (const s of srcs) if (!atual.imgs.includes(s) && !logos.has(s)) atual.imgs.push(s);
      }
      if (q.depois) novo();
    }
    // junta as páginas: título sozinho numa página passa para as imagens da seguinte; mesmo título seguido = um anexo
    const anexos = [];
    for (const [k, t] of trechos.entries()) {
      if (soTitulados && t.explicito === null) continue;
      if (!t.imgs.length) {
        const prox = trechos[k + 1];
        if ((t.titulo || t.explicito !== null) && prox && prox.imgs.length && prox.titulo === null && (!soTitulados || t.explicito !== null)) { prox.titulo = t.titulo; prox.explicito = t.explicito; continue; }
        // sem imagem: só um "ANEXO n – título" explícito vira anexo (com aviso), e não se for o título repetido do anterior
        if (t.explicito === null) continue;
        if (!t.titulo) { avisos.push(`O Word tem "ANEXO ${t.explicito}" sem título e sem imagem; não foi criado anexo.`); continue; }
      }
      const ult = anexos[anexos.length - 1];
      if (ult && t.titulo && ult.titulo && norm(ult.titulo) === norm(t.titulo)) { for (const s of t.imgs) if (!ult.imgs.includes(s)) ult.imgs.push(s); continue; }
      anexos.push({titulo: t.titulo, num: t.explicito, imgs: t.imgs, emf: t.emf});
    }
    for (const [ia, an] of anexos.entries()) {
      const imgs = an.imgs, antes = an.emf;
      const tit = an.titulo ? (an.titulo === an.titulo.toUpperCase() ? frase(an.titulo) : an.titulo) : 'Anexo ' + (an.num || ia + 1);
      // mais imagens do que cabem numa página de anexo: continua em outro anexo, sem perder nenhuma
      for (let k = 0; k === 0 || k < imgs.length; k += MAX_IMG_ANEXO) {
        const a = {titulo: k ? tit + ' (continuação)' : tit, imagens: []};
        for (const src of imgs.slice(k, k + MAX_IMG_ANEXO)) a.imagens.push({img: src, linha: a.imagens.length < MAX_POR_LINHA ? 1 : 2, aspecto: await aspectoDe(src)});
        out.push(a);
      }
      if (!imgs.length) avisos.push(`Anexo "${tit}": ${ctx.emf > antes ? 'a imagem está num formato antigo (WMF/EMF) que não consegui converter' : 'não encontrei imagem'}; insira a imagem no anexo.`);
      else if (imgs.length > MAX_IMG_ANEXO) avisos.push(`Anexo "${tit}": tinha ${imgs.length} imagens; como cada anexo tem até ${MAX_IMG_ANEXO}, as demais foram para "${tit} (continuação)".`);
    }
    return out;
  }

  // Tabela "HISTÓRICO DE REVISÕES" (ou "CONTROLE DE REVISÃO"): linha de títulos com Descrição/Motivo e Data;
  // as colunas são achadas pelos títulos. Linhas sem descrição e sem data (as vazias do modelo) ficam de fora.
  async function lerHistorico(tabelas, ctx) {
    const COLS = [['item', /^(ITEM|REV|N[O°º]?\b)/], ['descricao', /^(DESCRI|MOTIVO|ALTERACAO\b|ALTERACOES)/], ['foto', /^FOTO/],
      ['data', /^DATA/], ['solicitante', /SOLICITANTE/], ['responsavel', /RESPONSAVEL|REVISADO|ELABORADO/], ['obs', /^OBS/]];
    const dataIso = s => { const m = /(\d{1,2})\s*\/\s*(\d{1,2})\s*\/\s*(\d{2,4})/.exec(s || ''); if (!m) return ''; const a = m[3].length === 2 ? '20' + m[3] : m[3]; return `${a}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`; };
    const vazio = s => !s || /^[-–—\s.]*$/.test(s);
    for (const tbl of tabelas) {
      const linhas = grade(tbl);
      let marcou = false, mapa = null;
      const out = [];
      for (const linha of linhas) {
        const textos = linha.map(x => norm(textoCel(x.tc).replace(/\s+/g, ' ').trim()));
        if (!mapa) {
          if (textos.some(t => /HISTORICO DE REVIS|CONTROLE DE REVIS/.test(t))) { marcou = true; continue; }
          const m = {};
          for (const [k, x] of linha.entries()) for (const [campo, rx] of COLS) if (m[campo] === undefined && textos[k].length <= 40 && rx.test(textos[k])) { m[campo] = x; break; }
          if ((marcou || m.solicitante || m.responsavel) && m.descricao && m.data) mapa = m;
          continue;
        }
        const celula = campo => mapa[campo] && linha.find(x => sobrepoe(x, mapa[campo].c0, mapa[campo].c1));
        const txt = campo => { const x = celula(campo); return x && !x.cont ? textoCel(x.tc).replace(/\s+/g, ' ').trim() : ''; };
        const desc = txt('descricao'), data = txt('data');
        if (vazio(desc) && vazio(data)) continue;
        if (/^(HISTORICO|NOME COMPLETO|REGISTRO DE TREINAMENTO)/.test(norm(textos[0] || ''))) break;
        const xf = celula('foto');
        const foto = xf ? ((await imagensDe(xf.tc, ctx))[0] || '') : '';
        out.push({descricao: vazio(desc) ? '' : desc, foto, data: dataIso(data), solicitante: nomePessoa(vazio(txt('solicitante')) ? '' : txt('solicitante')),
          responsavel: nomePessoa(vazio(txt('responsavel')) ? '' : txt('responsavel')), obs: vazio(txt('obs')) ? '' : txt('obs')});
      }
      if (out.length) return out;
    }
    return [];
  }

  async function importar(arquivo) {
    const fase = s => { if (window.IMPORTAR_DEBUG) window.IMPORTAR_FASE = s; };
    const buf = await arquivo.arrayBuffer();
    fase('zip'); const z = await lerZip(buf); fase('zip ok');
    let editadoNoWord = false;
    // 1) Word exportado pelo gerador
    for (const n of z.nomes.filter(n => /^customXml\/item\d+\.xml$/i.test(n))) {
      const t = await z.texto(n);
      if (t && t.includes('urn:gerador-po')) {
        const m = /<!\[CDATA\[([\s\S]*)\]\]>/.exec(t.replace(/\]\]\]\]><!\[CDATA\[>/g, ']]>'));
        if (!m) continue;
        const dados = JSON.parse(m[1]);
        // se o Word foi editado depois de exportado, os textos não batem mais: lê o documento em vez dos dados guardados
        const dig = /\btexto="([^"]+)"/.exec(t.slice(0, 300));
        const igual = dig ? digitalTexto(await z.texto('word/document.xml')) === dig[1] : await conferePalavras(z, dados);
        if (igual) return {doc: dados, avisos: [], completo: true};
        editadoNoWord = dados;           // quem abriu escolhe: o que foi editado no Word ou o que o gerador salvou
      }
    }
    // 2) qualquer outro Word
    fase('document.xml'); const docXml = xml(await z.texto('word/document.xml')); fase('document ok');
    const relsXml = xml((await z.texto('word/_rels/document.xml.rels')) || '<Relationships/>');
    const rels = {};
    for (const r of todos(relsXml, 'Relationship')) rels[r.getAttribute('Id')] = 'word/' + r.getAttribute('Target').replace(/^\.?\//, '').replace(/^word\//, '');
    const cache = {};
    const ctx = {
      rels, emf: 0, cabTcs: [], ultTr: null, cabNome: '',
      async dataUrl(alvo) {
        if (cache[alvo] !== undefined) return cache[alvo];
        const d = await z.ler(alvo);
        if (!d) return (cache[alvo] = null);
        const ext = alvo.split('.').pop().toLowerCase();
        // EMF (check-list colado do Excel): o navegador não abre, então o leitor próprio desenha e devolve um PNG
        // WMF (formato mais antigo): convertido pelo mesmo leitor (21/10/2026)
        if (ext === 'emf' || ext === 'wmf') {
          let png = null;
          try {
            const b = d.buffer.slice(d.byteOffset, d.byteOffset + d.byteLength);
            png = typeof EMF === 'undefined' ? null : ext === 'emf' ? EMF.paraPNG(b, 1600) : EMF.wmfParaPNG(b, 1600);
          } catch (e) { png = null; }
          if (!png) ctx.emf++;
          return (cache[alvo] = png);
        }
        const mime = {png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', webp: 'image/webp'}[ext];
        if (!mime) return (cache[alvo] = null);
        let bin = ''; for (let i = 0; i < d.length; i += 0x8000) bin += String.fromCharCode.apply(null, d.subarray(i, i + 0x8000));
        return (cache[alvo] = `data:${mime};base64,${btoa(bin)}`);
      },
    };
    const avisos = [];
    if (editadoNoWord) avisos.push('Este Word foi exportado pelo gerador e depois alterado no Word; as informações foram lidas do documento. Confira o resultado.');
    const novo = novoDoc();
    novo.etapas = []; novo.historico = []; novo.anexos = [];
    const c = novo.cab;
    c.revAuto = false; c.tempoCiclo = 'N/A';

    const tabelas = todos(docXml, 'tbl').filter(t => t.namespaceURI === W && !t.parentNode.closest?.('tbl'));
    // ---------- cabeçalho: procura os rótulos nas células antes da tabela de etapas ----------
    const celulasCab = [];
    for (const tbl of tabelas.slice(0, 2)) {     // o cabeçalho fica na primeira tabela (ou na segunda, se a primeira começar pelas etapas)
      for (const linha of grade(tbl)) {
        if (ehLinhaTitulos(linha) || ehLinhaEtapaSolta(linha, ctx)) break;     // (PO sem linha de títulos: para na 1ª etapa)
        celulasCab.push(...linha.map(x => textoCel(x.tc)));
        // imagens do cabeçalho (logotipo), menos a linha dos EPI's básicos: não contam como anexo quando o cabeçalho se repete
        if (!linha.some(x => /EPI/.test(norm(textoCel(x.tc))))) ctx.cabTcs.push(...linha.map(x => x.tc));
      }
      if (celulasCab.some(Boolean)) break;
    }
    const todoCab = celulasCab.join('\n');
    // código: "PO.001", "PO 31", "PO.TI.070", "PO.037B", "PO.MAN - 03", "PO. OP-01", "PO - 117", "PO.MAN.13"
    // (só maiúsculas: "Porta 6B" não é código); revisão depois dele: "REV:01", "REV.000", "VER:01" (erro comum)
    const cr = codigoRev(todoCab);
    const mCod = cr.codigo && cr.revisao;
    if (cr.codigo) c.codigo = cr.codigo;
    if (cr.revisao) c.revisao = cr.revisao;
    const depois = (rotulo) => {
      for (let i = 0; i < celulasCab.length; i++) {
        const linhas = celulasCab[i].split('\n').map(s => s.trim());
        const k = linhas.findIndex(s => norm(s).startsWith(norm(rotulo)));
        if (k < 0) continue;
        const resto = linhas[k].slice(linhas[k].indexOf(':') + 1).trim();
        if (resto) return resto;
        const prox = linhas.slice(k + 1).find(s => s && !/^DATA/i.test(norm(s)));
        if (prox) return prox;
        const vizinho = (celulasCab[i + 1] || '').split('\n').map(s => s.trim()).find(Boolean);
        if (vizinho && !/:$/.test(vizinho)) return vizinho;
      }
      return '';
    };
    if (!mCod) {     // modelo do gerador: célula "AMP.009" e "Revisão n°: 02"; ou só "PO.TI.070", sem revisão
      const cod = c.codigo || celulasCab.map(s => s.trim()).find(s => /^([A-Z]{1,4}\.)?\d{2,4}$/.test(s));
      const rev = (depois('Revisão n') || depois('Revisao n') || depois('Rev')).replace(/\D/g, '');
      if (cod) c.codigo = cod;
      if (rev) c.revisao = rev.padStart(2, '0');
      if (!cod || !rev) avisos.push(!cod && !rev ? 'Não encontrei o código e a revisão do PO; preencha no cabeçalho.' : !cod ? 'Não encontrei o código do PO; preencha no cabeçalho.' : 'Não encontrei a revisão do PO; preencha no cabeçalho.');
    }
    c.elaborado = nomePessoa(depois('Elaborado por'));
    c.aprovado = nomePessoa(depois('Aprovado por'));
    const tc = depois('Tempo de ciclo'); if (tc) c.tempoCiclo = tc;
    const mData = /DATA:?\s*(\d{1,2})\s*\/\s*(\d{1,2})\s*\/\s*(\d{4})/i.exec(todoCab);
    if (mData) c.data = `${mData[3]}-${mData[2].padStart(2, '0')}-${mData[1].padStart(2, '0')}`;
    // título: o maior texto do cabeçalho que não é rótulo
    const ignorar = /^(APR\b|APR N|ANEXO|PROCEDIMENTO|ELABORADO|APROVADO|DOCUMENTA|REVIS|DATA|EPI|LEGENDA|FERRAMENTA|MATERIA|EQUIPAMENTO|TEMPO DE CICLO|QUALIDADE|SEGURANCA|MEIO AMBIENTE|ENERGIA|PO[\s.])/;
    const candidatos = celulasCab.flatMap(t => t.split('\n')).map(s => s.trim()).filter(s => s.length > 5 && !ignorar.test(norm(s)) && norm(s) !== norm(c.elaborado) && norm(s) !== norm(c.aprovado));
    c.nome = (candidatos.sort((a, b) => b.length - a.length)[0] || '').replace(/\s+/g, ' ');
    ctx.cabNome = c.nome;
    if (c.nome) c.nome = frase(c.nome);
    const ref = depois('Documentação Referência') || depois('Documentacao Referencia');
    if (ref) avisos.push(`Documentação de referência do Word ("${ref}") não tem campo no modelo novo e não foi levada.`);
    if (!c.elaborado) avisos.push('"Elaborado por" não encontrado.');
    if (!c.aprovado) avisos.push('"Aprovado por" não encontrado.');
    // histórico de revisões do Word (05/11/2026): mantém as linhas como estão; sem histórico, uma linha "Importado do Word"
    novo.historico = await lerHistorico(tabelas, ctx);
    if (!novo.historico.length) novo.historico.push({descricao: 'Importado do Word: ' + arquivo.name, foto: '', data: c.data || '', solicitante: c.elaborado, responsavel: c.aprovado, obs: ''});

    // EPI's básicos: imagens da linha que tem o rótulo "EPI's básicos"
    let basicosAchados = false;
    for (const tbl of tabelas.slice(0, 1)) {
      for (const linha of grade(tbl)) {
        const k = linha.findIndex(x => /EPI.*BASICO/.test(norm(textoCel(x.tc))));
        if (k < 0) continue;
        const srcs = [];
        // só as células até o próximo rótulo (no modelo novo a legenda vem na mesma linha)
        for (const [j, x] of linha.slice(k).entries()) {
          if (j && /[A-Za-zÀ-ú]{3}/.test(textoCel(x.tc))) break;   // um "l" solto na célula das imagens não é rótulo
          srcs.push(...await imagensDe(x.tc, ctx));
        }
        // placas antigas coladas nos EPIs básicos ("PROIBIDO USO DE ADORNOS", "USO OBRIGATÓRIO DE: UNIFORME DO TRABALHO"):
        // são bem mais largas que altas e não se usam mais; ficam de fora (07/11/2026, pedido do usuário)
        let placas = 0;
        for (let i = srcs.length - 1; i >= 0; i--) {
          const im = await carregar(srcs[i]);
          if (im && im.naturalWidth > 1.8 * im.naturalHeight) { srcs.splice(i, 1); placas++; }
        }
        if (placas) avisos.push(`EPI's básicos: ${placas} placa(s) antiga(s) (ex.: "Proibido uso de adornos") ficaram de fora; não são pictogramas de EPI.`);
        if (srcs.length) {
          basicosAchados = true;
          const itens = await pictogramas(srcs, 'epi', []);
          c.episBasicos = itens.map(it => it.nome ? {nome: it.nome} : {nome: '', img: it.img});
          const nao = itens.filter(it => !it.nome).length;
          if (nao) avisos.push(`EPI's básicos: ${nao} de ${itens.length} imagem(ns) não reconhecida(s) na biblioteca; ficaram com a imagem original.`);
        }
        break;
      }
    }
    if (!basicosAchados) avisos.push("EPI's básicos não encontrados; ficou a lista padrão.");

    fase('etapas');
    // ---------- etapas ----------
    let mapa = null;            // papel -> [c0, c1] na grade
    let ultimaEtapa = null, mapaAnt = null, nTitulos = 0;
    const temTitulos = tabelas.some(t => grade(t).some(ehLinhaTitulos)), deduzidas = new Set();
    let epiTexto = 0;
    for (const tbl of tabelas) {
      const linhas = grade(tbl);
      mapa = null;   // cada tabela tem a sua grade de colunas...
      // ...a não ser que a tabela de etapas tenha sido partida (um parágrafo no meio): mesma quantidade de colunas
      // e a primeira linha já é uma etapa com número
      const l0 = linhas[0];
      if (mapaAnt && l0 && !ehLinhaTitulos(l0) && l0.length === nTitulos && mapaAnt.num) {
        const cn = l0.find(x => sobrepoe(x, mapaAnt.num[0], mapaAnt.num[1]));
        if (cn && /^\d{1,2}\b/.test(textoCel(cn.tc))) mapa = mapaAnt;
      }
      for (let i = 0; i < linhas.length; i++) {
        const linha = linhas[i];
        if (ehLinhaTitulos(linha)) {
          mapa = {};
          for (const x of linha) { const p = papel(textoCel(x.tc)); if (p && !mapa[p]) mapa[p] = [x.c0, x.c1]; }
          for (const p of completarMapa(mapa, linha, linhas.slice(i + 1, i + 16), ctx)) deduzidas.add(p);
          mapaAnt = mapa; nTitulos = linha.length;
          continue;
        }
        // PO sem a linha de títulos: as colunas são deduzidas pelo conteúdo das linhas de etapa
        if (!mapa && !temTitulos && !nTitulos && ehLinhaEtapaSolta(linha, ctx)) {
          mapa = deduzirMapa(linhas.slice(i).filter(l => l.length === linha.length), ctx);
          if (mapa) {
            mapaAnt = mapa; nTitulos = linha.length;
            avisos.push('O Word não tem a linha de títulos das colunas (Nº, Etapa, Foto, Perigos...); as colunas foram deduzidas pelo conteúdo. Confira se fotos, descrição, perigos e EPIs ficaram no lugar certo.');
          }
        }
        if (!mapa) continue;
        const celDe = p => mapa[p] ? linha.filter(x => sobrepoe(x, mapa[p][0], mapa[p][1])) : [];
        const txtDe = p => celDe(p).filter(x => !x.cont).map(x => textoCel(x.tc)).join('\n').trim();
        const todaLinha = norm(linha.map(x => textoCel(x.tc)).join(' '));
        if (/^ANEXO\b/.test(todaLinha) || /ANEXO (\d|[IVX]+\b)/.test(todaLinha) && linha.length <= 3) { mapa = null; continue; }
        // sublinha (Equipamento, Ferramenta...) do modelo novo: rótulo na coluna da descrição
        const celsDesc = celDe('desc').filter(x => !x.cont);
        if (ultimaEtapa && celsDesc.length >= 2 && ROTULOS_SUB.includes(singular(textoCel(celsDesc[0].tc)))) {
          const tipo = TIPOS_SUB[ROTULOS_SUB.indexOf(singular(textoCel(celsDesc[0].tc)))];
          const valor = celsDesc.slice(1).map(x => textoCel(x.tc)).join(' ').trim();
          if (valor) ultimaEtapa.sublinhas.push({tipo, valor});
          ctx.ultTr = linha[0].tc.parentNode;
          continue;
        }
        const num = txtDe('num').replace(/\D+/g, ' ').trim().split(' ')[0];
        const nome = txtDe('nome').replace(/\s+/g, ' ').trim();
        const descPars = celsDesc.flatMap(x => parsDe(x.tc));
        if (!num && !nome && !descPars.some(p => textoPar(p).trim())) continue;
        ctx.ultTr = linha[0].tc.parentNode;     // os anexos começam depois da última linha de etapa
        // continuação da mesma etapa (modelo novo repete Nº e nome quando a etapa passa de página)
        let e;
        if (ultimaEtapa && num && ultimaEtapa._num === num && norm(ultimaEtapa.nome) === norm(nome)) e = ultimaEtapa;
        else {
          e = novaEtapa(); e._num = num; e.nome = nome ? (nome === nome.toUpperCase() ? frase(nome) : nome.charAt(0).toUpperCase() + nome.slice(1)) : '';
          novo.etapas.push(e); ultimaEtapa = e;
        }
        // descrição: cada parágrafo vira um item (sem marcadores de lista digitados)
        // um item por parágrafo ou por quebra de linha; "1 - ... 2 - ... 3 - ..." no mesmo parágrafo vira um item cada
        const linhasDesc = descPars.flatMap(p => htmlPar(p).split('\n'))
          .flatMap(h => /^\s*1\s*[-–]\s/.test(textoPuro(h)) && /\s2\s*[-–]\s/.test(h) ? h.split(/\s+(?=\d{1,2}\s*[-–]\s)/) : [h])
          .map(h => h.trim().replace(/^(?:\s|&nbsp;)*(?:[•▪●◦\-–]|\d{1,2}\s*[.)]|\d{1,2}\s*[-–](?=\s))\s*/, '')).filter(h => textoPuro(h));
        if (linhasDesc.length) e.descricao = [e.descricao, ...linhasDesc].filter(Boolean).join('\n');
        // fotos e legendas: mantém as fotos lado a lado (até 2 por linha) e o tamanho do Word
        for (const x of celDe('fotos').filter(x => !x.cont)) {
          for (const g of await fotosDaCelula(x.tc, ctx)) {
            const legenda = (g.legenda || '').replace(/^FOTO\s*\d+\s*[–-]?\s*/i, '').trim();
            // mais de 2 no mesmo parágrafo: vão em pares
            for (let k = 0; k < g.fotos.length; k += 2) {
              const par = g.fotos.slice(k, k + 2), ini = e.fotos.length;
              for (const [j, f] of par.entries()) {
                const aspecto = await aspectoDe(f.src);
                const foto = {img: f.src, legenda: k === 0 && j === 0 ? legenda : '', aspecto};
                // formas desenhadas no Word sobre a foto: viram formas do DESENHAR (gravadas na foto e editáveis)
                if (f.formas && f.formas.length && f.caixa && aspecto && typeof formasNaURL === 'function') {
                  foto.original = f.src;
                  foto.formas = formasNaFoto(f.formas, f.caixa.w, aspecto);
                  foto.img = await formasNaURL(f.src, foto.formas);
                  ctx.nFormas = (ctx.nFormas || 0) + foto.formas.length;
                  if (f.conferir) { foto.conferirFormas = true; (ctx.etapasConferir = ctx.etapasConferir || new Set()).add(novo.etapas.indexOf(e) + 1); }
                }
                // o tamanho do Word só vale se a imagem não foi esticada (senão fica no automático)
                if (f.w && f.h && aspecto && Math.abs(f.w / f.h / aspecto - 1) < 0.04) {
                  const lim = par.length === 2 ? FOTO_LADO : FOTO_MAX, esc = Math.min(1, lim.larg / 10 / f.w, lim.alt / 10 / f.h);
                  foto.largura = +(f.w * esc).toFixed(2); foto.altura = +(f.h * esc).toFixed(2);
                }
                e.fotos.push(foto);
              }
              if (par.length === 2) { e.juntar = e.juntar || []; e.juntar[ini] = true; }
            }
          }
        }
        // pontos de atenção (só no modelo novo; no antigo fica Segurança)
        const tp = norm(txtDe('pontos'));
        if (tp) e.pontos = PONTOS.filter(p => tp.includes(norm(p.nome))).map(p => p.id);
        // perigos e EPIs (imagem reconhecida pela biblioteca; nome escrito na célula quando houver)
        for (const [papelCol, lista, tipo] of [['perigos', 'perigos', 'perigo'], ['epis', 'epis', 'epi']]) {
          const cels = celDe(papelCol).filter(x => !x.cont);
          const srcs = []; const nomes = [];
          for (const x of cels) {
            srcs.push(...await imagensDe(x.tc, ctx));
            nomes.push(...textoCel(x.tc).split('\n').map(s => s.trim()).filter(s => s && !/^N\/?A$/i.test(s)));
          }
          // EPI escrito só com o nome ("Luva multitato, creme resistente a óleo e graxa"): entra com o nome escrito e,
          // se a biblioteca tiver um pictograma com esse nome, com ele; senão fica sem pictograma, para escolher no editor
          if (!srcs.length) {
            if (tipo !== 'epi') continue;
            const itens = nomesEpiTexto(nomes.join('\n')).map(t => ({nome: epiPorNome(t), rotulo: t.charAt(0).toUpperCase() + t.slice(1)}))
              .filter(it => !e.epis.some(j => norm(j.rotulo) === norm(it.rotulo)));
            e.epis.push(...itens);
            epiTexto += itens.length;
            const nao = itens.filter(it => !it.nome).length;
            if (nao) avisos.push(`Etapa ${novo.etapas.indexOf(e) + 1} (${e.nome || 'sem nome'}): ${nao} de ${itens.length} EPI(s) escrito(s) só com o nome não reconhecido(s) na biblioteca; ficou só o nome, sem pictograma.`);
            continue;
          }
          const rotulos = srcs.length === nomes.length ? nomes : [];
          const itens = (await pictogramas(srcs, tipo, rotulos))
            .filter(it => !e[lista].some(j => (it.nome ? j.nome === it.nome : j.img === it.img) && norm(j.rotulo) === norm(it.rotulo)));
          e[lista].push(...itens.map(({conferir, novo: _n, ...it}) => it));
          // EPIs escritos na célula que não batem com nenhuma imagem ("Capacete de segurança, luva multitato" ao lado das
          // imagens de botina e óculos): entram também, só com o nome, para não se perderem
          if (tipo === 'epi' && nomes.length && !rotulos.length) {
            const extra = nomesEpiTexto(nomes.join('\n')).filter(t => !e.epis.some(j => norm(j.rotulo) === norm(t) || (j.nome && (j.nome === epiPorNome(t) || nomeBate(t, j.nome)))))
              .map(t => ({nome: epiPorNome(t), rotulo: t.charAt(0).toUpperCase() + t.slice(1)}));
            e.epis.push(...extra); epiTexto += extra.length;
          }
          const nao = itens.filter(it => !it.nome && it.novo).length;
          if (nao) avisos.push(`Etapa ${novo.etapas.indexOf(e) + 1} (${e.nome || 'sem nome'}): ${nao} de ${itens.length} ${tipo === 'perigo' ? 'perigo(s)' : 'EPI(s)'} não reconhecido(s) na biblioteca; ficou a imagem original, sem nome.`);
        }
        // imagem solta dentro da descrição (pictograma flutuante ancorado na célula errada): vai para os perigos se for
        // um pictograma reconhecido; senão fica de fora com aviso
        for (const x of celsDesc) for (const src of await imagensDe(x.tc, ctx)) {
          const r = await reconhecer(src, 'perigo');
          if (r && r.nome && !e.perigos.some(j => j.nome === r.nome)) e.perigos.push({nome: r.nome, rotulo: nomePlanilha(r.nome) || r.rotulo || rotuloPadrao(r.nome)});
          else if (!r || !r.nome) avisos.push(`Etapa ${novo.etapas.indexOf(e) + 1} (${e.nome || 'sem nome'}): uma imagem dentro da descrição não foi levada; se for importante, insira na etapa.`);
        }
      }
    }
    if (deduzidas.size) avisos.push(`A coluna de ${[...deduzidas].join(' e da ')} não tinha título no Word; foi reconhecida pelo conteúdo. Confira.`);
    if (epiTexto) avisos.push(`${epiTexto} EPI(s) estavam escritos só com o nome (sem imagem) no Word: ficaram com o nome escrito e, quando a biblioteca tem um com o mesmo nome, com esse pictograma. Confira.`);
    if (ctx.nFormas) avisos.push(`${ctx.nFormas} forma(s) desenhada(s) no Word sobre as fotos (círculos, setas, caixas de texto...) vieram junto e podem ser mudadas em AJUSTAR > DESENHAR.`);
    if (ctx.etapasConferir && ctx.etapasConferir.size) avisos.push(`Formas para conferir nas fotos da(s) etapa(s) ${[...ctx.etapasConferir].join(', ')}: a posição no Word não é exata; abra AJUSTAR e confira (a foto tem o selo "conferir formas").`);
    if (ctx.formasFora) avisos.push(`${ctx.formasFora} forma(s) do Word sobre as fotos não foram levadas (grupos com pictograma, formas só em parte sobre a foto ou de tipo diferente).`);
    if (ctx.marcacoes) avisos.push(`${ctx.marcacoes} marcação(ões) escrita(s) sobre as fotos (caixas de texto) não foram levadas; se forem importantes, ponha na legenda.`);
    fase('anexos');
    // anexos: título "ANEXO 1 – ...", "Anexo I:", "ANEXO 2)"... e todas as imagens que vêm depois dele
    // (mesma linha, linhas e tabelas seguintes, parágrafos fora de tabela) até o próximo anexo ou o fim do documento
    for (const a of await lerAnexos(docXml, ctx, avisos)) novo.anexos.push(a);
    if (!novo.etapas.length) {
      avisos.unshift('Não encontrei a tabela de etapas (títulos como Nº, Etapa, Foto, Descrição, Perigos). O documento foi aberto só com o cabeçalho.');
      novo.etapas.push(novaEtapa());
    }
    novo.etapas.forEach(e => delete e._num);
    const nPer = novo.etapas.reduce((s, e) => s + e.perigos.length, 0), nEpi = novo.etapas.reduce((s, e) => s + e.epis.length, 0);
    if (nPer + nEpi) avisos.push('Confira os nomes embaixo dos perigos e EPIs: no Word eles fazem parte da imagem, então o gerador usou o nome do pictograma da biblioteca.');
    const nDup = novo.etapas.map(e => e.nome).filter((n, i, a) => n && a.indexOf(n) !== i).length;
    if (nDup) avisos.push(`${nDup} etapa(s) com nome repetido; confira se não é a mesma etapa.`);
    if (ctx.emf) avisos.push(`${ctx.emf} imagem(ns) em formato antigo (WMF/EMF) que não consegui converter ficaram de fora.`);
    return {doc: novo, avisos, completo: false, guardado: editadoNoWord || null};
  }
  // peças usadas também pela importação do Excel (importar_excel.js)
  const ajuda = {lerZip, xml, todos, norm, esc, papel, pictogramas, frase, assinatura, assinaturaInteira, correlacao, codigoRev, nomePessoa};
  return {importar, digitalTexto, ajuda};
})();

// janela com o que não foi reconhecido
function mostrarAvisosImportacao(nome, avisos, n) {
  const m = $('#modal');
  const fechar = () => { m.hidden = true; m.replaceChildren(); };
  // pictogramas não reconhecidos viram uma linha só (abre a lista por etapa); o resto é dividido entre
  // o que precisa de ação e o que é só para conferir
  const RX_PICTO = /^(Etapa \d+.*|EPI's básicos): (\d+) de \d+ .*não reconhecid/;
  const CONFERIR = /^(\d+ forma\(s\) desenhada|\d+ forma\(s\) do Word|O Word tem "ANEXO|Confira os nomes|A coluna de .* não tinha título|\d+ EPI\(s\) estavam escritos|Documentação de referência|\d+ etapa\(s\) com nome repetido|A data do cabeçalho|\d+ marcação|\d+ imagem\(ns\) pequena|Anexo .*: tinha \d+ imagens)/;
  const pictos = avisos.filter(a => RX_PICTO.test(a));
  const atencao = avisos.filter(a => !RX_PICTO.test(a) && !CONFERIR.test(a));
  const conferir = avisos.filter(a => !RX_PICTO.test(a) && CONFERIR.test(a));
  const lista = itens => el('ul', {style: 'margin:0 0 10px;padding-left:20px;line-height:1.5'}, itens);
  const nPictos = pictos.reduce((s, a) => s + +RX_PICTO.exec(a)[2], 0);
  // sem o "•" da lista: a seta do "abrir lista" já marca o item (05/11/2026)
  const linhaPictos = pictos.length ? el('li', {style: 'list-style:none;margin-left:-14px'},
    el('details', {},
      el('summary', {style: 'cursor:pointer'}, `${nPictos} pictograma(s) não reconhecido(s) em ${pictos.length} ${pictos.length > 1 ? 'lugares' : 'lugar'}: ficaram com a imagem do documento e a marca "conferir" no editor (ou só com o nome, onde o documento não tinha imagem). Clique na imagem para escolher o certo; o gerador aprende a troca para as próximas importações.`),
      el('ul', {style: 'margin:4px 0 0;padding-left:18px;font-size:12.5px;color:var(--ui-suave)'}, pictos.map(a => el('li', {}, a.replace(/; ficou.*$|; ficaram.*$/, '')))))) : null;
  const titulo = t => el('p', {style: 'margin:0 0 4px;font-weight:600'}, t);
  m.replaceChildren(el('div', {class: 'janela', style: 'height:auto;max-height:80vh;width:min(720px,94vw)', onclick: e => e.stopPropagation()},
    el('header', {}, el('b', {class: 'quebra'}, (/\.xlsx$/i.test(nome) ? 'Excel' : 'Word') + ' aberto: ', el('span', {style: 'text-transform:none'}, nome)), el('span', {style: 'flex:1'}), el('button', {class: 'btn primario', onclick: fechar}, 'Entendi')),
    el('div', {style: 'padding:12px 16px;overflow:auto'},
      el('p', {style: 'margin:0 0 10px'}, `${n} etapa(s) lida(s).` + (avisos.length ? ' Confira os pontos abaixo antes de salvar:' : '')),
      atencao.length ? [titulo('Precisa de atenção'), lista(atencao.map(a => el('li', {}, a)))] : null,
      conferir.length || linhaPictos ? [titulo('Para conferir'), lista([linhaPictos, ...conferir.map(a => el('li', {}, a))])] : null,
      el('p', {class: 'dica', style: 'margin-top:10px'}, 'O documento foi aberto como novo: use "Salvar" para gravar o PO em Word.'))));
  m.onclick = fechar;
  m.hidden = false;
}
