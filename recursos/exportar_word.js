// Exportação do PO para Word (.docx) no layout do PO.XX.000 - Modelo Padrão_Rev01.
// Cada página é uma tabela com altura exata das linhas, para o Word quebrar as páginas igual ao PDF.
'use strict';

async function gerarWord() {
  const esc = XL.esc;
  const AZUL = '002060', CINZA = '808080';
  const TW = mm => Math.round(mm * 56.6929);            // mm -> twips
  const EMU = mm => Math.round(mm * 36000);             // mm -> EMU
  const G17 = [197, 300, 1281, 2539, 943, 993, 568, 931, 627, 573, 277, 22, 2397, 2529, 846, 990, 247];
  const GH = [704, 3544, 1172, 1379, 993, 2480, 568, 1913, 3507];
  const PAG = {cab: 31, col: 10, inteira: 158, meia: 79};   // mm (um pouco abaixo do PDF para o Word não empurrar linhas)

  // ---------------- imagens ----------------
  const midias = [];                         // {url, ext, id}
  async function img(src) {
    const im = await imagemXL(src);
    if (!im) return null;
    let m = midias.find(x => x.url === im.url);
    if (!m) { m = {url: im.url, ext: im.ext, id: 'rIdImg' + (midias.length + 1), w: im.w, h: im.h}; midias.push(m); }
    return m;
  }
  let docPr = 0;
  const desenho = (m, wmm, hmm, moldura) => {
    docPr++;
    return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${EMU(wmm)}" cy="${EMU(hmm)}"/><wp:docPr id="${docPr}" name="Imagem ${docPr}"/>` +
      `<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:pic><pic:nvPicPr><pic:cNvPr id="${docPr}" name="img${docPr}.${m.ext}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${m.id}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
      `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${EMU(wmm)}" cy="${EMU(hmm)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
      `${moldura ? '<a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>' : ''}</pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
  };

  // ---------------- texto ----------------
  const run = (t, o = {}) => `<w:r><w:rPr><w:rFonts w:ascii="${o.fonte || 'Arial'}" w:hAnsi="${o.fonte || 'Arial'}" w:cs="${o.fonte || 'Arial'}"/>${o.b ? '<w:b/><w:bCs/>' : ''}${o.i ? '<w:i/>' : ''}${o.u ? '<w:u w:val="single"/>' : ''}` +
    `<w:color w:val="${o.cor || '000000'}"/><w:sz w:val="${Math.round((o.sz || 8) * 2)}"/><w:szCs w:val="${Math.round((o.sz || 8) * 2)}"/>` +
    `${o.fundo ? `<w:shd w:val="clear" w:color="auto" w:fill="${o.fundo}"/>` : ''}</w:rPr><w:t xml:space="preserve">${esc(t)}</w:t></w:r>`;
  const par = (conteudo, o = {}) => `<w:p><w:pPr><w:spacing w:before="${o.antes || 0}" w:after="${o.depois || 0}" w:line="${o.linha || 240}" w:lineRule="${o.exata ? 'exact' : 'auto'}"/>` +
    `${o.esq != null ? `<w:ind w:left="${o.esq}" w:hanging="${o.pend || 0}"/>` : o.ind ? `<w:ind w:left="${o.ind}" w:hanging="${o.ind}"/>` : ''}<w:jc w:val="${o.jc || 'center'}"/><w:rPr><w:sz w:val="${Math.round((o.sz || 8) * 2)}"/></w:rPr></w:pPr>${conteudo}</w:p>`;
  const txt = (t, o = {}) => par(run(t, o), o);

  // ---------------- tabela em grade (mesclagens como no PDF) ----------------
  // celulas: {c1, c2, r1, r2 (1-based, inclusivos), xml, fundo, rot, va, sem: {l, r, t, b}}
  function tabela(grade, alturasMm, celulas, extra = '', nCab = 0) {
    const bordas = '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(b => `<w:${b} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`).join('') + '</w:tblBorders>';
    let linhas = '';
    alturasMm.forEach((hmm, k) => {
      const r = k + 1;
      const doLinha = celulas.filter(c => c.r1 <= r && c.r2 >= r).sort((a, b) => a.c1 - b.c1);
      let col = 1, tcs = '';
      const vazio = (c1, c2) => `<w:tc><w:tcPr><w:tcW w:w="${grade.slice(c1 - 1, c2).reduce((s, x) => s + x, 0)}" w:type="dxa"/>${c2 > c1 ? `<w:gridSpan w:val="${c2 - c1 + 1}"/>` : ''}</w:tcPr><w:p/></w:tc>`;
      for (const c of doLinha) {
        if (c.c1 > col) tcs += vazio(col, c.c1 - 1);
        const w = grade.slice(c.c1 - 1, c.c2).reduce((s, x) => s + x, 0);
        const sem = c.sem || {};
        const tb = Object.keys(sem).length ? `<w:tcBorders>${[['t', 'top'], ['l', 'left'], ['b', 'bottom'], ['r', 'right']].filter(([k2]) => sem[k2]).map(([, n]) => `<w:${n} w:val="nil"/>`).join('')}</w:tcBorders>` : '';
        tcs += `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${c.c2 > c.c1 ? `<w:gridSpan w:val="${c.c2 - c.c1 + 1}"/>` : ''}` +
          `${c.r2 > c.r1 ? (r === c.r1 ? '<w:vMerge w:val="restart"/>' : '<w:vMerge/>') : ''}${tb}` +
          `${c.fundo ? `<w:shd w:val="clear" w:color="auto" w:fill="${c.fundo}"/>` : ''}${c.rot ? '<w:textDirection w:val="btLr"/>' : ''}<w:vAlign w:val="${c.va || 'center'}"/></w:tcPr>` +
          `${r === c.r1 && c.xml ? c.xml : '<w:p/>'}</w:tc>`;
        col = c.c2 + 1;
      }
      if (col <= grade.length) tcs += vazio(col, grade.length);
      linhas += `<w:tr><w:trPr><w:cantSplit/>${r <= nCab ? '<w:tblHeader/>' : ''}<w:trHeight w:val="${TW(hmm)}" w:hRule="exact"/></w:trPr>${tcs}</w:tr>`;
    });
    return `<w:tbl><w:tblPr><w:tblW w:w="${grade.reduce((s, x) => s + x, 0)}" w:type="dxa"/>${bordas}<w:tblLayout w:type="fixed"/>` +
      `<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="28" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="28" w:type="dxa"/></w:tblCellMar>${extra}</w:tblPr>` +
      `<w:tblGrid>${grade.map(x => `<w:gridCol w:w="${x}"/>`).join('')}</w:tblGrid>${linhas}</w:tbl>`;
  }
  const quebra = '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="2"/></w:rPr><w:br w:type="page"/></w:r></w:p>';
  const tabelaSemBorda = (larguras, celulasXml) => `<w:tbl><w:tblPr><w:tblW w:w="${larguras.reduce((s, x) => s + x, 0)}" w:type="dxa"/><w:jc w:val="center"/><w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(b => `<w:${b} w:val="nil"/>`).join('')}</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="28" w:type="dxa"/><w:right w:w="28" w:type="dxa"/></w:tblCellMar></w:tblPr>` +
    `<w:tblGrid>${larguras.map(x => `<w:gridCol w:w="${x}"/>`).join('')}</w:tblGrid><w:tr>${celulasXml.map((x, k) => `<w:tc><w:tcPr><w:tcW w:w="${larguras[k]}" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>${x}</w:tc>`).join('')}</w:tr></w:tbl>`;

  // ---------------- cabeçalho (mesma grade do PDF) ----------------
  const c = doc.cab, rv = revisaoAtual(), l = listasCabLinha();
  const icones = {
    qualidade: await img(BIB.icones.qualidade), meio_ambiente: await img(BIB.icones.meio_ambiente),
    energia: await img(BIB.icones.energia), seguranca: await img(cruzSegurancaPNG()),
  };
  const basicos = [];
  for (const b of c.episBasicos) { const m = await img(imgEpi(b)); if (m) basicos.push(m); }
  const n = basicos.length;
  const t1 = Math.min(12, 60 / Math.max(n, 1) - 1.5, 13), t2 = Math.min(6.25, 60 / Math.ceil(n / 2) - 1.5);
  const lado = n > 4 && t2 > t1 ? t2 : t1;
  const linhasBas = n > 4 && t2 > t1 ? [basicos.slice(0, Math.ceil(n / 2)), basicos.slice(Math.ceil(n / 2))] : [basicos];
  const xmlBasicos = linhasBas.map(lin => par(lin.map(m => desenho(m, lado, lado) + run('  ', {sz: 6})).join(''))).join('') || '<w:p/>';
  const ic = k => icones[k] ? par(desenho(icones[k], 5, 5)) : '<w:p/>';
  const B = (t, o = {}) => txt(t, o);
  function celulasCabecalho() {
    return [
      {c1: 1, c2: 3, r1: 1, r2: 3, fundo: AZUL, xml: B('PROCEDIMENTO OPERACIONAL', {sz: 7, b: true, cor: 'FFFFFF'})},
      {c1: 4, c2: 4, r1: 1, r2: 3, xml: B(c.codigo || 'CÓDIGO DO PO', {sz: 16, b: true})},
      {c1: 5, c2: 12, r1: 1, r2: 3, xml: B((c.nome || 'NOME DO PO').toUpperCase(), {sz: 11, b: true})},
      {c1: 13, c2: 13, r1: 1, r2: 1, xml: B('Revisão n°:', {sz: 7.5, b: true, jc: 'right'})},
      {c1: 14, c2: 14, r1: 1, r2: 1, xml: B(rv.rev, {sz: 7.5, b: true})},
      {c1: 15, c2: 15, r1: 1, r2: 1, xml: B('Data:', {sz: 7.5, b: true})},
      {c1: 16, c2: 17, r1: 1, r2: 1, xml: B(rv.data || '__/__/____', {sz: 7.5, b: true})},
      {c1: 13, c2: 13, r1: 2, r2: 2, xml: B('Elaborado por:', {sz: 7.5, b: true, jc: 'right'})},
      {c1: 14, c2: 14, r1: 2, r2: 2, xml: B(c.elaborado, {sz: 7.5, jc: 'left'})},
      {c1: 13, c2: 13, r1: 3, r2: 3, xml: B('Aprovado por:', {sz: 7.5, b: true, jc: 'right'})},
      {c1: 14, c2: 14, r1: 3, r2: 3, xml: B(c.aprovado, {sz: 7.5, jc: 'left'})},
      {c1: 15, c2: 15, r1: 2, r2: 3, xml: B('Tempo de ciclo:', {sz: 7.5, b: true})},
      {c1: 16, c2: 17, r1: 2, r2: 3, xml: B(c.tempoCiclo, {sz: 7.5, b: true})},
      {c1: 1, c2: 3, r1: 4, r2: 7, fundo: AZUL, xml: B("EPI'S BÁSICOS", {sz: 8, b: true, cor: 'FFFFFF'})},
      {c1: 4, c2: 5, r1: 4, r2: 7, xml: xmlBasicos},
      {c1: 6, c2: 6, r1: 4, r2: 7, fundo: AZUL, xml: B('LEGENDA', {sz: 8, b: true, cor: 'FFFFFF'})},
      {c1: 7, c2: 7, r1: 4, r2: 5, xml: ic('qualidade'), sem: {r: 1, b: 1}},
      {c1: 8, c2: 8, r1: 4, r2: 5, xml: B('Qualidade', {sz: 6.5, b: true}), sem: {l: 1, r: 1, b: 1}},
      {c1: 9, c2: 9, r1: 4, r2: 5, xml: ic('seguranca'), sem: {l: 1, r: 1, b: 1}},
      {c1: 10, c2: 12, r1: 4, r2: 5, xml: B('Segurança', {sz: 6.5, b: true}), sem: {l: 1, b: 1}},
      {c1: 7, c2: 7, r1: 6, r2: 7, xml: ic('meio_ambiente'), sem: {r: 1, t: 1}},
      {c1: 8, c2: 8, r1: 6, r2: 7, xml: B('Meio Ambiente', {sz: 6.5, b: true}), sem: {l: 1, r: 1, t: 1}},
      {c1: 9, c2: 9, r1: 6, r2: 7, xml: ic('energia'), sem: {l: 1, r: 1, t: 1}},
      {c1: 10, c2: 12, r1: 6, r2: 7, xml: B('Energia', {sz: 6.5, b: true}), sem: {l: 1, t: 1}},
      {c1: 13, c2: 13, r1: 4, r2: 4, fundo: AZUL, xml: B('FERRAMENTAS', {sz: 7.5, b: true, cor: 'FFFFFF'})},
      {c1: 14, c2: 17, r1: 4, r2: 4, xml: B(l.ferramentas, {sz: 7.5, jc: 'left'})},
      {c1: 13, c2: 13, r1: 5, r2: 6, fundo: AZUL, xml: B('MATERIAIS', {sz: 7.5, b: true, cor: 'FFFFFF'})},
      {c1: 14, c2: 17, r1: 5, r2: 6, xml: B(l.materiais, {sz: 7.5, jc: 'left'})},
      {c1: 13, c2: 13, r1: 7, r2: 7, fundo: AZUL, xml: B('EQUIPAMENTOS', {sz: 7.5, b: true, cor: 'FFFFFF'})},
      {c1: 14, c2: 17, r1: 7, r2: 7, xml: B(l.equipamentos, {sz: 7.5, jc: 'left'})},
    ];
  }
  const ALT_CAB = [8, 4, 4, 5, 2.5, 2.5, 5];
  const titulos = r => {
    const T = (t, sz) => B(t, {sz, b: true, cor: 'FFFFFF'});
    return [
      {c1: 1, c2: 1, r1: r, r2: r, fundo: AZUL, xml: T('Nº', 5)},
      {c1: 2, c2: 2, r1: r, r2: r, fundo: AZUL, rot: true, xml: T('NOME ETAPA', 5)},
      {c1: 3, c2: 5, r1: r, r2: r, fundo: AZUL, xml: T('FOTOS', 8)},
      {c1: 6, c2: 6, r1: r, r2: r, fundo: AZUL, xml: T('PONTOS DE ATENÇÃO', 7)},
      {c1: 7, c2: 13, r1: r, r2: r, fundo: AZUL, xml: T('DESCRIÇÃO', 8)},
      {c1: 14, c2: 14, r1: r, r2: r, fundo: AZUL, xml: T('PERIGOS', 8)},
      {c1: 15, c2: 17, r1: r, r2: r, fundo: AZUL, xml: T("EPI'S ESPECÍFICOS", 8)},
    ];
  };

  // ---------------- blocos das etapas ----------------
  function runsDescricao(html) {
    const out = [];
    const corpo = new DOMParser().parseFromString('<body>' + limparLinha(html), 'text/html').body;
    const andar = (nó, f) => {
      if (nó.nodeType === 3) { if (nó.nodeValue) out.push(run(nó.nodeValue, {sz: 8, ...f})); return; }
      if (nó.nodeType !== 1) return;
      const t = nó.tagName.toLowerCase(), g = {...f};
      if (t === 'b') g.b = true; if (t === 'i') g.i = true; if (t === 'u') g.u = true;
      const tam = /font-size:\s*([\d.]+)pt/i.exec(nó.getAttribute('style') || '');
      if (tam) g.sz = +tam[1];
      const corT = t !== 'mark' && /(?:^|;)\s*color:\s*#([0-9a-f]{6})/i.exec(nó.getAttribute('style') || '');
      if (corT) g.cor = corT[1].toUpperCase();
      const fonteT = fonteDe((/font-family:\s*([^;]+)/i.exec(nó.getAttribute('style') || '') || [])[1]);
      if (fonteT) g.fonte = fonteT;
      if (t === 'mark') {
        const st = nó.getAttribute('style') || '';
        const m = /background-color:\s*([^;]+)/i.exec(st), m2 = /(?:^|;)\s*color:\s*([^;]+)/i.exec(st);
        const hex = x => { const s = (x || '').trim(); if (/^#[0-9a-f]{6}$/i.test(s)) return s.slice(1).toUpperCase(); const r = /rgba?\((\d+),\s*(\d+),\s*(\d+)/i.exec(s); return r ? r.slice(1, 4).map(v => (+v).toString(16).padStart(2, '0')).join('').toUpperCase() : null; };
        if (m && hex(m[1])) g.fundo = hex(m[1]);
        if (m2 && hex(m2[1])) g.cor = hex(m2[1]);
      }
      nó.childNodes.forEach(x => andar(x, g));
    };
    corpo.childNodes.forEach(x => andar(x, {}));
    return out.join('');
  }
  async function pilhaXml(itens, larguraMm) {
    let x = '';
    for (const it of itens) {
      const m = await img(it.src);
      if (m) { const t = caberEm(m.w / m.h, 21.6, 18); x += par(desenho(m, t.w, t.h), {antes: 60}); }
      x += txt(it.rotulo || '', {sz: 7, b: true, depois: 40});
    }
    return x || '<w:p/>';
  }
  async function blocoXml(b, num) {
    const H = PAG[b.tam], nS = b.subs.length;
    const leg = doc.opcoes.legendaFotos ? 5 : 0;
    const reserva = b.grupos.reduce((s, g) => s + reservaLegenda(g), 0);
    const maxH = b.grupos.length ? Math.min(FOTO_MAX.alt, (H - 4 - reserva - vaoFotos() * (b.grupos.length - 1)) / b.grupos.length - 1) : 0;
    // fotos
    // espaço fixo entre as linhas de fotos (só entre elas: o conjunto fica centralizado na célula):
    // 3 mm sem legenda; com legenda, 1 mm visível das letras até a foto de baixo (06/11/2026)
    const vao3 = `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="${TW(vaoFotos())}" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/></w:rPr></w:pPr></w:p>`;
    let fotos = '';
    for (const g of b.grupos) {
      if (fotos) fotos += vao3;
      const partes = [];
      for (const f of g) {
        const lim = g.length === 2 ? FOTO_LADO : FOTO_MAX;
        const m = await img(f.img);
        if (!m) continue;
        const fa = f.aspecto ? f : Object.assign({}, f, {aspecto: m.w / m.h});
        const t = tamanhoFoto(fa, maxH, lim) || caberEm(fa.aspecto, lim.larg, Math.min(lim.alt, maxH));
        const legenda = leg ? `FOTO ${pad2(f.num)}` + (f.legenda ? ` – ${f.legenda}` : '') : '';
        partes.push({xml: par(desenho(m, t.w, t.h, true)) + (legenda ? txt(legenda, {sz: 6.5, b: true, antes: TW(0.7), linha: 140, exata: true}) : ''), w: t.w});   // legenda: 1 mm visível da foto até as letras (06/11/2026)
      }
      // lado a lado: célula de cada foto = foto + margens (0,5 mm de cada lado) e uma coluna de 2 mm no meio: 3 mm entre as fotos
      if (partes.length === 2) fotos += tabelaSemBorda([TW(Math.max(partes[0].w, 30) + 1), TW(2), TW(Math.max(partes[1].w, 30) + 1)], [partes[0].xml, '<w:p/>', partes[1].xml]);
      else if (partes.length) fotos += partes[0].xml;
    }
    // descrição
    // cada linha: marcador e recuo do nível (4 mm por nível), alinhamento escolhido no editor
    const desc = b.bullets.map(it => par((it.mk ? run(it.mk + '\t', {sz: 8}) : '') + runsDescricao(it.html),
      {jc: {j: 'both', e: 'left', c: 'center', d: 'right'}[it.a] || 'both', esq: TW((it.n - 1) * 4 + (it.mk ? 4 : 0)), pend: it.mk ? TW(4) : 0, depois: 57})).join('');
    const pontos = b.etapa.pontos.map(k => icones[k] ? par(desenho(icones[k], 13, 13), {depois: 170}) : '').join('') || '<w:p/>';
    const nR = 1 + nS, alturas = [H - b.hSubs.reduce((a, x) => a + x, 0), ...b.hSubs];
    const celulas = [
      {c1: 1, c2: 1, r1: 1, r2: nR, fundo: AZUL, xml: txt(String(num), {sz: 6.5, b: true, cor: 'FFFFFF'})},
      {c1: 2, c2: 2, r1: 1, r2: nR, fundo: AZUL, rot: true, xml: txt((b.etapa.nome || '').toUpperCase(), {sz: 5, b: true, cor: 'FFFFFF'})},
      {c1: 3, c2: 5, r1: 1, r2: nR, xml: fotos || '<w:p/>'},
      {c1: 6, c2: 6, r1: 1, r2: nR, xml: pontos},
      {c1: 7, c2: 13, r1: 1, r2: 1, xml: desc || '<w:p/>'},
      ...b.subs.flatMap((s, k) => [
        {c1: 7, c2: 10, r1: 2 + k, r2: 2 + k, fundo: AZUL, xml: txt(s.tipo, {sz: 8, b: true, cor: 'FFFFFF'})},
        {c1: 11, c2: 13, r1: 2 + k, r2: 2 + k, xml: txt(s.valor, {sz: 8, jc: 'both'})},
      ]),
      {c1: 14, c2: 14, r1: 1, r2: nR, xml: await pilhaXml(b.perigos.map(p => ({src: imgPerigo(p), rotulo: p.rotulo})))},
      {c1: 15, c2: 17, r1: 1, r2: nR, xml: episPO(b.etapa).length ? await pilhaXml(b.epis.map(p => ({src: imgEpi(p), rotulo: p.rotulo}))) : txt('N/A', {sz: 8, b: true})},
    ];
    return {alturas, celulas};
  }
  // Uma tabela só para PO, anexos e lista de treinamento: as 7 linhas do cabeçalho são marcadas como
  // "Repetir linhas de cabeçalho" do Word. Cada página soma sempre a mesma altura (168 mm abaixo do
  // cabeçalho), completando com uma linha vazia sem bordas, para o Word quebrar a página no lugar certo.
  const ALT_PAGINA = PAG.col + PAG.inteira;
  const docAlturas = [...ALT_CAB], docCelulas = celulasCabecalho();
  function pagina(partes, comTitulos) {
    const ini = docAlturas.length;
    if (comTitulos) { docAlturas.push(PAG.col); docCelulas.push(...titulos(docAlturas.length)); }
    for (const p of partes) {
      const base = docAlturas.length;
      docAlturas.push(...p.alturas);
      docCelulas.push(...p.celulas.map(x => Object.assign({}, x, {r1: x.r1 + base, r2: x.r2 + base})));
    }
    const usado = docAlturas.slice(ini).reduce((s, h) => s + h, 0);
    if (ALT_PAGINA - usado > 2) {
      docAlturas.push(ALT_PAGINA - usado);
      const r = docAlturas.length;
      docCelulas.push({c1: 1, c2: 17, r1: r, r2: r, xml: '<w:p/>', sem: {l: 1, r: 1, b: 1}});
    }
  }

  const paginas = [];
  prepararMedidor();
  const blocos = blocosDoc();
  let nf = 0;
  blocos.forEach(b => { b.grupos = b.grupos.map(g => g.map(f => Object.assign({}, f, {num: ++nf}))); });
  for (let i = 0; i < blocos.length; i++) {
    const b = blocos[i], prox = blocos[i + 1];
    if (b.tam === 'meia' && prox && prox.tam === 'meia' && !prox.novaPagina) { pagina([await blocoXml(b, b.num), await blocoXml(prox, prox.num)], true); i++; }
    else pagina([await blocoXml(b, b.num)], true);
  }
  // anexos
  for (let i = 0; i < doc.anexos.length; i++) {
    const a = doc.anexos[i];
    const imgsA = [];
    for (const im of a.imagens) { const m = await imagemXL(im.img); imgsA.push(im.aspecto || !m ? im : Object.assign({}, im, {aspecto: m.w / m.h})); }
    const alt = ALT_PAGINA - 5;
    const k = Math.min(280 / ANEXO_MAX.larg, (alt - 6) / ANEXO_MAX.alt) * 0.97;
    let xml = '';
    for (const ln of layoutAnexo(imgsA)) {
      let r = '';
      for (const it of ln) { const m = await img(it.im.img); if (m) r += desenho(m, it.w * k, it.h * k) + run('   ', {sz: 8}); }
      xml += par(r, {depois: 113});
    }
    pagina([{alturas: [5, alt], celulas: [
      {c1: 1, c2: 17, r1: 1, r2: 1, fundo: AZUL, xml: txt(`ANEXO ${i + 1} – ${(a.titulo || '').toUpperCase()}`, {sz: 9, b: true, cor: 'FFFFFF'})},
      {c1: 1, c2: 17, r1: 2, r2: 2, xml: xml || '<w:p/>'},
    ]}], false);
  }
  // lista de treinamento
  if (doc.treinamento.incluir) {
    const hl = (PAG.inteira) / LINHAS_TREINAMENTO;
    const T = t => txt(t, {sz: 8, b: true, cor: 'FFFFFF'});
    const cols = [[1, 5, 'NOME COMPLETO DO COLABORADOR'], [6, 6, 'NÚMERO REGISTRO'], [7, 12, 'TURNO'], [13, 13, 'DATA'], [14, 17, 'ASSINATURA']];
    const cel = cols.map(([c1, c2, t]) => ({c1, c2, r1: 1, r2: 1, fundo: AZUL, xml: T(t)}));
    for (let k = 0; k < LINHAS_TREINAMENTO; k++) cel.push(...cols.map(([c1, c2]) => ({c1, c2, r1: 2 + k, r2: 2 + k, xml: '<w:p/>'})));
    pagina([{alturas: [PAG.col, ...Array(LINHAS_TREINAMENTO).fill(hl)], celulas: cel}], false);
  }
  paginas.push(tabela(G17, docAlturas, docCelulas, '', ALT_CAB.length));
  // histórico de revisões (tabela própria, sem o cabeçalho do PO)
  // logotipo escolhido no cabeçalho, no lugar do VM3; o da Villares Metals saiu (05/11/2026)
  const vm3 = doc.cab.logo ? await img(doc.cab.logo) : null, tl = tamanhoLogoHist();
  const total = Math.max(LINHAS_HISTORICO, Math.ceil(doc.historico.length / LINHAS_HISTORICO) * LINHAS_HISTORICO);
  for (let ini = 0; ini < total; ini += LINHAS_HISTORICO) {
    const cel = [
      {c1: 1, c2: 3, r1: 1, r2: 1, fundo: AZUL, sem: {r: 1}, xml: vm3 ? par(desenho(vm3, tl.w, tl.h), {jc: 'left'}) : '<w:p/>'},
      {c1: 4, c2: 7, r1: 1, r2: 1, fundo: AZUL, sem: {l: 1, r: 1}, xml: txt('HISTÓRICO DE REVISÕES', {sz: 12, b: true, cor: 'FFFFFF'})},
      {c1: 8, c2: 9, r1: 1, r2: 1, fundo: AZUL, sem: {l: 1}, xml: '<w:p/>'},
      ...[[1, 1, 'Item'], [2, 2, 'Descrição da revisão'], [3, 4, 'Foto Alteração'], [5, 5, 'Data'], [6, 6, 'Nome Solicitante'], [7, 8, 'Responsável pela revisão'], [9, 9, 'Observações']]
        .map(([c1, c2, t]) => ({c1, c2, r1: 2, r2: 2, fundo: CINZA, xml: txt(t, {sz: 8, b: true, cor: 'FFFFFF'})})),
    ];
    for (let i = ini; i < ini + LINHAS_HISTORICO; i++) {
      const h = doc.historico[i] || {}, r = 3 + i - ini, V = t => txt(t, {sz: 8, b: true});
      // foto da alteração no tamanho do PO (máx. 4,4 × 2,5 cm); sem foto, "-"
      const fm = h.fotoImg && await img(h.fotoImg.img), tf = fm && tamanhoFoto(h.fotoImg, FOTO_HIST.alt, FOTO_HIST);
      const xmlFoto = tf ? par(desenho(fm, tf.w * 0.97, tf.h * 0.97)) : V(textoFotoHist(h));
      cel.push({c1: 1, c2: 1, r1: r, r2: r, xml: V(String(i))}, {c1: 2, c2: 2, r1: r, r2: r, xml: V(h.descricao || '')},
        {c1: 3, c2: 4, r1: r, r2: r, xml: xmlFoto}, {c1: 5, c2: 5, r1: r, r2: r, xml: V(fmtData(h.data))},
        {c1: 6, c2: 6, r1: r, r2: r, xml: V(h.solicitante || '')}, {c1: 7, c2: 8, r1: r, r2: r, xml: V(h.responsavel || '')},
        {c1: 9, c2: 9, r1: r, r2: r, xml: V(h.obs || '')});
    }
    paginas.push(tabela(GH, [10, 5, ...Array(LINHAS_HISTORICO).fill(26)], cel));
  }

  // ---------------- pacote .docx ----------------
  const corpo = paginas.join(quebra) + '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/></w:rPr></w:pPr></w:p>';
  const sect = `<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="${TW(5)}" w:right="${TW(5)}" w:bottom="${TW(4)}" w:left="${TW(5)}" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>`;
  const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>${corpo}${sect}</w:body></w:document>`;
  const estilos = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" w:eastAsia="Arial"/><w:sz w:val="16"/><w:szCs w:val="16"/><w:lang w:val="pt-BR"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="table" w:default="1" w:styleId="TabelaNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="28" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="28" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style></w:styles>`;
  const enc = new TextEncoder();
  const b64 = s => { const bin = atob(s); const u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i); return u; };
  const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdEstilos" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rIdDadosPO" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/>` +
    midias.map((m, i) => `<Relationship Id="${m.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/imagem${i + 1}.${m.ext}"/>`).join('') + `</Relationships>`;
  const arq = [
    {nome: '[Content_Types].xml', dados: enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/></Types>`)},
    {nome: '_rels/.rels', dados: enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`)},
    {nome: 'word/document.xml', dados: enc.encode(docXml)},
    {nome: 'word/styles.xml', dados: enc.encode(estilos)},
    {nome: 'word/_rels/document.xml.rels', dados: enc.encode(rels)},
    // dados completos do gerador guardados dentro do Word: o "Abrir" do gerador lê daqui e volta tudo como estava
    {nome: 'customXml/item1.xml', dados: enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><poGerador xmlns="urn:gerador-po" texto="${typeof IMPORTAR !== 'undefined' ? IMPORTAR.digitalTexto(docXml) : ''}"><![CDATA[${JSON.stringify(doc).replace(/]]>/g, ']]]]><![CDATA[>')}]]></poGerador>`)},
    {nome: 'customXml/itemProps1.xml', dados: enc.encode('<?xml version="1.0" encoding="UTF-8" standalone="no"?><ds:datastoreItem ds:itemID="{6E1B2C9A-3F4D-4B8E-9A7C-2D5E8F1A0B3C}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs><ds:schemaRef ds:uri="urn:gerador-po"/></ds:schemaRefs></ds:datastoreItem>')},
    {nome: 'customXml/_rels/item1.xml.rels', dados: enc.encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps1.xml"/></Relationships>')},
    ...midias.map((m, i) => ({nome: `word/media/imagem${i + 1}.${m.ext}`, dados: b64(m.url.split(',')[1])})),
  ];
  return XL.zip(arq);
}

async function exportarWord() {
  aviso('Montando o Word...');
  await compactarAgora(false);
  let blob;
  try { blob = await gerarWord(); }
  catch (e) { console.error(e); alert('Não consegui montar o Word: ' + erroAmigavel(e)); return; }
  // o "Salvar" do PO já grava este mesmo Word: a exportação sugere "- cópia" para não gravar por cima do arquivo de trabalho (05/11/2026)
  const nome = nomeArquivo() + ' - cópia.docx';
  if (window.showSaveFilePicker) {
    try {
      const h = await window.showSaveFilePicker({suggestedName: nome, types: [{description: 'Documento do Word', accept: {'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx']}}]});
      const w = await h.createWritable(); await w.write(blob); await w.close();
      aviso('Word salvo: ' + h.name); return;
    } catch (e) { if (e.name === 'AbortError') return; }
  }
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = nome;
  document.body.append(a); a.click(); a.remove();
  aviso('Word salvo na pasta Downloads: ' + nome);
}
