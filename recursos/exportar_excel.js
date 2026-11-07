// Exportação do PO para Excel (.xlsx) no layout do modelo B.00.183-00.
// O arquivo é montado aqui mesmo na página (sem precisar do Excel aberto): planilha, mesclagens,
// estilos, quebras de página e imagens posicionadas sobre as células.
'use strict';

const XL = (() => {
  // ---------------- zip (sem compressão) ----------------
  const CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  const crc32 = b => { let c = 0xFFFFFFFF; for (let i = 0; i < b.length; i++) c = CRC[(c ^ b[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; };
  function zip(arquivos) {           // arquivos: [{nome, dados: Uint8Array}]
    const enc = new TextEncoder(), partes = [], central = [];
    let pos = 0;
    for (const a of arquivos) {
      const nome = enc.encode(a.nome), d = a.dados, crc = crc32(d);
      const h = new DataView(new ArrayBuffer(30));
      h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true); h.setUint16(8, 0, true);
      h.setUint16(10, 0, true); h.setUint16(12, 0x21, true); h.setUint32(14, crc, true);
      h.setUint32(18, d.length, true); h.setUint32(22, d.length, true); h.setUint16(26, nome.length, true); h.setUint16(28, 0, true);
      partes.push(new Uint8Array(h.buffer), nome, d);
      const c = new DataView(new ArrayBuffer(46));
      c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(8, 0x0800, true);
      c.setUint16(10, 0, true); c.setUint16(12, 0, true); c.setUint16(14, 0x21, true); c.setUint32(16, crc, true);
      c.setUint32(20, d.length, true); c.setUint32(24, d.length, true); c.setUint16(28, nome.length, true);
      c.setUint32(42, pos, true);
      central.push(new Uint8Array(c.buffer), nome);
      pos += 30 + nome.length + d.length;
    }
    const tamCentral = central.reduce((s, x) => s + x.length, 0);
    const fim = new DataView(new ArrayBuffer(22));
    fim.setUint32(0, 0x06054b50, true); fim.setUint16(8, arquivos.length, true); fim.setUint16(10, arquivos.length, true);
    fim.setUint32(12, tamCentral, true); fim.setUint32(16, pos, true);
    return new Blob([...partes, ...central, new Uint8Array(fim.buffer)], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  }

  const esc = s => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
  const letra = n => { let s = ''; for (n++; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s; return s; };
  const colNum = l => l.split('').reduce((s, c) => s * 26 + c.charCodeAt(0) - 64, 0) - 1;

  // ---------------- estilos ----------------
  function Estilos() {
    const fonts = ['<font><sz val="10"/><name val="Arial"/></font>'], fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
    const borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>'], xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
    const idx = (lista, xml) => { let i = lista.indexOf(xml); if (i < 0) { lista.push(xml); i = lista.length - 1; } return i; };
    const cache = new Map();
    return {
      // e: {sz, b, i, cor, fundo, borda, h, v, wrap, rot, texto}
      id(e = {}) {
        const k = JSON.stringify(e);
        if (cache.has(k)) return cache.get(k);
        const f = idx(fonts, `<font>${e.b ? '<b/>' : ''}${e.i ? '<i/>' : ''}<sz val="${e.sz || 10}"/><color rgb="FF${e.cor || '000000'}"/><name val="Arial"/></font>`);
        const fi = e.fundo ? idx(fills, `<fill><patternFill patternType="solid"><fgColor rgb="FF${e.fundo}"/><bgColor indexed="64"/></patternFill></fill>`) : 0;
        const tem = t => e.borda === true || (typeof e.borda === 'string' && e.borda.includes(t[0]));   // borda: true ou 'ltrb' parcial
        const lado = t => tem(t) ? `<${t} style="thin"><color rgb="FF000000"/></${t}>` : `<${t}/>`;
        const bo = e.borda ? idx(borders, `<border>${lado('left')}${lado('right')}${lado('top')}${lado('bottom')}<diagonal/></border>`) : 0;
        const al = `<alignment horizontal="${e.h || 'general'}" vertical="${e.v || 'center'}"${e.wrap ? ' wrapText="1"' : ''}${e.rot ? ` textRotation="${e.rot}"` : ''}/>`;
        xfs.push(`<xf numFmtId="${e.texto ? 49 : 0}" fontId="${f}" fillId="${fi}" borderId="${bo}" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">${al}</xf>`);
        cache.set(k, xfs.length - 1);
        return xfs.length - 1;
      },
      xml() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
          `<fonts count="${fonts.length}">${fonts.join('')}</fonts><fills count="${fills.length}">${fills.join('')}</fills>` +
          `<borders count="${borders.length}">${borders.join('')}</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
          `<cellXfs count="${xfs.length}">${xfs.join('')}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
      },
    };
  }

  // ---------------- planilha ----------------
  function Planilha(nome, larguras, estilos) {       // larguras: {A: chars, ...}
    const celulas = new Map(), alturas = new Map(), mescla = [], quebras = [], desenhos = [];
    const pxCol = w => Math.trunc(((256 * w + Math.trunc(128 / 7)) / 256) * 7);
    const colsOrd = Object.keys(larguras).sort((a, b) => colNum(a) - colNum(b));
    const esquerda = l => { let s = 0; for (const c of colsOrd) { if (colNum(c) >= colNum(l)) break; s += pxCol(larguras[c]) * 0.75; } return s; };
    const largura = (a, b = a) => { let s = 0; for (const c of colsOrd) if (colNum(c) >= colNum(a) && colNum(c) <= colNum(b)) s += pxCol(larguras[c]) * 0.75; return s; };
    const p = {
      nome, quebras, desenhos,
      altura(r, h) { alturas.set(r, h); },
      topo(r) { let s = 0; for (let i = 1; i < r; i++) s += alturas.get(i) ?? 15; return s; },
      alturaLinhas(a, b) { let s = 0; for (let i = a; i <= b; i++) s += alturas.get(i) ?? 15; return s; },
      esquerda, largura,
      // posição em pt -> célula (0-based) + deslocamento dentro dela
      localizar(x, y) {
        let col = 0, acc = 0;
        for (let c = 0; c < 200; c++) {
          const w = larguras[letra(c)] != null ? pxCol(larguras[letra(c)]) * 0.75 : pxCol(8.43) * 0.75;
          if (acc + w > x) { col = c; break; }
          acc += w; col = c + 1;
        }
        let row = 0, accR = 0;
        for (let r = 1; r < 5000; r++) {
          const h = alturas.get(r) ?? 15;
          if (accR + h > y) { row = r - 1; break; }
          accR += h; row = r;
        }
        return {col, dx: Math.max(0, x - acc), row, dy: Math.max(0, y - accR)};
      },
      cel(ref, valor, estilo) { celulas.set(ref, {valor, s: estilos.id(estilo)}); },
      estilizar(area, estilo) {          // estilo em todas as células da área (para borda de mesclagem)
        const [a, b] = area.split(':'); const ma = /([A-Z]+)(\d+)/.exec(a), mb = /([A-Z]+)(\d+)/.exec(b || a);
        for (let r = +ma[2]; r <= +mb[2]; r++) for (let c = colNum(ma[1]); c <= colNum(mb[1]); c++) {
          const ref = letra(c) + r; const atual = celulas.get(ref);
          celulas.set(ref, {valor: atual ? atual.valor : undefined, s: estilos.id(estilo)});
        }
      },
      bloco(area, valor, estilo) {       // mescla + valor + estilo
        const [a, b] = area.split(':');
        if (b && b !== a) mescla.push(area);
        p.estilizar(area, estilo);
        celulas.set(a, {valor, s: estilos.id(estilo)});
      },
      xml(rIdDesenho) {
        const linhas = new Map();
        for (const [ref, c] of celulas) { const m = /([A-Z]+)(\d+)/.exec(ref); const r = +m[2]; if (!linhas.has(r)) linhas.set(r, []); linhas.get(r).push([colNum(m[1]), ref, c]); }
        for (const r of alturas.keys()) if (!linhas.has(r)) linhas.set(r, []);
        const corpo = [...linhas.keys()].sort((a, b) => a - b).map(r => {
          const cs = linhas.get(r).sort((a, b) => a[0] - b[0]).map(([, ref, c]) => {
            const v = c.valor;
            if (v === undefined || v === null || v === '') return `<c r="${ref}" s="${c.s}"/>`;
            if (typeof v === 'number') return `<c r="${ref}" s="${c.s}"><v>${v}</v></c>`;
            if (v && typeof v === 'object' && !Array.isArray(v) && v.f) return `<c r="${ref}" s="${c.s}" t="str"><f>${esc(v.f)}</f><v>${esc(v.v)}</v></c>`;
            if (Array.isArray(v)) return `<c r="${ref}" s="${c.s}" t="inlineStr"><is>${v.map(run =>
              `<r><rPr>${run.b ? '<b/>' : ''}${run.i ? '<i/>' : ''}${run.u ? '<u/>' : ''}<sz val="${run.sz}"/><color rgb="FF${run.cor || '000000'}"/><rFont val="${esc(run.fonte || 'Arial')}"/></rPr><t xml:space="preserve">${esc(run.t)}</t></r>`).join('')}</is></c>`;
            return `<c r="${ref}" s="${c.s}" t="inlineStr"><is><t xml:space="preserve">${esc(v)}</t></is></c>`;
          }).join('');
          const h = alturas.get(r);
          return `<row r="${r}"${h != null ? ` ht="${h.toFixed(2)}" customHeight="1"` : ''}>${cs}</row>`;
        }).join('');
        const cols = colsOrd.map(c => `<col min="${colNum(c) + 1}" max="${colNum(c) + 1}" width="${larguras[c]}" customWidth="1"/>`).join('');
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
          `<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><sheetViews><sheetView showGridLines="0" zoomScale="${p.zoom || 50}" workbookViewId="0"/></sheetViews>` +
          `<sheetFormatPr defaultRowHeight="15"/><cols>${cols}</cols><sheetData>${corpo}</sheetData>` +
          (mescla.length ? `<mergeCells count="${mescla.length}">${mescla.map(m => `<mergeCell ref="${m}"/>`).join('')}</mergeCells>` : '') +
          `<pageMargins left="${(p.margens || {}).l ?? 0.16}" right="${(p.margens || {}).r ?? 0.24}" top="${(p.margens || {}).t ?? 0.32}" bottom="${(p.margens || {}).b ?? 0.24}" header="0" footer="0"/>` +
          `<pageSetup paperSize="9" orientation="${p.retrato ? 'portrait' : 'landscape'}" fitToWidth="1" fitToHeight="${p.umaPagina ? 1 : 0}"/>` +
          (quebras.length ? `<rowBreaks count="${quebras.length}" manualBreakCount="${quebras.length}">${quebras.map(r => `<brk id="${r - 1}" max="16383" man="1"/>`).join('')}</rowBreaks>` : '') +
          (desenhos.length ? `<drawing r:id="${rIdDesenho}"/>` : '') + `</worksheet>`;
      },
    };
    return p;
  }

  // ---------------- desenhos (imagens e caixas de texto) ----------------
  const EMU = 12700;
  // ancora cada imagem/caixa na célula onde ela começa (com deslocamento), para acompanhar as linhas na impressão
  function desenhoXml(itens, relImg, localizar) {
    let id = 1;
    const ancora = (it, miolo) => {
      const l = localizar(it.x, it.y);
      return `<xdr:oneCellAnchor><xdr:from><xdr:col>${l.col}</xdr:col><xdr:colOff>${Math.round(l.dx * EMU)}</xdr:colOff><xdr:row>${l.row}</xdr:row><xdr:rowOff>${Math.round(l.dy * EMU)}</xdr:rowOff></xdr:from>` +
        `<xdr:ext cx="${Math.round(it.w * EMU)}" cy="${Math.round(it.h * EMU)}"/>${miolo}<xdr:clientData/></xdr:oneCellAnchor>`;
    };
    const xfrm = it => `<a:xfrm><a:off x="${Math.round(it.x * EMU)}" y="${Math.round(it.y * EMU)}"/><a:ext cx="${Math.round(it.w * EMU)}" cy="${Math.round(it.h * EMU)}"/></a:xfrm>`;
    const corpo = itens.map(it => {
      id++;
      if (it.tipo === 'img') return ancora(it, `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${id}" name="Imagem ${id}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>` +
        `<xdr:blipFill><a:blip r:embed="${relImg(it.img)}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr>${xfrm(it)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
        `${it.moldura ? '<a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>' : ''}</xdr:spPr></xdr:pic>`);
      return ancora(it, `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${id}" name="Texto ${id}"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr>` +
        `<xdr:spPr>${xfrm(it)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></xdr:spPr>` +
        `<xdr:txBody><a:bodyPr vertOverflow="overflow" horzOverflow="overflow" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"/><a:lstStyle/>` +
        `<a:p><a:pPr algn="${it.alinh || 'ctr'}"/><a:r><a:rPr lang="pt-BR" sz="${Math.round(it.sz * 100)}"${it.b ? ' b="1"' : ''}><a:solidFill><a:srgbClr val="${it.cor || '000000'}"/></a:solidFill><a:latin typeface="Arial"/><a:cs typeface="Arial"/></a:rPr><a:t>${esc(it.texto)}</a:t></a:r></a:p></xdr:txBody></xdr:sp>`);
    }).join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${corpo}</xdr:wsDr>`;
  }

  // ---------------- pasta de trabalho ----------------
  async function montar(planilhas, estilos, definidos, dados) {
    const enc = new TextEncoder(), arq = [], midias = new Map();   // dataURL -> nome
    const b64 = s => { const bin = atob(s); const u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i); return u; };
    const ct = [];
    planilhas.forEach((p, i) => {
      const n = i + 1;
      ct.push(`<Override PartName="/xl/worksheets/sheet${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`);
      let relsSheet = '';
      if (p.desenhos.length) {
        const rels = [];
        const relImg = img => {
          if (!midias.has(img.url)) midias.set(img.url, `image${midias.size + 1}.${img.ext}`);
          const alvo = `../media/${midias.get(img.url)}`;
          let r = rels.find(x => x.alvo === alvo);
          if (!r) { r = {id: `rId${rels.length + 1}`, alvo}; rels.push(r); }
          return r.id;
        };
        arq.push({nome: `xl/drawings/drawing${n}.xml`, dados: enc.encode(desenhoXml(p.desenhos, relImg, p.localizar))});
        arq.push({nome: `xl/drawings/_rels/drawing${n}.xml.rels`, dados: enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.map(r => `<Relationship Id="${r.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${r.alvo}"/>`).join('')}</Relationships>`)});
        ct.push(`<Override PartName="/xl/drawings/drawing${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`);
        relsSheet = `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${n}.xml"/>`;
      }
      arq.push({nome: `xl/worksheets/sheet${n}.xml`, dados: enc.encode(p.xml('rId1'))});
      if (relsSheet) arq.push({nome: `xl/worksheets/_rels/sheet${n}.xml.rels`, dados: enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relsSheet}</Relationships>`)});
    });
    for (const [url, nome] of midias) arq.push({nome: `xl/media/${nome}`, dados: b64(url.split(',')[1])});
    arq.push({nome: 'xl/styles.xml', dados: enc.encode(estilos.xml())});
    // dados completos do gerador guardados dentro do Excel: o "Abrir" do gerador lê daqui e volta tudo como estava
    let relDados = '';
    if (dados) {
      const dig = typeof IMPORTAR_XL !== 'undefined' ? IMPORTAR_XL.digitalFolhas(planilhas.map(p => p.xml('rId1'))) : '';
      arq.push({nome: 'customXml/item1.xml', dados: enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><poGerador xmlns="urn:gerador-po" texto="${dig}"><![CDATA[${dados.replace(/]]>/g, ']]]]><![CDATA[>')}]]></poGerador>`)});
      arq.push({nome: 'customXml/itemProps1.xml', dados: enc.encode('<?xml version="1.0" encoding="UTF-8" standalone="no"?><ds:datastoreItem ds:itemID="{6E1B2C9A-3F4D-4B8E-9A7C-2D5E8F1A0B3D}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs><ds:schemaRef ds:uri="urn:gerador-po"/></ds:schemaRefs></ds:datastoreItem>')});
      arq.push({nome: 'customXml/_rels/item1.xml.rels', dados: enc.encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps1.xml"/></Relationships>')});
      ct.push('<Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>');
      relDados = '<Relationship Id="rIdDadosPO" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/>';
    }
    arq.push({nome: 'xl/workbook.xml', dados: enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets>${planilhas.map((p, i) => `<sheet name="${esc(p.nome)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets>${definidos.length ? `<definedNames>${definidos.map(d => `<definedName name="${d.nome}" localSheetId="${d.folha}">${esc(d.ref)}</definedName>`).join('')}</definedNames>` : ''}</workbook>`)});
    arq.push({nome: 'xl/_rels/workbook.xml.rels', dados: enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${planilhas.map((p, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}<Relationship Id="rId${planilhas.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>${relDados}</Relationships>`)});
    arq.push({nome: '_rels/.rels', dados: enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`)});
    arq.push({nome: '[Content_Types].xml', dados: enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${ct.join('')}</Types>`)});
    arq.unshift(arq.pop());   // [Content_Types].xml primeiro
    return zip(arq);
  }
  return {Planilha, Estilos, montar, letra, zip, esc};
})();

// ---------------- imagens: tudo vira PNG/JPEG (o Excel não lê WebP) ----------------
const cacheImgXL = new Map();
function imagemXL(src) {
  if (cacheImgXL.has(src)) return cacheImgXL.get(src);
  const pr = new Promise(ok => {
    if (!src) return ok(null);
    const im = new Image();
    im.onload = () => {
      const jpeg = /^data:image\/jpe?g/.test(src);
      if (jpeg || /^data:image\/png/.test(src)) return ok({url: src.replace('image/jpg', 'image/jpeg'), ext: jpeg ? 'jpeg' : 'png', w: im.naturalWidth, h: im.naturalHeight});
      const c = document.createElement('canvas'); c.width = im.naturalWidth; c.height = im.naturalHeight;
      c.getContext('2d').drawImage(im, 0, 0);
      ok({url: c.toDataURL('image/png'), ext: 'png', w: im.naturalWidth, h: im.naturalHeight});
    };
    im.onerror = () => ok(null);
    im.src = src;
  });
  cacheImgXL.set(src, pr);
  return pr;
}
function cruzSegurancaPNG() {
  const c = document.createElement('canvas'); c.width = c.height = 200;
  const g = c.getContext('2d'); g.fillStyle = '#548235'; g.strokeStyle = '#2f4a1d'; g.lineWidth = 6;
  g.beginPath(); [[76, 20], [124, 20], [124, 76], [180, 76], [180, 124], [124, 124], [124, 180], [76, 180], [76, 124], [20, 124], [20, 76], [76, 76]].forEach(([x, y], i) => i ? g.lineTo(x, y) : g.moveTo(x, y));
  g.closePath(); g.fill(); g.stroke();
  return c.toDataURL('image/png');
}

// ---------------- o PO no layout do modelo ----------------
async function gerarExcel() {
  // tudo no tamanho real do gerador (mm -> pt), na mesma grade de 17 colunas do PDF
  const AZUL = '002060', AZ = AZUL, BRANCO = 'FFFFFF', CINZA = '808080';
  const MM = 72 / 25.4, ESC = MM;
  const ALT = {inteira: 158, meia: 79};          // mm do bloco (1 mm abaixo do PDF para o Excel não empurrar linhas)
  const FONTE = 8, FONTE_LEG = 6.5;   // FONTE_LEG: legenda das fotos
  const colunasMM = twips => { const o = {}; twips.forEach((t, i) => { o[XL.letra(i)] = Math.max(0, +((t / 20 / 0.75 - 2) / 7).toFixed(2)); }); return o; };
  const G17 = [197, 300, 1281, 2539, 943, 993, 568, 931, 627, 573, 277, 22, 2397, 2529, 846, 990, 247];
  const L = i => XL.letra(i - 1);                                // coluna da grade (1..17) -> letra
  const area = (c1, c2, r1, r2) => `${L(c1)}${r1}:${L(c2)}${r2}`;
  const est = XL.Estilos();
  const c = doc.cab, r = revisaoAtual(), l = listasCabLinha();
  const nomeFolha = ('PO.' + (c.codigo || 'XX.000')).replace(/[\[\]:*?\/\\]/g, '-').slice(0, 31);

  const icones = {
    qualidade: await imagemXL(BIB.icones.qualidade), meio_ambiente: await imagemXL(BIB.icones.meio_ambiente),
    energia: await imagemXL(BIB.icones.energia), seguranca: await imagemXL(cruzSegurancaPNG()),
  };
  const episBasicos = (await Promise.all(c.episBasicos.map(b => imagemXL(imgEpi(b))))).filter(Boolean);
  const txtW = (texto, sz, larg) => {      // altura aproximada de uma caixa de texto
    const porLinha = Math.max(1, Math.floor(larg / (sz * 0.55)));
    const linhas = String(texto).split('\n').reduce((s, t) => s + Math.max(1, Math.ceil(t.length / porLinha)), 0);
    return linhas * sz * 1.22 + 2;
  };

  function titulosColunas(p, linha) {
    p.altura(linha, 10 * MM);
    const T = {borda: true, wrap: true, h: 'center', b: true, cor: BRANCO, fundo: AZUL, sz: 8};
    p.bloco(area(1, 1, linha, linha), 'Nº', {...T, sz: 5});
    p.bloco(area(2, 2, linha, linha), 'NOME ETAPA', {...T, sz: 5, rot: 90});
    p.bloco(area(3, 5, linha, linha), 'FOTOS', T);
    p.bloco(area(6, 6, linha, linha), 'PONTOS DE ATENÇÃO', T);
    p.bloco(area(7, 13, linha, linha), 'DESCRIÇÃO', T);
    p.bloco(area(14, 14, linha, linha), 'PERIGOS', T);
    p.bloco(area(15, 17, linha, linha), "EPI'S ESPECÍFICOS", T);
  }
  // descrição com negrito/itálico/sublinhado (o realce não existe dentro de célula do Excel)
  // a descrição fica numa célula só: marcador e recuo do nível vão no texto (o alinhamento da célula é justificado)
  function runsDescricao(bullets) {
    const runs = [];
    bullets.forEach((it, k) => {
      if (k) runs.push({t: '\n', sz: FONTE});
      const recuo = it.n === 2 ? '      ' : '';
      if (recuo || it.mk) runs.push({t: recuo + (it.mk ? it.mk + ' ' : ''), sz: FONTE});
      const corpo = new DOMParser().parseFromString('<body>' + limparLinha(it.html), 'text/html').body;
      const andar = (n, f) => {
        if (n.nodeType === 3) { if (n.nodeValue) runs.push({t: n.nodeValue, sz: FONTE, ...f}); return; }
        if (n.nodeType !== 1) return;
        const t = n.tagName.toLowerCase(), g = {...f};
        if (t === 'b') g.b = true; if (t === 'i') g.i = true; if (t === 'u') g.u = true;
        const tam = /font-size:\s*([\d.]+)pt/i.exec(n.getAttribute('style') || '');
        if (tam) g.sz = +tam[1];
      const corT = t !== 'mark' && /(?:^|;)\s*color:\s*#([0-9a-f]{6})/i.exec(n.getAttribute('style') || '');
      if (corT) g.cor = corT[1].toUpperCase();
      const fonteT = fonteDe((/font-family:\s*([^;]+)/i.exec(n.getAttribute('style') || '') || [])[1]);
      if (fonteT) g.fonte = fonteT;
        n.childNodes.forEach(x => andar(x, g));
      };
      corpo.childNodes.forEach(x => andar(x, {}));
    });
    return runs;
  }
  async function pilha(p, itens, col1, col2, topo, alt) {
    const x0 = p.esquerda(col1), w = p.largura(col1, col2), gap = 1.5 * ESC;
    const fs = [];
    for (const it of itens) {
      const im = await imagemXL(it.src);
      let iw = 0, ih = 0;
      if (im) { const t = caberEm(im.w / im.h, 21.6 * ESC, 18 * ESC); iw = t.w * 0.97; ih = t.h * 0.97; }
      const th = txtW(it.rotulo, 7, w - 4);
      fs.push({im, iw, ih, th, rotulo: it.rotulo});
    }
    const total = fs.reduce((s, f) => s + f.ih * 1.05 + f.th + gap, 0);
    let y = topo + Math.max(2, (alt - total) / 2);
    for (const f of fs) {
      if (f.im) p.desenhos.push({tipo: 'img', img: f.im, x: x0 + (w - f.iw) / 2, y, w: f.iw, h: f.ih});
      y += f.ih * 1.05;
      p.desenhos.push({tipo: 'txt', texto: f.rotulo, sz: 7, b: true, x: x0 + 2, y, w: w - 4, h: f.th});
      y += f.th + gap;
    }
  }
  async function blocoEtapaXL(p, b, linha) {
    const H = ALT[b.tam] * MM;
    // o Excel usa linhas mais altas (Arial 8 = 11,25 pt por linha): a sublinha com texto quebrado ganha folga
    const nS = b.subs.length, hSub = b.hSubs.map(h => Math.max(h * MM, Math.max(1, Math.round((h - 1.6) / 3.25)) * 11.25 + 3)), principal = H - hSub.reduce((a, x) => a + x, 0), nPart = Math.ceil(principal / 400);
    for (let k = 0; k < nPart; k++) p.altura(linha + k, principal / nPart);
    for (let k = 0; k < nS; k++) p.altura(linha + nPart + k, hSub[k]);
    const ult = linha + nPart + nS - 1, ultP = linha + nPart - 1;
    const B = {borda: true, wrap: true, sz: FONTE};
    p.bloco(area(1, 1, linha, ult), b.num, {...B, b: true, cor: BRANCO, fundo: AZUL, h: 'center', sz: 6.5});
    p.bloco(area(2, 2, linha, ult), (b.etapa.nome || '').toUpperCase(), {...B, b: true, cor: BRANCO, fundo: AZUL, h: 'center', rot: 90, sz: 5});
    p.bloco(area(3, 5, linha, ult), '', B);
    p.bloco(area(6, 6, linha, ult), '', B);
    p.bloco(area(7, 13, linha, ultP), b.bullets.length ? runsDescricao(b.bullets) : '', {...B, h: 'justify'});
    b.subs.forEach((s, k) => {
      const rr = linha + nPart + k;
      p.bloco(area(7, 10, rr, rr), s.tipo, {...B, b: true, cor: BRANCO, fundo: AZUL, h: 'center'});
      p.bloco(area(11, 13, rr, rr), s.valor, {...B, h: 'justify'});
    });
    p.bloco(area(14, 14, linha, ult), '', B);
    p.bloco(area(15, 17, linha, ult), episPO(b.etapa).length ? '' : 'N/A', {...B, b: true, h: 'center'});
    const topo = p.topo(linha), alt = p.alturaLinhas(linha, ult);
    // fotos (mesmos tamanhos do PDF)
    const leg = doc.opcoes.legendaFotos ? 5 : 0;
    const reserva = b.grupos.reduce((s, g) => s + reservaLegenda(g), 0);
    const alts = alturasLinhas(b.grupos, ALT[b.tam] - 3 - reserva - vaoFotos() * (b.grupos.length - 1) - 0.5 * b.grupos.length);   // soma das alturas (07/11/2026)
    const linhasF = [];
    for (const [gi, g] of b.grupos.entries()) {
      const maxH = alts[gi];
      const fs = [];
      for (const f of g) {
        const lim = g.length === 2 ? FOTO_LADO : FOTO_MAX;
        const im = await imagemXL(f.img);
        const fa = f.aspecto ? f : Object.assign({}, f, {aspecto: im ? im.w / im.h : 4 / 3});   // usa a proporção real da foto
        const t = tamanhoFoto(fa, maxH, lim) || caberEm(fa.aspecto, lim.larg, Math.min(lim.alt, maxH));
        const legenda = leg ? `FOTO ${pad2(f.num)}` + (f.legenda ? ` – ${f.legenda}` : '') : '';
        // na impressão o Excel desenha a imagem um pouco maior que as linhas: reduz 3% e deixa folga para a legenda
        const w = t.w * ESC * 0.97, h = t.h * ESC * 0.97, lw = Math.max(w, 35 * ESC);
        fs.push({im, w, h, legenda, lw, lh: legenda ? txtW(legenda, FONTE_LEG, lw) : 0});
      }
      linhasF.push({fs, larg: fs.reduce((s, f) => s + f.w, 0), alt: Math.max(...fs.map(f => f.h + (f.lh ? f.lh + 2 : 0)))});
    }
    const soma = linhasF.reduce((s, x) => s + x.alt, 0);
    // espaço fixo entre as linhas de fotos (3 mm sem legenda; 1 mm com legenda), conjunto centralizado na altura da célula
    const vao = vaoFotos() * ESC + sobraDistribuir(b.etapa, linhasF.length, alt - soma - vaoFotos() * ESC * (linhasF.length - 1));   // DISTRIBUIR (07/11/2026)
    let y = topo + Math.max(0, (alt - soma - vao * (linhasF.length - 1)) / 2);
    const xF = p.esquerda(L(3)), wF = p.largura(L(3), L(5)), gapX = 1.4 * ESC;   // 1,4 mm entre as fotos lado a lado (07/11/2026)
    for (const lf of linhasF) {
      let x = xF + (wF - (lf.larg + gapX * (lf.fs.length - 1))) / 2;
      for (const f of lf.fs) {
        if (f.im) p.desenhos.push({tipo: 'img', img: f.im, x, y, w: f.w, h: f.h, moldura: true});
        if (f.legenda) p.desenhos.push({tipo: 'txt', texto: f.legenda, sz: FONTE_LEG, b: true, x: x + (f.w - f.lw) / 2, y: y + f.h + 2, w: f.lw, h: f.lh});
        x += f.w + gapX;
      }
      y += lf.alt + vao;
    }
    // pontos de atenção
    const lado = 13 * ESC, pts = b.etapa.pontos;
    let yp = topo + (alt - (pts.length * lado + (pts.length - 1) * 3 * ESC)) / 2;
    const xH = p.esquerda(L(6)), wH = p.largura(L(6));
    for (const k of pts) { if (icones[k]) p.desenhos.push({tipo: 'img', img: icones[k], x: xH + (wH - lado) / 2, y: yp, w: lado, h: lado}); yp += lado + 3 * ESC; }
    await pilha(p, b.perigos.map(x => ({src: imgPerigo(x), rotulo: x.rotulo})), L(14), L(14), topo, alt);
    await pilha(p, b.epis.map(x => ({src: imgEpi(x), rotulo: x.rotulo})), L(15), L(17), topo, alt);
    return ult + 1;
  }

  // ---------- aba do PO: cabeçalho (linhas 1 a 7, repete na impressão), etapas e anexos ----------
  const po = XL.Planilha(nomeFolha, colunasMM(G17), est);
  cabecalhoGerador(po);
  prepararMedidor();
  const blocos = blocosDoc();
  let nf = 0;
  blocos.forEach(b => { b.grupos = b.grupos.map(g => g.map(f => Object.assign({}, f, {num: ++nf}))); });
  let linha = 8, primeira = true;
  for (let i = 0; i < blocos.length; i++) {
    if (!primeira) po.quebras.push(linha);
    primeira = false;
    titulosColunas(po, linha); linha++;
    const b = blocos[i];
    linha = await blocoEtapaXL(po, b, linha);
    if (b.tam === 'meia' && blocos[i + 1] && blocos[i + 1].tam === 'meia' && !blocos[i + 1].novaPagina) { linha = await blocoEtapaXL(po, blocos[i + 1], linha); i++; }
  }
  // anexos: faixa com o título e as imagens numa célula ocupando o resto da página
  for (let i = 0; i < doc.anexos.length; i++) {
    const a = doc.anexos[i];
    if (linha > 8) po.quebras.push(linha);
    po.altura(linha, 5 * MM);
    po.bloco(area(1, 17, linha, linha), `ANEXO ${i + 1} – ${(a.titulo || '').toUpperCase()}`, {borda: true, sz: 9, b: true, cor: BRANCO, fundo: AZUL, h: 'center'});
    linha++;
    const altArea = (10 + ALT.inteira - 5) * MM, nPart = Math.ceil(altArea / 400);
    for (let k = 0; k < nPart; k++) po.altura(linha + k, altArea / nPart);
    po.bloco(area(1, 17, linha, linha + nPart - 1), '', {borda: true});
    const topo = po.topo(linha), xA = po.esquerda(L(1)), wA = po.largura(L(1), L(17));
    const k2 = Math.min(wA / (ANEXO_MAX.larg * MM), altArea / (ANEXO_MAX.alt * MM)) * MM * 0.88;   // mm do layout -> pt (folga: o Excel imprime imagens grandes um pouco maiores)
    const imgsA = [];
    for (const im of a.imagens) { const x = await imagemXL(im.img); imgsA.push(im.aspecto || !x ? im : Object.assign({}, im, {aspecto: x.w / x.h})); }
    const lins = layoutAnexo(imgsA);
    const altT = lins.reduce((s, ln) => s + Math.max(...ln.map(x => x.h)), 0) * k2 + 4 * k2 * (lins.length - 1);
    let y = topo + (altArea - altT) / 2;
    for (const ln of lins) {
      const larg = ln.reduce((s, x) => s + x.w, 0) * k2 + 4 * k2 * (ln.length - 1), hl = Math.max(...ln.map(x => x.h)) * k2;
      let x = xA + (wA - larg) / 2;
      for (const it of ln) {
        const im = await imagemXL(it.im.img);
        if (im) po.desenhos.push({tipo: 'img', img: im, x, y: y + (hl - it.h * k2) / 2, w: it.w * k2, h: it.h * k2});
        x += it.w * k2 + 4 * k2;
      }
      y += hl + 4 * k2;
    }
    linha += nPart;
  }
  po.zoom = 100;
  po.margens = {l: 0.16, r: 0.24, t: 0.16, b: 0.16};

  // ---------- lista de treinamento e revisões: mesmo layout das páginas do gerador (tamanho real) ----------

  // origem: nome da aba do PO; quando informado, os dados do cabeçalho viram fórmulas que espelham essa aba
  function cabecalhoGerador(p, origem) {
    const M = (c1, r1, valor) => origem ? {f: `'${origem.replace(/'/g, "''")}'!${L(c1)}${r1}&""`, v: valor} : valor;
    [8, 4, 4, 5, 2.5, 2.5, 5].forEach((h, k) => p.altura(k + 1, h * MM));
    const T = (c1, c2, r1, r2, v, e) => p.bloco(area(c1, c2, r1, r2), v, Object.assign({borda: true, wrap: true, h: 'center'}, e));
    T(1, 3, 1, 3, 'PROCEDIMENTO OPERACIONAL', {sz: 7, b: true, cor: BRANCO, fundo: AZ});
    T(4, 4, 1, 3, M(4, 1, c.codigo || 'CÓDIGO DO PO'), {sz: 16, b: true});
    T(5, 12, 1, 3, M(5, 1, (c.nome || 'NOME DO PO').toUpperCase()), {sz: 11, b: true});
    T(13, 13, 1, 1, 'Revisão n°:', {sz: 7.5, b: true, h: 'right'});
    T(14, 14, 1, 1, M(14, 1, r.rev), {sz: 7.5, b: true, texto: true});
    T(15, 15, 1, 1, 'Data:', {sz: 7.5, b: true});
    T(16, 17, 1, 1, M(16, 1, r.data || '__/__/____'), {sz: 7.5, b: true, texto: true});
    T(13, 13, 2, 2, 'Elaborado por:', {sz: 7.5, b: true, h: 'right'});
    T(14, 14, 2, 2, M(14, 2, c.elaborado), {sz: 7.5, h: 'left'});
    T(13, 13, 3, 3, 'Aprovado por:', {sz: 7.5, b: true, h: 'right'});
    T(14, 14, 3, 3, M(14, 3, c.aprovado), {sz: 7.5, h: 'left'});
    T(15, 15, 2, 3, 'Tempo de ciclo:', {sz: 7.5, b: true});
    T(16, 17, 2, 3, M(16, 2, c.tempoCiclo), {sz: 7.5, b: true});
    T(1, 3, 4, 7, "EPI'S BÁSICOS", {sz: 8, b: true, cor: BRANCO, fundo: AZ});
    T(4, 5, 4, 7, '', {});
    T(6, 6, 4, 7, 'LEGENDA', {sz: 8, b: true, cor: BRANCO, fundo: AZ});
    // legenda sem as linhas internas
    T(7, 7, 4, 5, '', {borda: 'lt'}); T(8, 8, 4, 5, 'Qualidade', {sz: 6.5, b: true, borda: 't'}); T(9, 9, 4, 5, '', {borda: 't'}); T(10, 12, 4, 5, 'Segurança', {sz: 6.5, b: true, borda: 'tr'});
    T(7, 7, 6, 7, '', {borda: 'lb'}); T(8, 8, 6, 7, 'Meio Ambiente', {sz: 6.5, b: true, borda: 'b'}); T(9, 9, 6, 7, '', {borda: 'b'}); T(10, 12, 6, 7, 'Energia', {sz: 6.5, b: true, borda: 'br'});
    T(13, 13, 4, 4, 'FERRAMENTAS', {sz: 7.5, b: true, cor: BRANCO, fundo: AZ}); T(14, 17, 4, 4, M(14, 4, l.ferramentas), {sz: 7.5, h: 'left'});
    T(13, 13, 5, 6, 'MATERIAIS', {sz: 7.5, b: true, cor: BRANCO, fundo: AZ}); T(14, 17, 5, 6, M(14, 5, l.materiais), {sz: 7.5, h: 'left'});
    T(13, 13, 7, 7, 'EQUIPAMENTOS', {sz: 7.5, b: true, cor: BRANCO, fundo: AZ}); T(14, 17, 7, 7, M(14, 7, l.equipamentos), {sz: 7.5, h: 'left'});
    // EPI's básicos (mesmo cálculo do gerador)
    const n = episBasicos.length;
    const t1 = Math.min(12, 60 / Math.max(n, 1) - 1.5, 13), t2 = Math.min(6.25, 60 / Math.ceil(n / 2) - 1.5);
    const duas = n > 4 && t2 > t1, lado = (duas ? t2 : t1) * MM;
    const linhasB = duas ? [episBasicos.slice(0, Math.ceil(n / 2)), episBasicos.slice(Math.ceil(n / 2))] : [episBasicos];
    const x0 = p.esquerda(L(4)), wA = p.largura(L(4), L(5)), y0 = p.topo(4), hA = p.alturaLinhas(4, 7);
    linhasB.forEach((lin, k) => {
      const passo = wA / lin.length, y = y0 + (hA - linhasB.length * lado - (linhasB.length - 1) * 1.5 * MM) / 2 + k * (lado + 1.5 * MM);
      lin.forEach((im, i) => p.desenhos.push({tipo: 'img', img: im, x: x0 + passo * i + (passo - lado) / 2, y, w: lado, h: lado}));
    });
    // ícones da legenda (5 mm, centralizados na célula)
    [[7, 4, 5, 'qualidade'], [9, 4, 5, 'seguranca'], [7, 6, 7, 'meio_ambiente'], [9, 6, 7, 'energia']].forEach(([col, r1, r2, k]) => {
      if (!icones[k]) return;
      const s = 5 * MM, x = p.esquerda(L(col)) + (p.largura(L(col)) - s) / 2, y = p.topo(r1) + (p.alturaLinhas(r1, r2) - s) / 2;
      p.desenhos.push({tipo: 'img', img: icones[k], x, y, w: s, h: s});
    });
  }

  const tr = XL.Planilha('Lista de Treinamento', colunasMM(G17), est);
  cabecalhoGerador(tr, po.nome);
  tr.altura(8, 10 * MM);
  const colsT = [[1, 5, 'NOME COMPLETO DO COLABORADOR'], [6, 6, 'NÚMERO REGISTRO'], [7, 12, 'TURNO'], [13, 13, 'DATA'], [14, 17, 'ASSINATURA']];
  colsT.forEach(([c1, c2, t]) => tr.bloco(area(c1, c2, 8, 8), t, {borda: true, wrap: true, h: 'center', sz: 8, b: true, cor: BRANCO, fundo: AZ}));
  for (let k = 0; k < LINHAS_TREINAMENTO; k++) {
    tr.altura(9 + k, (159 - 0.25) / LINHAS_TREINAMENTO * MM);
    colsT.forEach(([c1, c2]) => tr.bloco(area(c1, c2, 9 + k, 9 + k), '', {borda: true, sz: 8}));
  }
  tr.zoom = 100; tr.umaPagina = true;

  const GH = [704, 3544, 1172, 1379, 993, 2480, 568, 1913, 3507];
  const rv = XL.Planilha('Revisões', colunasMM(GH), est);
  rv.altura(1, 10 * MM); rv.altura(2, 5 * MM);
  rv.bloco('A1:I1', 'HISTÓRICO DE REVISÕES', {borda: true, sz: 12, b: true, cor: BRANCO, fundo: AZ, h: 'center'});
  // logotipo escolhido no cabeçalho, no lugar do VM3; o da Villares Metals saiu (05/11/2026)
  const vm3 = doc.cab.logo ? await imagemXL(doc.cab.logo) : null, tl = tamanhoLogoHist();
  if (vm3) rv.desenhos.push({tipo: 'img', img: vm3, x: 2 * MM, y: (10 - tl.h) / 2 * MM, w: tl.w * MM, h: tl.h * MM});
  [['A', 'A', 'Item'], ['B', 'B', 'Descrição da revisão'], ['C', 'D', 'Foto Alteração'], ['E', 'E', 'Data'], ['F', 'F', 'Nome Solicitante'], ['G', 'H', 'Responsável pela revisão'], ['I', 'I', 'Observações']]
    .forEach(([a, b, t]) => rv.bloco(a === b ? `${a}2` : `${a}2:${b}2`, t, {borda: true, sz: 8, b: true, cor: BRANCO, fundo: CINZA, h: 'center'}));
  const nHist = Math.max(LINHAS_HISTORICO, Math.ceil(doc.historico.length / LINHAS_HISTORICO) * LINHAS_HISTORICO);
  for (let i = 0; i < nHist; i++) {
    const h = doc.historico[i] || {}, ln = 3 + i, V = {borda: true, sz: 8, b: true, h: 'center', wrap: true, texto: true};
    rv.altura(ln, 26 * MM);
    rv.bloco(`A${ln}`, String(i), V); rv.bloco(`B${ln}`, h.descricao || '', V);
    rv.bloco(`C${ln}:D${ln}`, h.fotoImg ? '' : textoFotoHist(h), V);
    // foto da alteração: centralizada na célula, no tamanho do PO (máx. 4,4 × 2,5 cm)
    const fim = h.fotoImg && await imagemXL(h.fotoImg.img), tf = fim && tamanhoFoto(h.fotoImg, FOTO_HIST.alt, FOTO_HIST);
    if (tf) {
      const k = 0.97, w = tf.w * MM * k, hh = tf.h * MM * k, x0 = rv.largura('A', 'B'), y0 = 15 * MM + i * 26 * MM;
      rv.desenhos.push({tipo: 'img', img: fim, x: x0 + (rv.largura('C', 'D') - w) / 2, y: y0 + (26 * MM - hh) / 2, w, h: hh});
    }
    rv.bloco(`E${ln}`, fmtData(h.data), V); rv.bloco(`F${ln}`, h.solicitante || '', V); rv.bloco(`G${ln}:H${ln}`, h.responsavel || '', V);
    rv.bloco(`I${ln}`, h.obs || '', V);
    if (i > 0 && i % LINHAS_HISTORICO === 0) rv.quebras.push(ln);
  }
  rv.zoom = 100; rv.umaPagina = nHist <= LINHAS_HISTORICO;

  const q = n => `'${n.replace(/'/g, "''")}'`;
  return XL.montar([po, tr, rv], est, [
    {nome: '_xlnm.Print_Area', folha: 0, ref: `${q(po.nome)}!$A$1:$Q$${linha - 1}`},
    {nome: '_xlnm.Print_Titles', folha: 0, ref: `${q(po.nome)}!$1:$7`},
    {nome: '_xlnm.Print_Area', folha: 1, ref: `${q(tr.nome)}!$A$1:$Q$${8 + LINHAS_TREINAMENTO}`},
    {nome: '_xlnm.Print_Area', folha: 2, ref: `${q(rv.nome)}!$A$1:$I$${2 + nHist}`},
    {nome: '_xlnm.Print_Titles', folha: 2, ref: `${q(rv.nome)}!$1:$2`},  ], JSON.stringify(doc));
}

async function exportarExcel() {
  aviso('Montando o Excel...');
  await compactarAgora(false);
  let blob;
  try { blob = await gerarExcel(); }
  catch (e) { console.error(e); alert('Não consegui montar o Excel: ' + erroAmigavel(e)); return; }
  const nome = nomeArquivo() + '.xlsx';
  if (window.showSaveFilePicker) {
    try {
      const h = await window.showSaveFilePicker({suggestedName: nome, types: [{description: 'Planilha do Excel', accept: {'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx']}}]});
      const w = await h.createWritable(); await w.write(blob); await w.close();
      aviso('Excel salvo: ' + h.name); return;
    } catch (e) { if (e.name === 'AbortError') return; }
  }
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = nome;
  document.body.append(a); a.click(); a.remove();
  aviso('Excel salvo na pasta Downloads: ' + nome);
}
