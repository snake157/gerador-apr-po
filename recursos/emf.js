// Leitor de EMF (Enhanced Metafile) e de WMF (Windows Metafile, o formato antigo) para os importadores de Word e Excel:
// desenha a imagem num canvas e devolve um PNG.
// O navegador não abre EMF; ele é comum quando um check-list do Excel é colado no Word como imagem.
// Cobre os registros GDI que o Excel/Word usam (linhas, retângulos, polígonos, preenchimentos, textos, bitmaps).
// Os EMF "duplos" (EMF+ com cópia GDI) também funcionam, porque a parte GDI vem junto. Se não der, devolve null.
'use strict';

const EMF = (() => {
  const corDe = c => `rgb(${c & 255},${(c >> 8) & 255},${(c >> 16) & 255})`;
  const STOCK = {
    0x80000000: {tipo: 'brush', estilo: 0, cor: 0xFFFFFF}, 0x80000001: {tipo: 'brush', estilo: 0, cor: 0xC0C0C0}, 0x80000002: {tipo: 'brush', estilo: 0, cor: 0x808080},
    0x80000003: {tipo: 'brush', estilo: 0, cor: 0x404040}, 0x80000004: {tipo: 'brush', estilo: 0, cor: 0}, 0x80000005: {tipo: 'brush', estilo: 1, cor: 0},
    0x80000006: {tipo: 'pen', estilo: 0, largura: 0, cor: 0xFFFFFF}, 0x80000007: {tipo: 'pen', estilo: 0, largura: 0, cor: 0}, 0x80000008: {tipo: 'pen', estilo: 5, largura: 0, cor: 0},
    0x8000000A: {tipo: 'font', altura: -13, peso: 400, italico: 0, nome: 'Arial', esc: 0}, 0x8000000D: {tipo: 'font', altura: -13, peso: 400, italico: 0, nome: 'Arial', esc: 0},
    0x80000011: {tipo: 'font', altura: -11, peso: 400, italico: 0, nome: 'Arial', esc: 0},
  };
  const mult = (a, b) => [a[0] * b[0] + a[1] * b[2], a[0] * b[1] + a[1] * b[3], a[2] * b[0] + a[3] * b[2], a[2] * b[1] + a[3] * b[3], a[4] * b[0] + a[5] * b[2] + b[4], a[4] * b[1] + a[5] * b[3] + b[5]];

  // DIB (BITMAPINFO + bits) -> canvas
  function dib(dv, offBmi, offBits) {
    const w = dv.getInt32(offBmi + 4, true), hRaw = dv.getInt32(offBmi + 8, true), bpp = dv.getUint16(offBmi + 14, true), comp = dv.getUint32(offBmi + 16, true);
    const hdr = dv.getUint32(offBmi, true), h = Math.abs(hRaw);
    if (!w || !h || w > 8000 || h > 8000 || (comp !== 0 && comp !== 3)) return null;
    let nCores = dv.getUint32(offBmi + 32, true);
    if (!nCores && bpp <= 8) nCores = 1 << bpp;
    const pal = [];
    for (let i = 0; i < nCores; i++) { const p = offBmi + hdr + (comp === 3 ? 12 : 0) + i * 4; pal.push([dv.getUint8(p + 2), dv.getUint8(p + 1), dv.getUint8(p)]); }
    const c = document.createElement('canvas'); c.width = w; c.height = h;
    const g = c.getContext('2d'), img = g.createImageData(w, h);
    const linha = Math.floor((w * bpp + 31) / 32) * 4;
    for (let y = 0; y < h; y++) {
      const yy = hRaw > 0 ? h - 1 - y : y, base = offBits + y * linha;
      for (let x = 0; x < w; x++) {
        let r, gg, b;
        if (bpp === 24 || bpp === 32) { const p = base + x * (bpp / 8); b = dv.getUint8(p); gg = dv.getUint8(p + 1); r = dv.getUint8(p + 2); }
        else if (bpp === 16) { const v = dv.getUint16(base + x * 2, true); r = ((v >> 10) & 31) << 3; gg = ((v >> 5) & 31) << 3; b = (v & 31) << 3; }
        else { const bitPos = x * bpp, byte = dv.getUint8(base + (bitPos >> 3)), idx = (byte >> (8 - bpp - (bitPos & 7))) & ((1 << bpp) - 1); [r, gg, b] = pal[idx] || [0, 0, 0]; }
        const o = (yy * w + x) * 4; img.data[o] = r; img.data[o + 1] = gg; img.data[o + 2] = b; img.data[o + 3] = 255;
      }
    }
    g.putImageData(img, 0, 0);
    return c;
  }

  // buf: ArrayBuffer do .emf; larguraPx: largura do PNG
  function paraPNG(buf, larguraPx = 1600) {
    const dv = new DataView(buf);
    if (dv.byteLength < 88 || dv.getUint32(0, true) !== 1 || dv.getUint32(40, true) !== 0x464D4520) return null;
    const bx0 = dv.getInt32(8, true), by0 = dv.getInt32(12, true), bx1 = dv.getInt32(16, true), by1 = dv.getInt32(20, true);
    const bw = bx1 - bx0 + 1, bh = by1 - by0 + 1;
    if (bw < 2 || bh < 2) return null;
    const k = Math.min(larguraPx / bw, 6000 / bh);
    const cv = document.createElement('canvas'); cv.width = Math.max(1, Math.round(bw * k)); cv.height = Math.max(1, Math.round(bh * k));
    const g = cv.getContext('2d');
    g.fillStyle = '#fff'; g.fillRect(0, 0, cv.width, cv.height);
    const base = [k, 0, 0, k, -bx0 * k, -by0 * k];          // dispositivo -> canvas
    const novoDC = () => ({modo: 1, wox: 0, woy: 0, wex: 1, wey: 1, vox: 0, voy: 0, vex: 1, vey: 1, mundo: [1, 0, 0, 1, 0, 0],
      pen: STOCK[0x80000007], brush: STOCK[0x80000000], font: STOCK[0x8000000D], corTexto: 0, corFundo: 0xFFFFFF, fundoOpaco: false, alinh: 0, x: 0, y: 0, enchimento: 1, clip: null});
    let dc = novoDC();
    const pilha = [], objs = {};
    // lógico -> canvas
    const matriz = () => {
      let m = dc.mundo;
      if (dc.modo === 7 || dc.modo === 8) { const sx = dc.vex / dc.wex, sy = dc.vey / dc.wey; m = mult(m, [sx, 0, 0, sy, dc.vox - dc.wox * sx, dc.voy - dc.woy * sy]); }
      else m = mult(m, [1, 0, 0, 1, dc.vox - dc.wox, dc.voy - dc.woy]);
      return mult(m, base);
    };
    const escala = m => Math.sqrt(Math.abs(m[0] * m[3] - m[1] * m[2])) || 1;
    let caminho = null, emCaminho = false;
    const iniciar = () => { if (!emCaminho) { g.beginPath(); } };
    const aplicarClip = () => {
      g.restore(); g.save();
      if (dc.clip) { g.setTransform(1, 0, 0, 1, 0, 0); g.beginPath(); g.rect(dc.clip[0], dc.clip[1], dc.clip[2] - dc.clip[0], dc.clip[3] - dc.clip[1]); g.clip(); }
    };
    g.save();
    const tracar = () => {
      const p = dc.pen;
      if (!p || p.estilo === 5) return;
      const m = matriz();
      g.setTransform(1, 0, 0, 1, 0, 0);
      g.lineWidth = Math.max(0.8, (p.largura || 0) * escala(m)); g.strokeStyle = corDe(p.cor);
      g.setLineDash(p.estilo === 1 ? [6, 3] : p.estilo === 2 ? [1.5, 2] : p.estilo === 3 ? [6, 2, 1.5, 2] : []);
      g.stroke(caminho);
    };
    const encher = () => {
      const b = dc.brush;
      if (!b || b.estilo === 1) return;
      g.setTransform(1, 0, 0, 1, 0, 0); g.fillStyle = corDe(b.cor);
      g.fill(caminho, dc.enchimento === 1 ? 'evenodd' : 'nonzero');
    };
    // caminho em coordenadas de canvas
    const pt = (m, x, y) => [x * m[0] + y * m[2] + m[4], x * m[1] + y * m[3] + m[5]];
    const novoCaminho = () => { if (!emCaminho) caminho = new Path2D(); };
    const fimForma = (fill, stroke) => { if (emCaminho) return; if (fill) encher(); if (stroke) tracar(); };
    const poligono = (pts, fechar, fill) => {
      novoCaminho(); const m = matriz();
      pts.forEach(([x, y], i) => { const [a, b] = pt(m, x, y); if (i === 0) caminho.moveTo(a, b); else caminho.lineTo(a, b); });
      if (fechar) caminho.closePath();
      fimForma(fill, true);
    };
    const ler16 = (p, n) => { const r = []; for (let i = 0; i < n; i++) r.push([dv.getInt16(p + i * 4, true), dv.getInt16(p + i * 4 + 2, true)]); return r; };
    const ler32 = (p, n) => { const r = []; for (let i = 0; i < n; i++) r.push([dv.getInt32(p + i * 8, true), dv.getInt32(p + i * 8 + 4, true)]); return r; };
    let desenhou = 0, p = 0, voltas = 0;
    try {
      while (p + 8 <= dv.byteLength && voltas++ < 500000) {
        const t = dv.getUint32(p, true), sz = dv.getUint32(p + 4, true);
        if (sz < 8 || p + sz > dv.byteLength) break;
        const q = p + 8;
        switch (t) {
          case 14: p = dv.byteLength; continue;                                  // EOF
          case 17: dc.modo = dv.getUint32(q, true); break;                        // SETMAPMODE
          case 9: dc.wex = dv.getInt32(q, true) || 1; dc.wey = dv.getInt32(q + 4, true) || 1; break;
          case 10: dc.wox = dv.getInt32(q, true); dc.woy = dv.getInt32(q + 4, true); break;
          case 11: dc.vex = dv.getInt32(q, true) || 1; dc.vey = dv.getInt32(q + 4, true) || 1; break;
          case 12: dc.vox = dv.getInt32(q, true); dc.voy = dv.getInt32(q + 4, true); break;
          case 35: dc.mundo = [0, 4, 8, 12, 16, 20].map(o => dv.getFloat32(q + o, true)); break;   // SETWORLDTRANSFORM
          case 36: {                                                                              // MODIFYWORLDTRANSFORM
            const x = [0, 4, 8, 12, 16, 20].map(o => dv.getFloat32(q + o, true)), modo = dv.getUint32(q + 24, true);
            dc.mundo = modo === 1 ? [1, 0, 0, 1, 0, 0] : modo === 2 ? mult(x, dc.mundo) : modo === 3 ? mult(dc.mundo, x) : x; break;
          }
          case 33: pilha.push(Object.assign({}, dc)); break;                     // SAVEDC
          case 34: { const n = dv.getInt32(q, true); let r = null; for (let i = 0; i < Math.abs(n || 1) && pilha.length; i++) r = pilha.pop(); if (r) { dc = r; aplicarClip(); } break; }
          case 24: dc.corTexto = dv.getUint32(q, true); break;
          case 25: dc.corFundo = dv.getUint32(q, true); break;
          case 18: dc.fundoOpaco = dv.getUint32(q, true) === 2; break;
          case 22: dc.alinh = dv.getUint32(q, true); break;
          case 19: dc.enchimento = dv.getUint32(q, true); break;
          case 38: objs[dv.getUint32(q, true)] = {tipo: 'pen', estilo: dv.getUint32(q + 4, true) & 15, largura: dv.getInt32(q + 8, true), cor: dv.getUint32(q + 16, true)}; break;
          case 95: objs[dv.getUint32(q, true)] = {tipo: 'pen', estilo: dv.getUint32(q + 20, true) & 15, largura: dv.getUint32(q + 24, true), cor: dv.getUint32(q + 32, true)}; break;
          case 39: objs[dv.getUint32(q, true)] = {tipo: 'brush', estilo: dv.getUint32(q + 4, true), cor: dv.getUint32(q + 8, true)}; break;
          case 82: {                                                                              // EXTCREATEFONTINDIRECTW
            let nome = ''; for (let i = 0; i < 32; i++) { const ch = dv.getUint16(q + 4 + 28 + i * 2, true); if (!ch) break; nome += String.fromCharCode(ch); }
            objs[dv.getUint32(q, true)] = {tipo: 'font', altura: dv.getInt32(q + 4, true), esc: dv.getInt32(q + 12, true), peso: dv.getInt32(q + 20, true), italico: dv.getUint8(q + 24), nome: nome || 'Arial'}; break;
          }
          case 37: { const h = dv.getUint32(q, true), o = h >= 0x80000000 ? STOCK[h] : objs[h]; if (o) dc[o.tipo] = o; break; }
          case 40: delete objs[dv.getUint32(q, true)]; break;
          case 27: dc.x = dv.getInt32(q, true); dc.y = dv.getInt32(q + 4, true); if (emCaminho) { const [a, b] = pt(matriz(), dc.x, dc.y); caminho.moveTo(a, b); } break;
          case 54: {                                                                              // LINETO
            const x = dv.getInt32(q, true), y = dv.getInt32(q + 4, true), m = matriz();
            if (emCaminho) { const [a, b] = pt(m, x, y); caminho.lineTo(a, b); }
            else { caminho = new Path2D(); const [a, b] = pt(m, dc.x, dc.y), [c2, d2] = pt(m, x, y); caminho.moveTo(a, b); caminho.lineTo(c2, d2); tracar(); desenhou++; }
            dc.x = x; dc.y = y; break;
          }
          case 43: case 44: {                                                                     // RECTANGLE / ROUNDRECT
            const [l, tp, r, b] = [0, 4, 8, 12].map(o => dv.getInt32(q + o, true));
            if (t === 44) {
              const rw = dv.getInt32(q + 16, true) / 2, rh = dv.getInt32(q + 20, true) / 2, m = matriz(); novoCaminho();
              const P = (x, y) => pt(m, x, y), seg = [[l + rw, tp], [r - rw, tp]];
              caminho.moveTo(...P(...seg[0])); caminho.lineTo(...P(r - rw, tp)); caminho.quadraticCurveTo(...P(r, tp), ...P(r, tp + rh)); caminho.lineTo(...P(r, b - rh));
              caminho.quadraticCurveTo(...P(r, b), ...P(r - rw, b)); caminho.lineTo(...P(l + rw, b)); caminho.quadraticCurveTo(...P(l, b), ...P(l, b - rh));
              caminho.lineTo(...P(l, tp + rh)); caminho.quadraticCurveTo(...P(l, tp), ...P(l + rw, tp)); caminho.closePath(); fimForma(true, true);
            } else poligono([[l, tp], [r, tp], [r, b], [l, b]], true, true);
            desenhou++; break;
          }
          case 42: {                                                                              // ELLIPSE
            const [l, tp, r, b] = [0, 4, 8, 12].map(o => dv.getInt32(q + o, true)), m = matriz(); novoCaminho();
            const [cx, cy] = pt(m, (l + r) / 2, (tp + b) / 2), [ex, ey] = pt(m, r, b), [sx, sy] = pt(m, l, tp);
            caminho.ellipse(cx, cy, Math.abs(ex - sx) / 2, Math.abs(ey - sy) / 2, 0, 0, Math.PI * 2); fimForma(true, true); desenhou++; break;
          }
          case 86: case 87: case 3: case 4: case 89: case 6: {                                    // POLYGON / POLYLINE (16 e 32 bits), POLYLINETO
            const n = dv.getUint32(q + 16, true), pts = (t === 3 || t === 4 || t === 6) ? ler32(q + 20, n) : ler16(q + 20, n);
            if (t === 89 || t === 6) { pts.unshift([dc.x, dc.y]); dc.x = pts[pts.length - 1][0]; dc.y = pts[pts.length - 1][1]; }
            if (emCaminho && (t === 89 || t === 6)) { const m = matriz(); pts.slice(1).forEach(([x, y]) => caminho.lineTo(...pt(m, x, y))); }
            else poligono(pts, t === 86 || t === 3, t === 86 || t === 3);
            desenhou++; break;
          }
          case 91: case 92: case 7: case 8: {                                                     // POLYPOLYLINE / POLYPOLYGON
            const nPol = dv.getUint32(q + 16, true), cnt = dv.getUint32(q + 20, true), b16 = t === 91 || t === 92;
            let off = q + 24 + nPol * 4; const m = matriz(); novoCaminho();
            for (let i = 0; i < nPol; i++) {
              const n = dv.getUint32(q + 24 + i * 4, true), pts = b16 ? ler16(off, n) : ler32(off, n); off += n * (b16 ? 4 : 8);
              pts.forEach(([x, y], j) => { const [a, b] = pt(m, x, y); if (j === 0) caminho.moveTo(a, b); else caminho.lineTo(a, b); });
              if (t === 92 || t === 8) caminho.closePath();
            }
            fimForma(t === 92 || t === 8, true); desenhou += cnt ? 1 : 0; break;
          }
          case 85: case 88: case 2: case 5: {                                                     // POLYBEZIERTO16 / POLYBEZIER16
            const n = dv.getUint32(q + 16, true), pts = (t === 2 || t === 5) ? ler32(q + 20, n) : ler16(q + 20, n), m = matriz();
            novoCaminho();
            let ini = 0;
            if (t === 88 || t === 2) { caminho.moveTo(...pt(m, ...pts[0])); ini = 1; } else if (!emCaminho) caminho.moveTo(...pt(m, dc.x, dc.y));
            for (let i = ini; i + 2 < pts.length; i += 3) caminho.bezierCurveTo(...pt(m, ...pts[i]), ...pt(m, ...pts[i + 1]), ...pt(m, ...pts[i + 2]));
            if (pts.length) { dc.x = pts[pts.length - 1][0]; dc.y = pts[pts.length - 1][1]; }
            fimForma(false, true); desenhou++; break;
          }
          case 59: emCaminho = true; caminho = new Path2D(); break;               // BEGINPATH
          case 60: emCaminho = false; break;                                      // ENDPATH
          case 61: if (caminho) caminho.closePath(); break;                       // CLOSEFIGURE
          case 62: if (caminho) { encher(); desenhou++; } break;                  // FILLPATH
          case 63: if (caminho) { encher(); tracar(); desenhou++; } break;        // STROKEANDFILLPATH
          case 64: if (caminho) { tracar(); desenhou++; } break;                  // STROKEPATH
          case 30: {                                                              // INTERSECTCLIPRECT
            const [l, tp, r, b] = [0, 4, 8, 12].map(o => dv.getInt32(q + o, true)), m = matriz();
            const [x0, y0] = pt(m, l, tp), [x1, y1] = pt(m, r, b);
            let c = [Math.min(x0, x1), Math.min(y0, y1), Math.max(x0, x1), Math.max(y0, y1)];
            if (dc.clip) c = [Math.max(c[0], dc.clip[0]), Math.max(c[1], dc.clip[1]), Math.min(c[2], dc.clip[2]), Math.min(c[3], dc.clip[3])];
            dc.clip = c; aplicarClip(); break;
          }
          case 75: {                                                              // EXTSELECTCLIPRGN: região em unidades do dispositivo
            const cb = dv.getUint32(q, true), modo = dv.getUint32(q + 4, true);
            if (!cb) { dc.clip = null; aplicarClip(); break; }
            const r0 = q + 8, n = dv.getUint32(r0 + 8, true);
            let c = null;
            for (let i = 0; i < n; i++) {
              const o = r0 + 32 + i * 16, [l, tp, r, b] = [0, 4, 8, 12].map(k2 => dv.getInt32(o + k2, true));
              const [x0, y0] = pt(base, l, tp), [x1, y1] = pt(base, r, b);
              c = c ? [Math.min(c[0], x0), Math.min(c[1], y0), Math.max(c[2], x1), Math.max(c[3], y1)] : [x0, y0, x1, y1];
            }
            if (!c) break;
            if (modo === 1 && dc.clip) c = [Math.max(c[0], dc.clip[0]), Math.max(c[1], dc.clip[1]), Math.min(c[2], dc.clip[2]), Math.min(c[3], dc.clip[3])];
            else if (modo === 2 && dc.clip) c = [Math.min(c[0], dc.clip[0]), Math.min(c[1], dc.clip[1]), Math.max(c[2], dc.clip[2]), Math.max(c[3], dc.clip[3])];
            else if (modo === 4) break;                                           // diferença: ignora
            dc.clip = c; aplicarClip(); break;
          }
          case 76: case 81: case 77: {                                            // BITBLT / STRETCHDIBITS / STRETCHBLT
            const m = matriz();
            let xd, yd, wd, hd, rop, offBmi, cbBmi, offBits;
            if (t === 81) { xd = dv.getInt32(q + 16, true); yd = dv.getInt32(q + 20, true); offBmi = dv.getUint32(q + 40, true); cbBmi = dv.getUint32(q + 44, true); offBits = dv.getUint32(q + 48, true); rop = dv.getUint32(q + 60, true); wd = dv.getInt32(q + 64, true); hd = dv.getInt32(q + 68, true); }
            else { xd = dv.getInt32(q + 16, true); yd = dv.getInt32(q + 20, true); wd = dv.getInt32(q + 24, true); hd = dv.getInt32(q + 28, true); rop = dv.getUint32(q + 32, true); offBmi = dv.getUint32(q + 76, true); cbBmi = dv.getUint32(q + 80, true); offBits = dv.getUint32(q + 84, true);
              if (t === 77) { wd = dv.getInt32(q + 24, true); hd = dv.getInt32(q + 28, true); } }
            const [ax, ay] = pt(m, xd, yd), [bx, by] = pt(m, xd + wd, yd + hd);
            g.setTransform(1, 0, 0, 1, 0, 0);
            if (cbBmi && offBmi) { const c = dib(dv, p + offBmi, p + offBits); if (c) { g.drawImage(c, Math.min(ax, bx), Math.min(ay, by), Math.abs(bx - ax), Math.abs(by - ay)); desenhou++; } }
            else {
              // só PATCOPY/WHITENESS/BLACKNESS pintam; os demais (ex.: 0xAA0029, "não mexe") não
              const cor = rop === 0x00FF0062 ? 'rgb(255,255,255)' : rop === 0x00000042 ? 'rgb(0,0,0)' : (rop === 0x00F00021 && dc.brush && dc.brush.estilo !== 1 ? corDe(dc.brush.cor) : null);
              if (cor) { g.fillStyle = cor; g.fillRect(Math.min(ax, bx), Math.min(ay, by), Math.abs(bx - ax), Math.abs(by - ay)); desenhou++; }
            }
            break;
          }
          case 84: case 83: {                                                     // EXTTEXTOUTW / EXTTEXTOUTA
            const x = dv.getInt32(q + 28, true), y = dv.getInt32(q + 32, true), n = dv.getUint32(q + 36, true), offStr = dv.getUint32(q + 40, true), opc = dv.getUint32(q + 44, true), offDx = dv.getUint32(q + 64, true);
            if (!n) break;
            let s = '';
            for (let i = 0; i < n; i++) s += t === 84 ? String.fromCharCode(dv.getUint16(p + offStr + i * 2, true)) : String.fromCharCode(dv.getUint8(p + offStr + i));
            const f = dc.font || STOCK[0x8000000D], m = matriz();
            const alt = Math.abs(f.altura) || 12;
            const ux = (dc.alinh & 1) ? dc.x : x, uy = (dc.alinh & 1) ? dc.y : y;
            g.setTransform(m[0], m[1], m[2], m[3], m[4], m[5]);
            const sinalY = Math.sign(m[3]) || 1;
            g.font = `${f.italico ? 'italic ' : ''}${f.peso >= 600 ? 'bold ' : ''}${alt * (f.altura > 0 ? 0.8 : 1)}px '${f.nome}', Arial, sans-serif`;
            g.fillStyle = corDe(dc.corTexto);
            g.textBaseline = (dc.alinh & 24) === 24 ? 'alphabetic' : (dc.alinh & 8) ? 'bottom' : 'top';
            const ha = dc.alinh & 6;
            g.save();
            g.translate(ux, uy); if (f.esc) g.rotate(-f.esc / 10 * Math.PI / 180 * sinalY); if (sinalY < 0) g.scale(1, -1);
            if (offDx && !(ha === 6 || ha === 2)) {
              let cx = 0;
              for (let i = 0; i < n; i++) { g.textAlign = 'left'; g.fillText(s[i], cx, 0); cx += dv.getInt32(p + offDx + i * 4 * (opc & 0x2000 ? 2 : 1), true); }
            } else { g.textAlign = ha === 6 ? 'center' : ha === 2 ? 'right' : 'left'; g.fillText(s, 0, 0); }
            g.restore();
            desenhou++; break;
          }
          default: break;
        }
        p += sz;
      }
    } catch (e) { console.warn('EMF:', e); }
    g.restore();
    if (!desenhou) return null;
    return cv.toDataURL('image/png');
  }

  // ---------------- WMF (Windows Metafile, formato de 16 bits, mais antigo que o EMF) ----------------
  // Mesmo uso do paraPNG: desenha num canvas e devolve um PNG (ou null). Muitos WMF do Word trazem o EMF
  // inteiro dentro (comentários "WMFC"): nesse caso o EMF é que é desenhado, com o leitor acima.
  function wmfParaPNG(buf, larguraPx = 1600) {
    const dv = new DataView(buf);
    let p = 0, caixa = null;
    if (dv.byteLength >= 22 && dv.getUint32(0, true) === 0x9AC6CDD7) {        // cabeçalho "placeable"
      const l = dv.getInt16(6, true), t = dv.getInt16(8, true), r = dv.getInt16(10, true), b = dv.getInt16(12, true);
      caixa = [l, t, r, b]; p = 22;
    }
    if (p + 18 > dv.byteLength) return null;
    const tipo = dv.getUint16(p, true), tamCab = dv.getUint16(p + 2, true);
    if ((tipo !== 1 && tipo !== 2) || tamCab !== 9) return null;
    const ini = p + tamCab * 2;
    const registros = [];
    for (let q = ini, voltas = 0; q + 6 <= dv.byteLength && voltas < 500000; voltas++) {
      const sz = dv.getUint32(q, true) * 2, fn = dv.getUint16(q + 4, true);
      if (sz < 6 || q + sz > dv.byteLength) break;
      registros.push([q, sz, fn]);
      if (fn === 0) break;
      q += sz;
    }
    // EMF guardado dentro do WMF (META_ESCAPE / MFCOMMENT "WMFC"): junta os pedaços e usa o leitor de EMF
    try {
      let emf = null, pos = 0;
      for (const [q, sz, fn] of registros) {
        if (fn !== 0x0626 || sz < 6 + 4 + 34 || dv.getUint16(q + 6, true) !== 0x000F || dv.getUint32(q + 10, true) !== 0x43464D57) continue;
        const d = q + 10, cur = dv.getUint32(d + 22, true), total = dv.getUint32(d + 30, true);
        if (!emf) { if (!total || total > 200e6) break; emf = new Uint8Array(total); }
        if (pos + cur > emf.length || d + 34 + cur > q + sz) { emf = null; break; }
        emf.set(new Uint8Array(buf, d + 34, cur), pos); pos += cur;
      }
      if (emf && pos === emf.length) { const png = paraPNG(emf.buffer, larguraPx); if (png) return png; }
    } catch (e) { /* desenha pelos registros WMF */ }

    const i16 = o => dv.getInt16(o, true), u16 = o => dv.getUint16(o, true), cor = o => dv.getUint32(o, true) & 0xFFFFFF;
    // janela lógica (SETWINDOWORG / SETWINDOWEXT) e tamanho do desenho
    let wox = 0, woy = 0, wex = 0, wey = 0;
    for (const [q, , fn] of registros) {
      if (fn === 0x020B && !wex) { woy = i16(q + 6); wox = i16(q + 8); }
      if (fn === 0x020C && !wex) { wey = i16(q + 6); wex = i16(q + 8); }
    }
    if (!wex || !wey) { if (!caixa) return null; wox = caixa[0]; woy = caixa[1]; wex = caixa[2] - caixa[0]; wey = caixa[3] - caixa[1]; }
    if (!wex || !wey) return null;
    const prop = caixa && caixa[2] !== caixa[0] ? Math.abs((caixa[3] - caixa[1]) / (caixa[2] - caixa[0])) : Math.abs(wey / wex);
    const W = larguraPx, H = Math.max(1, Math.min(6000, Math.round(larguraPx * prop)));
    const cv = document.createElement('canvas'); cv.width = W; cv.height = H;
    const g = cv.getContext('2d');
    g.fillStyle = '#fff'; g.fillRect(0, 0, W, H);

    const novoDC = () => ({wox, woy, wex, wey, pen: {estilo: 0, largura: 0, cor: 0}, brush: {estilo: 0, cor: 0xFFFFFF}, font: {altura: -12, peso: 400, italico: 0, nome: 'Arial', esc: 0},
      corTexto: 0, corFundo: 0xFFFFFF, fundoOpaco: true, alinh: 0, x: 0, y: 0, enchimento: 1, clip: null});
    let dc = novoDC();
    const pilha = [], objs = [];
    const novoObj = o => { let i = 0; while (objs[i]) i++; objs[i] = o; };
    const X = x => (x - dc.wox) / dc.wex * W, Y = y => (y - dc.woy) / dc.wey * H;
    const kx = () => Math.abs(W / dc.wex), ky = () => Math.abs(H / dc.wey);
    const aplicarClip = () => {
      g.restore(); g.save();
      if (dc.clip) { g.beginPath(); g.rect(dc.clip[0], dc.clip[1], dc.clip[2] - dc.clip[0], dc.clip[3] - dc.clip[1]); g.clip(); }
    };
    const pincel = () => {
      const b = dc.brush;
      if (!b || b.estilo === 1) return null;
      if (b.padrao) return g.createPattern(b.padrao, 'repeat');
      return corDe(b.cor);
    };
    const formar = (caminho, encher, tracar = true) => {
      if (encher) { const f = pincel(); if (f) { g.fillStyle = f; g.fill(caminho, dc.enchimento === 1 ? 'evenodd' : 'nonzero'); } }
      const pn = dc.pen;
      if (tracar && pn && pn.estilo !== 5) {
        g.lineWidth = Math.max(0.8, (pn.largura || 0) * kx()); g.strokeStyle = corDe(pn.cor);
        g.setLineDash(pn.estilo === 1 ? [6, 3] : pn.estilo === 2 ? [1.5, 2] : pn.estilo === 3 ? [6, 2, 1.5, 2] : []);
        g.stroke(caminho);
      }
    };
    const pontos = (o, n) => { const r = []; for (let i = 0; i < n; i++) r.push([X(i16(o + i * 4)), Y(i16(o + i * 4 + 2))]); return r; };
    const caminhoDe = (pts, fechar) => { const c = new Path2D(); pts.forEach(([a, b], i) => i ? c.lineTo(a, b) : c.moveTo(a, b)); if (fechar) c.closePath(); return c; };
    // retângulo nos parâmetros (bottom, right, top, left) a partir de o
    const ret = o => { const b = Y(i16(o)), r = X(i16(o + 2)), t = Y(i16(o + 4)), l = X(i16(o + 6)); return [Math.min(l, r), Math.min(t, b), Math.max(l, r), Math.max(t, b)]; };
    // DIB (BITMAPINFOHEADER + paleta + bits) começando em o
    const dibEm = o => {
      const hdr = dv.getUint32(o, true);
      if (hdr === 12) return null;                                              // BITMAPCOREHEADER: raro, fica de fora
      const bpp = u16(o + 14), comp = dv.getUint32(o + 16, true);
      let n = dv.getUint32(o + 32, true); if (!n && bpp <= 8) n = 1 << bpp;
      return dib(dv, o, o + hdr + (comp === 3 ? 12 : 0) + n * 4);
    };
    const texto = (s, x, y) => {
      if (!s) return;
      const f = dc.font, alt = Math.max(1, Math.abs(f.altura || 12) * (f.altura > 0 ? 0.8 : 1) * ky());
      const usaCP = dc.alinh & 1, px = usaCP ? X(dc.x) : X(x), py = usaCP ? Y(dc.y) : Y(y);
      g.save();
      g.font = `${f.italico ? 'italic ' : ''}${f.peso >= 600 ? 'bold ' : ''}${alt}px '${f.nome}', Arial, sans-serif`;
      g.fillStyle = corDe(dc.corTexto);
      g.textBaseline = (dc.alinh & 24) === 24 ? 'alphabetic' : (dc.alinh & 8) ? 'bottom' : 'top';
      const ha = dc.alinh & 6;
      g.textAlign = ha === 6 ? 'center' : ha === 2 ? 'right' : 'left';
      g.translate(px, py); if (f.esc) g.rotate(-f.esc / 10 * Math.PI / 180);
      g.fillText(s, 0, 0);
      g.restore();
    };
    const dec = (() => { try { return new TextDecoder('windows-1252'); } catch (e) { return null; } })();
    const lerTexto = (o, n) => { const b = new Uint8Array(buf, o, n); return dec ? dec.decode(b) : String.fromCharCode(...b); };

    let desenhou = 0;
    g.save();
    try {
      for (const [q, sz, fn] of registros) {
        const a = q + 6;                                                        // parâmetros
        switch (fn) {
          case 0x0000: break;                                                   // EOF
          case 0x020B: dc.woy = i16(a); dc.wox = i16(a + 2); break;             // SETWINDOWORG (y, x)
          case 0x020C: dc.wey = i16(a) || dc.wey; dc.wex = i16(a + 2) || dc.wex; break;   // SETWINDOWEXT (y, x)
          case 0x020F: dc.woy += i16(a); dc.wox += i16(a + 2); break;           // OFFSETWINDOWORG
          case 0x0410: { const yD = i16(a) || 1, yN = i16(a + 2), xD = i16(a + 4) || 1, xN = i16(a + 6); dc.wey = dc.wey * yN / yD || dc.wey; dc.wex = dc.wex * xN / xD || dc.wex; break; }   // SCALEWINDOWEXT
          case 0x001E: pilha.push(Object.assign({}, dc)); break;                // SAVEDC
          case 0x0127: { const n = i16(a); let r = null; for (let i = 0; i < Math.abs(n || 1) && pilha.length; i++) r = pilha.pop(); if (r) { dc = r; aplicarClip(); } break; }   // RESTOREDC
          case 0x0209: dc.corTexto = cor(a); break;                             // SETTEXTCOLOR
          case 0x0201: dc.corFundo = cor(a); break;                             // SETBKCOLOR
          case 0x0102: dc.fundoOpaco = u16(a) === 2; break;                     // SETBKMODE
          case 0x012E: dc.alinh = u16(a); break;                                // SETTEXTALIGN
          case 0x0106: dc.enchimento = u16(a); break;                           // SETPOLYFILLMODE
          case 0x02FA: novoObj({tipo: 'pen', estilo: u16(a) & 15, largura: i16(a + 2), cor: cor(a + 6)}); break;          // CREATEPENINDIRECT
          case 0x02FC: novoObj({tipo: 'brush', estilo: u16(a), cor: cor(a + 2)}); break;                                  // CREATEBRUSHINDIRECT
          case 0x0142: novoObj({tipo: 'brush', estilo: 3, cor: 0x808080, padrao: dibEm(a + 4)}); break;                   // DIBCREATEPATTERNBRUSH
          case 0x01F9: novoObj({tipo: 'brush', estilo: 0, cor: 0x808080}); break;                                        // CREATEPATTERNBRUSH (antigo)
          case 0x02FB: {                                                                                                // CREATEFONTINDIRECT
            let nome = ''; for (let i = 0; i < 32 && a + 18 + i < q + sz; i++) { const c = dv.getUint8(a + 18 + i); if (!c) break; nome += String.fromCharCode(c); }
            novoObj({tipo: 'font', altura: i16(a), esc: i16(a + 4), peso: i16(a + 8), italico: dv.getUint8(a + 10), nome: nome || 'Arial'}); break;
          }
          case 0x00F7: case 0x06FF: novoObj({tipo: 'outro'}); break;           // CREATEPALETTE / CREATEREGION: ocupam lugar na tabela
          case 0x012D: { const o = objs[u16(a)]; if (o && o.tipo !== 'outro') dc[o.tipo] = o; break; }                    // SELECTOBJECT
          case 0x01F0: objs[u16(a)] = undefined; break;                         // DELETEOBJECT
          case 0x0214: dc.y = i16(a); dc.x = i16(a + 2); break;                 // MOVETO (y, x)
          case 0x0213: {                                                        // LINETO (y, x)
            const y = i16(a), x = i16(a + 2);
            formar(caminhoDe([[X(dc.x), Y(dc.y)], [X(x), Y(y)]], false), false); dc.x = x; dc.y = y; desenhou++; break;
          }
          case 0x041B: { const [l, t, r, b] = ret(a); formar(caminhoDe([[l, t], [r, t], [r, b], [l, b]], true), true); desenhou++; break; }   // RECTANGLE
          case 0x061C: {                                                        // ROUNDRECT (height, width, rect)
            const rh = i16(a) * ky() / 2, rw = i16(a + 2) * kx() / 2, [l, t, r, b] = ret(a + 4), c = new Path2D();
            c.moveTo(l + rw, t); c.lineTo(r - rw, t); c.quadraticCurveTo(r, t, r, t + rh); c.lineTo(r, b - rh); c.quadraticCurveTo(r, b, r - rw, b);
            c.lineTo(l + rw, b); c.quadraticCurveTo(l, b, l, b - rh); c.lineTo(l, t + rh); c.quadraticCurveTo(l, t, l + rw, t); c.closePath();
            formar(c, true); desenhou++; break;
          }
          case 0x0418: { const [l, t, r, b] = ret(a), c = new Path2D(); c.ellipse((l + r) / 2, (t + b) / 2, (r - l) / 2, (b - t) / 2, 0, 0, Math.PI * 2); formar(c, true); desenhou++; break; }   // ELLIPSE
          case 0x0817: case 0x081A: case 0x0830: {                              // ARC / PIE / CHORD
            const ye = Y(i16(a)), xe = X(i16(a + 2)), ys = Y(i16(a + 4)), xs = X(i16(a + 6)), [l, t, r, b] = ret(a + 8);
            const cx = (l + r) / 2, cy = (t + b) / 2, rx = (r - l) / 2 || 1, ry = (b - t) / 2 || 1;
            const ang = (x, y) => Math.atan2((y - cy) / ry, (x - cx) / rx), c = new Path2D();
            if (fn === 0x081A) c.moveTo(cx, cy);
            c.ellipse(cx, cy, rx, ry, 0, ang(xs, ys), ang(xe, ye), true);     // sentido anti-horário (padrão do GDI)
            if (fn !== 0x0817) c.closePath();
            formar(c, fn !== 0x0817); desenhou++; break;
          }
          case 0x0324: case 0x0325: { const n = i16(a); formar(caminhoDe(pontos(a + 2, n), fn === 0x0324), fn === 0x0324); desenhou++; break; }   // POLYGON / POLYLINE
          case 0x0538: {                                                        // POLYPOLYGON
            const nPol = u16(a), c = new Path2D(); let o = a + 2 + nPol * 2;
            for (let i = 0; i < nPol; i++) { const n = u16(a + 2 + i * 2); pontos(o, n).forEach(([x, y], j) => j ? c.lineTo(x, y) : c.moveTo(x, y)); c.closePath(); o += n * 4; }
            formar(c, true); desenhou++; break;
          }
          case 0x0416: {                                                        // INTERSECTCLIPRECT
            let c = ret(a);
            if (dc.clip) c = [Math.max(c[0], dc.clip[0]), Math.max(c[1], dc.clip[1]), Math.min(c[2], dc.clip[2]), Math.min(c[3], dc.clip[3])];
            dc.clip = c; aplicarClip(); break;
          }
          case 0x012C: dc.clip = null; aplicarClip(); break;                    // SELECTCLIPREGION (região nula = sem recorte)
          case 0x0521: {                                                        // TEXTOUT
            const n = i16(a), o = a + 2 + ((n + 1) & ~1);
            texto(lerTexto(a + 2, n), i16(o + 2), i16(o)); desenhou++; break;
          }
          case 0x0A32: {                                                        // EXTTEXTOUT
            const y = i16(a), x = i16(a + 2), n = i16(a + 4), op = u16(a + 6);
            let o = a + 8;
            if (op & 0x0006) {
              if (op & 0x0002) { const [l, t, r, b] = [X(i16(o)), Y(i16(o + 2)), X(i16(o + 4)), Y(i16(o + 6))]; g.fillStyle = corDe(dc.corFundo); g.fillRect(Math.min(l, r), Math.min(t, b), Math.abs(r - l), Math.abs(b - t)); }
              o += 8;
            }
            if (n > 0 && o + n <= q + sz) { texto(lerTexto(o, n), x, y); desenhou++; }
            break;
          }
          case 0x061D: {                                                        // PATBLT (rop, height, width, y, x)
            const rop = dv.getUint32(a, true), h = i16(a + 4), w = i16(a + 6), y = i16(a + 8), x = i16(a + 10);
            const f = rop === 0x00FF0062 ? '#fff' : rop === 0x00000042 ? '#000' : rop === 0x00F00021 ? pincel() : null;
            if (f) { g.fillStyle = f; const x0 = X(x), y0 = Y(y), x1 = X(x + w), y1 = Y(y + h); g.fillRect(Math.min(x0, x1), Math.min(y0, y1), Math.abs(x1 - x0), Math.abs(y1 - y0)); desenhou++; }
            break;
          }
          case 0x0F43: case 0x0B41: case 0x0940: {                              // STRETCHDIB / DIBSTRETCHBLT / DIBBITBLT
            let o = a + 4, hd, wd, yd, xd, sh = 0, sw = 0, sy = 0, sx = 0;
            const comBitmap = sz > 6 + ((fn >> 8) + 3) * 2 - 6;
            if (fn === 0x0F43) { o += 2; sh = i16(o); sw = i16(o + 2); sy = i16(o + 4); sx = i16(o + 6); hd = i16(o + 8); wd = i16(o + 10); yd = i16(o + 12); xd = i16(o + 14); o += 16; }
            else if (fn === 0x0B41) { sh = i16(o); sw = i16(o + 2); sy = i16(o + 4); sx = i16(o + 6); hd = i16(o + 8); wd = i16(o + 10); yd = i16(o + 12); xd = i16(o + 14); o += 16; }
            else { sy = i16(o); sx = i16(o + 2); hd = i16(o + 4); wd = i16(o + 6); yd = i16(o + 8); xd = i16(o + 10); sh = hd; sw = wd; o += 12; }
            if (!comBitmap || o + 40 > q + sz) break;
            const c = dibEm(o);
            if (!c) break;
            const x0 = X(xd), y0 = Y(yd), x1 = X(xd + wd), y1 = Y(yd + hd);
            // origem do recorte: no DIB "de baixo para cima" o y conta a partir de baixo
            const srcW = Math.abs(sw) || c.width, srcH = Math.abs(sh) || c.height;
            const srcY = dv.getInt32(o + 8, true) > 0 ? c.height - sy - srcH : sy;
            g.drawImage(c, sx, Math.max(0, srcY), srcW, srcH, Math.min(x0, x1), Math.min(y0, y1), Math.abs(x1 - x0), Math.abs(y1 - y0));
            desenhou++; break;
          }
          default: break;
        }
      }
    } catch (e) { console.warn('WMF:', e); }
    g.restore();
    if (!desenhou) return null;
    return cv.toDataURL('image/png');
  }
  return {paraPNG, wmfParaPNG};
})();
