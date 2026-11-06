// Leitor de Excel antigo (.xls, Excel 97-2003) feito no próprio gerador: lê o arquivo "composto" (CFB),
// acha o fluxo Workbook e os registros BIFF8 com os textos e números das células e as mesclagens.
// Devolve as abas no mesmo formato do leitor de .xlsx da APR: [{nome, ler() -> {A1: 'texto', '~A2': 'texto da mesclada'}}].
// Não lê imagens nem formatação (a APR só precisa dos textos).
'use strict';

const XLS = (() => {
  const ehXLS = buf => { const b = new Uint8Array(buf, 0, Math.min(8, buf.byteLength)); return b[0] === 0xD0 && b[1] === 0xCF && b[2] === 0x11 && b[3] === 0xE0; };

  // ---------- arquivo composto (CFB) ----------
  function cfb(buf) {
    const dv = new DataView(buf), u8 = new Uint8Array(buf);
    const ssz = 1 << dv.getUint16(30, true), mssz = 1 << dv.getUint16(32, true);
    const nFat = dv.getUint32(44, true), dir0 = dv.getUint32(48, true), corte = dv.getUint32(56, true);
    const mfat0 = dv.getUint32(60, true), nMfat = dv.getUint32(64, true);
    let difat = dv.getUint32(68, true), nDifat = dv.getUint32(72, true);
    const off = s => (s + 1) * ssz;
    const fatSec = [];
    for (let i = 0; i < 109 && fatSec.length < nFat; i++) fatSec.push(dv.getUint32(76 + i * 4, true));
    while (fatSec.length < nFat && difat < 0xFFFFFFFA && nDifat-- > 0) {
      for (let i = 0; i < ssz / 4 - 1 && fatSec.length < nFat; i++) fatSec.push(dv.getUint32(off(difat) + i * 4, true));
      difat = dv.getUint32(off(difat) + ssz - 4, true);
    }
    const fat = new Uint32Array(fatSec.length * ssz / 4);
    fatSec.forEach((s, k) => { for (let i = 0; i < ssz / 4; i++) if (off(s) + i * 4 + 4 <= buf.byteLength) fat[k * ssz / 4 + i] = dv.getUint32(off(s) + i * 4, true); });
    const cadeia = (s, tab) => { const r = []; let n = 0; while (s < 0xFFFFFFFA && s < tab.length && n++ < 1e6) { r.push(s); s = tab[s]; } return r; };
    const lerGrande = (s, tam) => {
      const out = new Uint8Array(tam); let p = 0;
      for (const x of cadeia(s, fat)) { const n = Math.min(ssz, tam - p); if (n <= 0) break; out.set(u8.subarray(off(x), off(x) + n), p); p += n; }
      return out;
    };
    // diretório
    const dirB = lerGrande(dir0, cadeia(dir0, fat).length * ssz), dd = new DataView(dirB.buffer);
    const ent = [];
    for (let p = 0; p + 128 <= dirB.length; p += 128) {
      const nl = dd.getUint16(p + 64, true); let nome = '';
      for (let i = 0; i < Math.max(0, nl / 2 - 1); i++) nome += String.fromCharCode(dd.getUint16(p + i * 2, true));
      ent.push({nome, tipo: dd.getUint8(p + 66), ini: dd.getUint32(p + 116, true), tam: dd.getUint32(p + 120, true)});
    }
    const raiz = ent.find(e => e.tipo === 5);
    let mini = null, mfat = null;
    const lerMini = (s, tam) => {
      if (!mini) {
        mini = lerGrande(raiz.ini, raiz.tam);
        const b = lerGrande(mfat0, nMfat * ssz), d2 = new DataView(b.buffer);
        mfat = new Uint32Array(b.length / 4); for (let i = 0; i < mfat.length; i++) mfat[i] = d2.getUint32(i * 4, true);
      }
      const out = new Uint8Array(tam); let p = 0;
      for (const x of cadeia(s, mfat)) { const n = Math.min(mssz, tam - p); if (n <= 0) break; out.set(mini.subarray(x * mssz, x * mssz + n), p); p += n; }
      return out;
    };
    return nome => {
      const e = ent.find(x => x.tipo === 2 && x.nome.toLowerCase() === nome.toLowerCase());
      if (!e) return null;
      return e.tam < corte ? lerMini(e.ini, e.tam) : lerGrande(e.ini, e.tam);
    };
  }

  // ---------- BIFF8 ----------
  const col = n => { let s = ''; n++; while (n) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; };
  const numTxt = v => Number.isInteger(v) ? String(v) : String(+v.toPrecision(15));
  function rk(v) {
    let x;
    if (v & 2) x = v >> 2;
    else { const b = new DataView(new ArrayBuffer(8)); b.setUint32(4, v & 0xFFFFFFFC, true); b.setUint32(0, 0, true); x = b.getFloat64(0, true); }
    return v & 1 ? x / 100 : x;
  }
  // lê um XLUnicodeString num trecho contínuo (LABEL, STRING, nomes das abas)
  function texto(dv, p, curto) {
    const n = curto ? dv.getUint8(p) : dv.getUint16(p, true); p += curto ? 1 : 2;
    const fl = dv.getUint8(p++); let s = '';
    for (let i = 0; i < n; i++) s += String.fromCharCode(fl & 1 ? dv.getUint16(p + i * 2, true) : dv.getUint8(p + i));
    return s;
  }
  // tabela de textos compartilhados (SST + CONTINUE): um texto pode passar para o próximo bloco no meio
  function sst(blocos) {
    const out = [];
    let b = 0, p = 8;            // pula o total e os únicos
    const dv = () => blocos[b];
    const garantir = () => { if (p >= dv().byteLength && b + 1 < blocos.length) { b++; p = 0; } };
    const u8 = () => { garantir(); return dv().getUint8(p++); };
    const u16 = () => { garantir(); if (p + 2 > dv().byteLength) { const a = u8(), c = u8(); return a | c << 8; } const v = dv().getUint16(p, true); p += 2; return v; };
    const u32 = () => u16() | u16() << 16;
    const pular = n => { while (n > 0) { garantir(); const k = Math.min(n, dv().byteLength - p); p += k; n -= k; if (k === 0) break; } };
    const total = blocos[0].getUint32(4, true);
    for (let i = 0; i < total && b < blocos.length; i++) {
      if (p >= dv().byteLength && b + 1 >= blocos.length) break;
      const n = u16(); let fl = u8();
      const nRun = fl & 8 ? u16() : 0, nExt = fl & 4 ? u32() : 0;
      let s = '', falta = n;
      while (falta > 0) {
        if (p >= dv().byteLength) { if (b + 1 >= blocos.length) break; b++; p = 0; fl = (fl & ~1) | (dv().getUint8(p++) & 1); }   // novo bloco: novo byte de "16 bits?"
        const larg = fl & 1 ? 2 : 1, cabe = Math.min(falta, Math.floor((dv().byteLength - p) / larg));
        for (let k = 0; k < cabe; k++) s += String.fromCharCode(larg === 2 ? dv().getUint16(p + k * 2, true) : dv().getUint8(p + k));
        p += cabe * larg; falta -= cabe;
        if (cabe === 0 && falta > 0 && p < dv().byteLength) break;
      }
      pular(nRun * 4); pular(nExt);
      out.push(s);
    }
    return out;
  }

  function lerPlanilhas(buf) {
    const ler = cfb(buf);
    const wb = ler('Workbook') || ler('Book');
    if (!wb) throw new Error('não achei a planilha dentro do arquivo .xls');
    const dv = new DataView(wb.buffer, wb.byteOffset, wb.byteLength);
    // 1º passo: abas e textos compartilhados (no começo do arquivo, até o 1º EOF)
    const abas = []; let textos = [];
    let p = 0, sstBlocos = null;
    while (p + 4 <= dv.byteLength) {
      const t = dv.getUint16(p, true), n = dv.getUint16(p + 2, true), q = p + 4;
      if (t === 0x0809 && p > 0 && abas.length) { /* BOF de uma aba */ }
      if (t === 0x0085) abas.push({pos: dv.getUint32(q, true), tipo: dv.getUint8(q + 5), nome: texto(dv, q + 6, true)});
      else if (t === 0x00FC) sstBlocos = [new DataView(wb.buffer, wb.byteOffset + q, n)];
      else if (t === 0x003C && sstBlocos && sstBlocos.fim !== true) sstBlocos.push(new DataView(wb.buffer, wb.byteOffset + q, n));
      else if (sstBlocos) sstBlocos.fim = true;
      if (t === 0x000A) break;              // EOF do "globals"
      p = q + n;
    }
    if (sstBlocos) textos = sst(sstBlocos);
    return abas.filter(a => a.tipo === 0).map(a => ({nome: a.nome, ler: async () => {
      const cel = {}, merges = [];
      let p = a.pos, formula = null;
      while (p + 4 <= dv.byteLength) {
        const t = dv.getUint16(p, true), n = dv.getUint16(p + 2, true), q = p + 4;
        if (q + n > dv.byteLength) break;
        const pos = () => col(dv.getUint16(q + 2, true)) + (dv.getUint16(q, true) + 1);
        if (t === 0x00FD) { const s = textos[dv.getUint32(q + 6, true)]; if (s) cel[pos()] = s; }                       // LABELSST
        else if (t === 0x0204 || t === 0x00D6) { const s = texto(dv, q + 6); if (s) cel[pos()] = s; }                    // LABEL / RSTRING
        else if (t === 0x0203) cel[pos()] = numTxt(dv.getFloat64(q + 6, true));                                           // NUMBER
        else if (t === 0x027E) cel[pos()] = numTxt(rk(dv.getInt32(q + 6, true)));                                         // RK
        else if (t === 0x00BD) {                                                                                           // MULRK
          const r = dv.getUint16(q, true), c0 = dv.getUint16(q + 2, true), k = (n - 6) / 6;
          for (let i = 0; i < k; i++) cel[col(c0 + i) + (r + 1)] = numTxt(rk(dv.getInt32(q + 4 + i * 6 + 2, true)));
        } else if (t === 0x0006) {                                                                                         // FORMULA: resultado guardado
          formula = null;
          if (dv.getUint16(q + 12, true) === 0xFFFF) { if (dv.getUint8(q + 6) === 0) formula = pos(); else if (dv.getUint8(q + 6) === 1) cel[pos()] = dv.getUint8(q + 8) ? 'VERDADEIRO' : 'FALSO'; }
          else cel[pos()] = numTxt(dv.getFloat64(q + 6, true));
        } else if (t === 0x0207 && formula) { const s = texto(dv, q); if (s) cel[formula] = s; formula = null; }          // STRING (da fórmula)
        else if (t === 0x0205 && dv.getUint8(q + 7) === 0) cel[pos()] = dv.getUint8(q + 6) ? 'VERDADEIRO' : 'FALSO';        // BOOLERR
        else if (t === 0x00E5) { const k = dv.getUint16(q, true); for (let i = 0; i < k; i++) { const o = q + 2 + i * 8; merges.push([dv.getUint16(o, true), dv.getUint16(o + 2, true), dv.getUint16(o + 4, true), dv.getUint16(o + 6, true)]); } }
        else if (t === 0x000A) break;
        p = q + n;
      }
      // célula mesclada: as outras células do bloco valem o mesmo que a primeira
      for (const [r0, r1, c0, c1] of merges) {
        const v = cel[col(c0) + (r0 + 1)];
        if (v === undefined) continue;
        for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) { const k = col(c) + (r + 1); if (!(r === r0 && c === c0) && !(k in cel)) cel['~' + k] = v; }
      }
      return cel;
    }}));
  }
  return {ehXLS, lerPlanilhas};
})();
