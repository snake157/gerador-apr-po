// "Atualizar biblioteca": o próprio navegador (Edge/Chrome) lê a pasta Pictogramas,
// reduz os pictogramas de Perigos, EPIs (e EPIs Fundo 02, se existir) e grava Pictogramas\biblioteca.js,
// que o gerador carrega no lugar do biblioteca.js da pasta do gerador.
'use strict';

const ATUALIZAR_BIB = (() => {
  const EXT = /\.(png|jpe?g)$/i;
  function carregar(arq) {
    return new Promise((ok, erro) => {
      const url = URL.createObjectURL(arq), im = new Image();
      im.onload = () => { URL.revokeObjectURL(url); ok(im); };
      im.onerror = () => { URL.revokeObjectURL(url); erro(new Error('não consegui abrir ' + arq.name)); };
      im.src = url;
    });
  }
  // miniatura (lado maior = lado px): WebP 0,88 para pictogramas, PNG para os ícones
  async function dataUrl(arq, lado = 360, formato = 'webp') {
    const im = await carregar(arq);
    const k = Math.min(1, lado / Math.max(im.naturalWidth, im.naturalHeight));
    const c = document.createElement('canvas');
    c.width = Math.max(1, Math.round(im.naturalWidth * k)); c.height = Math.max(1, Math.round(im.naturalHeight * k));
    const g = c.getContext('2d'); g.imageSmoothingQuality = 'high';
    g.drawImage(im, 0, 0, c.width, c.height);
    return formato === 'webp' ? c.toDataURL('image/webp', 0.88) : c.toDataURL('image/png');
  }
  async function subpasta(dir, nome) {
    for await (const [n, h] of dir.entries()) if (h.kind === 'directory' && n.toLowerCase() === nome.toLowerCase()) return h;
    return null;
  }
  async function pasta(dir, progresso) {
    if (!dir) return [];
    const arqs = [];
    for await (const [n, h] of dir.entries()) if (h.kind === 'file' && EXT.test(n)) arqs.push([n, h]);
    arqs.sort((a, b) => a[0].toLowerCase() < b[0].toLowerCase() ? -1 : a[0].toLowerCase() > b[0].toLowerCase() ? 1 : 0);
    const itens = [];
    for (const [n, h] of arqs) {
      try { itens.push({nome: n.replace(/\.[^.]+$/, ''), img: await dataUrl(await h.getFile())}); }
      catch (e) { progresso.erros.push(n); }
      progresso.feitos++; progresso.mostrar();
    }
    return itens;
  }
  // pic: a pasta Pictogramas (com Perigos e EPIs dentro)
  async function montar(pic, aviso) {
    const dirs = {perigos: await subpasta(pic, 'Perigos'), epis1: await subpasta(pic, 'EPIs'), epis2: await subpasta(pic, 'EPIs Fundo 02')};
    if (!dirs.perigos && !dirs.epis1) throw new Error('a pasta escolhida não tem as pastas "Perigos" e "EPIs". Escolha a pasta "Pictogramas" do gerador.');
    let total = 0;
    for (const d of Object.values(dirs)) if (d) for await (const [n, h] of d.entries()) if (h.kind === 'file' && EXT.test(n)) total++;
    const progresso = {feitos: 0, erros: [], mostrar: () => aviso(`Lendo pictogramas... ${progresso.feitos} de ${total}`)};
    const bib = {perigos: await pasta(dirs.perigos, progresso), epis1: await pasta(dirs.epis1, progresso), epis2: await pasta(dirs.epis2, progresso), icones: {}};
    // ícones da legenda e logos: continuam os da biblioteca atual
    bib.icones = Object.assign({}, BIB.icones || {});
    return {bib, erros: progresso.erros};
  }
  async function gravar(raiz, bib) {
    const h = await raiz.getFileHandle('biblioteca.js', {create: true});
    const w = await h.createWritable();
    await w.write('window.BIBLIOTECA = ' + JSON.stringify(bib) + ';\n');
    await w.close();
  }
  async function executar() {
    if (!window.showDirectoryPicker) { alert('Este navegador não permite escolher pastas. Use o Edge ou o Chrome.'); return; }
    // O navegador não deixa a página abrir pastas sozinha: na 1ª vez a pasta Pictogramas é escolhida e fica
    // lembrada; nas próximas basta confirmar a permissão do navegador (se ele pedir).
    let raiz = await rascunho.ler('pastaPictogramas');
    if (raiz) {
      try {
        let perm = await raiz.queryPermission({mode: 'readwrite'});
        if (perm !== 'granted') perm = await raiz.requestPermission({mode: 'readwrite'});
        if (perm !== 'granted') raiz = null;
        else if (!(await subpasta(raiz, 'Perigos')) && !(await subpasta(raiz, 'EPIs'))) raiz = null;   // pasta mudou de lugar
      } catch (e) { raiz = null; }
    }
    if (!raiz) {
      if (!(await perguntar('Na próxima janela, escolha a pasta "Pictogramas" do gerador (a que tem as pastas Perigos e EPIs). O navegador vai pedir permissão para gravar o biblioteca.js nela.\n\nIsso só é preciso na primeira vez: depois o gerador lembra a pasta.', {sim: 'ESCOLHER A PASTA', titulo: 'ATUALIZAR BIBLIOTECA'}))) return;
      try { raiz = await window.showDirectoryPicker({id: 'pictogramas', mode: 'readwrite'}); }
      catch (e) { if (e.name !== 'AbortError') alert('Não consegui abrir a pasta: ' + erroAmigavel(e)); return; }
    }
    try {
      const {bib, erros} = await montar(raiz, aviso);
      await gravar(raiz, bib);
      rascunho.salvar(raiz, 'pastaPictogramas');   // lembra a pasta para a próxima vez
      // passa a valer já, sem recarregar a página
      Object.assign(BIB, bib);
      renderEditor(); renderPrevia();
      aviso(`Biblioteca atualizada: ${bib.perigos.length} perigos e ${bib.epis1.length} EPIs.`);
      alert(`Biblioteca atualizada.\n\nPerigos: ${bib.perigos.length}\nEPIs: ${bib.epis1.length}` + (bib.epis2.length ? `\nEPIs Fundo 02: ${bib.epis2.length}` : '') +
        (erros.length ? `\n\nNão consegui ler: ${erros.join(', ')}` : '') + '\n\nO arquivo biblioteca.js da pasta Pictogramas foi regravado.');
    } catch (e) { console.error(e); alert('Não consegui atualizar a biblioteca: ' + erroAmigavel(e)); }
  }
  return {executar, montar};
})();
