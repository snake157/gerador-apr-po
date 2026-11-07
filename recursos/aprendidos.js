// Pictogramas aprendidos: quando você troca uma imagem não reconhecida (trazida de um Word ou Excel antigo)
// pelo pictograma certo da biblioteca, o gerador guarda a "assinatura" daquela imagem com o nome escolhido.
// Na próxima importação a mesma imagem já vem com o nome. Fica guardado no navegador e no arquivo
// Pictogramas\aprendidos.js (para valer em outros computadores que usam a mesma pasta).
'use strict';

const APRENDIDOS = (() => {
  const CHAVE = 'pictosAprendidos';
  const LIMIAR = 0.9;        // bem mais exigente que o reconhecimento normal: na dúvida, fica a imagem original
  let lista = Array.isArray(window.PICTOS_APRENDIDOS) ? window.PICTOS_APRENDIDOS.slice() : [];   // [{tipo, nome, h, a}]
  const pronto = (async () => {
    const salvos = await rascunho.ler(CHAVE);
    if (Array.isArray(salvos)) for (const x of salvos) if (!lista.some(y => y.h === x.h && y.tipo === x.tipo)) lista.push(x);
    for (const x of lista) x.nome = x.tipo === 'perigo' ? nomePerigoAtual(x.nome) : nomeEpiAtual(x.nome);   // nomes antigos (07/11/2026: EPIs também)
  })();

  // identificação exata da imagem (mesmo arquivo no Word)
  function hash(s) {
    let h = 0x811c9dc5;
    for (let i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 0x01000193) >>> 0; }
    return s.length.toString(36) + '-' + h.toString(36);
  }
  // assinatura (Float32Array normalizada) <-> texto compacto (1 byte por ponto)
  const comprimir = v => { let b = ''; for (const x of v) b += String.fromCharCode(Math.max(0, Math.min(255, Math.round(x * 32 + 128)))); return btoa(b); };
  const expandir = s => { const b = atob(s), v = new Float32Array(b.length); for (let i = 0; i < b.length; i++) v[i] = (b.charCodeAt(i) - 128) / 32; return v; };
  const nomesDe = tipo => new Set((tipo === 'perigo' ? BIB.perigos : listaEpis()).map(x => x.nome));

  // Pictogramas dos POs de treino (recursos\pictos_treino.js): nome da biblioteca (ou vazio = não existe nela)
  // e o nome escrito na imagem. Valem depois das trocas que você fizer (as suas sempre ganham).
  // (07/11/2026) o arquivo do treino (~1 MB) só é baixado na primeira importação, e não mais ao abrir o gerador
  let treino = [], treinoPronto = null;
  const porHash = {};
  function carregarTreino() {
    if (!treinoPronto) treinoPronto = (async () => {
      if (!window.PICTOS_TREINO) { try { await carregarScript('recursos/pictos_treino.js'); } catch (e) { /* sem treino: só as trocas aprendidas */ } }
      treino = Array.isArray(window.PICTOS_TREINO) ? window.PICTOS_TREINO : [];
      for (const x of treino) if (x.nome) x.nome = x.tipo === 'perigo' ? nomePerigoAtual(x.nome) : nomeEpiAtual(x.nome);   // nomes antigos dos perigos e EPIs
      for (const x of treino) porHash[x.tipo + ':' + x.h] = x;
    })();
    return treinoPronto;
  }
  const LIMIAR_INTEIRA = 0.95, LIMIAR_SIMBOLO = 0.92;
  const resposta = (x, d, comRotulo) => ({nome: x.nome || '', rotulo: comRotulo ? x.rotulo || '' : '', d, treino: true});

  // nome aprendido para esta imagem, ou null. inteiraFn: devolve a assinatura da imagem inteira (sob demanda)
  async function buscar(tipo, src, variantes, inteiraFn) {
    await pronto;
    await carregarTreino();
    if (!src) return null;
    const validos = nomesDe(tipo), h = hash(src);
    const exato = lista.find(x => x.tipo === tipo && x.h === h && validos.has(x.nome));
    if (exato) return {nome: exato.nome, d: 1};
    const tr = porHash[tipo + ':' + h];
    if (tr && (!tr.nome || validos.has(tr.nome))) return resposta(tr, 1, true);
    const {correlacao} = IMPORTAR.ajuda;
    if (variantes && variantes.length) {
      let melhor = null;
      for (const x of lista) {
        if (x.tipo !== tipo || !x.a || !validos.has(x.nome)) continue;
        const a = x._v || (x._v = expandir(x.a));
        if (a.length !== variantes[0].length) continue;
        const c = Math.max(...variantes.map(v => correlacao(v, a)));
        if (!melhor || c > melhor.c) melhor = {nome: x.nome, c};
      }
      if (melhor && melhor.c >= LIMIAR) return {nome: melhor.nome, d: melhor.c};
    }
    if (!treino.length) return null;
    // mesma imagem do treino em outro arquivo (redimensionada, recomprimida): compara a imagem inteira
    const w = inteiraFn ? await inteiraFn() : null;
    if (w) {
      let melhor = null;
      for (const x of treino) {
        if (x.tipo !== tipo || !x.w || (x.nome && !validos.has(x.nome))) continue;
        const a = x._w || (x._w = expandir(x.w));
        if (a.length !== w.length) continue;
        const c = correlacao(w, a);
        if (!melhor || c > melhor.c) melhor = {x, c};
      }
      if (melhor && melhor.c >= LIMIAR_INTEIRA) return resposta(melhor.x, melhor.c, true);
    }
    // só o símbolo parecido (nome escrito diferente): vale se todas as do treino parecidas concordam no nome
    if (variantes && variantes.length) {
      const nomes = new Set(); let cMax = 0;
      for (const x of treino) {
        if (x.tipo !== tipo || !x.a) continue;
        const a = x._v || (x._v = expandir(x.a));
        if (a.length !== variantes[0].length) continue;
        const c = Math.max(...variantes.map(v => correlacao(v, a)));
        if (c >= LIMIAR_SIMBOLO) { nomes.add(x.nome || ''); cMax = Math.max(cMax, c); }
      }
      if (nomes.size === 1) { const nome = [...nomes][0]; if (nome && validos.has(nome)) return {nome, rotulo: '', d: cMax, treino: true}; }
    }
    return null;
  }

  const paraGravar = () => lista.map(({_v, ...x}) => x);
  async function gravarArquivo(raiz) {
    const h = await raiz.getFileHandle('aprendidos.js', {create: true});
    const w = await h.createWritable();
    await w.write('// Pictogramas aprendidos pelo gerador (trocas feitas em imagens não reconhecidas). Pode apagar para esquecer tudo.\n' +
      'window.PICTOS_APRENDIDOS = ' + JSON.stringify(paraGravar()) + ';\n');
    await w.close();
  }
  // grava no arquivo Pictogramas\aprendidos.js a cada troca (pedido do usuário, 21/10/2026): usa a pasta Pictogramas
  // já lembrada (pede a permissão do navegador, se precisar); sem ela, pergunta uma vez e pede para escolher a pasta.
  // "AGORA NÃO" vale até fechar a página (o aprendido fica no navegador e vai para o arquivo na próxima troca).
  let recusou = false;
  const ehPastaPictogramas = async d => { for (const n of ['Perigos', 'EPIs']) { try { await d.getDirectoryHandle(n); return true; } catch (e) {} } return false; };
  async function gravarSePuder() {
    try {
      let raiz = await rascunho.ler('pastaPictogramas');
      if (raiz) {
        let p = await raiz.queryPermission({mode: 'readwrite'});
        if (p !== 'granted') { try { p = await raiz.requestPermission({mode: 'readwrite'}); } catch (e) { /* sem o clique: pergunta abaixo */ } }
        if (p === 'granted') { await gravarArquivo(raiz); return; }
      }
      if (recusou || !window.showDirectoryPicker || typeof perguntar !== 'function') return;
      const sim = await perguntar('O gerador aprendeu esta troca de pictograma. Para guardar no arquivo Pictogramas\\aprendidos.js (e valer também em outros computadores), ' +
        (raiz ? 'permita gravar na pasta Pictogramas.' : 'escolha a pasta "Pictogramas" do gerador na próxima janela.'),
        {sim: raiz ? 'PERMITIR' : 'ESCOLHER A PASTA', nao: 'AGORA NÃO', titulo: 'PICTOGRAMAS APRENDIDOS'});
      if (!sim) { recusou = true; return; }
      if (raiz && (await raiz.requestPermission({mode: 'readwrite'})) === 'granted') { await gravarArquivo(raiz); return; }
      raiz = await window.showDirectoryPicker({id: 'pictogramas', mode: 'readwrite'});
      if (!(await ehPastaPictogramas(raiz))) { avisar('Essa pasta não é a "Pictogramas" do gerador (a que tem as pastas Perigos e EPIs). O aprendido ficou só no navegador por enquanto.'); return; }
      await rascunho.salvar(raiz, 'pastaPictogramas');
      await gravarArquivo(raiz);
      aviso('Pictogramas aprendidos gravados em Pictogramas\\aprendidos.js.');
    } catch (e) { if (e.name !== 'AbortError') console.warn('aprendidos.js não gravado:', e); }
  }

  // troca feita no editor: imagem original (src) -> pictograma escolhido (nome)
  async function aprender(tipo, src, nome) {
    if (!src || !nome || !/^data:image\//.test(src)) return;
    await pronto;
    const h = hash(src);
    let a = null;
    try { const v = await IMPORTAR.ajuda.assinatura(src, false); if (v) a = comprimir(v); } catch (e) {}
    lista = lista.filter(x => !(x.tipo === tipo && x.h === h));
    lista.push({tipo, nome, h, a});
    await rascunho.salvar(paraGravar(), CHAVE);
    await gravarSePuder();
  }
  return {buscar, aprender, gravarArquivo, total: () => lista.length};
})();
