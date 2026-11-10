// Pictogramas aprendidos: quando você troca uma imagem não reconhecida (trazida de um Word ou Excel antigo)
// pelo pictograma certo da biblioteca, o gerador guarda a "assinatura" daquela imagem com o nome escolhido.
// Na próxima importação a mesma imagem já vem com o nome.
// CONGELADO (10/11/2026, pedido do usuário): com o gerador online e usado por várias pessoas, o gerador não aprende
// mais trocas novas (cada um preferindo um pictograma ia ficar mudando o padrão) e não pede mais a pasta Pictogramas.
// Continua usando o que já estava aprendido: Pictogramas\aprendidos.js, o treino (recursos\pictos_treino.js) e as
// trocas já guardadas neste navegador antes do congelamento.
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
  // assinatura (texto compacto, 1 byte por ponto) -> Float32Array normalizada
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

  // troca feita no editor: congelado, não aprende mais nem grava nada (10/11/2026)
  async function aprender() {}
  return {buscar, aprender, total: () => lista.length};
})();
