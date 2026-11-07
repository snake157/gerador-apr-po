// Gerador de APR (Análise Preliminar de Riscos), na aba "APR" do gerador.
// - A APR é um documento próprio, salvo em Excel (Salvar / Abrir / Salvar como, como o PO): o mesmo Excel do
//   "Exportar Excel", com a APR inteira guardada numa parte oculta (customXml). O .apr.json antigo ainda abre.
// - Prévia igual ao Excel exportado: mesmas colunas, alturas, páginas, logotipo, cabeçalho repetido e rodapé.
// - Exporta Excel (no "APR.XX.000 - Modelo Padrão", embutido em recursos/modelo_apr.js) e PDF (impressão da prévia).
// - Abre APR antiga em Excel (.xlsx ou .xls antigo, via recursos\xls.js): colunas achadas pelos títulos.
// - Complementos com o PO aberto na aba PO: "Importar do PO" (etapas, perigos, EPIs) e "Comparar com o PO":
//   tela de conferência no lugar da prévia, etapa por etapa (faixa com as etapas); etapa, perigos e EPIs em grade
//   (APR | situação / corrigir | PO), corrigindo a APR ou o PO; medidas da APR (M1, M2...) ao lado da descrição
//   do PO frase por frase, com "Editar descrição do PO" / "Escrever frase no PO" no editor de texto da aba PO.
//   É a mesma conferência ("Comparar com APR" na aba PO abre esta tela, sem trocar de aba).
'use strict';

const APR = (() => {
  const SITUACOES = [['R', 'R – Rotineira'], ['NR', 'NR – Não Rotineira']];
  const N = s => String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
  const esc = s => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');

  // ---------------- documento ----------------
  // metodo do perigo: vazio = o do cabeçalho (nº do PO); era da etapa até 28/10/2026 (pedido do usuário)
  const novaLinha = (perigo = '') => ({id: uid(), perigo, risco: '', epc: '', epis: '', metodo: '', po: '', pd: '', medidas: '', confirmadas: []});
  const novaEtapa = (nome = '') => ({id: uid(), nome, situacao: 'R', linhas: [novaLinha()]});
  function novaAPR() {
    return {
      tipo: 'apr', versao: 1,
      cab: {numero: '', area: '', departamento: '', maquina: '', atividade: '', avaliado: '', revAuto: true, revisao: '', data: '', episBasicos: '', metodo: '',
        empresa: '', elaboracao: ''},     // empresa e elaboração: só na Ordem de Serviço (vazios na APR nova)
      etapas: [novaEtapa()],
      historico: [{data: '', motivo: 'Criação da APR', responsavel: '', origem: 'Emissão'}],
      logo: '',      // logotipo escolhido (PNG em data URL); vazio = o do modelo
    };
  }
  function normalizar(d) {
    const b = novaAPR();
    d.tipo = 'apr'; d.versao = 1;
    d.cab = Object.assign(b.cab, d.cab || {});
    // .apr.json antigo (método na etapa): o método da etapa passa para cada perigo dela que não tem o seu
    d.etapas = (d.etapas || []).map(e => {
      const mEt = (e.metodo || '').trim();
      const ne = Object.assign(novaEtapa(), e, {
        linhas: (e.linhas || []).map(l => Object.assign(novaLinha(), l, {confirmadas: l.confirmadas || [], metodo: (l.metodo || '').trim() ? l.metodo : mEt,
          perigo: renomeado(l.perigo || '')})),   // perigos com o nome corrigido na planilha (07/11/2026)
      });
      delete ne.metodo;
      return ne;
    });
    d.historico = (d.historico || []).map(h => Object.assign({data: '', motivo: '', responsavel: '', origem: ''}, h));
    d.logo = typeof d.logo === 'string' && /^data:image\/png;base64,/.test(d.logo) ? d.logo : '';
    return d;
  }
  let apr = novaAPR();
  let arquivo = null, nomeAtual = '', nomeOrigem = '', alterado = false, carregando = false, salvoAs = '';

  const numero = v => { const n = parseInt(v, 10); return n >= 1 && n <= 5 ? n : null; };
  const riscoDe = l => { const po = numero(l.po), pd = numero(l.pd); return po && pd ? po * pd : null; };
  function nivel(r) {
    if (!r) return null;
    return r <= 4 ? {nome: 'Baixo', fundo: '#00B050', cor: '#fff'} : r <= 11 ? {nome: 'Moderado', fundo: '#FFFF00', cor: '#000'}
      : r <= 15 ? {nome: 'Alto', fundo: '#E46C0A', cor: '#000'} : {nome: 'Crítico', fundo: '#FF0000', cor: '#fff'};
  }
  function ultimaRev() { let u = -1; apr.historico.forEach((h, i) => { if ((h.motivo || '').trim()) u = i; }); return u; }
  function cabAPR() {
    const c = apr.cab, u = ultimaRev();
    const rev = c.revAuto ? pad2(Math.max(0, u)) : (c.revisao || '00');
    const data = c.revAuto ? (u >= 0 ? fmtData(apr.historico[u].data) : '') : fmtData(c.data);
    const num = (c.numero || '').trim() || 'XX.000';
    return {numero: num, area: c.area, departamento: c.departamento, maquina: c.maquina, atividade: c.atividade, avaliado: c.avaliado,
      data, rev, episBasicos: c.episBasicos, metodo: (c.metodo || '').trim() || 'PO.' + num};
  }
  const metodoDe = (l, ca) => (l && (l.metodo || '').trim()) || ca.metodo;
  function nomeAPR() {
    const c = cabAPR();
    return `APR.${c.numero} - ${(c.atividade || 'Sem nome').trim()}_Rev${c.rev}`.replace(/[\\/:*?"<>|]/g, '-');
  }
  const textoEpis = l => (l.epis || '').trim() || 'N/A';
  const textoEpc = l => (l.epc || '').trim() || 'N/A';
  const dividirEpis = t => String(t || '').split(/[,;\n]|\s+e\s+(?=[A-ZÁÉÍÓÚ])/).map(s => s.trim().replace(/\.$/, '')).filter(s => s && !/^(n\/?a|-|nenhum)$/i.test(s));

  // ---------------- tabela de Perigos e Riscos (recursos/perigos_riscos.js) ----------------
  // Lista fechada, exatamente como na planilha: cada linha é um perigo (com o número "1 - ", "2 - "... quando o
  // mesmo perigo tem categorias de risco diferentes) atrelado a um risco e ao potencial de dano (PD) sugerido.
  // Na APR o perigo é escolhido da lista; o risco vem junto e não muda; o PD vem preenchido e pode ser trocado.
  // perigo da planilha com o nome corrigido em 07/11/2026 (ex.: "Oxigenio" > "Oxigênio"): o nome antigo vira o novo
  const renomeado = p => { const m = window.PERIGOS_RENOMEADOS || {}; if (m[p]) return m[p]; const k = Object.keys(m).find(x => N(x) === N(p)); return k ? m[k] : p; };
  let TAB = [];
  function montarTabela() {
    TAB = (window.TABELA_PERIGOS || []).filter(r => String(r.perigo || '').trim())
      .map(r => ({perigo: renomeado(String(r.perigo).trim()), risco: String(r.risco || '').trim(), pd: numero(r.pd), obs: String(r.obs || '').trim()}));
  }
  montarTabela();
  const daTabela = perigo => TAB.find(r => r.perigo === String(perigo || '').trim()) || TAB.find(r => N(r.perigo) === N(perigo));
  function aplicarPerigo(l, r) { l.perigo = r.perigo; l.risco = r.risco; l.pd = r.pd ? String(r.pd) : ''; }
  // perigo escrito numa APR antiga ("Queda de Materiais e tombamento") -> linha da tabela com o mesmo texto,
  // sem ligar para pontuação, maiúsculas e o número "1 - "; havendo "1 -/2 -/3 -", vale a do PD da APR
  // (ou a de PD mais próximo). Texto diferente continua como está e aparece "fora da tabela".
  const chavePerigo = s => N(String(s || '').replace(/^\s*\d\s*-\s*/, '')).replace(/[^a-z0-9]+/g, ' ').trim();
  // "trocas seguras" da revisão de perigos das APRs (planilha "Revisão de perigos - responder", 26/10/2026): texto antigo ->
  // perigo da lista (o nº 1/2/3 sai pelo PD). Os duvidosos (Batida Contra, Atropelamento...) não entram: ficam fora da lista
  const TROCAS_SEGURAS = {
    'Ruído': 'Exposição ao Ruído',
    'Contato com material cortante ou perfurantes': 'Contato com material cortante ou perfurante.',
    'Inalação de aerosdisperssoides': 'Inalação de aerodispersoides.', 'Inalação de aerodisperssoides': 'Inalação de aerodispersoides.',
    'Inalação de aerosdispersoides': 'Inalação de aerodispersoides.',
    'Queda, esbarrões e tropeços': 'Quedas, Esbarrões e Tropeços',
    'Exposição a condições climaticas adversas (exposição ao sol, frio, chuva)': 'Exposição a condições climáticas adversas (exposição ao sol, frio, chuva e etc)',
    'Exposição a Radiação Não- Ionizante': 'Exposição à Radiação Não Ionizante',
    'Projeção ou Respingo de Metal Líquido': 'Projeção, Respingo ou Transbordo de Metal Líquido',
    'Contato com partes quentes': 'Exposição ao Calor e contato com partes quentes',
    'Postura Inadequada no Trabalho': 'Postura Inadequada no Trabalho e Esforço físico excessivo',
    'Postura Inadequada no Trabalho e Esforço fisico excessivo': 'Postura Inadequada no Trabalho e Esforço físico excessivo',
    'Projeção Fagulhas, Cavacos ou Particulas Volantes': 'Projeção de Fagulhas, Cavacos ou Partículas Volantes.',
    'Queda de materiais': 'Queda de Materiais e tombamento.', 'Queda de materiais e tombamentos': 'Queda de Materiais e tombamento.',
    'Quedas de Materiais e Tombamento': 'Queda de Materiais e tombamento.', 'Queda de materiais, Tombamento': 'Queda de Materiais e tombamento.',
    'Tombamento': 'Queda de Materiais e tombamento.',
    'Queda de materiais, Batida Contra': 'Queda de materiais, Tombamento, batida Contra e Atingido por Aprisionamento entre partes.',
    'Queda de materiais, Tombamento, Batida contra': 'Queda de materiais, Tombamento, batida Contra e Atingido por Aprisionamento entre partes.',
    'Movimentação de carga, Queda de materiais, Tombamento, Batida contra, Aprisionamento entre partes': 'Queda de materiais, Tombamento, batida Contra e Atingido por Aprisionamento entre partes.',
    // duvidosos da mesma família 1/2/3 (07/11/2026, regra do usuário: o PD decide o nº, PD 3 = 1, PD 4 = 2, PD 5 = 3)
    'Batida Contra': 'Queda de materiais, Tombamento, batida Contra e Atingido por Aprisionamento entre partes.',
    'Aprisionamento entre partes': 'Queda de materiais, Tombamento, batida Contra e Atingido por Aprisionamento entre partes.',
    'Atingido por Aprisionamento entre partes': 'Queda de materiais, Tombamento, batida Contra e Atingido por Aprisionamento entre partes.',
    'Queda de materiais, Atingido por Aprisionamento entre partes': 'Queda de materiais, Tombamento, batida Contra e Atingido por Aprisionamento entre partes.',
    'Queda de materiais, Tombamento, Atingido por Aprisionamento entre partes': 'Queda de materiais, Tombamento, batida Contra e Atingido por Aprisionamento entre partes.',
    'Batida Contra e Atingido por Aprisionamento entre partes': 'Queda de materiais, Tombamento, batida Contra e Atingido por Aprisionamento entre partes.',
    'Atropelamento': 'Atropelamento, colisão.', 'Atropelamento, Batida Contra': 'Atropelamento, colisão.',
    'Queda de materiais, Batida Contra e Atropelamento': 'Atropelamento, colisão.',
    'Queda de nível diferente': 'Queda de nível diferente e queda de materiais.',
    // restantes, sugestões aprovadas pelo usuário (07/11/2026)
    'Projeção, batida contra': 'Resvalamento e Projeções',
    'Contato com material cortante ou perfurante, Tropeções, Movimentação de carga, queda de materiais': 'Contato com material cortante ou perfurante.',
    'Queda de nível diferente, Queda de materiais, Batida Contra': 'Queda de nível diferente e queda de materiais.',
    'Projeção de escória': 'Projeção, Respingo ou Transbordo de Metal Líquido',
    'Chuvas intensas': 'Exposição a condições climáticas adversas (exposição ao sol, frio, chuva e etc)',
    'Projeção': 'Projeção de Fagulhas, Cavacos ou Partículas Volantes.',
    'Tropeços, quedas, batida contra': 'Quedas, Esbarrões e Tropeços',
    'Queda de material, tropeços e quedas': 'Quedas, Esbarrões e Tropeços',
    'Esforço Físico, queda do material': 'Postura Inadequada no Trabalho e Esforço físico excessivo',
    'Atropelamento, Queda de material': 'Atropelamento, colisão.',
    'Exposição a poeiras sem proteção': 'Inalação de aerodispersoides.',
  };
  let trocas = null;
  function perigoDaTabela(perigo, pd) {
    if (!trocas) { trocas = new Map(); for (const [a, b] of Object.entries(TROCAS_SEGURAS)) trocas.set(chavePerigo(a), chavePerigo(b)); }
    let k = chavePerigo(perigo);
    if (k && !TAB.some(r => chavePerigo(r.perigo) === k) && trocas.has(k)) k = trocas.get(k);
    const cands = k ? TAB.filter(r => chavePerigo(r.perigo) === k) : [];
    if (!cands.length) return null;
    const alvo = numero(pd);
    return cands.slice().sort((a, b) => Math.abs((a.pd || 0) - alvo) - Math.abs((b.pd || 0) - alvo))[0];
  }
  // só uma sugestão para o editor ("Ruído" ~ "Exposição ao Ruído"): palavras em comum, sem "de", "e", "ao"...
  const palavrasPerigo = s => new Set(chavePerigo(s).split(' ').filter(w => w.length > 2 && !/^(com|por|para|ao|aos|das|dos|nos|nas|sem)$/.test(w)).map(w => w.replace(/s$/, '')));
  function parecidoNaTabela(perigo, pd) {
    const a = palavrasPerigo(perigo);
    if (!a.size) return null;
    let melhor = null;
    for (const r of TAB) {
      const b = palavrasPerigo(r.perigo), comum = [...a].filter(w => b.has(w)).length;
      const j = comum / (a.size + b.size - comum);
      if (comum && (!melhor || j > melhor.j || (j === melhor.j && Math.abs((r.pd || 0) - (numero(pd) || 0)) < Math.abs((melhor.r.pd || 0) - (numero(pd) || 0))))) melhor = {r, j};
    }
    return melhor && melhor.j >= 0.4 ? melhor.r : null;
  }
  // janela para escolher o perigo da tabela, com busca pelo nome (ou pelo risco)
  function escolherPerigo(titulo, atual, aoEscolher) {
    const md = $('#modal');
    const busca = el('input', {type: 'text', placeholder: 'Buscar pelo nome do perigo ou do risco...'});
    const lista = el('div', {class: 'pp-lista'});
    const fechar = () => { md.hidden = true; md.replaceChildren(); };
    const escolher = r => { fechar(); aoEscolher(r); };
    let achados = [];
    const desenhar = () => {
      const q = N(busca.value).split(' ').filter(Boolean);
      achados = TAB.filter(r => { const tx = N(r.perigo + ' ' + r.risco); return q.every(w => tx.includes(w)); });
      lista.replaceChildren(...(achados.length ? achados.map(r => el('button', {class: 'pp-item' + (N(r.perigo) === N(atual) ? ' atual' : ''), onclick: () => escolher(r)},
        el('b', {}, r.perigo), el('span', {}, `${r.risco}${r.pd ? ' · PD ' + r.pd : ''}`))) : [el('p', {class: 'dica', style: 'padding:8px'}, 'Nenhum perigo encontrado.')]));
    };
    busca.addEventListener('input', desenhar);
    busca.addEventListener('keydown', ev => { if (ev.key === 'Enter' && achados.length) { ev.preventDefault(); escolher(achados[0]); } else if (ev.key === 'Escape') fechar(); });
    md.replaceChildren(el('div', {class: 'janela pp-janela', onclick: ev => ev.stopPropagation()},
      el('header', {}, el('b', {}, titulo), busca, el('button', {class: 'btn', onclick: fechar}, 'Fechar')),
      atual ? el('div', {class: 'dica', style: 'padding:6px 12px 0'}, 'Atual: ' + atual) : null,
      lista));
    md.onclick = fechar; md.hidden = false;
    desenhar(); busca.focus();
    const a = lista.querySelector('.pp-item.atual'); if (a) a.scrollIntoView({block: 'center'});
  }
  // EPI's específicos: escolher da lista da biblioteca (com busca), sem perder o texto digitado à mão (21/10/2026).
  // Cada clique põe ou tira o EPI do campo (separado por vírgula); o que foi escrito à mão continua lá.
  function escolherEpis(l, aoMudar) {
    const md = $('#modal');
    const busca = el('input', {type: 'text', placeholder: 'Buscar o EPI...'});
    const lista = el('div', {class: 'pp-lista'});
    const fechar = () => { md.hidden = true; md.replaceChildren(); aoMudar(); };
    const vistos = new Set();
    const epis = listaEpis().map(p => ({rot: rotuloPadrao(p.nome), img: p.img})).filter(p => !vistos.has(N(p.rot)) && vistos.add(N(p.rot)));
    const noCampo = () => dividirEpis(l.epis);
    const alternar = rot => {
      const atuais = noCampo();
      const novo = atuais.some(t => N(t) === N(rot)) ? atuais.filter(t => N(t) !== N(rot)) : atuais.concat(rot);
      l.epis = novo.join(', '); mudou(); desenhar();
    };
    const desenhar = () => {
      const q = N(busca.value).split(' ').filter(Boolean), tem = new Set(noCampo().map(N));
      const achados = epis.filter(p => q.every(w => N(p.rot).includes(w)));
      lista.replaceChildren(el('div', {class: 'dica', style: 'margin:0 0 4px'}, 'No campo: ' + (l.epis.trim() || 'N/A')),
        ...(achados.length ? achados.map(p => el('button', {class: 'pp-item pe-item' + (tem.has(N(p.rot)) ? ' atual' : ''), onclick: () => alternar(p.rot)},
          p.img ? el('img', {src: p.img, alt: ''}) : null, el('b', {}, (tem.has(N(p.rot)) ? '✓ ' : '') + p.rot))) : [el('p', {class: 'dica', style: 'padding:8px'}, 'Nenhum EPI encontrado. Escreva o nome direto no campo.')]));
    };
    busca.addEventListener('input', desenhar);
    busca.addEventListener('keydown', ev => { if (ev.key === 'Escape') fechar(); });
    md.replaceChildren(el('div', {class: 'janela pp-janela', onclick: ev => ev.stopPropagation()},
      el('header', {}, el('b', {}, "EPI's específicos"), busca, el('button', {class: 'btn primario', onclick: fechar}, 'OK')),
      lista));
    md.onclick = fechar; md.hidden = false;
    desenhar(); busca.focus();
  }
  // janela para editar um texto curto (nome da etapa, EPI, medida)
  function pedirTexto(titulo, valor, aoSalvar, linhas = 1) {
    const md = $('#modal');
    const fechar = () => { md.hidden = true; md.replaceChildren(); };
    const inp = linhas > 1 ? el('textarea', {rows: linhas, value: valor}) : el('input', {type: 'text', value: valor});
    const ok = () => { const v = inp.value.trim(); if (!v) { inp.focus(); return; } fechar(); if (v !== valor) aoSalvar(v); };
    inp.addEventListener('keydown', ev => { if (ev.key === 'Enter' && (linhas === 1 || ev.ctrlKey)) { ev.preventDefault(); ok(); } else if (ev.key === 'Escape') fechar(); });
    md.replaceChildren(el('div', {class: 'janela pt-janela', onclick: ev => ev.stopPropagation()},
      el('header', {}, el('b', {}, titulo), el('span', {style: 'flex:1'}), el('button', {class: 'btn', onclick: fechar}, 'Cancelar'), el('button', {class: 'btn primario', onclick: ok}, 'Salvar')),
      el('div', {style: 'padding:12px'}, inp)));
    md.onclick = fechar; md.hidden = false;
    inp.focus(); inp.select();
  }
  const linhaDe = txt => { const l = novaLinha(txt); perigoEscolhido(l); return l; };
  // texto vindo do PO ou de uma APR antiga: se for um perigo da tabela, entra com o risco e o PD dela
  function perigoEscolhido(l) {
    const r = daTabela(l.perigo);
    if (!r) return false;
    aplicarPerigo(l, r);
    return true;
  }
  // Menu > Atualizar Tabela de Perigos: o usuário escolhe a planilha Excel; o gerador grava recursos\perigos_riscos.js.
  async function lerTabelaXlsx(f) {
    const L = porLinha(await (await lerPlanilhas(await f.arrayBuffer()))[0].ler());
    const nums = Object.keys(L).map(Number).sort((a, b) => a - b);
    const hr = nums.find(n => Object.values(L[n]).some(v => N(v) === 'perigo'));
    if (!hr) throw new Error('não achei a coluna "Perigo"');
    const col = {};
    for (const [c, v] of Object.entries(L[hr])) { const t = N(v); if (t === 'perigo') col.p = c; else if (t.startsWith('risco')) col.r = c; else if (t.startsWith('potencial') || t === 'pd') col.d = c; else if (t.startsWith('observ')) col.o = c; }
    const linhas = [];
    for (const n of nums.filter(n => n > hr)) {
      const R = L[n], p = String(R[col.p] || '').trim();
      if (p) linhas.push({perigo: p, risco: String(R[col.r] || '').trim(), pd: numero(R[col.d]), obs: String(R[col.o] || '').trim()});
    }
    if (!linhas.length) throw new Error('a planilha não tem linhas de perigo');
    return linhas;
  }
  const textoTabela = (linhas, nome) => '// Tabela de Perigos e Riscos da APR (perigo, risco / consequência, potencial de dano e observação).\n' +
    `// Gerada a partir de "${nome}" (${new Date().toLocaleDateString('pt-BR')}). Para atualizar: aba APR > Menu > "Atualizar Tabela de Perigos".\n` +
    'window.TABELA_PERIGOS = ' + JSON.stringify(linhas, null, 0) + ';\n';
  // Pasta recursos já autorizada antes (lembrada pelo navegador), se ainda tiver permissão; senão null.
  async function pastaRecursos() {
    const dir = await rascunho.ler('apr:pastaRecursos');
    if (!dir || !dir.queryPermission) return null;
    try { return (await dir.queryPermission({mode: 'readwrite'})) === 'granted' ? dir : null; } catch (e) { return null; }
  }
  async function escolherPlanilha() {
    if (window.showOpenFilePicker) {
      const [h] = await window.showOpenFilePicker({id: 'planilha-perigos', types: [{description: 'Planilha do Excel', accept: {'application/vnd.ms-excel': ['.xlsx', '.xls']}}]});
      return h.getFile();
    }
    return new Promise((ok, no) => {
      const i = document.createElement('input'); i.type = 'file'; i.accept = '.xlsx,.xls';
      i.onchange = () => i.files[0] ? ok(i.files[0]) : no({name: 'AbortError'});
      i.click();
    });
  }
  async function gravarTabela(texto) {
    const dir = await pastaRecursos();
    if (dir) { const h = await dir.getFileHandle('perigos_riscos.js', {create: true}); const w = await h.createWritable(); await w.write(texto); await w.close(); return 'gravado'; }
    if (window.showSaveFilePicker) {
      await avisar('Na próxima janela, salve o arquivo "perigos_riscos.js" dentro da pasta "recursos" do gerador (confirme a substituição do arquivo atual).');
      const h = await window.showSaveFilePicker({id: 'pasta-recursos', suggestedName: 'perigos_riscos.js', types: [{description: 'JavaScript', accept: {'text/javascript': ['.js']}}]});
      const w = await h.createWritable(); await w.write(texto); await w.close(); return 'gravado';
    }
    const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([texto], {type: 'text/javascript'})); a.download = 'perigos_riscos.js'; a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 5000);
    return 'baixado';
  }
  async function atualizarTabela() {
    let f;
    try { f = await escolherPlanilha(); } catch (e) { if (e.name !== 'AbortError') alert('Não consegui abrir o arquivo: ' + erroAmigavel(e)); return; }
    let linhas;
    try { linhas = await lerTabelaXlsx(f); } catch (e) { alert(`Não consegui ler "${f.name}": ` + erroAmigavel(e)); return; }
    let como;
    try { como = await gravarTabela(textoTabela(linhas, f.name)); } catch (e) { if (e.name !== 'AbortError') alert('Li a planilha, mas não consegui gravar o perigos_riscos.js: ' + erroAmigavel(e)); como = null; }
    window.TABELA_PERIGOS = linhas; montarTabela(); renderEditor(); renderPrevia();
    aviso(`Tabela atualizada a partir de "${f.name}": ${linhas.length} perigos.` + (como === 'baixado' ? ' Salvei o perigos_riscos.js nos downloads: copie-o para a pasta "recursos" do gerador.' : como ? '' : ' (vale só até fechar a página: o arquivo perigos_riscos.js não foi gravado)'));
  }
  // ---------------- alterações, rascunho, arquivo ----------------
  let tPrevia, tRascunho;
  function mudou(estrutura) {
    if (editandoPO) return;                // edição no pop-up do PO: só vale ao salvar
    if (!carregando && typeof registrarHist === 'function') registrarHist('apr', estrutura);
    if (estrutura) renderEditor();
    clearTimeout(tPrevia); tPrevia = setTimeout(() => { if (abaAtual === 'apr' || conf) renderPrevia(); }, 350);   // 07/11/2026: era 180
    clearTimeout(tRascunho); tRascunho = setTimeout(() => rascunho.salvar(JSON.stringify(apr), 'apr:rascunho'), 800);
    if (!carregando) { alterado = true; mostrarArquivo(); }
  }
  function mostrarArquivo() {
    if (!window.DOC_INICIAL) rascunho.salvar(alterado, 'apr:alterado');
    const s = $('#statusArquivo');
    if (!s || abaAtual !== 'apr') return;
    const estado = alterado ? ' – não salvo' : salvoAs ? ` – salvo ${salvoAs}` : '';
    // Excel feito fora do gerador abre como APR nova: a barra mostra de onde veio (07/11/2026)
    linhaStatus(s, nomeAtual || (nomeOrigem ? 'Aberta de ' + nomeOrigem : 'APR nova, sem arquivo'), estado, !!(nomeAtual || nomeOrigem));   // no HTML: "…" no meio
  }
  function definirArquivo(h, nome) {
    arquivo = h; nomeAtual = nome || ''; nomeOrigem = '';
    rascunho.salvar(h || null, 'apr:arquivo'); rascunho.salvar(nomeAtual, 'apr:nome');
  }
  function carregar(d) {
    carregando = true; apr = normalizar(d); conf = null;
    if (abaAtual === 'apr') { renderEditor(); renderPrevia(); }
    carregando = false; alterado = false; salvoAs = ''; mostrarArquivo();
    rascunho.salvar(JSON.stringify(apr), 'apr:rascunho');
    if (typeof zerarHist === 'function') zerarHist('apr');
  }
  // desfazer / refazer: volta a um estado guardado
  function restaurar(d) {
    apr = d;
    if (abaAtual === 'apr' || conf) { renderEditor(); renderPrevia(); }   // com a conferência aberta na aba PO, o PO também voltou
    rascunho.salvar(JSON.stringify(apr), 'apr:rascunho');
    alterado = true; mostrarArquivo();
  }
  async function nova() {
    if (alterado && !(await perguntar('Começar uma APR nova? As alterações não salvas serão perdidas.', {sim: 'COMEÇAR NOVA', titulo: 'APR NOVA'}))) return;
    definirArquivo(null, ''); carregar(novaAPR());
  }
  // "Salvar": o mesmo Excel do "Exportar Excel", com a APR inteira guardada dentro dele (guardarDados).
  // Só grava por cima de um .xlsx (um .apr.json antigo guardado no rascunho pede o local de novo);
  // APR nova: escolhe o local ANTES de montar o Excel (o navegador só abre a janela logo depois do clique).
  const TIPOS_XLSX = [{description: 'Planilha do Excel', accept: {'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx']}}];
  const arquivoExcel = () => arquivo && /\.xlsx$/i.test(arquivo.name || '') ? arquivo : null;
  async function salvar(comoNovo) {
    const sugerido = nomeAPR() + '.xlsx';
    const atual = arquivoExcel();
    try {
      let h = null;
      if (window.showSaveFilePicker && !(atual && !comoNovo))
        h = await window.showSaveFilePicker({id: 'salvar-apr', suggestedName: comoNovo && /\.xlsx$/i.test(nomeAtual) ? nomeAtual : sugerido, types: TIPOS_XLSX});
      aviso('Salvando a APR em Excel...');
      const blob = await gerarExcel();
      if (!window.showSaveFilePicker) {
        const a = el('a', {href: URL.createObjectURL(blob), download: sugerido});
        document.body.append(a); a.click(); a.remove();
        definirArquivo(null, sugerido); aviso('Salvo na pasta Downloads: ' + sugerido);
      } else if (!h) {
        await gravarArquivo(atual, blob); aviso('Salvo: ' + nomeAtual);
      } else {
        await gravarArquivo(h, blob); definirArquivo(h, h.name); aviso('Salvo: ' + h.name);
      }
      alterado = false; salvoAs = new Date().toLocaleTimeString('pt-BR').slice(0, 5); mostrarArquivo();
      return true;
    } catch (e) {
      if (e.name === 'SecurityError' || e.name === 'NotAllowedError') return 'clique';   // quem chamou pergunta de novo
      if (e.name !== 'AbortError') { console.error(e); alert('Não consegui salvar: ' + erroGravar(e)); }
      return false;
    }
  }
  async function abrir() {
    if (alterado && !(await perguntar('Abrir outra APR? As alterações não salvas serão perdidas.', {sim: 'ABRIR OUTRA', titulo: 'ABRIR APR'}))) return;
    let f, h = null;
    try {
      if (window.showOpenFilePicker) {
        [h] = await window.showOpenFilePicker({id: 'abrir-apr', types: [{description: 'APR em Excel (ou .apr.json antigo)', accept: {'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'], 'application/vnd.ms-excel': ['.xls'], 'application/json': ['.json']}}]});
        f = await h.getFile();
      } else f = await new Promise(ok => { const i = el('input', {type: 'file', accept: '.xlsx,.xls,.json', onchange: () => ok(i.files[0])}); i.click(); });
    } catch (e) { return; }
    if (f) await abrirArquivo(f, h, true);
  }
  // Excel salvo pelo gerador: abre com a APR guardada dentro dele e o "Salvar" grava nele de novo.
  // Excel salvo pelo gerador e depois mexido no Excel: pergunta qual versão usar. Outra APR em Excel (.xlsx/.xls)
  // passa pela leitura da APR antiga e abre como APR nova (sem arquivo). .apr.json antigo: abre só para passar para Excel.
  async function abrirArquivo(f, h, jaConfirmou) {
    if (!jaConfirmou && alterado && !(await perguntar('Abrir outra APR? As alterações não salvas da APR serão perdidas.', {sim: 'ABRIR OUTRA', titulo: 'ABRIR APR'}))) return;
    if (/\.xlsx?$/i.test(f.name)) {
      let g = null;
      try { g = await lerGuardado(await f.arrayBuffer()); } catch (e) { console.error(e); }
      if (abaAtual !== 'apr') trocarAba('apr');
      if (g && g.igual) { definirArquivo(h || null, f.name); carregar(g.dados); aviso('Aberto: ' + f.name); return; }
      if (g) {
        const i = await escolher(`"${f.name}" foi alterada direto no Excel depois de salva pelo gerador.\n\nUSAR O DO EXCEL: lê o que está na planilha agora (confira o resultado); as ligações do comparador continuam nas etapas e perigos que não mudaram.\n\nUSAR O DO GERADOR: abre como o gerador salvou da última vez; o que foi mudado no Excel fica de fora.`,
          ['USAR O DO EXCEL', 'USAR O DO GERADOR', 'CANCELAR'], {titulo: 'EDITADA NO EXCEL'});
        if (i < 0 || i === 2) return;
        if (i === 1) { definirArquivo(h || null, f.name); carregar(g.dados); aviso('Aberta como o gerador salvou: ' + f.name); return; }
      }
      aviso('Lendo a APR em Excel...');
      let d;
      try { d = await lerExcel(f); }
      catch (e) { console.error(e); alert('Não consegui ler essa APR: ' + erroAmigavel(e)); return; }
      if (g) { d = juntarGuardado(d, g.dados); definirArquivo(h || null, f.name); } else { definirArquivo(null, ''); nomeOrigem = f.name; }
      carregar(d); alterado = true; mostrarArquivo();
      const qt = `${d.etapas.length} etapas, ${d.etapas.reduce((s, e) => s + e.linhas.length, 0)} perigos`;
      aviso(g ? `APR lida da planilha: ${qt}. Confira e use "Salvar".` : `APR lida: ${qt}. Use "Salvar" para gravar a APR em Excel.`);
      return;
    }
    let d;
    try { d = JSON.parse(await f.text()); } catch (e) { alert('Esse arquivo não é uma APR do gerador.'); return; }
    abrirDados(d, null, f.name);
  }
  // .apr.json (formato antigo): abre como APR sem arquivo; o "Salvar" grava em Excel
  function abrirDados(d, h, nome) {
    if (!d || d.tipo !== 'apr' || !Array.isArray(d.etapas)) {
      alert(d && d.cab && Array.isArray(d.etapas) ? 'Esse arquivo é um PO. Abra na aba PO.' : 'Esse arquivo não é uma APR do gerador.');
      return;
    }
    if (abaAtual !== 'apr') trocarAba('apr');
    definirArquivo(null, ''); nomeOrigem = nome || ''; carregar(d); alterado = true; mostrarArquivo();
    aviso('Aberto: ' + nome + ' (formato antigo). Use "Salvar" para gravar a APR em Excel.');
  }
  async function iniciar() {
    if (window.DOC_INICIAL) return;
    const r = await rascunho.ler('apr:rascunho');
    if (r && (await rascunho.ler('apr:alterado')) === true) {
      let d = null; try { d = normalizar(JSON.parse(r)); } catch (e) { /* rascunho inválido */ }
      const nome = (await rascunho.ler('apr:nome')) || '';
      if (d && await perguntar(`A APR "${nome || 'APR.' + (d.cab.numero || 'XX.000') + ' - ' + (d.cab.atividade || 'Sem nome')}" ficou com alterações não salvas da última vez.`, {sim: 'RECUPERAR', nao: 'COMEÇAR NOVA', titulo: 'ALTERAÇÕES NÃO SALVAS'})) {
        apr = d; arquivo = await rascunho.ler('apr:arquivo'); nomeAtual = nome; alterado = true;
      } else rascunho.salvar(false, 'apr:alterado');
    }
    if (typeof zerarHist === 'function') zerarHist('apr');
    if (abaAtual === 'apr') { renderEditor(); renderPrevia(); }
  }

  // ---------------- PO aberto na aba PO ----------------
  const temPO = () => doc && doc.etapas.some(e => (e.nome || '').trim() || e.perigos.length);
  const rotPO = x => x.rotulo || (x.nome ? rotuloPadrao(x.nome) : '');
  const descricaoPO = e => String(e.descricao || '').split('\n').map(l => textoPuro(l)).filter(Boolean);
  async function garantirPO(acao) {
    if (temPO()) return true;
    if (!(await perguntar(`Para ${acao}, o PO precisa estar aberto na aba PO.`, {sim: 'ESCOLHER O PO', titulo: 'FALTA O PO'}))) return false;
    await abrirPO();
    return temPO();
  }
  async function puxarDoPO() {
    if (!(await garantirPO('puxar as etapas'))) return;
    const temDados = apr.etapas.some(e => e.linhas.some(l => l.perigo || l.risco || l.medidas));
    if (temDados && !(await perguntar(`Trocar as etapas da APR pelas do PO "${nomeArquivo()}"?\n\nOs dados da APR (risco, EPC, EPIs, PO/PD, medidas) dos perigos com o mesmo nome na mesma etapa são mantidos.`, {sim: 'TROCAR', titulo: 'IMPORTAR DO PO'}))) return;
    const c = apr.cab;
    if (!c.numero) c.numero = doc.cab.codigo || '';
    if (!c.atividade) c.atividade = doc.cab.nome || '';
    if (!c.episBasicos) c.episBasicos = doc.cab.episBasicos.map(b => b.rotulo || rotuloPadrao(b.nome)).join(', ');
    const antigas = apr.etapas;
    apr.etapas = doc.etapas.map((pe, i) => {
      const ae = antigas[i];
      const e = novaEtapa(pe.nome || '');
      if (ae) { e.id = ae.id; e.situacao = ae.situacao; }
      const pers = perigosPO(pe);
      e.linhas = pers.length ? pers.map(p => {
        const velha = ae && ae.linhas.find(l => N(l.perigo) === N(rotPO(p)));
        if (velha) return Object.assign({}, velha, {perigo: rotPO(p)});
        const l = novaLinha(rotPO(p)); perigoEscolhido(l); return l;
      }) : [novaLinha()];
      // etapa com um perigo só: os EPIs do PO valem para ele
      if (e.linhas.length === 1 && !e.linhas[0].epis) e.linhas[0].epis = episPO(pe).map(rotPO).join(', ');
      return e;
    });
    mudou(true);
    aviso(`Etapas puxadas do PO: ${apr.etapas.length} etapas, ${apr.etapas.reduce((s, e) => s + e.linhas.length, 0)} perigos.`);
  }
  // ---------------- botões da aba PO: Importar da APR, Comparar com APR, Pendências ----------------
  const temAPR = () => apr.etapas.some(e => (e.nome || '').trim() || e.linhas.some(l => l.perigo.trim()));
  async function garantirAPR(acao) {
    if (temAPR()) return true;
    if (!(await perguntar(`Para ${acao}, a APR precisa estar aberta na aba APR.`, {sim: 'ESCOLHER A APR', titulo: 'FALTA A APR'}))) return false;
    await abrir();
    return temAPR();
  }
  // pictograma da biblioteca com o mesmo nome do texto da APR ('' se não achar)
  function pictoDe(texto, lista) {
    let melhor = '', nota = 0.8;
    for (const p of lista) { const s = N(rotuloPadrao(p.nome)) === N(texto) ? 2 : semelhanca(rotuloPadrao(p.nome), texto); if (s > nota) { nota = s; melhor = p.nome; } }
    return melhor;
  }
  // traz da APR para o PO aberto: etapas que faltam, perigos e EPIs que faltam (nada que já está no PO é trocado)
  // ---- histórico de revisões entre APR e PO (05/11/2026, pedido do usuário) ----
  // copia item por item: Motivo <-> Descrição, Data e Responsável; os campos que só existem de um lado (Origem na APR;
  // Solicitante, Foto e Observações no PO) ficam como estão no item de mesmo número; na linha 0 o texto padrão de criação
  // troca de documento ("Criação do P.O." <-> "Criação da APR")
  const histUsado = x => !!((x.motivo || x.descricao || '').trim() || x.data || (x.responsavel || '').trim());
  // histórico de verdade (não só a linha 0 padrão sem data): serve de origem mesmo sem etapas
  const histReal = h => (h || []).some(x => x.data || (x.responsavel || '').trim()) || (h || []).filter(x => (x.motivo || x.descricao || '').trim()).length > 1;
  function textoHist(t, i, paraAPR) {
    t = t || '';
    if (i === 0 && /^\s*cria[çc][ãa]o\s+d[aoe]\s+(p\.?\s?o\.?|apr|procedimento.*|an[áa]lise.*)\s*$/i.test(t)) return paraAPR ? 'Criação da APR' : 'Criação do P.O.';
    return t;
  }
  async function histDoPO() {
    if (!temPO() && !histReal(doc.historico) && !(await garantirPO('importar o histórico de revisões'))) return;
    const src = doc.historico || [];
    if (!src.some(histUsado)) { aviso('O histórico de revisões do PO está vazio.'); return; }
    if (apr.historico.some(histUsado) && !(await perguntar(`Trocar o controle de revisão da APR pelo histórico de revisões do PO "${nomeArquivo()}"?\n\n${src.length} item(ns). A origem de cada item da APR é mantida.`, {sim: 'TROCAR', titulo: 'IMPORTAR DO PO'}))) return;
    const velho = apr.historico;
    apr.historico = src.map((p, i) => Object.assign({data: '', motivo: '', responsavel: '', origem: ''}, velho[i], {
      motivo: textoHist(p.descricao, i, true), data: p.data || '', responsavel: p.responsavel || ''}));
    mudou(true);
    aviso(`Controle de revisão importado do PO: ${apr.historico.length} item(ns).`);
  }
  async function histParaPO() {
    if (!temAPR() && !histReal(apr.historico) && !(await garantirAPR('importar o histórico de revisões'))) return;
    const src = apr.historico;
    if (!src.some(histUsado)) { aviso('O controle de revisão da APR está vazio.'); return; }
    if ((doc.historico || []).some(histUsado) && !(await perguntar(`Trocar o histórico de revisões do PO pelo controle de revisão da APR "${nomeAPR ? nomeAPR() : ''}"?\n\n${src.length} item(ns). Solicitante, foto e observações de cada item do PO são mantidos.`, {sim: 'TROCAR', titulo: 'IMPORTAR DA APR'}))) return;
    const velho = doc.historico || [];
    doc.historico = src.map((x, i) => Object.assign({descricao: '', foto: '', data: '', solicitante: '', responsavel: '', obs: ''}, velho[i], {
      descricao: textoHist(x.motivo, i, false), data: x.data || '', responsavel: x.responsavel || ''}));
    mudouPO(true);
    aviso(`Histórico de revisões importado da APR: ${doc.historico.length} item(ns).`);
  }
  async function importarParaPO() {
    if (!(await garantirAPR('importar'))) return;
    if (!(await perguntar(`Importar para o PO as etapas, os perigos e os EPIs da APR "${nomeAPR ? nomeAPR() : ''}"?\n\nSó entra o que falta no PO: o que já existe nele não é alterado.`, {sim: 'IMPORTAR', titulo: 'IMPORTAR DA APR'}))) return;
    const c = apr.cab;
    if (!doc.cab.codigo && c.numero) doc.cab.codigo = c.numero;
    if (!doc.cab.nome && c.atividade) doc.cab.nome = c.atividade;
    const semPicto = new Set();
    let novasE = 0, novosP = 0, novosE = 0;
    const lib = listaEpis();
    const juntar = (pe, ae) => {
      for (const l of ae.linhas) {
        const t = l.perigo.trim();
        // já existe no PO com o mesmo nome OU com o mesmo pictograma (ex.: "Exposição a condições climáticas adversas" e
        // "Exposição a Condições Climáticas"): não entra de novo (05/11/2026)
        const contem = (a, b) => { a = N(a); b = N(b); return a.length >= 10 && b.length >= 10 && (a.includes(b) || b.includes(a)); };
        if (t && !perigosPO(pe).some(p => N(rotPO(p)) === N(t) || contem(rotPO(p), t))) {
          const nome = pictoDe(t, BIB.perigos);
          if (!nome) semPicto.add(t);
          else if (!perigosPO(pe).some(p => p.nome === nome || (!p.nome && pictoDe(rotPO(p), BIB.perigos) === nome))) { pe.perigos.push({nome, rotulo: t}); novosP++; }
        }
        for (const t2 of dividirEpis(l.epis)) {
          if (episPO(pe).some(p => N(rotPO(p)) === N(t2))) continue;
          const nome = pictoDe(t2, lib);
          if (nome && episPO(pe).some(p => p.nome === nome || (!p.nome && pictoDe(rotPO(p), lib) === nome))) continue;
          if (nome) { pe.epis.push({nome, rotulo: t2}); novosE++; } else semPicto.add(t2);
        }
      }
    };
    const pares = alinharEtapas(apr.etapas, doc.etapas);
    const novas = [];
    for (const q of pares) {
      if (q.a < 0) continue;
      const ae = apr.etapas[q.a];
      let pe = q.p >= 0 ? doc.etapas[q.p] : null;
      if (!pe) { pe = window.novaEtapa(); pe.nome = ae.nome; novasE++; novas.push({pe, depois: pares.slice(0, pares.indexOf(q)).reverse().find(r => r.p >= 0)}); }
      juntar(pe, ae);
    }
    for (const {pe, depois} of novas) { const k = depois ? doc.etapas.indexOf(doc.etapas[depois.p]) + 1 : 0; doc.etapas.splice(k, 0, pe); }
    mudarPO();
    if (typeof renderEditor === 'function') renderEditor();
    const resumo = `Entraram no PO: ${novasE} etapa(s), ${novosP} perigo(s) e ${novosE} EPI(s).` + (semPicto.size ? `\n\nSem pictograma na biblioteca (não entraram): ${[...semPicto].join(', ')}.` : '');
    aviso(resumo.split('\n')[0]);
    alert(resumo);     // resumo numa janela, para não passar despercebido (05/11/2026)
  }
  // a conferência é uma só (a da aba APR) e abre igual nas duas abas, sem trocar de aba
  // ao fechar, avisa o que ficou sem salvar (a conferência corrige os dois) e oferece salvar (21/10/2026)
  async function fecharConf() {
    const poAlt = typeof poAlterado === 'function' && poAlterado(), aprAlt = alterado;
    if (poAlt || aprAlt) {
      const qual = poAlt && aprAlt ? 'A APR e o PO têm' : aprAlt ? 'A APR tem' : 'O PO tem';
      const rot = poAlt && aprAlt ? 'SALVAR APR E PO' : aprAlt ? 'SALVAR A APR' : 'SALVAR O PO';
      const i = await escolher(`${qual} alterações não salvas.`, [rot, 'FECHAR SEM SALVAR', 'CONTINUAR NA CONFERÊNCIA'], {titulo: 'FECHAR CONFERÊNCIA'});
      if (i < 0 || i === 2) return;
      if (i === 0) {
        // cada um pode abrir a janela "Salvar como" (documento novo); o navegador só abre uma por clique
        const salvarUm = async (fn, nome) => {
          let r = await fn();
          while (r === 'clique') {
            if (!(await perguntar(`Falta salvar ${nome}.`, {sim: nome === 'a APR' ? 'SALVAR A APR' : 'SALVAR O PO', titulo: 'SALVAR'}))) return false;
            r = await fn();
          }
          return r === true;
        };
        if (aprAlt && !(await salvarUm(() => salvar(false), 'a APR'))) return;
        if (poAlt && !(await salvarUm(() => salvarPO(false), 'o PO'))) return;
      }
    }
    const voltar = conf && conf.recolheu;
    conf = null; renderEditor(); renderPrevia();
    if (voltar && document.body.classList.contains('sem-editor')) colunaEditor(true);
  }
  function validarPO() {
    const L = [], c = doc.cab;
    for (const [k, r] of [['codigo', 'Código do PO'], ['nome', 'Nome da atividade'], ['elaborado', 'Elaborado por'], ['aprovado', 'Aprovado por']])
      if (!String(c[k] || '').trim()) L.push({i: -1, texto: `Falta preencher no cabeçalho: ${r}.`, campo: 'cab:' + k});
    // pictogramas "conferir" (imagem trazida do documento, não reconhecida) também são pendência (05/11/2026)
    const conferirB = (c.episBasicos || []).filter(b => !b.nome && b.img).length;
    if (conferirB) L.push({i: -1, texto: `${conferirB} EPI(s) básico(s) com pictograma para conferir.`, campo: 'cab:epis-basicos'});
    doc.etapas.forEach((e, i) => {
      const falta = [];
      if (!(e.nome || '').trim()) falta.push('nome');
      if (!descricaoPO(e).length) falta.push('descrição');
      if (!perigosPO(e).length) falta.push('perigos');
      const prim = {nome: 'nome', 'descrição': 'descricao', perigos: 'perigos'}[falta[0]];
      if (falta.length) L.push({i, nome: e.nome, texto: `Falta: ${falta.join(', ')}.`, campo: e.id + ':' + prim});
      for (const p of perigosPO(e)) if (!p.nome && !p.img) L.push({i, nome: e.nome, texto: `"${p.rotulo || 'Item'}" está sem pictograma.`, campo: e.id + ':perigos'});
      for (const p of episPO(e)) if (!p.nome && !p.img) L.push({i, nome: e.nome, texto: `"${p.rotulo || 'Item'}" está sem pictograma.`, campo: e.id + ':epis'});
      const confP = perigosPO(e).filter(p => !p.nome && p.img).length, confE = episPO(e).filter(p => !p.nome && p.img).length;
      if (confP) L.push({i, nome: e.nome, texto: `${confP} perigo(s) com pictograma para conferir.`, campo: e.id + ':perigos'});
      if (confE) L.push({i, nome: e.nome, texto: `${confE} EPI(s) com pictograma para conferir.`, campo: e.id + ':epis'});
      const fora = perigosPO(e).filter(p => p.nome || !p.img).map(rotPO).filter(t => t && !daTabela(t));   // os "conferir" já aparecem acima
      if (fora.length) L.push({i, nome: e.nome, texto: `Perigo fora da lista: ${fora.map(t => '"' + t + '"').join(', ')}.`, campo: e.id + ':perigos'});
    });
    return L;
  }
  // "IR PARA" das pendências do PO: abre a aba PO e vai ao campo (cabeçalho ou parte da etapa), como na APR
  function irParaPO(chave) {
    if (abaAtual !== 'po') trocarAba('po');
    setTimeout(() => {
      const [a, b] = String(chave).split(':');
      const card = a === 'cab' ? null : document.getElementById('ed-' + a);
      const alvo = a === 'cab' ? document.querySelector(`#editor [data-cab-ed~="${b}"]`) : card && (card.querySelector(`[data-campo="${b}"]`) || card);
      if (!alvo) return;
      const det = alvo.closest('details'); if (det) det.open = true;
      alvo.scrollIntoView({behavior: 'smooth', block: 'center'});
      alvo.classList.add('destaque'); setTimeout(() => alvo.classList.remove('destaque'), 1500);
      const foco = b === 'descricao' && card ? card.querySelector('.rich')
        : alvo.matches('input, textarea') ? alvo : alvo.querySelector('input[type=text], textarea');
      if (foco && b !== 'perigos' && b !== 'epis') setTimeout(() => foco.focus({preventScroll: true}), 350);
    }, 60);
  }
  // aoExportar: antes de exportar o PO (PDF, Excel, Word), a lista aparece com "Exportar mesmo assim" (21/10/2026)
  function pendenciasPO(aoExportar) {
    const m = $('#modal'), fechar = () => { m.hidden = true; m.replaceChildren(); };
    const L = validarPO(), grupos = new Map();
    for (const it of L) { if (!grupos.has(it.i)) grupos.set(it.i, []); grupos.get(it.i).push(it); }
    m.replaceChildren(el('div', {class: 'janela', onclick: ev => ev.stopPropagation()},
      el('header', {}, el('b', {}, aoExportar ? 'O que falta preencher no PO' : 'Pendências do PO'), el('span', {style: 'flex:1'}),
        aoExportar ? el('button', {class: 'btn primario', onclick: () => { fechar(); aoExportar(); }}, 'Exportar mesmo assim') : null,
        el('button', {class: 'btn', onclick: fechar}, 'Fechar')),
      el('div', {style: 'flex:1;overflow:auto;padding:12px'},
        L.length ? null : el('p', {class: 'apr-ok'}, '✓ Nada pendente no PO.'),
        [...grupos.values()].map(its => el('div', {class: 'apr-rel-grupo'},
          el('div', {class: 'apr-rel-tit'}, its[0].i < 0 ? 'Cabeçalho' : `Etapa ${its[0].i + 1}${its[0].nome ? ' – ' + its[0].nome : ''}`),
          its.map(it => el('div', {class: 'apr-rel-item'}, el('span', {}, it.texto),
            botao('IR PARA', () => { fechar(); irParaPO(it.campo); }))))))));
    m.onclick = fechar; m.hidden = false;
  }

  // ---------------- palavras (medidas de controle × frases do PO) ----------------
  const PARADAS = new Set(('a o e as os de da do das dos em no na nos nas ao aos com sem para pra por pelo pela pelos pelas um uma uns umas ou que se ' +
    'nao sua seu suas seus entre sobre ate apos antes durante quando como mais menos muito todo toda todos todas cada ser estar ter deve devem ' +
    'uso usar utilizar utilizacao local item itens etc').split(' '));
  const raiz = w => w.length > 5 ? w.slice(0, 5) : w.replace(/s$/, '');
  // calibrado com as APRs × POs da pasta "Treinamento do Gerador" (18/10/2026):
  // "Proibido fumar" = "Não fumar"; "subir" = "acesso"; "20 Km/h" = "20km/h"; números contam mesmo curtos
  const SINONIMOS = {proibido: 'nao', proibida: 'nao', proibidos: 'nao', proibicao: 'nao', proibir: 'nao', evitar: 'nao', jamais: 'nao', nunca: 'nao',
    subir: 'acesso', descer: 'acesso', subida: 'acesso', descida: 'acesso', acessar: 'acesso', acessos: 'acesso', checklist: 'checklist'};
  const preparar = t => N(t).replace(/\bcheck\s+list\b/g, 'checklist').replace(/\b(\d+)\s+(km|kg|m|cm|mm|h|v|kv|t|ton|min|bar|psi)\b/g, '$1$2');
  const palavras = t => preparar(t).split(' ').map(w => SINONIMOS[w] || w).filter(w => (w.length >= 3 || /^\d+$/.test(w)) && !PARADAS.has(w));
  // uma medida por linha (ou por ";"); linha que continua a anterior ("(Caminhão, trator...)" ou começando em
  // minúscula, sem marcador) volta a fazer parte dela
  const itensMedidas = txt => {
    const out = [];
    for (const bruto of String(txt || '').split(/\n|;/)) {
      const marcador = /^\s*[\-–—•*·]/.test(bruto);
      const s = bruto.replace(/^[\s\-–—•*·]+/, '').replace(/[\s.;,]+$/, '').trim();
      if (!s) continue;
      if (out.length && !marcador && (/^\(/.test(s) || /^[a-zà-ú]/.test(s))) out[out.length - 1] += ' ' + s;
      else out.push(s);
    }
    return out;
  };
  function corpoDe(texto) {
    const s = new Set();
    for (const w of palavras(texto)) { s.add(w); s.add(raiz(w)); if (w.length >= 5) s.add('p4:' + w.slice(0, 4)); }
    if (/\d+\s*km\s*h?\b/.test(preparar(texto))) { s.add('velocidade'); s.add(raiz('velocidade')); }   // "não ultrapassar 20km/h" fala de velocidade
    return s;
  }
  // a palavra aparece? (inteira, pelo radical de 5 letras ou, nas longas, de 4: limpeza ~ limpo, organização ~ organizado)
  const temPalavra = (corpo, w) => corpo.has(w) || corpo.has(raiz(w)) || (w.length >= 6 && corpo.has('p4:' + w.slice(0, 4)));
  function gruposMedida(m) {
    // o que está entre parênteses são exemplos: "Sinal sonoro (Buzina, sirene de ré)" vale se o PO citar a buzina
    const g = [m.replace(/\([^)]*\)/g, ' ')];
    for (const x of m.matchAll(/\(([^)]*)\)/g)) g.push(...x[1].split(/,|;|\bou\b|\be\b/));
    return g.map(t => [...new Set(palavras(t))]).filter(ws => ws.length);
  }
  // frases do PO numa etapa: as linhas da descrição (numeradas como aparecem), os recursos e os EPIs
  function frasesPO(pe, po) {
    po = po || doc;
    if (!pe) return [];
    const fr = [];
    String(pe.descricao || '').split('\n').forEach((l, k) => { const t = textoPuro(l); if (t) fr.push({linha: k, rot: String(fr.length + 1), t}); });
    // cada item da sublinha ("Caminhão; Guindaste") é um recurso; grupo = a sublinha inteira (ligações antigas)
    pe.sublinhas.forEach(s => {
      const its = String(s.valor || '').split(';').map(v => v.trim()).filter(Boolean);
      its.forEach(v => fr.push({linha: -1, rot: s.tipo, t: v, recurso: true, grupo: its.length > 1 ? N(s.valor) : ''}));
    });
    episPO(pe).forEach(x => fr.push({linha: -1, rot: 'EPI', t: rotPO(x), recurso: true}));       // "Cinto de segurança" como medida vale pelo EPI do PO
    // a etapa "Todas as atividades acima" do PO vale para todas as etapas (organização, EPIs gerais...)
    for (const te of (po && po.etapas) || []) {
      if (te === pe || !ehTodasAcima(te.nome)) continue;
      String(te.descricao || '').split('\n').forEach((l, k) => { const t = textoPuro(l); if (t) fr.push({linha: -1, rot: 'Todas ' + (k + 1), t, recurso: true}); });
      episPO(te).forEach(x => fr.push({linha: -1, rot: 'EPI (todas)', t: rotPO(x), recurso: true}));
    }
    // EPI's básicos do cabeçalho do PO valem em todas as etapas
    if (po && po.cab) (po.cab.episBasicos || []).forEach(b => { const t = b.rotulo || (b.nome ? rotuloPadrao(b.nome) : ''); if (t) fr.push({linha: -1, rot: 'EPI básico', t, recurso: true}); });
    for (const f of fr) { f.n = N(f.t); f.corpo = corpoDe(f.t); }
    return fr;
  }
  // quanto a frase cita a medida (0 a 1) e quais palavras dela batem
  // limiar calibrado com os 69 pares APR × PO do treino: de 0,72 para cima a frase sugerida era a certa em quase todos
  const AUTO_PESO = 0.72;
  // peso de cada palavra: as que aparecem em quase todas as frases do PO ("atividade", "área") valem menos
  function peso(frases, w) {
    const c = frases._idf || (frases._idf = new Map());
    if (!c.has(w)) { const df = frases.filter(f => temPalavra(f.corpo, w)).length; c.set(w, Math.log((1 + frases.length) / (1 + df)) + 1); }
    return c.get(w);
  }
  function nota(m, f, frases) {
    let melhor = {v: 0, vp: 0, ws: [], hits: 0};
    for (const ws of gruposMedida(m)) {
      const hits = ws.filter(w => temPalavra(f.corpo, w));
      const v = hits.length / ws.length;
      const tot = frases ? ws.reduce((s, w) => s + peso(frases, w), 0) : ws.length;
      const vp = frases ? hits.reduce((s, w) => s + peso(frases, w), 0) / (tot || 1) : v;
      if (vp > melhor.vp || (vp === melhor.vp && v > melhor.v)) melhor = {v, vp, ws, hits: hits.length};
    }
    return melhor;
  }
  // as 3 frases mais parecidas; empate: a frase mais curta (mais específica) primeiro
  const sugestoes = (m, frases) => frases.map(f => ({f, ...nota(m, f, frases)})).filter(s => s.v >= 0.3)
    .sort((a, b) => b.vp - a.vp || b.v - a.v || a.f.t.length - b.f.t.length).slice(0, 3);
  // as 3 frases da descrição mais parecidas (painel do comparador e atalhos 1-3; EPIs e recursos vêm à parte)
  const parecidas = (m, frases) => frases.filter(f => !f.recurso).map(f => ({f, ...nota(m, f, frases)})).filter(s => s.v >= 0.3)
    .sort((a, b) => b.vp - a.vp || b.v - a.v || a.f.t.length - b.f.t.length).slice(0, 3);
  // sublinhas do PO que aparecem como "recursos utilizados" no painel, com o nome mostrado
  const ROT_RECURSO = {Ferramenta: 'Ferramenta', Equipamento: 'Equipamento', 'Documentação': 'Documento', Material: 'Materiais'};
  const regraEpi = (m, l, pe) => { const nm = ' ' + N(m) + ' '; return / epis? /.test(nm) ? (dividirEpis(l.epis).length > 0 || (pe && episPO(pe).length > 0)) : / epcs? /.test(nm) ? dividirEpis(l.epc).length > 0 : null; };
  // situação de cada medida: ok (ligada a uma frase ou confirmada), auto (achada sozinha), nao (não está no PO) ou conferir
  function situacao(m, l, pe, frases) {
    const k = N(m), c = (l.conf || {})[k];
    if (c && c.estado === 'nao') return {estado: 'nao'};
    if (c && c.estado === 'ok') {
      if (!c.frase) return {estado: 'ok'};
      const f = frases.find(x => x.n === c.frase) || frases.find(x => x.grupo && x.grupo === c.frase);
      if (f) return {estado: 'ok', f};
    }
    if ((l.confirmadas || []).includes(k)) return {estado: 'ok'};
    const r = regraEpi(m, l, pe);
    if (r) return {estado: 'auto', regra: true};
    const s = sugestoes(m, frases)[0];
    // todas as palavras da medida na frase; ou quase todas, faltando só palavras comuns (medida de 3+ palavras)
    // (um número diferente nunca vale: "20 km/h" na APR e "10 km/h" no PO ficam para conferir)
    const numerosOk = x => x.ws.filter(w => /^\d/.test(w)).every(w => temPalavra(x.f.corpo, w));
    if (s && numerosOk(s) && (s.v >= 1 || (s.vp >= AUTO_PESO && s.hits >= 2 && s.ws.length >= 3))) return {estado: 'auto', f: s.f};
    return {estado: 'conferir'};
  }
  // troca um EPI escrito na linha da APR por outro texto (sem repetir)
  function trocarEpi(l, de, para) {
    l.epis = dividirEpis(l.epis).map(t => N(t) === N(de) ? para : t).filter((t, k, a) => a.findIndex(x => N(x) === N(t)) === k).join(', ');
  }
  function ligar(l, m, estado, f) {
    l.conf = l.conf || {};
    if (estado) l.conf[N(m)] = {estado, frase: f ? f.n : ''}; else delete l.conf[N(m)];
  }

  // ---------------- o que falta preencher ----------------
  function validar() {
    const L = [];
    const add = (ae, i, texto, campo) => L.push({ae, i, texto, campo});
    const c = apr.cab;
    for (const [k, r] of [['numero', 'Nº'], ['area', 'Área'], ['departamento', 'Departamento'], ['maquina', 'Máquina/Equipamento'], ['atividade', 'Atividade'], ['avaliado', 'Avaliado por']])
      if (!String(c[k] || '').trim()) add(null, -1, `Falta preencher no cabeçalho: ${r}.`, 'cab:' + k);
    apr.etapas.forEach((ae, i) => {
      if (!(ae.nome || '').trim()) add(ae, i, 'Etapa sem nome.', ae.id + ':nome');
      for (const l of ae.linhas) {
        const falta = [];
        if (!l.perigo.trim()) falta.push('perigo');
        if (!l.risco.trim()) falta.push('risco (efeito / dano)');
        if (!numero(l.po)) falta.push('PO');
        if (!numero(l.pd)) falta.push('PD');
        if (!l.medidas.trim()) falta.push('medidas de controle');
        if (falta.length) add(ae, i, `${l.perigo ? `"${l.perigo}"` : 'Perigo sem nome'}: falta ${falta.join(', ')}.`, l.id + ':' + (l.perigo.trim() ? 'risco' : 'perigo'));
        // perigo fora da lista de Perigos e Riscos (07/11/2026)
        if (l.perigo.trim() && !daTabela(l.perigo)) add(ae, i, `"${l.perigo}": perigo fora da lista. Escolha o perigo da lista.`, l.id + ':perigo');
      }
    });
    return L;
  }
  const botao = (t, f, tit, cls = 'btn mini') => el('button', {class: cls, title: tit || '', onclick: f}, t);
  function relatorio(aoExportar) {
    const m = $('#modal');
    const fechar = () => { m.hidden = true; m.replaceChildren(); };
    const L = validar();
    const grupos = new Map();
    for (const it of L) { const k = it.i < 0 ? 'cab' : 'e' + it.i; if (!grupos.has(k)) grupos.set(k, []); grupos.get(k).push(it); }
    m.replaceChildren(el('div', {class: 'janela', onclick: ev => ev.stopPropagation()},
      el('header', {}, el('b', {}, 'O que falta preencher na APR'), el('span', {style: 'flex:1'}),
        aoExportar ? el('button', {class: 'btn primario', onclick: () => { fechar(); aoExportar(); }}, 'Exportar mesmo assim') : null,
        el('button', {class: 'btn', onclick: fechar}, 'Fechar')),
      el('div', {style: 'flex:1;overflow:auto;padding:12px'},
        L.length ? null : el('p', {class: 'apr-ok'}, '✓ Todos os campos da APR estão preenchidos.'),
        [...grupos.values()].map(its => el('div', {class: 'apr-rel-grupo'},
          el('div', {class: 'apr-rel-tit'}, its[0].i < 0 ? 'Cabeçalho' : `Etapa ${its[0].i + 1}${its[0].ae.nome ? ' – ' + its[0].ae.nome : ''}`),
          its.map(it => el('div', {class: 'apr-rel-item'}, el('span', {}, it.texto), botao('IR PARA', () => { fechar(); irPara(it.campo); }))))))));
    m.onclick = fechar; m.hidden = false;
  }
  function irPara(chave) {
    if (abaAtual !== 'apr') trocarAba('apr');
    setTimeout(() => {
      const alvo = document.querySelector(`#editor [data-apr-ed="${chave}"]`) || document.getElementById('apr-' + String(chave).split(':')[0]);
      if (!alvo) return;
      const det = alvo.closest('details'); if (det) det.open = true;
      alvo.scrollIntoView({behavior: 'smooth', block: 'center'});
      alvo.classList.add('destaque'); setTimeout(() => alvo.classList.remove('destaque'), 1500);
      const foco = alvo.matches('input, textarea, select') ? alvo : alvo.querySelector('input[type=text], textarea, select');
      if (foco) setTimeout(() => foco.focus({preventScroll: true}), 350);
    }, 60);
  }

  // ---------------- conferência APR × PO ----------------
  // Tela no lugar da prévia: etapa por etapa, perigos / EPIs / nome lado a lado (corrigindo a APR ou o PO)
  // e as medidas de controle uma a uma, com as frases do PO mais parecidas ao lado.
  let conf = null;                 // {i: etapa, sel: medida selecionada, tudo: descrição inteira aberta}
  let editandoPO = false;          // pop-up com a descrição do PO aberto: as edições ainda não valem
  function semelhanca(a, b) {
    const bi = s => { const t = N(s).replace(/ /g, ''), r = []; for (let i = 0; i < t.length - 1; i++) r.push(t.slice(i, i + 2)); return r; };
    const x = bi(a), y = bi(b);
    if (!x.length || !y.length) return 0;
    const mp = new Map(); y.forEach(g => mp.set(g, (mp.get(g) || 0) + 1));
    let c = 0; x.forEach(g => { if (mp.get(g)) { c++; mp.set(g, mp.get(g) - 1); } });
    return 2 * c / (x.length + y.length);
  }
  // casa os itens da APR com os do PO: iguais primeiro, depois os parecidos (diferença de escrita)
  function parear(aprs, pos) {
    const linhas = [], livresP = [...pos];
    const resto = [];
    for (const a of aprs) { const k = livresP.findIndex(p => N(p.t) === N(a.t)); if (k >= 0) { linhas.push({a, p: livresP[k], tipo: 'igual'}); livresP.splice(k, 1); } else resto.push(a); }
    for (const a of resto) {
      let melhor = -1, nota = 0.45;
      livresP.forEach((p, k) => { const s = semelhanca(a.t, p.t); if (s > nota) { nota = s; melhor = k; } });
      if (melhor >= 0) { linhas.push({a, p: livresP[melhor], tipo: 'dif'}); livresP.splice(melhor, 1); }
      else linhas.push({a, p: null, tipo: 'soAPR'});
    }
    for (const p of livresP) linhas.push({a: null, p, tipo: 'soPO'});
    return linhas;
  }
  // casa as etapas da APR com as do PO pelo nome (a APR pode ter etapas a mais, a menos ou em outra ordem):
  // devolve [{a, p}] com o índice na APR e no PO (-1 = só de um lado), na ordem da APR. Primeiro os nomes mais
  // parecidos; depois, uma etapa que sobrou fica com a que sobrou do PO na mesma posição relativa (vizinhas já
  // casadas), que é o caso de etapa renomeada.
  function alinharEtapas(aes, pes) {
    const na = aes.length, np = pes.length, MIN = 0.5;
    // a.poNome: etapa do PO escolhida à mão na conferência (vale mais que qualquer nome parecido)
    // palavras repetidas em muitas etapas ("Manutenção no ...") não contam na semelhança
    const df = new Map(), nomes = [...aes, ...pes].map(e => N(e.nome));
    for (const s of nomes) for (const w of new Set(s.split(' '))) df.set(w, (df.get(w) || 0) + 1);
    const lim = Math.max(3, 0.3 * nomes.length);
    const miolo = s => { const r = s.split(' ').filter(w => w.length > 2 && !PARADAS.has(w) && df.get(w) <= lim).join(' '); return r || s; };
    const sim = aes.map(a => pes.map(p => {
      if (a.poNome) return N(p.nome) === N(a.poNome) ? 3 : 0;
      const x = N(a.nome), y = N(p.nome);
      if (!x || !y) return 0;
      if (x === y) return 1;
      // radicais em comum ("Basculamento da caçamba" ~ "Bascular o material") e letras em comum (erros de digitação)
      const rz = s => new Set(miolo(s).split(' ').filter(w => w.length > 2 && !PARADAS.has(w)).map(w => w.slice(0, 5)));
      const ra = rz(x), rb = rz(y), com = [...ra].filter(w => rb.has(w)).length;
      const dice = ra.size + rb.size ? 2 * com / (ra.size + rb.size) : 0;
      return Math.max(dice, 0.9 * semelhanca(miolo(x), miolo(y)));
    }));
    // a mesma posição relativa desempata nomes parecidos com mais de uma etapa
    const perto = (i, j) => 0.15 * Math.abs((i + 0.5) / na - (j + 0.5) / np);
    const pA = new Array(na).fill(-1), pP = new Array(np).fill(-1);
    const cands = [];
    for (let i = 0; i < na; i++) for (let j = 0; j < np; j++) if (sim[i][j] >= MIN) cands.push([sim[i][j] - perto(i, j), i, j]);
    cands.sort((x, y) => y[0] - x[0]);
    for (const [, i, j] of cands) if (pA[i] < 0 && pP[j] < 0) { pA[i] = j; pP[j] = i; }
    // depois, entre as que sobraram, nomes menos parecidos ("Transitar entre as áreas" ~ "Trânsito pela área")
    const c2 = [];
    for (let i = 0; i < na; i++) for (let j = 0; j < np; j++) if (pA[i] < 0 && pP[j] < 0 && !aes[i].poNome && sim[i][j] >= 0.35) c2.push([sim[i][j] - perto(i, j), i, j]);
    c2.sort((x, y) => y[0] - x[0]);
    for (const [, i, j] of c2) if (pA[i] < 0 && pP[j] < 0) { pA[i] = j; pP[j] = i; }
    // sobras na mesma posição relativa: vizinha de cima casada com a vizinha de cima do PO (ou começo/fim)
    for (let rodada = 0; rodada < 2; rodada++) for (let i = 0; i < na; i++) {
      if (pA[i] >= 0 || aes[i].poNome) continue;
      const cand = [];
      for (let j = 0; j < np; j++) {
        if (pP[j] >= 0) continue;
        const cima = i === 0 ? j === 0 : pA[i - 1] >= 0 && pA[i - 1] === j - 1;
        const baixo = i === na - 1 ? j === np - 1 : pA[i + 1] >= 0 && pA[i + 1] === j + 1;
        if ((cima && baixo) || ((cima || baixo) && sim[i][j] >= 0.2)) cand.push(j);    // entre duas casadas: é ela, mesmo renomeada
      }
      if (cand.length === 1) { pA[i] = cand[0]; pP[cand[0]] = i; }
    }
    // etapa do PO que junta várias da APR (o nome do PO lista "Carregar... Cortar... Descarregar..."):
    // a etapa da APR que sobrou e cujo nome está quase todo dentro do nome de uma etapa do PO usa essa etapa
    const raizes = s => N(s).split(' ').filter(w => w.length > 2 && !PARADAS.has(w)).map(w => w.slice(0, 5));
    for (let i = 0; i < na; i++) {
      if (pA[i] >= 0 || aes[i].poNome) continue;
      const ra = raizes(aes[i].nome);
      if (ra.length < 2) continue;
      let melhor = -1, nota = 0.8;
      for (let j = 0; j < np; j++) {
        const rb = new Set(raizes(pes[j].nome));
        if (rb.size < ra.length + 2) continue;
        const v = ra.filter(w => rb.has(w)).length / ra.length;
        if (v >= nota) { nota = v; melhor = j; }
      }
      if (melhor >= 0) { pA[i] = melhor; if (pP[melhor] < 0) pP[melhor] = i; }
    }
    // ordem: a da APR; a etapa só do PO entra logo depois da etapa casada com a etapa anterior do PO
    const out = [];
    const soPO = j => out.push({a: -1, p: j});
    for (let j = 0; j < np && pP[j] < 0; j++) soPO(j);
    for (let i = 0; i < na; i++) {
      out.push({a: i, p: pA[i]});
      if (pA[i] >= 0) for (let j = pA[i] + 1; j < np && pP[j] < 0 && !out.some(q => q.p === j); j++) soPO(j);
    }
    for (let j = 0; j < np; j++) if (pP[j] < 0 && !out.some(q => q.p === j)) soPO(j);
    return out;
  }
  const ehTodasAcima = nome => /todas?\s+as\s+a+tividades|todas?\s+as\s+etapas/.test(N(nome));
  const paresEtapas = () => alinharEtapas(apr.etapas, doc.etapas);
  function medidasDaEtapa(ae) {
    const r = [];
    for (const l of ae.linhas) for (const m of itensMedidas(l.medidas)) r.push({l, m});
    return r;
  }
  function progresso() {
    let feitas = 0, total = 0;
    for (const par of paresEtapas()) {
      const ae = apr.etapas[par.a], pe = doc.etapas[par.p], fr = frasesPO(pe);
      if (ae) for (const {l, m} of medidasDaEtapa(ae)) { total++; if (situacao(m, l, pe, fr).estado !== 'conferir') feitas++; }
    }
    return {feitas, total};
  }
  // abre (ou fecha) a conferência na aba em que se está: a mesma tela, o mesmo estado e o mesmo desfazer nas duas abas
  async function abrirConferencia() {
    if (conf) { fecharConf(); return; }
    const aba = abaAtual;
    if (!(await garantirAPR('comparar')) || !(await garantirPO('comparar'))) return;
    if (abaAtual !== aba) trocarAba(aba);                       // abrir a APR ou o PO agora não muda a aba
    // códigos diferentes (ex.: PO.AMP.009 x APR 107): avisa antes de abrir (05/11/2026)
    const cod = s => String(s || '').toUpperCase().replace(/^(APR|PO)\b[\s.\-–]*/, '').replace(/[^A-Z0-9]/g, '');
    const ca = cod(apr.cab.numero), cp = cod(doc.cab.codigo);
    if (ca && cp && ca !== cp && !(await perguntar(`A APR (${apr.cab.numero}) e o PO (${doc.cab.codigo}) têm códigos diferentes. Comparar assim mesmo?`, {sim: 'COMPARAR ASSIM MESMO', nao: 'CANCELAR', titulo: 'CÓDIGOS DIFERENTES'}))) return;
    conf = {i: 0, sel: 0, tudo: false};
    if (typeof zerarHist === 'function') zerarHist('conf');     // desfazer do comparador: APR e PO juntos
    renderEditor(); renderPrevia();
    // a conferência precisa da largura toda: o painel de edição recolhe sozinho e volta ao fechar (28/10/2026)
    if (typeof colunaEditor === 'function' && !document.body.classList.contains('sem-editor')) { conf.recolheu = true; colunaEditor(false); }
  }
  const compararDoPO = abrirConferencia;                        // botão "Comparar com APR" da aba PO
  const mudarPO = () => { if (typeof mudouPO === 'function') mudouPO(true); renderPrevia(); };
  function destacarFrase(t, ws) {
    const frag = document.createDocumentFragment(), alvo = corpoDe(ws.join(' '));
    let ult = 0;
    for (const x of t.matchAll(/[\p{L}\p{N}]+/gu)) {
      const w = N(x[0]);
      if (w.length >= 3 && !PARADAS.has(w) && ws.some(m => raiz(m) === raiz(w) || (m.length >= 6 && w.length >= 5 && m.slice(0, 4) === w.slice(0, 4)) || m === w) && alvo.size) {
        frag.append(t.slice(ult, x.index), el('mark', {}, x[0])); ult = x.index + x[0].length;
      }
    }
    frag.append(t.slice(ult));
    return frag;
  }
  // escolher à mão qual etapa do PO corresponde à etapa da APR (quando os nomes são bem diferentes)
  const selPO = (a, rot, atual = -1) => el('label', {class: 'cf-par'}, rot + ' ',
    el('select', {onchange: ev => {
      const v = ev.target.value;
      if (v === '') delete a.poNome; else a.poNome = doc.etapas[+v].nome;
      mudou(true);
      const k = paresEtapas().findIndex(q => q.a === apr.etapas.indexOf(a)); if (k >= 0) conf.i = k;
      conf.trocar = false; renderPrevia();
    }}, el('option', {value: ''}, a.poNome || atual < 0 ? '(AUTOMÁTICO, PELO NOME)' : 'ESCOLHA...'),
      doc.etapas.map((p, k) => el('option', {value: k, selected: a.poNome ? N(p.nome) === N(a.poNome) : k === atual}, `PO ${k + 1} – ${p.nome || '(sem nome)'}`))));
  const selAPR = p => el('label', {class: 'cf-par'}, 'Comparar com a etapa da APR: ',
    el('select', {onchange: ev => {
      if (ev.target.value === '') return;
      const a = apr.etapas[+ev.target.value]; a.poNome = p.nome; mudou(true);
      const k = paresEtapas().findIndex(q => q.a === +ev.target.value); if (k >= 0) conf.i = k;
      renderPrevia();
    }}, el('option', {value: ''}, 'ESCOLHA...'), apr.etapas.map((a, k) => el('option', {value: k}, `APR ${k + 1} – ${a.nome || '(sem nome)'}`))));
  // etapa que só existe de um lado: criar do outro ou escolher a etapa correspondente
  function semPar(par, pares, ae, pe) {
    const i = pares.indexOf(par);
    return el('div', {class: 'cf-card'}, !ae
      ? [el('p', {}, `A etapa ${par.p + 1} do PO ("${pe.nome}") não tem etapa parecida na APR.`), botao('CRIAR A ETAPA NA APR', () => {
          const e = novaEtapa(pe.nome); const ps = perigosPO(pe); if (ps.length) e.linhas = ps.map(p => linhaDe(rotPO(p)));
          const ant = pares.slice(0, i).reverse().find(q => q.a >= 0); apr.etapas.splice(ant ? ant.a + 1 : 0, 0, e); mudou(true); }, 'Cria a etapa na APR, na mesma posição do PO', 'btn'),
        el('div', {style: 'margin-top:8px'}, 'Ou, se é a mesma etapa com outro nome: ', selAPR(pe))]
      : [el('p', {}, `A etapa ${par.a + 1} da APR ("${ae.nome}") não tem etapa parecida no PO.`), botao('Criar a etapa no PO', () => {
          const e = window.novaEtapa(); e.nome = ae.nome;
          const ant = pares.slice(0, i).reverse().find(q => q.p >= 0); doc.etapas.splice(ant ? ant.p + 1 : 0, 0, e); mudarPO(); }, 'Cria a etapa no PO, na mesma posição da APR', 'btn'),
        el('div', {style: 'margin-top:8px'}, 'Ou, se é a mesma etapa com outro nome: ', selPO(ae, 'Comparar com a etapa do PO:'))]);
  }

  // ---------------- comparador da aba APR (4ª versão, 19/10/2026, modelo aprovado pelo usuário) ----------------
  // Cabeçalho com a faixa de etapas; grade "Etapa, perigos e EPIs" (APR | situação / corrigir | PO); embaixo, as
  // medidas da APR (M1, M2...) ao lado da descrição do PO frase por frase. As conferências continuam salvas como
  // antes: linha.conf[N(medida)] = {estado, frase} e etapa.poDispensa (frases do PO que não precisam de medida).
  const curto = (s, n) => { s = String(s || '').replace(/\s+/g, ' ').trim(); return s.length > n ? s.slice(0, n - 1).replace(/\s+\S*$/, '') + '…' : s; };
  const maiuscula = s => s.charAt(0).toUpperCase() + s.slice(1);
  const chavePar = q => q.a + '|' + q.p;
  const nomeDoPar = q => (apr.etapas[q.a] || doc.etapas[q.p] || {}).nome || '(sem nome)';
  const numPar = q => q.a >= 0 ? String(q.a + 1) : 'PO ' + (q.p + 1);
  // etapas da APR comparadas com a etapa p do PO (uma etapa do PO pode valer para várias da APR)
  const aprsDoPO = (p, pares) => pares.filter(q => q.p === p && q.a >= 0).map(q => apr.etapas[q.a]);
  function linhasGrade(ae, pe) {
    const L = [{tipo: 'etapa', rot: 'Etapa', sit: N(ae.nome) === N(pe.nome) ? 'igual' : 'dif', a: {t: ae.nome}, p: {t: pe.nome}}];
    const perAPR = [...new Map(ae.linhas.filter(l => l.perigo.trim()).map(l => [N(l.perigo), {t: l.perigo}])).values()];
    for (const r of parear(perAPR, perigosPO(pe).map(p => ({t: rotPO(p), item: p})))) L.push({tipo: 'perigo', rot: 'Perigo', sit: r.tipo, a: r.a, p: r.p});
    const epiAPR = [...new Map(ae.linhas.flatMap(l => dividirEpis(l.epis)).map(t => [N(t), {t}])).values()];
    for (const r of parear(epiAPR, episPO(pe).map(x => ({t: rotPO(x), item: x})))) L.push({tipo: 'epi', rot: 'EPI', sit: r.tipo, a: r.a, p: r.p});
    // associações feitas com ASSOCIAR (05/11/2026): a linha "só na APR" e a "só no PO" viram uma linha divergente
    // (na posição da primeira); associação que não vale mais (nome já corrigido, item tirado) é esquecida
    const k = chaveAssoc(ae, pe);
    if (conf && conf.assoc && conf.assoc[k]) conf.assoc[k] = conf.assoc[k].filter(q => {
      const ia = L.findIndex(x => x.tipo === 'perigo' && x.sit === 'soAPR' && N(x.a.t) === q.a);
      const ip = L.findIndex(x => x.tipo === 'perigo' && x.sit === 'soPO' && N(x.p.t) === q.p);
      if (ia < 0 || ip < 0) return false;
      const i = Math.min(ia, ip), j = Math.max(ia, ip);
      const m = {tipo: 'perigo', rot: L[ia].rot, sit: 'dif', a: L[ia].a, p: L[ip].p, assoc: q};
      L.splice(j, 1); L[i] = m;
      return true;
    });
    return L;
  }
  const chaveAssoc = (ae, pe) => (ae.id || apr.etapas.indexOf(ae)) + '|' + doc.etapas.indexOf(pe);
  // frase da descrição do PO -> medidas (de todas as etapas da APR comparadas com esta etapa do PO) que ela explica
  function ligacoes(pe, aes, fr) {
    const mp = new Map();
    for (const ae of aes) medidasDaEtapa(ae).forEach(({l, m}, k) => {
      const s = situacao(m, l, pe, fr);
      if ((s.estado === 'ok' || s.estado === 'auto') && s.f) { if (!mp.has(s.f.n)) mp.set(s.f.n, []); mp.get(s.f.n).push({ae, l, m, k}); }
    });
    return mp;
  }
  const dispensada = (aes, f) => aes.some(ae => (ae.poDispensa || []).includes(f.n));
  function resumoEtapa(par, pares) {
    const ae = apr.etapas[par.a], pe = doc.etapas[par.p];
    if (!ae || !pe) return {ok: 0, pend: 1};
    let ok = 0, pend = 0;
    const conta = b => { if (b) ok++; else pend++; };
    linhasGrade(ae, pe).forEach(x => conta(x.sit === 'igual'));
    // só as medidas da APR contam (as frases do PO sem medida na APR não são mais "a decidir")
    const fr = frasesPO(pe);
    medidasDaEtapa(ae).forEach(({l, m}) => conta(situacao(m, l, pe, fr).estado !== 'conferir'));
    return {ok, pend};
  }
  // onde a medida está no PO e a cor do chip (verde = o PO explica, vermelho = não explica, cinza = vale pelos
  // EPIs da etapa ou não aplicável no PO)
  function ondeMedida(s) {
    if (s.estado === 'conferir') return {cor: 'mr', onde: '—'};
    if (s.estado === 'nao') return {cor: 'mc', onde: 'Não aplicável no PO'};
    if (s.regra) return {cor: 'mc', onde: 'EPI / EPC da Etapa'};
    if (!s.f) return {cor: 'mv', onde: 'confirmada'};
    const r = s.f.rot;
    // EPI e recurso: só "EPI" / "Recurso"; o nome do item fica na dica (title)
    // "Frase n": passando o mouse aparece a frase do PO (05/11/2026)
    return {cor: 'mv', dica: !s.f.recurso || /^Todas /.test(r) ? 'Frase do PO: ' + s.f.t :/^EPI/.test(r) ? 'EPI: ' + s.f.t : (ROT_RECURSO[r] || r) + ': ' + s.f.t, onde: !s.f.recurso ? 'Frase ' + r : /^EPI/.test(r) ? 'EPI'
      : /^Todas /.test(r) ? 'Todas as atividades: Frase ' + r.slice(6) : 'Recurso'};
  }
  // lista do ASSOCIAR (05/11/2026): janelinha presa à tela (a grade tem overflow:hidden e cortaria a lista)
  let assocDe = null;
  const foraAssoc = ev => { const p = document.getElementById('cf2-apop'); if (p && !p.contains(ev.target) && !(assocDe && assocDe.contains(ev.target))) fecharAssoc(); };
  const escAssoc = ev => { if (ev.key === 'Escape') { ev.stopPropagation(); ev.preventDefault(); fecharAssoc(); } };
  function fecharAssoc() {
    const p = document.getElementById('cf2-apop'); if (p) p.remove();
    assocDe = null;
    document.removeEventListener('mousedown', foraAssoc, true); document.removeEventListener('keydown', escAssoc, true);
    window.removeEventListener('resize', fecharAssoc); const pv = $('#previa'); if (pv) pv.removeEventListener('scroll', fecharAssoc);
  }
  function irEtapa2(i) {
    if (i < 0 || i >= paresEtapas().length) return;
    conf.i = i; conf.real = null; conf.incluir = null; conf.escolha = null; conf.trocar = false; conf.tudo = false;
    renderPrevia(); $('#previa').scrollTop = 0;
  }
  function desenharComparador() {
    fecharAssoc();
    const pares = paresEtapas(), n = pares.length;
    if (conf.i >= n) conf.i = Math.max(0, n - 1);
    conf.vistas = conf.vistas || new Set(); conf.corr = conf.corr || {}; conf.tirados = conf.tirados || {};
    const par = pares[conf.i] || {a: -1, p: -1}, cp = chavePar(par);
    conf.vistas.add(cp);
    const ae = apr.etapas[par.a], pe = doc.etapas[par.p];
    const res = resumoEtapa(par, pares);
    const rotPar = q => q.a >= 0 && q.p >= 0 ? (q.a === q.p ? `Etapa ${q.a + 1}` : `APR ${q.a + 1} / PO ${q.p + 1}`) : q.a >= 0 ? `APR ${q.a + 1} (só na APR)` : `PO ${q.p + 1} (só no PO)`;
    const topo = el('div', {class: 'cf-topo cf2-topo'},
      el('div', {class: 'cf2-nav'},
        el('span', {class: 'cf2-titulo', title: 'Setas ← → do teclado mudam a etapa'}, 'CONFERÊNCIA APR X PO'),
        el('span', {style: 'flex:1'}),
        el('span', {class: 'cf2-cont'}, `Etapa ${numPar(par)}: `, el('b', {class: 'ok'}, `${res.ok} conferidos`), ' · ', el('b', {class: res.pend ? 'pend' : 'ok'}, `${res.pend} a decidir`)),
        botao('Fechar conferência', fecharConf, '', 'btn cf2-fechar')),
      (() => {
        // faixa com 5 etapas por vez: as setas ‹ › mostram as 5 anteriores / seguintes (sem trocar a etapa aberta);
        // ao mudar de etapa, a faixa vai para o grupo da etapa atual
        const POR = 5, maxIni = Math.max(0, n - POR);
        if (conf.ini == null || (conf.iniDe !== conf.i && (conf.i < conf.ini || conf.i >= conf.ini + POR))) conf.ini = Math.floor(conf.i / POR) * POR;
        conf.iniDe = conf.i;
        conf.ini = Math.max(0, Math.min(conf.ini, maxIni));
        const seta = (t, ok, d, tit) => el('button', {class: 'cf2-et cf2-seta', disabled: !ok, title: tit, onclick: () => { conf.ini = Math.max(0, Math.min(conf.ini + d, maxIni)); renderPrevia(); }}, t);
        return el('div', {class: 'cf2-faixa', style: `grid-template-columns:36px repeat(${Math.max(1, Math.min(n, POR))},minmax(0,1fr)) 36px`},
          seta('‹', conf.ini > 0, -POR, 'Etapas anteriores'),
          pares.slice(conf.ini, conf.ini + POR).map((q, j) => {
            const k = conf.ini + j, r = k === conf.i ? res : resumoEtapa(q, pares), visto = conf.vistas.has(chavePar(q));
            const cor = visto && !r.pend ? 'ok' : 'nv';          // verde = aberta e tudo resolvido; cinza = o resto
            return el('button', {class: 'cf2-et ' + cor + (k === conf.i ? ' atual' : ''), onclick: () => irEtapa2(k),
              title: `${rotPar(q)} – ${nomeDoPar(q)}\n` + (!visto ? 'Ainda não visitada' : r.pend ? `${r.pend} a decidir` : 'Tudo resolvido')},
              el('span', {class: 't'}, el('span', {}, `${numPar(q)} · ${nomeDoPar(q)}`)), el('i'));
          }),
          seta('›', conf.ini < maxIni, POR, 'Próximas etapas'));
      })());
    const corpo = el('div', {class: 'cf-corpo'});
    const mostrar = () => { $('#paginas').replaceChildren(el('div', {class: 'apr-conf cf2'}, topo, corpo)); };
    if (!ae || !pe) { corpo.append(semPar(par, pares, ae, pe)); mostrar(); return; }

    // ---- etapa, perigos e EPIs ----
    // botões que corrigem a APR (borda verde-oliva) e os que corrigem o PO (borda azul)
    // texto em maiúsculo pelo CSS; "<" e ">" na altura das maiúsculas
    const rotB = t => { const m = /^(<\s)?(.*?)(\s>)?$/.exec(t); return [m[1] ? el('span', {class: 'cf2-sa'}, '<') : null, m[1] ? ' ' : null, m[2], m[3] ? ' ' : null, m[3] ? el('span', {class: 'cf2-sa'}, '>') : null].filter(x => x != null); };
    const bA = (t, f, tit) => el('button', {class: 'cf2-b a', title: tit || '', onclick: f}, ...rotB(t)), bP = (t, f, tit) => el('button', {class: 'cf2-b p', title: tit || '', onclick: f}, ...rotB(t));
    const ca = cabAPR();
    const grade = el('div', {class: 'cf2-grade'},
      el('div', {class: 'cf2-gt'}),
      el('div', {class: 'cf2-gp'}, el('span', {class: 'cf2-pil apr'}, `APR.${ca.numero}`)),
      el('div', {class: 'cf2-gp meio'}, el('label', {class: 'cf2-so'}, el('input', {type: 'checkbox', checked: !!conf.so, onchange: ev => { conf.so = ev.target.checked; renderPrevia(); }}), 'mostrar apenas divergências')),
      el('div', {class: 'cf2-gp'}, el('span', {class: 'cf2-pil po'}, `PO.${(doc.cab.codigo || 'XX.000').trim()}`)));
    const corrigido = (x, final, rot) => { conf.corr[cp + '|' + x.tipo + '|' + (x.tipo === 'etapa' ? '' : N(final))] = rot; };
    const tirado = (x, lado) => { (conf.tirados[cp] = conf.tirados[cp] || []).push({tipo: x.tipo, rot: x.rot, lado, t: (lado === 'po' ? x.p : x.a).t}); };
    const linhasDoPerigo = t => ae.linhas.filter(l => N(l.perigo) === N(t));
    const linhasDoEpi = t => ae.linhas.filter(l => dividirEpis(l.epis).some(e => N(e) === N(t)));
    // células da APR e do PO: perigo = campo que abre a lista padrão de perigos (escolherPerigo); etapa e EPI =
    // texto editável na própria grade. O texto vale ao sair do campo (ou Enter); enquanto digita, só a etiqueta
    // do meio acompanha (igual / diferente), para o campo não perder o foco.
    const fixarPar = () => { if (!ae.poNome) ae.poNome = pe.nome; };      // editar o nome não pode descasar as etapas
    const gravarTexto = (x, lado, v) => {
      const velho = (lado === 'apr' ? x.a : x.p).t;
      if (!v || v === velho) return false;
      if (x.tipo === 'etapa') {
        if (lado === 'apr') { fixarPar(); ae.nome = v; mudou(true); }
        else {
          apr.etapas.forEach(a => { if (a.poNome && N(a.poNome) === N(pe.nome)) a.poNome = v; });
          fixarPar(); ae.poNome = v; pe.nome = v; mudou(true); mudarPO();
        }
      } else if (lado === 'apr') { linhasDoEpi(velho).forEach(l => trocarEpi(l, velho, v)); mudou(true); }
      else { x.p.item.rotulo = v; mudarPO(); }
      return true;
    };
    const campoTexto = (x, lado, tag) => {
      const it = lado === 'apr' ? x.a : x.p, outro = lado === 'apr' ? x.p : x.a;
      const etapa = x.tipo === 'etapa';
      const c = el(etapa ? 'textarea' : 'input', etapa ? {class: 'cf2-ed', rows: 2, spellcheck: 'true'} : {class: 'cf2-ed', type: 'text', spellcheck: 'true'});
      c.value = it.t;
      // nome da etapa em uma linha só: centralizado na altura do campo (textarea não centraliza sozinho)
      const centrar = () => {
        if (!etapa || !c.isConnected) return;
        c.style.paddingTop = c.style.paddingBottom = '';
        const cs = getComputedStyle(c), pad = parseFloat(cs.paddingTop) + parseFloat(cs.paddingBottom), lh = parseFloat(cs.lineHeight);
        const alt = c.clientHeight;
        c.style.height = '0'; const texto = c.scrollHeight - pad; c.style.height = '';     // altura do texto (o scrollHeight nunca é menor que o campo)
        if (texto < lh * 1.5) c.style.paddingTop = c.style.paddingBottom = Math.max(0, (alt - lh) / 2) + 'px';
      };
      if (etapa) { setTimeout(centrar, 0); c.addEventListener('input', centrar); }
      let feito = false;
      const fim = () => {
        if (feito || !c.isConnected) return;                        // a grade foi refeita por outro motivo
        const v = c.value.replace(/\s+/g, ' ').trim();
        if (v === it.t) return;
        feito = true;
        if (!gravarTexto(x, lado, v)) { c.value = it.t; renderPrevia(); }
      };
      c.addEventListener('input', () => {
        if (!outro || !tag) return;
        const igual = N(c.value) === N(outro.t);
        tag.className = 'cf2-tag ' + (igual ? 'igual' : 'dif'); tag.textContent = igual ? '✓ informações compatíveis' : 'informações divergentes';
      });
      c.addEventListener('keydown', ev => {
        if (ev.key === 'Enter') { ev.preventDefault(); fim(); c.blur(); }
        else if (ev.key === 'Escape') { c.value = it.t; c.dispatchEvent(new Event('input')); c.blur(); }
      });
      c.addEventListener('focus', () => { feito = false; });
      c.addEventListener('blur', fim);
      return c;
    };
    // pictograma do perigo no PO: a mesma troca do editor do PO (trocarPicto), mas marcando o PO como alterado
    const trocarPictoPO = item => {
      const atual = {img: imgPerigo(item), nome: item.rotulo || (item.nome ? rotuloPadrao(item.nome) : 'Imagem trazida do documento, sem nome')};
      abrirSeletor('Trocar perigo', BIB.perigos, nome => {
        if ('rotulo' in item && (!item.rotulo || item.rotulo === rotuloPadrao(item.nome))) item.rotulo = rotuloPadrao(nome);
        if (item.img && typeof APRENDIDOS !== 'undefined') APRENDIDOS.aprender('perigo', item.img, nome);
        item.nome = nome; delete item.img;
        mudarPO();
      }, atual);
    };
    const campoPerigo = (x, lado) => {
      const it = lado === 'apr' ? x.a : x.p, tab = daTabela(it.t);
      if (lado === 'po') {
        // no PO: só o nome do perigo e, à direita, o pictograma (clique nele troca o pictograma)
        const src = imgPerigo(it.item);
        return el('div', {class: 'cf2-psel cf2-ppo', title: it.t + '\nClique para escolher da lista de perigos', onclick: () =>
          escolherPerigo('Trocar o perigo no PO (o pictograma continua o mesmo)', it.t, row => { it.item.rotulo = row.perigo; mudarPO(); })},
          // "FORA DA LISTA" na 2ª linha, embaixo do nome (05/11/2026, pedido do usuário); o campo não muda de tamanho
          el('span', {class: 'nm2' + (tab ? '' : ' com-fora')}, el('span', {class: 'nm'}, it.t), tab ? null : el('span', {class: 'cf2-fora'}, 'fora da lista')),
          el('button', {class: 'cf2-pic', title: 'Trocar o pictograma', onclick: ev => { ev.stopPropagation(); trocarPictoPO(it.item); }},
            src ? el('img', {src, alt: ''}) : el('span', {}, 'sem pictograma')),
          el('span', {class: 'cf2-tri'}, '▼'));
      }
      const risco = tab ? tab.risco : lado === 'apr' ? (linhasDoPerigo(it.t)[0] || {}).risco : '';
      return el('button', {class: 'cf2-psel', title: it.t + (risco ? '\nRisco: ' + risco : '') + '\nClique para escolher da lista de perigos', onclick: () => lado === 'apr'
        ? escolherPerigo('Trocar o perigo na APR', it.t, row => { linhasDoPerigo(it.t).forEach(l => aplicarPerigo(l, row)); mudou(true); })
        : escolherPerigo('Trocar o perigo no PO (o pictograma continua o mesmo)', it.t, row => { it.item.rotulo = row.perigo; mudarPO(); })},
        el('span', {class: 'l1'}, el('span', {class: 'nm'}, it.t)),
        el('span', {class: 'l2'}, el('span', {class: 'rs'}, 'Risco: ' + (risco || '—')), tab ? null : el('span', {class: 'cf2-fora'}, 'fora da lista')),
        el('span', {class: 'cf2-tri'}, '▼'));
    };
    // ASSOCIAR (05/11/2026, modelo aprovado pelo usuário): no quadro vazio de um perigo que só está de um lado.
    // Abre a lista dos perigos sem par do outro lado nesta etapa; escolhido um, as duas linhas viram uma só
    // ("informações divergentes") e USAR DO PO / USAR DA APR decidem o nome. A associação vale enquanto a
    // conferência está aberta (conf.assoc), como as demais marcações da grade.
    const semParOutro = x => linhas.filter(y => y.tipo === 'perigo' && y.sit === (x.sit === 'soAPR' ? 'soPO' : 'soAPR'));
    const abrirAssociar = (x, bt) => {
      fecharAssoc();
      const lado = x.sit === 'soAPR' ? 'po' : 'apr', cands = semParOutro(x);
      const pop = el('div', {class: 'cf2-apop', id: 'cf2-apop'},
        el('div', {class: 'cf2-apt'}, lado === 'po' ? 'Associar com perigo do PO' : 'Associar com perigo da APR'),
        cands.length ? cands.map(y => {
          const it = lado === 'po' ? y.p : y.a, src = lado === 'po' ? imgPerigo(it.item) : null;
          return el('button', {class: 'cf2-aop', title: it.t, onclick: () => {
            fecharAssoc();
            conf.assoc = conf.assoc || {}; const ka = chaveAssoc(ae, pe);
            (conf.assoc[ka] = conf.assoc[ka] || []).push({a: N((x.a || y.a).t), p: N((x.p || y.p).t)});
            renderPrevia();
          }}, src ? el('img', {src, alt: ''}) : null, el('span', {}, it.t));
        }) : el('div', {class: 'cf2-av'}, 'Nenhum perigo sem par nesta etapa'),
        el('button', {class: 'cf2-b t cf2-acan', onclick: fecharAssoc}, 'Cancelar'));
      document.body.append(pop);
      const r = bt.getBoundingClientRect(), w = Math.max(r.width, 280);
      pop.style.left = Math.max(8, Math.min(r.left, innerWidth - w - 8)) + 'px'; pop.style.width = w + 'px'; pop.style.top = (r.bottom + 4) + 'px';
      const h = pop.offsetHeight; if (r.bottom + 4 + h > innerHeight - 8) pop.style.top = Math.max(8, r.top - 4 - h) + 'px';
      assocDe = bt;
      document.addEventListener('mousedown', foraAssoc, true); document.addEventListener('keydown', escAssoc, true);
      window.addEventListener('resize', fecharAssoc); const pv = $('#previa'); if (pv) pv.addEventListener('scroll', fecharAssoc);
      const o = pop.querySelector('.cf2-aop'); if (o) o.focus();
    };
    const vazioAssoc = (lado, x) => {
      const bt = el('button', {class: 'cf2-b t cf2-assoc', title: `Associar com um perigo ${lado === 'po' ? 'do PO' : 'da APR'} que está com outro nome`,
        onclick: ev => { ev.stopPropagation(); if (assocDe === bt) fecharAssoc(); else abrirAssociar(x, bt); }}, 'Associar');
      return el('div', {class: 'cf2-vazio cf2-vazio2'}, el('span', {}, lado === 'apr' ? 'Não consta na APR' : 'Não consta no PO'), bt);
    };
    const celula = (lado, it, cls, x, tag) => el('div', {class: 'cf2-c ' + cls}, !it
      ? (x.tipo === 'perigo' ? vazioAssoc(lado, x) : el('div', {class: 'cf2-vazio'}, lado === 'apr' ? 'Não consta na APR' : 'Não consta no PO'))
      : x.tipo === 'perigo' ? campoPerigo(x, lado) : campoTexto(x, lado, tag));
    const etiqueta = (cls, t) => el('span', {class: 'cf2-tag ' + cls}, t);
    const linhas = linhasGrade(ae, pe);
    let escondidos = 0;
    for (const x of linhas) {
      const chaveC = cp + '|' + x.tipo + '|' + (x.tipo === 'etapa' ? '' : N((x.a || x.p).t));
      if (x.sit === 'igual' && !conf.corr[chaveC] && conf.so) { escondidos++; continue; }
      let meio;
      if (x.sit === 'igual') meio = [conf.corr[chaveC] ? etiqueta('feito', '✓ ' + conf.corr[chaveC]) : etiqueta('igual', '✓ informações compatíveis')];
      else if (x.sit === 'dif') { meio = [etiqueta('dif', 'informações divergentes'), el('div', {class: 'cf2-bts'},
        bA('< Usar do PO', () => {
          if (x.tipo === 'etapa') ae.nome = pe.nome;
          else if (x.tipo === 'perigo') linhasDoPerigo(x.a.t).forEach(l => { l.perigo = x.p.t; });
          else linhasDoEpi(x.a.t).forEach(l => trocarEpi(l, x.a.t, x.p.t));
          corrigido(x, x.p.t, 'APR corrigida'); mudou(true);
        }, 'Corrige a APR com o texto do PO'),
        bP('Usar da APR >', () => {
          if (x.tipo === 'etapa') pe.nome = ae.nome; else x.p.item.rotulo = x.a.t;
          corrigido(x, x.a.t, 'PO corrigido'); mudarPO();
        }, 'Corrige o PO com o texto da APR'))];
      if (x.assoc) {
        // nome da APR sugerido quando ele é da lista de perigos e o do PO não
        if (daTabela(x.a.t) && !daTabela(x.p.t)) {
          const b = meio[1].lastChild; b.classList.add('cf2-sug'); b.title = 'Sugerido: o nome da APR está na lista de perigos. ' + b.title;
          meio.push(el('div', {class: 'cf2-shint'}, 'Sugerido: nome da APR (está na lista de perigos)'));
        }
        meio.push(botao('Desfazer associação', () => { const ka = chaveAssoc(ae, pe); conf.assoc[ka] = (conf.assoc[ka] || []).filter(q => q !== x.assoc); renderPrevia(); }, 'Volta as duas linhas separadas', 'cf2-b t cf2-troca'));
      }
      }
      else if (x.sit === 'soPO') {
        const kEsc = chaveC;
        const incluir = l => { l.epis = [...dividirEpis(l.epis), x.p.t].join(', '); conf.escolha = null; corrigido(x, x.p.t, 'APR corrigida'); mudou(true); };
        meio = [etiqueta('so', 'consta apenas no PO'), el('div', {class: 'cf2-bts'},
          bA('< Incluir na APR', () => {
            if (x.tipo === 'perigo') {
              const vazia = ae.linhas.length === 1 && !ae.linhas[0].perigo.trim() && !ae.linhas[0].medidas.trim() ? ae.linhas[0] : null;
              if (vazia) { const nl = linhaDe(x.p.t); Object.assign(vazia, {perigo: nl.perigo, risco: nl.risco, pd: nl.pd}); } else ae.linhas.push(linhaDe(x.p.t));
              corrigido(x, x.p.t, 'APR corrigida'); mudou(true);
            } else if (ae.linhas.length <= 1) incluir(ae.linhas[0]);
            else { conf.escolha = conf.escolha === kEsc ? null : kEsc; renderPrevia(); }
          }, x.tipo === 'epi' ? 'Escreve o EPI na APR (escolha o perigo)' : 'Cria a linha deste perigo na APR'),
          bP('Tirar do PO >', async () => {
            const lista = x.tipo === 'perigo' ? pe.perigos : pe.epis;
            if (!(await perguntar(`Tirar ${x.tipo === 'perigo' ? 'o perigo' : 'o EPI'} "${x.p.t}" do PO?`, {sim: 'TIRAR DO PO', titulo: 'TIRAR DO PO'}))) return;
            lista.splice(lista.indexOf(x.p.item), 1); tirado(x, 'po'); mudarPO();
          }, 'Apaga este item do PO')),
          conf.escolha === kEsc ? el('select', {class: 'apr-sel cf2-esc', onchange: ev => { const l = ae.linhas[+ev.target.value]; if (l) incluir(l); }},
            el('option', {value: ''}, 'No perigo da APR...'), ae.linhas.map((l, k) => el('option', {value: k}, l.perigo || `(perigo ${k + 1})`))) : null];
      } else meio = [etiqueta('so', 'consta apenas na APR'), el('div', {class: 'cf2-bts'},
        bA('< Tirar da APR', async () => {
          if (x.tipo === 'perigo') {
            if (!(await perguntar(`Tirar "${x.a.t}" da APR (com os dados de risco e medidas)?`, {sim: 'TIRAR DA APR', titulo: 'TIRAR DA APR'}))) return;
            const la = linhasDoPerigo(x.a.t); ae.linhas = ae.linhas.filter(l => !la.includes(l)); if (!ae.linhas.length) ae.linhas.push(novaLinha());
          } else linhasDoEpi(x.a.t).forEach(l => { l.epis = dividirEpis(l.epis).filter(t => N(t) !== N(x.a.t)).join(', '); });
          tirado(x, 'apr'); mudou(true);
        }, 'Apaga este item da APR'),
        bP('Incluir no PO >', () => abrirSeletor(`Pictograma para "${x.a.t}"`, x.tipo === 'perigo' ? BIB.perigos : listaEpis(), nome => {
          (x.tipo === 'perigo' ? pe.perigos : pe.epis).push({nome, rotulo: x.a.t}); corrigido(x, x.a.t, 'PO corrigido'); mudarPO();
        }, null, x.a.t), 'Escolhe o pictograma e inclui no PO'))];
      // casamento errado das etapas: "trocar etapa do PO" abre o seletor (só na linha "Etapa")
      if (x.tipo === 'etapa') meio.push(conf.trocar
        ? selPO(ae, '', par.p)
        : botao('Trocar etapa do PO', () => { conf.trocar = true; renderPrevia(); }, 'Comparar esta etapa da APR com outra etapa do PO', 'cf2-b t cf2-troca'));
      const cls = x.sit === 'dif' ? 'cf2-r-dif' : x.sit === 'igual' ? '' : 'cf2-r-so';
      const tag = x.sit === 'igual' || x.sit === 'dif' ? meio[0] : null;
      grade.append(el('div', {class: 'cf2-rot ' + cls}, x.rot), celula('apr', x.a, cls, x, tag), el('div', {class: 'cf2-meio ' + cls}, meio), celula('po', x.p, cls, x, tag));
    }
    for (const t of conf.tirados[cp] || []) {
      const sumiu = el('div', {class: 'cf2-c'}, el('span', {class: 'tx cf2-riscado'}, t.t));
      grade.append(el('div', {class: 'cf2-rot'}, t.rot), t.lado === 'apr' ? sumiu : el('div', {class: 'cf2-c'}, el('div', {class: 'cf2-vazio'}, 'Não consta na APR')),
        el('div', {class: 'cf2-meio'}, etiqueta('feito', t.lado === 'po' ? '✓ PO corrigido' : '✓ APR corrigida')),
        t.lado === 'po' ? sumiu : el('div', {class: 'cf2-c'}, el('div', {class: 'cf2-vazio'}, 'Não consta no PO')));
    }
    if (escondidos) grade.append(el('div', {class: 'cf2-escondidos'}, `✓ ${escondidos} ${escondidos === 1 ? 'item igual escondido' : 'itens iguais escondidos'}`));
    corpo.append(el('div', {class: 'cf-card cf2-card'}, el('div', {class: 'cf2-secao'}, 'ETAPA, PERIGOS E EPIS'), grade));

    // ---- medidas de controle da APR (19/10/2026: só a APR; a descrição do PO aparece no painel da medida) ----
    const fr = frasesPO(pe), ms = medidasDaEtapa(ae);
    const sits = ms.map(({l, m}) => situacao(m, l, pe, fr));
    const real = conf.real || {};
    const abrir = (k, ev) => { if (ev) ev.stopPropagation(); conf.real = real.m === k ? null : {m: k}; conf.tudo = false; renderPrevia(); };
    const tagM = (k, cor) => el('span', {class: 'cf2-m ' + cor}, 'M' + (k + 1));
    const bAz = (t, f, tit, cls = 'a') => el('button', {class: 'cf2-pb ' + cls, title: tit || '', onclick: f}, t);
    const naoAplica = (l, m, k, s) => {
      if (s.estado === 'nao') ligar(l, m, null); else { ligar(l, m, 'nao'); avancar(ae, pe, k); }
      mudou(); renderPrevia();
    };
    // painel da medida clicada: as frases do PO mais parecidas (ou a descrição completa), os EPIs e os recursos
    // do PO na etapa, cada um com VINCULAR; embaixo, os 4 botões
    function painel(l, m, s, k) {
      const ws = gruposMedida(m).flat();
      const vinc = f => {
        const lig = s.f === f, auto = lig && s.estado === 'auto';
        return el('button', {class: 'cf2-vb' + (lig && !auto ? ' lig' : ''), title: lig && !auto ? 'Clique para desfazer o vínculo' : 'Vincula a medida a este item do PO',
          onclick: () => { if (lig && !auto) ligar(l, m, null); else { ligar(l, m, 'ok', f); avancar(ae, pe, k); } mudou(); renderPrevia(); }},
          auto ? 'Confirmar vínculo' : lig ? '✓ Vinculada' : 'Vincular');
      };
      const item = (n, t, f) => el('div', {class: 'cf2-sf' + (s.f === f ? ' ligada' : '')}, el('span', {class: 'n'}, n), el('span', {class: 'tx'}, t), vinc(f));
      const descr = fr.filter(f => !f.recurso), sug = parecidas(m, fr).map(x => x.f);
      if (s.f && !s.f.recurso && !sug.includes(s.f)) sug.push(s.f);
      const lista = conf.tudo ? descr : sug;
      const epis = fr.filter(f => f.rot === 'EPI'), recs = fr.filter((f, j) => f.recurso && ROT_RECURSO[f.rot] && fr.findIndex(x => x.rot === f.rot && x.n === f.n) === j);
      return el('div', {class: 'cf2-painel', onclick: ev => ev.stopPropagation()},
        el('div', {class: 'cf2-pt'}, conf.tudo ? 'Descrição completa' : 'Frases do PO mais parecidas'),
        lista.map(f => item(f.rot + '.', destacarFrase(f.t, ws), f)),
        lista.length ? null : el('div', {class: 'cf2-pv'}, conf.tudo ? 'Esta etapa do PO não tem descrição.' : 'Nenhuma frase parecida nesta etapa do PO.'),
        el('div', {class: 'cf2-pt sub'}, 'EPIs'),
        epis.map(f => item('', f.t, f)),
        epis.length ? null : el('div', {class: 'cf2-pv'}, 'Nenhum EPI nesta etapa do PO.'),
        el('div', {class: 'cf2-pt sub'}, 'Recursos utilizados'),
        recs.map(f => item('', ROT_RECURSO[f.rot] + ': ' + f.t, f)),
        recs.length ? null : el('div', {class: 'cf2-pv'}, 'Nenhum recurso nesta etapa do PO.'),
        el('div', {class: 'cf2-p4'},
          bAz('Não aplicável no PO', () => naoAplica(l, m, k, s), s.estado === 'nao' ? 'Já marcada: clique para desfazer' : 'Marca a medida como não aplicável no PO', 'c' + (s.estado === 'nao' ? ' on' : '')),
          bAz('Editar texto no PO', () => editarDescricao(pe, {l, m, k}), 'Abre a descrição do PO com a frase sugerida no fim'),
          bAz(conf.tudo ? 'Fechar descrição' : 'Descrição completa', () => { conf.tudo = !conf.tudo; renderPrevia(); }),
          bAz('Fechar', () => { conf.real = null; conf.tudo = false; renderPrevia(); }, '', 'c')));
    }
    // texto da medida editável na própria linha: grava no Enter ou ao sair, Esc desfaz
    const campoMedida = (l, m, k) => {
      const c = el('input', {class: 'cf2-mi', type: 'text', spellcheck: 'true', 'data-k': k});
      c.value = m;
      let feito = false;
      const fim = () => {
        if (feito || !c.isConnected) return;
        const v = c.value.replace(/\s+/g, ' ').trim();
        if (!v || v === m) { c.value = m; return; }
        feito = true;
        if (trocarMedida(l, m, v)) { mudou(true); renderPrevia(); } else { c.value = m; feito = false; }
      };
      c.addEventListener('click', ev => {
        ev.stopPropagation();
        if (real.m !== k) { const p = c.selectionStart; conf.real = {m: k}; conf.tudo = false; conf.foco = {k, p}; renderPrevia(); }
      });
      c.addEventListener('keydown', ev => {
        if (ev.key === 'Enter') { ev.preventDefault(); fim(); c.blur(); }
        else if (ev.key === 'Escape') { ev.stopPropagation(); c.value = m; c.blur(); }
      });
      c.addEventListener('focus', () => { feito = false; });
      c.addEventListener('blur', fim);
      return c;
    };
    const quadro = el('div', {class: 'cf2-col cf2-quadro'}, el('div', {class: 'cf2-ch apr cf2-chc'}, 'MEDIDAS DE CONTROLE'));
    let ultL = null;
    ms.forEach(({l, m}, k) => {
      if (l !== ultL) { quadro.append(el('div', {class: 'cf2-per'}, l.perigo || '(perigo sem nome)')); ultL = l; }
      const s = sits[k], o = ondeMedida(s);
      quadro.append(el('div', {class: 'cf2-med' + (real.m === k ? ' real' : ''), title: s.estado === 'auto' ? 'Achada automaticamente' : '', onclick: ev => abrir(k, ev)},
        tagM(k, o.cor), campoMedida(l, m, k), el('span', {class: 'onde', title: o.dica || ''}, o.onde)));
      if (real.m === k) quadro.append(painel(l, m, s, k));
    });
    if (!ms.length) quadro.append(el('p', {class: 'dica'}, 'Esta etapa da APR não tem medidas de controle.'));
    quadro.append(el('div', {class: 'cf2-leg'},
      el('span', {}, el('span', {class: 'cf2-m mv'}, 'M'), ' o PO explica'),
      el('span', {}, el('span', {class: 'cf2-m mr'}, 'M'), ' o PO não explica'),
      el('span', {}, el('span', {class: 'cf2-m mc'}, 'M'), ' vale pelos EPIs ou Não aplicável no PO')));
    const pend = ms.map((x, k) => ({...x, k})).filter(x => sits[x.k].estado === 'conferir');
    quadro.append(pend.length
      ? el('div', {class: 'cf2-pend'}, el('div', {class: 'cf2-pend-t'}, `Medidas da APR que este texto ainda não explica (${pend.length})`),
        pend.map(({l, m, k}) => el('div', {class: 'cf2-pend-i' + (real.m === k ? ' real' : '')},
          el('div', {class: 'cf2-pend-tx', onclick: ev => abrir(k, ev)}, tagM(k, 'mr'), el('span', {}, m, el('span', {class: 'cf2-pper'}, ' – ' + (l.perigo || '')))),
          bAz('Editar texto no PO', () => editarDescricao(pe, {l, m, k}), 'Abre a descrição do PO com a frase sugerida no fim'),
          bAz('Não aplicável no PO', () => { ligar(l, m, 'nao'); mudou(); renderPrevia(); }, 'Marca a medida como não aplicável no PO', 'c'))))
      : el('div', {class: 'cf2-pend ok'}, '✓ Todas as medidas desta etapa estão explicadas no PO.'));
    corpo.append(quadro,
      el('div', {class: 'cf2-dica'}, 'Atalhos: ← → mudam a etapa; ↑ ↓ escolhem a medida; 1 a 3 vinculam a frase parecida; N = não aplicável no PO; Esc fecha o painel.'));
    mostrar();
    if (conf.foco) {
      const c = $(`#paginas .cf2-mi[data-k="${conf.foco.k}"]`), p = conf.foco.p;
      conf.foco = null;
      if (c) { c.focus({preventScroll: true}); try { c.setSelectionRange(p, p); } catch (e) {} }
    }
    // a medida aberta fica presa logo abaixo do cabeçalho da conferência (que também fica preso no topo)
    const r = $('#paginas .cf2-med.real');
    if (r) r.style.top = (topo.offsetHeight + (parseFloat(getComputedStyle(topo).top) || 0)) + 'px';
    if (r && conf.rolar) r.scrollIntoView({block: 'nearest'});
    conf.rolar = false;
  }
  // depois de decidir uma medida, vai para a próxima ainda sem frase do PO (como no comparador anterior)
  function avancar(ae, pe, k) {
    const fr = frasesPO(pe), ms = medidasDaEtapa(ae);
    const j = ms.findIndex((x, i) => i > k && situacao(x.m, x.l, pe, fr).estado === 'conferir');
    conf.real = j >= 0 ? {m: j} : null; conf.tudo = false; conf.rolar = true;
  }
  // troca o texto de uma medida na linha da APR (a conferência dela vai junto); false se não achou o trecho
  function trocarMedida(l, m, v) {
    let k = l.medidas.indexOf(m), tam = m.length;
    if (k < 0) {                       // medida que juntou duas linhas: acha o trecho com a quebra de linha no meio
      const r = new RegExp(m.split(/\s+/).map(x => x.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('[\\s;]+')).exec(l.medidas);
      if (!r) return false;
      k = r.index; tam = r[0].length;
    }
    l.medidas = l.medidas.slice(0, k) + v + l.medidas.slice(k + tam);
    if (l.conf && l.conf[N(m)]) { l.conf[N(v)] = l.conf[N(m)]; delete l.conf[N(m)]; }
    if ((l.confirmadas || []).includes(N(m))) l.confirmadas = l.confirmadas.map(c => c === N(m) ? N(v) : c);
    return true;
  }
  // janela "Editar descrição do PO": a descrição inteira da etapa no mesmo editor da aba PO. Aberta por "Escrever
  // frase no PO" (alvo = medida), a frase sugerida já vem no fim e a frase nova/alterada fica ligada à medida.
  // Frases inalteradas mantêm as ligações; frase alterada leva as ligações da frase antiga mais parecida.
  // "Editar no PO" de uma frase sugerida (linha >= 0): não insere nada, o cursor vai para o fim dessa frase
  function editarDescricao(pe, alvo, linha = -1) {
    const md = $('#modal'), ip = doc.etapas.indexOf(pe);
    const tmp = Object.assign({}, pe);
    editandoPO = true;
    const fechar = () => { editandoPO = false; md.hidden = true; md.replaceChildren(); };
    const ed = editorTexto(tmp), caixa = ed.querySelector('.rich');
    const antes = String(pe.descricao || '').split('\n').map(x => N(textoPuro(x))).filter(Boolean), antesS = new Set(antes);
    let nova = null;
    // a frase nova entra como último item (termina em "."); o item que era o último passa a terminar em ";" (05/11/2026)
    const sugerida = alvo && linha < 0 ? maiuscula(alvo.m.replace(/[\s.;,]+$/, '')) + '.' : '';
    if (sugerida) {
      nova = document.createElement('div'); nova.textContent = sugerida;
      const ult = caixa.lastElementChild;
      const cheios = [...caixa.children].filter(d => d.textContent.trim());
      const ant = cheios[cheios.length - 1];
      if (ant) {
        const w = document.createTreeWalker(ant, NodeFilter.SHOW_TEXT); let tn = null;
        while (w.nextNode()) if (w.currentNode.textContent.trim()) tn = w.currentNode;
        if (tn && /\.\s*$/.test(tn.textContent)) tn.textContent = tn.textContent.replace(/\.(\s*)$/, ';$1');
      }
      if (ult && !ult.textContent.trim()) ult.replaceWith(nova); else caixa.append(nova);
      caixa.dispatchEvent(new Event('input'));
    }
    const salvar = () => {
      if (String(tmp.descricao || '') === String(pe.descricao || '')) {
        // sem mudança: "Editar no PO" de uma frase liga a medida a ela (como no comparador anterior)
        const f = alvo && linha >= 0 ? frasesPO(pe).find(x => !x.recurso && x.linha === linha) : null;
        fechar();
        if (f) { ligar(alvo.l, alvo.m, 'ok', f); mudou(); renderPrevia(); }
        return;
      }
      pe.descricao = tmp.descricao;
      const depois = frasesPO(pe).filter(f => !f.recurso), depoisS = new Set(depois.map(f => f.n));
      const novas = depois.filter(f => !antesS.has(f.n));
      const troca = new Map();
      for (const t of antes) {
        if (depoisS.has(t)) continue;
        let m = null, v = 0.5;
        for (const f of novas) { const s = semelhanca(t, f.n); if (s > v) { v = s; m = f.n; } }
        if (m) troca.set(t, m);
      }
      const aes = aprsDoPO(ip, paresEtapas());
      for (const ae of aes) {
        for (const l of ae.linhas) for (const c of Object.values(l.conf || {})) if (c.estado === 'ok' && troca.has(c.frase)) c.frase = troca.get(c.frase);
        if (ae.poDispensa) ae.poDispensa = ae.poDispensa.map(x => troca.get(x) || x);
      }
      if (alvo) {
        const f = novas.find(x => x.n === N(sugerida)) || (linha >= 0 ? depois.find(x => x.linha === linha && !antesS.has(x.n)) : null) || novas[novas.length - 1]
          || (linha >= 0 ? depois.find(x => x.linha === linha) : null);
        if (f) ligar(alvo.l, alvo.m, 'ok', f);
      }
      fechar(); mudarPO(); mudou();
      aviso('Descrição do PO atualizada (salve o PO na aba PO).');
    };
    md.replaceChildren(el('div', {class: 'janela apr-pop cf2-pop', onclick: ev => ev.stopPropagation()},
      el('header', {}, el('b', {class: 'quebra'}, `Editar descrição do PO – Etapa ${ip + 1}: ${pe.nome || ''}`), el('span', {style: 'flex:1'}),
        el('button', {class: 'btn', onclick: fechar}, 'Cancelar'),
        el('button', {class: 'btn primario', title: 'Aplica o texto no PO (o arquivo é salvo pelo SALVAR)', onclick: salvar}, 'Aplicar')),
      el('div', {class: 'apr-pop-corpo'},
        alvo ? [el('div', {class: 'cf-h'}, 'Medida da APR que precisa ser explicada'),
          el('div', {class: 'apr-pop-medida'}, el('span', {class: 'cf2-m mr'}, 'M' + (alvo.k + 1)), el('span', {style: 'flex:1'}, alvo.m, el('span', {class: 'cf-per2'}, ' – ' + (alvo.l.perigo || ''))))] : null,
        el('div', {class: 'cf-h'}, sugerida ? 'Descrição completa da etapa no PO (a frase sugerida já está no fim; ajuste à vontade)' : 'Descrição completa da etapa no PO'),
        ed)));
    md.onclick = null; md.hidden = false;
    setTimeout(() => {
      const d = nova && nova.isConnected ? nova : linha >= 0 && caixa.children[linha] ? caixa.children[linha] : caixa.lastElementChild;
      if (d) { if (nova || linha >= 0) d.classList.add('cf-alvo'); d.scrollIntoView({block: 'center'}); const r = document.createRange(); r.selectNodeContents(d); r.collapse(false); const s = getSelection(); s.removeAllRanges(); s.addRange(r); }
      caixa.focus();
    }, 50);
  }
  // comparador (aba APR ou PO): clicar em qualquer lugar da tela fora do painel azul da medida fecha a edição
  // (a linha da medida aberta e as janelas por cima não contam; clicar de novo na medida já fecha sozinho)
  document.addEventListener('click', ev => {
    if (!conf || !conf.real) return;
    if (ev.target.closest && ev.target.closest('.cf2-painel, .cf2-med.real, #modal')) return;
    conf.real = null; conf.tudo = false; renderPrevia();
  }, true);
  document.addEventListener('keydown', ev => {
    if (!conf || !$('#modal').hidden || ev.ctrlKey || ev.metaKey || ev.altKey) return;
    if (ev.target.closest && ev.target.closest('input, textarea, select, [contenteditable]')) return;
    const par = paresEtapas()[conf.i] || {a: -1, p: -1}, ae = apr.etapas[par.a], pe = doc.etapas[par.p];
    if (ev.key === 'ArrowLeft' || ev.key === 'ArrowRight') { ev.preventDefault(); irEtapa2(conf.i + (ev.key === 'ArrowRight' ? 1 : -1)); return; }
    if (!ae || !pe) return;
    // a medida realçada é a "selecionada"
    const ms = medidasDaEtapa(ae), k = conf.real && conf.real.m != null ? conf.real.m : -1, atual = ms[k];
    if (ev.key === 'ArrowDown' || ev.key === 'ArrowUp') {
      if (!ms.length) return;
      ev.preventDefault(); conf.real = {m: k < 0 ? 0 : Math.max(0, Math.min(ms.length - 1, k + (ev.key === 'ArrowDown' ? 1 : -1)))}; conf.rolar = true; renderPrevia();
    } else if (ev.key === 'Escape' && conf.real) { conf.real = null; conf.incluir = null; conf.tudo = false; renderPrevia(); }
    else if (atual && /^[1-3]$/.test(ev.key)) { const s = parecidas(atual.m, frasesPO(pe))[+ev.key - 1]; if (s) { ev.preventDefault(); ligar(atual.l, atual.m, 'ok', s.f); avancar(ae, pe, k); mudou(); renderPrevia(); } }
    else if (atual && ev.key.toLowerCase() === 'n') { ev.preventDefault(); ligar(atual.l, atual.m, 'nao'); avancar(ae, pe, k); mudou(); renderPrevia(); }
  });

  // pop-up: a medida em cima e a descrição da etapa do PO embaixo, editável, para escrever onde preferir
  function popupPO(pe, i, l, m, linha) {
    const md = $('#modal');
    const tmp = Object.assign({}, pe);                 // cópia: só vale para o PO ao clicar em "Salvar no PO"
    editandoPO = true;
    const fechar = () => { editandoPO = false; md.hidden = true; md.replaceChildren(); };
    const ed = editorTexto(tmp), caixa = ed.querySelector('.rich');
    const inserir = () => {
      const s = getSelection();
      let d = s.rangeCount && caixa.contains(s.anchorNode) ? (s.anchorNode.nodeType === 1 ? s.anchorNode : s.anchorNode.parentElement).closest('.rich > div') : null;
      const nova = document.createElement('div'); nova.textContent = m + ';';
      if (d && !d.textContent.trim()) d.replaceWith(nova); else if (d) d.after(nova); else caixa.append(nova);
      caixa.dispatchEvent(new Event('input'));
      const r = document.createRange(); r.selectNodeContents(nova); r.collapse(false); s.removeAllRanges(); s.addRange(r); caixa.focus();
    };
    const antes = String(pe.descricao || '').split('\n').map(x => N(textoPuro(x)));
    md.replaceChildren(el('div', {class: 'janela apr-pop', onclick: ev => ev.stopPropagation()},
      el('header', {}, el('b', {class: 'quebra'}, `${linha >= 0 ? 'Editar' : 'Incluir'} no PO – Etapa ${i + 1}: ${pe.nome || ''}`), el('span', {style: 'flex:1'}),
        el('button', {class: 'btn', onclick: fechar}, 'Cancelar'),
        el('button', {class: 'btn primario', onclick: () => {
          pe.descricao = tmp.descricao;
          // a medida fica ligada à frase nova ou alterada
          const depois = frasesPO(pe).filter(f => !f.recurso);
          const f = depois.find(x => !antes.includes(x.n)) || (linha >= 0 ? depois.find(x => x.linha === linha) : null);
          if (f) ligar(l, m, 'ok', f);
          fechar(); mudarPO(); mudou();
          aviso('Descrição do PO atualizada (salve o PO na aba PO).');
        }}, 'Salvar no PO')),
      el('div', {class: 'apr-pop-corpo'},
        el('div', {class: 'cf-h'}, 'Medida de controle da APR'),
        el('div', {class: 'apr-pop-medida'}, el('span', {style: 'flex:1'}, m),
          el('button', {class: 'btn mini', onclick: () => { navigator.clipboard && navigator.clipboard.writeText(m); aviso('Medida copiada.'); }}, 'Copiar'),
          el('button', {class: 'btn mini', title: 'Coloca a medida numa linha nova, logo abaixo da linha onde está o cursor', onmousedown: ev => ev.preventDefault(), onclick: inserir}, 'Inserir aqui')),
        el('div', {class: 'cf-h'}, 'Descrição da etapa no PO (escreva ou cole a medida onde preferir)'),
        ed)));
    md.onclick = null; md.hidden = false;
    // "Editar no PO": cursor no fim da frase escolhida
    setTimeout(() => {
      const d = linha >= 0 ? caixa.children[linha] : null;
      if (d) { d.classList.add('cf-alvo'); d.scrollIntoView({block: 'center'}); const r = document.createRange(); r.selectNodeContents(d); r.collapse(false); const s = getSelection(); s.removeAllRanges(); s.addRange(r); }
      caixa.focus();
    }, 50);
  }

  // ---------------- editor ----------------
  const bloco = (chave, ...f) => el('div', {'data-apr-ed': chave}, ...f);
  function edCabecalho() {
    const c = apr.cab, ca = cabAPR(), u = ultimaRev();
    return secao('Cabeçalho da APR', true,
      // logotipo da empresa (cabeçalho da APR e da Ordem de Serviço): 1º item, pedido do usuário (28/10/2026)
      bloco('cab:logo',
        el('label', {class: 'campo'}, el('span', {}, 'Logotipo'),
          apr.logo ? el('div', {class: 'apr-logo-ed'}, el('img', {src: apr.logo, alt: 'Logotipo'})) : null),
        el('div', {class: 'linha-botoes'}, botao(apr.logo ? 'Trocar logotipo' : 'Inserir logotipo', trocarLogo, 'Escolher a imagem do logotipo da APR e da Ordem de Serviço (prévia, Excel e PDF)', 'btn')),
        apr.logo ? el('div', {class: 'linha-botoes'}, botao('Remover logotipo', () => { apr.logo = ''; mudou(true); }, 'Deixar a APR sem logotipo', 'btn')) : null),
      el('div', {class: 'dupla'},
        bloco('cab:numero', campo('Nº da APR (ex.: AMP.009)', c, 'numero', {placeholder: 'XX.000'})),
        bloco('cab:metodo', campo('Método / Instrução de Seg.', c, 'metodo', {placeholder: ca.metodo}))),
      bloco('cab:atividade', campo('Atividade', c, 'atividade', {placeholder: 'Nome da atividade'})),
      el('div', {class: 'dupla'}, bloco('cab:area', campo('Área', c, 'area')), bloco('cab:departamento', campo('Departamento', c, 'departamento'))),
      el('div', {class: 'dupla'}, bloco('cab:maquina', campo('Máquina/Equipamento', c, 'maquina')), bloco('cab:avaliado', campo('Avaliado por', c, 'avaliado'))),
      bloco('cab:rev', caixa('Revisão e data automáticas pelo controle de revisão', c, 'revAuto')),
      c.revAuto ? el('div', {class: 'dica', id: 'dicaRevApr'}, u < 0 ? 'Preencha o controle de revisão da APR (mais abaixo).' : `Usando o item ${u}: Revisão ${pad2(u)}, data ${fmtData(apr.historico[u].data) || '(sem data)'}.`)
        : el('div', {class: 'dupla'}, bloco('cab:rev', campo('Revisão nº', c, 'revisao')), bloco('cab:data', campo('Data', c, 'data', {tipo: 'date'}))),
      bloco('cab:episBasicos', campo("EPI's básicos obrigatórios", c, 'episBasicos', {tipo: 'textarea', linhas: 2})));
  }
  // logotipo: a imagem escolhida é guardada em PNG (fundo transparente mantido), com no máximo 1200 px de lado e
  // a proporção original; o ajuste ao espaço de cada cabeçalho é feito na prévia e no Excel (sem distorcer)
  function trocarLogo() {
    const inp = el('input', {type: 'file', accept: 'image/*'});
    inp.addEventListener('change', () => {
      const f = inp.files[0]; if (!f) return;
      const url = URL.createObjectURL(f), img = new Image();
      img.onload = () => {
        const k = Math.min(1, 1200 / Math.max(img.naturalWidth, img.naturalHeight));
        const c = document.createElement('canvas');
        c.width = Math.max(1, Math.round(img.naturalWidth * k)); c.height = Math.max(1, Math.round(img.naturalHeight * k));
        c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
        URL.revokeObjectURL(url);
        apr.logo = c.toDataURL('image/png'); mudou(true);
      };
      img.onerror = () => { URL.revokeObjectURL(url); aviso('Não consegui ler a imagem ' + f.name); };
      img.src = url;
    });
    inp.click();
  }
  function edLinha(ae, l, j) {
    const oTab = daTabela(l.perigo);
    const foraDaTabela = l.perigo.trim() && !oTab;
    const r = riscoDe(l), nv = nivel(r);
    const sel = k => el('select', {onchange: ev => { l[k] = ev.target.value; mudou(true); }},
      el('option', {value: ''}, '–'), [1, 2, 3, 4, 5].map(n => el('option', {value: n, selected: numero(l[k]) === n}, n)));
    return el('div', {class: 'apr-perigo', 'data-apr-ed': l.id + ':linha'},
      el('div', {class: 'apr-per-topo'},
        // perigo e risco numa linha só, com "…"; o texto inteiro aparece ao passar o mouse (23/10/2026); o número
        // fica dentro do campo, que começa na mesma borda do risco e dos outros campos (28/10/2026)
        bloco(l.id + ':perigo', el('button', {class: 'apr-perigo-sel' + (foraDaTabela ? ' fora' : '') + (l.perigo.trim() ? '' : ' vazio'),
          title: l.perigo.trim() ? l.perigo + (foraDaTabela ? ' (fora da lista)' : '') : 'Escolher o perigo na tabela de Perigos e Riscos (com busca)',
          onclick: () => escolherPerigo('Escolher o perigo', l.perigo, r => { aplicarPerigo(l, r); mudou(true); })},
          el('span', {}, el('span', {class: 'apr-num'}, `${j + 1}. `), l.perigo.trim() ? l.perigo + (foraDaTabela ? ' (fora da lista)' : '') : 'Escolha o perigo...'), el('span', {class: 'seta'}, '▾'))),
        botao('↑', () => mover(ae.linhas, j, -1), 'Subir'), botao('↓', () => mover(ae.linhas, j, 1), 'Descer'),
        botao('✕', async () => { if (await perguntar(`Excluir o perigo "${l.perigo || j + 1}"?`, {sim: 'EXCLUIR', titulo: 'EXCLUIR PERIGO'})) { ae.linhas.splice(j, 1); mudou(true); } }, 'Excluir', 'btn mini perigo')),
      foraDaTabela ? (() => {
        const sug = parecidoNaTabela(l.perigo, l.pd);
        return el('div', {class: 'apr-pdtab mudou'}, 'Este perigo está fora da lista (veio do PO ou de uma APR antiga). Escolha o perigo da lista.',
          sug ? [el('br'), `Parecido na lista: "${sug.perigo}" (risco: ${sug.risco}). `,
            el('button', {class: 'btn mini', onclick: () => { aplicarPerigo(l, sug); mudou(true); }}, 'Usar este')] : null);
      })() : null,
      oTab && l.risco !== oTab.risco ? el('div', {class: 'apr-pdtab mudou'}, `O risco está diferente da tabela ("${oTab.risco}").`,
        el('button', {class: 'btn mini', onclick: () => { l.risco = oTab.risco; mudou(true); }}, 'Usar o risco da tabela')) : null,
      bloco(l.id + ':risco', el('div', {class: 'campo'}, el('span', {}, 'Risco (efeito / dano)'),
        el('div', {class: 'apr-fixo' + (l.risco ? '' : ' vazio'), title: l.risco || 'O risco é o da tabela para este perigo'}, l.risco || 'Escolha o perigo'))),
      el('div', {class: 'dupla'},
        bloco(l.id + ':epc', campo("EPC's", l, 'epc', {placeholder: 'N/A'})),
        bloco(l.id + ':epis', (() => {
          // texto livre (separado por vírgula) e, ao lado, a lista da biblioteca
          const c = campo("EPI's específicos", l, 'epis', {placeholder: 'N/A (separe com vírgula)'});
          const inp = c.querySelector('input');
          c.replaceChildren(c.firstChild, el('div', {class: 'apr-epi-linha'}, inp,
            botao('LISTA', () => escolherEpis(l, () => { inp.value = l.epis; renderPrevia(); }), 'Escolher os EPIs da biblioteca (o texto digitado continua valendo)', 'btn')));
          return c;
        })())),
      // método em cada perigo, depois de EPIs/EPCs (era só no 1º perigo, da etapa; pedido do usuário, 28/10/2026)
      bloco(l.id + ':metodo', campo('Método / Instrução de Segurança', l, 'metodo', {placeholder: cabAPR().metodo})),
      // PO e PD dividem a coluna de EPC's; o risco ocupa a coluna de EPI's específicos (23/10/2026)
      bloco(l.id + ':po', el('div', {class: 'apr-aval'},
        el('div', {class: 'apr-popd'}, el('label', {}, 'PO ', sel('po')), el('label', {}, 'PD ', sel('pd'))),
        el('span', {class: 'apr-risco', style: nv ? `background:${nv.fundo};color:${nv.cor}` : ''}, r ? `RISCO ${r} – ${nv.nome.toUpperCase()}` : 'RISCO: ESCOLHA PO E PD'))),
      oTab && oTab.pd ? el('div', {class: 'apr-pdtab' + (numero(l.pd) !== oTab.pd ? ' mudou' : '')},
        numero(l.pd) === oTab.pd ? `PD ${oTab.pd} pela tabela.` : `PD da tabela: ${oTab.pd} (alterado para ${numero(l.pd) || '–'}).`,
        oTab.obs ? el('span', {}, ' Quando muda: ' + oTab.obs) : null,
        numero(l.pd) !== oTab.pd ? el('button', {class: 'btn mini', onclick: () => { l.pd = String(oTab.pd); mudou(true); }}, 'Voltar ao da tabela') : null) : null,
      bloco(l.id + ':medidas', el('div', {class: 'campo'}, el('span', {}, 'Medidas de controle (uma por linha)'),
        campoMedidas(l))));
  }
  // medidas de controle: cada medida começa com o marcador "- " (pedido do usuário, 19/10/2026). Enter abre a
  // linha seguinte já com "- "; Shift+Enter quebra a linha sem marcador (continuação da mesma medida); campo
  // vazio ganha "- " ao entrar; ao sair, linha só com o marcador some.
  function campoMedidas(l) {
    const MARC = '- ';
    const t = el('textarea', {rows: Math.max(3, itensMedidas(l.medidas).length + 1), placeholder: '- Medida 1;\n- Medida 2;', value: l.medidas,
      oninput: ev => { l.medidas = ev.target.value; mudou(); }});
    const trocar = (txt, ini, fim, cursor) => { t.setRangeText(txt, ini, fim, cursor); t.dispatchEvent(new Event('input')); };
    t.addEventListener('focus', () => { if (!t.value.trim()) { t.value = ''; trocar(MARC, 0, 0, 'end'); } });
    t.addEventListener('keydown', ev => {
      if (ev.key !== 'Enter' || ev.shiftKey || ev.ctrlKey || ev.altKey || ev.metaKey || ev.isComposing) return;
      ev.preventDefault();
      trocar('\n' + MARC, t.selectionStart, t.selectionEnd, 'end');
    });
    t.addEventListener('blur', () => {
      const limpo = t.value.split('\n').filter(x => !/^\s*[-–—]\s*$/.test(x)).join('\n');
      if (limpo !== t.value) { t.value = limpo; l.medidas = limpo; mudou(); }
    });
    return t;
  }
  function edEtapa(ae, i) {
    return el('div', {class: 'etapa', id: 'apr-' + ae.id},
      el('div', {class: 'etapa-topo'},
        el('b', {}, `Etapa ${i + 1}`),
        botao('↑', () => mover(apr.etapas, i, -1), 'Subir etapa'), botao('↓', () => mover(apr.etapas, i, 1), 'Descer etapa'),
        botao('+ abaixo', () => { apr.etapas.splice(i + 1, 0, novaEtapa()); mudou(true); }, 'Inserir uma etapa nova logo abaixo', 'btn mini larg'),
        botao('Duplicar', () => { duplicarEtapa(apr.etapas, i); mudou(true); }, 'Criar uma cópia completa desta etapa logo abaixo', 'btn mini larg'),
        botao('Excluir', async () => { if (await perguntar(`Excluir a etapa ${i + 1}${ae.nome ? ' (' + ae.nome + ')' : ''}?`, {sim: 'EXCLUIR', titulo: 'EXCLUIR ETAPA'})) { if (typeof sumir === 'function') await sumir('apr-' + ae.id); apr.etapas.splice(i, 1); mudou(true); } }, '', 'btn mini perigo larg')),
      bloco(ae.id + ':nome', campo('Etapa do processo', ae, 'nome')),
      bloco(ae.id + ':situacao', selecao('Situação', ae, 'situacao', SITUACOES, true)),
      el('div', {class: 'sub'}, 'Perigos'),
      ae.linhas.map((l, j) => edLinha(ae, l, j)),
      el('div', {class: 'linha-botoes'}, botao('+ Perigo', () => { ae.linhas.push(novaLinha()); mudou(true); }, '', 'btn')));
  }
  // campos que só aparecem na aba Ordem de Serviço (vazios na APR nova)
  function edOrdemServico() {
    const c = apr.cab;
    return secao('Ordem de Serviço', true,
      el('div', {class: 'dupla'},
        bloco('cab:empresa', campo('Empresa', c, 'empresa')),
        bloco('cab:elaboracao', campo('Elaboração', c, 'elaboracao'))));
  }
  function edHistorico() {
    const h = apr.historico;
    // cada campo com o título logo acima (`campoHist`, o mesmo do histórico do PO) e as duas linhas de cada item com
    // a mesma borda direita (28/10/2026, pedido do usuário)
    const ch = (t, attrs, ...f) => typeof campoHist === 'function' ? campoHist(t, attrs, ...f) : el('div', attrs, ...f);
    return secao('Controle de revisão da APR', false,
      el('table', {class: 'hist'}, el('tbody', {},
        h.map((x, i) => [
        el('tr', {}, el('td', {class: 'hist-n'}, i),
          el('td', {colspan: 4}, el('div', {class: 'hist-linha hist-desc'},
            ch('Motivo', {'data-apr-ed': 'hist:' + i}, el('input', {type: 'text', value: x.motivo, oninput: ev => { x.motivo = ev.target.value; mudou(); }})),
            botao('✕', async () => { if (await perguntar(`Excluir o item ${i} do controle de revisão?`, {sim: 'EXCLUIR', titulo: 'EXCLUIR ITEM'})) { h.splice(i, 1); mudou(true); } }, '', 'btn mini perigo')))),
        el('tr', {}, el('td'), el('td', {colspan: 4}, el('div', {class: 'hist-linha hist-tres'},
          ch('Data', {}, el('input', {type: 'date', value: x.data, title: 'Data', oninput: ev => {
            x.data = ev.target.value; mudou();
            const d = document.getElementById('dicaRevApr'); const u = ultimaRev(); if (d && u >= 0) d.textContent = `Usando o item ${u}: Revisão ${pad2(u)}, data ${fmtData(h[u].data) || '(sem data)'}.`;
          }})),
          ch('Responsável', {}, el('input', {type: 'text', value: x.responsavel, oninput: ev => { x.responsavel = ev.target.value; mudou(); }})),
          ch('Origem', {}, el('input', {type: 'text', value: x.origem, list: 'aprOrigens', oninput: ev => { x.origem = ev.target.value; mudou(); }}))))),
        el('tr', {}, el('td', {colspan: 5, style: 'height:6px;border-bottom:1px solid #e5e7eb'}))]))),
      el('div', {class: 'linha-botoes', 'data-apr-ed': 'hist:novo'}, botao('+ Revisão', () => {
        h.push({data: new Date().toISOString().slice(0, 10), motivo: '', responsavel: '', origem: ''}); mudou(true);
      }, '', 'btn')),
      el('div', {class: 'linha-botoes'}, botao('Importar do PO', histDoPO, 'Copia o histórico de revisões do PO aberto na aba PO', 'btn')));
  }
  function editor() {
    const origens = [...new Set(['Emissão', 'Auditoria Interna', 'Revisão Periódica', 'Acidente Sem Lesão', 'Acidente Com Lesão', ...apr.historico.map(x => x.origem).filter(Boolean)])];
    return [
      el('datalist', {id: 'aprOrigens'}, origens.map(o => el('option', {value: o}))),
      el('div', {class: 'apr-topo'},
        botao('Importar do PO', puxarDoPO, 'Traz as etapas, os perigos e os EPIs do PO aberto na aba PO', 'btn'),
        botao(conf ? 'Fechar conferência' : 'Comparar com PO', abrirConferencia, 'Confere a APR com o PO aberto na aba PO, etapa por etapa', conf ? 'btn ativo' : 'btn'),
        botao('Pendências', () => relatorio(), 'Lista o que falta preencher na APR', 'btn')),
      // a linha "X campo(s) a preencher / ✓ Campos preenchidos" saiu a pedido do usuário (19/10/2026): "Pendências" mostra a lista
      edCabecalho(),
      secao('Etapas da APR', true, apr.etapas.map((ae, i) => edEtapa(ae, i)),
        el('div', {class: 'linha-botoes'}, botao('+ Adicionar etapa', () => { apr.etapas.push(novaEtapa()); mudou(true); if (typeof realcarNovo === 'function') realcarNovo(apr.etapas[apr.etapas.length - 1].id); setTimeout(() => $('#editor .etapa:last-of-type')?.scrollIntoView({behavior: 'smooth'}), 50); }, '', 'btn primario'))),
      edOrdemServico(),
      edHistorico(),
    ];
  }
  function menu() {
    return [
      ['Nova APR', nova, 'Começar uma APR nova'],
      ['Abrir', abrir, 'Abrir uma APR em Excel (ou .apr.json antigo) (Ctrl+O)'],
      ['Salvar como', () => salvar(true), 'Escolher onde salvar (Ctrl+Shift+S)'],
      null,
      ['Exportar PDF', pdf, 'Abre a impressão para salvar a APR em PDF', 'exp pdf'],
      ['Exportar Excel', () => exportar(), 'Salva a APR em .xlsx, no modelo da APR', 'exp xls'],
      null,
      ['Atualizar Tabela de Perigos', atualizarTabela, 'Escolha a planilha de Perigos e Riscos (Excel) para atualizar a lista do gerador'],
    ];
  }

  // ---------------- layout (o mesmo para a prévia e para o Excel) ----------------
  // largura das colunas do modelo (caracteres) -> pixels, como o Excel faz com a fonte padrão Arial 10 (7 px por dígito)
  const COLS = 'ABCDEFGHIJKLMNOPQRSTUVWX'.split('');
  const LCOL = {A: 5.6640625, B: 5.109375, C: 7.5546875, D: 3.5546875, E: 6.33203125, F: 1.5546875, G: 3.44140625, H: 2.5546875, I: 8.33203125, J: 25.5546875,
    K: 7, L: 5.44140625, M: 4.44140625, N: 3.109375, O: 5.109375, P: 1.109375, Q: 20.44140625, R: 21.88671875, S: 19.44140625, T: 6.33203125, U: 6.33203125,
    V: 10.5546875, W: 20.88671875, X: 39.33203125};
  const pxCol = c => Math.trunc((256 * LCOL[c] + 18) / 256 * 7);
  const px = (...cols) => cols.reduce((s, c) => s + pxCol(c), 0);
  const LARG_TOTAL = px(...COLS);
  const MARGEM_MM = 5;                                                   // 0,5 cm em todos os lados
  const ALT_CAB = [48, 48, 42, 40.5, 18.6, 31.5];                        // pt das linhas 1 a 6 (repetidas em cada página)
  const ALT_ROD = [18, 16.8, 16.8, 15, 15, 15, 15.6];                    // pt do rodapé
  // Na impressão o Excel ajusta a planilha à largura da página em 56,87% (medido no PDF do Excel com margens de 0,5 cm:
  // a folha toda ocupa 200 mm de altura quando as linhas somam 200 mm / 0,5687).
  // A prévia usa a mesma escala: alturas e fontes em pt iguais ao Excel, colunas alargadas até a largura da página.
  const ESCALA_XL = 0.5687;
  // a prévia ocupa a largura toda (287 mm), descontada a borda grossa de 2 px da tabela, que fica para fora
  const ESCALA = ((297 - 2 * MARGEM_MM) / 25.4 * 96 - 2 * ESCALA_XL) / LARG_TOTAL;
  const LARGURA_PREVIA = c => pxCol(c) * ESCALA / ESCALA_XL;
  const LARG_PREVIA = COLS.reduce((s, c) => s + LARGURA_PREVIA(c), 0);
  const ALT_PAGINA = ((210 - 2 * MARGEM_MM) / 25.4 * 72) / ESCALA_XL;   // pt da folha que cabem numa página
  // pt de segurança: o Excel arredonda a altura de cada linha para cima (até 0,1 pt), e uma linha a mais que não caiba
  // iria para a página seguinte; por isso fica 1 pt + 0,1 pt por linha da página
  const FOLGA = 1, FOLGA_LINHA = 0.1;  const ALT_LINHA = 16.8;       // pt por linha de texto Arial 12 (como no modelo)
  const medidor = document.createElement('canvas').getContext('2d');
  function linhas(t, larg, fonte = '16px Arial') {
    medidor.font = fonte;
    const util = larg - 8;
    let n = 0;
    for (const par of String(t || '').replace(/\s+$/, '').split('\n')) {
      const ws = par.split(/\s+/).filter(Boolean);
      if (!ws.length) { n++; continue; }
      let l = '', c = 1;
      for (const w of ws) {
        const s = l ? l + ' ' + w : w;
        if (medidor.measureText(s).width * 1.04 > util && l) { c++; l = w; } else l = s;
        while (medidor.measureText(l).width * 1.04 > util && l.length > 1) { c++; l = l.slice(Math.ceil(l.length * util / medidor.measureText(l).width / 1.04)); }
      }
      n += c;
    }
    return Math.max(1, n);
  }
  function layout() {
    const ca = cabAPR();
    const etapas = apr.etapas.map((ae, i) => {
      const lista = ae.linhas.length ? ae.linhas : [null];
      const alts = lista.map(l => !l ? 2 * ALT_LINHA : Math.max(2, linhas(l.perigo, px('J')), linhas(l.risco, px('K', 'L', 'M', 'N', 'O', 'P')),
        linhas(textoEpc(l), px('Q')), linhas(textoEpis(l), px('R')), linhas(metodoDe(l, ca), px('S')), linhas(l.medidas, px('W', 'X'))) * ALT_LINHA);
      // o nome da etapa (mesclado nas linhas da etapa) também precisa caber
      const precisa = linhas(ae.nome, px('B', 'C', 'D', 'E', 'F', 'G', 'H')) * ALT_LINHA;
      const soma = alts.reduce((a, b) => a + b, 0);
      if (soma < precisa) alts[alts.length - 1] += precisa - soma;
      return {ae, i, lista, alts, h: Math.max(soma, precisa)};
    });
    // páginas: a etapa não é dividida; o rodapé vai na última página (ou numa página só para ele, se não couber)
    const folga = n => FOLGA_LINHA * n, folgaEt = et => folga(et.lista.length);
    const util = ALT_PAGINA - FOLGA - ALT_CAB.reduce((a, b) => a + b, 0) - folga(ALT_CAB.length), rod = ALT_ROD.reduce((a, b) => a + b, 0) + folga(ALT_ROD.length);
    const paginas = [[]];
    let usado = 0;
    for (const et of etapas) {
      if (usado > 0 && usado + et.h + folgaEt(et) > util) { paginas.push([]); usado = 0; }
      paginas[paginas.length - 1].push(et); usado += et.h + folgaEt(et);
    }
    const rodapeSozinho = usado > 0 && usado + rod > util;
    // a sobra de cada página é repartida entre as linhas das etapas dela (proporcional à altura de cada linha),
    // para a tabela terminar rente à margem de baixo; a última página das etapas termina onde acabar (pedido do usuário)
    paginas.slice(0, paginas.length - 1).forEach(p => {
      const total = p.reduce((s, et) => s + et.h, 0), alvo = util - p.reduce((s, et) => s + folgaEt(et), 0);
      if (!total || total >= alvo) return;
      const f = alvo / total;
      for (const et of p) { et.alts = et.alts.map(a => a * f); et.h *= f; }
    });
    return {ca, etapas, paginas, rodapeSozinho};
  }

  // ---------------- prévia (igual ao Excel) ----------------
  // o logotipo da Siderol do modelo não é mais usado (pedido do usuário, 28/10/2026): sem logotipo escolhido, sai em branco
  const logoAtual = () => apr.logo;
  function previa() {
    if (conf && temPO()) { desenharComparador(); return; }
    conf = null;
    if (abaAtual !== 'apr') { renderPrevia(); return; }   // conferência fechada (o PO ficou vazio) na aba PO: volta a prévia do PO
    const marca = new Map();
    const {ca, paginas, rodapeSozinho} = layout();
    const W = c => LARGURA_PREVIA(c).toFixed(2) + 'px', H = pt => (pt * 4 / 3).toFixed(2) + 'px';
    const td = (chave, cont, cls = '', extra = {}) => {
      const p = chave && marca.get(chave);
      return el('td', Object.assign({'data-apr': chave || null, class: cls + (p ? ' apr-prob' : ''), title: p ? p.join('\n') : null}, extra), cont);
    };
    const rico = (rot, val) => [el('b', {}, rot), el('br'), el('span', {style: 'font-weight:400'}, val || '')];
    const cabecalho = () => [
      el('tr', {style: `height:${H(ALT_CAB[0])}`},
        el('td', {colspan: 6, class: 'c'}, logoAtual() ? el('img', {src: logoAtual(), alt: '', class: 'apr-logo'}) : null),
        el('td', {colspan: 13, class: 'c apr-titulo'}, 'ANÁLISE PRELIMINAR DE RISCOS (APR)'),
        td('cab:rev', rico('DATA:', ca.data), 'c', {colspan: 4}),
        td('cab:numero', rico('Nº', ca.numero), 'c b')),
      el('tr', {style: `height:${H(ALT_CAB[1])}`},
        td('cab:area', rico('ÁREA:', ca.area), 'e', {colspan: 6}), td('cab:departamento', rico('DEPARTAMENTO:', ca.departamento), 'e', {colspan: 4}),
        td('cab:maquina', rico('MÁQUINA/EQUIPAMENTO:', ca.maquina), 'e', {colspan: 7}), td('cab:atividade', rico('ATIVIDADE:', ca.atividade), 'e', {colspan: 2}),
        td('cab:avaliado', rico('AVALIADO POR:', ca.avaliado), 'c', {colspan: 4}), td('cab:rev', rico('REVISÃO:', ca.rev), 'c')),
      el('tr', {style: `height:${H(ALT_CAB[2])}`}, td('cab:episBasicos', rico("EPI's BÁSICOS OBRIGATÓRIOS:", ca.episBasicos), 'e', {colspan: 24})),
      el('tr', {class: 'cinza b', style: `height:${H(ALT_CAB[3])}`},
        el('td', {rowspan: 3, class: 'c'}, 'Nº'), el('td', {rowspan: 3, colspan: 7, class: 'c', style: 'font-size:18pt'}, 'Etapa do processo'),
        el('td', {rowspan: 3, class: 'c'}, el('div', {class: 'apr-vert'}, 'Situação')), el('td', {rowspan: 3, class: 'c'}, 'Perigo'),
        el('td', {rowspan: 3, colspan: 6, class: 'c'}, 'Risco', el('br'), '(Efeito / Dano)'),
        el('td', {colspan: 3, class: 'c'}, 'Ações', el('br'), el('i', {}, 'Necessárias para eliminar ou minimizar os riscos')),
        el('td', {colspan: 3, class: 'c'}, 'Avaliação do Risco'), el('td', {rowspan: 3, colspan: 2, class: 'c'}, 'Medidas de Controle')),
      el('tr', {class: 'cinza b', style: `height:${H(ALT_CAB[4])}`},
        ...["EPC's", "EPI's específicos", 'Método / Instrução de Segurança', 'PO', 'PD', 'RISCO'].map(t => el('td', {rowspan: 2, class: 'c'}, t))),
      el('tr', {class: 'cinza', style: `height:${H(ALT_CAB[5])}`}),
    ];
    const linhasEtapa = et => {
      const {ae, i, lista, alts} = et;
      const epc = lista.map(l => l ? textoEpc(l) : ''), epis = lista.map(l => l ? textoEpis(l) : ''), met = lista.map(l => metodoDe(l, ca));
      const span = (lst, k) => { if (k > 0 && lst[k] === lst[k - 1]) return 0; let f = k; while (f + 1 < lst.length && lst[f + 1] === lst[k]) f++; return f - k + 1; };
      return lista.map((l, k) => {
        const tr = el('tr', {style: `height:${H(alts[k])}`});
        if (k === 0) tr.append(td(ae.id + ':nome', String(i + 1), 'c', {rowspan: lista.length}), td(ae.id + ':nome', ae.nome, 'c', {rowspan: lista.length, colspan: 7}),
          td(ae.id + ':situacao', ae.situacao, 'c', {rowspan: lista.length}));
        tr.append(td(l && l.id + ':perigo', l ? l.perigo : '', 'c'), td(l && l.id + ':risco', l ? l.risco : '', 'c', {colspan: 6}));
        const sq = span(epc, k), sr = span(epis, k), ss = span(met, k);
        if (sq) tr.append(td(l && l.id + ':epc', epc[k], 'c', {rowspan: sq}));
        if (sr) tr.append(td(l && l.id + ':epis', epis[k], 'c apr-cinza', {rowspan: sr}));
        // método do perigo; iguais em linhas seguidas ficam numa célula só (como EPC / EPI)
        if (ss) tr.append(td(l && l.id + ':metodo', met[k], 'c', {rowspan: ss}));
        const r = l && riscoDe(l), nv = nivel(r);
        tr.append(td(l && l.id + ':po', l && numero(l.po) || '', 'c'), td(l && l.id + ':po', l && numero(l.pd) || '', 'c'),
          td(l && l.id + ':po', r || '', 'c b', {style: nv ? `background:${nv.fundo};color:${nv.cor}` : ''}));
        // medidas como escritas (inclusive "- " e ";")
        const cel = td(l && l.id + ':medidas', '', 'e', {colspan: 2});
        if (l) for (const linha of String(l.medidas).replace(/\s+$/, '').split('\n')) cel.append(el('div', {}, linha || '\u00a0'));
        tr.append(cel);
        return tr;
      });
    };
    // rodapé igual ao do modelo: as 4 caixas de risco estão nas mesmas posições dos retângulos do Excel
    const rodape = () => {
      const x = (...c) => c.reduce((s, k) => s + LARGURA_PREVIA(k), 0);
      const caixas = [
        ['Risco Baixo - (De 1 a 4)', '#00B050', '#fff', 0, x('F', 'G', 'H', 'I') + 63, '10.5pt'],
        ['Risco Moderado - (De 5 a 11)', '#FFFF00', '#000', x('F', 'G', 'H', 'I') + 62, x('F', 'G', 'H', 'I', 'J', 'K', 'L') + 11, '11pt'],
        ['Risco Alto - (De 12 a 15)', '#FF6600', '#000', x('F', 'G', 'H', 'I', 'J', 'K', 'L') + 11, x('F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P') + 107, '11pt'],
        ['Risco Crítico - (De 16 a 25)', '#FF0000', '#fff', x('F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P') + 106, x('F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R'), '11pt'],
      ];
      const pd = [' 1 – Incômodo', ' 2 – Lesão Leve', ' 3 – Lesão Pequena', ' 4 – Lesão Grave / Dano Irreversível', ' 5 – Fatalidade /', ' Incapacidade total permanente'];
      const po = ['1 – Improvável', '2 – Pouco provável', '3 – Ocasional', '4 – Provável', '5 – Frequente', ''];
      const sit = [[el('b', {}, 'R'), ' – Rotineira'], [el('b', {}, 'NR'), ' – Não Rotineira'], '', '', '', ''];
      const ass = t => '\n\n___________________________________\n' + t;
      const ult = ALT_ROD.length - 2;
      return [
        el('tr', {class: 'rd rd1', style: `height:${H(ALT_ROD[0])}`},
          el('td', {colspan: 5, class: 'c rd-tit rd-l'}, 'Situação:'),
          el('td', {colspan: 13, class: 'rd-caixas rd-l rd-r'}, el('div', {}, caixas.map(([t, f, c, a, b, fs]) =>
            el('span', {style: `left:${a}px;width:${b - a}px;background:${f};color:${c};font-size:${fs}`}, t)))),
          el('td', {colspan: 4, class: 'c rd-tit rd-r'}, 'PD:  Potencial de Dano'),
          el('td', {colspan: 2, class: 'c rd-tit rd-l', style: 'vertical-align:top'}, 'PO: Probabilidade de Ocorrência')),
        ...ALT_ROD.slice(1).map((h, k) => el('tr', {class: 'rd' + (k === ult ? ' rd-fim' : ''), style: `height:${H(h)}`},
          el('td', {colspan: 5, class: 'c rd-l'}, sit[k]),
          k === 0 ? el('td', {colspan: 8, rowspan: 6, class: 'c b rd-l rd-ass'}, ass('Supervisor')) : null,
          k === 0 ? el('td', {colspan: 5, rowspan: 6, class: 'c b rd-r rd-ass'}, ass('Emitente')) : null,
          el('td', {colspan: 4, class: 'e rd-r rd-leg'}, pd[k]),
          el('td', {colspan: 2, class: 'e rd-l rd-leg', style: 'vertical-align:top'}, po[k]))),
      ];
    };
    const folha = conteudo => el('div', {class: 'pagina apr-pagina'}, el('div', {class: 'apr-folha', style: `zoom:${ESCALA_XL}`},
      el('table', {class: 'apr-t'}, el('colgroup', {}, COLS.map(c => el('col', {style: `width:${W(c)}`}))), el('tbody', {}, conteudo))));
    const pags = paginas.map((p, k) => folha([...cabecalho(), ...p.flatMap(linhasEtapa), ...(k === paginas.length - 1 && !rodapeSozinho ? rodape() : [])]));
    if (rodapeSozinho) pags.push(folha([...cabecalho(), ...rodape()]));
    $('#paginas').replaceChildren(...pags);
    ajustarZoom();
    abasExtras = desenharAbas();
  }
  // as outras abas do Excel exportado, desenhadas a partir do próprio arquivo (iguais à impressão do Excel)
  const ABAS_EXTRAS = ['Ordem de Serviço', 'Controle de revisão', 'Matriz de Risco'];
  let abasExtras = Promise.resolve(), vezAbas = 0;
  const imagensAbas = {};
  let logoAbas = '';
  let abasGuardadas = null;
  async function desenharAbas() {
    const vez = ++vezAbas;
    const colocar = pags => {
      const cont = $('#paginas');
      cont.querySelectorAll('.xl-pagina').forEach(p => p.remove());
      cont.append(...pags);
    };
    // (07/11/2026) a APR não mudou desde o último desenho (ex.: só trocou de aba APR/PO): põe as mesmas páginas
    // na hora, sem montar e reler o Excel de novo (era o que travava a troca de aba)
    let chave = JSON.stringify(apr);
    if (abasGuardadas && abasGuardadas.chave === chave) { colocar(abasGuardadas.pags); return; }
    await new Promise(ok => setTimeout(ok, 120));          // espera a digitação parar
    if (vez !== vezAbas || typeof PLANILHA === 'undefined') return;
    if (logoAbas !== apr.logo) {                           // logotipo trocado: as imagens guardadas das abas não valem mais
      for (const k in imagensAbas) { URL.revokeObjectURL(imagensAbas[k]); delete imagensAbas[k]; }
      logoAbas = apr.logo;
    }
    chave = JSON.stringify(apr);
    let arqs;
    try { arqs = await montarArquivos(); } catch (e) { console.error(e); return; }
    const pags = [];
    for (const nome of ABAS_EXTRAS) { try { const p = await PLANILHA.pagina(arqs, nome, imagensAbas); if (p) pags.push(p); } catch (e) { console.error(nome, e); } }
    abasGuardadas = {chave, pags};
    if (vez !== vezAbas || conf || abaAtual !== 'apr') return;
    colocar(pags);
  }
  function cliquePrevia(ev) {
    const c = ev.target.closest('[data-apr]');
    if (!c || ev.target.closest('button')) return;
    const k = c.dataset.apr;
    const alvo = document.querySelector(`#editor [data-apr-ed="${k}"]`) || document.getElementById('apr-' + k.split(':')[0]);
    if (!alvo) return;
    const det = alvo.closest('details'); if (det) det.open = true;
    alvo.scrollIntoView({behavior: 'smooth', block: 'center'});
    alvo.classList.add('destaque'); setTimeout(() => alvo.classList.remove('destaque'), 1500);
    const pc = palavraClicada(ev);
    const foco = alvo.matches('input, textarea, select') ? alvo : alvo.querySelector('input[type=text], textarea, select');
    if (foco) setTimeout(() => { foco.focus({preventScroll: true}); if (foco.matches('input[type=text], textarea')) irParaPalavra(foco, pc); }, 350);
  }
  async function pdf() {
    if (abaAtual !== 'apr') trocarAba('apr');
    const antes = conf; conf = null;                       // a impressão é sempre da APR
    imgTelaOriginal = true;                               // imagens no tamanho original (07/11/2026)
    renderPrevia();
    await abasExtras;                                     // Ordem de Serviço, Controle de revisão e Matriz de Risco
    await esperarImagens($('#paginas'));
    const titulo = document.title; document.title = nomeAPR();
    setTimeout(() => { window.print(); document.title = titulo; conf = antes; imgTelaOriginal = false; renderPrevia(); }, 150);
  }

  // ---------------- Excel ----------------
  const serial = iso => { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || ''); return m ? (Date.UTC(+m[1], m[2] - 1, +m[3]) - Date.UTC(1899, 11, 30)) / 86400000 : null; };
  const cel = (ref, s, v) => {
    if (v === undefined || v === null || v === '') return `<c r="${ref}" s="${s}"/>`;
    if (typeof v === 'number') return `<c r="${ref}" s="${s}"><v>${v}</v></c>`;
    if (typeof v === 'object') return `<c r="${ref}" s="${s}"><f>${esc(v.f)}</f>${v.v !== undefined && v.v !== '' ? `<v>${esc(v.v)}</v>` : ''}</c>`;
    return `<c r="${ref}" s="${s}" t="inlineStr"><is><t xml:space="preserve">${esc(v)}</t></is></c>`;
  };
  const rico = (rotulo, valor) => `<is><r><t xml:space="preserve">${esc(rotulo)}</t></r><r><rPr><sz val="12"/><rFont val="Arial"/><family val="2"/></rPr><t xml:space="preserve">${esc(valor)}</t></r></is>`;
  function trocaCel(xml, ref, conteudo, tipo = 'inlineStr') {
    const re = new RegExp(`<c r="${ref}"( s="\\d+")?[^>]*?(?:/>|>[\\s\\S]*?</c>)`);
    if (!re.test(xml)) throw new Error('célula ' + ref + ' não encontrada no modelo');
    return xml.replace(re, (m, s) => `<c r="${ref}"${s || ''}${tipo ? ` t="${tipo}"` : ''}>${conteudo}</c>`);
  }
  function trocaValor(xml, ref, texto) {     // valor guardado de uma fórmula (o Excel recalcula ao abrir)
    return xml.replace(new RegExp(`(<c r="${ref}"[^>]*>(?:<f[^>]*>[\\s\\S]*?</f>|<f[^>]*/>))(?:<v>[\\s\\S]*?</v>|<v/>)?`), (m, a) => `${a}<v>${esc(texto)}</v>`);
  }
  const desloca = (xml, d) => xml.replace(/<row r="(\d+)"/g, (m, n) => `<row r="${+n + d}"`).replace(/ r="([A-Z]+)(\d+)"/g, (m, c, n) => ` r="${c}${+n + d}"`);

  let modeloLido = null;
  function lerModelo() {
    if (!modeloLido) modeloLido = (async () => {
      // (07/11/2026) o modelo (~1,6 MB) só é baixado quando a APR precisa dele, e não mais ao abrir o gerador
      if (!window.MODELO_APR) { try { await carregarScript('recursos/modelo_apr.js'); } catch (e) { modeloLido = null; } }
      if (!window.MODELO_APR) { modeloLido = null; throw new Error('não encontrei o modelo da APR (recursos/modelo_apr.js)'); }
      const bin = Uint8Array.from(atob(window.MODELO_APR), c => c.charCodeAt(0));
      const z = await IMPORTAR.ajuda.lerZip(bin.buffer);
      const m = {};
      for (const n of z.nomes) m[n] = await z.ler(n);
      return m;
    })();
    return modeloLido;
  }
  // ---------------- APR inteira guardada dentro do Excel (customXml) ----------------
  // Vai junto no "Salvar" e no "Exportar Excel": ligações do comparador, método por perigo, logotipo, ids...
  // "texto" = impressão digital das células escritas (sem fórmulas, que o Excel recalcula): ao abrir, se não bate,
  // a planilha foi mexida direto no Excel.
  const URN_APR = 'urn:gerador-apr';
  function digitalAPR(T) {           // T(nome da parte do pacote) -> texto ('' se não existe)
    const {xml} = IMPORTAR.ajuda;
    const todosT = (n, tag) => [...n.getElementsByTagNameNS('*', tag)];
    const ss = [];
    const tss = T('xl/sharedStrings.xml');
    if (tss) for (const si of todosT(xml(tss), 'si')) ss.push(todosT(si, 't').filter(t => !t.closest('rPh')).map(t => t.textContent).join(''));
    const rels = {};
    for (const r of todosT(xml(T('xl/_rels/workbook.xml.rels')), 'Relationship')) rels[r.getAttribute('Id')] = r.getAttribute('Target');
    let s = '';
    for (const sh of todosT(xml(T('xl/workbook.xml')), 'sheet')) {
      const id = sh.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id') || sh.getAttribute('r:id');
      const t = T('xl/' + (rels[id] || '').replace(/^\/?(xl\/)?/, ''));
      if (!t) continue;
      for (const c of todosT(xml(t), 'c')) {
        if (todosT(c, 'f').length) continue;
        const tp = c.getAttribute('t'), v = todosT(c, 'v')[0], vt = v ? v.textContent.trim() : '';
        if (tp === 'inlineStr') s += todosT(c, 't').map(n => n.textContent).join('');
        else if (v && tp === 's') s += ss[+vt] || '';
        else if (v && tp !== 'e') s += vt && isFinite(+vt) ? String(Math.round(+vt * 1e6) / 1e6) : vt;
      }
    }
    s = N(s).replace(/ /g, '');
    let h = 0x811c9dc5;
    for (let i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 0x01000193) >>> 0; }
    return s.length + '-' + h.toString(16);
  }
  function guardarDados(arqs) {
    const dec = new TextDecoder(), enc = new TextEncoder();
    const T = n => arqs[n] ? dec.decode(arqs[n]) : '';
    let k = 1; while (arqs[`customXml/item${k}.xml`]) k++;
    const dados = JSON.stringify(apr).replace(/]]>/g, ']]]]><![CDATA[>');
    arqs[`customXml/item${k}.xml`] = enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><aprGerador xmlns="${URN_APR}" texto="${digitalAPR(T)}"><![CDATA[${dados}]]></aprGerador>`);
    arqs[`customXml/itemProps${k}.xml`] = enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="no"?><ds:datastoreItem ds:itemID="{8C2D4E6F-1A3B-4C5D-9E7F-0A1B2C3D4E5F}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs><ds:schemaRef ds:uri="${URN_APR}"/></ds:schemaRefs></ds:datastoreItem>`);
    arqs[`customXml/_rels/item${k}.xml.rels`] = enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps${k}.xml"/></Relationships>`);
    arqs['[Content_Types].xml'] = enc.encode(T('[Content_Types].xml').replace('</Types>', `<Override PartName="/customXml/itemProps${k}.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/></Types>`));
    arqs['xl/_rels/workbook.xml.rels'] = enc.encode(T('xl/_rels/workbook.xml.rels').replace('</Relationships>', `<Relationship Id="rIdDadosAPR" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item${k}.xml"/></Relationships>`));
  }
  // Excel aberto: a APR guardada pelo gerador e se o texto das células ainda é o mesmo (null = não tem)
  async function lerGuardado(buf) {
    const u8 = new Uint8Array(buf);
    if (u8[0] !== 0x50 || u8[1] !== 0x4B) return null;           // .xls antigo / com senha: não é zip
    const z = await IMPORTAR.ajuda.lerZip(buf);
    for (const n of z.nomes.filter(n => /^customXml\/item\d+\.xml$/i.test(n))) {
      const t = await z.texto(n);
      if (!t || !t.includes(URN_APR)) continue;
      const m = /<!\[CDATA\[([\s\S]*)\]\]>/.exec(t.replace(/\]\]\]\]><!\[CDATA\[>/g, ']]>'));
      if (!m) continue;
      const dig = /\btexto="([^"]+)"/.exec(t.slice(0, 300));
      const textos = {};
      for (const k of z.nomes) if (/^xl\/(workbook\.xml|_rels\/workbook\.xml\.rels|sharedStrings\.xml|worksheets\/[^/]+\.xml)$/i.test(k)) textos[k] = await z.texto(k);
      return {dados: JSON.parse(m[1]), igual: !!dig && digitalAPR(k => textos[k] || '') === dig[1]};
    }
    return null;
  }
  // APR lida da planilha mexida no Excel: o que só o gerador guarda (ligações do comparador, etapa do PO escolhida
  // à mão, ids, logotipo, empresa/elaboração) volta para as etapas (mesmo nome) e perigos (mesmo perigo) que continuam
  function juntarGuardado(d, g) {
    g = normalizar(g);
    const usadas = new Set();
    d.etapas.forEach((e, i) => {
      let k = g.etapas.findIndex((x, j) => !usadas.has(j) && N(x.nome) === N(e.nome));
      if (k < 0 && g.etapas.length === d.etapas.length && !usadas.has(i)) k = i;    // nome mudado no Excel, mesma posição
      if (k < 0) return;
      usadas.add(k);
      const ge = g.etapas[k], usL = new Set();
      e.id = ge.id;
      if (ge.poNome) e.poNome = ge.poNome;
      if (ge.poDispensa) e.poDispensa = ge.poDispensa;
      e.linhas.forEach(l => {
        const j = ge.linhas.findIndex((x, q) => !usL.has(q) && N(x.perigo) === N(l.perigo));
        if (j < 0) return;
        usL.add(j);
        const gl = ge.linhas[j];
        l.id = gl.id;
        if (gl.conf) l.conf = gl.conf;
        l.confirmadas = gl.confirmadas || [];
      });
    });
    d.logo = g.logo;
    for (const c of ['empresa', 'elaboracao']) if (!(d.cab[c] || '').trim()) d.cab[c] = g.cab[c] || '';
    return d;
  }
  async function gerarExcel() {
    const arqs = await montarArquivos();
    guardarDados(arqs);
    const nomes = Object.keys(arqs).sort((a, b) => (a === '[Content_Types].xml' ? -1 : b === '[Content_Types].xml' ? 1 : 0));
    return XL.zip(nomes.map(nome => ({nome, dados: arqs[nome]})));
  }
  const LOGO_VAZIO = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4AWJiYGBgAAAAAP//XRcpzQAAAAZJREFUAwAADwADJDd96QAAAABJRU5ErkJggg==';
  async function montarArquivos() {
    const arqs = Object.assign({}, await lerModelo());
    const dec = new TextDecoder(), enc = new TextEncoder();
    const T = n => dec.decode(arqs[n]), S = (n, t) => { arqs[n] = enc.encode(t); };
    const {ca, etapas, paginas, rodapeSozinho} = layout();

    // ---- logotipo escolhido: troca as imagens do modelo (A1 da Planilha Padrão e Ordem de Serviço) ----
    // A1 é imagem na célula: o Excel ajusta pela largura ou pela altura, o que limitar primeiro, sem distorcer.
    // Ordem de Serviço: a altura do modelo fica (360000 EMU) e a largura acompanha a proporção da imagem.
    // Sem logotipo escolhido, as imagens do modelo (Siderol) viram um PNG transparente de 1×1.
    {
      const png = Uint8Array.from(atob((apr.logo || LOGO_VAZIO).split(',')[1]), c => c.charCodeAt(0));
      const dv = new DataView(png.buffer), wPx = dv.getUint32(16), hPx = dv.getUint32(20);     // IHDR do PNG
      const alvos = new Set();
      for (const rel of ['xl/richData/_rels/richValueRel.xml.rels', 'xl/drawings/_rels/drawing2.xml.rels'])
        if (arqs[rel]) for (const m of T(rel).matchAll(/Target="\.\.\/media\/([^"]+)"/g)) alvos.add('xl/media/' + m[1]);
      for (const a of alvos) if (arqs[a]) arqs[a] = png;
      const dr2 = 'xl/drawings/drawing2.xml';
      if (arqs[dr2] && wPx && hPx) S(dr2, T(dr2).replace(/(<(?:xdr|a):ext cx=")\d+(" cy=")(\d+)"/g, (m, a, b, cy) => `${a}${Math.round(+cy * wPx / hPx)}${b}${cy}"`));
    }

    // ---- Planilha Padrão ----
    const ss = (T('xl/sharedStrings.xml').match(/<si>[\s\S]*?<\/si>/g) || []).map(si => (si.match(/<t[^>]*>[^<]*<\/t>/g) || []).map(t => t.replace(/<[^>]+>/g, '')).join(''));
    let s1 = T('xl/worksheets/sheet1.xml');
    const iSD = s1.indexOf('<sheetData>') + 11, fSD = s1.indexOf('</sheetData>');
    const rows = s1.slice(iSD, fSD).match(/<row [^>]*?(?:\/>|>[\s\S]*?<\/row>)/g);
    const nRow = r => +/<row r="(\d+)"/.exec(r)[1];
    const RODAPE = nRow(rows.find(x => { const m = /<c r="A\d+"[^>]*t="s"[^>]*><v>(\d+)<\/v>/.exec(x); return nRow(x) > 6 && m && /^situa/i.test(ss[+m[1]].trim()); }) || '<row r="0"');
    if (!RODAPE) throw new Error('não achei o rodapé ("Situação:") no modelo da APR');
    // estilos das células: linha comum = 1ª linha de dados do modelo; última linha = a que fica logo acima do rodapé
    const estilos = n => { const x = rows.find(y => nRow(y) === n) || ''; const o = {}; for (const m of x.matchAll(/<c r="([A-Z]+)\d+" s="(\d+)"/g)) o[m[1]] = m[2]; return o; };
    const EST = estilos(7), EST_ULT = estilos(RODAPE - 1);
    let cab = rows.filter(r => nRow(r) <= 6).join('');
    const rod = rows.filter(r => nRow(r) >= RODAPE);
    const fimRodape = Math.max(...rod.filter(x => /<v>/.test(x)).map(nRow));     // fim da área de impressão, como no modelo
    // rótulo de cada célula do cabeçalho exatamente como no modelo ("DATA:" + quebra, "ATIVIDADE: " + espaço...):
    // as fórmulas da Ordem de Serviço e do Controle de revisão tiram esse rótulo para mostrar só o valor
    const siBruto = T('xl/sharedStrings.xml').match(/<si>[\s\S]*?<\/si>/g) || [];
    const rotuloDe = (ref, padrao) => {
      const m = new RegExp(`<c r="${ref}"[^>]*t="s"[^>]*><v>(\\d+)</v>`).exec(cab);
      const t1 = m && siBruto[+m[1]] && /<t[^>]*>([^<]*)<\/t>/.exec(siBruto[+m[1]]);
      if (!t1) return padrao;
      const txt = t1[1].replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/\r\n/g, '\n');
      return /^Nº/.test(txt) ? 'Nº\n' : txt;
    };
    const R_ = {T1: rotuloDe('T1', 'DATA:\n'), A2: rotuloDe('A2', 'ÁREA:\n'), G2: rotuloDe('G2', 'DEPARTAMENTO:\n'), K2: rotuloDe('K2', 'MÁQUINA/EQUIPAMENTO:\n'),
      R2: 'ATIVIDADE:\n', T2: rotuloDe('T2', 'AVALIADO POR:\n'), X2: rotuloDe('X2', 'REVISÃO:\n'), A3: rotuloDe('A3', "EPI's BÁSICOS OBRIGATÓRIOS:\n")};
    cab = trocaCel(cab, 'T1', rico(R_.T1, ca.data));
    cab = trocaCel(cab, 'X1', `<is><t xml:space="preserve">Nº\n${esc(ca.numero)}</t></is>`);
    cab = trocaCel(cab, 'A2', rico(R_.A2, ca.area));
    cab = trocaCel(cab, 'G2', rico(R_.G2, ca.departamento));
    cab = trocaCel(cab, 'K2', rico(R_.K2, ca.maquina));
    cab = trocaCel(cab, 'R2', rico(R_.R2, ca.atividade));
    cab = trocaCel(cab, 'T2', rico(R_.T2, ca.avaliado));
    cab = trocaCel(cab, 'X2', rico(R_.X2, ca.rev));
    cab = trocaCel(cab, 'A3', rico(R_.A3, ca.episBasicos));

    const totalLinhas = etapas.reduce((s, x) => s + x.lista.length, 0);
    const primeiraDePagina = new Set(paginas.slice(1).map(p => p[0]));
    const linhasXml = [], merges = [], quebras = [];
    let r = 7;
    for (const et of etapas) {
      const {ae, i, lista, alts} = et;
      const r0 = r;
      if (primeiraDePagina.has(et)) quebras.push(r0 - 1);
      const epc = lista.map(l => l ? textoEpc(l) : ''), epis = lista.map(l => l ? textoEpis(l) : ''), met = lista.map(l => metodoDe(l, ca));
      lista.forEach((l, k) => {
        const risco = l ? riscoDe(l) : null;
        const st = r - 7 === totalLinhas - 1 ? EST_ULT : EST;
        const S_ = L => st[L] || EST[L] || '0';
        const igualAcima = lst => k > 0 && lst[k] === lst[k - 1];     // EPC / EPI / método iguais em linhas seguidas ficam numa célula só
        const c = [];
        c.push(cel('A' + r, S_('A'), k === 0 ? i + 1 : ''));
        for (const L of 'BCDEFGH') c.push(cel(L + r, S_(L), k === 0 && L === 'B' ? ae.nome || '' : ''));
        c.push(cel('I' + r, S_('I'), k === 0 ? ae.situacao || 'R' : ''));
        c.push(cel('J' + r, S_('J'), l ? l.perigo : ''));
        for (const L of 'KLMNOP') c.push(cel(L + r, S_(L), L === 'K' && l ? l.risco : ''));
        c.push(cel('Q' + r, S_('Q'), igualAcima(epc) ? '' : epc[k]));
        c.push(cel('R' + r, S_('R'), igualAcima(epis) ? '' : epis[k]));
        c.push(cel('S' + r, S_('S'), igualAcima(met) ? '' : met[k]));
        c.push(cel('T' + r, S_('T'), l && numero(l.po) ? numero(l.po) : ''));
        c.push(cel('U' + r, S_('U'), l && numero(l.pd) ? numero(l.pd) : ''));
        c.push(cel('V' + r, S_('V'), risco ? {f: `T${r}*U${r}`, v: risco} : ''));
        c.push(cel('W' + r, S_('W'), l ? l.medidas.replace(/\s+$/, '') : ''), cel('X' + r, S_('X'), ''));
        linhasXml.push(`<row r="${r}" spans="1:27" s="4" customFormat="1" ht="${alts[k].toFixed(1)}" customHeight="1" x14ac:dyDescent="0.25">${c.join('')}</row>`);
        merges.push(`K${r}:P${r}`, `W${r}:X${r}`);
        r++;
      });
      const r1 = r - 1;
      merges.push(`B${r0}:H${r1}`);
      if (r1 > r0) merges.push(`A${r0}:A${r1}`, `I${r0}:I${r1}`);
      for (const [L, lst] of [['Q', epc], ['R', epis], ['S', met]])
        for (let k = 0; k < lst.length;) { let f = k; while (f + 1 < lst.length && lst[f + 1] === lst[k]) f++; if (f > k) merges.push(`${L}${r0 + k}:${L}${r0 + f}`); k = f + 1; }
    }
    const ultDados = r - 1, d = r - RODAPE;
    if (rodapeSozinho) quebras.push(ultDados);

    const mergesModelo = [...s1.matchAll(/<mergeCell ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"\/>/g)];
    const todos = mergesModelo.filter(m => +m[2] <= 6).map(m => `${m[1]}${m[2]}:${m[3]}${m[4]}`)
      .concat(mergesModelo.filter(m => +m[2] >= RODAPE).map(m => `${m[1]}${+m[2] + d}:${m[3]}${+m[4] + d}`), merges);
    const ultima = Math.max(...rows.map(nRow)) + d;
    s1 = s1.slice(0, iSD) + cab + linhasXml.join('') + desloca(rod.join(''), d) + s1.slice(fSD);
    s1 = s1.replace(/<dimension ref="[^"]*"\/>/, `<dimension ref="A1:X${ultima}"/>`);
    s1 = s1.replace(/<mergeCells count="\d+">[\s\S]*?<\/mergeCells>/, `<mergeCells count="${todos.length}">${todos.map(m => `<mergeCell ref="${m}"/>`).join('')}</mergeCells>`);
    // cor do risco: as mesmas regras do modelo, valendo para todas as linhas
    const cf = (s1.match(/<conditionalFormatting[\s\S]*?<\/conditionalFormatting>/) || [''])[0].replace(/sqref="[^"]*"/, `sqref="V7:V${Math.max(7, ultDados)}"`);
    let primeiro = true;
    s1 = s1.replace(/<conditionalFormatting[\s\S]*?<\/conditionalFormatting>/g, () => { const t = primeiro ? cf : ''; primeiro = false; return t; });
    const brk = quebras.length ? `<rowBreaks count="${quebras.length}" manualBreakCount="${quebras.length}">${quebras.map(q => `<brk id="${q}" max="23" man="1"/>`).join('')}</rowBreaks>` : '';
    s1 = s1.replace(/<rowBreaks[\s\S]*?<\/rowBreaks>/, brk);
    s1 = s1.replace(/<selection [^>]*\/>/, '<selection activeCell="A7" sqref="A7"/>');
    S('xl/worksheets/sheet1.xml', s1);

    // retângulos da legenda de risco (sobre o rodapé)
    const dr = 'xl/drawings/drawing1.xml';
    if (arqs[dr]) S(dr, T(dr).replace(/<xdr:row>(\d+)<\/xdr:row>/g, (m, n) => `<xdr:row>${+n >= RODAPE - 1 ? +n + d : n}</xdr:row>`));

    // ---- Ordem de Serviço: valores guardados das fórmulas ----
    let s3 = T('xl/worksheets/sheet3.xml');
    for (const [ref, v] of [['P3', ca.numero], ['E4', ca.departamento], ['P4', ca.rev], ['E5', ca.maquina], ['P5', ca.data], ['E6', ca.atividade]]) s3 = trocaValor(s3, ref, v);
    // Empresa (E3) e Elaboração (J3): no modelo eram textos fixos (Siderol / Erisvaldo); agora vêm do cabeçalho, vazios por padrão
    for (const [ref, v] of [['E3', apr.cab.empresa], ['J3', apr.cab.elaboracao]])
      s3 = (v || '').trim() ? trocaCel(s3, ref, `<is><t xml:space="preserve">${esc(v.trim())}</t></is>`) : trocaCel(s3, ref, '', '');
    // a atividade sai em duas linhas ("ATIVIDADE:" e o nome embaixo), como os outros campos; a fórmula do modelo
    // só tirava "ATIVIDADE: " com espaço, então ela passa a tirar o rótulo das duas formas
    s3 = s3.replace(/(<c r="E6"[^>]*>)<f>[^<]*<\/f>/, (m, a) => `${a}<f>TRIM(SUBSTITUTE(SUBSTITUTE('Planilha Padrão'!R2,"ATIVIDADE:"&amp;CHAR(10),"",1),"ATIVIDADE:","",1))</f>`);    S('xl/worksheets/sheet3.xml', s3);

    // ---- Controle de revisão: linhas 8 a 26 do modelo ----
    let s4 = T('xl/worksheets/sheet4.xml');
    s4 = trocaValor(s4, 'B6', 'APR Nº: ' + ca.numero);
    s4 = trocaValor(s4, 'C6', 'Departamento: ' + ca.departamento);
    s4 = trocaValor(s4, 'D6', 'Área: ' + ca.area);
    s4 = trocaValor(s4, 'D27', 'Data de Emissão: ' + ca.data);
    const linhas4 = s4.match(/<row [^>]*?(?:\/>|>[\s\S]*?<\/row>)/g);
    const linha4 = n => linhas4.find(x => nRow(x) === n) || '';
    const est4 = n => { const o = {}; for (const m of linha4(n).matchAll(/<c r="([A-Z]+)\d+" s="(\d+)"/g)) o[m[1]] = m[2]; return o; };
    const vazia = [...Array(19).keys()].map(k => 8 + k).find(n => !/<v>/.test(linha4(n)));
    const E4D = est4(9), E4V = est4(vazia || 26);
    const hist = apr.historico.filter(h => (h.motivo || '').trim() || h.data).slice(0, 19);
    for (let k = 0; k < 19; k++) {
      const rr = 8 + k, h = hist[k];
      let xmlLinha;
      if (h) {
        const ht = Math.max(1, linhas(h.motivo, Math.trunc((256 * 61.33203125 + 18) / 256 * 7), '13.33px Arial')) * 13.2;
        xmlLinha = `<row r="${rr}" spans="1:5"${ht > 13.2 ? ` ht="${ht.toFixed(1)}" customHeight="1"` : ''} x14ac:dyDescent="0.25"><c r="A${rr}" s="${E4D.A || 0}"/>${cel('B' + rr, E4D.B, serial(h.data) ?? '')}${cel('C' + rr, E4D.C, h.motivo || '')}${cel('D' + rr, E4D.D, h.responsavel || '')}${cel('E' + rr, E4D.E, h.origem || '')}</row>`;
      } else if (/<v>/.test(linha4(rr))) {
        xmlLinha = `<row r="${rr}" spans="1:5" x14ac:dyDescent="0.25"><c r="A${rr}" s="${E4V.A || 0}"/><c r="B${rr}" s="${E4V.B}"/><c r="C${rr}" s="${E4V.C}"/><c r="D${rr}" s="${E4V.D}"/><c r="E${rr}" s="${E4V.E}"/></row>`;
      } else continue;
      s4 = s4.replace(new RegExp(`<row r="${rr}"[^>]*?(?:/>|>[\\s\\S]*?</row>)`), xmlLinha);
    }
    S('xl/worksheets/sheet4.xml', s4);

    // ---- margens e centralização de cada aba: 0,5 cm em todas, menos a Matriz de Risco (1,0 cm);
    //      Ordem de Serviço e Controle de revisão centralizados na folha ----
    const caminhoAba = nome => {
      const wbx = T('xl/workbook.xml'), rid = (new RegExp(`<sheet [^>]*name="${nome}"[^>]*r:id="(rId\\d+)"`).exec(wbx) || [])[1];
      const alvo = rid && (new RegExp(`Id="${rid}"[^>]*Target="([^"]+)"|Target="([^"]+)"[^>]*Id="${rid}"`).exec(T('xl/_rels/workbook.xml.rels')) || []);
      return alvo ? 'xl/' + (alvo[1] || alvo[2]).replace(/^\/?(xl\/)?/, '') : null;
    };
    const MARGEM_ABA = {'Matriz de Risco': 10};
    const CENTRO_ABA = {'Ordem de Serviço': true, 'Controle de revisão': true, 'Matriz de Risco': true};
    for (const nome of ['Planilha Padrão', 'Perigos', 'Ordem de Serviço', 'Controle de revisão', 'Matriz de Risco']) {
      const n = caminhoAba(nome);
      if (!n || !arqs[n]) continue;
      const M = (MARGEM_ABA[nome] || MARGEM_MM) / 25.4;
      let x = T(n).replace(/<pageMargins [^>]*\/>/, `<pageMargins left="${M}" right="${M}" top="${M}" bottom="${M}" header="${M}" footer="${M}"/>`);
      if (CENTRO_ABA[nome]) {
        const po = '<printOptions horizontalCentered="1" verticalCentered="1"/>';
        x = /<printOptions[^>]*\/>/.test(x) ? x.replace(/<printOptions[^>]*\/>/, po) : x.replace(/<pageMargins /, po + '<pageMargins ');
      }
      S(n, x);
    }

    // ---- pasta de trabalho: área de impressão, recalcular ao abrir, sem a cadeia de cálculo antiga ----
    let wb = T('xl/workbook.xml');
    wb = wb.replace(/(<definedName name="_xlnm\.Print_Area" localSheetId="0">[^<]*!\$A\$1:\$X\$)\d+/, `$1${fimRodape + d}`);
    wb = wb.replace(/<calcPr ([^>]*?)\/>/, (m, a) => `<calcPr ${a.replace(/\s*fullCalcOnLoad="\d"/, '')} fullCalcOnLoad="1"/>`);
    S('xl/workbook.xml', wb);
    delete arqs['xl/calcChain.xml'];
    S('xl/_rels/workbook.xml.rels', T('xl/_rels/workbook.xml.rels').replace(/<Relationship [^>]*Target="calcChain\.xml"\/>/, ''));
    S('[Content_Types].xml', T('[Content_Types].xml').replace(/<Override PartName="\/([^"]+)"[^>]*\/>/g, (m, p) => arqs[p] ? m : ''));
    return arqs;
  }
  async function exportar(semConferir) {
    if (!semConferir && validar().length) { relatorio(() => exportar(true)); return; }
    aviso('Montando a APR...');
    let blob;
    try { blob = await gerarExcel(); }
    catch (e) { console.error(e); alert('Não consegui montar a APR: ' + erroAmigavel(e)); return; }
    // o "Salvar" da APR já grava este mesmo Excel: a exportação sugere "- cópia" (05/11/2026)
    const nome = nomeAPR() + ' - cópia.xlsx';
    if (window.showSaveFilePicker) {
      try {
        const h = await window.showSaveFilePicker({id: 'exportar-apr', suggestedName: nome, types: [{description: 'Planilha do Excel', accept: {'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx']}}]});
        const w = await h.createWritable(); await w.write(blob); await w.close();
        aviso('APR salva: ' + h.name); return;
      } catch (e) { if (e.name === 'AbortError') return; }
    }
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = nome;
    document.body.append(a); a.click(); a.remove();
    aviso('APR salva na pasta Downloads: ' + nome);
  }

  // ---------------- APR antiga em Excel ----------------
  async function lerPlanilhas(buf) {
    // .xlsx com senha para abrir: fica guardado criptografado (dentro de um arquivo no formato do .xls antigo)
    const u8 = new Uint8Array(buf);
    if (u8[0] === 0xD0 && u8[1] === 0xCF && new TextDecoder('utf-16le').decode(u8.subarray(0, Math.min(u8.length, 1 << 20))).includes('EncryptedPackage'))
      throw new Error('a planilha está protegida por senha. Abra no Excel, tire a senha (Arquivo > Informações > Proteger pasta de trabalho > Criptografar com senha, apagar a senha), salve e abra de novo');
    if (typeof XLS !== 'undefined' && XLS.ehXLS(buf)) return XLS.lerPlanilhas(buf);   // Excel antigo (.xls)
    const {lerZip, xml} = IMPORTAR.ajuda;
    const z = await lerZip(buf);
    const ss = [];
    const tss = await z.texto('xl/sharedStrings.xml');
    if (tss) for (const si of xml(tss).getElementsByTagNameNS('*', 'si')) ss.push([...si.getElementsByTagNameNS('*', 't')].filter(t => !t.closest('rPh')).map(t => t.textContent).join(''));
    const wb = xml(await z.texto('xl/workbook.xml'));
    const rels = {};
    for (const r of xml(await z.texto('xl/_rels/workbook.xml.rels')).getElementsByTagNameNS('*', 'Relationship')) rels[r.getAttribute('Id')] = r.getAttribute('Target');
    const folhas = [];
    for (const s of wb.getElementsByTagNameNS('*', 'sheet')) {
      const id = s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id') || s.getAttribute('r:id');
      const alvo = 'xl/' + (rels[id] || '').replace(/^\/?(xl\/)?/, '');
      folhas.push({nome: s.getAttribute('name'), ler: async () => {
        const x = xml(await z.texto(alvo));
        const cel = {};
        for (const c of x.getElementsByTagNameNS('*', 'c')) {
          const t = c.getAttribute('t'), v = c.getElementsByTagNameNS('*', 'v')[0];
          let val = '';
          if (t === 'inlineStr') val = [...c.getElementsByTagNameNS('*', 't')].map(n => n.textContent).join('');
          else if (v) val = t === 's' ? ss[+v.textContent] || '' : v.textContent;
          if (t === 'e') val = '';
          if (val !== '') cel[c.getAttribute('r')] = val;
        }
        // célula mesclada: as outras células do bloco valem o mesmo que a primeira (também quando ocupa várias colunas)
        const nCol = s => [...s].reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0), letra = n => { let s = ''; while (n) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; };
        for (const m of x.getElementsByTagNameNS('*', 'mergeCell')) {
          const r = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(m.getAttribute('ref') || '');
          if (!r || !cel[r[1] + r[2]]) continue;
          for (let n = +r[2]; n <= +r[4]; n++) for (let k = nCol(r[1]); k <= nCol(r[3]); k++) { const ref = letra(k) + n; if (ref !== r[1] + r[2] && !(ref in cel)) cel['~' + ref] = cel[r[1] + r[2]]; }
        }
        return cel;
      }});
    }
    return folhas;
  }
  const porLinha = cel => { const L = {}; for (const [k, v] of Object.entries(cel)) { const m = /^([A-Z]+)(\d+)$/.exec(k); if (!m) continue; (L[+m[2]] = L[+m[2]] || {})[m[1]] = v; } return L; };
  function dataISO(t) {
    t = String(t || '').trim();
    const m = /(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(t);
    if (m) return `${m[3]}-${pad2(m[2])}-${pad2(m[1])}`;
    if (/^\d{4,5}(\.\d+)?$/.test(t)) return new Date(Date.UTC(1899, 11, 30) + Math.floor(+t) * 86400000).toISOString().slice(0, 10);
    return '';
  }
  const depoisDe = (t, rotulo) => String(t || '').replace(new RegExp('^\\s*' + rotulo + '\\s*:?\\s*', 'i'), '').trim();
  async function ehAPR(arquivo) {
    try { const c = await (await lerPlanilhas(await arquivo.arrayBuffer()))[0].ler(); return Object.values(c).some(v => /AN[ÁA]LISE PRELIMINAR DE RISCO/i.test(v)); }
    catch (e) { return false; }
  }
  async function lerExcel(arquivo) {
    const folhas = await lerPlanilhas(await arquivo.arrayBuffer());
    const cel0 = await folhas[0].ler(), L = porLinha(cel0);
    const nums = Object.keys(L).map(Number).sort((a, b) => a - b);
    const d = novaAPR(), c = d.cab;
    const hr0 = nums.find(n => Object.values(L[n]).some(v => N(v) === 'perigo')) || 4;
    const umaLinha = s => String(s || '').replace(/\s+/g, ' ').trim();
    // cabeçalho: rótulo e valor na mesma célula ("ÁREA: Oficina") ou o valor na célula de baixo (modelo antigo)
    const ROT_CAB = /^(data|n[ºo°]|area|departamento|maquina|atividade|avaliado|revisao|epi s basicos)/;
    for (const n of nums.filter(n => n < hr0)) for (const [Lt, v0] of Object.entries(L[n])) {
      let t = String(v0); const nt = N(t);
      if (ROT_CAB.test(nt) && /:\s*$/.test(t.trim())) {
        const baixo = L[n + 1] && L[n + 1][Lt];
        if (baixo && !ROT_CAB.test(N(baixo))) t = t.trim() + ' ' + String(baixo);
      }
      // "APR 31", "APR.MAN.12", "APR.AMP.010", "APR.TI.11"
      const mApr = /^APR[\s.]*(?:N[º°o]?\s*)?(?:([A-Z]{2,4})[\s.\-–]+)?(\d+[A-Z]?)\b/.exec(t.trim());
      if (mApr && !c.numero) { c.numero = (mApr[1] ? mApr[1] + '.' : '') + mApr[2]; continue; }
      if (nt.startsWith('data')) {
        c.data = dataISO(depoisDe(t, 'DATA'));
        // "DATA:" vazio e a data na célula à direita (às vezes como número de série do Excel: 44298)
        if (!c.data && !depoisDe(t, 'DATA')) {
          const prox = Object.entries(L[n]).filter(([k]) => k.length > Lt.length || (k.length === Lt.length && k > Lt)).sort((a, b) => a[0].length - b[0].length || (a[0] < b[0] ? -1 : 1))[0];
          if (prox && /^(\d{5}(\.\d+)?|\d{1,2}\/\d{1,2}\/\d{4})$/.test(String(prox[1]).trim())) c.data = dataISO(prox[1]);
        }
      }
      else if (/^\d{1,3}[A-Z]?$/.test(t.trim()) && n === nums[0] && !c.numero) c.numero = t.trim();   // só "14", no lugar do Nº
      else if (/^n[ºo°]/i.test(t.trim())) {         // "Nº: APR 116" = 116
        const v = depoisDe(t, 'N[ºo°]?'), m2 = /^APR[\s.]*(?:([A-Z]{2,4})[\s.\-–]+)?(\d+[A-Z]?)\b/i.exec(v);
        c.numero = m2 ? (m2[1] ? m2[1] + '.' : '') + m2[2] : v;
      }
      else if (nt.startsWith('area')) c.area = umaLinha(depoisDe(t, '[ÁA]REA'));
      else if (nt.startsWith('departamento')) c.departamento = umaLinha(depoisDe(t, 'DEPARTAMENTO'));
      else if (nt.startsWith('maquina')) c.maquina = umaLinha(depoisDe(t, 'M[ÁA]QUINA/EQUIPAMENTO'));
      else if (nt.startsWith('atividade')) c.atividade = umaLinha(depoisDe(t, 'ATIVIDADE'));
      else if (nt.startsWith('avaliado')) c.avaliado = umaLinha(depoisDe(t, 'AVALIADO POR'));
      else if (nt.startsWith('revisao')) c.revisao = pad2(parseInt(depoisDe(t, 'REVIS[ÃA]O'), 10) || 0);
      else if (nt.startsWith('epi s basicos')) c.episBasicos = depoisDe(t, "EPI'?s B[ÁA]SICOS OBRIGAT[ÓO]RIOS").replace(/\s*\n\s*/g, ' ');
    }
    // colunas pelos títulos
    const hr = nums.find(n => Object.values(L[n]).some(v => N(v) === 'perigo'));
    if (!hr) throw new Error('não achei a linha de títulos (Perigo, Risco, Medidas de Controle)');
    const col = {};
    for (const n of [hr, hr + 1]) for (const [Lt, v] of Object.entries(L[n] || {})) {
      const t = N(v);
      if (t === 'n' || t === 'no') col.num = Lt;
      else if (t.startsWith('etapa')) col.etapa = Lt;
      else if (t.startsWith('situacao')) col.sit = Lt;
      else if (t === 'perigo') col.perigo = Lt;
      else if (t.startsWith('risco efeito')) col.risco = Lt;
      else if (t.startsWith('epc')) col.epc = Lt;
      else if (t.startsWith('epi')) col.epi = Lt;
      else if (t.startsWith('metodo')) col.metodo = Lt;
      else if (t === 'po') col.po = Lt;
      else if (t === 'pd') col.pd = Lt;
      else if (t.startsWith('medidas')) col.medidas = Lt;
    }
    for (const k of ['num', 'etapa', 'perigo', 'medidas']) if (!col[k]) throw new Error('não achei a coluna "' + k + '" da APR');
    const na = s => /^(n\/?a|-)$/i.test(String(s).trim()) ? '' : String(s).trim();
    d.etapas = [];
    let atual = null;
    for (const n of nums.filter(n => n > hr + 1)) {
      const R = L[n] || {}, v = k => col[k] ? String(R[col[k]] ?? '').trim() : '';
      const mesclado = k => col[k] ? String(cel0['~' + col[k] + n] ?? '').trim() : '';
      if (/^situa/i.test(String(R.A || R[col.num] || '').trim())) break;           // rodapé
      if (v('num') && /^\d+$/.test(v('num'))) {
        atual = novaEtapa(v('etapa').replace(/\s+/g, ' ')); atual.situacao = /^nr/i.test(v('sit')) ? 'NR' : 'R'; atual.linhas = []; d.etapas.push(atual);
      }
      if (!atual || !v('perigo')) continue;
      const l = novaLinha(v('perigo').replace(/\s+/g, ' '));
      // EPC / EPI mesclados em várias linhas: valem para todas elas
      Object.assign(l, {risco: v('risco'), epc: na(v('epc') || mesclado('epc')), epis: na(v('epi') || mesclado('epi')), metodo: umaLinha(v('metodo') || mesclado('metodo')),
        po: numero(v('po')) ? String(numero(v('po'))) : '', pd: numero(v('pd')) ? String(numero(v('pd'))) : '', medidas: v('medidas')});
      // mesmo perigo da tabela com outra pontuação/maiúsculas: passa a usar o nome exato da tabela (risco e PD da APR ficam)
      const tab = perigoDaTabela(l.perigo, l.pd);
      if (tab) l.perigo = tab.perigo;
      atual.linhas.push(l);
    }
    for (const e of d.etapas) if (!e.linhas.length) e.linhas.push(novaLinha());
    if (!d.etapas.length) d.etapas.push(novaEtapa());
    // método: o mais comum vai para o cabeçalho (se não for o "PO.nº" padrão); só os perigos diferentes ficam com o seu
    const todas = d.etapas.flatMap(e => e.linhas), contaM = new Map();
    for (const l of todas) if (l.metodo) contaM.set(N(l.metodo), (contaM.get(N(l.metodo)) || 0) + 1);
    const padraoM = N('PO.' + (c.numero || 'XX.000'));         // empate: fica o "PO.nº" padrão
    const comum = [...contaM].sort((a, b) => b[1] - a[1] || (b[0] === padraoM) - (a[0] === padraoM))[0];
    const mComum = comum ? todas.find(l => N(l.metodo) === comum[0]).metodo : '';
    if (mComum && N(mComum) !== padraoM) c.metodo = mComum;
    for (const l of todas) if (!l.metodo || N(l.metodo) === N(mComum)) l.metodo = '';
    // controle de revisão
    d.historico = [];
    const fc = folhas.find(f => /controle de revis/i.test(f.nome));
    if (fc) {
      const lin = porLinha(await fc.ler()), ns = Object.keys(lin).map(Number).sort((a, b) => a - b);
      let h = ns.find(n => Object.values(lin[n]).some(v => N(v) === 'motivo'));
      const cm = {};
      if (h) for (const [Lt, v] of Object.entries(lin[h])) { const t = N(v); if (t === 'data') cm.data = Lt; else if (t === 'motivo') cm.motivo = Lt; else if (t.startsWith('respons')) cm.resp = Lt; else if (t === 'origem') cm.origem = Lt; }
      else {
        // sem a linha "Data | Motivo | Responsável | Origem": a 1ª linha com data e texto ao lado define as colunas
        const ehData = v => /^(\d{5}(\.\d+)?|\d{1,2}\/\d{1,2}\/\d{4})$/.test(String(v).trim());
        const i1 = ns.find(n => { const es = Object.entries(lin[n]).sort((a, b) => a[0].length - b[0].length || (a[0] < b[0] ? -1 : 1)); return es.length >= 2 && ehData(es[0][1]) && /[a-zà-ú]{3}/i.test(es[1][1]); });
        if (i1) {
          const ls = Object.keys(lin[i1]).sort((a, b) => a.length - b.length || (a < b ? -1 : 1));
          Object.assign(cm, {data: ls[0], motivo: ls[1], resp: ls[2], origem: ls[3]});
          h = i1 - 1;
        }
      }
      if (h) {
        for (const n of ns.filter(n => n > h)) {
          const x = lin[n], motivo = String(x[cm.motivo] || '').trim();
          if (!motivo) continue;
          if (/^data de emiss/i.test(motivo)) break;
          d.historico.push({data: dataISO(x[cm.data]), motivo, responsavel: String(x[cm.resp] || '').trim(), origem: String(x[cm.origem] || '').trim()});
        }
      }
    }
    // revisão e data automáticas se baterem com o controle de revisão
    const u = d.historico.length - 1;
    c.revAuto = u >= 0 && pad2(u) === (c.revisao || '00') && (!c.data || d.historico[u].data === c.data);
    return d;
  }

  // ---------------- estilos ----------------
  document.head.append(el('style', {}, `
.abas{display:flex;gap:4px;padding:8px 10px 0;position:sticky;top:var(--alt-barra,0px);z-index:4;background:var(--ui-bg)}   /* presas logo abaixo da barra azul (--alt-barra, medida no renderEditor) */
.abas button{flex:1;border:1px solid var(--ui-borda);border-bottom:0;background:#e5e7eb;color:var(--ui-suave);border-radius:7px 7px 0 0;padding:6px 0;font:inherit;font-weight:700;cursor:pointer}
.abas button.on{background:var(--ui-card);color:var(--azul)}
.apr-topo{display:flex;gap:6px;align-items:stretch;margin:10px 12px 2px}
.apr-topo .btn{flex:1 1 0;min-width:0;text-align:center;white-space:nowrap}
.apr-topo .btn.ativo{background:#dbeafe;border-color:#93c5fd;color:var(--azul)}
.apr-estado{margin:2px 12px 0;font-size:12px;font-weight:600}.apr-estado .falta{color:#b91c1c}.apr-estado .ok{color:#15803d}
.apr-perigo{border:1px solid var(--ui-borda);border-radius:6px;background:#fff;padding:6px 8px;margin:6px 0}
.apr-per-topo{display:flex;align-items:center;gap:4px}
.apr-per-topo>div{flex:1}.apr-per-topo input{font-weight:600}
.apr-aviso{background:#fef2f2;border:1px solid #fecaca;color:#991b1b;border-radius:5px;padding:4px 6px;margin:6px 0;font-size:12px}
.apr-aviso>div{margin:2px 0}
.apr-sel{width:auto!important;max-width:220px;font-size:12px!important;padding:2px 4px!important}
.apr-aval{display:grid;grid-template-columns:1fr 1fr;gap:8px;align-items:center;margin:6px 0;font-size:13px}   /* mesmas colunas da linha EPC's | EPI's */
.apr-popd{display:grid;grid-template-columns:1fr 1fr;gap:8px}
.apr-popd label{display:flex;align-items:center;gap:5px;min-width:0}
.apr-aval select{flex:1;min-width:0;width:auto;height:32px;box-sizing:border-box}
.apr-risco{box-sizing:border-box;height:32px;line-height:32px;display:block;min-width:0;text-align:center;border-radius:5px;padding:0 8px;font-weight:700;font-size:12px;background:#f3f4f6;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.apr-tabriscos{display:flex;flex-wrap:wrap;gap:4px;align-items:center;margin:6px 0 2px;font-size:12px}
.apr-tabriscos .rot{color:var(--ui-suave)}
.apr-chip{border:1px solid var(--ui-borda);background:#f9fafb;border-radius:12px;padding:2px 9px;font:inherit;font-size:12px;cursor:pointer;text-align:left}
.apr-chip:hover{border-color:var(--ui-acento)}.apr-chip.on{background:#dcfce7;border-color:#86efac}
.apr-chip b{color:#374151}
.apr-pdtab{font-size:11.5px;color:#166534;margin:-2px 0 6px;display:flex;flex-wrap:wrap;gap:4px;align-items:center}
.apr-pdtab span{color:var(--ui-suave)}.apr-pdtab.mudou{color:#92400e}
.apr-perigo-sel{width:100%;display:flex;align-items:center;gap:6px;text-align:left;font:inherit;font-weight:600;border:1px solid var(--ui-borda);border-radius:5px;background:#fff;padding:5px 7px;cursor:pointer;color:var(--ui-texto)}
.apr-perigo-sel{box-sizing:border-box;height:32px;min-width:0}.apr-per-topo>div{min-width:0}
.apr-perigo-sel>span:first-child{flex:1;min-width:0;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.apr-perigo-sel .seta{color:var(--ui-suave);font-weight:400}
.apr-perigo-sel:hover{border-color:var(--ui-acento)}.apr-perigo-sel.vazio{color:#9ca3af;font-weight:400}
.pp-janela header input{flex:1}
.pp-lista{flex:1;overflow:auto;padding:8px 12px 12px;display:flex;flex-direction:column;gap:4px}
.pp-item{display:flex;flex-direction:column;align-items:flex-start;gap:1px;text-align:left;border:1px solid var(--ui-borda);background:#fff;border-radius:6px;padding:6px 10px;cursor:pointer;font:inherit}
.pp-item:hover{border-color:var(--ui-acento);background:#eff6ff}.pp-item.atual{border-color:#86efac;background:#f0fdf4}
.pp-item span{font-size:12px;color:var(--ui-suave)}
.pp-item.pe-item{flex-direction:row;align-items:center;gap:10px}.pe-item img{width:34px;height:34px;object-fit:contain;flex:none}
.apr-epi-linha{display:flex;gap:4px}.apr-epi-linha input{flex:1;min-width:0}.apr-epi-linha .btn{flex:none;padding:4px 7px;font-size:11.5px}
.pt-janela{height:auto}
.cf-tab td.cf-edc{width:1%;white-space:nowrap;padding-right:14px}
.cf-medida{display:flex;gap:8px;align-items:center;justify-content:space-between}
.apr-perigo-sel.fora{border-color:#f59e0b;background:#fffbeb}
.apr-fixo{box-sizing:border-box;border:1px solid var(--ui-borda);border-radius:5px;background:#f3f4f6;padding:0 7px;height:32px;line-height:30px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;cursor:default}
.apr-fixo.vazio{color:#9ca3af}
.apr-medidas{margin-top:4px;font-size:12px}
.apr-med{display:grid;grid-template-columns:14px 1fr auto;gap:4px;align-items:start;padding:1px 0}
.apr-med.ok .ic,.apr-med.conf .ic{color:#15803d;font-weight:700}.apr-med.falta .ic{color:#b91c1c;font-weight:700}
.apr-med .st{font-size:11px;color:var(--ui-suave);white-space:nowrap;display:flex;gap:3px;align-items:center}
.apr-med.falta .st{color:#b91c1c}
.apr-po{border:1px solid #bfdbfe;background:#eff6ff;border-radius:6px;padding:6px 8px;margin:8px 0;font-size:12px}
.apr-po-tit{font-weight:700;color:#1e40af;cursor:pointer}
.apr-po-sub{font-weight:700;color:#374151;margin:6px 0 2px}
.apr-po-item{display:flex;gap:5px;align-items:center;flex-wrap:wrap;padding:1px 0}
.apr-po-item .ic{font-weight:700}.apr-po-item.ok .ic{color:#15803d}.apr-po-item.falta .ic{color:#b91c1c}.apr-po-item.falta>span:nth-child(2){color:#991b1b}
.apr-po-desc{background:#fff;border:1px solid #dbeafe;border-radius:4px;padding:4px 6px;max-height:220px;overflow:auto;line-height:1.35}
.apr-po-desc>div{margin:0 0 3px}
.apr-po-desc mark{background:#bbf7d0;color:inherit;border-radius:2px}
.apr-rel-grupo{border:1px solid var(--ui-borda);border-radius:7px;margin:0 0 10px;overflow:hidden}
.apr-rel-tit{background:#f3f4f6;font-weight:700;color:var(--azul);padding:6px 10px;display:flex;justify-content:space-between;align-items:center}
.apr-rel-item{padding:5px 10px;border-top:1px solid #f1f5f9;font-size:13px;display:flex;gap:8px;justify-content:space-between;align-items:center}
.apr-rel-acoes{display:flex;gap:4px;align-items:center;flex:none}
.apr-rel-item:not(.t-campos):not(.t-cab) span:first-child{color:#b91c1c}
.apr-ok{color:#15803d;font-weight:700}
.apr-resumo{max-width:297mm;margin:0 auto 4mm;font-size:12px;padding:6px 10px;border-radius:6px;display:flex;gap:10px;align-items:center}
.apr-resumo.falta{background:#fef2f2;color:#991b1b}.apr-resumo.ok{background:#f0fdf4;color:#166534}
.apr-pagina{padding:5mm}
.apr-folha{width:${LARG_PREVIA}px}
.apr-t{border-collapse:collapse;table-layout:fixed;width:${LARG_PREVIA}px;font:12pt Arial,Helvetica,sans-serif;color:#000;border:2px solid #000}
.apr-t td{border:1px solid #000;padding:0 3px;overflow:hidden;white-space:pre-line;line-height:1.15;vertical-align:middle;cursor:pointer}
.apr-t td.c{text-align:center}.apr-t td.e{text-align:left}.apr-t .b,.apr-t .b td{font-weight:700}
.apr-t tr.cinza td{background:#C0C0C0}
.apr-t td.apr-cinza{background:#C0C0C0}
.apr-t .apr-titulo{font-size:18pt;font-weight:700;color:#339966}
.apr-t .apr-logo{max-width:100%;max-height:60px;display:block;margin:0 auto}
.apr-logo-ed{border:1px solid #d1d5db;border-radius:6px;padding:6px;margin:4px 0 6px;background:repeating-conic-gradient(#f3f4f6 0 25%,#fff 0 50%) 0 0/16px 16px;text-align:center}
.apr-logo-ed img{max-width:100%;max-height:70px;display:inline-block}
.apr-t .apr-vert{writing-mode:vertical-rl;transform:rotate(180deg);margin:0 auto}
.apr-t td[colspan="2"].e{white-space:pre-wrap}
.apr-t tr.rd td{border:0;cursor:default;font-size:12pt;padding:0 2px}
.apr-t tr.rd1 td{border-top:2px solid #000;border-bottom:2px solid #000}
.apr-t tr.rd-fim td{border-bottom:2px solid #000}
.apr-t tr.rd td.rd-l{border-left:2px solid #000}.apr-t tr.rd td.rd-r{border-right:2px solid #000}
.apr-t tr.rd td.rd-tit{font-size:14pt;font-weight:700;text-decoration:underline;background:#f2f2f2;line-height:1}
.apr-t tr.rd td.rd-leg{font-size:11pt;white-space:pre}
.apr-t tr.rd td.rd-ass{border-top:2px solid #000;border-bottom:2px solid #000;white-space:pre-line;line-height:1.15}
.apr-t tr.rd td.rd-caixas{padding:0}
.apr-t .rd-caixas>div{position:relative;height:24px}
.apr-t .rd-caixas span{position:absolute;top:0;bottom:0;display:flex;align-items:center;justify-content:center;font-weight:700;border:2px solid #000;white-space:nowrap;overflow:hidden;font-family:Arial,Helvetica,sans-serif}
.apr-t .apr-prob{background:#fee2e2!important;outline:2px solid #dc2626;outline-offset:-2px}
.apr-t .apr-nao{color:#b91c1c;text-decoration:underline wavy #dc2626}
#paginas:has(.apr-conf){zoom:1!important}
.apr-conf{max-width:1100px;margin:0 auto;font:13px/1.4 'Segoe UI',Arial,sans-serif;color:var(--ui-texto)}
.cf-topo{position:sticky;top:-16px;z-index:3;display:flex;flex-wrap:wrap;gap:10px;align-items:center;background:#fff;border-radius:8px;padding:8px 10px;margin-bottom:10px}
.cf-nav{display:flex;gap:4px;align-items:center}.cf-nav select{width:auto;max-width:420px}
.cf-prog{font-weight:700;color:#166534}.cf-po{flex:1;color:var(--ui-suave);font-size:12px;text-align:right}
.cf-card{background:#fff;border-radius:8px;padding:10px 12px;margin-bottom:10px}
.cf-h{font-weight:700;color:#374151;margin:2px 0 6px;display:flex;align-items:center;gap:6px}
.cf-par{font-weight:400;font-size:12.5px;color:var(--ui-suave)}.cf-par select{font-size:12.5px;max-width:340px}
.cf-tab{width:100%;border-collapse:collapse}
.cf-tab th{text-align:left;font-size:12px;color:var(--ui-suave);font-weight:600;padding:3px 6px}
.cf-tab td{border-top:1px solid #eef0f3;padding:5px 6px;vertical-align:middle}
.cf-tab .cf-rot{color:var(--ui-suave);font-size:12px;width:52px}
.cf-acoes{white-space:nowrap;text-align:right}.cf-acoes>*{margin-left:4px}
.cf-ok{color:#15803d;font-weight:600}.cf-falta{color:#9ca3af}
.cf-tag{font-size:11px;font-weight:700;border-radius:9px;padding:1px 8px}
.cf-tag.dif{background:#fef3c7;color:#92400e}.cf-tag.soAPR,.cf-tag.soPO{background:#fee2e2;color:#991b1b}
.cf-duas{display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1.3fr);gap:12px}
.cf-lista{border-right:1px solid #eef0f3;padding-right:10px;max-height:70vh;overflow:auto}
.cf-per{font-weight:700;color:var(--azul);margin:8px 0 3px}
.cf-med{display:grid;grid-template-columns:16px 1fr auto;gap:6px;align-items:start;padding:5px 6px;border-radius:6px;cursor:pointer;border:1px solid transparent}
.cf-med:hover{background:#f8fafc}.cf-med.sel{background:#eff6ff;border-color:#93c5fd}
.cf-ic{font-weight:700}.cf-ic.ok,.cf-ic.auto{color:#15803d}.cf-ic.nao{color:#b91c1c}.cf-ic.conferir{color:#9ca3af}
.cf-st{font-size:11px;white-space:nowrap}.cf-st.ok{color:#15803d}.cf-st.no{color:#b91c1c}.cf-st.pend{color:#9ca3af}
.cf-medida{font-size:15px;font-weight:600;background:#f8fafc;border-radius:6px;padding:6px 8px;margin-bottom:8px}.cf-per2{font-size:12px;font-weight:400;color:var(--ui-suave)}
.cf-frase{display:grid;grid-template-columns:auto 1fr auto;gap:8px;align-items:start;border:1px solid var(--ui-borda);border-radius:6px;padding:7px 8px;margin:0 0 6px}
.cf-frase.ligada{border-color:#86efac;background:#f0fdf4}
.cf-n{color:var(--ui-suave);font-size:12px;white-space:nowrap}.cf-fb{display:flex;gap:4px}
.cf-frase mark{background:#bbf7d0;color:inherit;border-radius:2px;padding:0 1px}
.cf-botoes{display:flex;flex-wrap:wrap;gap:6px;margin:6px 0}
.cf-tudo{border-top:1px dashed var(--ui-borda);padding-top:6px;margin-top:4px}
.cf-aviso{background:#fef2f2;color:#991b1b;border-radius:6px;padding:5px 8px;margin-bottom:8px;display:flex;gap:8px;align-items:center}
.cf-ok-box{background:#f0fdf4;color:#166534;border-radius:6px;padding:5px 8px;margin-bottom:8px;display:flex;gap:8px;align-items:center}
.apr-pop{height:min(88vh,860px)}
.apr-pop-corpo{flex:1;overflow:auto;padding:12px}
.apr-pop-medida{display:flex;gap:6px;align-items:center;background:#fef3c7;border-radius:6px;padding:8px 10px;font-weight:600;margin-bottom:10px}
.apr-pop .rich{max-height:none;min-height:220px}
.apr-pop .rich>div.cf-alvo{background:#fef9c3}
.cf2{max-width:1280px;font:14px/1.4 Arial,Helvetica,sans-serif;color:#1f2937}
.cf2 button,.cf2 select,.cf2 input,.cf2-pop{font-family:Arial,Helvetica,sans-serif}
.cf2-topo{display:block;padding:12px 14px 12px}
.cf2-nav{display:flex;gap:8px;align-items:center;flex-wrap:wrap}.cf2-nav select{width:auto;max-width:440px;font-size:14px}
.cf2-cont{font-size:14px;color:#6b7280}.cf2-cont b.ok{color:#256029}.cf2-cont b.pend{color:#8a4b00}
.cf2-faixa{display:grid;gap:6px;padding-top:12px}
.cf2-titulo{font:700 18px Arial,Helvetica,sans-serif;color:#5D6532;letter-spacing:.3px}
.cf2-et{display:flex;flex-direction:column;gap:5px;height:66px;min-width:0;border:1px solid #dcdfd5;border-radius:5px;background:#fff;padding:6px 8px 7px;cursor:pointer;color:#1f2937;font:700 11.5px/1.2 Arial,Helvetica,sans-serif;overflow:hidden}
.cf2-et .t{flex:1;min-height:0;display:flex;align-items:center;justify-content:center;text-align:center;text-transform:uppercase;overflow:hidden;overflow-wrap:anywhere}
.cf2-et .t>span{display:-webkit-box;-webkit-line-clamp:3;-webkit-box-orient:vertical;overflow:hidden}
.cf2-seta{align-items:center;justify-content:center;padding:0;font:700 22px/1 Arial,Helvetica,sans-serif;color:#4b5563}
.cf2-seta:disabled{color:#d1d5db;cursor:default;border-color:#eceee8}.cf2-seta:disabled:hover{border-color:#eceee8}
.cf2-et i{display:block;height:4px;border-radius:2px;background:#c9cdbf}
.cf2-et.ok i{background:#3f8f46}.cf2-et.pend i{background:#e0a030}
.cf2-et:hover{border-color:#5D6532}
.cf2-et.atual{background:#f4f5ef;border-color:#5D6532;box-shadow:0 0 0 1px #5D6532}
.cf2-card{padding:0;overflow:hidden}
.cf2-so{text-transform:uppercase;font-weight:400;font-size:12px;color:#6b7280;display:inline-flex;gap:6px;align-items:center;cursor:pointer}
.cf2-secao{text-align:center;padding:12px 14px 10px;border-bottom:1px solid #e9eae5;font:700 16px Arial,Helvetica,sans-serif;color:#5D6532;letter-spacing:.3px}
.cf2-grade{display:grid;grid-template-columns:68px minmax(0,1fr) minmax(280px,0.95fr) minmax(0,1fr)}
.cf2-grade>div{padding:10px 12px;border-top:1px solid #e9eae5;display:flex;align-items:center;gap:6px;min-width:0}
.cf2-gt,.cf2-gp{border-top:0!important;padding:12px!important}
.cf2-gp.meio{justify-content:center}
.cf2-pil{display:block;width:100%;text-align:center;border-radius:4px;padding:7px 12px;color:#fff;font-weight:700;font-size:14px}
.cf2-pil.apr{background:#5D6532}.cf2-pil.po{background:#1F4E79}
.cf2-rot{color:#6b7280;font-size:12px;text-transform:uppercase;padding-left:10px!important;padding-right:4px!important}
.cf2-c .tx{flex:1;min-width:0;overflow-wrap:anywhere}
.cf2-meio{flex-direction:column;justify-content:center;align-items:center}
.cf2-r-dif{background:#fffbf2}.cf2-r-so{background:#fffafa}
.cf2-bts{display:flex;gap:6px;flex-wrap:nowrap;justify-content:center}
.cf2-b{width:120px;box-sizing:border-box;text-align:center;text-transform:uppercase;background:#fff;border:1px solid;border-radius:4px;padding:3px 2px;font:700 12px Arial,Helvetica,sans-serif;cursor:pointer;white-space:nowrap}
.cf2-sa{display:inline-block;font-size:12px;line-height:1;transform:scale(1.15,1.45);transform-origin:center 60%}
.cf2-b.a{border-color:#5D6532;color:#4a5128}.cf2-b.a:hover{background:#f4f5ef}
.cf2-b.p{border-color:#1F4E79;color:#1F4E79}.cf2-b.p:hover{background:#eef3f8}
.cf2-vazio{flex:1;box-sizing:border-box;height:54px;display:flex;align-items:center;border:1px dashed #d9b8b3;border-radius:4px;color:#7a5651;padding:0 10px;font-size:13px;text-align:left}
.cf2-psel{position:relative;flex:1;min-width:0;box-sizing:border-box;height:54px;display:flex;flex-direction:column;justify-content:center;gap:2px;text-align:left;background:#fff;border:1px solid #d3d7c8;border-radius:4px;padding:4px 30px 4px 10px;cursor:pointer;font:15px Arial,Helvetica,sans-serif;color:#1f2937}
.cf2-psel:hover,.cf2-ed:hover{border-color:#5D6532}.cf2-psel:focus-visible,.cf2-ed:focus{outline:2px solid #5D6532;outline-offset:-1px}
.cf2-psel .l1{display:flex;gap:6px;align-items:center;min-width:0;max-width:100%}.cf2-psel .nm{white-space:nowrap;overflow:hidden;text-overflow:ellipsis;min-width:0}
.cf2-psel .l2{display:block;max-width:100%;font-size:12px;color:#6b6f63;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.cf2-ppo{flex-direction:row;align-items:center;justify-content:flex-start;gap:8px;padding:4px 28px 4px 10px}
.cf2-ppo .nm2{flex:1;min-width:0;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;line-height:1.25}
.cf2-ppo .cf2-fora{margin-left:6px;vertical-align:1px}
/* "FORA DA LISTA" na 2ª linha (05/11/2026): APR junto do "Risco:", PO embaixo do nome; o nome usa a 1ª linha inteira */
.cf2-psel .l2{display:flex;gap:6px;align-items:center}.cf2-psel .l2 .rs{min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.cf2-ppo .nm2.com-fora{display:flex;flex-direction:column;align-items:flex-start;gap:3px}
.cf2-ppo .nm2.com-fora .nm{max-width:100%;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.cf2-ppo .nm2.com-fora .cf2-fora{margin-left:0;white-space:nowrap}
.cf2-pic{flex:none;width:44px;height:44px;box-sizing:border-box;display:flex;align-items:center;justify-content:center;border:1px solid transparent;border-radius:4px;background:#fff;padding:1px;cursor:pointer}
.cf2-pic:hover{border-color:#1F4E79}.cf2-pic img{max-width:100%;max-height:100%;display:block}.cf2-pic span{font-size:9px;line-height:1.1;color:#9ca3af;text-align:center}
.cf2-tri{position:absolute;right:10px;top:50%;transform:translateY(-50%);font-size:12px;color:#5D6532}
.cf2-fora{flex:none;font-size:11px;font-weight:700;text-transform:uppercase;background:#fdf0d5;color:#8a4b00;border-radius:8px;padding:1px 6px}
.cf2-ed{flex:1;min-width:0;box-sizing:border-box;height:54px;background:#fff;border:1px solid #d3d7c8;border-radius:4px;font:15px/1.3 Arial,Helvetica,sans-serif;color:#1f2937}
input.cf2-ed{padding:0 10px}textarea.cf2-ed{padding:6px 10px;resize:none;overflow-y:auto;scrollbar-width:thin}
.cf2-riscado{text-decoration:line-through;color:#9ca3af}
.cf2-tag{display:inline-block;width:246px;box-sizing:border-box;text-align:center;font-size:12px;font-weight:700;border-radius:10px;padding:2px 6px;white-space:nowrap;border:1px solid transparent}
.cf2-tag{text-transform:uppercase}.cf2-tag.feito{width:246px}
.cf2-tag.igual,.cf2-tag.feito{color:#256029;background:#e6f2e6}
.cf2-tag.dif{background:#fdf0d5;color:#8a4b00;border-color:#f0d9a8}
.cf2-tag.so{background:#fde8e8;color:#a12a2a;border-color:#f3c2c2}
.cf2-troca{text-transform:uppercase}
.cf2-lnk{border:0;background:none;color:#6b7280;font-size:11px;text-decoration:underline;cursor:pointer;padding:0 2px;flex:none}
.cf2-lnk:hover{color:#1F4E79}
.cf2-meio .cf-par select{max-width:230px;font-size:12px}
.cf2-esc{max-width:220px;font-size:12px}
.cf2-escondidos{grid-column:1/-1;justify-content:center;color:#256029;font-size:13px}
.cf2-col{background:#fff;border-radius:8px;overflow:hidden;padding-bottom:6px}
.cf2-quadro{margin-bottom:8px;overflow:visible}.cf2-quadro .cf2-ch{border-radius:8px 8px 0 0}.cf2-ch.cf2-chc{justify-content:center}
.cf2-ch{color:#fff;font-weight:700;padding:8px 12px;display:flex;gap:8px;align-items:center;justify-content:space-between;margin-bottom:4px}
.cf2-ch.apr{background:#5D6532}.cf2-ch.po{background:#1F4E79}
.cf2-per{font:700 14px Arial,Helvetica,sans-serif;color:#5D6532;padding:10px 16px 3px}
.cf2-med{display:grid;grid-template-columns:44px minmax(0,1fr) 120px;gap:8px;align-items:center;padding:5px 16px;cursor:pointer}
.cf2-med:hover{background:#f8fafc}
.cf2-pend-i.real{background:#fefce8}
.cf2-med.real{position:sticky;top:0;z-index:2;background:#fff;box-shadow:0 2px 6px rgba(0,0,0,.15);border-left:4px solid #1F4E79;padding-left:12px}
.cf2-med .onde{font-size:12px;color:#6b6f63;text-align:right}
.cf2-mi{box-sizing:border-box;width:100%;height:36px;border:1px solid #d3d7c8;border-radius:4px;padding:0 10px;font:15px Arial,Helvetica,sans-serif;color:#1f2937;background:#fff}
.cf2-mi:focus{outline:2px solid #5D6532;outline-offset:0}
.cf2-m{display:inline-block;font-size:11px;font-weight:700;border-radius:4px;padding:2px 6px;white-space:nowrap;border:0;font-family:Arial,Helvetica,sans-serif;text-align:center}
.cf2-med .cf2-m{justify-self:start}
.cf2-m.mv{background:#e4f1e5;color:#256029}.cf2-m.mr{background:#fbe4e1;color:#9b1c13}.cf2-m.mc{background:#eceee6;color:#3d4321}
.cf2-painel{margin:2px 16px 8px 70px;padding:10px 12px;background:#f4f7fb;border:1px solid #1F4E79;border-radius:6px;cursor:default}
.cf2-pt{font:700 12px Arial,Helvetica,sans-serif;color:#1F4E79;text-transform:uppercase;margin:0 0 6px}
.cf2-pt.sub{margin-top:10px}
.cf2-pv{font-size:12px;color:#6b6f63;margin:0 0 5px}
.cf2-sf{display:grid;grid-template-columns:auto 1fr auto;gap:8px;align-items:center;border:1px solid #e5e7eb;background:#fff;border-radius:5px;padding:6px 8px;margin:0 0 5px}
.cf2-sf.ligada{border-color:#86efac;background:#f0fdf4}
.cf2-sf .n{color:#6b7280;font-size:12px;white-space:nowrap}
.cf2-sf mark{background:#bbf7d0;color:inherit;border-radius:2px;padding:0 1px}
.cf2-vb{font:700 12px Arial,Helvetica,sans-serif;text-transform:uppercase;padding:4px 9px;border:1px solid #1F4E79;color:#1F4E79;background:#fff;border-radius:4px;cursor:pointer;white-space:nowrap}
.cf2-vb.lig{border-color:#256029;color:#256029;background:#e4f1e5}
.cf2-p4{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:6px;margin-top:10px}
.cf2-pb{box-sizing:border-box;font:700 12px Arial,Helvetica,sans-serif;text-transform:uppercase;padding:6px 4px;border:1px solid;border-radius:4px;cursor:pointer;text-align:center;background:#fff}
.cf2-pb.a{border-color:#1F4E79;color:#1F4E79}
.cf2-pb.c{border-color:#c9ccbd;color:#3d4321;background:#eceee6}
.cf2-pb.c.on{border-color:#3d4321}
.cf2-leg{display:flex;gap:18px;flex-wrap:wrap;padding:10px 16px 12px;font-size:12px;color:#6b6f63}
.cf2-pend{margin:0 16px 16px;border:1px solid #eeaaa2;background:#fffafa;border-radius:6px;padding:8px 10px}
.cf2-pend.ok{border-color:#bfe0c2;background:#f5fbf5;color:#256029;font-weight:700}
.cf2-pend-t{font:700 15px Arial,Helvetica,sans-serif;color:#9b1c13;margin-bottom:6px}
.cf2-pend-i{display:grid;grid-template-columns:minmax(0,1fr) 170px 170px;gap:6px;align-items:center;padding:5px 4px;border-top:1px solid #f3d3cf}
.cf2-pend-tx{display:flex;gap:8px;align-items:center;cursor:pointer;min-width:0}
.cf2-pper{font-size:12px;color:#6b7280}
.cf2-dica{background:#fff;border-radius:6px;padding:6px 10px;font-size:12px;color:#6b7280}
/* alinhamento (28/10/2026, aprovado pelo usuário): botões com 32 px, FECHAR CONFERÊNCIA em maiúsculo, os dois
   blocos com a mesma faixa de título e a coluna "onde está no PO" das medidas com o texto centralizado */
.cf2-b,.cf2-pb,.cf2-vb,.cf2 .cf2-fechar{box-sizing:border-box;height:32px;display:inline-flex;align-items:center;justify-content:center;gap:3px}
.cf2 .cf2-fechar{text-transform:uppercase;font:700 12px Arial,Helvetica,sans-serif;padding:0 12px;min-width:170px}
.cf2-b.t{width:246px;margin-top:6px;border-color:#9ca3af;color:#4b5563}.cf2-b.t:hover{background:#f3f4f6}
.cf2-secao{background:#5D6532;color:#fff;font-size:14px;text-align:center;padding:8px 12px;border-bottom:0;border-radius:8px 8px 0 0}
.cf2-med .onde{text-align:center}
/* ASSOCIAR (05/11/2026): botão dentro do quadro vazio (o quadro mantém os 54 px) e a lista presa à tela */
.cf2-vazio.cf2-vazio2{flex-direction:column;justify-content:center;align-items:stretch;gap:3px;padding:3px 8px;text-align:center}
.cf2-vazio2 .cf2-assoc{width:100%;height:26px;margin-top:0;border-style:dashed;text-transform:uppercase}
.cf2-vazio2 .cf2-assoc:hover{border-style:solid}
.cf2-b.p.cf2-sug{box-shadow:0 0 0 3px #d9e3b8}
.cf2-shint{font-size:11px;color:#6b7280;text-align:center;margin-top:4px}
.cf2-apop{position:fixed;z-index:1000;box-sizing:border-box;background:#fff;border:1px solid #c9ccbd;border-radius:6px;box-shadow:0 6px 18px rgba(0,0,0,.18);padding:8px;display:flex;flex-direction:column;gap:6px;font-family:Arial,Helvetica,sans-serif;max-height:60vh;overflow-y:auto}
.cf2-apt{font:700 12px Arial,Helvetica,sans-serif;color:#6b7280;text-transform:uppercase;text-align:center}
.cf2-aop{display:flex;align-items:center;gap:8px;width:100%;box-sizing:border-box;min-height:36px;border:1px solid #d3d7c8;border-radius:4px;background:#fff;padding:4px 8px;cursor:pointer;font:14px Arial,Helvetica,sans-serif;color:#1f2937;text-align:left}
.cf2-aop:hover,.cf2-aop:focus-visible{border-color:#5D6532;background:#f4f5ef;outline:0}
.cf2-aop img{width:28px;height:28px;object-fit:contain;flex:none}
.cf2-aop span{flex:1;min-width:0;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.cf2-av{font-size:12px;color:#6b7280;text-align:center;padding:4px}
.cf2-apop .cf2-acan{width:100%;margin-top:0;text-transform:uppercase}
.apr-perigo-sel .apr-num{color:var(--ui-suave);font-weight:700}
@media print{.apr-resumo{display:none}.apr-t .apr-prob{background:none!important;outline:0}.apr-t .apr-nao{color:inherit;text-decoration:none}}
`));

  setTimeout(iniciar, 0);
  return {_lerPlanilhas: lerPlanilhas, alinharEtapas, mudou, restaurar, editor, previa, cliquePrevia, menu, salvar, abrir, abrirArquivo, abrirDados, nova, pdf, exportar, gerarExcel, ehAPR, lerExcel, puxarDoPO, abrirConferencia, importarParaPO, histParaPO, compararDoPO, pendenciasPO,
    validar, validarPO, relatorio, situacao, frasesPO, sugestoes, popupPO, nomeAPR, mostrarArquivo, alterado: () => alterado,
    get doc() { return apr; }, set doc(d) { apr = normalizar(d); }, get conf() { return conf; }};
})();
