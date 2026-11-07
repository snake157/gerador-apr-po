// Menu > ADMINISTRADOR (07/11/2026, tela aprovada): com a senha de aprovador, inclui, edita e exclui
// os perigos da tabela (Perigo, Risco, PD) e os pictogramas de perigo e de EPI.
// As mudanças ficam marcadas e só são gravadas em SALVAR ALTERAÇÕES, todas de uma vez no GitHub,
// pelo carteiro do Cloudflare (functions/api/[acao].js, rota "administrar").
'use strict';

const ADMIN = (() => {
  const A = () => ENVIOS.ajuda;
  const estilo = document.createElement('style');
  estilo.textContent = `
.janela.envio.admin{width:min(1000px,96vw);height:min(90vh,860px);max-height:none}
.admin .adm-abas{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));border:1px solid var(--env-cor);border-radius:5px;overflow:hidden;margin:12px 12px 0;flex:none}
.admin .adm-abas button{border:0;background:#fff;color:var(--env-cor);font:700 12px Arial,Helvetica,sans-serif;text-transform:uppercase;height:32px;cursor:pointer;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;padding:0 6px}
.admin .adm-abas button.sel{background:var(--env-cor);color:#fff}
.admin .adm-barra{display:grid;grid-template-columns:minmax(0,1fr) 140px;gap:8px;padding:10px 12px;flex:none}
.admin .adm-corpo{padding:8px 12px 12px;overflow:auto;flex:1;min-height:0}
.admin table{width:100%;border-collapse:collapse;font-size:12.5px;table-layout:fixed}
.admin th{position:sticky;top:-8px;background:#f6f7f3;font:700 11px Arial,Helvetica,sans-serif;text-transform:uppercase;letter-spacing:.03em;color:#374151;text-align:left;padding:7px 6px;border-bottom:1px solid var(--ui-borda);z-index:1}
.admin th.pd,.admin td.pd{text-align:center}
.admin td{padding:6px;border-bottom:1px solid var(--ui-borda);vertical-align:middle;overflow-wrap:anywhere}
.admin td.pd{font-weight:700;font-variant-numeric:tabular-nums}
.admin tr.editando td{background:#f3f5ec}
.admin tr.alterado td:first-child{box-shadow:inset 3px 0 0 #15803d}
.admin .adm-acoes{display:grid;grid-template-columns:1fr;gap:4px}
.admin .adm-acoes .btn,.admin .adm-cartao .btn{min-height:26px;font-size:11px;padding:2px 6px}
.admin .adm-grade{display:grid;grid-template-columns:repeat(auto-fill,minmax(130px,1fr));gap:14px 10px;padding-top:8px}
.admin .adm-cartao{position:relative;border:1px solid var(--ui-borda);border-radius:8px;padding:8px;display:flex;flex-direction:column;gap:6px;align-items:center;text-align:center;min-width:0}
.admin .adm-cartao img{width:84px;height:84px;object-fit:contain}
.admin .adm-cartao .nome{font-size:11.5px;font-weight:700;line-height:1.25;min-height:2.5em;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;width:100%}
.admin .adm-cartao .adm-dois{display:grid;grid-template-columns:1fr;gap:4px;width:100%;margin-top:auto}
.admin .adm-cartao.alterado{border-color:#15803d;box-shadow:inset 0 0 0 1px #15803d}
.admin .adm-cartao .env-selo{position:absolute;top:-9px;left:50%;transform:translateX(-50%);white-space:nowrap;box-shadow:0 0 0 2px #fff}
.admin td .env-selo{margin-top:3px}
.admin .adm-editor{border:1px solid var(--env-cor);border-radius:8px;padding:10px;display:grid;grid-template-columns:96px minmax(0,1fr);gap:10px;margin-bottom:8px;background:#f3f5ec}
.admin .adm-editor img,.admin .adm-editor .vazio{width:96px;height:96px;object-fit:contain;background:#fff;border-radius:6px;display:flex;align-items:center;justify-content:center;font-size:11px;color:var(--ui-suave);text-align:center}
.admin .adm-editor .campos{display:flex;flex-direction:column;gap:6px;min-width:0}
.admin .adm-status{font-size:12px;color:var(--ui-suave);text-align:center;padding:6px 12px 0;flex:none}
.admin .adm-msg{text-align:center;color:var(--ui-suave);font-size:12.5px;padding:18px 0}
`;
  document.head.append(estilo);

  const variacao = /\s\(\d+\)$/;
  let senha = '', aba = 'tabela', editando = null, tab = [], pics = {}, excluidos = 0, salvando = false;
  let corpo, busca, status, fecharJanela;

  function carregar() {
    tab = (window.TABELA_PERIGOS || []).map(r => ({perigo: String(r.perigo || '').trim(), risco: String(r.risco || '').trim(), pd: Number(r.pd) || 3, _orig: String(r.perigo || '').trim()}));
    pics = {perigos: (BIB.perigos || []).map(p => ({nome: p.nome, img: p.img, _orig: p.nome})),
      epis1: (BIB.epis1 || []).map(p => ({nome: p.nome, img: p.img, _orig: p.nome}))};
    excluidos = 0; editando = null;
  }
  const mudancas = () => excluidos + tab.filter(r => r._alt).length + pics.perigos.filter(p => p._alt).length + pics.epis1.filter(p => p._alt).length;
  function mostrarStatus() {
    const n = mudancas();
    status.textContent = n ? `${n} ${n === 1 ? 'alteração ainda não salva' : 'alterações ainda não salvas'}.` : 'Nenhuma alteração.';
  }
  const nomesTab = () => tab.map(r => r.perigo);
  const pictosDoPerigo = nome => pics.perigos.filter(p => p.nome === nome || (p.nome.startsWith(nome + ' (') && variacao.test(p.nome) && p.nome.replace(variacao, '') === nome));

  // ---------- desenho ----------
  function desenhar() {
    const {selo, div} = A();
    const q = A().N(busca.value);
    if (aba === 'tabela') {
      const cols = el('colgroup', {}, ...['36%', '40%', '9%', '15%'].map(w => el('col', {style: 'width:' + w})));
      const linhas = tab.map((r, i) => [r, i]).filter(([r]) => !q || A().N(r.perigo + ' ' + r.risco).includes(q)).map(([r, i]) => {
        if (editando === i) {
          const p = el('textarea', {'aria-label': 'Perigo'}), ri = el('textarea', {'aria-label': 'Risco'});
          p.value = r.perigo; ri.value = r.risco;
          const pd = el('select', {'aria-label': 'PD'}, ...[1, 2, 3, 4, 5].map(k => el('option', {value: k}, String(k))));
          pd.value = String(r.pd);
          return el('tr', {class: 'editando'}, el('td', {}, p), el('td', {}, ri), el('td', {}, pd),
            el('td', {}, div('adm-acoes',
              el('button', {class: 'btn primario', onclick: () => okPerigo(i, p.value, ri.value, Number(pd.value))}, 'OK'),
              el('button', {class: 'btn', onclick: () => cancelarEdicao()}, 'Cancelar'))));
        }
        return el('tr', {class: r._alt ? 'alterado' : ''},
          el('td', {}, r.perigo, r._alt ? el('br') : '', r._alt ? selo(r._alt, true) : ''), el('td', {}, r.risco), el('td', {class: 'pd'}, String(r.pd)),
          el('td', {}, div('adm-acoes',
            el('button', {class: 'btn', onclick: () => { editando = i; desenhar(); }}, 'Editar'),
            el('button', {class: 'btn', onclick: () => excluirPerigo(i)}, 'Excluir'))));
      });
      corpo.replaceChildren(el('table', {}, cols, el('thead', {}, el('tr', {}, el('th', {}, 'Perigo'), el('th', {}, 'Risco'), el('th', {class: 'pd'}, 'PD'), el('th', {}))),
        el('tbody', {}, ...linhas)), linhas.length ? '' : div('adm-msg', 'Nenhum perigo encontrado.'));
    } else {
      const lista = pics[aba];
      const cartoes = lista.map((p, i) => [p, i]).filter(([p]) => !q || A().N(p.nome).includes(q)).map(([p, i]) =>
        div('adm-cartao' + (p._alt ? ' alterado' : ''), el('img', {src: p.img, alt: ''}), el('span', {class: 'nome', title: p.nome}, p.nome),
          p._alt ? selo(p._alt, true) : '',
          div('adm-dois', el('button', {class: 'btn', onclick: () => { editando = i; desenhar(); corpo.scrollTop = 0; }}, 'Editar'),
            el('button', {class: 'btn', onclick: () => excluirPicto(i)}, 'Excluir'))));
      corpo.replaceChildren(editando !== null ? editorPicto() : '', div('adm-grade', ...cartoes), cartoes.length ? '' : div('adm-msg', 'Nenhum pictograma encontrado.'));
    }
    mostrarStatus();
  }

  // ---------- perigos ----------
  async function okPerigo(i, perigo, risco, pd) {
    perigo = perigo.trim().replace(/\s+/g, ' '); risco = risco.trim();
    if (!perigo || !risco) { alert('Preencha o PERIGO e o RISCO.'); return; }
    if (tab.some((r, k) => k !== i && A().N(r.perigo) === A().N(perigo))) { alert(`O perigo "${perigo}" já está na tabela.`); return; }
    const r = tab[i], antes = r.perigo;
    if (r.perigo === perigo && r.risco === risco && r.pd === pd) { cancelarEdicao(); return; }
    if (r._orig && antes !== perigo) {
      const ps = pictosDoPerigo(antes);
      if (ps.length && await perguntar(`"${antes}" tem ${ps.length} pictograma(s) com esse nome. Renomear também para "${perigo}"?`, {sim: 'Renomear também', nao: 'Só o perigo', titulo: 'Renomear perigo'}))
        for (const p of ps) { p.nome = perigo + p.nome.slice(antes.length); p._alt = p._alt || 'Editado'; }
    }
    Object.assign(r, {perigo, risco, pd, _alt: r._alt === 'Novo' ? 'Novo' : 'Editado'});
    editando = null; desenhar();
  }
  function cancelarEdicao() {
    const lista = aba === 'tabela' ? tab : pics[aba];
    const it = lista[editando];
    if (it && it._alt === 'Novo' && it._vazio) lista.splice(editando, 1);   // + NOVO cancelado antes do primeiro OK
    editando = null; desenhar();
  }
  async function excluirPerigo(i) {
    const r = tab[i], ps = pictosDoPerigo(r.perigo);
    let tambem = false;
    if (ps.length) {
      const k = await escolher(`"${r.perigo}" tem ${ps.length} pictograma(s) com esse nome.`, ['Excluir perigo e pictogramas', 'Excluir só o perigo', 'Voltar'], {titulo: 'Excluir perigo'});
      if (k < 0 || k === 2) return;
      tambem = k === 0;
    } else if (!(await perguntar(`Excluir "${r.perigo}" da tabela? APRs já feitas não mudam.`, {sim: 'Excluir', nao: 'Voltar', titulo: 'Excluir perigo'}))) return;
    if (r._orig) excluidos++;
    tab.splice(i, 1);
    if (tambem) for (const p of ps) { const k = pics.perigos.indexOf(p); if (p._orig) excluidos++; pics.perigos.splice(k, 1); }
    editando = null; desenhar();
  }

  // ---------- pictogramas ----------
  function editorPicto() {
    const {div} = A();
    const lista = pics[aba], p = lista[editando];
    const nome = el('input', {type: 'text', value: p.nome, list: 'admNomes', autocomplete: 'off', 'aria-label': 'Nome'});
    const dl = el('datalist', {id: 'admNomes'}, ...(aba === 'perigos' ? nomesTab() : []).map(n => el('option', {value: n})));
    const arq = el('input', {type: 'file', accept: 'image/png,image/jpeg', hidden: true});
    let nova = null;
    const imgEl = p.img ? el('img', {src: p.img, alt: ''}) : div('vazio', 'Nenhuma imagem');
    const caixa = div('adm-editor', imgEl, div('campos',
      el('span', {class: 'env-rot'}, aba === 'perigos' ? 'Nome (da tabela de perigos)' : 'Nome do EPI'), nome, dl,
      div('env-linha',
        el('button', {class: 'btn', onclick: () => arq.click()}, p.img ? 'Trocar imagem' : 'Escolher imagem'),
        el('button', {class: 'btn primario', onclick: () => okPicto(nome.value, nova)}, 'OK'),
        el('button', {class: 'btn', onclick: () => cancelarEdicao()}, 'Cancelar')), arq));
    arq.onchange = async () => {
      const f = arq.files[0]; arq.value = '';
      if (!f) return;
      if (f.size > 2 * 1024 * 1024) { alert('A imagem tem mais de 2 MB. Diminua o tamanho e tente de novo.'); return; }
      try {
        const {im} = await A().abrirImagem(f);
        nova = {mini: A().miniatura(im), arquivo: await A().lerBase64(f), nomeArquivo: f.name};
        caixa.firstChild.replaceWith(el('img', {src: nova.mini, alt: ''}));
      } catch (e) { alert('Não consegui abrir a imagem: ' + erroAmigavel(e)); }
    };
    setTimeout(() => nome.focus(), 0);
    return caixa;
  }
  function okPicto(nome, nova) {
    const lista = pics[aba], p = lista[editando];
    nome = nome.trim().replace(/\s+/g, ' ');
    if (!nome) { alert('Escreva o nome do pictograma.'); return; }
    if (/[\\/:*?"<>|[\]]/.test(nome)) { alert('O nome não pode ter \\ / : * ? " < > | [ ]'); return; }
    if (lista.some((x, k) => k !== editando && x.nome === nome)) { alert(`Já existe um pictograma "${nome}". Use outro nome (ex.: "${nome} (2)").`); return; }
    if (aba === 'perigos' && nome !== p._orig && !nomesTab().includes(nome.replace(variacao, ''))) {
      alert(`"${nome.replace(variacao, '')}" não está na tabela de perigos. Escolha um nome da lista (pode terminar com " (2)", " (3)"...).`); return;
    }
    if (!p.img && !nova) { alert('Escolha a imagem do pictograma.'); return; }
    if (nome === p.nome && !nova) { cancelarEdicao(); return; }
    p.nome = nome;
    if (nova) Object.assign(p, {img: nova.mini, mini: nova.mini, arquivo: nova.arquivo, nomeArquivo: nova.nomeArquivo});
    p._alt = p._orig ? 'Editado' : 'Novo'; delete p._vazio;
    editando = null; desenhar();
  }
  async function excluirPicto(i) {
    const p = pics[aba][i];
    if (!(await perguntar(`Excluir o pictograma "${p.nome}"? POs já feitos com ele ficam sem a imagem ao abrir.`, {sim: 'Excluir', nao: 'Voltar', titulo: 'Excluir pictograma'}))) return;
    if (p._orig) excluidos++;
    pics[aba].splice(i, 1);
    editando = null; desenhar();
  }

  function novo() {
    if (editando !== null) { alert('Termine (OK) ou cancele a edição aberta antes.'); return; }
    busca.value = '';
    if (aba === 'tabela') tab.unshift({perigo: '', risco: '', pd: 3, _orig: null, _alt: 'Novo', _vazio: true});
    else pics[aba].unshift({nome: '', img: '', _orig: null, _alt: 'Novo', _vazio: true});
    editando = 0; desenhar(); corpo.scrollTop = 0;
  }

  // ---------- salvar ----------
  function pedido() {
    const d = {senha};
    const tabMudou = excluidos && true;   // excluir perigo ou pictograma
    if (tab.some(r => r._alt) || tab.length !== (window.TABELA_PERIGOS || []).length || tabMudou) {
      d.tabela = tab.map(r => ({perigo: r.perigo, risco: r.risco, pd: r.pd}));
      d.renomeadosTabela = Object.fromEntries(tab.filter(r => r._orig && r._orig !== r.perigo).map(r => [r._orig, r.perigo]));
    }
    d.pictos = [];
    for (const l of ['perigos', 'epis1']) {
      const ficam = new Set(pics[l].filter(p => p._orig).map(p => p._orig));
      for (const p of BIB[l] || []) if (!ficam.has(p.nome)) d.pictos.push({lista: l, acao: 'excluir', nome: p.nome});
    }
    for (const l of ['perigos', 'epis1']) for (const p of pics[l]) if (p._orig && (p.nome !== p._orig || p.arquivo))
      d.pictos.push({lista: l, acao: 'editar', nome: p._orig, novoNome: p.nome, ...(p.arquivo ? {mini: p.mini, arquivo: p.arquivo, nomeArquivo: p.nomeArquivo} : {})});
    for (const l of ['perigos', 'epis1']) for (const p of pics[l]) if (!p._orig)
      d.pictos.push({lista: l, acao: 'novo', nome: p.nome, mini: p.mini, arquivo: p.arquivo, nomeArquivo: p.nomeArquivo});
    return d;
  }
  // depois de salvar: vale já nesta página (os outros veem quando o site publicar de novo)
  function aplicarAqui(d) {
    const ren = (v, m) => { const o = window[v] = window[v] || {}; for (const k in o) if (m[o[k]]) o[k] = m[o[k]]; Object.assign(o, m); };
    if (d.tabela) { window.TABELA_PERIGOS = d.tabela; ren('PERIGOS_RENOMEADOS', d.renomeadosTabela || {}); if (typeof APR !== 'undefined' && APR.montarTabela) APR.montarTabela(); }
    const renPer = {}, renEpi = {};
    for (const o of d.pictos) if (o.acao === 'editar' && o.novoNome !== o.nome) (o.lista === 'perigos' ? renPer : renEpi)[o.nome] = o.novoNome;
    ren('PERIGOS_NOMES_ANTIGOS', renPer); ren('EPIS_NOMES_ANTIGOS', renEpi);
    for (const l of ['perigos', 'epis1']) BIB[l] = pics[l].map(p => ({nome: p.nome, img: p.img}));
    try { renderEditor(); renderPrevia(); } catch (e) { console.warn(e); }
  }
  async function salvar(botao) {
    if (salvando) return false;
    if (editando !== null) { alert('Termine (OK) ou cancele a edição aberta antes de salvar.'); return false; }
    if (!mudancas()) { aviso('Nenhuma alteração para salvar.'); return false; }
    const d = pedido();
    salvando = true;
    if (botao) { botao.disabled = true; botao.textContent = 'Salvando…'; }
    try {
      await A().chamar('administrar', d);
      aplicarAqui(d); carregar(); desenhar();
      aviso('Alterações salvas. Em cerca de 1 minuto todos veem no site.');
      return true;
    } catch (e) {
      if (e.status === 401) A().guardar('envio:senhaAdm', '');
      alert('Não consegui salvar: ' + erroAmigavel(e));
      return false;
    } finally {
      salvando = false;
      if (botao) { botao.disabled = false; botao.textContent = 'Salvar alterações'; }
    }
  }

  async function abrir() {
    senha = A().guardado('envio:senhaAdm') || await A().pedirSenha();
    if (!senha) return;
    carregar(); aba = 'tabela';
    const {div} = A();
    const abas = div('adm-abas', ...[['tabela', 'Perigos e riscos'], ['perigos', 'Pictogramas de perigo'], ['epis1', 'Pictogramas de EPI']].map(([a, t]) =>
      el('button', {type: 'button', class: a === aba ? 'sel' : '', 'data-aba': a}, t)));
    abas.onclick = ev => {
      const b = ev.target.closest('button'); if (!b || b.dataset.aba === aba) return;
      if (editando !== null) { alert('Termine (OK) ou cancele a edição aberta antes.'); return; }
      aba = b.dataset.aba; busca.value = '';
      abas.querySelectorAll('button').forEach(x => x.classList.toggle('sel', x === b));
      desenhar(); corpo.scrollTop = 0;
    };
    busca = el('input', {type: 'text', placeholder: 'Buscar...', 'aria-label': 'Buscar'});
    busca.oninput = () => desenhar();
    corpo = div('adm-corpo'); status = div('adm-status');
    const fechar = async () => {
      if (mudancas()) {
        const k = await escolher('Há alterações não salvas.', ['Salvar', 'Descartar', 'Voltar'], {titulo: 'Fechar administrador'});
        if (k === 0) { if (!(await salvar())) return; }
        else if (k !== 1) return;
      }
      fecharJanela();
    };
    fecharJanela = A().janela('admin', 'Administrador',
      el('div', {style: 'display:flex;flex-direction:column;min-height:0;flex:1'}, abas,
        div('adm-barra', busca, el('button', {class: 'btn primario', onclick: novo}, '+ Novo')), corpo, status),
      () => [
        el('button', {class: 'btn primario', onclick: ev => salvar(ev.currentTarget)}, 'Salvar alterações'),
        el('button', {class: 'btn', onclick: async () => {
          if (!mudancas() && editando === null) return;
          if (await perguntar('Descartar todas as alterações que ainda não foram salvas?', {sim: 'Descartar', nao: 'Voltar', titulo: 'Desfazer tudo'})) { carregar(); desenhar(); }
        }}, 'Desfazer tudo'),
        el('button', {class: 'btn', onclick: fechar}, 'Fechar')]);
    desenhar();
  }
  return {abrir};
})();
