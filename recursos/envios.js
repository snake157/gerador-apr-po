// Envio compartilhado (07/11/2026): qualquer pessoa manda um pictograma ou um perigo novo pelo gerador online,
// sem senha; o envio fica PENDENTE e só o aprovador (senha) aprova ou recusa em Menu > APROVAR ENVIOS.
// Quem grava no GitHub é o "carteiro" do Cloudflare (functions/api/[acao].js); depois de aprovado,
// o site publica de novo em ~1 minuto e todos passam a ver.
'use strict';

const ENVIOS = (() => {
  // no site do Cloudflare o carteiro está no mesmo endereço; aberto de outro lugar (GitHub Pages, cópia no PC), usa o do Cloudflare
  const SITE = 'https://gerador.apr-po.workers.dev';
  const API = (/\.(pages|workers)\.dev$/.test(location.hostname) ? location.origin : SITE) + '/api/';
  const PD = {1: '1 - Incômodo', 2: '2 - Primeiros socorros', 3: '3 - Reversível com afastamento', 4: '4 - Grave / irreversível parcial', 5: '5 - Fatal'};
  const guardado = k => { try { return localStorage.getItem(k) || ''; } catch (e) { return ''; } };
  const guardar = (k, v) => { try { if (v) localStorage.setItem(k, v); else localStorage.removeItem(k); } catch (e) {} };
  const N = s => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().replace(/[^A-Z0-9]+/g, ' ').trim();
  let pendentesTotal = null;
  const rotuloAprovar = () => 'APROVAR ENVIOS' + (pendentesTotal ? ` (${pendentesTotal})` : '');
  // o menu já desenhado mostra o número novo sem redesenhar a tela
  function mostrarTotal(n) {
    pendentesTotal = n;
    for (const b of document.querySelectorAll('.lista-menu button')) if (b.textContent.startsWith('APROVAR ENVIOS')) b.textContent = rotuloAprovar();
  }

  const estilo = document.createElement('style');
  estilo.textContent = `
.janela.envio{width:min(460px,94vw);height:auto;max-height:min(90vh,820px)}
.janela.envio.aprovar{width:min(560px,94vw)}
.janela.envio header{justify-content:center}
.envio .env-corpo{padding:12px;overflow:auto;display:flex;flex-direction:column;gap:10px;min-height:0}
.envio .env-campo{display:flex;flex-direction:column;gap:3px}
.envio .env-rot{font:700 11px Arial,Helvetica,sans-serif;text-transform:uppercase;letter-spacing:.04em;color:#374151}
.envio input[type=text],.envio input[type=password],.envio select,.envio textarea{box-sizing:border-box;width:100%;height:32px;border:1px solid var(--ui-borda);border-radius:5px;padding:0 8px;font:13px Arial,Helvetica,sans-serif;color:#111827;background:#fff}
.envio textarea{height:46px;padding:6px 8px;resize:vertical}
.envio .env-chave{display:grid;grid-template-columns:1fr 1fr;border:1px solid var(--env-cor);border-radius:5px;overflow:hidden}
.envio .env-chave button{border:0;background:#fff;color:var(--env-cor);font:700 12px Arial,Helvetica,sans-serif;text-transform:uppercase;height:30px;cursor:pointer}
.envio .env-chave button.sel{background:var(--env-cor);color:#fff}
.envio .env-previa{display:grid;grid-template-columns:96px minmax(0,1fr);gap:10px;align-items:center;border:1px dashed var(--ui-borda);border-radius:6px;padding:8px;min-height:96px}
.envio .env-previa img{width:96px;height:96px;object-fit:contain}
.envio .env-previa .vazio{width:96px;height:96px;display:flex;align-items:center;justify-content:center;font-size:11px;color:var(--ui-suave);text-align:center;background:#f6f7f3;border-radius:4px}
.envio .env-info{font-size:12px;color:#374151;min-width:0;overflow-wrap:anywhere;display:flex;flex-direction:column;gap:3px;align-items:flex-start}
.envio .env-selo{display:inline-block;font:700 10.5px Arial,Helvetica,sans-serif;text-transform:uppercase;letter-spacing:.04em;border-radius:999px;padding:2px 8px}
.envio .env-selo.verde{background:#dcfce7;color:#15803d}
.envio .env-selo.cinza{background:#eef0f2;color:#6b7280}
.envio .env-resumo{font-size:12px;color:#374151;background:#f6f7f3;border-radius:6px;padding:8px;display:flex;flex-direction:column;gap:3px;align-items:flex-start}
.envio .env-linha{display:grid;grid-auto-flow:column;grid-auto-columns:minmax(0,1fr);gap:8px}
.envio .btn{box-sizing:border-box;min-height:32px;text-transform:uppercase;letter-spacing:.02em;text-align:center;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.envio .env-item{display:grid;grid-template-columns:64px minmax(0,1fr);gap:10px;padding:10px 0;border-bottom:1px solid var(--ui-borda)}
.envio .env-item:last-child{border-bottom:0}
.envio .env-item > img,.envio .env-item > .env-icone{width:64px;height:64px;object-fit:contain}
.envio .env-icone{display:flex;align-items:center;justify-content:center;border-radius:6px;background:#16a34a;color:#fff;font:700 10px Arial,Helvetica,sans-serif;text-transform:uppercase}
.envio .env-dados{display:flex;flex-direction:column;gap:4px;min-width:0}
.envio .env-nome{font-weight:700;font-size:13px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.envio .env-quem{font-size:11.5px;color:var(--ui-suave);overflow-wrap:anywhere}
.envio .env-compara{display:flex;gap:6px;align-items:center;font-size:11px;color:var(--ui-suave);flex-wrap:wrap}
.envio .env-compara img{width:34px;height:34px;object-fit:contain}
.envio .env-msg{text-align:center;color:var(--ui-suave);font-size:12.5px;padding:18px 0}
`;
  document.head.append(estilo);

  const cor = () => (typeof abaAtual !== 'undefined' && abaAtual === 'apr') ? '#5D6532' : '#002060';
  const div = (cls, ...f) => el('div', {class: cls}, ...f);
  const campo = (rotulo, controle, id) => { if (id) controle.id = id; return div('env-campo', el('label', {class: 'env-rot', for: id}, rotulo), controle); };
  const selo = (txt, verde) => el('span', {class: 'env-selo ' + (verde ? 'verde' : 'cinza')}, txt);
  const dataBR = iso => { const d = new Date(iso); return isNaN(d) ? '' : d.toLocaleDateString('pt-BR') + ' ' + d.toLocaleTimeString('pt-BR', {hour: '2-digit', minute: '2-digit'}); };

  // abre a janela no #modal do gerador; devolve a função que fecha
  function janela(classe, titulo, corpo, botoes) {
    const m = $('#modal');
    const fechar = () => { m.hidden = true; m.replaceChildren(); m.onclick = null; };
    const j = el('div', {class: 'janela envio ' + classe, onclick: ev => ev.stopPropagation()},
      el('header', {}, el('b', {}, titulo)), corpo, el('div', {class: 'janela-bts'}, ...botoes(fechar)));
    j.style.setProperty('--env-cor', cor());
    m.replaceChildren(j); m.onclick = fechar; m.hidden = false;
    return fechar;
  }

  async function chamar(acao, dados) {
    let r;
    try { r = await fetch(API + acao, {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(dados)}); }
    catch (e) { throw new Error('não consegui falar com o servidor de envios. Confira a internet e tente de novo.'); }
    let j = {};
    try { j = await r.json(); } catch (e) {}
    if (!r.ok) { const x = new Error(j.erro || `o servidor de envios respondeu ${r.status}`); x.status = r.status; throw x; }
    return j;
  }

  // ---------- imagens ----------
  const lerBase64 = f => new Promise((ok, no) => { const r = new FileReader(); r.onload = () => ok(String(r.result).split(',')[1]); r.onerror = () => no(r.error); r.readAsDataURL(f); });
  function abrirImagem(f) {
    return new Promise((ok, no) => {
      const url = URL.createObjectURL(f), im = new Image();
      im.onload = () => ok({im, url}); im.onerror = () => { URL.revokeObjectURL(url); no(new Error('não consegui abrir a imagem')); };
      im.src = url;
    });
  }
  // a mesma miniatura do ATUALIZAR BIBLIOTECA: lado maior 360 px, WebP 0,88
  function miniatura(im, lado = 360) {
    const k = Math.min(1, lado / Math.max(im.naturalWidth, im.naturalHeight));
    const c = document.createElement('canvas');
    c.width = Math.max(1, Math.round(im.naturalWidth * k)); c.height = Math.max(1, Math.round(im.naturalHeight * k));
    const g = c.getContext('2d'); g.imageSmoothingQuality = 'high'; g.drawImage(im, 0, 0, c.width, c.height);
    return c.toDataURL('image/webp', 0.88);
  }
  const listaBib = tipo => (typeof BIB !== 'undefined' && BIB[tipo === 'perigo' ? 'perigos' : 'epis1']) || [];
  const nomesTabela = () => (window.TABELA_PERIGOS || []).map(r => String(r.perigo || '').trim()).filter(Boolean);

  // ---------- ENVIAR PICTOGRAMA ----------
  function enviarPictograma() {
    let tipo = 'perigo', arq = null, img = null;
    const lista = el('datalist', {id: 'envNomes'});
    const nome = el('input', {type: 'text', list: 'envNomes', autocomplete: 'off', placeholder: 'Comece a digitar para buscar'});
    const rotNome = el('label', {class: 'env-rot', for: 'envNome'}, '');
    nome.id = 'envNome';
    const bts = [el('button', {type: 'button', class: 'sel'}, 'Perigo'), el('button', {type: 'button'}, 'EPI')];
    const previa = div('env-previa'), info = div('env-info');
    const obs = el('textarea', {}), autor = el('input', {type: 'text', value: guardado('envio:autor'), placeholder: 'Nome e sobrenome'});
    const escolher = el('input', {type: 'file', accept: 'image/png,image/jpeg', hidden: true});
    const mostrar = () => {
      const existe = listaBib(tipo).find(p => p.nome === nome.value.trim());
      previa.replaceChildren(img ? el('img', {src: img.url, alt: ''}) : div('vazio', 'Nenhuma imagem'), info);
      info.replaceChildren(...(arq ? [el('span', {}, arq.name), el('span', {}, `${img.im.naturalWidth} × ${img.im.naturalHeight} px`)] : [el('span', {}, 'PNG ou JPG, até 2 MB')]),
        nome.value.trim() ? (existe ? selo('Já existe: vai substituir', false) : selo('Novo', true)) : '');
    };
    const trocarTipo = t => {
      tipo = t; bts[0].classList.toggle('sel', t === 'perigo'); bts[1].classList.toggle('sel', t === 'epi');
      rotNome.textContent = t === 'perigo' ? 'Nome (da tabela de perigos)' : 'Nome do EPI';
      const nomes = t === 'perigo' ? nomesTabela() : [...new Set(listaBib('epi').map(p => p.nome))];
      lista.replaceChildren(...nomes.map(n => el('option', {value: n})));
      nome.value = ''; mostrar();
    };
    bts[0].onclick = () => trocarTipo('perigo'); bts[1].onclick = () => trocarTipo('epi');
    nome.oninput = mostrar;
    escolher.onchange = async () => {
      const f = escolher.files[0]; escolher.value = '';
      if (!f) return;
      if (f.size > 2 * 1024 * 1024) { alert('A imagem tem mais de 2 MB. Diminua o tamanho e tente de novo.'); return; }
      try { if (img) URL.revokeObjectURL(img.url); img = await abrirImagem(f); arq = f; mostrar(); }
      catch (e) { alert('Não consegui abrir a imagem: ' + erroAmigavel(e)); }
    };
    const fechar = janela('', 'Enviar pictograma',
      div('env-corpo',
        div('env-campo', el('span', {class: 'env-rot'}, 'Tipo'), div('env-chave', ...bts)),
        div('env-campo', rotNome, nome, lista),
        div('env-campo', el('span', {class: 'env-rot'}, 'Imagem'), previa, el('button', {type: 'button', class: 'btn', onclick: () => escolher.click()}, 'Escolher imagem'), escolher),
        campo('Observação (opcional)', obs, 'envObs1'),
        campo('Seu nome', autor, 'envAutor1')),
      f => [
        el('button', {class: 'btn primario', onclick: async ev => {
          const n = nome.value.trim(), a = autor.value.trim();
          if (!n) { alert('Escolha o nome do pictograma.'); return; }
          if (tipo === 'perigo' && !nomesTabela().includes(n)) { alert(`"${n}" não está na tabela de perigos. Escolha um nome da lista (se o perigo for novo, envie primeiro em Menu > ENVIAR PERIGO).`); return; }
          if (!arq) { alert('Escolha a imagem do pictograma.'); return; }
          if (!a) { alert('Escreva o seu nome.'); autor.focus(); return; }
          const b = ev.currentTarget; b.disabled = true; b.textContent = 'Enviando…';
          try {
            await chamar('enviar', {autor: a, tipo, nome: n, obs: obs.value, nomeArquivo: arq.name, arquivo: await lerBase64(arq), mini: miniatura(img.im)});
            guardar('envio:autor', a); f(); avisar('Enviado, aguardando aprovação.', 'Enviar pictograma');
          } catch (e) { b.disabled = false; b.textContent = 'Enviar'; alert('Não consegui enviar: ' + erroAmigavel(e)); }
        }}, 'Enviar'),
        el('button', {class: 'btn', onclick: f}, 'Cancelar')]);
    trocarTipo('perigo');
    nome.focus();
    return fechar;
  }

  // ---------- ENVIAR PERIGO ----------
  function parecidos(texto) {
    const q = N(texto).split(' ').filter(w => w.length > 2);
    if (!q.length) return [];
    return nomesTabela().map(p => { const w = N(p).split(' '); return [p, q.filter(x => w.some(y => y.startsWith(x) || x.startsWith(y))).length]; })
      .filter(([, n]) => n >= Math.max(1, Math.ceil(q.length / 2))).sort((a, b) => b[1] - a[1]).slice(0, 3).map(([p]) => p);
  }
  function enviarPerigo() {
    const perigo = el('input', {type: 'text', placeholder: 'Ex.: Superfície quente'});
    const risco = el('textarea', {placeholder: 'Ex.: Queimaduras por contato com partes quentes.'});
    const resumo = div('env-resumo');
    const obs = el('textarea', {}), autor = el('input', {type: 'text', value: guardado('envio:autor'), placeholder: 'Nome e sobrenome'});
    const igual = () => nomesTabela().find(p => N(p) === N(perigo.value));
    const mostrar = () => {
      const t = perigo.value.trim();
      if (!t) { resumo.replaceChildren(el('span', {}, 'Se já existir um perigo com nome parecido, o gerador avisa aqui antes de enviar.')); return; }
      const ja = igual(), par = parecidos(t).filter(p => p !== ja);
      resumo.replaceChildren(ja ? selo('Já existe na tabela', false) : selo('Não existe na tabela', true),
        ...(ja ? [el('span', {}, ja)] : par.length ? [el('span', {}, 'Parecidos na tabela:'), ...par.map(p => el('span', {}, '• ' + p))] : [el('span', {}, 'Nenhum perigo parecido na tabela.')]));
    };
    perigo.oninput = mostrar;
    janela('', 'Enviar perigo',
      div('env-corpo', campo('Perigo', perigo, 'envPerigo'), campo('Risco', risco, 'envRisco'), resumo,
        campo('Observação (opcional)', obs, 'envObs2'), campo('Seu nome', autor, 'envAutor2')),
      f => [
        el('button', {class: 'btn primario', onclick: async ev => {
          const p = perigo.value.trim(), r = risco.value.trim(), a = autor.value.trim();
          if (!p || !r) { alert('Preencha o PERIGO e o RISCO.'); return; }
          if (igual()) { alert(`"${igual()}" já está na tabela de perigos.`); return; }
          if (!a) { alert('Escreva o seu nome.'); autor.focus(); return; }
          const b = ev.currentTarget; b.disabled = true; b.textContent = 'Enviando…';
          try {
            await chamar('enviar', {autor: a, tipo: 'perigo-novo', perigo: p, risco: r, obs: obs.value});
            guardar('envio:autor', a); f(); avisar('Enviado, aguardando aprovação.', 'Enviar perigo');
          } catch (e) { b.disabled = false; b.textContent = 'Enviar'; alert('Não consegui enviar: ' + erroAmigavel(e)); }
        }}, 'Enviar'),
        el('button', {class: 'btn', onclick: f}, 'Cancelar')]);
    mostrar(); perigo.focus();
  }

  // ---------- APROVAR ENVIOS ----------
  function pedirSenha() {
    return new Promise(ok => {
      const s = el('input', {type: 'password', autocomplete: 'off', 'aria-label': 'Senha de aprovador'});
      const tentar = async b => {
        if (!s.value) return;
        b.disabled = true;
        try { await chamar('entrar', {senha: s.value}); guardar('envio:senhaAdm', s.value); fechar(); ok(s.value); }
        catch (e) { b.disabled = false; alert(e.status === 401 ? 'Senha incorreta.' : 'Não consegui entrar: ' + erroAmigavel(e)); }
      };
      const fechar = janela('', 'Senha de aprovador', div('env-corpo', s), f => [
        el('button', {class: 'btn primario', onclick: ev => tentar(ev.currentTarget)}, 'Entrar'),
        el('button', {class: 'btn', onclick: () => { f(); ok(null); }}, 'Cancelar')]);
      s.onkeydown = ev => { if (ev.key === 'Enter') { ev.preventDefault(); tentar($('#modal .btn.primario')); } };
      s.focus();
    });
  }
  // depois de aprovar: vale já nesta página (as outras pessoas veem quando o site publicar de novo)
  function aplicarAqui(r) {
    if (r.tipo === 'perigo-novo' && r.linha) {
      (window.TABELA_PERIGOS = window.TABELA_PERIGOS || []).push(r.linha);
      if (typeof APR !== 'undefined' && APR.montarTabela) APR.montarTabela();
    } else if (typeof BIB !== 'undefined') {
      const l = BIB[r.tipo === 'perigo' ? 'perigos' : 'epis1'], i = l.findIndex(p => p.nome === r.nome);
      if (i >= 0) l[i] = {nome: r.nome, img: r.mini}; else l.push({nome: r.nome, img: r.mini});
    }
    try { renderEditor(); renderPrevia(); } catch (e) { console.warn(e); }
  }
  async function aprovarEnvios() {
    let senha = guardado('envio:senhaAdm') || await pedirSenha();
    if (!senha) return;
    const corpo = div('env-corpo', div('env-msg', 'Buscando envios…'));
    janela('aprovar', 'Aprovar envios', corpo, f => [el('button', {class: 'btn', onclick: f}, 'Fechar')]);
    let envios;
    try { envios = (await chamar('pendentes', {senha})).envios; }
    catch (e) {
      if (e.status === 401) { guardar('envio:senhaAdm', ''); $('#modal').hidden = true; aprovarEnvios(); return; }
      corpo.replaceChildren(div('env-msg', 'Não consegui buscar os envios: ' + erroAmigavel(e))); return;
    }
    const vazio = () => { if (!corpo.querySelector('.env-item')) corpo.replaceChildren(div('env-msg', 'Nenhum envio aguardando aprovação.')); };
    const atualizarTotal = () => mostrarTotal(corpo.querySelectorAll('.env-item').length);
    corpo.replaceChildren(...envios.map(e => {
      const novo = e.tipo === 'perigo-novo';
      // pictograma (aprovado 07/11/2026): o aprovador pode trocar TIPO e NOME e escolher SUBSTITUIR ou NOVO
      const pic = novo ? null : {tipo: e.tipo, modo: 'substituir', area: div('env-dados'),
        nome: el('input', {type: 'text', value: e.nome, list: 'envLista' + e.id, autocomplete: 'off', 'aria-label': 'Nome'}),
        lista: el('datalist', {id: 'envLista' + e.id})};
      if (pic) {
        const chaveTipo = div('env-chave', ...['perigo', 'epi'].map(t => el('button', {type: 'button', 'data-t': t}, t === 'perigo' ? 'Perigo' : 'EPI')));
        const chaveModo = div('env-chave', ...['substituir', 'novo'].map(t => el('button', {type: 'button', 'data-m': t}, t === 'novo' ? 'Novo' : 'Substituir')));
        const mostrar = () => {
          chaveTipo.querySelectorAll('button').forEach(b => b.classList.toggle('sel', b.dataset.t === pic.tipo));
          chaveModo.querySelectorAll('button').forEach(b => b.classList.toggle('sel', b.dataset.m === pic.modo));
          const n = pic.nome.value.trim(), atual = listaBib(pic.tipo).find(p => p.nome === n);
          pic.area.replaceChildren(
            atual ? div('env-compara', 'Atual', el('img', {src: atual.img, alt: ''}), '→ novo', el('img', {src: e.mini, alt: ''})) : '',
            atual ? div('env-campo', el('span', {class: 'env-rot'}, 'Já existe com este nome'), chaveModo) : el('span', {}, selo('Novo: não existe com este nome', true)),
            atual ? el('span', {class: 'env-quem'}, `SUBSTITUIR: troca o atual. NOVO: entra ao lado do atual como "${n} (2)", e no gerador dá para escolher qual usar.`) : '');
        };
        const encherLista = () => pic.lista.replaceChildren(...(pic.tipo === 'perigo' ? nomesTabela() : [...new Set(listaBib('epi').map(p => p.nome))]).map(n => el('option', {value: n})));
        chaveTipo.onclick = ev => { const t = ev.target.dataset.t; if (t && t !== pic.tipo) { pic.tipo = t; encherLista(); mostrar(); } };
        chaveModo.onclick = ev => { const m = ev.target.dataset.m; if (m) { pic.modo = m; mostrar(); } };
        pic.nome.oninput = mostrar;
        pic.campos = [div('env-campo', el('span', {class: 'env-rot'}, 'Tipo'), chaveTipo),
          div('env-campo', el('span', {class: 'env-rot'}, 'Nome (você pode mudar)'), pic.nome, pic.lista), pic.area];
        encherLista(); mostrar();
      }
      const campos = novo ? {
        perigo: el('input', {type: 'text', value: e.perigo, 'aria-label': 'Perigo'}),
        risco: el('textarea', {'aria-label': 'Risco'}, e.risco),
        pd: el('select', {'aria-label': 'PD'}, el('option', {value: ''}, 'Escolha o PD'), ...Object.entries(PD).map(([v, t]) => el('option', {value: v}, t))),
      } : null;
      if (campos) campos.risco.value = e.risco;
      const item = div('env-item',
        novo ? div('env-icone', 'Perigo') : el('img', {src: e.mini, alt: ''}),
        div('env-dados',
          el('span', {class: 'env-nome', title: e.nome}, novo ? 'Perigo novo' : e.nome),
          el('span', {class: 'env-quem'}, [novo ? '' : e.tipo === 'perigo' ? 'Perigo' : 'EPI', e.autor, dataBR(e.data)].filter(Boolean).join(' · ')),
          e.obs ? el('span', {class: 'env-quem'}, 'Obs.: ' + e.obs) : '',
          ...(novo ? [div('env-campo', el('span', {class: 'env-rot'}, 'Perigo'), campos.perigo), div('env-campo', el('span', {class: 'env-rot'}, 'Risco'), campos.risco),
                      div('env-campo', el('span', {class: 'env-rot'}, 'PD (você define)'), campos.pd)]
            : pic.campos),
          div('env-linha',
            el('button', {class: 'btn primario', onclick: async ev => {
              const d = {senha, id: e.id};
              if (novo) {
                Object.assign(d, {perigo: campos.perigo.value.trim(), risco: campos.risco.value.trim(), pd: Number(campos.pd.value)});
                if (!d.perigo || !d.risco) { alert('Preencha o PERIGO e o RISCO.'); return; }
                if (!d.pd) { alert('Escolha o PD antes de aprovar.'); campos.pd.focus(); return; }
              } else {
                Object.assign(d, {tipo: pic.tipo, nome: pic.nome.value.trim(), modo: pic.modo});
                if (!d.nome) { alert('Escreva o nome do pictograma.'); pic.nome.focus(); return; }
                if (d.tipo === 'perigo' && !nomesTabela().includes(d.nome)) { alert(`"${d.nome}" não está na tabela de perigos. Escolha um nome da lista.`); pic.nome.focus(); return; }
              }
              const b = ev.currentTarget; b.disabled = true; b.textContent = 'Aprovando…';
              try {
                const r = await chamar('aprovar', d);
                item.remove(); atualizarTotal(); vazio(); aplicarAqui(r);
                aviso(`Aprovado: ${novo ? d.perigo : r.nome}. Em cerca de 1 minuto todos veem no site.`);
              } catch (x) { b.disabled = false; b.textContent = 'Aprovar'; alert('Não consegui aprovar: ' + erroAmigavel(x)); if (x.status === 404) { item.remove(); vazio(); } }
            }}, 'Aprovar'),
            el('button', {class: 'btn', onclick: async ev => {
              const b = ev.currentTarget;   // depois do await o evento já não tem o botão
              if (!(await perguntar(`Recusar "${novo ? e.perigo : e.nome}" enviado por ${e.autor}? O envio é apagado e nada muda no gerador.`, {sim: 'Recusar', nao: 'Voltar', titulo: 'Recusar envio'}))) return;
              b.disabled = true;
              try { await chamar('recusar', {senha, id: e.id}); item.remove(); atualizarTotal(); vazio(); }
              catch (x) { b.disabled = false; alert('Não consegui recusar: ' + erroAmigavel(x)); if (x.status === 404) { item.remove(); vazio(); } }
            }}, 'Recusar'))));
      return item;
    }));
    atualizarTotal(); vazio();
  }

  // número de pendentes no menu, só para quem já entrou como aprovador neste computador
  async function contar() {
    const senha = guardado('envio:senhaAdm');
    if (!senha) return;
    try { mostrarTotal((await chamar('contar', {senha})).total); } catch (e) { if (e.status === 401) guardar('envio:senhaAdm', ''); }
  }
  if (location.protocol.startsWith('http')) setTimeout(contar, 3000);

  const menu = () => [
    ['ENVIAR PICTOGRAMA', enviarPictograma, 'Mandar um pictograma de perigo ou de EPI para todos (fica aguardando aprovação)'],
    ['ENVIAR PERIGO', enviarPerigo, 'Mandar um perigo novo para a tabela de perigos (fica aguardando aprovação)'],
    [rotuloAprovar(), aprovarEnvios, 'Aprovar ou recusar os envios (só com a senha de aprovador)'],
  ];
  return {menu, enviarPictograma, enviarPerigo, aprovarEnvios, API};
})();
