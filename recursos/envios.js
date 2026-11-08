// Envio compartilhado (07/11/2026): qualquer pessoa manda um pictograma ou um perigo novo pelo gerador online,
// sem senha (Menu > INCLUIR PERIGO / PICTOGRAMA); o envio fica PENDENTE e só o aprovador (senha) aprova ou recusa
// em Menu > ADMINISTRADOR, aba APROVAR ENVIOS (08/11/2026: os dois itens de envio e os dois de aprovação viraram um só).
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
  const total = () => pendentesTotal ? ` (${pendentesTotal})` : '';
  const rotuloAdm = () => 'ADMINISTRADOR' + total();
  // o menu já desenhado mostra o número novo sem redesenhar a tela
  function mostrarTotal(n) {
    pendentesTotal = n;
    for (const b of document.querySelectorAll('.lista-menu button')) if (b.textContent.startsWith('ADMINISTRADOR')) b.textContent = rotuloAdm();
    for (const b of document.querySelectorAll('.adm-abas button[data-aba=envios]')) b.textContent = 'Aprovar envios' + total();
  }

  const estilo = document.createElement('style');
  estilo.textContent = `
#editor .lista-menu{width:280px}
#editor .lista-menu .menu-adm{display:flex;align-items:center;gap:6px}
.envio .env-abas{display:grid;grid-template-columns:1fr 1fr;border:1px solid var(--env-cor);border-radius:5px;overflow:hidden;flex:none}
.envio .env-abas button{border:0;background:#fff;color:var(--env-cor);font:700 12px Arial,Helvetica,sans-serif;text-transform:uppercase;height:32px;cursor:pointer}
.envio .env-abas button.sel{background:var(--env-cor);color:#fff}
#editor .lista-menu .menu-adm>button:first-child{flex:1;min-width:0}
#editor .lista-menu button.menu-sair{flex:none;width:auto;height:28px;display:inline-flex;align-items:center;gap:5px;padding:0 10px;margin-right:4px;background:#e5484d;color:#fff;font-weight:700;border-radius:5px}
#editor .lista-menu button.menu-sair:hover{background:#c9363b;color:#fff}
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
.env-lupa{position:fixed;z-index:95;background:#fff;border:1px solid var(--ui-borda);border-radius:8px;box-shadow:0 8px 24px rgba(0,0,0,.25);padding:8px;display:flex;flex-direction:column;align-items:center;gap:6px;pointer-events:none}
.env-lupa img{width:76px;height:76px;object-fit:contain}
.env-lupa span{font:700 11px Arial,Helvetica,sans-serif;text-transform:uppercase;letter-spacing:.04em;color:#374151}
.envio .env-msg{text-align:center;color:var(--ui-suave);font-size:12.5px;padding:18px 0}
`;
  document.head.append(estilo);

  const cor = () => (typeof abaAtual !== 'undefined' && abaAtual === 'apr') ? '#5D6532' : '#002060';
  const div = (cls, ...f) => el('div', {class: cls}, ...f);
  const campo = (rotulo, controle, id) => { if (id) controle.id = id; return div('env-campo', el('label', {class: 'env-rot', for: id}, rotulo), controle); };
  const selo = (txt, verde) => el('span', {class: 'env-selo ' + (verde ? 'verde' : 'cinza')}, txt);
  const dataBR = iso => { const d = new Date(iso); return isNaN(d) ? '' : d.toLocaleDateString('pt-BR') + ' ' + d.toLocaleTimeString('pt-BR', {hour: '2-digit', minute: '2-digit'}); };

  // miniatura pequena: com o mouse em cima abre uma visão maior ao lado (08/11/2026), perigo ou EPI
  function lupa(img, legenda) {
    let caixa = null;
    const tirar = () => { if (caixa) { caixa.remove(); caixa = null; } };
    img.style.cursor = 'zoom-in';
    img.addEventListener('mouseenter', () => {
      tirar();
      const r = img.getBoundingClientRect(), lado = 76;
      caixa = el('div', {class: 'env-lupa'}, el('img', {src: img.src, alt: ''}), legenda ? el('span', {}, legenda) : '');
      const x = Math.min(window.innerWidth - lado - 32, Math.max(8, r.left + r.width / 2 - lado / 2 - 10));
      const y = r.top - lado - 38 > 8 ? r.top - lado - 38 : r.bottom + 8;
      caixa.style.left = x + 'px'; caixa.style.top = y + 'px';
      document.body.append(caixa);
    });
    img.addEventListener('mouseleave', tirar);
    img.addEventListener('click', tirar);
    return img;
  }

  // abre a janela no #modal do gerador; devolve a função que fecha
  function janela(classe, titulo, corpo, botoes) {
    const m = $('#modal');
    const fechar = () => { m.hidden = true; m.replaceChildren(); m.onclick = null; };
    const j = el('div', {class: 'janela envio ' + classe, onclick: ev => ev.stopPropagation()},
      el('header', {}, el('b', {}, titulo)), corpo, el('div', {class: 'janela-bts'}, ...botoes(fechar)));
    j.style.setProperty('--env-cor', cor());
    // 07/11/2026: clicar fora não fecha (para não perder o que foi preenchido); fecha pelo botão ou pelo Esc
    m.replaceChildren(j); m.onclick = null; m.hidden = false;
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

  // ---------- INCLUIR PERIGO / PICTOGRAMA ----------
  // cada aba devolve {corpo, enviar(botão), foco}; a janela troca a aba sem perder o que foi preenchido
  function formPictograma() {
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
    const corpo = div('env-corpo',
        div('env-campo', el('span', {class: 'env-rot'}, 'Tipo'), div('env-chave', ...bts)),
        div('env-campo', rotNome, nome, lista),
        div('env-campo', el('span', {class: 'env-rot'}, 'Imagem'), previa, el('button', {type: 'button', class: 'btn', onclick: () => escolher.click()}, 'Escolher imagem'), escolher),
        campo('Observação (opcional)', obs, 'envObs1'),
        campo('Seu nome', autor, 'envAutor1'));
    const enviar = async () => {
      const n = nome.value.trim(), a = autor.value.trim();
      if (!n) { alert('Escolha o nome do pictograma.'); return false; }
      if (tipo === 'perigo' && !nomesTabela().includes(n)) { alert(`"${n}" não está na tabela de perigos. Escolha um nome da lista (se o perigo for novo, inclua primeiro o perigo, na aba PERIGO).`); return false; }
      if (!arq) { alert('Escolha a imagem do pictograma.'); return false; }
      if (!a) { alert('Escreva o seu nome.'); autor.focus(); return false; }
      await chamar('enviar', {autor: a, tipo, nome: n, obs: obs.value, nomeArquivo: arq.name, arquivo: await lerBase64(arq), mini: miniatura(img.im)});
      guardar('envio:autor', a); return true;
    };
    trocarTipo('perigo');
    return {corpo, enviar, foco: nome};
  }

  function parecidos(texto) {
    const q = N(texto).split(' ').filter(w => w.length > 2);
    if (!q.length) return [];
    return nomesTabela().map(p => { const w = N(p).split(' '); return [p, q.filter(x => w.some(y => y.startsWith(x) || x.startsWith(y))).length]; })
      .filter(([, n]) => n >= Math.max(1, Math.ceil(q.length / 2))).sort((a, b) => b[1] - a[1]).slice(0, 3).map(([p]) => p);
  }
  function formPerigo() {
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
    const corpo = div('env-corpo', campo('Perigo', perigo, 'envPerigo'), campo('Risco', risco, 'envRisco'), resumo,
      campo('Observação (opcional)', obs, 'envObs2'), campo('Seu nome', autor, 'envAutor2'));
    const enviar = async () => {
      const p = perigo.value.trim(), r = risco.value.trim(), a = autor.value.trim();
      if (!p || !r) { alert('Preencha o PERIGO e o RISCO.'); return false; }
      if (igual()) { alert(`"${igual()}" já está na tabela de perigos.`); return false; }
      if (!a) { alert('Escreva o seu nome.'); autor.focus(); return false; }
      await chamar('enviar', {autor: a, tipo: 'perigo-novo', perigo: p, risco: r, obs: obs.value});
      guardar('envio:autor', a); return true;
    };
    mostrar();
    return {corpo, enviar, foco: perigo};
  }
  // Menu > INCLUIR PERIGO / PICTOGRAMA (08/11/2026): uma janela com as abas PERIGO e PICTOGRAMA
  function incluir(aba = 'perigo') {
    const forms = {perigo: formPerigo(), pictograma: formPictograma()};
    const abas = div('env-abas', ...[['perigo', 'Perigo'], ['pictograma', 'Pictograma']].map(([a, t]) => el('button', {type: 'button', 'data-aba': a}, t)));
    const lugar = div('env-lugar');
    lugar.style.cssText = 'display:flex;flex-direction:column;min-height:0;overflow:auto';
    const mostrarAba = a => {
      aba = a;
      abas.querySelectorAll('button').forEach(b => b.classList.toggle('sel', b.dataset.aba === a));
      lugar.replaceChildren(forms[a].corpo); forms[a].foco.focus();
    };
    abas.onclick = ev => { const a = ev.target.closest('button')?.dataset.aba; if (a && a !== aba) mostrarAba(a); };
    const topo = div('', abas); topo.style.cssText = 'padding:12px 12px 0';
    janela('', 'Incluir perigo / pictograma', el('div', {style: 'display:flex;flex-direction:column;min-height:0;flex:1'}, topo, lugar), f => [
      el('button', {class: 'btn primario', onclick: async ev => {
        const b = ev.currentTarget; b.disabled = true; b.textContent = 'Enviando…';
        try {
          if (await forms[aba].enviar()) { f(); avisar('Enviado, aguardando aprovação.', 'Incluir ' + (aba === 'perigo' ? 'perigo' : 'pictograma')); return; }
        } catch (e) { alert('Não consegui enviar: ' + erroAmigavel(e)); }
        b.disabled = false; b.textContent = 'Enviar';
      }}, 'Enviar'),
      el('button', {class: 'btn', onclick: f}, 'Cancelar')]);
    mostrarAba(aba);
  }

  // ---------- APROVAR ENVIOS ----------
  function pedirSenha() {
    return new Promise(ok => {
      const s = el('input', {type: 'password', autocomplete: 'off', 'aria-label': 'Senha de aprovador'});
      const tentar = async b => {
        if (!s.value) return;
        b.disabled = true;
        try { await chamar('entrar', {senha: s.value}); guardar('envio:senhaAdm', s.value); fechar(); redesenhar(); ok(s.value); }
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
  // aba APROVAR ENVIOS do ADMINISTRADOR (08/11/2026): enche o "corpo" com os envios pendentes.
  // antes(): false impede aprovar (o ADMINISTRADOR tem alterações não salvas); depois(): o ADMINISTRADOR relê as listas
  async function montarAprovacao(corpo, senha, {antes = () => true, depois = () => {}} = {}) {
    corpo.replaceChildren(div('env-msg', 'Buscando envios…'));
    let envios;
    try { envios = (await chamar('pendentes', {senha})).envios; }
    catch (e) { corpo.replaceChildren(div('env-msg', 'Não consegui buscar os envios: ' + erroAmigavel(e))); if (e.status === 401) guardar('envio:senhaAdm', ''); return; }
    const vazio = () => { if (!corpo.querySelector('.env-item')) corpo.replaceChildren(div('env-msg', 'Nenhum envio aguardando aprovação.')); };
    const atualizarTotal = () => mostrarTotal(corpo.querySelectorAll('.env-item').length);
    corpo.replaceChildren(...envios.map(e => {
      const novo = e.tipo === 'perigo-novo';
      // pictograma (aprovado 07/11/2026): o aprovador pode trocar TIPO e NOME e escolher SUBSTITUIR ou NOVO
      // 08/11/2026: cada item tem uma linha de campos só (cabeçalho numa linha), para APROVAR / RECUSAR ficarem
      // centralizados no item e alinhados com os campos
      const pic = novo ? null : {tipo: e.tipo, modo: 'substituir', area: div('env-campo env-modo'),
        nome: el('input', {type: 'text', value: e.nome, list: 'envLista' + e.id, autocomplete: 'off', 'aria-label': 'Nome'}),
        lista: el('datalist', {id: 'envLista' + e.id}), estado: el('span', {class: 'env-estado'})};
      if (pic) {
        const chaveTipo = div('env-chave', ...['perigo', 'epi'].map(t => el('button', {type: 'button', 'data-t': t}, t === 'perigo' ? 'Perigo' : 'EPI')));
        const chaveModo = div('env-chave', ...['substituir', 'novo'].map(t => el('button', {type: 'button', 'data-m': t}, t === 'novo' ? 'Novo' : 'Substituir')));
        const mostrar = () => {
          chaveTipo.querySelectorAll('button').forEach(b => b.classList.toggle('sel', b.dataset.t === pic.tipo));
          chaveModo.querySelectorAll('button').forEach(b => b.classList.toggle('sel', b.dataset.m === pic.modo));
          const n = pic.nome.value.trim(), atual = listaBib(pic.tipo).find(p => p.nome === n);
          // 08/11/2026: o selo NOVO fica na linha do rótulo NOME, para não aumentar a altura do item
          pic.estado.replaceChildren(atual ? '' : selo('Novo', true));
          pic.area.title = atual ? `SUBSTITUIR: troca o atual. NOVO: entra ao lado do atual como "${n} (2)", e no gerador dá para escolher qual usar.` : '';
          pic.area.replaceChildren(...(atual ? [el('span', {class: 'env-rot'}, 'Já existe', lupa(el('img', {class: 'env-atual', src: atual.img, alt: 'atual'}), 'Atual')), chaveModo] : []));
        };
        const encherLista = () => pic.lista.replaceChildren(...(pic.tipo === 'perigo' ? nomesTabela() : [...new Set(listaBib('epi').map(p => p.nome))]).map(n => el('option', {value: n})));
        chaveTipo.onclick = ev => { const t = ev.target.dataset.t; if (t && t !== pic.tipo) { pic.tipo = t; encherLista(); mostrar(); } };
        chaveModo.onclick = ev => { const m = ev.target.dataset.m; if (m) { pic.modo = m; mostrar(); } };
        pic.nome.oninput = mostrar;
        pic.campos = [div('env-campo', el('span', {class: 'env-rot'}, 'Tipo'), chaveTipo),
          div('env-campo', el('span', {class: 'env-rot'}, 'Nome (você pode mudar) ', pic.estado), pic.nome, pic.lista), pic.area];
        encherLista(); mostrar();
      }
      const campos = novo ? {
        perigo: el('input', {type: 'text', value: e.perigo, 'aria-label': 'Perigo'}),
        risco: el('textarea', {'aria-label': 'Risco'}, e.risco),
        pd: el('select', {'aria-label': 'PD'}, el('option', {value: ''}, 'Escolha o PD'), ...Object.entries(PD).map(([v, t]) => el('option', {value: v}, t))),
      } : null;
      if (campos) campos.risco.value = e.risco;
      const acoes = div('env-acoes',
            el('button', {class: 'btn primario', onclick: async ev => {
              if (!antes()) return;
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
                item.remove(); atualizarTotal(); vazio(); aplicarAqui(r); depois();
                aviso(`Aprovado: ${novo ? d.perigo : r.nome}. Em cerca de 1 minuto todos veem no site.`);
              } catch (x) { b.disabled = false; b.textContent = 'Aprovar'; alert('Não consegui aprovar: ' + erroAmigavel(x)); if (x.status === 404) { item.remove(); vazio(); } }
            }}, 'Aprovar'),
            el('button', {class: 'btn', onclick: async ev => {
              const b = ev.currentTarget;   // depois do await o evento já não tem o botão
              if (!(await perguntar(`Recusar "${novo ? e.perigo : e.nome}" enviado por ${e.autor}? O envio é apagado e nada muda no gerador.`, {sim: 'Recusar', nao: 'Voltar', titulo: 'Recusar envio'}))) return;
              b.disabled = true;
              try { await chamar('recusar', {senha, id: e.id}); item.remove(); atualizarTotal(); vazio(); }
              catch (x) { b.disabled = false; alert('Não consegui recusar: ' + erroAmigavel(x)); if (x.status === 404) { item.remove(); vazio(); } }
            }}, 'Recusar'));
      const titulo = novo ? 'Perigo novo' : e.nome;
      const meta = [novo ? '' : e.tipo === 'perigo' ? 'Perigo' : 'EPI', e.autor, dataBR(e.data), e.obs ? 'Obs.: ' + e.obs : ''].filter(Boolean).join(' · ');
      const item = div('env-item',
        novo ? div('env-icone', 'Perigo') : el('img', {src: e.mini, alt: ''}),
        div('env-dados',
          el('div', {class: 'env-cab', title: titulo + ' · ' + meta}, el('b', {}, titulo), el('span', {class: 'env-quem'}, ' · ' + meta)),
          div('env-campos ' + (novo ? 'perigo' : 'picto'), ...(novo ? [div('env-campo', el('span', {class: 'env-rot'}, 'Perigo'), campos.perigo), div('env-campo', el('span', {class: 'env-rot'}, 'Risco'), campos.risco),
                      div('env-campo', el('span', {class: 'env-rot'}, 'PD (você define)'), campos.pd)]
            : pic.campos))), acoes);
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
    ['INCLUIR PERIGO / PICTOGRAMA', () => incluir(), 'Mandar um perigo novo para a tabela ou um pictograma de perigo ou de EPI (fica aguardando aprovação)'],
    // 07/11/2026: a senha fica memorizada neste computador; depois de entrar, o SAIR (vermelho) fica na mesma linha
    // do ADMINISTRADOR e esquece a senha. 08/11/2026: APROVAR ENVIOS virou uma aba do ADMINISTRADOR
    guardado('envio:senhaAdm') ? linhaAdministrador() : [rotuloAdm(), () => ADMIN.abrir(), TXT_ADM],
  ];
  const TXT_ADM = 'Aprovar os envios e incluir, editar e excluir perigos da tabela e pictogramas de perigo e de EPI (só com a senha de aprovador)';
  const linhaAdministrador = () => el('div', {class: 'menu-adm'},
    el('button', {type: 'button', title: TXT_ADM, onclick: () => { fecharMenu(); ADMIN.abrir(); }}, rotuloAdm()),
    botaoSair());
  function botaoSair() {
    const b = el('button', {type: 'button', class: 'menu-sair', title: 'Sair do administrador: esquece a senha de aprovador neste computador', onclick: () => { fecharMenu(); sair(); }});
    b.innerHTML = '<svg viewBox="0 0 24 24" width="15" height="15" aria-hidden="true"><path fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round" d="M10 4H5a1 1 0 0 0-1 1v14a1 1 0 0 0 1 1h5M15 8l4 4-4 4M19 12H9"/></svg><span>Sair</span>';
    return b;
  }
  async function sair() {
    if (!(await perguntar('Sair do administrador neste computador? Para usar o ADMINISTRADOR de novo, a senha será pedida.', {sim: 'Sair', nao: 'Voltar', titulo: 'Sair do administrador'}))) return;
    guardar('envio:senhaAdm', ''); pendentesTotal = 0;
    redesenhar();
    aviso('Você saiu do administrador.');
  }
  // o Menu é montado de novo para mostrar ou esconder o SAIR
  const redesenhar = () => { try { renderEditor(); } catch (e) {} };
  // usado pelo ADMINISTRADOR (recursos/administrador.js)
  const ajuda = {chamar, janela, pedirSenha, guardado, guardar, abrirImagem, miniatura, lerBase64, selo, div, PD, N, montarAprovacao, pendentes: () => pendentesTotal || 0};
  return {menu, incluir, API, ajuda};
})();
