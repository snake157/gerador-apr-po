// Janelas de aviso e de pergunta no estilo do gerador, no lugar das caixas do navegador (alert / confirm).
// - avisar(texto): mostra o aviso com o botão "OK" (devolve uma Promise, que termina ao fechar).
// - perguntar(texto, {sim, nao, titulo}): botões com o que cada um faz; devolve true (sim) ou false (não / Esc).
// - escolher(texto, [botões], {titulo}): vários botões; devolve o índice do botão (ou -1 com Esc).
// Os botões têm a mesma largura e ocupam a linha; o primeiro é o principal (Enter), na cor da aba aberta.
// O alert() da página passa a usar o avisar(). Uma janela de cada vez: as outras esperam na fila.
'use strict';

(() => {
  const estilo = document.createElement('style');
  estilo.textContent = `
.jp-fundo{position:fixed;inset:0;background:rgba(0,0,0,.45);display:flex;align-items:center;justify-content:center;z-index:90}
.jp-fundo .jp{background:#fff;border-radius:10px;width:min(460px,94vw);max-height:86vh;display:flex;flex-direction:column;gap:14px;padding:18px;box-shadow:0 10px 30px rgba(0,0,0,.3);font:14px/1.45 "Segoe UI",Arial,sans-serif;color:#111827}
.jp-fundo .jp-tit{font:700 14px Arial,Helvetica,sans-serif;text-transform:uppercase;letter-spacing:.03em;text-align:center;color:var(--jp-cor)}
.jp-fundo .jp-txt{white-space:pre-line;text-align:center;overflow:auto}
.jp-fundo .jp-bts{display:grid;grid-auto-columns:minmax(0,1fr);grid-auto-flow:column;gap:8px}
.jp-fundo .jp-bts.muitos{grid-auto-flow:row}
.jp-fundo button{font:700 12.5px Arial,Helvetica,sans-serif;text-transform:uppercase;letter-spacing:.03em;text-align:center;padding:8px 6px;border-radius:5px;cursor:pointer;border:1px solid var(--jp-cor);background:#fff;color:var(--jp-cor)}
.jp-fundo button.principal{background:var(--jp-cor);color:#fff}
.jp-fundo button:hover{filter:brightness(1.15)}
.jp-fundo button:focus-visible{outline:2px solid var(--jp-cor);outline-offset:2px}
`;
  document.head.append(estilo);

  let fila = Promise.resolve();
  const corDaAba = () => (typeof abaAtual !== 'undefined' && abaAtual === 'apr') ? '#5D6532' : '#002060';

  function mostrar(texto, botoes, titulo) {
    const vez = fila.then(() => new Promise(ok => {
      const fundo = document.createElement('div');
      fundo.id = 'janelaPergunta'; fundo.className = 'jp-fundo';
      fundo.style.setProperty('--jp-cor', corDaAba());
      const caixa = document.createElement('div'); caixa.className = 'jp';
      if (titulo) { const t = document.createElement('div'); t.className = 'jp-tit'; t.textContent = titulo; caixa.append(t); }
      const tx = document.createElement('div'); tx.className = 'jp-txt'; tx.textContent = String(texto ?? ''); caixa.append(tx);
      const bts = document.createElement('div'); bts.className = 'jp-bts' + (botoes.length > 2 ? ' muitos' : '');
      const anterior = document.activeElement;
      const fechar = i => {
        document.removeEventListener('keydown', teclas, true);
        // some aos poucos (26/10/2026); sem id, a próxima janela da fila já pode abrir
        if (matchMedia('(prefers-reduced-motion: reduce)').matches) fundo.remove();
        else { fundo.removeAttribute('id'); fundo.classList.add('modal-saindo', 'fim'); setTimeout(() => fundo.remove(), 200); }
        if (anterior && anterior.focus && document.contains(anterior)) { try { anterior.focus({preventScroll: true}); } catch (e) {} }
        ok(i);
      };
      botoes.forEach((rot, i) => {
        const b = document.createElement('button');
        b.type = 'button'; b.textContent = rot; if (i === 0) b.className = 'principal';
        b.addEventListener('click', ev => { ev.stopPropagation(); fechar(i); });
        bts.append(b);
      });
      caixa.append(bts); fundo.append(caixa);
      // tecla: Enter = botão em foco (ou o principal); Esc = cancelar (o último botão; aviso de um botão só fecha)
      const teclas = ev => {
        if (ev.key === 'Escape') { ev.preventDefault(); ev.stopPropagation(); fechar(botoes.length > 1 ? -1 : 0); }
        else if (ev.key === 'Enter') { ev.preventDefault(); ev.stopPropagation(); const f = [...bts.children].indexOf(document.activeElement); fechar(f >= 0 ? f : 0); }
        else if (ev.ctrlKey || ev.metaKey) { ev.preventDefault(); ev.stopPropagation(); }   // Ctrl+S, Ctrl+Z... não agem por baixo da janela
      };
      document.addEventListener('keydown', teclas, true);
      document.body.append(fundo);
      bts.firstChild.focus();
    }));
    fila = vez.catch(() => {});
    return vez;
  }

  window.avisar = (texto, titulo = 'AVISO') => mostrar(texto, ['OK'], titulo).then(() => undefined);
  window.perguntar = (texto, o = {}) => mostrar(texto, [o.sim || 'OK', o.nao || 'CANCELAR'], o.titulo || 'CONFIRMAR').then(i => i === 0);
  window.escolher = (texto, botoes, o = {}) => mostrar(texto, botoes, o.titulo || 'ESCOLHA');
  window.alert = texto => { window.avisar(texto); };
})();
