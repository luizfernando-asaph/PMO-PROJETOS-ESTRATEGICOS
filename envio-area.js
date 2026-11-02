/*
 * envio-area.js — cadastro de áreas/e-mails e envio do painel de tarefas por área.
 * Dados: window.BD_SNAPSHOT (dados/tarefas-bd.js) ou uma planilha importada pelo usuário.
 * Persistência: localStorage deste navegador (exportável em .json).
 */
(function () {
  'use strict';

  var K = { areas: 'pmoEnvio.areas', notas: 'pmoEnvio.notas', dados: 'pmoEnvio.dados',
            hist: 'pmoEnvio.historico', msg: 'pmoEnvio.mensagem' };
  var SITE_URL = 'https://luizfernando-asaph.github.io/Painel-gestao-projetos/';
  var SHEETJS_URL = 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js';
  var DOMINIO = 'meirelesefreitas.com.br';

  // Áreas sugeridas no primeiro acesso (sem e-mail — cada usuário cadastra os seus)
  var AREAS_PADRAO = [
    { nome: 'Departamento Pessoal', grupos: ['DP', 'Pessoal'] },
    { nome: 'RH / Gente', grupos: ['RH', 'Gente'] },
    { nome: 'SESMT', grupos: ['SESMT'] },
    { nome: 'Treinamento', grupos: ['Treinamento', 'T&D'] },
    { nome: 'Suprimentos', grupos: ['Suprimentos'] },
    { nome: 'Facilities e Base de Apoio', grupos: ['Facilities', 'Base de Apoio'] },
    { nome: 'Frotas', grupos: ['Frotas'] },
    { nome: 'Tecnologia (TIC)', grupos: ['Tecnologia (TIC)'] },
    { nome: 'Projetos / PMO', grupos: ['Projetos', 'Implantação', 'Planejamento', 'Roadmap', 'Sem grupo'] }
  ];

  // ---------- utilidades ----------
  function $(id) { return document.getElementById(id); }
  function esc(s) {
    return String(s === null || s === undefined ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }
  function ler(k, def) {
    try { var v = localStorage.getItem(k); return v ? JSON.parse(v) : def; } catch (e) { return def; }
  }
  function gravar(k, v) {
    try { localStorage.setItem(k, JSON.stringify(v)); return true; }
    catch (e) { toast('Não foi possível salvar neste navegador (armazenamento cheio ou bloqueado).'); return false; }
  }
  function uid() { return 'a' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6); }
  function hojeIso() {
    var d = new Date();
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  }
  function br(iso) { if (!iso) return '—'; var p = iso.split('-'); return p[2] + '/' + p[1] + '/' + p[0]; }
  function dias(aIso, bIso) { return Math.round((Date.parse(bIso) - Date.parse(aIso)) / 86400000); }
  function toast(t) {
    var el = $('toast'); el.textContent = t; el.style.display = 'block';
    clearTimeout(toast._t); toast._t = setTimeout(function () { el.style.display = 'none'; }, 3200);
  }
  function baixar(nome, conteudo, tipo) {
    var blob = conteudo instanceof Blob ? conteudo : new Blob([conteudo], { type: tipo || 'text/plain;charset=utf-8' });
    var a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = nome;
    document.body.appendChild(a); a.click(); setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 500);
  }
  function separaEmails(txt) {
    return String(txt || '').split(/[;,\s]+/).map(function (s) { return s.trim(); }).filter(Boolean);
  }
  function emailValido(e) { return /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(e); }
  function avisoDominio(lista) {
    var sus = lista.filter(function (e) {
      var d = (e.split('@')[1] || '').toLowerCase();
      return d && d !== DOMINIO && /meireles|freitas/.test(d);
    });
    return sus.length ? 'Confira o domínio: ' + sus.join(', ') + ' — o domínio da empresa é @' + DOMINIO + '.' : '';
  }

  // ---------- estado ----------
  var HOJE = hojeIso();
  var dados = null;       // { meta, tarefas }
  var areas = ler(K.areas, null);
  var notas = ler(K.notas, {});
  var historico = ler(K.hist, []);
  var selId = null;
  var desmarcadas = {};   // chave -> true (não incluir no e-mail), por sessão
  var editId = null;

  function carregaDados() {
    var imp = ler(K.dados, null);
    if (imp && imp.tarefas && imp.tarefas.length) {
      dados = imp; $('btnVoltarSnapshot').hidden = false;
    } else {
      var s = window.BD_SNAPSHOT || { tarefas: [], avisos: [] };
      dados = { meta: { origem: 'publicado', fonte: s.fonte, fonteUrl: s.fonteUrl, geradoEm: s.geradoEm, avisos: s.avisos || [] }, tarefas: s.tarefas || [] };
      $('btnVoltarSnapshot').hidden = true;
    }
    var m = dados.meta;
    var quando = m.origem === 'importado'
      ? 'planilha importada <b>' + esc(m.arquivo) + '</b> em ' + esc(m.importadoEm)
      : 'snapshot publicado de <b>' + esc(m.fonte || 'PMO_STATUS REPORT') + '</b>, gerado em ' + br(m.geradoEm);
    $('fonteInfo').innerHTML = 'Dados: ' + quando + ' · ' + dados.tarefas.length + ' tarefas' +
      (m.fonteUrl ? ' · <a style="color:var(--lime)" target="_blank" rel="noopener" href="' + esc(m.fonteUrl) + '">abrir planilha-fonte</a>' : '');
    $('footFonte').textContent = m.origem === 'importado' ? m.arquivo : (m.fonte || 'PMO_STATUS REPORT');
    var av = (m.avisos || []);
    $('avisosDados').innerHTML = av.length
      ? '<div class="callout">⚠ Ajustes automáticos de data feitos na leitura da planilha: ' + av.map(esc).join(' ') +
        ' As tarefas afetadas mostram o ajuste na observação. Vale corrigir na origem.</div>'
      : '';
  }

  if (!areas) {
    areas = AREAS_PADRAO.map(function (a) { return { id: uid(), nome: a.nome, para: [], cc: [], grupos: a.grupos.slice() }; });
    gravar(K.areas, areas);
  }

  // ---------- regras de tarefa ----------
  function aberta(t) { return t.status !== 'Concluído' && t.status !== 'Cancelado'; }
  // Se a planilha traz fim anterior ao início (erro de digitação), o prazo considerado é o início
  function fimEfetivo(t) { return (t.inicio && t.fim && t.fim < t.inicio) ? t.inicio : t.fim; }
  function atrasada(t) { var f = fimEfetivo(t); return aberta(t) && !!f && f < HOJE; }
  function venceEm(t) { var f = fimEfetivo(t); return f ? dias(HOJE, f) : null; }

  // Nota de data em linguagem curta. No e-mail só vai o que afeta o prazo da área.
  function notaDataCurta(t, paraEmail) {
    if (!t.notaData) return '';
    var n = t.notaData;
    if (paraEmail) {
      if (/fim anterior ao início/.test(n)) return 'Data de fim a confirmar (anterior ao início na planilha)';
      if (/não confirmado/.test(n)) return 'Data a confirmar';
      return '';
    }
    return 'Conferir na planilha: ' + n;
  }
  function obsAuto(t, paraEmail) {
    var p = [];
    if (aberta(t) && fimEfetivo(t)) {
      var d = venceEm(t);
      if (d < 0) p.push('<span class="late">Atrasada há ' + (-d) + ' dia' + (-d === 1 ? '' : 's') + '</span>');
      else if (d === 0) p.push('<span class="soon">Vence hoje</span>');
      else if (d <= 15) p.push('<span class="soon">Vence em ' + d + ' dia' + (d === 1 ? '' : 's') + '</span>');
    }
    if (aberta(t) && t.status === 'A Iniciar' && t.inicio && t.inicio < HOJE && !(t.fim && t.fim < HOJE)) p.push('Início previsto já passou');
    if (t.local && t.local !== t.grupo && t.local !== t.grupoOriginal) p.push('Local/tipo: ' + esc(t.local));
    if (t.qtde) p.push('Qtde: ' + esc(t.qtde));
    if (t.sla) p.push('SLA: ' + esc(/^\d+([.,]\d+)?$/.test(t.sla) ? t.sla + ' dia(s)' : t.sla));
    var nd = notaDataCurta(t, paraEmail); if (nd) p.push(esc(nd));
    return p;
  }
  function obsTexto(t) {
    var auto = obsAuto(t, true).map(function (h) { return h.replace(/<[^>]+>/g, ''); });
    var man = (notas[t.chave] || '').trim();
    if (man) auto.unshift(man);
    return auto.join(' · ');
  }
  function statusExib(t) { return atrasada(t) ? 'Atrasada' : t.status; }
  function badgeCls(t) {
    if (atrasada(t)) return 'st-late';
    if (t.status === 'Concluído') return 'st-ok';
    if (t.status === 'Em Andamento') return 'st-prog';
    return 'st-todo';
  }

  function tarefasDaArea(a) {
    var set = {}; (a.grupos || []).forEach(function (g) { set[g] = true; });
    return dados.tarefas.filter(function (t) { return set[t.grupo]; });
  }
  // Projetos com prazo mais urgente primeiro; dentro do projeto, atrasadas e depois por data de fim
  function ordena(lista) {
    var urg = {};
    lista.forEach(function (t) { var f = t.fim || '9999-12-31'; if (!urg[t.projeto] || f < urg[t.projeto]) urg[t.projeto] = f; });
    return lista.slice().sort(function (x, y) {
      if (x.projeto !== y.projeto) return urg[x.projeto].localeCompare(urg[y.projeto]) || x.projeto.localeCompare(y.projeto, 'pt-BR', { numeric: true });
      var ax = atrasada(x) ? 0 : 1, ay = atrasada(y) ? 0 : 1;
      if (ax !== ay) return ax - ay;
      return (x.fim || '9999').localeCompare(y.fim || '9999') || (x.linha || 0) - (y.linha || 0);
    });
  }

  // ---------- 1. cadastro ----------
  function contagemGrupos() {
    var c = {};
    dados.tarefas.forEach(function (t) {
      c[t.grupo] = c[t.grupo] || { total: 0, abertas: 0 };
      c[t.grupo].total++; if (aberta(t)) c[t.grupo].abertas++;
    });
    return c;
  }

  function renderAreas() {
    var el = $('listaAreas');
    if (!areas.length) { el.innerHTML = '<div class="empty">Nenhuma área cadastrada. Clique em "+ Nova área".</div>'; }
    else {
      el.innerHTML = areas.map(function (a) {
        var ts = tarefasDaArea(a), ab = ts.filter(aberta), at = ab.filter(atrasada);
        var prox = ab.filter(function (t) { var d = venceEm(t); return d !== null && d >= 0 && d <= 15; });
        var mails = (a.para || []).length ? esc(a.para.join(', ')) + (a.cc && a.cc.length ? ' · cc: ' + esc(a.cc.join(', ')) : '') : 'Sem e-mail cadastrado';
        return '<div class="card area' + (a.id === selId ? ' sel' : '') + '" data-id="' + a.id + '" tabindex="0" role="button" aria-pressed="' + (a.id === selId) + '">' +
          '<div class="acts"><button data-ed="' + a.id + '" title="Editar">Editar</button><button class="danger" data-del="' + a.id + '" title="Excluir">✕</button></div>' +
          '<h3>' + esc(a.nome) + '</h3>' +
          '<div class="mails' + ((a.para || []).length ? '' : ' vazio') + '">' + mails + '</div>' +
          '<div class="chips">' + (a.grupos || []).map(function (g) { return '<span class="chip">' + esc(g) + '</span>'; }).join('') + '</div>' +
          '<div class="nums"><div><b>' + ab.length + '</b>pendentes</div><div class="late"><b>' + at.length + '</b>atrasadas</div><div><b>' + prox.length + '</b>vencem em 15 dias</div></div>' +
          '</div>';
      }).join('');
    }
    // grupos com pendências que não pertencem a nenhuma área
    var usados = {}; areas.forEach(function (a) { (a.grupos || []).forEach(function (g) { usados[g] = true; }); });
    var c = contagemGrupos();
    var orf = Object.keys(c).filter(function (g) { return !usados[g] && c[g].abertas > 0; });
    var box = $('semArea');
    if (orf.length) {
      box.hidden = false;
      box.innerHTML = '⚠ Grupos com tarefas pendentes que não estão em nenhuma área: ' +
        orf.map(function (g) { return '<b>' + esc(g) + '</b> (' + c[g].abertas + ')'; }).join(', ') + '. Inclua-os em alguma área para que essas tarefas sejam enviadas.';
    } else box.hidden = true;
  }

  function abreForm(a) {
    editId = a ? a.id : null;
    $('formTitulo').textContent = a ? 'Editar área' : 'Nova área';
    $('fNome').value = a ? a.nome : '';
    $('fPara').value = a ? (a.para || []).join('; ') : '';
    $('fCc').value = a ? (a.cc || []).join('; ') : '';
    $('formErro').textContent = '';
    var c = contagemGrupos();
    var marcados = {}; (a ? a.grupos : []).forEach(function (g) { marcados[g] = true; });
    var nomes = Object.keys(c);
    (a ? a.grupos : []).forEach(function (g) { if (nomes.indexOf(g) < 0) nomes.push(g); });
    nomes.sort(function (x, y) { return x.localeCompare(y, 'pt-BR'); });
    $('fGrupos').innerHTML = nomes.map(function (g) {
      var n = c[g] ? c[g].abertas : 0;
      return '<label><input type="checkbox" value="' + esc(g) + '"' + (marcados[g] ? ' checked' : '') + '> ' + esc(g) + '<small>(' + n + ')</small></label>';
    }).join('');
    validaEmailsForm();
    $('formArea').hidden = false;
    $('fNome').focus();
    $('formArea').scrollIntoView({ behavior: 'smooth', block: 'center' });
  }
  function validaEmailsForm() {
    var para = separaEmails($('fPara').value), cc = separaEmails($('fCc').value);
    var inv = para.concat(cc).filter(function (e) { return !emailValido(e); });
    var dom = avisoDominio(para.concat(cc));
    $('fParaHint').innerHTML = inv.length ? '<span class="err">E-mail inválido: ' + esc(inv.join(', ')) + '</span>'
      : (dom ? '<span class="warn">' + esc(dom) + '</span>' : 'Separe vários e-mails com vírgula ou ponto e vírgula.');
    return inv;
  }
  function salvaForm() {
    var nome = $('fNome').value.trim();
    var para = separaEmails($('fPara').value), cc = separaEmails($('fCc').value);
    var grupos = Array.prototype.map.call(document.querySelectorAll('#fGrupos input:checked'), function (i) { return i.value; });
    var inv = validaEmailsForm();
    var erro = !nome ? 'Informe o nome da área.' : !para.length ? 'Informe pelo menos um e-mail de destino.'
      : inv.length ? 'Corrija os e-mails inválidos.' : !grupos.length ? 'Marque pelo menos um grupo da planilha.' : '';
    var dup = areas.filter(function (a) { return a.id !== editId && a.nome.toLowerCase() === nome.toLowerCase(); });
    if (!erro && dup.length) erro = 'Já existe uma área com esse nome.';
    if (erro) { $('formErro').textContent = erro; return; }
    if (editId) {
      areas = areas.map(function (a) { return a.id === editId ? { id: a.id, nome: nome, para: para, cc: cc, grupos: grupos } : a; });
      selId = editId;
    } else {
      var n = { id: uid(), nome: nome, para: para, cc: cc, grupos: grupos };
      areas.push(n); selId = n.id;
    }
    gravar(K.areas, areas);
    $('formArea').hidden = true;
    toast('Área salva.');
    renderTudo();
  }

  // ---------- 2. painel ----------
  function areaSel() { return areas.filter(function (a) { return a.id === selId; })[0] || null; }

  function filtradas() {
    var a = areaSel(); if (!a) return [];
    var st = $('filtroStatus').value, pj = $('filtroProjeto').value, q = $('filtroBusca').value.trim().toLowerCase();
    return ordena(tarefasDaArea(a).filter(function (t) {
      if (st === 'abertas' && !aberta(t)) return false;
      if (st === 'atrasadas' && !atrasada(t)) return false;
      if (st === 'prox') { var d = venceEm(t); if (!(atrasada(t) || (aberta(t) && d !== null && d <= 15))) return false; }
      if (pj && t.projeto !== pj) return false;
      if (q) {
        var txt = (t.atividade + ' ' + t.acao + ' ' + t.projeto + ' ' + obsTexto(t)).toLowerCase();
        if (txt.indexOf(q) < 0) return false;
      }
      return true;
    }));
  }
  function selecionadas() { return filtradas().filter(function (t) { return !desmarcadas[t.chave]; }); }

  function renderPainel() {
    var a = areaSel();
    if (!a) {
      $('painelTitulo').textContent = '2. Painel da área';
      $('painelSub').textContent = 'Selecione uma área acima.'; $('painelCorpo').hidden = true; return;
    }
    $('painelTitulo').textContent = '2. Painel — ' + a.nome;
    $('painelSub').textContent = 'Grupos: ' + a.grupos.join(', ') + '. Hoje: ' + br(HOJE) + '.';
    $('painelCorpo').hidden = false;

    var projs = {}; tarefasDaArea(a).forEach(function (t) { projs[t.projeto] = true; });
    var sel = $('filtroProjeto'), atual = sel.value;
    sel.innerHTML = '<option value="">Todos os projetos</option>' + Object.keys(projs).sort(function (x, y) { return x.localeCompare(y, 'pt-BR', { numeric: true }); })
      .map(function (p) { return '<option' + (p === atual ? ' selected' : '') + '>' + esc(p) + '</option>'; }).join('');

    var lista = filtradas();
    var tb = $('tbody');
    if (!lista.length) {
      tb.innerHTML = '<tr><td colspan="7" class="empty">Nenhuma tarefa nesse filtro.</td></tr>';
    } else {
      var html = [], proj = null;
      lista.forEach(function (t) {
        if (t.projeto !== proj) { proj = t.projeto; html.push('<tr class="proj"><td colspan="7">' + esc(proj) + '</td></tr>'); }
        var off = !!desmarcadas[t.chave];
        html.push('<tr' + (off ? ' class="off"' : '') + '>' +
          '<td><input type="checkbox" data-k="' + esc(t.chave) + '"' + (off ? '' : ' checked') + ' aria-label="Incluir no e-mail"></td>' +
          '<td>' + esc(t.acao) + '<div class="hint" style="margin:0">' + esc(t.grupo) + (t.linha ? ' · linha ' + t.linha + ' da BD' : '') + '</div></td>' +
          '<td>' + esc(t.atividade) + '</td>' +
          '<td class="nowrap">' + br(t.inicio) + '</td>' +
          '<td class="nowrap">' + br(t.fim) + '</td>' +
          '<td><span class="badge ' + badgeCls(t) + '">' + esc(statusExib(t)) + '</span>' + (atrasada(t) ? '<div class="hint" style="margin:2px 0 0">' + esc(t.status) + ' na planilha</div>' : '') + '</td>' +
          '<td class="obs"><div class="auto">' + (obsAuto(t).join(' · ') || '—') + '</div>' +
          '<input type="text" data-nota="' + esc(t.chave) + '" value="' + esc(notas[t.chave] || '') + '" placeholder="Observação para a área (opcional)"></td>' +
          '</tr>');
      });
      tb.innerHTML = html.join('');
    }
    var nSel = lista.filter(function (t) { return !desmarcadas[t.chave]; }).length;
    var nAt = lista.filter(atrasada).length;
    $('contagem').textContent = lista.length + ' tarefa(s) no filtro · ' + nAt + ' atrasada(s) · ' + nSel + ' marcada(s) para o e-mail';
  }

  // ---------- 3. e-mail ----------
  function assuntoPadrao(a) { return '[PMO] Painel de tarefas — ' + a.nome + ' — ' + br(HOJE); }
  function msgPadrao(a) {
    return 'Olá, equipe ' + a.nome + '.\n\nSegue o painel das tarefas da área nos projetos em implantação, com prazos e status atualizados. ' +
      'Por favor, responda este e-mail com as atualizações ou eventuais impedimentos.';
  }
  var assuntoEditado = {}, msgEditada = ler(K.msg, {});

  function resumo(lista) {
    var at = lista.filter(atrasada).length;
    var and = lista.filter(function (t) { return aberta(t) && t.status === 'Em Andamento'; }).length;
    var ini = lista.filter(function (t) { return aberta(t) && t.status === 'A Iniciar'; }).length;
    var conc = lista.filter(function (t) { return t.status === 'Concluído'; }).length;
    var prox = lista.filter(function (t) { var d = venceEm(t); return aberta(t) && d !== null && d >= 0 && d <= 7; }).length;
    return { total: lista.length, at: at, and: and, ini: ini, conc: conc, prox: prox };
  }

  function emailHtml(a, lista, msg) {
    var r = resumo(lista);
    var cor = function (t) { return atrasada(t) ? '#B91C1C' : t.status === 'Concluído' ? '#15803D' : t.status === 'Em Andamento' ? '#1D4ED8' : '#4B5563'; };
    var bg = function (t) { return atrasada(t) ? '#FDECEC' : t.status === 'Concluído' ? '#EEFBF1' : t.status === 'Em Andamento' ? '#EAF1FF' : '#F3F4F6'; };
    var th = 'style="background:#0F2921;color:#fff;text-align:left;padding:8px 10px;font-size:12px;border:1px solid #0F2921"';
    var td = 'style="padding:7px 10px;border:1px solid #E5E7EB;font-size:13px;vertical-align:top"';
    var rows = [], proj = null;
    lista.forEach(function (t) {
      if (t.projeto !== proj) {
        proj = t.projeto;
        rows.push('<tr><td colspan="7" style="background:#E8F3EC;color:#0F2921;font-weight:700;padding:7px 10px;border:1px solid #D1E7DA;font-size:13px">' + esc(proj) + '</td></tr>');
      }
      rows.push('<tr>' +
        '<td ' + td + '>' + esc(t.projeto) + '</td>' +
        '<td ' + td + '>' + esc(t.acao) + '</td>' +
        '<td ' + td + '><b>' + esc(t.atividade) + '</b></td>' +
        '<td ' + td + ' nowrap>' + br(t.inicio) + '</td>' +
        '<td ' + td + ' nowrap>' + br(t.fim) + '</td>' +
        '<td ' + td + '><span style="display:inline-block;padding:2px 8px;border-radius:10px;font-weight:700;font-size:12px;color:' + cor(t) + ';background:' + bg(t) + '">' + esc(statusExib(t)) + '</span></td>' +
        '<td ' + td + '>' + esc(obsTexto(t) || '—') + '</td>' +
        '</tr>');
    });
    var kpi = function (n, l, c) { return '<td style="padding:8px 14px;border:1px solid #E5E7EB;text-align:center"><div style="font-size:20px;font-weight:800;color:' + c + '">' + n + '</div><div style="font-size:11px;color:#6B7280">' + l + '</div></td>'; };
    return '<div style="font-family:Segoe UI,Arial,sans-serif;color:#1F2937;font-size:14px;line-height:1.5">' +
      '<p style="margin:0 0 14px">' + esc(msg).replace(/\n/g, '<br>') + '</p>' +
      '<table cellspacing="0" cellpadding="0" style="border-collapse:collapse;margin:0 0 16px"><tr>' +
      kpi(r.total, 'tarefas listadas', '#0F2921') + kpi(r.at, 'atrasadas', '#B91C1C') + kpi(r.prox, 'vencem em 7 dias', '#C2660A') +
      kpi(r.and, 'em andamento', '#1D4ED8') + kpi(r.ini, 'a iniciar', '#4B5563') + (r.conc ? kpi(r.conc, 'concluídas', '#15803D') : '') +
      '</tr></table>' +
      '<table cellspacing="0" cellpadding="0" style="border-collapse:collapse;width:100%">' +
      '<thead><tr><th ' + th + '>Projeto</th><th ' + th + '>Ação</th><th ' + th + '>Atividade a ser realizada</th><th ' + th + '>Início</th><th ' + th + '>Fim</th><th ' + th + '>Status</th><th ' + th + '>Observação</th></tr></thead>' +
      '<tbody>' + rows.join('') + '</tbody></table>' +
      '<p style="margin:16px 0 0;font-size:12px;color:#6B7280">Área: ' + esc(a.nome) + ' (grupos ' + esc(a.grupos.join(', ')) + ') · posição de ' + br(HOJE) +
      ' · fonte: ' + esc(dados.meta.origem === 'importado' ? dados.meta.arquivo : dados.meta.fonte) +
      ' · <a href="' + SITE_URL + '" style="color:#15803D">Painel de Gestão de Projetos</a></p></div>';
  }

  function emailTexto(a, lista, msg) {
    var r = resumo(lista);
    var out = [msg, '', 'Resumo: ' + r.total + ' tarefas · ' + r.at + ' atrasadas · ' + r.prox + ' vencem em 7 dias · ' + r.and + ' em andamento · ' + r.ini + ' a iniciar' + (r.conc ? ' · ' + r.conc + ' concluídas' : ''), ''];
    var proj = null;
    lista.forEach(function (t) {
      if (t.projeto !== proj) { proj = t.projeto; out.push(''); out.push('■ ' + proj.toUpperCase()); }
      out.push('• ' + t.atividade + ' [' + statusExib(t) + ']');
      out.push('   Ação: ' + t.acao + ' | Início: ' + br(t.inicio) + ' | Fim: ' + br(t.fim));
      var o = obsTexto(t); if (o) out.push('   Obs.: ' + o);
    });
    out.push(''); out.push('Posição de ' + br(HOJE) + ' · ' + SITE_URL);
    return out.join('\n');
  }

  function renderEnvio() {
    var a = areaSel();
    if (!a) { $('envioSub').textContent = 'Selecione uma área com e-mail cadastrado.'; $('envioCorpo').hidden = true; return; }
    if (!(a.para || []).length) {
      $('envioSub').innerHTML = 'A área <b>' + esc(a.nome) + '</b> ainda não tem e-mail. Clique em "Editar" no cartão da área para cadastrar.';
      $('envioCorpo').hidden = true; return;
    }
    $('envioSub').textContent = 'Revise a prévia e escolha como enviar. Só vão as tarefas marcadas no painel acima.';
    $('envioCorpo').hidden = false;
    $('envPara').innerHTML = esc(a.para.join(', ')) + (a.cc && a.cc.length ? '<br><span style="color:var(--muted)">cc: ' + esc(a.cc.join(', ')) + '</span>' : '');
    if (!assuntoEditado[a.id]) $('envAssunto').value = assuntoPadrao(a); else $('envAssunto').value = assuntoEditado[a.id];
    $('envMsg').value = msgEditada[a.id] || msgPadrao(a);
    atualizaPrevia();
    var h = historico.filter(function (x) { return x.areaId === a.id; }).slice(-3).reverse();
    $('historico').innerHTML = h.length ? 'Últimos envios preparados: ' + h.map(function (x) { return esc(x.quando) + ' (' + x.qtd + ' tarefas, ' + esc(x.metodo) + ')'; }).join(' · ') : '';
  }
  function atualizaPrevia() {
    var a = areaSel(); if (!a || !(a.para || []).length) return;
    var lista = selecionadas();
    $('preview').innerHTML = lista.length ? emailHtml(a, lista, $('envMsg').value) : '<p style="color:#6B7280">Nenhuma tarefa marcada para envio.</p>';
    ['btnGmail', 'btnCopiar', 'btnEml', 'btnMailto'].forEach(function (id) { $(id).disabled = !lista.length; });
  }

  function registra(metodo) {
    var a = areaSel();
    var d = new Date();
    historico.push({ areaId: a.id, area: a.nome, quando: br(HOJE) + ' ' + String(d.getHours()).padStart(2, '0') + ':' + String(d.getMinutes()).padStart(2, '0'), qtd: selecionadas().length, metodo: metodo });
    historico = historico.slice(-60); gravar(K.hist, historico); renderEnvio();
  }

  function b64utf8(s) {
    var bytes = new TextEncoder().encode(s), bin = '';
    for (var i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    return btoa(bin).replace(/.{76}/g, '$&\r\n');
  }
  function geraEml(a, assunto, html, texto) {
    var b = '=_pmo_' + Date.now().toString(36);
    return [
      'To: ' + a.para.join(', '),
      a.cc && a.cc.length ? 'Cc: ' + a.cc.join(', ') : null,
      'Subject: =?UTF-8?B?' + btoa(String.fromCharCode.apply(null, new TextEncoder().encode(assunto))) + '?=',
      'X-Unsent: 1',
      'MIME-Version: 1.0',
      'Content-Type: multipart/alternative; boundary="' + b + '"',
      '',
      '--' + b,
      'Content-Type: text/plain; charset=UTF-8',
      'Content-Transfer-Encoding: base64',
      '',
      b64utf8(texto),
      '--' + b,
      'Content-Type: text/html; charset=UTF-8',
      'Content-Transfer-Encoding: base64',
      '',
      b64utf8('<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body>' + html + '</body></html>'),
      '--' + b + '--',
      ''
    ].filter(function (l) { return l !== null; }).join('\r\n');
  }

  function acaoGmail() {
    var a = areaSel(), lista = selecionadas(), assunto = $('envAssunto').value;
    var corpo = emailTexto(a, lista, $('envMsg').value);
    var url = 'https://mail.google.com/mail/?view=cm&fs=1&to=' + encodeURIComponent(a.para.join(',')) +
      (a.cc && a.cc.length ? '&cc=' + encodeURIComponent(a.cc.join(',')) : '') +
      '&su=' + encodeURIComponent(assunto) + '&body=' + encodeURIComponent(corpo);
    if (url.length > 7500) {
      url = 'https://mail.google.com/mail/?view=cm&fs=1&to=' + encodeURIComponent(a.para.join(',')) +
        (a.cc && a.cc.length ? '&cc=' + encodeURIComponent(a.cc.join(',')) : '') + '&su=' + encodeURIComponent(assunto);
      copiaFormatado(true);
      $('envAviso').innerHTML = '<span class="warn">A lista é grande demais para ir pelo link do Gmail. Abrimos o rascunho com destinatários e assunto e <b>copiamos o e-mail formatado</b>: é só colar (Ctrl+V) no corpo.</span>';
    } else $('envAviso').textContent = '';
    window.open(url, '_blank', 'noopener');
    registra('Gmail');
  }
  function copiaFormatado(silencioso) {
    var a = areaSel(), lista = selecionadas();
    var html = emailHtml(a, lista, $('envMsg').value), texto = emailTexto(a, lista, $('envMsg').value);
    var feito = function () { if (!silencioso) toast('E-mail copiado. Cole no corpo da mensagem (Ctrl+V).'); };
    if (navigator.clipboard && window.ClipboardItem) {
      navigator.clipboard.write([new ClipboardItem({
        'text/html': new Blob([html], { type: 'text/html' }),
        'text/plain': new Blob([texto], { type: 'text/plain' })
      })]).then(feito, function () { copiaFallback(); feito(); });
    } else { copiaFallback(); feito(); }
    if (!silencioso) registra('copiado');
  }
  function copiaFallback() {
    var r = document.createRange(); r.selectNodeContents($('preview'));
    var s = window.getSelection(); s.removeAllRanges(); s.addRange(r);
    try { document.execCommand('copy'); } catch (e) { /* ignora */ }
    s.removeAllRanges();
  }
  function acaoEml() {
    var a = areaSel(), lista = selecionadas(), assunto = $('envAssunto').value;
    var eml = geraEml(a, assunto, emailHtml(a, lista, $('envMsg').value), emailTexto(a, lista, $('envMsg').value));
    var nome = 'tarefas-' + a.nome.normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9]+/gi, '-').toLowerCase() + '-' + HOJE + '.eml';
    baixar(nome, eml, 'message/rfc822');
    $('envAviso').textContent = 'Abra o arquivo baixado: o Outlook mostra a mensagem como rascunho, pronta para "Enviar".';
    registra('Outlook .eml');
  }
  function acaoMailto() {
    var a = areaSel(), lista = selecionadas();
    var corpo = emailTexto(a, lista, $('envMsg').value);
    var url = 'mailto:' + a.para.map(encodeURIComponent).join(',') + '?subject=' + encodeURIComponent($('envAssunto').value) +
      (a.cc && a.cc.length ? '&cc=' + a.cc.map(encodeURIComponent).join(',') : '') + '&body=' + encodeURIComponent(corpo);
    if (url.length > 1900) $('envAviso').innerHTML = '<span class="warn">A lista é longa e alguns programas de e-mail cortam o texto. Se faltar conteúdo, use "Copiar e-mail formatado" ou o arquivo .eml.</span>';
    window.location.href = url;
    registra('app de e-mail');
  }

  function acaoCsv() {
    var a = areaSel(), lista = selecionadas();
    var cols = ['Projeto', 'Ação', 'Atividade a ser realizada', 'Início', 'Fim', 'Status', 'Status na planilha', 'Observação', 'Grupo', 'Linha BD'];
    var q = function (v) { return '"' + String(v === null || v === undefined ? '' : v).replace(/"/g, '""') + '"'; };
    var linhas = [cols.map(q).join(';')].concat(lista.map(function (t) {
      return [t.projeto, t.acao, t.atividade, br(t.inicio), br(t.fim), statusExib(t), t.status, obsTexto(t), t.grupo, t.linha].map(q).join(';');
    }));
    baixar('tarefas-' + a.nome.replace(/[^a-z0-9]+/gi, '-').toLowerCase() + '-' + HOJE + '.csv', '﻿' + linhas.join('\r\n'), 'text/csv;charset=utf-8');
  }

  // ---------- importação / exportação ----------
  function exportaCadastro() {
    baixar('cadastro-areas-pmo-' + HOJE + '.json', JSON.stringify({ tipo: 'pmo-envio-areas', versao: 1, exportadoEm: HOJE, areas: areas, notas: notas }, null, 2), 'application/json');
  }
  function importaCadastro(file) {
    var fr = new FileReader();
    fr.onload = function () {
      try {
        var j = JSON.parse(fr.result);
        if (!j || !Array.isArray(j.areas)) throw new Error('formato');
        var validas = j.areas.filter(function (a) { return a && typeof a.nome === 'string' && Array.isArray(a.grupos); }).map(function (a) {
          return { id: a.id || uid(), nome: String(a.nome).trim(),
                   para: (a.para || []).map(String).filter(emailValido), cc: (a.cc || []).map(String).filter(emailValido),
                   grupos: a.grupos.map(String) };
        });
        if (!validas.length) throw new Error('vazio');
        var porNome = {}; areas.forEach(function (a) { porNome[a.nome.toLowerCase()] = a; });
        var novas = 0, atualizadas = 0;
        validas.forEach(function (v) {
          var ex = porNome[v.nome.toLowerCase()];
          if (ex) { ex.para = v.para; ex.cc = v.cc; ex.grupos = v.grupos; atualizadas++; }
          else { areas.push(v); novas++; }
        });
        if (j.notas && typeof j.notas === 'object') Object.keys(j.notas).forEach(function (k) { if (typeof j.notas[k] === 'string') notas[k] = j.notas[k]; });
        gravar(K.areas, areas); gravar(K.notas, notas);
        toast('Cadastro importado: ' + novas + ' área(s) nova(s), ' + atualizadas + ' atualizada(s).');
        renderTudo();
      } catch (e) { toast('Arquivo inválido: use um .json exportado por esta página.'); }
    };
    fr.readAsText(file);
  }

  function carregaSheetJS() {
    return new Promise(function (ok, falha) {
      if (window.XLSX) return ok(window.XLSX);
      var s = document.createElement('script'); s.src = SHEETJS_URL;
      s.onload = function () { ok(window.XLSX); }; s.onerror = function () { falha(new Error('Não foi possível carregar o leitor de planilhas (sem internet?).')); };
      document.head.appendChild(s);
    });
  }
  function importaXlsx(file) {
    toast('Lendo a planilha…');
    carregaSheetJS().then(function (XLSX) {
      return file.arrayBuffer().then(function (buf) {
        var wb = XLSX.read(buf, { type: 'array', cellDates: false });
        var res = window.BDParser.parseWorkbook(XLSX, wb);
        if (!res.tarefas.length) throw new Error('Nenhuma tarefa encontrada na aba BD.');
        var d = new Date();
        dados = { meta: { origem: 'importado', arquivo: file.name, importadoEm: br(HOJE) + ' ' + String(d.getHours()).padStart(2, '0') + ':' + String(d.getMinutes()).padStart(2, '0'), avisos: res.avisos }, tarefas: res.tarefas };
        if (!gravar(K.dados, dados)) toast('Planilha lida, mas não ficou salva: ao recarregar a página, volta aos dados publicados.');
        else toast(res.tarefas.length + ' tarefas carregadas de ' + file.name + '.');
        carregaDados(); renderTudo();
      });
    }).catch(function (e) { toast(e.message || 'Falha ao ler a planilha.'); });
  }

  // ---------- render geral e eventos ----------
  function renderTudo() { renderAreas(); renderPainel(); renderEnvio(); }

  $('listaAreas').addEventListener('click', function (e) {
    var ed = e.target.getAttribute('data-ed'), del = e.target.getAttribute('data-del');
    if (ed) { e.stopPropagation(); abreForm(areas.filter(function (a) { return a.id === ed; })[0]); return; }
    if (del) {
      e.stopPropagation();
      var a = areas.filter(function (x) { return x.id === del; })[0];
      if (a && confirm('Excluir a área "' + a.nome + '"? As tarefas continuam na planilha; só o cadastro é removido.')) {
        areas = areas.filter(function (x) { return x.id !== del; }); if (selId === del) selId = null;
        gravar(K.areas, areas); renderTudo();
      }
      return;
    }
    var card = e.target.closest('.area'); if (!card) return;
    selId = card.getAttribute('data-id'); desmarcadas = {}; $('filtroProjeto').value = ''; $('filtroBusca').value = '';
    renderTudo();
    $('painel').scrollIntoView({ behavior: 'smooth' });
  });
  $('listaAreas').addEventListener('keydown', function (e) {
    if ((e.key === 'Enter' || e.key === ' ') && e.target.classList.contains('area')) { e.preventDefault(); e.target.click(); }
  });
  $('btnNovaArea').onclick = function () { abreForm(null); };
  $('btnCancelarArea').onclick = function () { $('formArea').hidden = true; };
  $('btnSalvarArea').onclick = salvaForm;
  $('fPara').addEventListener('input', validaEmailsForm);
  $('fCc').addEventListener('input', validaEmailsForm);
  $('btnExportCad').onclick = exportaCadastro;
  $('btnImportCad').onclick = function () { $('fileCad').click(); };
  $('fileCad').onchange = function () { if (this.files[0]) importaCadastro(this.files[0]); this.value = ''; };
  $('btnImportXlsx').onclick = function () { $('fileXlsx').click(); };
  $('fileXlsx').onchange = function () { if (this.files[0]) importaXlsx(this.files[0]); this.value = ''; };
  $('btnVoltarSnapshot').onclick = function () {
    try { localStorage.removeItem(K.dados); } catch (e) { /* ignora */ }
    carregaDados(); renderTudo(); toast('Voltamos aos dados publicados no site.');
  };

  ['filtroStatus', 'filtroProjeto'].forEach(function (id) { $(id).addEventListener('change', function () { renderPainel(); atualizaPrevia(); }); });
  $('filtroBusca').addEventListener('input', function () { renderPainel(); atualizaPrevia(); });
  $('tbody').addEventListener('change', function (e) {
    var k = e.target.getAttribute('data-k');
    if (k) { if (e.target.checked) delete desmarcadas[k]; else desmarcadas[k] = true; e.target.closest('tr').classList.toggle('off', !e.target.checked); atualizaPrevia();
      var lista = filtradas(); $('contagem').textContent = lista.length + ' tarefa(s) no filtro · ' + lista.filter(atrasada).length + ' atrasada(s) · ' + lista.filter(function (t) { return !desmarcadas[t.chave]; }).length + ' marcada(s) para o e-mail'; }
  });
  $('tbody').addEventListener('input', function (e) {
    var k = e.target.getAttribute('data-nota');
    if (k !== null) {
      var v = e.target.value; if (v.trim()) notas[k] = v; else delete notas[k];
      clearTimeout(this._t); this._t = setTimeout(function () { gravar(K.notas, notas); atualizaPrevia(); }, 400);
    }
  });
  $('btnMarcarTodas').onclick = function () { filtradas().forEach(function (t) { delete desmarcadas[t.chave]; }); renderPainel(); atualizaPrevia(); };
  $('btnDesmarcarTodas').onclick = function () { filtradas().forEach(function (t) { desmarcadas[t.chave] = true; }); renderPainel(); atualizaPrevia(); };
  $('btnCsv').onclick = acaoCsv;
  $('envAssunto').addEventListener('input', function () { var a = areaSel(); if (a) assuntoEditado[a.id] = this.value; });
  $('envMsg').addEventListener('input', function () { var a = areaSel(); if (a) { msgEditada[a.id] = this.value; gravar(K.msg, msgEditada); } atualizaPrevia(); });
  $('btnGmail').onclick = acaoGmail;
  $('btnCopiar').onclick = function () { copiaFormatado(false); };
  $('btnEml').onclick = acaoEml;
  $('btnMailto').onclick = acaoMailto;

  carregaDados();
  renderTudo();
})();
