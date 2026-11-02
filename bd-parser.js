/*
 * bd-parser.js — lê a aba "BD" do PMO_STATUS REPORT e devolve a lista de tarefas normalizada.
 * Funciona no navegador (window.BDParser) e no Node (module.exports).
 *
 * Regras principais:
 *  - Colunas usadas (aba BD): C Projeto · D Marco · E Grupo · F Id · G Ação · H Tipo/Local ·
 *    I Qtde · K SLA · L Status · M Início · N Dia · O Fim · P Dia.
 *  - Linhas-resumo de marco (sem Id e com Grupo = "MARCO n") não viram tarefa; o texto da
 *    coluna Ação delas vira o nome do marco (ex.: "MARCO 1 · Atração, Seleção e Admissão").
 *  - Datas: a planilha mistura dia/mês e mês/dia. Para cada data, testamos as leituras
 *    possíveis e escolhemos a que bate com o dia da semana registrado nas colunas "Dia".
 *    Anos 19xx são tratados como erro de digitação de 20xx (ex.: 1926 → 2026).
 */
(function (root) {
  'use strict';

  var COL = { projeto: 2, marco: 3, grupo: 4, id: 5, acao: 6, local: 7, qtde: 8, sla: 10,
              status: 11, inicio: 12, diaIni: 13, fim: 14, diaFim: 15 };

  var WD = { dom: 0, seg: 1, ter: 2, qua: 3, qui: 4, sex: 5, sab: 6 };

  function semAcento(s) {
    return String(s).normalize('NFD').replace(/[̀-ͯ]/g, '');
  }
  function limpa(s) {
    if (s === null || s === undefined) return '';
    return String(s).replace(/\s+/g, ' ').trim();
  }

  // Nome canônico do grupo (corrige caixa e erros de digitação da planilha)
  var GRUPOS = {
    'dp': 'DP', 'rh': 'RH', 'sesmt': 'SESMT', 'suprimentos': 'Suprimentos',
    'base apoio': 'Base de Apoio', 'base de apoio': 'Base de Apoio',
    'projetos': 'Projetos', 'facilities': 'Facilities', 'faiclities': 'Facilities',
    'frotas': 'Frotas', 'pessoal': 'Pessoal', 'qualidade': 'Qualidade', 't&d': 'T&D',
    'treinamento': 'Treinamento', 'implantacao': 'Implantação', 'meireles': 'Meireles',
    'tic': 'Tecnologia (TIC)', 'tecnologia': 'Tecnologia (TIC)', 'tecmologia': 'Tecnologia (TIC)',
    'tecnolologia': 'Tecnologia (TIC)', 'comercial': 'Comercial', 'gente': 'Gente',
    'cliente': 'Cliente', 'roadmap': 'Roadmap', 'inteligencia neg.': 'Inteligência de Negócio',
    'planejamento': 'Planejamento', 'operacao': 'Operação'
  };
  function grupoCanonico(g) {
    var s = limpa(g);
    if (!s) return 'Sem grupo';
    var k = semAcento(s).toLowerCase();
    return GRUPOS[k] || s;
  }

  function diaSemana(v) {
    var k = semAcento(limpa(v)).toLowerCase().replace(/[^a-z]/g, '').slice(0, 3);
    return Object.prototype.hasOwnProperty.call(WD, k) ? WD[k] : null;
  }

  function valida(y, m, d) {
    if (!(m >= 1 && m <= 12 && d >= 1 && d <= 31)) return false;
    var dt = new Date(Date.UTC(y, m - 1, d));
    return dt.getUTCMonth() === m - 1 && dt.getUTCDate() === d;
  }
  function wd(y, m, d) { return new Date(Date.UTC(y, m - 1, d)).getUTCDay(); }

  // Converte número serial do Excel em {y,m,d}
  function serialParaYMD(n) {
    var ms = Math.round((n - 25569) * 86400000);
    var dt = new Date(ms);
    return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate() };
  }

  /*
   * Resolve uma data da BD.
   * Retorna { iso: 'AAAA-MM-DD' | null, nota: string | '' }
   */
  function resolveData(valor, diaTexto, refIso) {
    if (valor === null || valor === undefined || valor === '') return { iso: null, nota: '' };
    var cands = [];
    if (typeof valor === 'number') {
      var p = serialParaYMD(valor);
      cands.push({ y: p.y, m: p.m, d: p.d, padrao: true });
      if (p.d <= 12 && p.d !== p.m) cands.push({ y: p.y, m: p.d, d: p.m });
    } else if (valor instanceof Date) {
      var q = { y: valor.getUTCFullYear(), m: valor.getUTCMonth() + 1, d: valor.getUTCDate() };
      cands.push({ y: q.y, m: q.m, d: q.d, padrao: true });
      if (q.d <= 12 && q.d !== q.m) cands.push({ y: q.y, m: q.d, d: q.m });
    } else {
      var mm = limpa(valor).match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/);
      if (!mm) return { iso: null, nota: 'data inválida na planilha: "' + limpa(valor) + '"' };
      var a = +mm[1], b = +mm[2], y = +mm[3];
      // Strings da planilha vêm em mês/dia/ano (padrão), mas aceitamos dia/mês se o dia da semana confirmar
      if (valida(y, a, b)) cands.push({ y: y, m: a, d: b, padrao: true });
      if (a !== b && valida(y, b, a)) cands.push({ y: y, m: b, d: a, padrao: !cands.length });
    }
    cands = cands.filter(function (c) { return valida(c.y, c.m, c.d); });
    if (!cands.length) return { iso: null, nota: 'data inválida na planilha' };

    var alvo = diaSemana(diaTexto);
    var escolhida = null, confirmada = false;
    if (alvo !== null) {
      var ok = cands.filter(function (c) {
        return wd(c.y, c.m, c.d) === alvo || (c.y < 2000 && wd(c.y + 100, c.m, c.d) === alvo);
      });
      if (ok.length === 1) { escolhida = ok[0]; confirmada = true; }
      else if (ok.length > 1) { escolhida = maisProxima(ok, refIso) || ok.filter(function (c) { return c.padrao; })[0] || ok[0]; confirmada = !!refIso; }
    }
    if (!escolhida) escolhida = cands.filter(function (c) { return c.padrao; })[0] || cands[0];

    var notas = [];
    var yy = escolhida.y;
    if (yy < 2000) { notas.push('ano ' + yy + ' corrigido para ' + (yy + 100)); yy += 100; }
    if (!confirmada && cands.length > 1) notas.push('dia/mês não confirmado');
    var iso = yy + '-' + String(escolhida.m).padStart(2, '0') + '-' + String(escolhida.d).padStart(2, '0');
    return { iso: iso, nota: notas.join('; ') };
  }

  // Desempate: entre leituras igualmente válidas, a mais próxima da data de referência
  function maisProxima(cands, refIso) {
    if (!refIso) return null;
    var ref = Date.parse(refIso + 'T00:00:00Z');
    var best = null, dist = Infinity;
    cands.forEach(function (c) {
      var y = c.y < 2000 ? c.y + 100 : c.y;
      var dd = Math.abs(Date.UTC(y, c.m - 1, c.d) - ref);
      if (dd < dist) { dist = dd; best = c; }
    });
    return best;
  }

  function normalizaStatus(s) {
    var k = semAcento(limpa(s)).toLowerCase();
    if (k.indexOf('conclu') === 0) return 'Concluído';
    if (k.indexOf('andamento') >= 0) return 'Em Andamento';
    if (k.indexOf('iniciar') >= 0) return 'A Iniciar';
    if (k.indexOf('atras') >= 0) return 'Atrasado';
    if (k.indexOf('cancel') >= 0) return 'Cancelado';
    if (k.indexOf('paralis') >= 0 || k.indexOf('bloque') >= 0) return 'Paralisado';
    return limpa(s) || 'Sem status';
  }

  function formataQtde(v) {
    if (v === null || v === undefined || v === '') return '';
    if (typeof v === 'number') return String(Math.round(v * 100) / 100).replace('.', ',');
    return limpa(v);
  }

  /*
   * rows: matriz (array de arrays) da aba BD, começando na linha 1 da planilha (índice 0).
   * Retorna { tarefas: [...], avisos: [...] }
   */
  function parseRows(rows) {
    var marcos = {};   // projeto|marco -> rótulo
    var brutas = [];
    var avisos = [];
    var ultimoProjeto = '';

    for (var i = 0; i < rows.length; i++) {
      var r = rows[i] || [];
      var projeto = limpa(r[COL.projeto]);
      var marco = limpa(r[COL.marco]);
      var grupo = limpa(r[COL.grupo]);
      var acao = limpa(r[COL.acao]);
      if (!projeto) continue;
      if (marco === 'Marco' && grupo === 'Grupo') { ultimoProjeto = projeto; continue; } // cabeçalho de bloco
      var semId = r[COL.id] === null || r[COL.id] === undefined || r[COL.id] === '';
      var ehResumo = semId && (/^marco\s*\d+$/i.test(grupo) || (grupo && grupo === marco));
      if (ehResumo) {
        if (acao) marcos[projeto + '|' + marco.toUpperCase()] = acao;
        continue;
      }
      if (!acao) continue; // linha vazia de separação
      brutas.push({ linha: i + 1, r: r, projeto: projeto, marco: marco, grupo: grupo, acao: acao });
    }

    var vistos = {};
    var refPorProjeto = {};
    var tarefas = brutas.map(function (b) {
      var r = b.r;
      var ref = refPorProjeto[b.projeto] || null;
      var ini = resolveData(r[COL.inicio], r[COL.diaIni], ref);
      var fim = resolveData(r[COL.fim], r[COL.diaFim], ini.iso || ref);
      if (fim.iso || ini.iso) refPorProjeto[b.projeto] = fim.iso || ini.iso;
      var marcoNome = b.marco ? b.marco.replace(/^marco/i, 'MARCO') : '';
      var rotulo = marcos[b.projeto + '|' + b.marco.toUpperCase()];
      var acaoNome = marcoNome && rotulo ? marcoNome + ' · ' + rotulo : (marcoNome || '—');
      var notasData = [];
      if (ini.nota) notasData.push('início: ' + ini.nota);
      if (fim.nota) notasData.push('fim: ' + fim.nota);
      if (ini.iso && fim.iso && fim.iso < ini.iso) notasData.push('fim anterior ao início na planilha');
      var grupoFinal = grupoCanonico(b.grupo);
      if (/^marco\s*\d+$/i.test(b.grupo)) { notasData.push('grupo preenchido como "' + b.grupo + '" na planilha'); grupoFinal = 'Sem grupo'; }

      var base = [b.projeto, b.marco, grupoFinal, b.acao].join('|');
      vistos[base] = (vistos[base] || 0) + 1;

      return {
        chave: base + '#' + vistos[base],
        linha: b.linha,
        projeto: b.projeto,
        marco: marcoNome,
        acao: acaoNome,
        atividade: b.acao,
        grupo: grupoFinal,
        grupoOriginal: b.grupo,
        local: limpa(r[COL.local]),
        qtde: formataQtde(r[COL.qtde]),
        sla: limpa(r[COL.sla]),
        status: normalizaStatus(r[COL.status]),
        inicio: ini.iso,
        fim: fim.iso,
        notaData: notasData.join('; ')
      };
    });

    var corrigidas = tarefas.filter(function (t) { return /corrigido/.test(t.notaData); }).length;
    var naoConf = tarefas.filter(function (t) { return /não confirmado/.test(t.notaData); }).length;
    if (corrigidas) avisos.push(corrigidas + ' tarefa(s) com ano 19xx na planilha, tratado como 20xx.');
    if (naoConf) avisos.push(naoConf + ' tarefa(s) com dia/mês que não pôde ser confirmado pelo dia da semana.');
    return { tarefas: tarefas, avisos: avisos };
  }

  // Lê um workbook SheetJS (XLSX.read(..., {cellDates:false}))
  function parseWorkbook(XLSX, wb) {
    var nome = wb.SheetNames.filter(function (n) { return limpa(n).toUpperCase() === 'BD'; })[0];
    if (!nome) throw new Error('A planilha não tem uma aba chamada "BD".');
    var ws = wb.Sheets[nome];
    var rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
    // sheet_to_json começa no início do range usado; ajusta para a linha 1 real
    var ref = XLSX.utils.decode_range(ws['!ref']);
    var pad = [];
    for (var i = 0; i < ref.s.r; i++) pad.push([]);
    if (ref.s.c > 0) rows = rows.map(function (row) { return new Array(ref.s.c).fill(null).concat(row); });
    return parseRows(pad.concat(rows));
  }

  var api = { parseRows: parseRows, parseWorkbook: parseWorkbook, resolveData: resolveData,
              grupoCanonico: grupoCanonico };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.BDParser = api;
})(typeof window !== 'undefined' ? window : this);
