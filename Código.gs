// ============================================================
// COBERTURA DE ESCALAS — SAD BH
// Codigo.gs — versão completa e unificada
// ============================================================

const SPREADSHEET_ID = '1HNHPvB65_ilfNyQfsLU-Y-j8ly664sNA0GAOshA_ZHQ';
const SENHA_GESTOR   = 'Sad@pbh2026';
const PERFIS = { GESTOR: 'gestor', FERISTA: 'ferista', EQUIPE: 'equipe' };

// ============================================================
// ENTRY POINTS
// ============================================================

function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};

  if (params.acao) {
    try {
      let resultado;
      switch (params.acao) {
        case 'escala':           resultado = getEscalaSemanal(params.inicio); break;
        case 'feristas':         resultado = getFeristasSemanal(params.inicio); break;
        case 'tecnicos':         resultado = getTecnicosSemanal(params.inicio); break;
        case 'gestor':           resultado = getDadosGestor(params.inicio); break;
        case 'semanas':          resultado = getSemanasDisponiveis(); break;
        case 'equipes_lista':    resultado = getEquipesLista(); break;
        case 'listar_ausencias': resultado = listarAusencias(params.inicio); break;
        case 'meu_acesso':       resultado = { ok: true }; break;
        case 'cad_acessos':      resultado = getCadAcessos(); break;
        default: resultado = { erro: 'Acao nao reconhecida: ' + params.acao };
      }
      return jsonResponse(resultado);
    } catch (err) {
      return jsonResponse({ erro: err.message });
    }
  }

  const email  = Session.getActiveUser().getEmail();
  const acesso = verificarAcesso(email);

  if (!acesso) {
    return HtmlService.createHtmlOutputFromFile('acesso_negado')
      .setTitle('Acesso negado - SAD BH')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  const pagina = params.p || 'frontend_integracao';
  return HtmlService.createHtmlOutputFromFile(pagina)
    .setTitle('Cobertura de Escalas - SAD BH')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  const email  = Session.getActiveUser().getEmail();
  const acesso = verificarAcesso(email);
  if (!acesso) return jsonResponse({ ok: false, erro: 'Acesso negado.' });
  if (acesso.perfil !== PERFIS.GESTOR) return jsonResponse({ ok: false, erro: 'Apenas o gestor pode realizar esta acao.' });

  try {
    const body = JSON.parse(e.postData.contents);
    let resultado;
    switch (body.acao || '') {
      case 'autorizar_extra':         resultado = autorizarExtra(body); break;
      case 'registrar_ausencia':      resultado = registrarAusencia(body); break;
      case 'registrar_ausencia_lote': resultado = registrarAusenciaLote(body); break;
      case 'excluir_ausencia':        resultado = excluirAusencia(body); break;
      case 'gerar_escala':            resultado = gerarEscalaViaWeb(body); break;
      case 'salvar_ajustes':          resultado = salvarAjustes(body); break;
      case 'publicar_escala':         resultado = publicarEscala(body); break;
      case 'salvar_acesso':           resultado = salvarAcesso(body); break;
      case 'excluir_acesso':          resultado = excluirAcesso(body); break;
      default: resultado = { erro: 'Acao POST nao reconhecida: ' + body.acao };
    }
    return jsonResponse(resultado);
  } catch (err) {
    return jsonResponse({ erro: err.message });
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// AUTENTICACAO POR E-MAIL
// ============================================================

function verificarAcesso(email) {
  if (!email) return null;
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName('cad_acessos');

  if (!sh) {
    sh = ss.insertSheet('cad_acessos');
    sh.getRange(1, 1, 1, 4).setValues([['Email','Perfil','Nome','Referencia']]);
    sh.getRange(1, 1, 1, 4).setFontWeight('bold');
    sh.appendRow([email, 'gestor', 'Gestor', '']);
  }

  if (sh.getLastRow() < 2) return null;
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();

  for (const r of vals) {
    const emailCad = (r[0] || '').toString().trim().toLowerCase();
    const perfil   = (r[1] || '').toString().trim().toLowerCase();
    const nome     = (r[2] || '').toString().trim();
    const ref      = (r[3] || '').toString().trim();
    if (emailCad === email.toLowerCase().trim()) {
      return { ok: true, email, perfil, nome, ref };
    }
  }
  return null;
}

function getCadAcessos() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName('cad_acessos');
  if (!sh || sh.getLastRow() < 2) return { ok: true, acessos: [] };
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
  const acessos = vals.map((r, i) => ({
    row:   i + 2,
    email: (r[0] || '').toString().trim(),
    perfil:(r[1] || '').toString().trim(),
    nome:  (r[2] || '').toString().trim(),
    ref:   (r[3] || '').toString().trim()
  })).filter(a => a.email);
  return { ok: true, acessos };
}

function salvarAcesso(body) {
  const { email, perfil, nome, ref } = body;
  if (!email || !perfil) return { ok: false, erro: 'Email e perfil obrigatorios.' };
  if (!['gestor','ferista','equipe'].includes(perfil)) return { ok: false, erro: 'Perfil invalido.' };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName('cad_acessos');
  if (!sh) return { ok: false, erro: 'Aba cad_acessos nao encontrada.' };

  if (sh.getLastRow() > 1) {
    const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < vals.length; i++) {
      if ((vals[i][0] || '').toString().trim().toLowerCase() === email.toLowerCase()) {
        sh.getRange(i + 2, 1, 1, 4).setValues([[email, perfil, nome || '', ref || '']]);
        return { ok: true, acao: 'atualizado' };
      }
    }
  }
  sh.appendRow([email, perfil, nome || '', ref || '']);
  return { ok: true, acao: 'criado' };
}

function excluirAcesso(body) {
  if (!body.row) return { ok: false, erro: 'Linha nao informada.' };
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName('cad_acessos');
  if (!sh) return { ok: false, erro: 'Aba nao encontrada.' };
  sh.deleteRow(body.row);
  return { ok: true };
}

// ============================================================
// GET: ESCALA SEMANAL
// ============================================================

function getEscalaSemanal(inicioParam) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tz = ss.getSpreadsheetTimeZone();
  const { inicio, fim } = resolverSemana(inicioParam, tz);
  const diasSemana = getDiasSemana(inicio);

  const shAuto = ss.getSheetByName('ALOCACAO_AUTO');
  const autoVals = lerAbaSemana(shAuto, inicio, fim, 8);

  const shEq = ss.getSheetByName('cad_equipes');
  const eqVals = shEq && shEq.getLastRow() > 1
    ? shEq.getRange(2, 1, shEq.getLastRow() - 1, 5).getValues() : [];

  const turnoEquipe = new Map();
  const equipesUnicas = new Set();
  eqVals.forEach(r => {
  const eq = toText(r[0]);
  const turno = normTurno(r[1]);
  if (!eq) return;
  equipesUnicas.add(eq);
  if (turno) turnoEquipe.set(`${eq}|${turno}`, turno);
  });

  const shBase = ss.getSheetByName('escala_base');
  const baseVals = lerAbaSemana(shBase, inicio, fim, 4);
  const extras = getExtrasAtivos(ss, inicio, fim, tz);

  const mapaAlocacao = new Map();
  autoVals.forEach(r => {
    if (!(r[0] instanceof Date)) return;
    const k = `${toText(r[2])}|${normTurno(r[1])}|${keyDia(r[0], tz)}`;
    const tipo = toText(r[6]).toUpperCase();
    mapaAlocacao.set(k, {
      tipo: tipo === 'COBERTURA' ? 'cob' : 'apo',
      nome: toText(r[5]) || '--',
      cat:  toText(r[3]).toLowerCase() === 'medico' ? 'med' : 'enf'
    });
  });

  extras.forEach(ex => {
    const k = `${ex.equipe}|${ex.turno}|${ex.dKey}`;
    mapaAlocacao.set(k, {
      tipo: 'ext',
      nome: ex.profissional || '--',
      cat:  ex.categoria === 'medico' ? 'med' : ex.categoria === 'enfermeiro' ? 'enf' : 'tec'
    });
  });

  const mapaBuracos = new Map();
  baseVals.forEach(r => {
  if (!(r[0] instanceof Date) || !toText(r[1])) return;
      const cat = toText(r[2]).toLowerCase() === 'medico' ? 'med' : 'enf';
      mapaBuracos.set(`${toText(r[1])}|${normTurno(r[3])}|${keyDia(r[0], tz)}`, cat);
   });

  const todasEquipes = new Set();
  eqVals.forEach(r => { if (toText(r[0])) todasEquipes.add(toText(r[0])); });
  baseVals.forEach(r => { if (toText(r[1])) todasEquipes.add(toText(r[1])); });

  const resultado = [];
  todasEquipes.forEach(equipe => {
  const temManha = turnoEquipe.has(`${equipe}|MANHA`);
  const temTarde = turnoEquipe.has(`${equipe}|TARDE`);
  const turnos = [];
  if (temManha) turnos.push('MANHA');
  if (temTarde) turnos.push('TARDE');
  if (!turnos.length) turnos.push('MANHA', 'TARDE');

    const turnosData = turnos.map(turno => {
      const dias = diasSemana.map(d => {
        const k   = `${equipe}|${turno}|${keyDia(d, tz)}`;
        const aloc = mapaAlocacao.get(k);
        if (aloc) return aloc;
          if (mapaBuracos.has(k)) return { tipo: 'sem', nome: '--', cat: mapaBuracos.get(k) };
          return { tipo: 'vaga', nome: '', cat: '' };
      });
      return { turno: turno === 'MANHA' ? 'manha' : 'tarde', dias };
    });

    resultado.push({ equipe, turnos: turnosData });
  });

  resultado.sort((a, b) => a.equipe.localeCompare(b.equipe));

  return {
    ok: true,
    semana: {
      inicio: Utilities.formatDate(inicio, tz, 'dd/MM/yyyy'),
      fim:    Utilities.formatDate(new Date(fim.getTime() - 86400000), tz, 'dd/MM/yyyy'),
      label:  formatarLabelSemana(inicio, fim, tz)
    },
    dias: diasSemana.map(d => Utilities.formatDate(d, tz, 'dd/MM')),
    equipes: resultado
  };
}

// ============================================================
// GET: FERISTAS SEMANAL
// ============================================================

function getFeristasSemanal(inicioParam) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tz = ss.getSpreadsheetTimeZone();
  const { inicio, fim } = resolverSemana(inicioParam, tz);
  const diasSemana = getDiasSemana(inicio);

  const shFer = ss.getSheetByName('cad_feristas');
  const ferVals = shFer && shFer.getLastRow() > 1
    ? shFer.getRange(2, 1, shFer.getLastRow() - 1, 3).getValues() : [];

  const shAuto = ss.getSheetByName('ALOCACAO_AUTO');
  const autoVals = lerAbaSemana(shAuto, inicio, fim, 8);
  const extras = getExtrasAtivos(ss, inicio, fim, tz);

  const mapaFer = new Map();
  autoVals.forEach(r => {
    if (!(r[0] instanceof Date) || !toText(r[5])) return;
    const k = `${toText(r[5])}|${keyDia(r[0], tz)}`;
    if (!mapaFer.has(k)) mapaFer.set(k, []);
    mapaFer.get(k).push({
      tipo:  toText(r[6]).toUpperCase() === 'COBERTURA' ? 'cob' : 'apo',
      local: toText(r[2]),
      turno: normTurno(r[1]) === 'MANHA' ? 'manha' : 'tarde'
    });
  });

  extras.forEach(ex => {
    if (!ex.profissional) return;
    const k = `${ex.profissional}|${ex.dKey}`;
    if (!mapaFer.has(k)) mapaFer.set(k, []);
    mapaFer.get(k).push({ tipo: 'ext', local: ex.equipe, turno: ex.turno === 'MANHA' ? 'manha' : 'tarde' });
  });

  const resultado = [];
  ferVals.forEach(r => {
    const nome = toText(r[0]);
    const cat  = toText(r[1]).toLowerCase();
    const disp = toText(r[2]).toUpperCase();
    if (!nome || disp !== 'SIM') return;
    if (cat !== 'medico' && cat !== 'enfermeiro') return;

    const partes  = nome.split(' ').filter(Boolean);
    const iniciais = ((partes[0]||'')[0] + (partes[partes.length-1]||'')[0]).toUpperCase();

    const dias = diasSemana.map(d => ({
      slots: mapaFer.get(`${nome}|${keyDia(d, tz)}`) || []
    }));

    resultado.push({ nome, cat, ini: iniciais, dias });
  });

  return { ok: true, feristas: resultado };
}

// ============================================================
// GET: TECNICOS SEMANAL
// ============================================================

function getTecnicosSemanal(inicioParam) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tz = ss.getSpreadsheetTimeZone();
  const { inicio, fim } = resolverSemana(inicioParam, tz);
  const diasSemana = getDiasSemana(inicio);

  const shTec = ss.getSheetByName('ausencias_tecnicos');
  if (!shTec) return { ok: false, erro: 'Aba ausencias_tecnicos nao encontrada.' };

  const tecVals = shTec.getLastRow() > 1
    ? shTec.getRange(2, 1, shTec.getLastRow() - 1, 5).getValues() : [];

  const extras  = getExtrasAtivos(ss, inicio, fim, tz, 'tecnico');
  const mapaTec = new Map();

  tecVals.forEach(r => {
    if (!(r[0] instanceof Date) || !toText(r[1])) return;
    const d = onlyDate(r[0]);
    if (d < inicio || d >= fim) return;
    const k    = `${toText(r[1])}|${keyDia(r[0], tz)}`;
    const tipo = toText(r[3]).toLowerCase() === 'coberto' ? 'tec-ok' : 'tec-sem';
    mapaTec.set(k, { tipo, nome: toText(r[2]) || '--' });
  });

  extras.forEach(ex => {
    mapaTec.set(`${ex.equipe}|${ex.dKey}`, { tipo: 'tec-ext', nome: ex.profissional || 'Extra' });
  });

  const bases = [...new Set(tecVals.map(r => toText(r[1])).filter(Boolean))].sort();
  const resultado = bases.map(base => ({
    base,
    dias: diasSemana.map(d => mapaTec.get(`${base}|${keyDia(d, tz)}`) || { tipo: 'vaga', nome: '' })
  }));

  return { ok: true, tecnicos: resultado };
}

// ============================================================
// GET: DADOS DO GESTOR
// ============================================================

function getDadosGestor(inicioParam) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tz = ss.getSpreadsheetTimeZone();
  const { inicio, fim } = resolverSemana(inicioParam, tz);

  const shBase = ss.getSheetByName('escala_base');
  const shAuto = ss.getSheetByName('ALOCACAO_AUTO');
  const baseVals = lerAbaSemana(shBase, inicio, fim, 4);
  const autoVals = lerAbaSemana(shAuto, inicio, fim, 8);
  const extras   = getExtrasAtivos(ss, inicio, fim, tz);

  let cob=0, sem=0, apo=0, ext=extras.length;
  let cobM=0, cobE=0, semM=0, semE=0, apoM=0, apoE=0;

  const buracosSet  = new Set();
  const cobertosSet = new Set();

  baseVals.forEach(r => {
    if (!(r[0] instanceof Date) || !toText(r[1])) return;
    buracosSet.add(`${keyDia(r[0],tz)}|${normTurno(r[3])}|${toText(r[1])}|${toText(r[2]).toLowerCase()}`);
  });

  autoVals.forEach(r => {
    if (!(r[0] instanceof Date)) return;
    const cat  = toText(r[3]).toLowerCase();
    const tipo = toText(r[6]).toUpperCase();
    if (tipo === 'COBERTURA') {
      cobertosSet.add(`${keyDia(r[0],tz)}|${normTurno(r[1])}|${toText(r[2])}|${cat}`);
      cob++; cat === 'medico' ? cobM++ : cobE++;
    } else if (tipo === 'APOIO') {
      apo++; cat === 'medico' ? apoM++ : apoE++;
    }
  });

  const naoCobertos = [];
  buracosSet.forEach(k => {
    if (!cobertosSet.has(k)) {
      const pts = k.split('|');
      naoCobertos.push({ dKey: pts[0], turno: pts[1], equipe: pts[2], cat: pts[3] });
    }
  });

  sem = naoCobertos.length;
  naoCobertos.forEach(x => x.cat === 'medico' ? semM++ : semE++);
  const perc = cob+sem > 0 ? Math.round(cob/(cob+sem)*100) : 100;

  return {
    ok: true,
    stats: { cob, sem, apo, ext, perc },
    categorias: {
      medico:     { cob: cobM, apo: apoM, sem: semM },
      enfermeiro: { cob: cobE, apo: apoE, sem: semE }
    },
    naoCobertos
  };
}

// ============================================================
// GET: EQUIPES LISTA
// ============================================================

function getEquipesLista() {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh  = ss.getSheetByName('cad_equipes');
  if (!sh || sh.getLastRow() < 2) return { ok: true, equipes: [] };
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  const equipes = [...new Set(vals.map(r => toText(r[0])).filter(Boolean))].sort();
  return { ok: true, equipes };
}

// ============================================================
// GET: SEMANAS DISPONIVEIS
// ============================================================

function getSemanasDisponiveis() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tz = ss.getSpreadsheetTimeZone();
  const sh = ss.getSheetByName('escala_base');
  if (!sh || sh.getLastRow() < 2) return { ok: true, semanas: [] };

  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  const semanasSet = new Set();

  vals.forEach(r => {
    if (!(r[0] instanceof Date)) return;
    const d   = onlyDate(r[0]);
    const dow  = d.getDay();
    const diff = dow === 0 ? -6 : 1 - dow;
    const seg  = new Date(d);
    seg.setDate(d.getDate() + diff);
    semanasSet.add(keyDia(seg, tz));
  });

  const semanas = [...semanasSet].sort().reverse().map(k => {
    const p   = k.split('-');
    const seg = new Date(+p[0], +p[1]-1, +p[2]);
    const sex = new Date(seg); sex.setDate(seg.getDate() + 4);
    return {
      inicio: k,
      label:  `${Utilities.formatDate(seg, tz, 'dd/MM')} - ${Utilities.formatDate(sex, tz, 'dd/MM/yyyy')}`
    };
  });

  return { ok: true, semanas };
}

// ============================================================
// GET: LISTAR AUSENCIAS
// ============================================================

function listarAusencias(inicioParam) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tz = ss.getSpreadsheetTimeZone();
  const resultado = [];

  const shBase = ss.getSheetByName('escala_base');
  if (shBase && shBase.getLastRow() > 1) {
    shBase.getRange(2, 1, shBase.getLastRow()-1, 7).getValues().forEach((r, i) => {
      if (!(r[0] instanceof Date)) return;
      resultado.push({
        row:         i + 2,
        data:        keyDia(r[0], tz),
        equipe:      toText(r[1]),
        categoria:   toText(r[2]).toLowerCase(),
        turno:       normTurno(r[3]),
        motivo:      toText(r[4]),
        obs:         toText(r[5]),
        profissional:toText(r[6])
      });
    });
  }

  const shTec = ss.getSheetByName('ausencias_tecnicos');
  if (shTec && shTec.getLastRow() > 1) {
    shTec.getRange(2, 1, shTec.getLastRow()-1, 5).getValues().forEach((r, i) => {
      if (!(r[0] instanceof Date)) return;
      resultado.push({
        row:         -(i + 2),
        data:        keyDia(r[0], tz),
        equipe:      toText(r[1]),
        categoria:   'tecnico',
        turno:       '',
        profissional:toText(r[2]),
        motivo:      toText(r[4]),
        obs:         toText(r[4])
      });
    });
  }

  resultado.sort((a, b) => b.data.localeCompare(a.data));
  return { ok: true, ausencias: resultado };
}

// ============================================================
// POST: AUTORIZAR EXTRA
// ============================================================

function autorizarExtra(body) {
  const { equipe, turno, dias, categoria, tipo, profissional, motivo, obs } = body;
  if (!equipe || !dias || !dias.length) return { ok: false, erro: 'Campos obrigatorios ausentes.' };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName('cad_extras');
  if (!sh) {
    sh = ss.insertSheet('cad_extras');
    sh.getRange(1,1,1,9).setValues([['Data','Equipe','Turno','Categoria','Tipo','Profissional','Motivo','Obs','Autorizado_em']]);
    sh.getRange(1,1,1,9).setFontWeight('bold');
  }

  const agora  = new Date();
  const linhas = dias.map(dStr => {
    const p = dStr.split('/');
    return [
      new Date(+p[2], +p[1]-1, +p[0]),
      equipe||'', normTurno(turno)||'', categoria||'',
      tipo||'Plantao extra pago', profissional||'', motivo||'', obs||'', agora
    ];
  });

  sh.getRange(sh.getLastRow()+1, 1, linhas.length, 9).setValues(linhas);
  return { ok: true, gravados: linhas.length };
}

// ============================================================
// POST: REGISTRAR AUSENCIA (unitaria)
// ============================================================

function registrarAusencia(body) {
  const { data, equipe, categoria, turno, motivo, obs } = body;
  if (!data || !equipe || !categoria) return { ok: false, erro: 'Campos obrigatorios: data, equipe, categoria.' };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const p  = data.split('/');
  const dt = new Date(+p[2], +p[1]-1, +p[0]);

  if (categoria === 'tecnico') {
    let sh = ss.getSheetByName('ausencias_tecnicos');
    if (!sh) {
      sh = ss.insertSheet('ausencias_tecnicos');
      sh.getRange(1,1,1,5).setValues([['Data','Base','Profissional','Status','Obs']]);
    }
    sh.appendRow([dt, equipe, '', 'sem', obs||'']);
  } else {
    const sh = ss.getSheetByName('escala_base');
    if (!sh) return { ok: false, erro: 'Aba escala_base nao encontrada.' };
    sh.appendRow([dt, equipe, categoria, normTurno(turno), motivo||'', obs||'']);
  }
  return { ok: true };
}

// ============================================================
// POST: REGISTRAR AUSENCIA LOTE
// ============================================================

function registrarAusenciaLote(body) {
  const { categoria, equipe, turno, profissional, motivo, obs, dias } = body;
  if (!equipe || !motivo || !dias || !dias.length) return { ok: false, erro: 'Campos obrigatorios: equipe, motivo, dias.' };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let gravados = 0;

  if (categoria === 'tecnico') {
    let sh = ss.getSheetByName('ausencias_tecnicos');
    if (!sh) {
      sh = ss.insertSheet('ausencias_tecnicos');
      sh.getRange(1,1,1,5).setValues([['Data','Base','Profissional','Status','Motivo']]);
      sh.getRange(1,1,1,5).setFontWeight('bold');
    }
    const linhas = dias.map(d => {
      const p = d.split('-');
      return [new Date(+p[0],+p[1]-1,+p[2]), equipe, profissional||'', 'sem', motivo||''];
    });
    sh.getRange(sh.getLastRow()+1, 1, linhas.length, 5).setValues(linhas);
    gravados = linhas.length;
  } else {
    const sh = ss.getSheetByName('escala_base');
    if (!sh) return { ok: false, erro: 'Aba escala_base nao encontrada.' };
    if (sh.getLastColumn() < 7) sh.getRange(1,7).setValue('Profissional');
    const linhas = dias.map(d => {
      const p = d.split('-');
      return [new Date(+p[0],+p[1]-1,+p[2]), equipe, categoria, normTurno(turno), motivo||'', obs||'', profissional||''];
    });
    sh.getRange(sh.getLastRow()+1, 1, linhas.length, 7).setValues(linhas);
    gravados = linhas.length;
  }
  return { ok: true, gravados };
}

// ============================================================
// POST: EXCLUIR AUSENCIA
// ============================================================

function excluirAusencia(body) {
  const { row } = body;
  if (!row) return { ok: false, erro: 'Linha nao informada.' };
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(row > 0 ? 'escala_base' : 'ausencias_tecnicos');
  if (!sh) return { ok: false, erro: 'Aba nao encontrada.' };
  sh.deleteRow(Math.abs(row));
  return { ok: true };
}

// ============================================================
// POST: GERAR ESCALA VIA WEB
// ============================================================

function gerarEscalaViaWeb(body) {
  const { inicio } = body;
  if (!inicio) return { ok: false, erro: 'Data de inicio obrigatoria.' };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const shRegras = ss.getSheetByName('REGRAS_SEMANA');
  if (!shRegras) return { ok: false, erro: 'Aba REGRAS_SEMANA nao encontrada.' };

  const p = inicio.split('-');
  const dataInicio = new Date(+p[0], +p[1]-1, +p[2]);
  shRegras.getRange('A1').setValue('INICIO_SEMANA');
  shRegras.getRange('B1').setValue(dataInicio);

  try {
    // Chama diretamente as funções internas sem passar por getUi()
    const tz = ss.getSpreadsheetTimeZone();
    const fimSemana = new Date(dataInicio);
    fimSemana.setDate(dataInicio.getDate() + 5);

    const shEquipes  = ss.getSheetByName('cad_equipes');
    const shFeristas = ss.getSheetByName('cad_feristas');
    const shBuracos  = ss.getSheetByName('escala_base');
    const shRegrasAba = ss.getSheetByName('REGRAS_SEMANA');

    if (!shEquipes || !shFeristas || !shBuracos || !shRegrasAba) {
      return { ok: false, erro: 'Faltam abas obrigatorias.' };
    }

    // Roda a geração sem alertas de UI
    gerarEscalaSemanal_semUI();

    const stats = calcularStatsCobertura_(
      shBuracos,
      ss.getSheetByName('ALOCACAO_AUTO'),
      dataInicio, fimSemana, tz
    );

    return {
      ok: true,
      cobertos:    stats.global.cobertos,
      naoCobertos: stats.global.naoCobertos,
      apoios:      stats.global.apoios,
      perc:        stats.global.perc
    };
  } catch(err) {
    return { ok: false, erro: err.message };
  }
}

function gerarEscalaSemanal_semUI() {
  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone();

  const shEquipes  = ss.getSheetByName('cad_equipes');
  const shFeristas = ss.getSheetByName('cad_feristas');
  const shBuracos  = ss.getSheetByName('escala_base');
  const shRegras   = ss.getSheetByName('REGRAS_SEMANA');

  const inicioSemana = getInicioSemana_(shRegras);
  const fimSemana = new Date(inicioSemana);
  fimSemana.setDate(fimSemana.getDate() + 5);

  const regras = lerRegrasSemana_(shRegras);

  // Lê equipes
  const eqLast = shEquipes.getLastRow();
  const eqVals = eqLast > 1 ? shEquipes.getRange(2, 1, eqLast-1, 5).getValues() : [];
  const pesoEquipe = new Map();
  const turnoEquipe = new Map();
  eqVals.forEach(r => {
    const equipe = toText_(r[0]);
    const turno  = normTurno_(r[1]);
    const peso   = Number(r[4]);
    if (!equipe) return;
    pesoEquipe.set(equipe, isFinite(peso) ? peso : 0);
    turnoEquipe.set(equipe, turno);
  });

  const equipesOrdenadasPorPeso = Array.from(pesoEquipe.entries())
    .map(([equipe, peso]) => ({ equipe, peso }))
    .sort((a, b) => (b.peso||0) - (a.peso||0));

  // Lê feristas
  const ferLast = shFeristas.getLastRow();
  const ferVals = ferLast > 1 ? shFeristas.getRange(2, 1, ferLast-1, 3).getValues() : [];
  const feristasPorCat = { medico: [], enfermeiro: [] };
  const catPorFerista = new Map();
  ferVals.forEach(r => {
    const nome = toText_(r[0]);
    const cat  = toText_(r[1]).toLowerCase();
    const disp = toText_(r[2]).toUpperCase();
    if (!nome || disp !== 'SIM') return;
    if (cat !== 'medico' && cat !== 'enfermeiro') return;
    feristasPorCat[cat].push(nome);
    catPorFerista.set(nome, cat);
  });

  // Lê buracos
  const burLast = shBuracos.getLastRow();
  const burVals = burLast > 1 ? shBuracos.getRange(2, 1, burLast-1, 4).getValues() : [];
  let demandas = [];
  burVals.forEach((r, idx) => {
    const dt = r[0];
    const equipe = toText_(r[1]);
    const categoria = toText_(r[2]).toLowerCase();
    const turnoDigitado = normTurno_(r[3]);
    if (!(dt instanceof Date)) return;
    const data = onlyDate_(dt);
    if (data < inicioSemana || data >= fimSemana) return;
    if (!equipe) return;
    if (categoria !== 'medico' && categoria !== 'enfermeiro') return;
    const turnoCad = turnoEquipe.get(equipe) || '';
    const turno = turnoCad || turnoDigitado;
    if (!turno) return;
    demandas.push({
      data, equipe, categoria, turno,
      peso: pesoEquipe.get(equipe) || 0,
      origemRow: idx + 2
    });
  });

  demandas.sort((a, b) => {
    const da = a.data.getTime(), db = b.data.getTime();
    if (da !== db) return da - db;
    if (a.peso !== b.peso) return b.peso - a.peso;
    if (a.turno !== b.turno) return a.turno === 'MANHA' ? -1 : 1;
    return a.equipe.localeCompare(b.equipe);
  });

  const usadosDiaCat = new Map();
  const usadosDiaCatTurno = new Map();
  const teamWeekCount = new Map();
  const feristaWeekCount = new Map();
  const feristaTeamCount = new Map();
  const lastDayFeristaTeam = new Map();
  const apoioTeamDayCount = new Map();
  const apoioFeristaTeamCount = new Map();
  const apoioLastDayFeristaTeam = new Map();
  const apoioTeamWeekCount = new Map();
  const alocacoes = [];
  const diasSemana = [];

  for (let i = 0; i < 5; i++) {
    const d = new Date(inicioSemana);
    d.setDate(d.getDate() + i);
    diasSemana.push(d);
  }

  // Cobertura normal
  for (const dem of demandas) {
    const pool = feristasPorCat[dem.categoria] || [];
    if (!pool.length) continue;
    const escolhido = pickBestCandidate_(
      dem, pool, usadosDiaCat, usadosDiaCatTurno,
      teamWeekCount, feristaWeekCount, feristaTeamCount,
      lastDayFeristaTeam, regras, tz
    );
    if (!escolhido) continue;

    registrarUsoDiaCat_(usadosDiaCat, dem.data, dem.categoria, escolhido);
    registrarUsoDiaCatTurno_(usadosDiaCatTurno, dem.data, dem.categoria, escolhido, dem.turno);
    incrementarMapa_(teamWeekCount, `${dem.equipe}|${dem.categoria}`);
    incrementarMapa_(feristaWeekCount, `${escolhido}|${dem.categoria}`);
    incrementarMapa_(feristaTeamCount, `${escolhido}|${dem.equipe}|${dem.categoria}`);
    lastDayFeristaTeam.set(`${escolhido}|${dem.equipe}|${dem.categoria}`, keyDia_(dem.data, tz));

    alocacoes.push({
      data: dem.data, turno: dem.turno, equipe: dem.equipe,
      categoria: dem.categoria, peso: dem.peso,
      ferista: escolhido, tipo: 'COBERTURA', origemRow: dem.origemRow
    });
  }

  escreverAlocacaoAuto_(ss, alocacoes);
  escreverVisualSemanal_(ss, alocacoes, ferVals, diasSemana, tz);
  atualizarVisualEquipes_(alocacoes, inicioSemana);
  atualizarPublicarSemana_();
  formatarVisuais_();
  aplicarValidacoes_();
  atualizarDashboard_();
  atualizarHistoricoCobertura_();
}


// ============================================================
// POST: SALVAR AJUSTES
// ============================================================

function salvarAjustes(body) {
  const { inicio, alteracoes } = body;
  if (!alteracoes || !alteracoes.length) return { ok: true, gravados: 0 };

  const ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tz     = ss.getSpreadsheetTimeZone();
  const shAuto = ss.getSheetByName('ALOCACAO_AUTO');
  if (!shAuto || shAuto.getLastRow() < 2) return { ok: false, erro: 'ALOCACAO_AUTO vazia.' };

  const p = inicio.split('-');
  const inicioDate = new Date(+p[0], +p[1]-1, +p[2]); inicioDate.setHours(0,0,0,0);
  const vals = shAuto.getRange(2, 1, shAuto.getLastRow()-1, 8).getValues();
  let modificados = 0;

  alteracoes.forEach(alt => {
    const diaDate = new Date(inicioDate);
    diaDate.setDate(inicioDate.getDate() + alt.diaIdx);
    const dKey = keyDia(diaDate, tz);

    for (let i = 0; i < vals.length; i++) {
      const r = vals[i];
      if (!(r[0] instanceof Date)) continue;
      if (keyDia(r[0],tz) === dKey && normTurno(r[1]) === normTurno(alt.turno) && toText(r[2]) === alt.equipe) {
        const rowNum = i + 2;
        if (alt.feristaNovo === null) {
          shAuto.getRange(rowNum, 7).setValue('SEM_COBERTURA');
          shAuto.getRange(rowNum, 6).setValue('');
        } else {
          shAuto.getRange(rowNum, 6).setValue(alt.feristaNovo);
          shAuto.getRange(rowNum, 7).setValue('COBERTURA');
        }
        modificados++;
        break;
      }
    }
  });

  try {
    const alocs = lerTodasAlocacoes_(ss, tz);
    escreverVisualSemanal_(ss, alocs, lerFeristas_(ss), getDiasSemanaArr_(inicioDate), tz);
    atualizarVisualEquipes_(alocs, inicioDate);
  } catch(e) { Logger.log('Aviso visuais: ' + e.message); }

  return { ok: true, gravados: modificados };
}

// ============================================================
// POST: PUBLICAR ESCALA
// ============================================================

function publicarEscala(body) {
  try {
    atualizarPublicarSemana_();
    atualizarDashboard_();
    atualizarHistoricoCobertura_();
    return { ok: true };
  } catch(err) {
    return { ok: false, erro: err.message };
  }
}

// ============================================================
// HELPERS MODULO 3
// ============================================================

function lerTodasAlocacoes_(ss, tz) {
  const sh = ss.getSheetByName('ALOCACAO_AUTO');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow()-1, 8).getValues()
    .filter(r => r[0] instanceof Date)
    .map(r => ({
      data:      onlyDate_(r[0]),
      turno:     normTurno_(r[1]),
      equipe:    toText_(r[2]),
      categoria: toText_(r[3]).toLowerCase(),
      peso:      Number(r[4])||0,
      ferista:   toText_(r[5]),
      tipo:      toText_(r[6]).toUpperCase(),
      origemRow: r[7]
    }));
}

function lerFeristas_(ss) {
  const sh = ss.getSheetByName('cad_feristas');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow()-1, 3).getValues();
}

function getDiasSemanaArr_(inicio) {
  return [0,1,2,3,4].map(i => {
    const d = new Date(inicio);
    d.setDate(inicio.getDate() + i);
    d.setHours(0,0,0,0);
    return d;
  });
}

// ============================================================
// HELPERS GERAIS
// ============================================================

function resolverSemana(inicioParam, tz) {
  let inicio;
  if (inicioParam) {
    const p = inicioParam.split('-');
    inicio  = new Date(+p[0], +p[1]-1, +p[2]);
  } else {
    try {
      const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('REGRAS_SEMANA');
      const v  = sh.getRange('B1').getValue();
      if (v instanceof Date) inicio = new Date(v);
    } catch(e) {}
  }
  if (!inicio || isNaN(inicio)) {
    const hoje = new Date();
    const dow  = hoje.getDay();
    inicio     = new Date(hoje);
    inicio.setDate(hoje.getDate() + (dow === 0 ? -6 : 1 - dow));
  }
  inicio.setHours(0,0,0,0);
  const fim = new Date(inicio); fim.setDate(inicio.getDate() + 5);
  return { inicio, fim };
}

function getDiasSemana(inicio) {
  return [0,1,2,3,4].map(i => {
    const d = new Date(inicio); d.setDate(inicio.getDate() + i); d.setHours(0,0,0,0); return d;
  });
}

function lerAbaSemana(sh, inicio, fim, numCols) {
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow()-1, numCols).getValues().filter(r => {
    if (!(r[0] instanceof Date)) return false;
    const d = onlyDate(r[0]);
    return d >= inicio && d < fim;
  });
}

function getExtrasAtivos(ss, inicio, fim, tz, categoriaSo) {
  const sh = ss.getSheetByName('cad_extras');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow()-1, 9).getValues()
    .filter(r => {
      if (!(r[0] instanceof Date)) return false;
      const d = onlyDate(r[0]);
      if (d < inicio || d >= fim) return false;
      if (categoriaSo && toText(r[3]).toLowerCase() !== categoriaSo) return false;
      return true;
    })
    .map(r => ({
      dKey:         keyDia(r[0], tz),
      equipe:       toText(r[1]),
      turno:        normTurno(r[2]),
      categoria:    toText(r[3]).toLowerCase(),
      tipo:         toText(r[4]),
      profissional: toText(r[5]),
      motivo:       toText(r[6])
    }));
}

function formatarLabelSemana(inicio, fim, tz) {
  return Utilities.formatDate(inicio, tz, 'dd/MM') + ' - ' +
         Utilities.formatDate(new Date(fim.getTime()-86400000), tz, 'dd/MM/yyyy');
}

function toText(v)   { return (v || '').toString().trim(); }
function toText_(v)  { return (v || '').toString().trim(); }

function normTurno(v) {
  const t = toText(v).toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'');
  if (t === 'MANHA' || t === 'MANHA') return 'MANHA';
  if (t === 'TARDE') return 'TARDE';
  return t;
}
function normTurno_(v) { return normTurno(v); }

function onlyDate(d)  { const x = new Date(d); x.setHours(0,0,0,0); return x; }
function onlyDate_(d) { return onlyDate(d); }

function keyDia(d, tz) {
  return Utilities.formatDate(d, tz || SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
}
function keyDia_(d, tz) { return keyDia(d, tz); }

function getScriptUrl() { return ScriptApp.getService().getUrl(); }

function chamarAPI(params) {
  const acao = params.acao || '';
  switch (acao) {
    case 'listar_feristas': return listarFeristas();
    case 'escala':           return getEscalaSemanal(params.inicio);
    case 'feristas':         return getFeristasSemanal(params.inicio);
    case 'tecnicos':         return getTecnicosSemanal(params.inicio);
    case 'gestor':           return getDadosGestor(params.inicio);
    case 'semanas':          return getSemanasDisponiveis();
    case 'equipes_lista':    return getEquipesLista();
    case 'listar_ausencias': return listarAusencias(params.inicio);
    case 'cad_acessos':      return getCadAcessos();
    case 'meu_acesso': {
      const email = Session.getActiveUser().getEmail();
      return verificarAcesso(email) || { ok: false };
    }
    default: return { erro: 'Acao nao reconhecida: ' + acao };
  }
}
 
function chamarAPIPost(body) {
  const acao = body.acao || '';
  switch (acao) {
    case 'autorizar_extra':         return autorizarExtra(body);
    case 'registrar_ausencia':      return registrarAusencia(body);
    case 'registrar_ausencia_lote': return registrarAusenciaLote(body);
    case 'excluir_ausencia':        return excluirAusencia(body);
    case 'gerar_escala':            return gerarEscalaViaWeb(body);
    case 'salvar_ajustes':          return salvarAjustes(body);
    case 'publicar_escala':         return publicarEscala(body);
    case 'salvar_acesso':           return salvarAcesso(body);
    case 'excluir_acesso':          return excluirAcesso(body);
    case 'salvar_ferista':  return salvarFerista(body);
    case 'excluir_ferista': return excluirFerista(body);
    default: return { erro: 'Acao POST nao reconhecida: ' + acao };
  }
}


// ============================================================
// PATCH v2 — Gerenciar Feristas com disponibilidade
// Cole estas funções no final do Codigo.gs
// Substitui as funções listarFeristas, salvarFerista, excluirFerista
// ============================================================

// Colunas da aba cad_feristas:
// A: Nome | B: Categoria | C: Disponivel | D: Conselho
// E: Seg_M | F: Seg_T | G: Ter_M | H: Ter_T | I: Qua_M | J: Qua_T
// K: Qui_M | L: Qui_T | M: Sex_M | N: Sex_T

const DISP_COLS = ['seg_m','seg_t','ter_m','ter_t','qua_m','qua_t','qui_m','qui_t','sex_m','sex_t'];

function garantirCabecalhoFeristas_(sh) {
  const cabecalho = ['Nome','Categoria','Disponivel','Conselho',
    'Seg_M','Seg_T','Ter_M','Ter_T','Qua_M','Qua_T','Qui_M','Qui_T','Sex_M','Sex_T'];
  if (sh.getLastColumn() < cabecalho.length) {
    sh.getRange(1, 1, 1, cabecalho.length).setValues([cabecalho]);
    sh.getRange(1, 1, 1, cabecalho.length).setFontWeight('bold');
  }
}

function listarFeristas() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName('cad_feristas');
  if (!sh || sh.getLastRow() < 2) return { ok: true, feristas: [] };

  garantirCabecalhoFeristas_(sh);
  const numCols = Math.max(sh.getLastColumn(), 14);
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, numCols).getValues();

  const feristas = vals.map((r, i) => {
    const disp = {};
    DISP_COLS.forEach((k, idx) => {
      disp[k] = (r[4 + idx] || '').toString().trim().toUpperCase() === 'NAO' ? 'NAO' : 'SIM';
    });
    return {
      row:        i + 2,
      nome:       (r[0] || '').toString().trim(),
      categoria:  (r[1] || '').toString().trim().toLowerCase(),
      disponivel: (r[2] || '').toString().trim().toUpperCase() || 'SIM',
      conselho:   (r[3] || '').toString().trim(),
      disp
    };
  }).filter(f => f.nome);

  return { ok: true, feristas };
}

function salvarFerista(body) {
  const { nome, categoria, disponivel, conselho, row } = body;
  const dispObj = body.disp || {};

  if (!nome || !categoria) return { ok: false, erro: 'Nome e categoria obrigatorios.' };
  if (!['medico','enfermeiro'].includes(categoria)) return { ok: false, erro: 'Categoria invalida.' };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName('cad_feristas');
  if (!sh) return { ok: false, erro: 'Aba cad_feristas nao encontrada.' };

  garantirCabecalhoFeristas_(sh);

  const dispVals = DISP_COLS.map(k => (dispObj[k] || 'SIM').toUpperCase());
  const linha = [nome, categoria, disponivel || 'SIM', conselho || '', ...dispVals];

  if (row) {
    sh.getRange(row, 1, 1, linha.length).setValues([linha]);
    return { ok: true, acao: 'atualizado' };
  } else {
    if (sh.getLastRow() > 1) {
      const existing = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
      for (const r of existing) {
        if ((r[0] || '').toString().trim().toLowerCase() === nome.toLowerCase()) {
          return { ok: false, erro: 'Ja existe um ferista com esse nome.' };
        }
      }
    }
    sh.appendRow(linha);
    return { ok: true, acao: 'criado' };
  }
}

function excluirFerista(body) {
  const { row } = body;
  if (!row) return { ok: false, erro: 'Linha nao informada.' };
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName('cad_feristas');
  if (!sh) return { ok: false, erro: 'Aba nao encontrada.' };
  sh.deleteRow(row);
  return { ok: true };
}

// ============================================================
// PATCH — feristaPodeSerAlocado_ com disponibilidade fixa
// Substitua a função feristaPodeSerAlocado_ existente no Codigo.gs
// por esta versão que consulta a disponibilidade do cadastro
// ============================================================

function feristaPodeSerAlocadoV2_(ferista, dem, regras, mapDisp) {
  // 1) Verifica disponibilidade fixa do cadastro
  if (mapDisp && mapDisp.has(ferista)) {
    const disp = mapDisp.get(ferista);
    const dayCode = ['dom','seg','ter','qua','qui','sex','sab'][dem.data.getDay()];
    const turnoCode = dem.turno === 'MANHA' ? 'm' : 't';
    const key = dayCode + '_' + turnoCode;
    if (disp[key] === 'NAO') return false;
  }

  // 2) NAO_ALOCAR bloqueia
  const bloqueios = regras.filter(r => r.ferista === ferista && r.tipo === 'NAO_ALOCAR');
  for (const r of bloqueios) {
    if (!regraAplicaNoDia_(r, dem.data)) continue;
    if (r.equipe && dem.equipe && r.equipe !== dem.equipe) continue;
    if (r.turno && dem.turno && r.turno !== dem.turno) continue;
    return false;
  }

  // 3) OBRIGATORIO vira restricao dura no dia
  const obrigatorias = regras.filter(r => r.ferista === ferista && r.tipo === 'OBRIGATORIO');
  const obrigDoDia = obrigatorias.filter(r => regraAplicaNoDia_(r, dem.data));
  if (obrigDoDia.length > 0) {
    let compativel = false;
    for (const r of obrigDoDia) {
      const turnoOk  = !r.turno || r.turno === dem.turno;
      const equipeOk = !r.equipe || !dem.equipe || r.equipe === dem.equipe;
      if (turnoOk && equipeOk) { compativel = true; break; }
    }
    if (!compativel) return false;
  }

  // 4) PERMITIR_APENAS restringe
  const permitOnly = regras.filter(r => r.ferista === ferista && r.tipo === 'PERMITIR_APENAS');
  const aplicaveis = permitOnly.filter(r => {
    if (!regraAplicaNoDia_(r, dem.data)) return false;
    if (r.equipe && dem.equipe && r.equipe !== dem.equipe) return false;
    return true;
  });
  if (aplicaveis.length > 0) {
    const turnosPermitidos = new Set(aplicaveis.map(r => r.turno).filter(Boolean));
    if (turnosPermitidos.size > 0 && !turnosPermitidos.has(dem.turno)) return false;
  }

  return true;
}




