function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Feristas')
    .addItem('Gerar escala semanal', 'gerarEscalaSemanal')
    .addSeparator()
    .addItem('Exportar PUBLICAR_SEMANA (PDF)', 'exportarPublicarSemanaPDF')
    .addItem('Limpar saídas da semana', 'limparSemanaSaidas')
    .addToUi();
}

function gerarEscalaSemanal() {
  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone();

  const shEquipes  = ss.getSheetByName('cad_equipes');
  const shFeristas = ss.getSheetByName('cad_feristas');
  const shBuracos  = ss.getSheetByName('escala_base');
  const shRegras   = ss.getSheetByName('REGRAS_SEMANA');

  if (!shEquipes || !shFeristas || !shBuracos || !shRegras) {
    throw new Error('Faltam abas obrigatórias: cad_equipes, cad_feristas, escala_base, REGRAS_SEMANA.');
  }

  const inicioSemana = getInicioSemana_(shRegras);
  const fimSemana = new Date(inicioSemana);
  fimSemana.setDate(fimSemana.getDate() + 5); // exclusivo (Seg-Sex)

  const regras = lerRegrasSemana_(shRegras);
  const avisos = [];

  // =========================
  // CAD_EQUIPES
  // =========================
  const eqLast = shEquipes.getLastRow();
  const eqVals = eqLast > 1 ? shEquipes.getRange(2, 1, eqLast - 1, 5).getValues() : [];

  const pesoEquipe = new Map();
  const turnoEquipe = new Map();

  eqVals.forEach((r, i) => {
    const equipe = toText_(r[0]);
    const turno  = normTurno_(r[1]);
    const peso   = Number(r[4]);

    if (!equipe) return;

    pesoEquipe.set(equipe, isFinite(peso) ? peso : 0);
    turnoEquipe.set(equipe, turno);

    if (!turno) {
      avisos.push(`cad_equipes linha ${i + 2}: equipe "${equipe}" sem turno cadastrado.`);
    }
  });

  const equipesOrdenadasPorPeso = Array.from(pesoEquipe.entries())
    .map(([equipe, peso]) => ({ equipe, peso }))
    .sort((a, b) => (b.peso || 0) - (a.peso || 0));

  // =========================
  // CAD_FERISTAS
  // =========================
  const ferLast = shFeristas.getLastRow();
  const ferVals = ferLast > 1 ? shFeristas.getRange(2, 1, ferLast - 1, 3).getValues() : [];

  const feristasPorCat = { medico: [], enfermeiro: [] };
  const catPorFerista = new Map();

  ferVals.forEach((r, i) => {
    const nome = toText_(r[0]);
    const cat  = toText_(r[1]).toLowerCase();
    const disp = toText_(r[2]).toUpperCase();

    if (!nome || disp !== 'SIM') return;
    if (cat !== 'medico' && cat !== 'enfermeiro') {
      avisos.push(`cad_feristas linha ${i + 2}: categoria inválida para "${nome}".`);
      return;
    }

    feristasPorCat[cat].push(nome);
    catPorFerista.set(nome, cat);
  });

  // =========================
  // ESCALA_BASE -> DEMANDAS
  // =========================
  const burLast = shBuracos.getLastRow();
  const burVals = burLast > 1 ? shBuracos.getRange(2, 1, burLast - 1, 4).getValues() : [];

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

    if (!turno) {
      avisos.push(`escala_base linha ${idx + 2}: sem turno para equipe "${equipe}".`);
      return;
    }

    if (turnoCad && turnoDigitado && turnoCad !== turnoDigitado) {
      avisos.push(`escala_base linha ${idx + 2}: turno divergente em "${equipe}". Usando turno do cadastro (${turnoCad}).`);
    }

    demandas.push({
      data,
      equipe,
      categoria,
      turno,
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

  // =========================
  // CONTROLES DE ALOCAÇÃO
  // =========================
  const usadosDiaCat = new Map();        // yyyy-mm-dd|cat|ferista -> count total no dia
  const usadosDiaCatTurno = new Map();   // yyyy-mm-dd|cat|ferista|turno -> count no turno

  const teamWeekCount = new Map();
  const feristaWeekCount = new Map();
  const feristaTeamCount = new Map();
  const lastDayFeristaTeam = new Map();

  const apoioTeamDayCount = new Map();          // yyyy-mm-dd|cat|equipe
  const apoioFeristaTeamCount = new Map();      // ferista|equipe|cat
  const apoioLastDayFeristaTeam = new Map();    // ferista|equipe|cat -> yyyy-mm-dd
  const apoioTeamWeekCount = new Map();         // equipe|cat

  const alocacoes = [];
  const faltasNaoCobertas = [];

  const diasSemana = [];
  for (let i = 0; i < 5; i++) {
    const d = new Date(inicioSemana);
    d.setDate(d.getDate() + i);
    diasSemana.push(d);
  }

  // =========================
  // 1) REGRAS OBRIGATORIO
  // =========================
  const obrigRules = regras.filter(r => r.tipo === 'OBRIGATORIO');

  obrigRules.forEach(rule => {
    const ferista = rule.ferista;
    if (!catPorFerista.has(ferista)) return;
    const categoria = catPorFerista.get(ferista);

    for (const d of diasSemana) {
      if (!regraAplicaNoDia_(rule, d)) continue;

      const turnoObrigatorio = rule.turno || '';
      const equipesCandidatas = rule.equipe
        ? [rule.equipe]
        : (turnoObrigatorio
            ? equipesOrdenadasPorPeso
                .filter(e => (turnoEquipe.get(e.equipe) || '') === turnoObrigatorio)
                .map(e => e.equipe)
            : equipesOrdenadasPorPeso.map(e => e.equipe));

      let alocou = false;

      // primeiro tenta casar com um buraco real
      for (let i = 0; i < demandas.length; i++) {
        const dem = demandas[i];
        if (!sameDate_(dem.data, d)) continue;
        if (dem.categoria !== categoria) continue;
        if (rule.turno && dem.turno !== rule.turno) continue;
        if (rule.equipe && dem.equipe !== rule.equipe) continue;

        if (!feristaPodeSerAlocado_(ferista, dem, regras)) continue;
        if (!podeUsarNoTurno_(ferista, dem.data, dem.categoria, dem.turno, regras, usadosDiaCat, usadosDiaCatTurno, dem.equipe)) continue;

        registrarUsoDiaCat_(usadosDiaCat, dem.data, dem.categoria, ferista);
        registrarUsoDiaCatTurno_(usadosDiaCatTurno, dem.data, dem.categoria, ferista, dem.turno);

        incrementarMapa_(teamWeekCount, `${dem.equipe}|${dem.categoria}`);
        incrementarMapa_(feristaWeekCount, `${ferista}|${dem.categoria}`);
        incrementarMapa_(feristaTeamCount, `${ferista}|${dem.equipe}|${dem.categoria}`);
        lastDayFeristaTeam.set(`${ferista}|${dem.equipe}|${dem.categoria}`, keyDia_(dem.data, tz));

        alocacoes.push({
          data: dem.data,
          turno: dem.turno,
          equipe: dem.equipe,
          categoria: dem.categoria,
          peso: dem.peso,
          ferista,
          tipo: 'COBERTURA',
          origemRow: dem.origemRow
        });

        demandas.splice(i, 1);
        alocou = true;
        break;
      }

      if (alocou) continue;

      // se não havia buraco compatível, vira apoio
      for (const equipeApoio of equipesCandidatas) {
        const turnoApoio = turnoEquipe.get(equipeApoio) || rule.turno || 'MANHA';
        const demCheck = { data: d, equipe: equipeApoio, categoria, turno: turnoApoio };

        if (!feristaPodeSerAlocado_(ferista, demCheck, regras)) continue;
        if (!podeUsarNoTurno_(ferista, d, categoria, turnoApoio, regras, usadosDiaCat, usadosDiaCatTurno, equipeApoio)) continue;

        registrarUsoDiaCat_(usadosDiaCat, d, categoria, ferista);
        registrarUsoDiaCatTurno_(usadosDiaCatTurno, d, categoria, ferista, turnoApoio);
        registrarApoioMaps_(ferista, equipeApoio, categoria, d, apoioTeamDayCount, apoioFeristaTeamCount, apoioLastDayFeristaTeam, apoioTeamWeekCount, tz);

        alocacoes.push({
          data: d,
          turno: turnoApoio,
          equipe: equipeApoio,
          categoria,
          peso: pesoEquipe.get(equipeApoio) || 0,
          ferista,
          tipo: 'APOIO',
          origemRow: ''
        });

        alocou = true;
        break;
      }
    }
  });

  // =========================
  // 2) COBERTURA NORMAL
  // =========================
  for (const dem of demandas) {
    const pool = feristasPorCat[dem.categoria] || [];
    if (!pool.length) {
      faltasNaoCobertas.push(dem);
      continue;
    }

    const escolhido = pickBestCandidate_(
      dem,
      pool,
      usadosDiaCat,
      usadosDiaCatTurno,
      teamWeekCount,
      feristaWeekCount,
      feristaTeamCount,
      lastDayFeristaTeam,
      regras,
      tz
    );

    if (!escolhido) {
      faltasNaoCobertas.push(dem);
      continue;
    }

    registrarUsoDiaCat_(usadosDiaCat, dem.data, dem.categoria, escolhido);
    registrarUsoDiaCatTurno_(usadosDiaCatTurno, dem.data, dem.categoria, escolhido, dem.turno);

    incrementarMapa_(teamWeekCount, `${dem.equipe}|${dem.categoria}`);
    incrementarMapa_(feristaWeekCount, `${escolhido}|${dem.categoria}`);
    incrementarMapa_(feristaTeamCount, `${escolhido}|${dem.equipe}|${dem.categoria}`);
    lastDayFeristaTeam.set(`${escolhido}|${dem.equipe}|${dem.categoria}`, keyDia_(dem.data, tz));

    alocacoes.push({
      data: dem.data,
      turno: dem.turno,
      equipe: dem.equipe,
      categoria: dem.categoria,
      peso: dem.peso,
      ferista: escolhido,
      tipo: 'COBERTURA',
      origemRow: dem.origemRow
    });
  }

  // =========================
  // 3) APOIO SOBRANTE
  // =========================
  ['medico', 'enfermeiro'].forEach(categoria => {
    const pool = feristasPorCat[categoria] || [];
    if (!pool.length) return;

    for (const d of diasSemana) {
      for (const ferista of pool) {
        const turnoPreferido = turnoPreferidoParaApoio_(ferista, d, regras);
        const equipeApoio = escolherEquipeApoio_(
          turnoPreferido,
          d,
          categoria,
          ferista,
          equipesOrdenadasPorPeso,
          pesoEquipe,
          turnoEquipe,
          apoioTeamDayCount,
          apoioFeristaTeamCount,
          apoioLastDayFeristaTeam,
          apoioTeamWeekCount,
          regras,
          usadosDiaCat,
          usadosDiaCatTurno,
          tz
        );

        if (!equipeApoio) continue;

        const demCheck = { data: d, equipe: equipeApoio, categoria, turno: turnoPreferido };
        if (!feristaPodeSerAlocado_(ferista, demCheck, regras)) continue;
        if (!podeUsarNoTurno_(ferista, d, categoria, turnoPreferido, regras, usadosDiaCat, usadosDiaCatTurno, equipeApoio)) continue;

        registrarUsoDiaCat_(usadosDiaCat, d, categoria, ferista);
        registrarUsoDiaCatTurno_(usadosDiaCatTurno, d, categoria, ferista, turnoPreferido);
        registrarApoioMaps_(ferista, equipeApoio, categoria, d, apoioTeamDayCount, apoioFeristaTeamCount, apoioLastDayFeristaTeam, apoioTeamWeekCount, tz);

        alocacoes.push({
          data: d,
          turno: turnoPreferido,
          equipe: equipeApoio,
          categoria,
          peso: pesoEquipe.get(equipeApoio) || 0,
          ferista,
          tipo: 'APOIO',
          origemRow: ''
        });
      }
    }
  });

  // =========================
  // SAÍDAS
  // =========================
  escreverAlocacaoAuto_(ss, alocacoes);
  escreverVisualSemanal_(ss, alocacoes, ferVals, diasSemana, tz);
  atualizarVisualEquipes_(alocacoes, inicioSemana);
  atualizarChecks_(regras);
  atualizarPublicarSemana_();
  formatarVisuais_();
  aplicarValidacoes_();
  atualizarDashboard_();
  atualizarHistoricoCobertura_();

  if (avisos.length) {
    SpreadsheetApp.getUi().alert(
      'Atenção: encontrei inconsistências:\n\n' +
      avisos.slice(0, 25).map(x => '- ' + x).join('\n') +
      (avisos.length > 25 ? `\n... (+${avisos.length - 25} avisos)` : '')
    );
  }

  if (faltasNaoCobertas.length) {
    let msg = 'Existem buracos NÃO cobertos:\n';
    faltasNaoCobertas.slice(0, 20).forEach(x => {
      msg += `- ${Utilities.formatDate(x.data, tz, 'dd/MM')} | ${x.equipe} | ${x.categoria} | ${x.turno}\n`;
    });
    if (faltasNaoCobertas.length > 20) msg += `... (+${faltasNaoCobertas.length - 20} linhas)\n`;
    SpreadsheetApp.getUi().alert(msg);
  }
}

/* =========================
   HELPERS PRINCIPAIS
========================= */

function getInicioSemana_(shRegras) {
  const cfgKey = toText_(shRegras.getRange('A1').getValue()).toUpperCase();
  const cfgVal = shRegras.getRange('B1').getValue();
  if (cfgKey !== 'INICIO_SEMANA' || !(cfgVal instanceof Date)) {
    throw new Error('REGRAS_SEMANA precisa ter A1="INICIO_SEMANA" e B1 com a data da segunda-feira.');
  }
  const d = new Date(cfgVal);
  d.setHours(0, 0, 0, 0);
  return d;
}

function lerRegrasSemana_(shRegras) {
  const lastRow = shRegras.getLastRow();
  const vals = lastRow >= 3 ? shRegras.getRange(3, 1, lastRow - 2, 7).getValues() : [];

  const regras = [];

  vals.forEach(r => {
    const ferista = toText_(r[0]);
    const tipo = toText_(r[1]).toUpperCase();
    const dataRaw = r[2];
    const diaSemana = toText_(r[3]).toUpperCase();
    const turno = normTurno_(r[4]);
    const equipe = toText_(r[5]);
    const obs = toText_(r[6]);

    if (!ferista || !tipo) return;
    if (!['PERMITIR_APENAS', 'OBRIGATORIO', 'NAO_ALOCAR', 'DOBRAR'].includes(tipo)) return;

    let data = null;
    if (dataRaw instanceof Date) {
      data = new Date(dataRaw);
      data.setHours(0, 0, 0, 0);
    }

    regras.push({ ferista, tipo, data, diaSemana, turno, equipe, obs });
  });

  return regras;
}

function regraAplicaNoDia_(regra, data) {
  if (regra.data && !sameDate_(regra.data, data)) return false;
  if (regra.diaSemana && regra.diaSemana !== dayCodePT_(data)) return false;
  return true;
}

function feristaPodeSerAlocado_(ferista, dem, regras) {
  // =========================
  // 1) NAO_ALOCAR bloqueia
  // =========================
  const bloqueios = regras.filter(r => r.ferista === ferista && r.tipo === 'NAO_ALOCAR');
  for (const r of bloqueios) {
    if (!regraAplicaNoDia_(r, dem.data)) continue;
    if (r.equipe && dem.equipe && r.equipe !== dem.equipe) continue;
    if (r.turno && dem.turno && r.turno !== dem.turno) continue;
    return false;
  }

  // =========================
  // 2) OBRIGATORIO vira restrição dura no dia
  // =========================
  const obrigatorias = regras.filter(r => r.ferista === ferista && r.tipo === 'OBRIGATORIO');
  const obrigDoDia = obrigatorias.filter(r => regraAplicaNoDia_(r, dem.data));

  if (obrigDoDia.length > 0) {
    // Se houver obrigatoriedade no dia, o ferista só pode ser usado
    // em contextos compatíveis com pelo menos uma dessas regras.
    let compativel = false;

    for (const r of obrigDoDia) {
      const turnoOk = !r.turno || r.turno === dem.turno;
      const equipeOk = !r.equipe || !dem.equipe || r.equipe === dem.equipe;

      if (turnoOk && equipeOk) {
        compativel = true;
        break;
      }
    }

    if (!compativel) return false;
  }

  // =========================
  // 3) PERMITIR_APENAS restringe
  // =========================
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

function feristaPodeDobrar_(ferista, data, turno, equipe, regras) {
  const dobrar = regras.filter(r => r.ferista === ferista && r.tipo === 'DOBRAR');

  for (const r of dobrar) {
    if (!regraAplicaNoDia_(r, data)) continue;
    if (r.equipe && equipe && r.equipe !== equipe) continue;
    if (r.turno && turno && r.turno !== turno) continue;
    return true;
  }
  return false;
}

function podeUsarNoTurno_(ferista, data, categoria, turno, regras, usadosDiaCat, usadosDiaCatTurno, equipe) {
  const usosDia = getUsoDiaCat_(usadosDiaCat, data, categoria, ferista);
  const usosTurno = getUsoDiaCatTurno_(usadosDiaCatTurno, data, categoria, ferista, turno);

  // nunca pode repetir no mesmo turno
  if (usosTurno >= 1) return false;

  // sem uso no dia -> pode
  if (usosDia === 0) return true;

  // um uso no dia -> só pode se houver DOBRAR
  if (usosDia === 1) {
    return feristaPodeDobrar_(ferista, data, turno, equipe, regras);
  }

  // dois ou mais no dia -> nunca
  return false;
}

function pickBestCandidate_(dem, pool, usadosDiaCat, usadosDiaCatTurno, teamWeekCount, feristaWeekCount, feristaTeamCount, lastDayFeristaTeam, regras, tz) {
  const candidates = pool.filter(ferista => {
    if (!feristaPodeSerAlocado_(ferista, dem, regras)) return false;
    if (!podeUsarNoTurno_(ferista, dem.data, dem.categoria, dem.turno, regras, usadosDiaCat, usadosDiaCatTurno, dem.equipe)) return false;
    return true;
  });

  if (!candidates.length) return null;

  const teamKey = `${dem.equipe}|${dem.categoria}`;
  const teamCount = getMapa_(teamWeekCount, teamKey);

  let best = null;
  let bestScore = -1e18;

  const hasFresh = candidates.some(c => getMapa_(feristaTeamCount, `${c}|${dem.equipe}|${dem.categoria}`) === 0);

  for (const cand of candidates) {
    const fCount = getMapa_(feristaWeekCount, `${cand}|${dem.categoria}`);
    const ftCount = getMapa_(feristaTeamCount, `${cand}|${dem.equipe}|${dem.categoria}`);

    let score = 0;
    score += 10 / (1 + teamCount);
    score -= 2 * fCount;

    if (ftCount > 0) score += 6;

    const last = lastDayFeristaTeam.get(`${cand}|${dem.equipe}|${dem.categoria}`);
    if (last && isYesterdayKey_(last, dem.data)) score += 4;

    score -= 4 * ftCount;

    if (ftCount >= 2 && hasFresh) score -= 1e6;

    score += (dem.peso || 0) * 0.0001;

    if (score > bestScore) {
      bestScore = score;
      best = cand;
    }
  }

  return best;
}

function escolherEquipeApoio_(turno, data, categoria, ferista, equipesOrdenadasPorPeso, pesoEquipe, turnoEquipe, apoioTeamDayCount, apoioFeristaTeamCount, apoioLastDayFeristaTeam, apoioTeamWeekCount, regras, usadosDiaCat, usadosDiaCatTurno, tz) {
  let bestEquipe = null;
  let bestScore = -1e18;

  for (const e of equipesOrdenadasPorPeso) {
    const equipe = e.equipe;
    const peso = e.peso || 0;
    const turnoEq = turnoEquipe.get(equipe) || '';
    if (turnoEq && turnoEq !== turno) continue;

    const demCheck = { data, equipe, categoria, turno };
    if (!feristaPodeSerAlocado_(ferista, demCheck, regras)) continue;
    if (!podeUsarNoTurno_(ferista, data, categoria, turno, regras, usadosDiaCat, usadosDiaCatTurno, equipe)) continue;

    const dayCount = getMapa_(apoioTeamDayCount, `${keyDia_(data, tz)}|${categoria}|${equipe}`);
    const teamWeek = getMapa_(apoioTeamWeekCount, `${equipe}|${categoria}`);
    const ftCount = getMapa_(apoioFeristaTeamCount, `${ferista}|${equipe}|${categoria}`);

    let score = 0;

    // peso
    score += peso;

    // não concentrar apoio demais no mesmo dia
    score -= 30 * dayCount;

    // equipes com menos apoio semanal ganham bônus
    score += 20 / (1 + teamWeek);

    // forte continuidade do mesmo ferista no mesmo apoio
    if (ftCount > 0) score += 50 * ftCount;

    const last = apoioLastDayFeristaTeam.get(`${ferista}|${equipe}|${categoria}`);
    if (last && isYesterdayKey_(last, data)) score += 35;

    if (score > bestScore) {
      bestScore = score;
      bestEquipe = equipe;
    }
  }

  return bestEquipe;
}

function registrarApoioMaps_(ferista, equipe, categoria, data, apoioTeamDayCount, apoioFeristaTeamCount, apoioLastDayFeristaTeam, apoioTeamWeekCount, tz) {
  incrementarMapa_(apoioTeamDayCount, `${keyDia_(data, tz)}|${categoria}|${equipe}`);
  incrementarMapa_(apoioFeristaTeamCount, `${ferista}|${equipe}|${categoria}`);
  incrementarMapa_(apoioTeamWeekCount, `${equipe}|${categoria}`);
  apoioLastDayFeristaTeam.set(`${ferista}|${equipe}|${categoria}`, keyDia_(data, tz));
}

function turnoPreferidoParaApoio_(ferista, data, regras) {
  const permitOnly = regras.filter(r => r.ferista === ferista && r.tipo === 'PERMITIR_APENAS');
  const aplicaveis = permitOnly.filter(r => {
    if (!regraAplicaNoDia_(r, data)) return false;
    if (r.equipe) return false;
    return true;
  });

  if (aplicaveis.length > 0) {
    const turnos = new Set(aplicaveis.map(r => r.turno).filter(Boolean));
    if (turnos.has('MANHA')) return 'MANHA';
    if (turnos.has('TARDE')) return 'TARDE';
  }

  return 'MANHA';
}

/* =========================
   SAÍDAS
========================= */

function escreverAlocacaoAuto_(ss, alocacoes) {
  let sh = ss.getSheetByName('ALOCACAO_AUTO');
  if (!sh) sh = ss.insertSheet('ALOCACAO_AUTO');
  sh.clearContents();

  const header = [['Data', 'Turno', 'Equipe', 'Categoria', 'Peso', 'Ferista', 'Tipo', 'Origem(escala_base)']];
  sh.getRange(1, 1, 1, header[0].length).setValues(header);

  alocacoes.sort((a, b) => {
    const da = a.data.getTime(), db = b.data.getTime();
    if (da !== db) return da - db;
    if (a.turno !== b.turno) return a.turno === 'MANHA' ? -1 : 1;
    if (a.tipo !== b.tipo) return a.tipo === 'COBERTURA' ? -1 : 1;
    if (a.equipe !== b.equipe) return a.equipe.localeCompare(b.equipe);
    return a.ferista.localeCompare(b.ferista);
  });

  const vals = alocacoes.map(x => [
    x.data, x.turno, x.equipe, x.categoria, x.peso || 0, x.ferista, x.tipo, x.origemRow || ''
  ]);

  if (vals.length) sh.getRange(2, 1, vals.length, 8).setValues(vals);
}

function escreverVisualSemanal_(ss, alocacoes, ferVals, diasSemana, tz) {
  let sh = ss.getSheetByName('VISUAL_SEMANAL');
  if (!sh) sh = ss.insertSheet('VISUAL_SEMANAL');
  sh.clearContents();

  const nomesDiasPT = ['Dom', 'Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb'];

  const header = ['Ferista', 'Categoria'];
  diasSemana.forEach(d => {
    const nomeDia = nomesDiasPT[d.getDay()];
    const dataTxt = Utilities.formatDate(d, tz, 'dd/MM');
    header.push(`${nomeDia} ${dataTxt}`);
  });
  sh.getRange(1, 1, 1, header.length).setValues([header]);

  const listaFer = [];
  ferVals.forEach(r => {
    const nome = toText_(r[0]);
    const cat = toText_(r[1]).toLowerCase();
    const disp = toText_(r[2]).toUpperCase();
    if (!nome || disp !== 'SIM') return;
    if (cat === 'medico' || cat === 'enfermeiro') listaFer.push({ nome, cat });
  });
  listaFer.sort((a, b) => a.cat === b.cat ? a.nome.localeCompare(b.nome) : a.cat.localeCompare(b.cat));

  const mapFD = new Map();
  alocacoes.forEach(a => {
    const k = `${a.ferista}|${keyDia_(a.data, tz)}`;
    const txt = `${a.tipo}: ${a.equipe} — ${a.turno}`;
    if (!mapFD.has(k)) mapFD.set(k, []);
    mapFD.get(k).push(txt);
  });

  const vals = [];
  listaFer.forEach(f => {
    const row = [f.nome, f.cat];
    diasSemana.forEach(d => {
      const k = `${f.nome}|${keyDia_(d, tz)}`;
      row.push((mapFD.get(k) || ['']).join('\n'));
    });
    vals.push(row);
  });

  if (vals.length) sh.getRange(2, 1, vals.length, header.length).setValues(vals);
}

function atualizarVisualEquipes_(alocacoes, inicioSemana) {
  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone();
  const shEquipes = ss.getSheetByName('cad_equipes');
  if (!shEquipes) throw new Error('VISUAL_EQUIPES: falta cad_equipes.');

  const eqLast = shEquipes.getLastRow();
  const eqVals = eqLast > 1 ? shEquipes.getRange(2, 1, eqLast - 1, 5).getValues() : [];

  const equipes = [];
  const pesoEquipe = new Map();

  eqVals.forEach(r => {
    const equipe = toText_(r[0]);
    const peso = Number(r[4]);
    if (!equipe) return;
    if (!pesoEquipe.has(equipe)) equipes.push(equipe);
    pesoEquipe.set(equipe, isFinite(peso) ? peso : 0);
  });

  equipes.sort((a, b) => (pesoEquipe.get(b) || 0) - (pesoEquipe.get(a) || 0) || a.localeCompare(b));

  const dias = [];
  for (let i = 0; i < 5; i++) {
    const d = new Date(inicioSemana);
    d.setDate(d.getDate() + i);
    d.setHours(0,0,0,0);
    dias.push(d);
  }

  const nomesDiasPT = ['Dom', 'Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb'];
  const header = ['Equipe', 'Peso'];
  dias.forEach(d => {
    const nome = nomesDiasPT[d.getDay()];
    const dataTxt = Utilities.formatDate(d, tz, 'dd/MM');
    header.push(`${nome} ${dataTxt} M`);
    header.push(`${nome} ${dataTxt} T`);
  });

  const map = new Map();
  function mkey(equipe, d, turno) {
    return `${equipe}|${keyDia_(d, tz)}|${turno}`;
  }

  alocacoes.forEach(a => {
    const k = mkey(a.equipe, a.data, a.turno);
    const txt = `${a.tipo}: ${a.ferista}`;
    if (!map.has(k)) map.set(k, []);
    map.get(k).push(txt);
  });

  const vals = [];
  equipes.forEach(eq => {
    const row = [eq, pesoEquipe.get(eq) || 0];
    dias.forEach(d => {
      row.push((map.get(mkey(eq, d, 'MANHA')) || ['']).join('\n'));
      row.push((map.get(mkey(eq, d, 'TARDE')) || ['']).join('\n'));
    });
    vals.push(row);
  });

  let sh = ss.getSheetByName('VISUAL_EQUIPES');
  if (!sh) sh = ss.insertSheet('VISUAL_EQUIPES');
  sh.clearContents();

  sh.getRange(1, 1, 1, header.length).setValues([header]);
  if (vals.length) sh.getRange(2, 1, vals.length, header.length).setValues(vals);
}

function atualizarPublicarSemana_() {
  const ss = SpreadsheetApp.getActive();
  const shSrc = ss.getSheetByName('VISUAL_EQUIPES');
  if (!shSrc) throw new Error('PUBLICAR_SEMANA: falta VISUAL_EQUIPES.');

  let shPub = ss.getSheetByName('PUBLICAR_SEMANA');
  if (!shPub) shPub = ss.insertSheet('PUBLICAR_SEMANA');
  shPub.clearContents();
  shPub.clearFormats();

  const lastRow = shSrc.getLastRow();
  const lastCol = shSrc.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return;

  shPub.getRange(1, 1, lastRow, lastCol).setValues(shSrc.getRange(1, 1, lastRow, lastCol).getValues());
}

/* =========================
   DASHBOARD + HISTÓRICO
========================= */

function atualizarDashboard_() {
  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone();

  const shRegras = ss.getSheetByName('REGRAS_SEMANA');
  const shBase   = ss.getSheetByName('escala_base');
  const shAuto   = ss.getSheetByName('ALOCACAO_AUTO');

  if (!shRegras || !shBase || !shAuto) {
    throw new Error('DASHBOARD: faltam abas REGRAS_SEMANA, escala_base, ALOCACAO_AUTO.');
  }

  const inicioSemana = getInicioSemana_(shRegras);
  const fimSemana = new Date(inicioSemana);
  fimSemana.setDate(fimSemana.getDate() + 5);

  const stats = calcularStatsCobertura_(shBase, shAuto, inicioSemana, fimSemana, tz);

  let sh = ss.getSheetByName('DASHBOARD');
  if (!sh) sh = ss.insertSheet('DASHBOARD');
  sh.clearContents();
  sh.clearFormats();

  sh.getRange('A1').setValue('DASHBOARD — Escala de Feristas').setFontSize(14).setFontWeight('bold');

  const semanaTxt = `${Utilities.formatDate(inicioSemana, tz, 'dd/MM/yyyy')} a ${Utilities.formatDate(new Date(fimSemana.getTime() - 1), tz, 'dd/MM/yyyy')}`;
  sh.getRange('A3').setValue('Semana');
  sh.getRange('B3').setValue(semanaTxt);

  const table = [
    ['Indicador', 'Global', 'Médico', 'Enfermeiro'],
    ['Buracos na semana', stats.global.buracos, stats.medico.buracos, stats.enfermeiro.buracos],
    ['Cobertos', stats.global.cobertos, stats.medico.cobertos, stats.enfermeiro.cobertos],
    ['Não cobertos', stats.global.naoCobertos, stats.medico.naoCobertos, stats.enfermeiro.naoCobertos],
    ['Apoios gerados', stats.global.apoios, stats.medico.apoios, stats.enfermeiro.apoios],
    ['% cobertura', pctTxt_(stats.global.perc), pctTxt_(stats.medico.perc), pctTxt_(stats.enfermeiro.perc)]
  ];

  sh.getRange(5, 1, table.length, table[0].length).setValues(table);
  sh.getRange(5, 1, 1, 4).setFontWeight('bold');
  sh.getRange(5, 1, table.length, 4).setBorder(true, true, true, true, true, true);

  // Formatação correta: números nas contagens, texto na % cobertura
  sh.getRange(6, 2, 4, 3).setNumberFormat('0');
  sh.getRange(10, 2, 1, 3).setNumberFormat('@STRING@');

  sh.getRange(8, 2, 1, 3).setBackgrounds([[
    stats.global.naoCobertos > 0 ? '#f4cccc' : '#d9ead3',
    stats.medico.naoCobertos > 0 ? '#f4cccc' : '#d9ead3',
    stats.enfermeiro.naoCobertos > 0 ? '#f4cccc' : '#d9ead3'
  ]]);

  sh.getRange(10, 2, 1, 3).setBackgrounds([[
    corPerc_(stats.global.perc),
    corPerc_(stats.medico.perc),
    corPerc_(stats.enfermeiro.perc)
  ]]);

  sh.getRange('A13').setValue('Acessos rápidos').setFontWeight('bold');
  const links = [
    ['Publicar (PUBLICAR_SEMANA)', 'PUBLICAR_SEMANA'],
    ['Visão por equipes (VISUAL_EQUIPES)', 'VISUAL_EQUIPES'],
    ['Visão por feristas (VISUAL_SEMANAL)', 'VISUAL_SEMANAL'],
    ['Checagens (CHECKS)', 'CHECKS'],
    ['Alocação detalhada (ALOCACAO_AUTO)', 'ALOCACAO_AUTO'],
    ['Histórico (HISTORICO_COBERTURA)', 'HISTORICO_COBERTURA']
  ];

  sh.getRange(14, 1, links.length, 2).setValues(links);
  sh.getRange(14, 1, links.length, 2).setBorder(true, true, true, true, true, true);

  links.forEach((r, i) => {
    const target = ss.getSheetByName(r[1]);
    if (target) {
      const url = ss.getUrl() + '#gid=' + target.getSheetId();
      sh.getRange(14 + i, 2).setFormula(`=HYPERLINK("${url}";"${r[1]}")`);
    }
  });

  sh.setColumnWidth(1, 220);
  sh.setColumnWidth(2, 140);
  sh.setColumnWidth(3, 140);
  sh.setColumnWidth(4, 140);
}

function atualizarHistoricoCobertura_() {
  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone();

  const shRegras = ss.getSheetByName('REGRAS_SEMANA');
  const shBase   = ss.getSheetByName('escala_base');
  const shAuto   = ss.getSheetByName('ALOCACAO_AUTO');

  const inicioSemana = getInicioSemana_(shRegras);
  const fimSemana = new Date(inicioSemana);
  fimSemana.setDate(fimSemana.getDate() + 5);

  const stats = calcularStatsCobertura_(shBase, shAuto, inicioSemana, fimSemana, tz);

  let sh = ss.getSheetByName('HISTORICO_COBERTURA');
  if (!sh) sh = ss.insertSheet('HISTORICO_COBERTURA');

  const header = [
    'Semana_inicio', 'Semana_fim',
    'Buracos_global', 'Cobertos_global', 'Nao_cobertos_global', 'Apoios_global', 'Perc_global',
    'Buracos_medico', 'Cobertos_medico', 'Nao_cobertos_medico', 'Apoios_medico', 'Perc_medico',
    'Buracos_enfermeiro', 'Cobertos_enfermeiro', 'Nao_cobertos_enfermeiro', 'Apoios_enfermeiro', 'Perc_enfermeiro',
    'Atualizado_em'
  ];

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, header.length).setValues([header]);
  }

  const inicioTxt = Utilities.formatDate(inicioSemana, tz, 'yyyy-MM-dd');
  const fimTxt = Utilities.formatDate(new Date(fimSemana.getTime() - 1), tz, 'yyyy-MM-dd');

  const rowData = [
    inicioTxt, fimTxt,
    stats.global.buracos, stats.global.cobertos, stats.global.naoCobertos, stats.global.apoios, stats.global.perc,
    stats.medico.buracos, stats.medico.cobertos, stats.medico.naoCobertos, stats.medico.apoios, stats.medico.perc,
    stats.enfermeiro.buracos, stats.enfermeiro.cobertos, stats.enfermeiro.naoCobertos, stats.enfermeiro.apoios, stats.enfermeiro.perc,
    new Date()
  ];

  const lastRow = sh.getLastRow();
  let foundRow = 0;
  if (lastRow > 1) {
    const vals = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < vals.length; i++) {
      if (toText_(vals[i][0]) === inicioTxt) {
        foundRow = i + 2;
        break;
      }
    }
  }

  if (foundRow) {
    sh.getRange(foundRow, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sh.getRange(sh.getLastRow() + 1, 1, 1, rowData.length).setValues([rowData]);
  }
}

function calcularStatsCobertura_(shBase, shAuto, inicioSemana, fimSemana, tz) {
  const stats = {
    global: { buracos: 0, cobertos: 0, naoCobertos: 0, apoios: 0, perc: 0 },
    medico: { buracos: 0, cobertos: 0, naoCobertos: 0, apoios: 0, perc: 0 },
    enfermeiro: { buracos: 0, cobertos: 0, naoCobertos: 0, apoios: 0, perc: 0 }
  };

  const buracosSet = { global: new Set(), medico: new Set(), enfermeiro: new Set() };
  const coberturasSet = { global: new Set(), medico: new Set(), enfermeiro: new Set() };

  const baseLast = shBase.getLastRow();
  const baseVals = baseLast > 1 ? shBase.getRange(2, 1, baseLast - 1, 4).getValues() : [];

  baseVals.forEach(r => {
    const dt = r[0];
    const equipe = toText_(r[1]);
    const categoria = toText_(r[2]).toLowerCase();
    const turno = normTurno_(r[3]);

    if (!(dt instanceof Date)) return;
    const data = onlyDate_(dt);
    if (data < inicioSemana || data >= fimSemana) return;
    if (!equipe || !turno) return;
    if (categoria !== 'medico' && categoria !== 'enfermeiro') return;

    const k = `${keyDia_(data, tz)}|${turno}|${equipe}|${categoria}`;
    buracosSet.global.add(k);
    buracosSet[categoria].add(k);
  });

  const autoLast = shAuto.getLastRow();
  const autoVals = autoLast > 1 ? shAuto.getRange(2, 1, autoLast - 1, 8).getValues() : [];

  autoVals.forEach(r => {
    const dt = r[0];
    const turno = normTurno_(r[1]);
    const equipe = toText_(r[2]);
    const categoria = toText_(r[3]).toLowerCase();
    const tipo = toText_(r[6]).toUpperCase();

    if (!(dt instanceof Date)) return;
    const data = onlyDate_(dt);
    if (data < inicioSemana || data >= fimSemana) return;
    if (categoria !== 'medico' && categoria !== 'enfermeiro') return;

    if (tipo === 'COBERTURA') {
      const k = `${keyDia_(data, tz)}|${turno}|${equipe}|${categoria}`;
      coberturasSet.global.add(k);
      coberturasSet[categoria].add(k);
    } else if (tipo === 'APOIO') {
      stats.global.apoios++;
      stats[categoria].apoios++;
    }
  });

  ['global', 'medico', 'enfermeiro'].forEach(k => {
    stats[k].buracos = buracosSet[k].size;
    stats[k].cobertos = coberturasSet[k].size;
    stats[k].naoCobertos = Math.max(0, stats[k].buracos - stats[k].cobertos);
    stats[k].perc = stats[k].buracos > 0 ? stats[k].cobertos / stats[k].buracos : 1;
  });

  return stats;
}

/* =========================
   CHECKS
========================= */

function atualizarChecks_(regras) {
  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone();

  const shEquipes = ss.getSheetByName('cad_equipes');
  const shBase    = ss.getSheetByName('escala_base');
  const shAuto    = ss.getSheetByName('ALOCACAO_AUTO');
  const shRegras  = ss.getSheetByName('REGRAS_SEMANA');

  const inicioSemana = getInicioSemana_(shRegras);
  const fimSemana = new Date(inicioSemana);
  fimSemana.setDate(fimSemana.getDate() + 5);

  const eqLast = shEquipes.getLastRow();
  const eqVals = eqLast > 1 ? shEquipes.getRange(2, 1, eqLast - 1, 2).getValues() : [];
  const turnoEquipe = new Map();
  eqVals.forEach(r => {
    const equipe = toText_(r[0]);
    const turno = normTurno_(r[1]);
    if (equipe) turnoEquipe.set(equipe, turno);
  });

  const autoLast = shAuto.getLastRow();
  const autoVals = autoLast > 1 ? shAuto.getRange(2, 1, autoLast - 1, 8).getValues() : [];

  const turnoIncomp = [];
  const repetidos = [];
  const usoDia = new Map();
  const usoTurno = new Map();
  const coberturasSet = new Set();

  autoVals.forEach((r, i) => {
    const data = r[0];
    const turno = normTurno_(r[1]);
    const equipe = toText_(r[2]);
    const categoria = toText_(r[3]).toLowerCase();
    const ferista = toText_(r[5]);
    const tipo = toText_(r[6]).toUpperCase();

    if (!(data instanceof Date)) return;
    const d0 = onlyDate_(data);
    if (d0 < inicioSemana || d0 >= fimSemana) return;

    const turnoCad = turnoEquipe.get(equipe) || '';
    if (turnoCad && turno && turnoCad !== turno) {
      turnoIncomp.push([d0, equipe, turnoCad, turno, categoria, ferista, tipo, `ALOCACAO_AUTO linha ${i + 2}`]);
    }

    const kDia = `${keyDia_(d0, tz)}|${categoria}|${ferista}`;
    const kTurno = `${keyDia_(d0, tz)}|${categoria}|${ferista}|${turno}`;

    usoDia.set(kDia, (usoDia.get(kDia) || 0) + 1);
    usoTurno.set(kTurno, (usoTurno.get(kTurno) || 0) + 1);

    if ((usoTurno.get(kTurno) || 0) > 1) {
      repetidos.push([d0, categoria, ferista, equipe, turno, tipo, `Mais de uma alocação no mesmo turno (linha ${i + 2})`]);
    }

    if ((usoDia.get(kDia) || 0) > 2) {
      repetidos.push([d0, categoria, ferista, equipe, turno, tipo, `Mais de duas alocações no mesmo dia (linha ${i + 2})`]);
    }

    if ((usoDia.get(kDia) || 0) === 2 && !feristaPodeDobrar_(ferista, d0, turno, equipe, regras)) {
      repetidos.push([d0, categoria, ferista, equipe, turno, tipo, `Dobrou sem regra DOBRAR (linha ${i + 2})`]);
    }

    if (tipo === 'COBERTURA') {
      coberturasSet.add(`${keyDia_(d0, tz)}|${turno}|${equipe}|${categoria}`);
    }
  });

  const baseLast = shBase.getLastRow();
  const baseVals = baseLast > 1 ? shBase.getRange(2, 1, baseLast - 1, 4).getValues() : [];
  const naoCobertos = [];

  baseVals.forEach((r, i) => {
    const data = r[0];
    const equipe = toText_(r[1]);
    const categoria = toText_(r[2]).toLowerCase();
    const turno = normTurno_(r[3]);

    if (!(data instanceof Date)) return;
    const d0 = onlyDate_(data);
    if (d0 < inicioSemana || d0 >= fimSemana) return;

    const k = `${keyDia_(d0, tz)}|${turno}|${equipe}|${categoria}`;
    if (!coberturasSet.has(k)) {
      naoCobertos.push([d0, equipe, categoria, turno, `escala_base linha ${i + 2}`]);
    }
  });

  let sh = ss.getSheetByName('CHECKS');
  if (!sh) sh = ss.insertSheet('CHECKS');
  sh.clearContents();

  let row = 1;

  sh.getRange(row, 1).setValue('1) Turno incompatível com o cadastro da equipe').setFontWeight('bold'); row++;
  sh.getRange(row, 1, 1, 8).setValues([['Data', 'Equipe', 'Turno cadastro', 'Turno alocação', 'Categoria', 'Ferista', 'Tipo', 'Origem']]).setFontWeight('bold'); row++;
  if (turnoIncomp.length) {
    sh.getRange(row, 1, turnoIncomp.length, 8).setValues(turnoIncomp); row += turnoIncomp.length;
  } else {
    sh.getRange(row, 1).setValue('OK — nenhuma divergência encontrada.'); row++;
  }

  row += 2;
  sh.getRange(row, 1).setValue('2) Repetições indevidas').setFontWeight('bold'); row++;
  sh.getRange(row, 1, 1, 7).setValues([['Data', 'Categoria', 'Ferista', 'Equipe', 'Turno', 'Tipo', 'Detalhe']]).setFontWeight('bold'); row++;
  if (repetidos.length) {
    sh.getRange(row, 1, repetidos.length, 7).setValues(repetidos); row += repetidos.length;
  } else {
    sh.getRange(row, 1).setValue('OK — nenhuma repetição indevida encontrada.'); row++;
  }

  row += 2;
  sh.getRange(row, 1).setValue('3) Buracos não cobertos').setFontWeight('bold'); row++;
  sh.getRange(row, 1, 1, 5).setValues([['Data', 'Equipe', 'Categoria', 'Turno', 'Origem']]).setFontWeight('bold'); row++;
  if (naoCobertos.length) {
    sh.getRange(row, 1, naoCobertos.length, 5).setValues(naoCobertos); row += naoCobertos.length;
  } else {
    sh.getRange(row, 1).setValue('OK — todos os buracos foram cobertos.'); row++;
  }
}

/* =========================
   FORMATAÇÃO / VALIDAÇÃO
========================= */

function formatarVisuais_() {
  const ss = SpreadsheetApp.getActive();
  const shEq = ss.getSheetByName('VISUAL_EQUIPES');
  const shFe = ss.getSheetByName('VISUAL_SEMANAL');
  const shCk = ss.getSheetByName('CHECKS');

  if (shEq) aplicarFormatoVisual_(shEq, 2);
  if (shFe) aplicarFormatoVisual_(shFe, 2);
  if (shCk) aplicarFormatoChecks_(shCk);
}

function aplicarFormatoVisual_(sh, freezeCols) {
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 3) return;

  sh.setFrozenRows(1);
  sh.setFrozenColumns(freezeCols);
  sh.getRange(1, 1, lastRow, lastCol).setWrap(true).setVerticalAlignment('MIDDLE');
  sh.getRange(1, 1, 1, lastCol).setFontWeight('bold').setHorizontalAlignment('CENTER');

  const allBands = sh.getBandings();
  allBands.forEach(b => b.remove());
  sh.getRange(1, 1, lastRow, lastCol).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);

  const grid = sh.getRange(2, freezeCols + 1, Math.max(1, lastRow - 1), Math.max(1, lastCol - freezeCols));
  let rules = sh.getConditionalFormatRules() || [];
  rules = rules.filter(r => !r.getRanges().some(rg => intersects_(rg, grid)));

  rules.push(
    SpreadsheetApp.newConditionalFormatRule().whenTextContains('COBERTURA').setBackground('#d9ead3').setRanges([grid]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextContains('APOIO').setBackground('#cfe2f3').setRanges([grid]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=LEN(TRIM(INDIRECT("RC",FALSE)))=0').setBackground('#f4cccc').setRanges([grid]).build()
  );

  sh.setConditionalFormatRules(rules);

  sh.setColumnWidth(1, 220);
  if (freezeCols >= 2) sh.setColumnWidth(2, 90);
  for (let c = freezeCols + 1; c <= lastCol; c++) sh.setColumnWidth(c, 180);
}

function aplicarFormatoChecks_(sh) {
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 1) return;
  sh.getRange(1, 1, lastRow, lastCol).setWrap(true).setVerticalAlignment('MIDDLE');
  sh.autoResizeColumns(1, lastCol);
}

function aplicarValidacoes_() {
  const ss = SpreadsheetApp.getActive();

  const shEquipes = ss.getSheetByName('cad_equipes');
  const shFer     = ss.getSheetByName('cad_feristas');
  const shBase    = ss.getSheetByName('escala_base');
  const shRegras  = ss.getSheetByName('REGRAS_SEMANA');

  if (!shEquipes || !shFer || !shBase || !shRegras) return;

  function setListValidation(range, values, allowBlank) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(allowBlank ? [''].concat(values) : values, true)
      .setAllowInvalid(false)
      .build();
    range.setDataValidation(rule);
  }

  const maxEq = Math.max(1000, shEquipes.getMaxRows());
  const maxFer = Math.max(1000, shFer.getMaxRows());
  const maxBase = Math.max(2000, shBase.getMaxRows());
  const maxReg = Math.max(2000, shRegras.getMaxRows());

  setListValidation(shEquipes.getRange(2, 2, maxEq - 1, 1), ['MANHA', 'TARDE'], false);

  setListValidation(shFer.getRange(2, 2, maxFer - 1, 1), ['medico', 'enfermeiro'], false);
  setListValidation(shFer.getRange(2, 3, maxFer - 1, 1), ['SIM', 'NAO'], false);

  setListValidation(shBase.getRange(2, 3, maxBase - 1, 1), ['medico', 'enfermeiro'], false);
  setListValidation(shBase.getRange(2, 4, maxBase - 1, 1), ['MANHA', 'TARDE'], false);

  setListValidation(shRegras.getRange(3, 2, maxReg - 2, 1), ['PERMITIR_APENAS', 'OBRIGATORIO', 'NAO_ALOCAR', 'DOBRAR'], false);
  setListValidation(shRegras.getRange(3, 4, maxReg - 2, 1), ['SEG', 'TER', 'QUA', 'QUI', 'SEX'], true);
  setListValidation(shRegras.getRange(3, 5, maxReg - 2, 1), ['MANHA', 'TARDE'], true);
}

/* =========================
   UTILITÁRIOS GERAIS
========================= */

function toText_(v) {
  return (v || '').toString().trim();
}

function normTurno_(v) {
  const t = toText_(v).toUpperCase();
  if (t === 'MANHA' || t === 'MANHÃ') return 'MANHA';
  if (t === 'TARDE') return 'TARDE';
  return '';
}

function onlyDate_(d) {
  const x = new Date(d);
  x.setHours(0, 0, 0, 0);
  return x;
}

function sameDate_(a, b) {
  return onlyDate_(a).getTime() === onlyDate_(b).getTime();
}

function keyDia_(d, tz) {
  const tzUse = tz || SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  return Utilities.formatDate(d, tzUse, 'yyyy-MM-dd');
}

function dayCodePT_(d) {
  return ['DOM', 'SEG', 'TER', 'QUA', 'QUI', 'SEX', 'SAB'][d.getDay()] || '';
}

function isYesterdayKey_(yyyyMMdd, d) {
  const parts = yyyyMMdd.split('-').map(Number);
  const last = new Date(parts[0], parts[1] - 1, parts[2]); last.setHours(0,0,0,0);
  const prev = new Date(d); prev.setHours(0,0,0,0); prev.setDate(prev.getDate() - 1);
  return last.getTime() === prev.getTime();
}

function incrementarMapa_(map, key) {
  map.set(key, (map.get(key) || 0) + 1);
}

function getMapa_(map, key) {
  return map.get(key) || 0;
}

function getUsoDiaCat_(map, data, categoria, ferista) {
  return map.get(`${keyDia_(data)}|${categoria}|${ferista}`) || 0;
}

function registrarUsoDiaCat_(map, data, categoria, ferista) {
  const k = `${keyDia_(data)}|${categoria}|${ferista}`;
  map.set(k, (map.get(k) || 0) + 1);
}

function getUsoDiaCatTurno_(map, data, categoria, ferista, turno) {
  return map.get(`${keyDia_(data)}|${categoria}|${ferista}|${turno}`) || 0;
}

function registrarUsoDiaCatTurno_(map, data, categoria, ferista, turno) {
  const k = `${keyDia_(data)}|${categoria}|${ferista}|${turno}`;
  map.set(k, (map.get(k) || 0) + 1);
}

function pctTxt_(x) {
  return Math.round((x || 0) * 100) + '%';
}

function corPerc_(x) {
  if (x >= 0.95) return '#d9ead3';
  if (x >= 0.85) return '#fff2cc';
  return '#f4cccc';
}

function intersects_(r1, r2) {
  const r1r = r1.getRow(), r1c = r1.getColumn(), r1h = r1.getNumRows(), r1w = r1.getNumColumns();
  const r2r = r2.getRow(), r2c = r2.getColumn(), r2h = r2.getNumRows(), r2w = r2.getNumColumns();

  const r1Bottom = r1r + r1h - 1;
  const r2Bottom = r2r + r2h - 1;
  const r1Right = r1c + r1w - 1;
  const r2Right = r2c + r2w - 1;

  const rowsOverlap = !(r1Bottom < r2r || r2Bottom < r1r);
  const colsOverlap = !(r1Right < r2c || r2Right < r1c);
  return rowsOverlap && colsOverlap;
}

/* =========================
   EXPORTAR / LIMPAR
========================= */

function exportarPublicarSemanaPDF() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('PUBLICAR_SEMANA');
  if (!sh) throw new Error('Falta a aba PUBLICAR_SEMANA.');

  const gid = sh.getSheetId();
  const url = ss.getUrl().replace(/edit$/, '');

  const exportUrl =
    url +
    'export?format=pdf' +
    '&gid=' + gid +
    '&portrait=false' +
    '&fitw=true' +
    '&sheetnames=false&printtitle=false&pagenumbers=false' +
    '&gridlines=false&fzr=true';

  const token = ScriptApp.getOAuthToken();
  const resp = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });

  if (resp.getResponseCode() !== 200) {
    throw new Error('Falha ao exportar PDF. Código: ' + resp.getResponseCode());
  }

  const blob = resp.getBlob().setName('PUBLICAR_SEMANA.pdf');
  const file = DriveApp.createFile(blob);
  SpreadsheetApp.getUi().alert('PDF gerado no Drive:\n' + file.getUrl());
}

function limparSemanaSaidas() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  const resp = ui.alert(
    'Limpar saídas da semana',
    'Isso vai apagar o conteúdo das abas de saída. Não apaga cadastros nem escala_base.\n\nDeseja continuar?',
    ui.ButtonSet.YES_NO
  );

  if (resp !== ui.Button.YES) return;

  ['ALOCACAO_AUTO', 'VISUAL_SEMANAL', 'VISUAL_EQUIPES', 'PUBLICAR_SEMANA', 'CHECKS', 'DASHBOARD']
    .forEach(name => {
      const sh = ss.getSheetByName(name);
      if (sh) {
        sh.clearContents();
        sh.clearFormats();
      }
    });

  ui.alert('Saídas limpas.');
}
