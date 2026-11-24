/**
 * ============================================================================
 * AdminService.gs - Serviços Exclusivos da Área Administrativa
 * Portal: Manual de Gestão da Tecnologia — SEDUC
 * ============================================================================
 */

function validarPermissaoAdmin_() {
  if (!usuarioEhAdministrador()) {
    throw new Error('Acesso negado: privilégios de administrador necessários.');
  }
}

/**
 * Limpa o cache do portal para refletir alterações imediatas feitas na planilha.
 */
function limparCache() {
  validarPermissaoAdmin_();
  const cache = CacheService.getScriptCache();
  cache.removeAll([CACHE_KEY_CONTEUDOS, 'PORTAL_CONFIG_V1']);
  return { sucesso: true, mensagem: 'Cache limpo com sucesso.' };
}

/**
 * Consolida indicadores, métricas de uso e lista completa de conteúdos para o painel Admin.
 */
function getAdminDashboardData() {
  validarPermissaoAdmin_();

  const ss = getPlanilhaCatalogo_();
  const conteudos = carregarTodosRegistros_(true);

  // 1. Termos pesquisados & pesquisas sem resultado
  const sheetTermos = ss.getSheetByName(DEFAULTS.ABAS.TERMOS);
  const contagemTermos = {};
  const pesquisasSemResultado = [];

  if (sheetTermos && sheetTermos.getLastRow() > 1) {
    const dadosTermos = sheetTermos.getRange(2, 1, sheetTermos.getLastRow() - 1, 4).getValues();
    dadosTermos.forEach(([data, termo, qtd, filtros]) => {
      const t = String(termo || '').trim();
      if (!t || t === '[Filtro sem texto]') return;
      const chave = normalizarTexto(t);
      if (!contagemTermos[chave]) {
        contagemTermos[chave] = { termo: t, buscas: 0, ultimoResultado: Number(qtd) || 0 };
      }
      contagemTermos[chave].buscas++;
      contagemTermos[chave].ultimoResultado = Number(qtd) || 0;

      if (Number(qtd) === 0) {
        pesquisasSemResultado.push({
          data: data instanceof Date ? Utilities.formatDate(data, 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm') : String(data),
          termo: t,
          filtros: String(filtros || '')
        });
      }
    });
  }

  const topTermos = Object.values(contagemTermos)
    .sort((a, b) => b.buscas - a.buscas)
    .slice(0, 15);

  // 2. Conteúdos mais acessados
  const sheetAcessos = ss.getSheetByName(DEFAULTS.ABAS.ACESSOS);
  const contagemAcessos = {};
  let totalAcessos = 0;

  if (sheetAcessos && sheetAcessos.getLastRow() > 1) {
    const dadosAcessos = sheetAcessos.getRange(2, 1, sheetAcessos.getLastRow() - 1, 6).getValues();
    totalAcessos = dadosAcessos.length;
    dadosAcessos.forEach(([, , idCont, titulo, tema, tipo]) => {
      const k = String(idCont || titulo);
      if (!k) return;
      if (!contagemAcessos[k]) {
        contagemAcessos[k] = { id: idCont, titulo: titulo, tema: tema, tipo: tipo, acessos: 0 };
      }
      contagemAcessos[k].acessos++;
    });
  }

  const topAcessados = Object.values(contagemAcessos)
    .sort((a, b) => b.acessos - a.acessos)
    .slice(0, 15);

  // 3. Feedbacks recentes
  const sheetFeedback = ss.getSheetByName(DEFAULTS.ABAS.FEEDBACK);
  const feedbacks = [];
  if (sheetFeedback && sheetFeedback.getLastRow() > 1) {
    const dadosFb = sheetFeedback.getRange(2, 1, sheetFeedback.getLastRow() - 1, 5).getValues();
    dadosFb.slice(-20).reverse().forEach(([data, idCont, titulo, resp, coment]) => {
      feedbacks.push({
        data: data instanceof Date ? Utilities.formatDate(data, 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm') : String(data),
        id: idCont,
        titulo: titulo,
        resposta: resp,
        comentario: coment
      });
    });
  }

  // 4. Logs recentes
  const sheetLogs = ss.getSheetByName(DEFAULTS.ABAS.LOGS);
  const logs = [];
  if (sheetLogs && sheetLogs.getLastRow() > 1) {
    const dadosLogs = sheetLogs.getRange(2, 1, sheetLogs.getLastRow() - 1, 5).getValues();
    dadosLogs.slice(-25).reverse().forEach(([data, nivel, origem, msg, det]) => {
      logs.push({
        data: data instanceof Date ? Utilities.formatDate(data, 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm') : String(data),
        nivel: nivel,
        origem: origem,
        mensagem: msg,
        detalhes: det
      });
    });
  }

  return {
    resumo: {
      totalConteudos: conteudos.length,
      publicados: conteudos.filter(c => normalizarTexto(c.status) === 'publicado').length,
      arquivados: conteudos.filter(c => normalizarTexto(c.status) === 'arquivado').length,
      totalAcessos: totalAcessos,
      totalSemResultado: pesquisasSemResultado.length
    },
    conteudos: conteudos.map(sanitizarItemParaFrontend_),
    topTermos: topTermos,
    pesquisasSemResultado: pesquisasSemResultado.slice(-20).reverse(),
    topAcessados: topAcessados,
    feedbacks: feedbacks,
    logs: logs
  };
}

/**
 * Cadastra um novo conteúdo ou atualiza um registro existente na aba CONTEUDOS.
 */
function salvarConteudoAdmin(dados) {
  validarPermissaoAdmin_();
  if (!dados || !dados.titulo) {
    throw new Error('O título do conteúdo é obrigatório.');
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    throw new Error('Planilha ocupada. Tente novamente.');
  }

  try {
    const ss = getPlanilhaCatalogo_();
    const sheet = ss.getSheetByName(DEFAULTS.ABAS.CONTEUDOS);
    const lastRow = sheet.getLastRow();
    const tz = Session.getScriptTimeZone() || 'America/Sao_Paulo';
    const hoje = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy');

    const fileId = sanitizarEntradaTexto_(dados.fileId, 120);
    const urlFinal = sanitizarEntradaTexto_(dados.url, 500) ||
      (fileId ? 'https://drive.google.com/file/d/' + fileId + '/view' : '');

    if (dados.id) {
      // Edição de registro existente
      const ids = sheet.getRange(2, 1, Math.max(1, lastRow - 1), 1).getValues();
      let rowTarget = -1;
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0]).trim() === String(dados.id).trim()) {
          rowTarget = i + 2;
          break;
        }
      }
      if (rowTarget === -1) throw new Error('Registro não encontrado para edição.');

      const linhaAtual = sheet.getRange(rowTarget, 1, 1, 21).getValues()[0];
      linhaAtual[1] = sanitizarEntradaTexto_(dados.titulo, 250);
      linhaAtual[2] = sanitizarEntradaTexto_(dados.resumo, 1000);
      linhaAtual[3] = sanitizarEntradaTexto_(dados.tema, 120);
      linhaAtual[4] = sanitizarEntradaTexto_(dados.subtema, 120);
      linhaAtual[5] = sanitizarEntradaTexto_(dados.publico, 200);
      linhaAtual[6] = sanitizarEntradaTexto_(dados.tipo, 80);
      linhaAtual[7] = sanitizarEntradaTexto_(dados.palavrasChave, 400);
      linhaAtual[8] = fileId;
      linhaAtual[9] = urlFinal;
      linhaAtual[10] = sanitizarEntradaTexto_(dados.status || 'Publicado', 40);
      linhaAtual[11] = sanitizarEntradaTexto_(dados.versao || '1.0', 20);
      linhaAtual[12] = hoje;
      linhaAtual[13] = sanitizarEntradaTexto_(dados.responsavel || 'COEGD', 120);
      linhaAtual[14] = dados.destaque ? 'Sim' : 'Não';
      linhaAtual[15] = Number(dados.ordem) || linhaAtual[15] || 100;
      linhaAtual[17] = sanitizarEntradaTexto_(dados.relacionadosIds, 200);
      linhaAtual[20] = sanitizarEntradaTexto_(dados.objetivo, 1000);

      sheet.getRange(rowTarget, 1, 1, 21).setValues([linhaAtual]);
    } else {
      // Novo cadastro manual
      let maiorSeq = 0;
      if (lastRow > 1) {
        const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        ids.forEach(([val]) => {
          const m = String(val).match(/(\d+)$/);
          if (m && parseInt(m[1], 10) > maiorSeq) maiorSeq = parseInt(m[1], 10);
        });
      }
      const novoId = 'MG-' + String(maiorSeq + 1).padStart(4, '0');

      sheet.appendRow([
        novoId,
        sanitizarEntradaTexto_(dados.titulo, 250),
        sanitizarEntradaTexto_(dados.resumo, 1000),
        sanitizarEntradaTexto_(dados.tema || 'Orientações técnicas', 120),
        sanitizarEntradaTexto_(dados.subtema, 120),
        sanitizarEntradaTexto_(dados.publico || 'SEINTEC e SETEC', 200),
        sanitizarEntradaTexto_(dados.tipo || 'Orientação', 80),
        sanitizarEntradaTexto_(dados.palavrasChave, 400),
        fileId,
        urlFinal,
        sanitizarEntradaTexto_(dados.status || 'Publicado', 40),
        sanitizarEntradaTexto_(dados.versao || '1.0', 20),
        hoje,
        sanitizarEntradaTexto_(dados.responsavel || 'COEGD', 120),
        dados.destaque ? 'Sim' : 'Não',
        Number(dados.ordem) || maiorSeq + 1,
        '',
        sanitizarEntradaTexto_(dados.relacionadosIds, 200),
        'Cadastro manual via painel administrativo',
        '',
        sanitizarEntradaTexto_(dados.objetivo, 1000)
      ]);
    }

    limparCache();
    return { sucesso: true };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Altera rapidamente o status de publicação (Publicado / Arquivado).
 */
function alterarStatusConteudo(id, novoStatus) {
  validarPermissaoAdmin_();
  const statusValidos = ['Publicado', 'Revisão', 'Arquivado'];
  if (!statusValidos.includes(novoStatus)) throw new Error('Status inválido.');

  const ss = getPlanilhaCatalogo_();
  const sheet = ss.getSheetByName(DEFAULTS.ABAS.CONTEUDOS);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { sucesso: false };

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === String(id).trim()) {
      sheet.getRange(i + 2, 11).setValue(novoStatus);
      limparCache();
      return { sucesso: true };
    }
  }
  return { sucesso: false };
}

/**
 * Alterna o indicador de Destaque (Sim / Não).
 */
function alternarDestaqueConteudo(id) {
  validarPermissaoAdmin_();
  const ss = getPlanilhaCatalogo_();
  const sheet = ss.getSheetByName(DEFAULTS.ABAS.CONTEUDOS);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { sucesso: false };

  const dados = sheet.getRange(2, 1, lastRow - 1, 15).getValues();
  for (let i = 0; i < dados.length; i++) {
    if (String(dados[i][0]).trim() === String(id).trim()) {
      const atual = normalizarTexto(dados[i][14]) === 'sim';
      sheet.getRange(i + 2, 15).setValue(atual ? 'Não' : 'Sim');
      limparCache();
      return { sucesso: true };
    }
  }
  return { sucesso: false };
}

/**
 * Gera string CSV consolidada com métricas de uso e catálogo para download administrativo.
 */
function exportarRelatorioUsoCsv() {
  validarPermissaoAdmin_();
  const dash = getAdminDashboardData();
  const linhas = [
    ['RELATÓRIO DE USO — MANUAL DE GESTÃO DA TECNOLOGIA SEDUC'],
    ['Gerado em', Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm')],
    ['Total de Conteúdos', dash.resumo.totalConteudos],
    ['Conteúdos Publicados', dash.resumo.publicados],
    ['Total de Acessos Registrados', dash.resumo.totalAcessos],
    [],
    ['CONTEÚDOS MAIS ACESSADOS'],
    ['ID', 'Título', 'Tema', 'Tipo', 'Acessos']
  ];

  dash.topAcessados.forEach(a => {
    linhas.push([a.id, a.titulo, a.tema, a.tipo, a.acessos]);
  });

  linhas.push([], ['TERMOS MAIS PESQUISADOS'], ['Termo', 'Total de Buscas', 'Última Qtd. Resultados']);
  dash.topTermos.forEach(t => {
    linhas.push([t.termo, t.buscas, t.ultimoResultado]);
  });

  return linhas
    .map(cols => cols.map(c => '"' + String(c === undefined ? '' : c).replace(/"/g, '""') + '"').join(';'))
    .join('\n');
}