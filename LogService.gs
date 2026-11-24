/**
 * ============================================================================
 * LogService.gs - Telemetria, Feedback e Logs com LockService (LGPD)
 * Portal: Manual de Gestão da Tecnologia — SEDUC
 * ============================================================================
 */

function sanitizarEntradaTexto_(valor, maxLen) {
  return String(valor || '')
    .replace(/[<>]/g, '')
    .trim()
    .substring(0, maxLen || 500);
}

/**
 * Registra o acesso a um conteúdo na aba ACESSOS.
 */
function registrarAcesso(dados) {
  if (!dados || !dados.id) return { sucesso: false };

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return { sucesso: false };

  try {
    const ss = getPlanilhaCatalogo_();
    const sheet = ss.getSheetByName(DEFAULTS.ABAS.ACESSOS);
    const email = getUsuarioAtual_();
    // Armazena apenas identificador institucional anonimizado/parcial em conformidade com a LGPD
    const usuarioLgpd = email ? email.replace(/^(.{3}).*(@.*)$/, '$1***$2') : 'Sessão Institucional';

    sheet.appendRow([
      new Date(),
      usuarioLgpd,
      sanitizarEntradaTexto_(dados.id, 40),
      sanitizarEntradaTexto_(dados.titulo, 200),
      sanitizarEntradaTexto_(dados.tema, 100),
      sanitizarEntradaTexto_(dados.tipo, 60),
      sanitizarEntradaTexto_(dados.termoPesquisado, 150),
      sanitizarEntradaTexto_(dados.origem, 80)
    ]);

    return { sucesso: true };
  } catch (e) {
    Logger.log('Erro ao registrar acesso: ' + e.message);
    return { sucesso: false };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Registra termos pesquisados e volume de resultados na aba TERMOS_PESQUISADOS.
 */
function registrarPesquisa(dados) {
  if (!dados) return { sucesso: false };

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return { sucesso: false };

  try {
    const ss = getPlanilhaCatalogo_();
    const sheet = ss.getSheetByName(DEFAULTS.ABAS.TERMOS);
    const filtrosResumo = dados.filtros
      ? Object.entries(dados.filtros)
          .filter(([, v]) => v !== '' && v !== null && v !== undefined && v !== false)
          .map(([k, v]) => k + '=' + v)
          .join('; ')
      : '';

    sheet.appendRow([
      new Date(),
      sanitizarEntradaTexto_(dados.termo, 150),
      Number(dados.quantidade) || 0,
      sanitizarEntradaTexto_(filtrosResumo, 300)
    ]);

    return { sucesso: true };
  } catch (e) {
    Logger.log('Erro ao registrar pesquisa: ' + e.message);
    return { sucesso: false };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Registra avaliação de utilidade ("Sim" ou "Não" + comentário opcional) na aba FEEDBACK.
 */
function registrarFeedback(dados) {
  if (!dados || !dados.id || !dados.resposta) {
    throw new Error('Dados de feedback incompletos.');
  }

  const respostaNorm = normalizarTexto(dados.resposta) === 'sim' ? 'Sim' : 'Não';
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(8000)) {
    throw new Error('Servidor ocupado. Tente novamente em instantes.');
  }

  try {
    const ss = getPlanilhaCatalogo_();
    const sheet = ss.getSheetByName(DEFAULTS.ABAS.FEEDBACK);
    sheet.appendRow([
      new Date(),
      sanitizarEntradaTexto_(dados.id, 40),
      sanitizarEntradaTexto_(dados.titulo, 200),
      respostaNorm,
      sanitizarEntradaTexto_(dados.comentario, 800)
    ]);

    return { sucesso: true };
  } catch (e) {
    registrarLogErro('registrarFeedback', e.message, e.stack);
    throw new Error('Não foi possível registrar seu feedback.');
  } finally {
    lock.releaseLock();
  }
}

function registrarLogErro(origem, mensagem, detalhes) {
  gravarLog_('ERRO', origem, mensagem, detalhes);
}

function registrarLogInfo_(origem, mensagem, detalhes) {
  gravarLog_('INFO', origem, mensagem, detalhes);
}

function gravarLog_(nivel, origem, mensagem, detalhes) {
  try {
    const ss = getPlanilhaCatalogo_();
    const sheet = ss.getSheetByName(DEFAULTS.ABAS.LOGS);
    if (!sheet) return;
    sheet.appendRow([
      new Date(),
      nivel,
      sanitizarEntradaTexto_(origem, 100),
      sanitizarEntradaTexto_(mensagem, 300),
      sanitizarEntradaTexto_(detalhes, 1000)
    ]);
  } catch (e) {
    Logger.log('Falha ao gravar log: ' + e.message);
  }
}