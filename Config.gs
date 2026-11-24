/**
 * ============================================================================
 * Config.gs - Configurações Centrais, Inicialização e Controle de Acesso
 * Portal: Manual de Gestão da Tecnologia — SEDUC
 * ============================================================================
 */

const DEFAULTS = {
  DRIVE_FOLDER_ID: '1L2m2dJBIWaFYPH7qaokHDULBP_RQsmiV',
  NOME_PLANILHA: 'CATALOGO_MANUAL_GESTAO',
  NOME_PORTAL: 'Manual de Gestão da Tecnologia — SEDUC',
  SUBTITULO_PORTAL: 'Orientações, procedimentos, fluxos e ferramentas para a gestão da tecnologia nas Unidades Regionais de Ensino e escolas',
  AREA_RESPONSAVEL: 'COEGD — Coordenadoria de Gestão Digital / SEDUC-SP',
  VERSAO_PORTAL: '1.0.0',
  DOMINIO_INSTITUCIONAL: 'educacao.sp.gov.br',
  ADMINS_INICIAIS: 'ivan.brandao@educacao.sp.gov.br, robson.sakata@educacao.sp.gov.br',
  EMAIL_CONTATO: 'coegd.manual@educacao.sp.gov.br',
  LOGO_URL: '', // Opcional: URL pública ou Drive thumbnail; vazio usa o emblema vetorial oficial COEGD
  ABAS: {
    CONTEUDOS: 'CONTEUDOS',
    CONFIGURACOES: 'CONFIGURACOES',
    FAQ: 'FAQ',
    ACESSOS: 'ACESSOS',
    TERMOS: 'TERMOS_PESQUISADOS',
    FEEDBACK: 'FEEDBACK',
    LOGS: 'LOGS'
  }
};

/**
 * Executa a configuração inicial do ambiente:
 * 1. Cria a planilha CATALOGO_MANUAL_GESTAO na pasta principal (ou no Drive raiz).
 * 2. Cria todas as abas e cabeçalhos necessários.
 * 3. Grava os IDs essenciais em PropertiesService.getScriptProperties().
 * Execute esta função UMA VEZ manualmente pelo editor do Apps Script.
 */
function inicializarSistema() {
  const props = PropertiesService.getScriptProperties();
  let spreadsheetId = props.getProperty('SPREADSHEET_ID');
  let ss;

  if (spreadsheetId) {
    try {
      ss = SpreadsheetApp.openById(spreadsheetId);
    } catch (e) {
      ss = null;
    }
  }

  if (!ss) {
    ss = SpreadsheetApp.create(DEFAULTS.NOME_PLANILHA);
    spreadsheetId = ss.getId();
    props.setProperty('SPREADSHEET_ID', spreadsheetId);
  }

  props.setProperty('DRIVE_FOLDER_ID', DEFAULTS.DRIVE_FOLDER_ID);
  props.setProperty('DOMINIO_INSTITUCIONAL', DEFAULTS.DOMINIO_INSTITUCIONAL);

  configurarAbaConteudos_(ss);
  configurarAbaConfiguracoes_(ss);
  configurarAbaFaq_(ss);
  configurarAbaLogsMetricas_(ss);

  // Remove a aba padrão "Página1" ou "Sheet1" se ainda existir
  ['Página1', 'Sheet1'].forEach(nomePadrao => {
    const abaPadrao = ss.getSheetByName(nomePadrao);
    if (abaPadrao && ss.getSheets().length > 1) {
      ss.deleteSheet(abaPadrao);
    }
  });

  Logger.log('Sistema inicializado com sucesso!');
  Logger.log('ID da Planilha CATALOGO_MANUAL_GESTAO: ' + spreadsheetId);
  Logger.log('URL da Planilha: ' + ss.getUrl());

  return {
    sucesso: true,
    spreadsheetId: spreadsheetId,
    url: ss.getUrl()
  };
}

/**
 * Cria e formata a aba CONTEUDOS com as 21 colunas (A:U).
 */
function configurarAbaConteudos_(ss) {
  let sheet = ss.getSheetByName(DEFAULTS.ABAS.CONTEUDOS);
  if (!sheet) {
    sheet = ss.insertSheet(DEFAULTS.ABAS.CONTEUDOS);
  }

  const cabecalhos = [
    'ID', 'Título', 'Resumo', 'Tema', 'Subtema', 'Público', 'Tipo',
    'Palavras-chave', 'ID do arquivo', 'URL', 'Status', 'Versão',
    'Data da atualização', 'Responsável', 'Destaque', 'Ordem',
    'Texto indexado', 'Relacionados', 'Observações', 'Imagem', 'Objetivo'
  ];

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, cabecalhos.length).setValues([cabecalhos]);
    formatarCabecalho_(sheet, cabecalhos.length);
    sheet.setFrozenRows(1);

    // Validação de dados para Status (Coluna K) e Destaque (Coluna O)
    const ruleStatus = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Publicado', 'Revisão', 'Arquivado'], true)
      .build();
    sheet.getRange('K2:K1000').setDataValidation(ruleStatus);

    const ruleDestaque = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Sim', 'Não'], true)
      .build();
    sheet.getRange('O2:O1000').setDataValidation(ruleDestaque);
  }
}

/**
 * Cria e popula a aba CONFIGURACOES.
 */
function configurarAbaConfiguracoes_(ss) {
  let sheet = ss.getSheetByName(DEFAULTS.ABAS.CONFIGURACOES);
  if (!sheet) {
    sheet = ss.insertSheet(DEFAULTS.ABAS.CONFIGURACOES);
  }

  if (sheet.getLastRow() === 0) {
    const linhas = [
      ['Chave', 'Valor', 'Descrição'],
      ['NOME_PORTAL', DEFAULTS.NOME_PORTAL, 'Título principal exibido no cabeçalho e na página inicial'],
      ['SUBTITULO_PORTAL', DEFAULTS.SUBTITULO_PORTAL, 'Subtítulo institucional da área de destaque'],
      ['AREA_RESPONSAVEL', DEFAULTS.AREA_RESPONSAVEL, 'Unidade responsável pelo portal'],
      ['VERSAO_PORTAL', DEFAULTS.VERSAO_PORTAL, 'Versão atual da aplicação web'],
      ['ADMINISTRADORES', DEFAULTS.ADMINS_INICIAIS, 'E-mails institucionais com permissão de administrador (separados por vírgula)'],
      ['DOMINIO_PERMITIDO', DEFAULTS.DOMINIO_INSTITUCIONAL, 'Domínio institucional autorizado a acessar o portal'],
      ['EMAIL_CONTATO', DEFAULTS.EMAIL_CONTATO, 'E-mail de suporte exibido na seção Contato'],
      ['LOGO_URL', DEFAULTS.LOGO_URL, 'URL opcional da imagem do logotipo oficial (deixe vazio para usar o emblema padrão COEGD)'],
      ['PERMITIR_COPIA_DOWNLOAD', 'Sim', 'Exibe botões diretos para cópia e download dos materiais originais']
    ];
    sheet.getRange(1, 1, linhas.length, 3).setValues(linhas);
    formatarCabecalho_(sheet, 3);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, 3);
  }
}

/**
 * Cria e popula a aba FAQ com perguntas iniciais editáveis.
 */
function configurarAbaFaq_(ss) {
  let sheet = ss.getSheetByName(DEFAULTS.ABAS.FAQ);
  if (!sheet) {
    sheet = ss.insertSheet(DEFAULTS.ABAS.FAQ);
  }

  if (sheet.getLastRow() === 0) {
    const linhas = [
      ['ID', 'Categoria', 'Pergunta', 'Resposta', 'Ordem', 'Status'],
      ['FAQ-01', 'Acesso e Navegação', 'Como encontro rapidamente um modelo ou checklist?', 'Utilize a barra de busca central digitando palavras-chave como "checklist", "ata" ou "visita", ou acesse diretamente a aba "Modelos e ferramentas" no menu superior.', 1, 'Publicado'],
      ['FAQ-02', 'Documentos', 'Posso fazer uma cópia ou baixar os modelos disponíveis?', 'Sim. Ao abrir qualquer documento ou clicar em "Ver detalhes", você encontrará opções para abrir o arquivo original no Google Drive, fazer uma cópia para sua conta ou baixar o arquivo.', 2, 'Publicado'],
      ['FAQ-03', 'Atualização', 'Como saber se um procedimento está atualizado?', 'Cada cartão e janela de detalhes exibe a versão vigente e a data da última atualização sincronizada com o repositório oficial da COEGD.', 3, 'Publicado'],
      ['FAQ-04', 'Suporte', 'Não encontrei uma orientação específica. Como proceder?', 'Utilize a seção "Contato" no menu principal para encaminhar sua dúvida ou sugestão de novo material ao Grupo de Trabalho do Manual de Gestão.', 4, 'Publicado']
    ];
    sheet.getRange(1, 1, linhas.length, 6).setValues(linhas);
    formatarCabecalho_(sheet, 6);
    sheet.setFrozenRows(1);
  }
}

/**
 * Cria as abas de telemetria (ACESSOS, TERMOS_PESQUISADOS, FEEDBACK) e LOGS.
 */
function configurarAbaLogsMetricas_(ss) {
  const estruturas = [
    {
      nome: DEFAULTS.ABAS.ACESSOS,
      cols: ['Data e hora', 'Usuário', 'ID Conteúdo', 'Título', 'Tema', 'Tipo', 'Termo pesquisado', 'Origem do acesso']
    },
    {
      nome: DEFAULTS.ABAS.TERMOS,
      cols: ['Data e hora', 'Termo', 'Quantidade de resultados', 'Filtros utilizados']
    },
    {
      nome: DEFAULTS.ABAS.FEEDBACK,
      cols: ['Data e hora', 'ID do conteúdo', 'Título', 'Resposta', 'Comentário']
    },
    {
      nome: DEFAULTS.ABAS.LOGS,
      cols: ['Data e hora', 'Nível', 'Origem', 'Mensagem', 'Detalhes']
    }
  ];

  estruturas.forEach(item => {
    let sheet = ss.getSheetByName(item.nome);
    if (!sheet) {
      sheet = ss.insertSheet(item.nome);
    }
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, item.cols.length).setValues([item.cols]);
      formatarCabecalho_(sheet, item.cols.length);
      sheet.setFrozenRows(1);
    }
  });
}

function formatarCabecalho_(sheet, numCols) {
  sheet.getRange(1, 1, 1, numCols)
    .setBackground('#003366')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
}

/**
 * Retorna a instância da planilha de catálogo.
 */
function getPlanilhaCatalogo_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SPREADSHEET_ID');
  if (!id) {
    throw new Error('SPREADSHEET_ID não configurado. Execute inicializarSistema() primeiro.');
  }
  return SpreadsheetApp.openById(id);
}

/**
 * Obtém as configurações do portal combinando ScriptProperties e a aba CONFIGURACOES.
 */
function getConfiguracoes() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('PORTAL_CONFIG_V1');
  if (cached) {
    return JSON.parse(cached);
  }

  const config = {
    nomePortal: DEFAULTS.NOME_PORTAL,
    subtituloPortal: DEFAULTS.SUBTITULO_PORTAL,
    areaResponsavel: DEFAULTS.AREA_RESPONSAVEL,
    versaoPortal: DEFAULTS.VERSAO_PORTAL,
    administradores: DEFAULTS.ADMINS_INICIAIS.split(',').map(e => e.trim().toLowerCase()),
    dominioPermitido: DEFAULTS.DOMINIO_INSTITUCIONAL,
    emailContato: DEFAULTS.EMAIL_CONTATO,
    logoUrl: DEFAULTS.LOGO_URL,
    permitirCopiaDownload: true
  };

  try {
    const ss = getPlanilhaCatalogo_();
    const sheet = ss.getSheetByName(DEFAULTS.ABAS.CONFIGURACOES);
    if (sheet && sheet.getLastRow() > 1) {
      const dados = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
      const mapa = {};
      dados.forEach(([chave, valor]) => {
        if (chave) mapa[String(chave).trim()] = String(valor || '').trim();
      });

      if (mapa.NOME_PORTAL) config.nomePortal = mapa.NOME_PORTAL;
      if (mapa.SUBTITULO_PORTAL) config.subtituloPortal = mapa.SUBTITULO_PORTAL;
      if (mapa.AREA_RESPONSAVEL) config.areaResponsavel = mapa.AREA_RESPONSAVEL;
      if (mapa.VERSAO_PORTAL) config.versaoPortal = mapa.VERSAO_PORTAL;
      if (mapa.ADMINISTRADORES) {
        config.administradores = mapa.ADMINISTRADORES
          .split(',')
          .map(e => e.trim().toLowerCase())
          .filter(Boolean);
      }
      if (mapa.DOMINIO_PERMITIDO) config.dominioPermitido = mapa.DOMINIO_PERMITIDO.toLowerCase();
      if (mapa.EMAIL_CONTATO) config.emailContato = mapa.EMAIL_CONTATO;
      if (mapa.LOGO_URL !== undefined) config.logoUrl = mapa.LOGO_URL;
      if (mapa.PERMITIR_COPIA_DOWNLOAD) {
        config.permitirCopiaDownload = normalizarTexto(mapa.PERMITIR_COPIA_DOWNLOAD) === 'sim';
      }
    }
  } catch (e) {
    Logger.log('Aviso ao carregar CONFIGURACOES: ' + e.message);
  }

  cache.put('PORTAL_CONFIG_V1', JSON.stringify(config), 600);
  return config;
}

/**
 * Obtém o e-mail do usuário autenticado na sessão Google Workspace.
 */
function getUsuarioAtual_() {
  try {
    const email = Session.getActiveUser().getEmail();
    return email ? email.trim().toLowerCase() : '';
  } catch (e) {
    return '';
  }
}

/**
 * Verifica se o usuário pertence ao domínio institucional autorizado (@educacao.sp.gov.br).
 */
function usuarioAutorizadoNoDominio_() {
  const config = getConfiguracoes();
  const email = getUsuarioAtual_();
  if (!email) return true;
  if (config.administradores.includes(email)) return true;

  const dominiosAceitos = config.dominioPermitido
    .split(',')
    .map(d => d.trim().toLowerCase())
    .filter(Boolean);

  return dominiosAceitos.some(dom =>
    email.endsWith('@' + dom) || email.endsWith('.' + dom)
  );
}

/**
 * Verifica se o usuário atual consta na lista de administradores.
 */
function usuarioEhAdministrador() {
  const email = getUsuarioAtual_();
  if (!email) return false;
  const config = getConfiguracoes();
  return config.administradores.includes(email);
}