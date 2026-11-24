/**
 * ============================================================================
 * Code.gs - Ponto de Entrada Web App e Roteamento
 * Portal: Manual de Gestão da Tecnologia — SEDUC
 * ============================================================================
 */

/**
 * Renderiza a aplicação web principal.
 */
function doGet(e) {
  if (!usuarioAutorizadoNoDominio_()) {
    return HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"><title>Acesso Restrito</title>' +
      '<style>body{font-family:Arial,sans-serif;background:#F5F7FA;color:#37474F;display:flex;align-items:center;justify-content:center;height:100vh;margin:0;}' +
      '.box{background:#fff;padding:32px;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,.08);max-width:480px;text-align:center;border-top:4px solid #C62828;}' +
      'h1{color:#003366;font-size:20px;margin-top:0;}</style></head>' +
      '<body><div class="box"><h1>Acesso Institucional Restrito</h1>' +
      '<p>Este portal é exclusivo para profissionais autenticados com conta institucional <strong>@educacao.sp.gov.br</strong>.</p>' +
      '<p>Verifique a conta Google conectada no seu navegador e tente novamente.</p></div></body></html>'
    ).setTitle('Acesso Restrito — Manual de Gestão SEDUC');
  }

  const config = getConfiguracoes();
  const template = HtmlService.createTemplateFromFile('Index');
  template.config = config;
  template.isAdmin = usuarioEhAdministrador();
  template.userEmail = getUsuarioAtual_();

  return template.evaluate()
    .setTitle(config.nomePortal)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * Permite incluir arquivos HTML parciais (Styles, Scripts, Components, Admin).
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Carrega o pacote inicial de dados em uma única chamada assíncrona
 * para otimizar o tempo de carregamento da interface.
 */
function getInitialAppData() {
  try {
    const config = getConfiguracoes();
    const isAdmin = usuarioEhAdministrador();
    const temas = getTemas();
    const destaques = getDestaques();
    const faqs = getFaqs();
    const inicial = pesquisarConteudos('', {}, 1);

    return {
      sucesso: true,
      config: config,
      isAdmin: isAdmin,
      userEmail: getUsuarioAtual_(),
      temas: temas,
      destaques: destaques,
      faqs: faqs,
      resultadosIniciais: inicial
    };
  } catch (erro) {
    registrarLogErro('getInitialAppData', erro.message, erro.stack);
    return {
      sucesso: false,
      mensagem: 'Não foi possível carregar os dados iniciais do portal.'
    };
  }
}