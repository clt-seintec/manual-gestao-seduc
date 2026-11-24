/**
 * ============================================================================
 * SearchService.gs - Motor de Busca Ponderada, Filtros, Cache e Relacionados
 * Portal: Manual de Gestão da Tecnologia — SEDUC
 * ============================================================================
 */

const CACHE_KEY_CONTEUDOS = 'CATALOGO_CONTEUDOS_PUBLICADOS_V1';
const CACHE_TTL_SEGUNDOS = 600; // 10 minutos
const RESULTADOS_POR_PAGINA = 12;

/**
 * Normaliza texto removendo acentos, convertendo para minúsculas e removendo espaços extras.
 */
function normalizarTexto(texto) {
  return String(texto || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .trim();
}

/**
 * Carrega todos os conteúdos da aba CONTEUDOS (utilizando CacheService para registros publicados).
 */
function carregarTodosRegistros_(incluirNaoPublicados) {
  const cache = CacheService.getScriptCache();

  if (!incluirNaoPublicados) {
    const cached = cache.get(CACHE_KEY_CONTEUDOS);
    if (cached) {
      return JSON.parse(cached);
    }
  }

  const ss = getPlanilhaCatalogo_();
  const sheet = ss.getSheetByName(DEFAULTS.ABAS.CONTEUDOS);
  if (!sheet || sheet.getLastRow() <= 1) {
    return [];
  }

  const tz = Session.getScriptTimeZone() || 'America/Sao_Paulo';
  const dados = sheet.getRange(2, 1, sheet.getLastRow() - 1, 21).getValues();
  const itens = [];

  dados.forEach(row => {
    const id = String(row[0] || '').trim();
    const titulo = String(row[1] || '').trim();
    if (!id || !titulo) return;

    const status = String(row[10] || 'Publicado').trim();
    if (!incluirNaoPublicados && normalizarTexto(status) !== 'publicado') {
      return;
    }

    let dataAtualizacao = '';
    let ano = '';
    if (row[12] instanceof Date) {
      dataAtualizacao = Utilities.formatDate(row[12], tz, 'dd/MM/yyyy');
      ano = Utilities.formatDate(row[12], tz, 'yyyy');
    } else {
      dataAtualizacao = String(row[12] || '').trim();
      const matchAno = dataAtualizacao.match(/\b(20\d{2})\b/);
      ano = matchAno ? matchAno[1] : '';
    }

    const fileId = String(row[8] || '').trim();
    const urlOriginal = String(row[9] || '').trim();
    const links = gerarLinksAcessoDrive_(fileId, urlOriginal);

    itens.push({
      id: id,
      titulo: titulo,
      resumo: String(row[2] || '').trim(),
      tema: String(row[3] || 'Geral').trim(),
      subtema: String(row[4] || '').trim(),
      publico: String(row[5] || '').trim(),
      tipo: String(row[6] || 'Documento').trim(),
      palavrasChave: String(row[7] || '').trim(),
      fileId: fileId,
      url: links.urlVisualizar,
      urlCopiar: links.urlCopiar,
      urlDownload: links.urlDownload,
      status: status,
      versao: String(row[11] || '1.0').trim(),
      dataAtualizacao: dataAtualizacao,
      ano: ano,
      responsavel: String(row[13] || '').trim(),
      destaque: normalizarTexto(row[14]) === 'sim',
      ordem: Number(row[15]) || 9999,
      textoIndexado: String(row[16] || '').trim(),
      relacionadosIds: String(row[17] || '').trim(),
      observacoes: String(row[18] || '').trim(),
      imagem: String(row[19] || '').trim(),
      objetivo: String(row[20] || '').trim()
    });
  });

  // Ordena por Destaque (Sim primeiro), Ordem crescente e Título
  itens.sort((a, b) => {
    if (a.destaque !== b.destaque) return a.destaque ? -1 : 1;
    if (a.ordem !== b.ordem) return a.ordem - b.ordem;
    return a.titulo.localeCompare(b.titulo, 'pt-BR');
  });

  if (!incluirNaoPublicados) {
    try {
      // CacheService limita entradas individuais a 100KB; salva versão otimizada sem textoIndexado gigante se necessário
      const payload = JSON.stringify(itens);
      if (payload.length < 95000) {
        cache.put(CACHE_KEY_CONTEUDOS, payload, CACHE_TTL_SEGUNDOS);
      }
    } catch (e) {
      Logger.log('Aviso de tamanho de cache: ' + e.message);
    }
  }

  return itens;
}

/**
 * Calcula a pontuação de relevância de um item para os termos pesquisados.
 * Pesos: Título (10), Palavras-chave (8), Tema (6), Subtema (5),
 * Resumo (4), Texto indexado (2), Público (2), Tipo (1).
 */
function calcularRelevancia(item, termos) {
  if (!termos || termos.length === 0) return 1;

  const campos = {
    titulo: normalizarTexto(item.titulo),
    palavrasChave: normalizarTexto(item.palavrasChave),
    tema: normalizarTexto(item.tema),
    subtema: normalizarTexto(item.subtema),
    resumo: normalizarTexto(item.resumo),
    textoIndexado: normalizarTexto(item.textoIndexado),
    publico: normalizarTexto(item.publico),
    tipo: normalizarTexto(item.tipo)
  };

  let pontuacaoTotal = 0;
  let termosEncontrados = 0;

  termos.forEach(termo => {
    let pontosTermo = 0;

    if (campos.titulo.includes(termo)) {
      pontosTermo += 10;
      if (campos.titulo.startsWith(termo)) pontosTermo += 3;
    }
    if (campos.palavrasChave.includes(termo)) pontosTermo += 8;
    if (campos.tema.includes(termo)) pontosTermo += 6;
    if (campos.subtema.includes(termo)) pontosTermo += 5;
    if (campos.resumo.includes(termo)) pontosTermo += 4;
    if (campos.textoIndexado.includes(termo)) pontosTermo += 2;
    if (campos.publico.includes(termo)) pontosTermo += 2;
    if (campos.tipo.includes(termo)) pontosTermo += 1;

    if (pontosTermo > 0) {
      termosEncontrados++;
      pontuacaoTotal += pontosTermo;
    }
  });

  if (pontuacaoTotal === 0) return 0;

  // Bônus quando o conteúdo contempla TODAS as palavras pesquisadas
  if (termos.length > 1 && termosEncontrados === termos.length) {
    pontuacaoTotal = Math.round(pontuacaoTotal * 1.5) + 10;
  }

  return pontuacaoTotal;
}

/**
 * Retorna conteúdos filtrados (sem pontuação de busca textual).
 */
function getConteudos(filtros) {
  const f = filtros || {};
  const incluirAdmin = Boolean(f.incluirAdmin) && usuarioEhAdministrador();
  let lista = carregarTodosRegistros_(incluirAdmin);

  lista = aplicarFiltrosLista_(lista, f);
  return lista.map(sanitizarItemParaFrontend_);
}

/**
 * Pesquisa conteúdos com pontuação de relevância, filtros combinados,
 * sugestão de termos semelhantes e paginação (12 itens por página).
 */
function pesquisarConteudos(termo, filtros, pagina) {
  const f = filtros || {};
  const paginaAtual = Math.max(1, parseInt(pagina, 10) || 1);
  const termoLimpo = String(termo || '').trim();
  const termoNorm = normalizarTexto(termoLimpo);
  const tokens = termoNorm.split(/\s+/).filter(t => t.length >= 2);

  const todos = carregarTodosRegistros_(false);
  let filtrados = aplicarFiltrosLista_(todos, f);

  let ranqueados = [];

  if (tokens.length > 0) {
    filtrados.forEach(item => {
      const score = calcularRelevancia(item, tokens);
      if (score > 0) {
        ranqueados.push({ item: item, score: score });
      }
    });

    ranqueados.sort((a, b) => {
      if (b.score !== a.score) return b.score - a.score;
      if (a.item.destaque !== b.item.destaque) return a.item.destaque ? -1 : 1;
      return a.item.ordem - b.item.ordem;
    });
  } else {
    ranqueados = filtrados.map(item => ({ item: item, score: 0 }));
  }

  const totalResultados = ranqueados.length;
  const totalPaginas = Math.max(1, Math.ceil(totalResultados / RESULTADOS_POR_PAGINA));
  const paginaSegura = Math.min(paginaAtual, totalPaginas);
  const inicio = (paginaSegura - 1) * RESULTADOS_POR_PAGINA;
  const paginados = ranqueados
    .slice(inicio, inicio + RESULTADOS_POR_PAGINA)
    .map(r => {
      const dto = sanitizarItemParaFrontend_(r.item);
      dto.relevancia = r.score;
      return dto;
    });

  // Registra telemetria da pesquisa na primeira página
  if (paginaSegura === 1 && (tokens.length > 0 || Object.keys(f).some(k => f[k]))) {
    try {
      registrarPesquisa({
        termo: termoLimpo || '[Filtro sem texto]',
        quantidade: totalResultados,
        filtros: f
      });
    } catch (e) {
      // Falha de log não interrompe a busca do usuário
    }
  }

  // Sugere termos semelhantes quando nenhum resultado for encontrado
  let sugestoesSemelhantes = [];
  if (totalResultados === 0 && tokens.length > 0) {
    sugestoesSemelhantes = gerarSugestoesSemelhantes_(todos, tokens);
  }

  // Coleta opções dinâmicas disponíveis para os seletores de filtro
  const opcoesFiltros = extrairOpcoesFiltros_(todos);

  return {
    resultados: paginados,
    totalResultados: totalResultados,
    paginaAtual: paginaSegura,
    totalPaginas: totalPaginas,
    porPagina: RESULTADOS_POR_PAGINA,
    sugestoesSemelhantes: sugestoesSemelhantes,
    opcoesFiltros: opcoesFiltros
  };
}

/**
 * Aplica filtros combinados sobre a lista de conteúdos.
 */
function aplicarFiltrosLista_(lista, f) {
  return lista.filter(item => {
    if (f.tema && normalizarTexto(item.tema) !== normalizarTexto(f.tema)) return false;
    if (f.subtema && normalizarTexto(item.subtema) !== normalizarTexto(f.subtema)) return false;
    if (f.publico && !normalizarTexto(item.publico).includes(normalizarTexto(f.publico))) return false;
    if (f.tipo && normalizarTexto(item.tipo) !== normalizarTexto(f.tipo)) return false;
    if (f.responsavel && normalizarTexto(item.responsavel) !== normalizarTexto(f.responsavel)) return false;
    if (f.status && normalizarTexto(item.status) !== normalizarTexto(f.status)) return false;
    if (f.ano && String(item.ano) !== String(f.ano)) return false;
    if (f.destaque === true || normalizarTexto(f.destaque) === 'sim') {
      if (!item.destaque) return false;
    }
    if (f.apenasModelos) {
      const tiposModelo = ['checklist', 'planilha', 'modelo', 'formulario', 'ata', 'plano de acao'];
      if (!tiposModelo.some(tm => normalizarTexto(item.tipo).includes(tm))) return false;
    }
    if (f.apenasFluxos) {
      if (!normalizarTexto(item.tipo).includes('fluxo') && !normalizarTexto(item.tema).includes('fluxo')) {
        return false;
      }
    }
    return true;
  });
}

/**
 * Retorna os detalhes completos de um conteúdo por ID, incluindo até 4 materiais relacionados.
 */
function getConteudoPorId(id) {
  const idBusca = String(id || '').trim();
  if (!idBusca) return null;

  const todos = carregarTodosRegistros_(false);
  const encontrado = todos.find(item => item.id === idBusca);
  if (!encontrado) return null;

  const relacionados = obterRelacionados_(encontrado, todos, 4).map(sanitizarItemParaFrontend_);
  const dto = sanitizarItemParaFrontend_(encontrado);
  dto.relacionados = relacionados;
  return dto;
}

/**
 * Seleciona até 4 conteúdos relacionados:
 * Prioriza coluna Relacionados (R); se vazia, usa Tema -> Subtema -> Público -> Palavras-chave.
 */
function obterRelacionados_(itemBase, todos, limite) {
  const selecionados = [];
  const idsUsados = new Set([itemBase.id]);

  // 1. Relacionados explícitos na planilha
  if (itemBase.relacionadosIds) {
    const idsExplicitos = itemBase.relacionadosIds.split(',').map(s => s.trim()).filter(Boolean);
    idsExplicitos.forEach(relId => {
      if (selecionados.length >= limite || idsUsados.has(relId)) return;
      const match = todos.find(t => t.id === relId);
      if (match) {
        selecionados.push(match);
        idsUsados.add(match.id);
      }
    });
  }

  if (selecionados.length >= limite) return selecionados;

  // 2. Pontuação por afinidade automática (Tema, Subtema, Público, Palavras-chave)
  const kwBase = new Set(normalizarTexto(itemBase.palavrasChave).split(',').map(s => s.trim()).filter(Boolean));
  const candidatos = todos
    .filter(t => !idsUsados.has(t.id))
    .map(cand => {
      let afinidade = 0;
      if (normalizarTexto(cand.tema) === normalizarTexto(itemBase.tema)) afinidade += 10;
      if (itemBase.subtema && normalizarTexto(cand.subtema) === normalizarTexto(itemBase.subtema)) afinidade += 7;
      if (itemBase.publico && normalizarTexto(cand.publico) === normalizarTexto(itemBase.publico)) afinidade += 4;

      const kwCand = normalizarTexto(cand.palavrasChave).split(',').map(s => s.trim()).filter(Boolean);
      kwCand.forEach(k => {
        if (kwBase.has(k)) afinidade += 3;
      });

      return { item: cand, afinidade: afinidade };
    })
    .filter(c => c.afinidade > 0)
    .sort((a, b) => b.afinidade - a.afinidade);

  candidatos.slice(0, limite - selecionados.length).forEach(c => selecionados.push(c.item));
  return selecionados;
}

/**
 * Retorna os temas disponíveis dinamicamente a partir do catálogo.
 */
function getTemas() {
  const todos = carregarTodosRegistros_(false);
  const mapaTemas = new Map();

  todos.forEach(item => {
    const nomeTema = item.tema || 'Geral';
    if (!mapaTemas.has(nomeTema)) {
      mapaTemas.set(nomeTema, {
        nome: nomeTema,
        quantidade: 0,
        descricao: obterDescricaoTema_(nomeTema),
        icone: obterIconeTema_(nomeTema)
      });
    }
    mapaTemas.get(nomeTema).quantidade++;
  });

  return Array.from(mapaTemas.values()).sort((a, b) => a.nome.localeCompare(b.nome, 'pt-BR'));
}

/**
 * Retorna os conteúdos marcados como Destaque = Sim.
 */
function getDestaques() {
  const todos = carregarTodosRegistros_(false);
  return todos.filter(i => i.destaque).slice(0, 6).map(sanitizarItemParaFrontend_);
}

/**
 * Carrega as perguntas frequentes da aba FAQ.
 */
function getFaqs() {
  try {
    const ss = getPlanilhaCatalogo_();
    const sheet = ss.getSheetByName(DEFAULTS.ABAS.FAQ);
    if (!sheet || sheet.getLastRow() <= 1) return [];

    const dados = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
    return dados
      .filter(r => r[2] && r[3] && normalizarTexto(r[5] || 'Publicado') === 'publicado')
      .map(r => ({
        id: String(r[0] || ''),
        categoria: String(r[1] || 'Geral'),
        pergunta: String(r[2] || ''),
        resposta: String(r[3] || ''),
        ordem: Number(r[4]) || 99
      }))
      .sort((a, b) => a.ordem - b.ordem);
  } catch (e) {
    return [];
  }
}

function extrairOpcoesFiltros_(todos) {
  const sets = {
    temas: new Set(),
    subtemas: new Set(),
    publicos: new Set(),
    tipos: new Set(),
    responsaveis: new Set(),
    anos: new Set()
  };

  todos.forEach(i => {
    if (i.tema) sets.temas.add(i.tema);
    if (i.subtema) sets.subtemas.add(i.subtema);
    if (i.publico) sets.publicos.add(i.publico);
    if (i.tipo) sets.tipos.add(i.tipo);
    if (i.responsavel) sets.responsaveis.add(i.responsavel);
    if (i.ano) sets.anos.add(i.ano);
  });

  const toSortedArray = s => Array.from(s).sort((a, b) => String(a).localeCompare(String(b), 'pt-BR'));
  return {
    temas: toSortedArray(sets.temas),
    subtemas: toSortedArray(sets.subtemas),
    publicos: toSortedArray(sets.publicos),
    tipos: toSortedArray(sets.tipos),
    responsaveis: toSortedArray(sets.responsaveis),
    anos: Array.from(sets.anos).sort((a, b) => String(b).localeCompare(String(a)))
  };
}

function gerarSugestoesSemelhantes_(todos, tokens) {
  const sugestoes = new Set();
  todos.forEach(item => {
    const temaNorm = normalizarTexto(item.tema);
    const palavras = item.palavrasChave.split(',').map(p => p.trim()).filter(Boolean);
    tokens.forEach(tok => {
      const prefixo = tok.substring(0, Math.max(3, tok.length - 2));
      if (temaNorm.includes(prefixo)) sugestoes.add(item.tema);
      palavras.forEach(pw => {
        if (normalizarTexto(pw).includes(prefixo)) sugestoes.add(pw);
      });
    });
  });
  return Array.from(sugestoes).slice(0, 5);
}

function sanitizarItemParaFrontend_(item) {
  // Omite textoIndexado extenso no payload de listagem para manter respostas leves
  const copia = Object.assign({}, item);
  delete copia.textoIndexado;
  return copia;
}

function obterDescricaoTema_(tema) {
  const t = normalizarTexto(tema);
  if (t.includes('planejamento')) return 'Diretrizes, planos anuais, priorização e governança tecnológica.';
  if (t.includes('chamado')) return 'Abertura, triagem, atendimento e escalonamento de ocorrências.';
  if (t.includes('infraestrutura') || t.includes('conectividade')) return 'Redes escolares, links de internet, Wi-Fi e requisitos técnicos.';
  if (t.includes('equipamento') || t.includes('inventario')) return 'Controle patrimonial, tombamento, distribuição e manutenção.';
  if (t.includes('documental')) return 'Padronização de registros, atas, ofícios e arquivamento.';
  if (t.includes('monitoramento') || t.includes('indicador')) return 'Acompanhamento de metas, painéis e indicadores regionais.';
  if (t.includes('escolatec')) return 'Ações do programa EscolaTEC e ciclos de melhoria contínua.';
  if (t.includes('visita')) return 'Roteiros, checklists e relatórios de acompanhamento presencial.';
  if (t.includes('papel') || t.includes('responsabilidade')) return 'Atribuições de SEINTECs, SETECs, COEGD e gestores escolares.';
  if (t.includes('plano')) return 'Estruturação e execução de planos de ação nas UREs e escolas.';
  if (t.includes('fluxo')) return 'Diagramas operacionais e fluxos de decisão institucional.';
  if (t.includes('modelo') || t.includes('formulario')) return 'Templates editáveis, formulários padrão e listas de verificação.';
  return 'Orientações operacionais e materiais de referência da rede.';
}

function obterIconeTema_(tema) {
  const t = normalizarTexto(tema);
  if (t.includes('planejamento')) return 'compass';
  if (t.includes('chamado')) return 'headset';
  if (t.includes('infraestrutura') || t.includes('conectividade')) return 'wifi';
  if (t.includes('equipamento') || t.includes('inventario')) return 'laptop';
  if (t.includes('documental')) return 'folder';
  if (t.includes('monitoramento')) return 'chart';
  if (t.includes('visita')) return 'clipboard';
  if (t.includes('papel')) return 'users';
  if (t.includes('fluxo')) return 'git-branch';
  if (t.includes('modelo')) return 'file-check';
  return 'book';
}