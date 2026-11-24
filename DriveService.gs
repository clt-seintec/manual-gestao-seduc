/**
 * ============================================================================
 * DriveService.gs - Indexação do Google Drive e Extração Textual Segura
 * Portal: Manual de Gestão da Tecnologia — SEDUC
 * ============================================================================
 */

const LIMITE_TEXTO_INDEXADO = 15000; // Caracteres máximos por célula na planilha

/**
 * Percorre a pasta configurada no Google Drive e atualiza a aba CONTEUDOS.
 * - Preserva metadados editados manualmente (Tema, Subtema, Público, Resumo, etc.)
 * - Atualiza datas e URLs de arquivos existentes
 * - Cadastra novos arquivos encontrados
 * - Marca como "Arquivado" arquivos removidos ou na lixeira (nunca exclui linhas)
 */
function atualizarIndiceDoDrive() {
  const emailAtivo = Session.getActiveUser().getEmail();
  const emailDono = Session.getEffectiveUser().getEmail();
  const ehDonoOuTrigger = !emailAtivo || emailAtivo === emailDono;

  if (!ehDonoOuTrigger && !usuarioEhAdministrador()) {
    throw new Error('Acesso negado: apenas administradores podem atualizar o índice.');
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error('Outra atualização de índice já está em andamento. Aguarde alguns instantes.');
  }

  const stats = {
    novos: 0,
    atualizados: 0,
    arquivados: 0,
    erros: 0
  };

  try {
    const props = PropertiesService.getScriptProperties();
    const folderId = props.getProperty('DRIVE_FOLDER_ID') || DEFAULTS.DRIVE_FOLDER_ID;
    const pastaRaiz = DriveApp.getFolderById(folderId);

    const ss = getPlanilhaCatalogo_();
    const sheet = ss.getSheetByName(DEFAULTS.ABAS.CONTEUDOS);
    const lastRow = sheet.getLastRow();

    // Mapeia registros existentes pelo "ID do arquivo" (Coluna I = índice 8)
    const mapaExistentes = new Map();
    let maiorSequencial = 0;

    if (lastRow > 1) {
      const dados = sheet.getRange(2, 1, lastRow - 1, 21).getValues();
      dados.forEach((linha, idx) => {
        const idCatalogo = String(linha[0] || '');
        const matchSeq = idCatalogo.match(/(\d+)$/);
        if (matchSeq) {
          const num = parseInt(matchSeq[1], 10);
          if (num > maiorSequencial) maiorSequencial = num;
        }
        const fileId = String(linha[8] || '').trim();
        if (fileId) {
          mapaExistentes.set(fileId, {
            rowIndex: idx + 2,
            valores: linha
          });
        }
      });
    }

    // Varre arquivos recursivamente no Google Drive
    const arquivosEncontrados = new Map();
    varrerPastaRecursivamente_(pastaRaiz, '', arquivosEncontrados, stats);

    const novasLinhas = [];
    const tz = Session.getScriptTimeZone() || 'America/Sao_Paulo';

    arquivosEncontrados.forEach((info, fileId) => {
      try {
        const dataFormatada = Utilities.formatDate(info.lastUpdated, tz, 'dd/MM/yyyy');
        const tipoInferido = inferirTipoMaterial_(info.mimeType, info.name);
        const temaInferido = inferirTemaPorPasta_(info.caminhoPasta);

        if (mapaExistentes.has(fileId)) {
          // Atualiza registro existente preservando curadoria manual
          const reg = mapaExistentes.get(fileId);
          const v = reg.valores;
          const dataAnterior = v[12] instanceof Date
            ? Utilities.formatDate(v[12], tz, 'dd/MM/yyyy')
            : String(v[12] || '');

          const mudouArquivo = dataAnterior !== dataFormatada || !String(v[16] || '').trim();

          v[1] = v[1] || limparExtensaoNome_(info.name); // Mantém título curado ou usa nome
          v[9] = info.url; // Atualiza URL
          if (v[10] === 'Arquivado') {
            v[10] = 'Publicado'; // Reativa se voltou à pasta
          }
          v[12] = dataFormatada;
          if (!v[18]) {
            v[18] = 'Pasta: ' + (info.caminhoPasta || 'Raiz');
          }

          if (mudouArquivo) {
            const textoExtraido = extrairTextoArquivo_(info.file, info.mimeType);
            if (textoExtraido) {
              v[16] = textoExtraido;
            }
          }

          sheet.getRange(reg.rowIndex, 1, 1, 21).setValues([v]);
          stats.atualizados++;
        } else {
          // Cadastra novo conteúdo
          maiorSequencial++;
          const novoId = 'MG-' + String(maiorSequencial).padStart(4, '0');
          const tituloLimpo = limparExtensaoNome_(info.name);
          const textoExtraido = extrairTextoArquivo_(info.file, info.mimeType);

          novasLinhas.push([
            novoId,                                      // A: ID
            tituloLimpo,                                 // B: Título
            'Material técnico: ' + tituloLimpo + '.',    // C: Resumo inicial
            temaInferido,                                // D: Tema
            info.caminhoPasta || 'Geral',                // E: Subtema
            'SEINTEC e SETEC',                           // F: Público padrão
            tipoInferido,                                // G: Tipo
            gerarPalavrasChaveIniciais_(tituloLimpo, temaInferido, tipoInferido), // H: Palavras-chave
            fileId,                                      // I: ID do arquivo
            info.url,                                    // J: URL
            'Publicado',                                 // K: Status
            '1.0',                                       // L: Versão
            dataFormatada,                               // M: Data da atualização
            'COEGD / GT Manual',                         // N: Responsável
            'Não',                                       // O: Destaque
            maiorSequencial,                             // P: Ordem
            textoExtraido,                               // Q: Texto indexado
            '',                                          // R: Relacionados
            'Pasta: ' + (info.caminhoPasta || 'Raiz'),   // S: Observações
            '',                                          // T: Imagem
            ''                                           // U: Objetivo
          ]);
          stats.novos++;
        }
      } catch (errItem) {
        stats.erros++;
        registrarLogErro('atualizarIndiceDoDrive:item', 'Erro ao processar arquivo: ' + info.name, errItem.message);
      }
    });

    if (novasLinhas.length > 0) {
      const startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, novasLinhas.length, 21).setValues(novasLinhas);
    }

    // Marca como Arquivado registros cujo arquivo não existe mais na pasta
    mapaExistentes.forEach((reg, fileId) => {
      if (!arquivosEncontrados.has(fileId)) {
        const statusAtual = String(reg.valores[10] || '');
        if (statusAtual !== 'Arquivado') {
          sheet.getRange(reg.rowIndex, 11).setValue('Arquivado');
          stats.arquivados++;
        }
      }
    });

    limparCache();
    registrarLogInfo_('atualizarIndiceDoDrive', 'Indexação concluída', JSON.stringify(stats));
    return { sucesso: true, stats: stats };

  } catch (erro) {
    registrarLogErro('atualizarIndiceDoDrive', erro.message, erro.stack);
    throw new Error('Falha na indexação do Google Drive: ' + erro.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Varre recursivamente pastas e subpastas no Google Drive.
 */
function varrerPastaRecursivamente_(pasta, caminhoAtual, mapaDestino, stats) {
  const arquivos = pasta.getFiles();
  while (arquivos.hasNext()) {
    const file = arquivos.next();
    if (file.isTrashed()) continue;
    const mime = file.getMimeType();
    // Ignora atalhos quebrados ou a própria planilha de catálogo se estiver na pasta
    if (file.getName() === DEFAULTS.NOME_PLANILHA) continue;

    mapaDestino.set(file.getId(), {
      file: file,
      id: file.getId(),
      name: file.getName(),
      mimeType: mime,
      url: file.getUrl(),
      lastUpdated: file.getLastUpdated(),
      caminhoPasta: caminhoAtual
    });
  }

  const subpastas = pasta.getFolders();
  while (subpastas.hasNext()) {
    const sub = subpastas.next();
    if (sub.isTrashed()) continue;
    const proximoCaminho = caminhoAtual ? caminhoAtual + ' / ' + sub.getName() : sub.getName();
    varrerPastaRecursivamente_(sub, proximoCaminho, mapaDestino, stats);
  }
}

/**
 * Extrai texto de documentos Google nativos e arquivos de texto puro
 * sem modificar os arquivos originais.
 */
function extrairTextoArquivo_(file, mimeType) {
  try {
    let texto = '';

    if (mimeType === MimeType.GOOGLE_DOCS) {
      const doc = DocumentApp.openById(file.getId());
      texto = doc.getBody().getText();
    } else if (mimeType === MimeType.GOOGLE_SHEETS) {
      const ss = SpreadsheetApp.openById(file.getId());
      const sheets = ss.getSheets().slice(0, 5); // Limita às 5 primeiras abas por desempenho
      texto = sheets.map(s => {
        const vals = s.getDataRange().getValues();
        return vals.map(r => r.join(' ')).join('\n');
      }).join('\n');
    } else if (mimeType === MimeType.GOOGLE_SLIDES) {
      const pres = SlidesApp.openById(file.getId());
      texto = pres.getSlides().map(slide => {
        return slide.getShapes().map(sh => {
          try {
            return sh.getText().asString();
          } catch (e) {
            return '';
          }
        }).join(' ');
      }).join('\n');
    } else if (
      mimeType === MimeType.PLAIN_TEXT ||
      mimeType === 'text/csv' ||
      mimeType === 'text/markdown'
    ) {
      texto = file.getBlob().getDataAsString('UTF-8');
    }
    // Nota sobre PDF e Word (.docx):
    // Por segurança e integridade dos originais, não executamos conversão destrutiva.
    // Caso o serviço avançado "Drive API v3" seja habilitado no projeto, pode-se ativar
    // a extração temporária via Drive.Files.copy({mimeType: MimeType.GOOGLE_DOCS}) + exclusão imediata da cópia temporária.

    return String(texto || '')
      .replace(/\s+/g, ' ')
      .trim()
      .substring(0, LIMITE_TEXTO_INDEXADO);
  } catch (e) {
    registrarLogErro('extrairTextoArquivo_', 'Não foi possível extrair texto de ' + file.getName(), e.message);
    return '';
  }
}

function limparExtensaoNome_(nome) {
  return String(nome || '').replace(/\.(pdf|docx?|xlsx?|pptx?|txt|csv|png|jpe?g)$/i, '').trim();
}

function inferirTipoMaterial_(mimeType, nome) {
  const n = normalizarTexto(nome);
  if (n.includes('checklist')) return 'Checklist';
  if (n.includes('fluxo')) return 'Fluxo';
  if (n.includes('modelo') || n.includes('ata') || n.includes('formulario')) return 'Modelo';
  if (n.includes('manual') || n.includes('guia')) return 'Manual';
  if (n.includes('orientacao') || n.includes('protocolo')) return 'Orientação';

  if (mimeType === MimeType.GOOGLE_SHEETS || mimeType.includes('spreadsheet') || mimeType.includes('excel')) {
    return 'Planilha';
  }
  if (mimeType === MimeType.GOOGLE_SLIDES || mimeType.includes('presentation') || mimeType.includes('powerpoint')) {
    return 'Apresentação';
  }
  if (mimeType.startsWith('image/')) {
    return 'Fluxograma / Imagem';
  }
  return 'Documento';
}

function inferirTemaPorPasta_(caminhoPasta) {
  if (!caminhoPasta) return 'Orientações técnicas';
  const primeiraPasta = caminhoPasta.split('/')[0].trim();
  // Remove numeração inicial de pasta como "01 - Planejamento e governança"
  return primeiraPasta.replace(/^\d+[\s._-]+/, '').trim() || 'Orientações técnicas';
}

function gerarPalavrasChaveIniciais_(titulo, tema, tipo) {
  const palavras = (titulo + ' ' + tema + ' ' + tipo)
    .toLowerCase()
    .replace(/[^\w\sÀ-ÿ]/g, ' ')
    .split(/\s+/)
    .filter(p => p.length > 3);
  return Array.from(new Set(palavras)).slice(0, 10).join(', ');
}

/**
 * Gera URLs auxiliares para visualização, cópia e download conforme política do portal.
 */
function gerarLinksAcessoDrive_(fileId, urlOriginal) {
  if (!fileId) {
    return {
      urlVisualizar: urlOriginal || '',
      urlCopiar: '',
      urlDownload: ''
    };
  }
  const url = String(urlOriginal || '');
  const isGoogleNative = url.includes('docs.google.com/');

  return {
    urlVisualizar: url || ('https://drive.google.com/file/d/' + fileId + '/view'),
    urlCopiar: isGoogleNative
      ? url.replace(/\/(edit|view|preview).*$/, '/copy')
      : 'https://drive.google.com/uc?export=download&id=' + fileId,
    urlDownload: isGoogleNative
      ? url.replace(/\/(edit|view|preview).*$/, '/export?format=pdf')
      : 'https://drive.google.com/uc?export=download&id=' + fileId
  };
}