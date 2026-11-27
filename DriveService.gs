/**
 * ============================================================================
 * DriveService.gs - Sincronização, Links e Indexação Documental (Drive -> Sheets)
 * Portal: Manual de Gestão da Tecnologia — SEDUC
 * ============================================================================
 */

const MAX_CHARS_INDEXACAO = 45000;

/**
 * Gera os links de visualização, cópia e download usados pelo SearchService.gs
 * e pelo modal de detalhes do portal.
 */
function gerarLinksAcessoDrive_(fileId, tipoOuUrl, urlFallback) {
  const idLimpo = String(fileId || '').trim();
  const urlBase = String(urlFallback || (String(tipoOuUrl || '').startsWith('http') ? tipoOuUrl : '') || '').trim();

  if (!idLimpo) {
    return {
      url: urlBase || '#',
      urlCopiar: urlBase || '#',
      urlDownload: urlBase || '#'
    };
  }

  return {
    url: urlBase || ('https://drive.google.com/file/d/' + idLimpo + '/view'),
    urlCopiar: 'https://docs.google.com/document/d/' + idLimpo + '/copy',
    urlDownload: 'https://drive.google.com/uc?export=download&id=' + idLimpo
  };
}

/**
 * Localiza automaticamente o ID da pasta raiz do Google Drive no Config.gs,
 * nas Propriedades do Script ou na aba CONFIG da planilha.
 */
function obterIdPastaRaizDrive_() {
  if (typeof CONFIG !== 'undefined' && CONFIG) {
    const idGlobal = CONFIG.DRIVE_FOLDER_ID || CONFIG.ROOT_FOLDER_ID ||
      CONFIG.PASTA_RAIZ_ID || CONFIG.FOLDER_ID || CONFIG.ID_PASTA_DRIVE;
    if (idGlobal && String(idGlobal).trim().length > 10) {
      return String(idGlobal).trim();
    }
  }

  if (typeof getPortalConfig_ === 'function') {
    const cfg = getPortalConfig_();
    if (cfg) {
      const idCfg = cfg.rootFolderId || cfg.driveFolderId || cfg.folderId || cfg.pastaRaizId;
      if (idCfg && String(idCfg).trim().length > 10) return String(idCfg).trim();
    }
  }

  const props = PropertiesService.getScriptProperties().getProperties();
  const chavesPossiveis = ['ROOT_FOLDER_ID', 'DRIVE_FOLDER_ID', 'FOLDER_ID', 'PASTA_RAIZ_ID', 'ID_PASTA_DRIVE'];
  for (let i = 0; i < chavesPossiveis.length; i++) {
    const val = props[chavesPossiveis[i]];
    if (val && String(val).trim().length > 10) {
      return String(val).trim();
    }
  }

  const ss = obterPlanilhaCatalogo_();
  const sheetConfig = ss.getSheetByName('CONFIG');
  if (sheetConfig && sheetConfig.getLastRow() >= 1) {
    const dadosCfg = sheetConfig.getDataRange().getValues();
    for (let r = 0; r < dadosCfg.length; r++) {
      const chave = String(dadosCfg[r][0] || '').toUpperCase();
      const valor = String(dadosCfg[r][1] || '').trim();
      if ((chave.includes('FOLDER') || chave.includes('PASTA') || chave.includes('DRIVE')) && valor.length > 15) {
        return valor;
      }
    }
  }

  const sheetConteudos = ss.getSheetByName('CONTEUDOS');
  if (sheetConteudos && sheetConteudos.getLastRow() >= 2) {
    const fileIds = sheetConteudos.getRange(2, 9, Math.min(10, sheetConteudos.getLastRow() - 1), 1).getValues();
    for (let j = 0; j < fileIds.length; j++) {
      const fId = String(fileIds[j][0] || '').trim();
      if (fId) {
        try {
          const arquivo = DriveApp.getFileById(fId);
          const pais = arquivo.getParents();
          if (pais.hasNext()) {
            let pastaAtual = pais.next();
            for (let nivel = 0; nivel < 2; nivel++) {
              const avos = pastaAtual.getParents();
              if (!avos.hasNext()) break;
              const candidata = avos.next();
              if (candidata.getName().toLowerCase().includes('manual') || candidata.getName().toLowerCase().includes('seintec')) {
                pastaAtual = candidata;
              }
            }
            return pastaAtual.getId();
          }
        } catch (e) {}
      }
    }
  }

  throw new Error('ID da pasta raiz do Google Drive não encontrado.');
}

function obterPlanilhaCatalogo_() {
  if (typeof getCatalogSpreadsheet_ === 'function') {
    return getCatalogSpreadsheet_();
  }
  if (typeof CONFIG !== 'undefined' && CONFIG && CONFIG.SPREADSHEET_ID) {
    return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  }
  const propId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (propId) {
    return SpreadsheetApp.openById(propId);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Sincroniza todos os arquivos da pasta raiz do Drive com a aba CONTEUDOS
 * e preenche a Coluna Q (Texto indexado).
 */
function atualizarIndiceDoDrive() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  const stats = { novos: 0, atualizados: 0, arquivados: 0, erros: 0 };

  try {
    const rootFolderId = obterIdPastaRaizDrive_();
    Logger.log('Pasta raiz detectada: ' + rootFolderId);

    const ss = obterPlanilhaCatalogo_();
    const sheet = ss.getSheetByName('CONTEUDOS');
    if (!sheet) {
      throw new Error('Aba "CONTEUDOS" não encontrada na planilha de catálogo.');
    }

    const lastRow = sheet.getLastRow();
    const existentesPorFileId = new Map();
    let maiorNumeroId = 0;

    if (lastRow >= 2) {
      const dadosAtuais = sheet.getRange(2, 1, lastRow - 1, 19).getValues();
      dadosAtuais.forEach((row, idx) => {
        const idRegistro = String(row[0] || '').trim();
        const fileId = String(row[8] || '').trim();

        const matchNum = idRegistro.match(/(\d+)$/);
        if (matchNum) {
          const num = parseInt(matchNum[1], 10);
          if (num > maiorNumeroId) maiorNumeroId = num;
        }

        if (fileId) {
          existentesPorFileId.set(fileId, {
            rowIndex: idx + 2,
            valores: row
          });
        }
      });
    }

    const rootFolder = DriveApp.getFolderById(rootFolderId);
    const arquivosDrive = [];
    coletarArquivosRecursivo_(rootFolder, '', '', arquivosDrive);

    const fileIdsEncontrados = new Set();

    arquivosDrive.forEach((item) => {
      const file = item.file;
      const fileId = file.getId();
      fileIdsEncontrados.add(fileId);

      try {
        const nomeOriginal = file.getName();
        const tituloLimpo = limparTituloArquivo_(nomeOriginal);
        const tema = item.tema || 'Geral';
        const subtema = item.subtema || tema;
        const tipo = inferirTipoConteudo_(nomeOriginal, file.getMimeType());
        const publico = 'SEINTEC e SETEC';
        const palavrasChave = gerarPalavrasChave_(tituloLimpo, tema, subtema);
        const url = file.getUrl();
        const dataAtualizacao = Utilities.formatDate(
          file.getLastUpdated(),
          Session.getScriptTimeZone() || 'America/Sao_Paulo',
          'dd/MM/yyyy'
        );

        const textoIndexado = extrairTextoArquivoDrive_(file);
        const resumoExtraido = gerarResumoAutomatico_(textoIndexado, tituloLimpo);

        if (existentesPorFileId.has(fileId)) {
          const reg = existentesPorFileId.get(fileId);
          const row = reg.valores;

          row[1] = row[1] || tituloLimpo;
          if (!row[2] || String(row[2]).startsWith('Material técnico:')) {
            row[2] = resumoExtraido;
          }
          row[3] = tema;
          row[4] = subtema;
          row[6] = row[6] || tipo;
          row[7] = row[7] || palavrasChave;
          row[9] = url;
          row[12] = dataAtualizacao;
          if (textoIndexado) {
            row[16] = textoIndexado; // Coluna Q: Texto indexado
          }
          row[18] = 'Pasta: ' + (item.caminhoPasta || 'Raiz');

          sheet.getRange(reg.rowIndex, 1, 1, 19).setValues([row]);
          stats.atualizados++;
        } else {
          maiorNumeroId++;
          const novoId = 'MG-' + String(maiorNumeroId).padStart(4, '0');
          const novaOrdem = lastRow + stats.novos;

          const novaLinha = [
            novoId,
            tituloLimpo,
            resumoExtraido,
            tema,
            subtema,
            publico,
            tipo,
            palavrasChave,
            fileId,
            url,
            'Publicado',
            '1.0',
            dataAtualizacao,
            'COEGD / GT Manual',
            'Não',
            novaOrdem,
            textoIndexado,
            '',
            'Pasta: ' + (item.caminhoPasta || 'Raiz')
          ];

          sheet.appendRow(novaLinha);
          stats.novos++;
        }
      } catch (errArq) {
        stats.erros++;
        Logger.log('Erro no arquivo ' + file.getName() + ': ' + errArq.message);
      }
    });

    existentesPorFileId.forEach((reg, fileId) => {
      if (!fileIdsEncontrados.has(fileId) && reg.valores[10] === 'Publicado') {
        sheet.getRange(reg.rowIndex, 11).setValue('Arquivado');
        stats.arquivados++;
      }
    });

    if (typeof limparCache === 'function') {
      limparCache();
    }

    Logger.log('Sincronização concluída: ' + JSON.stringify(stats));
    return { sucesso: true, stats: stats };

  } finally {
    lock.releaseLock();
  }
}

function coletarArquivosRecursivo_(folder, temaAtual, subtemaAtual, acumulador) {
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    const nome = file.getName();
    if (nome.startsWith('~$') || nome.startsWith('.') || nome.startsWith('TEMP_INDEX_')) {
      continue;
    }
    acumulador.push({
      file: file,
      tema: temaAtual || 'Orientações Técnicas',
      subtema: subtemaAtual || temaAtual || 'Geral',
      caminhoPasta: subtemaAtual && subtemaAtual !== temaAtual
        ? temaAtual + ' / ' + subtemaAtual
        : temaAtual
    });
  }

  const subFolders = folder.getFolders();
  while (subFolders.hasNext()) {
    const sub = subFolders.next();
    const nomeSub = sub.getName();
    if (nomeSub.startsWith('.')) continue;

    const proximoTema = temaAtual ? temaAtual : nomeSub;
    const proximoSubtema = temaAtual ? nomeSub : nomeSub;
    coletarArquivosRecursivo_(sub, proximoTema, proximoSubtema, acumulador);
  }
}

/**
 * Extrai texto direto do XML interno de arquivos .docx (sem conversão Drive/DocumentApp).
 */
function extrairTextoDocxNativo_(file) {
  const zipBlob = file.getBlob().setContentType('application/zip');
  const entries = Utilities.unzip(zipBlob);
  const partesTexto = [];

  entries.forEach(entry => {
    const entryName = entry.getName();
    if (entryName === 'word/document.xml' || /^word\/(header|footer)\d*\.xml$/.test(entryName)) {
      const xml = entry.getDataAsString('UTF-8');
      const limpo = xml
        .replace(/<\/w:p>/g, '\n')
        .replace(/<\/w:tc>/g, ' | ')
        .replace(/<w:br[^>]*\/>/g, '\n')
        .replace(/<[^>]+>/g, '')
        .replace(/&amp;/g, '&')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'");
      if (limpo.trim()) {
        partesTexto.push(limpo.trim());
      }
    }
  });

  return partesTexto.join('\n\n');
}

/**
 * Extrai o texto de Google Docs, Sheets, Slides, Markdown/TXT, .docx (via XML nativo),
 * .pdf (via Drive API v3 OCR/Export) e planilhas Excel.
 */
function extrairTextoArquivoDrive_(file) {
  const mime = file.getMimeType();
  const fileId = file.getId();
  const nome = file.getName().toLowerCase();

  try {
    // 1. Documentos Google nativos
    if (mime === MimeType.GOOGLE_DOCS) {
      return sanitizarTextoIndexado_(DocumentApp.openById(fileId).getBody().getText());
    }

    // 2. Planilhas Google nativas
    if (mime === MimeType.GOOGLE_SHEETS) {
      const ss = SpreadsheetApp.openById(fileId);
      const textoSheets = ss.getSheets().map(s =>
        '[' + s.getName() + ']\n' +
        s.getDataRange().getDisplayValues().map(r => r.filter(Boolean).join(' | ')).join('\n')
      ).join('\n\n');
      return sanitizarTextoIndexado_(textoSheets);
    }

    // 3. Apresentações Google nativas
    if (mime === MimeType.GOOGLE_SLIDES) {
      const pres = SlidesApp.openById(fileId);
      const textoSlides = pres.getSlides().map((sl, i) =>
        '[Slide ' + (i + 1) + ']\n' +
        sl.getShapes().map(sh => sh.getText().asString()).join('\n')
      ).join('\n\n');
      return sanitizarTextoIndexado_(textoSlides);
    }

    // 4. Arquivos de texto simples ou Markdown
    if (mime === MimeType.PLAIN_TEXT || nome.endsWith('.md') || nome.endsWith('.txt') || nome.endsWith('.csv')) {
      return sanitizarTextoIndexado_(file.getBlob().getDataAsString('UTF-8'));
    }

    // 5. Arquivos Word modernos (.docx) -> Leitura direta do XML interno (rápido e imune a erros do DocumentApp)
    if (mime === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' || nome.endsWith('.docx')) {
      const textoDocx = extrairTextoDocxNativo_(file);
      if (textoDocx && textoDocx.trim().length > 10) {
        return sanitizarTextoIndexado_(textoDocx);
      }
    }

    // 6. Arquivos PDF (.pdf), Word antigo (.doc) e Excel (.xlsx/.xls) via Drive API v3 REST
    const isPdfOuDoc = mime === MimeType.PDF || mime === 'application/msword' ||
      nome.endsWith('.pdf') || nome.endsWith('.doc');

    const isExcel = mime === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
      mime === 'application/vnd.ms-excel' ||
      nome.endsWith('.xlsx') || nome.endsWith('.xls');

    if (isPdfOuDoc || isExcel) {
      const targetMime = isPdfOuDoc ? 'application/vnd.google-apps.document' : 'application/vnd.google-apps.spreadsheet';
      const token = ScriptApp.getOAuthToken();
      let tempId = null;

      try {
        const copyResp = UrlFetchApp.fetch(
          'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(fileId) + '/copy',
          {
            method: 'post',
            contentType: 'application/json',
            headers: { Authorization: 'Bearer ' + token },
            payload: JSON.stringify({
              name: 'TEMP_INDEX_' + file.getName(),
              mimeType: targetMime
            }),
            muteHttpExceptions: true
          }
        );

        if (copyResp.getResponseCode() === 200) {
          const copyData = JSON.parse(copyResp.getContentText());
          tempId = copyData && copyData.id;
        }

        if (tempId) {
          if (isPdfOuDoc) {
            const exportResp = UrlFetchApp.fetch(
              'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(tempId) + '/export?mimeType=text/plain',
              {
                method: 'get',
                headers: { Authorization: 'Bearer ' + token },
                muteHttpExceptions: true
              }
            );
            if (exportResp.getResponseCode() === 200) {
              return sanitizarTextoIndexado_(exportResp.getContentText('UTF-8'));
            }
          } else {
            const ssTemp = SpreadsheetApp.openById(tempId);
            const textoConvertido = ssTemp.getSheets().map(s =>
              '[' + s.getName() + ']\n' +
              s.getDataRange().getDisplayValues().map(r => r.filter(Boolean).join(' | ')).join('\n')
            ).join('\n\n');
            return sanitizarTextoIndexado_(textoConvertido);
          }
        }
      } finally {
        if (tempId) {
          try {
            DriveApp.getFileById(tempId).setTrashed(true);
          } catch (eDel) {}
        }
      }
    }
  } catch (e) {
    Logger.log('Falha ao extrair ' + file.getName() + ': ' + e.message);
  }

  return '';
}

function sanitizarTextoIndexado_(texto) {
  if (!texto) return '';
  return String(texto)
    .replace(/\r\n/g, '\n')
    .replace(/[ \t]+/g, ' ')
    .replace(/\n{3,}/g, '\n\n')
    .trim()
    .substring(0, MAX_CHARS_INDEXACAO);
}

function limparTituloArquivo_(nomeArquivo) {
  return String(nomeArquivo || '')
    .replace(/\.(docx?|xlsx?|pptx?|pdf|md|txt|csv)$/i, '')
    .trim();
}

function inferirTipoConteudo_(nomeArquivo, mimeType) {
  const n = String(nomeArquivo || '').toLowerCase();
  if (n.includes('checklist')) return 'Checklist';
  if (n.includes('fluxo') || mimeType.startsWith('image/')) return 'Fluxo';
  if (n.includes('manual') || n.includes('sumário') || n.includes('sumario')) return 'Manual';
  if (n.includes('planilha') || mimeType === MimeType.GOOGLE_SHEETS || n.endsWith('.xlsx') || n.endsWith('.xls')) return 'Planilha';
  if (n.includes('formulário') || n.includes('formulario') || n.includes('ficha') || n.includes('termo')) return 'Formulário';
  return 'Documento';
}

function gerarPalavrasChave_(titulo, tema, subtema) {
  const base = (titulo + ' ' + tema + ' ' + subtema)
    .toLowerCase()
    .replace(/[^\wÀ-ÿ\s]/g, ' ')
    .split(/\s+/)
    .filter(w => w.length >= 3);
  return Array.from(new Set(base)).slice(0, 10).join(', ');
}

function gerarResumoAutomatico_(textoIndexado, tituloLimpo) {
  if (textoIndexado && textoIndexado.length > 60) {
    const limpo = textoIndexado
      .replace(/^#+\s*/gm, '')
      .replace(/\s+/g, ' ')
      .trim();
    return limpo.substring(0, 260) + (limpo.length > 260 ? '...' : '');
  }
  return 'Material técnico: ' + tituloLimpo + '.';
}