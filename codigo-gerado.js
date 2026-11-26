/**
 * Ouvidoria — Yara Clube de Marília
 * Back-end de recebimento de manifestações.
 * Gera protocolo, grava na planilha, salva anexos no Drive e envia e-mail.
 */

// ===== CONFIGURAÇÃO (preencher antes de publicar) =====
const SPREADSHEET_ID   = 'https://drive.google.com/drive/my-drive';
const PASTA_ANEXOS_ID  = 'Manifestacoes';
const EMAIL_OUVIDORIA  = 'ouvidor.yaraclube@gmail.com';
const NOME_ABA         = 'Manifestacoes';

// ===== LIMITES =====
const LIMITE_ANEXO_BYTES = 10 * 1024 * 1024;  // 10 MB por arquivo
const LIMITE_TOTAL_BYTES = 25 * 1024 * 1024;  // 25 MB no total
const LIMITE_EMAIL_BYTES = 20 * 1024 * 1024;  // teto do e-mail

const CABECALHOS = [
  'PROTOCOLO', 'DATA_HORA_RECEBIMENTO', 'TIPO_MANIFESTACAO', 'ANONIMO',
  'NOME', 'SOBRENOME', 'CPF', 'EMAIL', 'TELEFONE', 'RELACAO_CLUBE',
  'ASSUNTO', 'LOCAL', 'LOCAL_ESPECIFICO', 'DATA_OCORRIDO',
  'HORARIO_INICIO', 'HORARIO_FIM', 'DESCRICAO', 'PESSOAS_ENVOLVIDAS',
  'TESTEMUNHAS', 'TIPOS_EVIDENCIA', 'PROVIDENCIAS', 'ASSINATURA',
  'TEXTO_LOCAL_DATA', 'ACEITE_VERACIDADE', 'ACEITE_CONFIDENCIALIDADE',
  'ANEXOS', 'LINKS_ANEXOS'
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('Requisição sem conteúdo.');
    }
    const dados = JSON.parse(e.postData.contents);
    validar(dados);

    const protocolo = gerarProtocolo();
    const anexos = salvarAnexos(protocolo, dados.anexos || []);
    registrarNaPlanilha(protocolo, dados, anexos);
    enviarEmail(protocolo, dados, anexos);

    return json({ ok: true, protocolo: protocolo });
  } catch (erro) {
    return json({ ok: false, erro: erro && erro.message ? erro.message : String(erro) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return json({ ok: true, servico: 'Ouvidoria Yara Clube', versao: '1.0' });
}

// ===== VALIDAÇÃO =====
function validar(d) {
  if (!d.tipoManifestacao) throw new Error('Tipo de manifestação não informado.');
  if (!d.assunto) throw new Error('Assunto principal não informado.');
  if (!d.local) throw new Error('Local do fato não informado.');
  if (!d.descricao || String(d.descricao).trim().length < 100) {
    throw new Error('A descrição deve ter no mínimo 100 caracteres.');
  }
  if (d.anonimo !== true) {
    if (!d.nome || !d.sobrenome) throw new Error('Nome e sobrenome são obrigatórios.');
    if (!d.email) throw new Error('E-mail é obrigatório.');
  }
  if (d.aceiteVeracidade !== true || d.aceiteConfidencialidade !== true) {
    throw new Error('As declarações de veracidade são obrigatórias.');
  }
  if (!d.assinatura) throw new Error('Assinatura do manifestante é obrigatória.');

  const lista = d.anexos || [];
  let total = 0;
  lista.forEach(function (a) {
    if (!a.nome || !a.base64) throw new Error('Anexo inválido: ' + (a.nome || 'sem nome'));
    if (a.tamanho > LIMITE_ANEXO_BYTES) {
      throw new Error('O arquivo ' + a.nome + ' excede 10 MB.');
    }
    total += (a.tamanho || 0);
  });
  if (total > LIMITE_TOTAL_BYTES) throw new Error('O total de anexos excede 25 MB.');
}

// ===== PROTOCOLO: OUV + AA + MM + sequência de 4 dígitos =====
function gerarProtocolo() {
  const agora = new Date();
  const aa = String(agora.getFullYear()).slice(-2);
  const mm = String(agora.getMonth() + 1).padStart(2, '0');
  const prefixo = 'OUV' + aa + mm;

  const aba = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(NOME_ABA);
  if (!aba) throw new Error('A aba ' + NOME_ABA + ' não existe na planilha.');

  const ultima = aba.getLastRow();
  let maior = 0;
  if (ultima > 1) {
    const valores = aba.getRange(2, 1, ultima - 1, 1).getValues();
    valores.forEach(function (linha) {
      const valor = String(linha[0] || '').trim().toUpperCase();
      if (valor.indexOf(prefixo) === 0) {
        const seq = parseInt(valor.slice(prefixo.length), 10);
        if (!isNaN(seq) && seq > maior) maior = seq;
      }
    });
  }
  return prefixo + String(maior + 1).padStart(4, '0');
}

// ===== ANEXOS NO DRIVE =====
function salvarAnexos(protocolo, lista) {
  if (!lista || lista.length === 0) return [];
  const pasta = DriveApp.getFolderById(PASTA_ANEXOS_ID);
  const subpasta = pasta.createFolder(protocolo);

  return lista.map(function (a) {
    const bytes = Utilities.base64Decode(a.base64);
    const blob = Utilities.newBlob(bytes, a.tipo || 'application/octet-stream', a.nome);
    const arquivo = subpasta.createFile(blob);
    return {
      nome: a.nome,
      tamanho: a.tamanho || 0,
      link: arquivo.getUrl()
    };
  });
}

// ===== PLANILHA =====
function registrarNaPlanilha(protocolo, d, anexos) {
  const aba = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(NOME_ABA);
  if (aba.getLastRow() === 0) {
    aba.appendRow(CABECALHOS);
    aba.getRange(1, 1, 1, CABECALHOS.length).setFontWeight('bold');
    aba.setFrozenRows(1);
  }
  const anonimo = d.anonimo === true;
  aba.appendRow([
    protocolo,
    Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm:ss'),
    d.tipoManifestacao || '',
    anonimo ? 'SIM' : 'NAO',
    anonimo ? '' : (d.nome || ''),
    anonimo ? '' : (d.sobrenome || ''),
    anonimo ? '' : (d.cpf || ''),
    anonimo ? '' : (d.email || ''),
    anonimo ? '' : (d.telefone || ''),
    d.relacao || '',
    d.assunto || '',
    d.local || '',
    d.localEspecifico || '',
    d.naoSeiData ? 'NÃO INFORMADA' : (d.dataOcorrido || ''),
    d.naoTenhoHorario ? 'NÃO INFORMADO' : (d.horarioInicio || ''),
    d.naoTenhoHorario ? 'NÃO INFORMADO' : (d.horarioFim || ''),
    d.descricao || '',
    d.pessoasEnvolvidas || '',
    d.testemunhas || '',
    (d.tiposEvidencia || []).join('; '),
    d.providencias || '',
    d.assinatura || '',
    d.localDataTexto || '',
    d.aceiteVeracidade === true ? 'SIM' : 'NAO',
    d.aceiteConfidencialidade === true ? 'SIM' : 'NAO',
    anexos.map(function (a) { return a.nome; }).join(' | '),
    anexos.map(function (a) { return a.link; }).join(' | ')
  ]);
}

// ===== E-MAIL =====
function enviarEmail(protocolo, d, anexos) {
  const assunto = 'PARA OUVIDORIA - ' + protocolo;
  const anonimo = d.anonimo === true;

  const linhas = [
    'MANIFESTAÇÃO RECEBIDA PELA OUVIDORIA',
    'Yara Clube de Marília',
    '',
    'PROTOCOLO: ' + protocolo,
    'RECEBIDA EM: ' + Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm'),
    '',
    '--- CLASSIFICAÇÃO ---',
    'Tipo de manifestação: ' + (d.tipoManifestacao || '-'),
    'Assunto principal: ' + (d.assunto || '-'),
    'Relação com o clube: ' + (d.relacao || '-'),
    'Local do fato: ' + (d.local || '-'),
    'Local específico: ' + (d.localEspecifico || '-'),
    'Data do ocorrido: ' + (d.naoSeiData ? 'não informada' : (d.dataOcorrido || '-')),
    'Horário: ' + (d.naoTenhoHorario ? 'não informado' : ((d.horarioInicio || '-') + ' a ' + (d.horarioFim || '-'))),
    '',
    '--- IDENTIFICAÇÃO ---',
    'Modo: ' + (anonimo ? 'ANÔNIMA' : 'IDENTIFICADA')
  ];

  if (!anonimo) {
    linhas.push(
      'Nome: ' + (d.nome || '') + ' ' + (d.sobrenome || ''),
      'CPF: ' + (d.cpf || '-'),
      'E-mail: ' + (d.email || '-'),
      'Telefone: ' + (d.telefone || '-')
    );
  }

  linhas.push(
    '',
    '--- RELATO ---',
    'Descrição detalhada:',
    (d.descricao || '-'),
    '',
    'Pessoas envolvidas ou citadas:',
    (d.pessoasEnvolvidas || 'não informado'),
    '',
    'Testemunhas:',
    (d.testemunhas || 'não informado'),
    '',
    'Providências já tomadas:',
    (d.providencias || 'não informado'),
    '',
    'Tipos de evidência declarados: ' + ((d.tiposEvidencia || []).join(', ') || 'nenhum'),
    '',
    '--- DECLARAÇÕES ---',
    'Veracidade: aceita',
    'Confidencialidade e não retaliação: aceita',
    'Assinatura: ' + (d.assinatura || '-'),
    (d.localDataTexto || ''),
    '',
    '--- ANEXOS (' + anexos.length + ') ---'
  );

  if (anexos.length === 0) {
    linhas.push('Nenhum anexo enviado.');
  } else {
    anexos.forEach(function (a) {
      linhas.push('- ' + a.nome + ' (' + formatarTamanho(a.tamanho) + ') — ' + a.link);
    });
  }

  linhas.push('', 'Mensagem gerada automaticamente pelo formulário web da Ouvidoria.');

  const corpo = linhas.join('\n');

  const arquivos = [];
  let acumulado = 0;
  (d.anexos || []).forEach(function (a) {
    if (acumulado + (a.tamanho || 0) <= LIMITE_EMAIL_BYTES) {
      arquivos.push(Utilities.newBlob(
        Utilities.base64Decode(a.base64),
        a.tipo || 'application/octet-stream',
        a.nome
      ));
      acumulado += (a.tamanho || 0);
    }
  });

  const opcoes = {
    to: EMAIL_OUVIDORIA,
    subject: assunto,
    body: corpo
  };
  if (arquivos.length > 0) opcoes.attachments = arquivos;
  if (arquivos.length < (d.anexos || []).length) {
    opcoes.body = corpo +
      '\n\nObservação: parte dos anexos excedeu o limite do e-mail e está disponível apenas pelos links acima.';
  }

  MailApp.sendEmail(opcoes);
}

function formatarTamanho(bytes) {
  if (!bytes) return '0 KB';
  if (bytes < 1024 * 1024) return Math.round(bytes / 1024) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(1).replace('.', ',') + ' MB';
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== UTILITÁRIO: criar cabeçalhos manualmente =====
function configurarPlanilha() {
  const aba = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(NOME_ABA);
  aba.clear();
  aba.appendRow(CABECALHOS);
  aba.getRange(1, 1, 1, CABECALHOS.length).setFontWeight('bold');
  aba.setFrozenRows(1);
}
