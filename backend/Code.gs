/**
 * Backend da confirmação de presença da Alice (Google Apps Script).
 *
 * A cada confirmação enviada pelo site (nomes da família e um WhatsApp):
 *  1. Grava uma linha por pessoa na planilha (Data, Nome, WhatsApp, Status), antes de qualquer envio.
 *  2. Avisa a organização (MAE_PHONE ou ORGANIZER_PHONE) e manda a confirmação ao WhatsApp informado.
 *  3. Atualiza a coluna Status dessas linhas com o resultado dos envios.
 *
 * Também recebe o webhook de mensagens da Z-API (formato Z_API) e grava na aba Mensagens.
 *
 * Configuração: Projeto > Configurações > Propriedades do script.
 * No Apps Script, as propriedades do script fazem o papel das variáveis de ambiente.
 * Nunca coloque esses valores no código.
 *
 *  SHEET_ID              ID da planilha (trecho entre /d/ e /edit na URL)
 *  SHEET_NAME            nome da aba, padrão: Respostas
 *  MAE_PHONE             WhatsApp que recebe o aviso, com DDI: 5521999999999 (se faltar, usa ORGANIZER_PHONE)
 *  WHATSAPP_PROVIDER     'omni' (padrão), 'zapi' ou 'meta'
 *
 *  Omni Z-API (padrão):
 *  OMNI_CHANNEL_ID       ID do canal (painel da Omni > Conexões)
 *  OMNI_SECRET_KEY       Secret Key (painel da Omni > Segurança)
 *  TEMPLATE_CONVIDADO    (opcional) nome do modelo aprovado para o convidado: {{primeiro_nome}} e {{nomes}}
 *  TEMPLATE_AVISO        (opcional) nome do modelo aprovado para o aviso: {{nomes}}
 *  TEMPLATE_LANG         (opcional) idioma do modelo, padrão pt_BR
 *
 *  Z-API por instância (alternativa): ZAPI_BASE_URL, ZAPI_INSTANCE_ID, ZAPI_TOKEN, ZAPI_CLIENT_TOKEN
 *  Meta direto (legado): WA_PHONE_NUMBER_ID, WA_TOKEN, TEMPLATE_ORGANIZER, TEMPLATE_LANG
 *  Webhook (opcionais): ZAPI_WEBHOOK_SECRET, MESSAGES_SHEET_NAME
 *  WEBHOOK_URL           URL /exec desta implantação (só para criarWebhookOmni)
 */

// ===== TEXTOS (edite aqui) =====
// {nome} = primeiro nome da lista. {nomes} = todos os nomes ("Ana Silva, João Silva e Bia Silva").
var TEXTO_CONFIRMACAO = 'Presença confirmada na festa de 15 anos da Alice: {nomes}.';
var TEXTO_CONVIDADO = 'Olá, {nome}! Presença confirmada na festa de 15 anos da Alice para: {nomes}.\n' +
  'Sexta, 11 de dezembro de 2026, às 20h30, no Sítio Monte Cristo ' +
  '(Rua Monte Cristo, s/n, Pau Grande).\nTraje: esporte fino. Até lá!';
var MAX_PESSOAS = 10;
// ===== FIM DOS TEXTOS =====

var OMNI_API = 'https://api.omni.z-api.io/v1/channels/';
var GRAPH_VERSION = 'v21.0';

function doPost(e) {
  var props = PropertiesService.getScriptProperties();
  var dados;
  try {
    dados = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, erro: 'Dados inválidos' });
  }

  // Webhook de mensagem recebida da Z-API: não é um formulário de presença.
  if (dados.type === 'ReceivedCallback' || (dados.phone !== undefined && dados.text !== undefined && dados.fromMe !== undefined)) {
    return receberMensagemZapi_(e, dados);
  }

  // Nomes da família. Aceita a lista nova (nomes) e o campo antigo (nome).
  var bruto = Array.isArray(dados.nomes) ? dados.nomes : [dados.nome];
  var nomes = bruto
    .map(function (n) { return String(n || '').trim().slice(0, 80); })
    .filter(function (n) { return n; });
  if (!nomes.length) {
    return json_({ ok: false, erro: 'Nome é obrigatório' });
  }
  if (nomes.length > MAX_PESSOAS) {
    return json_({ ok: false, erro: 'Máximo de ' + MAX_PESSOAS + ' pessoas por confirmação' });
  }
  // WhatsApp de quem confirma: 11 números (DDD + celular), sem +55. O 55 é acrescentado aqui.
  var digitos = String(dados.whatsapp || '').replace(/\D/g, '');
  if (digitos.length !== 11) {
    return json_({ ok: false, erro: 'Informe o WhatsApp com 11 números (DDD + celular)' });
  }
  var whatsapp = '55' + digitos;

  // 1. Salva primeiro, uma linha por pessoa. Se o envio falhar, as linhas continuam na planilha.
  var aba = SpreadsheetApp.openById(props.getProperty('SHEET_ID'))
    .getSheetByName(props.getProperty('SHEET_NAME') || 'Respostas');
  var agora = new Date();
  var linhas = nomes.map(function (n) {
    aba.appendRow([agora, n, whatsapp, 'Salvo']);
    return aba.getLastRow();
  });

  // 2. Avisa a organização e confirma no WhatsApp informado. Uma falha não impede a outra.
  // split/join e não replace: em replace, "$&" no nome vira trecho do modelo.
  var lista = juntarNomes_(nomes);
  var primeiro = nomes[0].split(/\s+/)[0];
  // Modelos aprovados (opcionais), com variáveis por nome, como o painel da Omni cria:
  // TEMPLATE_AVISO usa {{nomes}}; TEMPLATE_CONVIDADO usa {{primeiro_nome}} e {{nomes}}. Sem eles, vai texto livre.
  var destinos = [
    ['mae', props.getProperty('MAE_PHONE') || props.getProperty('ORGANIZER_PHONE'), TEXTO_CONFIRMACAO,
      { nome: props.getProperty('TEMPLATE_AVISO'), parametros: { nomes: lista } }],
    ['convidado', whatsapp, TEXTO_CONVIDADO,
      { nome: props.getProperty('TEMPLATE_CONVIDADO'), parametros: { primeiro_nome: primeiro, nomes: lista } }]
  ];
  var falhas = [];
  destinos.forEach(function (destino) {
    var texto = destino[2].split('{nomes}').join(lista).split('{nome}').join(primeiro);
    try {
      enviarAviso_(destino[1], texto, lista, destino[3]);
    } catch (err) {
      falhas.push(destino[0] + ': ' + err.message);
    }
  });

  // 3. Atualiza a coluna Status (a 4ª) das linhas desta confirmação.
  var status = falhas.length ? 'Falha no envio: ' + falhas.join('; ') : 'Enviado';
  linhas.forEach(function (linha) { aba.getRange(linha, 4).setValue(status); });

  return json_({ ok: true, status: status });
}

/**
 * Recebe uma mensagem do WhatsApp pelo webhook da Z-API e registra na aba de mensagens.
 * Não responde automaticamente. Se ZAPI_WEBHOOK_SECRET estiver definido, a URL do
 * webhook precisa trazer ?secret= com esse mesmo valor.
 */
function receberMensagemZapi_(e, dados) {
  var props = PropertiesService.getScriptProperties();
  var segredo = props.getProperty('ZAPI_WEBHOOK_SECRET');
  if (segredo && (!e.parameter || e.parameter.secret !== segredo)) {
    return json_({ ok: false, erro: 'Não autorizado' });
  }

  // Ignora mensagens enviadas pela própria instância e de grupos.
  if (dados.fromMe || dados.isGroup) {
    return json_({ ok: true, ignorado: true });
  }

  var texto = dados.text && dados.text.message
    ? String(dados.text.message)
    : '[' + (dados.type || 'mensagem sem texto') + ']';
  var quando = dados.momment || dados.moment ? new Date(Number(dados.momment || dados.moment)) : new Date();

  obterAbaMensagens_().appendRow([quando, String(dados.phone || ''), dados.senderName || '', texto]);
  return json_({ ok: true, status: 'Registrada' });
}

/** Aba das mensagens recebidas (MESSAGES_SHEET_NAME, padrão "Mensagens"); cria se não existir. */
function obterAbaMensagens_() {
  var props = PropertiesService.getScriptProperties();
  var planilha = SpreadsheetApp.openById(props.getProperty('SHEET_ID'));
  var nome = props.getProperty('MESSAGES_SHEET_NAME') || 'Mensagens';
  var aba = planilha.getSheetByName(nome);
  if (!aba) {
    aba = planilha.insertSheet(nome);
    aba.appendRow(['Recebida em', 'Telefone', 'Nome', 'Mensagem']);
  }
  return aba;
}

/** Envia a confirmação pelo provedor configurado. Lança erro se o envio falhar. */
function enviarAviso_(telefone, texto, nome, modelo) {
  var props = PropertiesService.getScriptProperties();
  var provedor = (props.getProperty('WHATSAPP_PROVIDER') || 'omni').toLowerCase();
  var numero = String(telefone || '').replace(/\D/g, '');

  if (provedor === 'omni') {
    // Com modelo aprovado configurado, usa o modelo (dispensa a janela de 24 horas).
    if (modelo && modelo.nome) return enviarModeloOmni_(numero, modelo.nome, modelo.parametros);
    return enviarTextoOmni_(numero, texto);
  }
  if (provedor === 'zapi') return enviarTextoZapi_(numero, texto);
  if (provedor === 'meta') return enviarModelo_(numero, props.getProperty('TEMPLATE_ORGANIZER'), [nome]);
  throw new Error('WHATSAPP_PROVIDER desconhecido: ' + provedor);
}

/**
 * Envia texto pela Omni Z-API: POST /v1/channels/{channelId}/messages, com Bearer.
 * No canal da Meta, texto livre só é entregue com a janela de 24 horas aberta.
 * Lança erro se a API recusar.
 */
function enviarTextoOmni_(telefone, mensagem) {
  return chamarOmni_(telefone, { type: 'TEXT', body: { message: mensagem } });
}

/**
 * Envia um modelo aprovado pela Meta pela Omni Z-API (funciona fora da janela de 24 horas).
 * parametros: { nome_da_variavel: valor }, com os mesmos nomes que estão entre {{ }} no modelo.
 */
function enviarModeloOmni_(telefone, nomeModelo, parametros) {
  var idioma = PropertiesService.getScriptProperties().getProperty('TEMPLATE_LANG') || 'pt_BR';
  return chamarOmni_(telefone, {
    type: 'TEMPLATE',
    attachments: [{
      template: {
        name: nomeModelo,
        language: { policy: 'deterministic', code: idioma },
        components: [{
          type: 'body',
          parameters: Object.keys(parametros || {}).map(function (nome) {
            return { type: 'text', parameter_name: nome, text: String(parametros[nome]) };
          })
        }]
      }
    }]
  });
}

/** POST de uma mensagem na Omni Z-API. Lança erro se a API recusar. */
function chamarOmni_(telefone, conteudo) {
  if (!telefone) throw new Error('telefone do destinatário ausente');

  var props = PropertiesService.getScriptProperties();
  var canal = props.getProperty('OMNI_CHANNEL_ID');
  var chave = props.getProperty('OMNI_SECRET_KEY');
  if (!canal || !chave) {
    throw new Error('Omni sem configuração (OMNI_CHANNEL_ID ou OMNI_SECRET_KEY)');
  }

  var resposta = UrlFetchApp.fetch(OMNI_API + encodeURIComponent(canal) + '/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + chave },
    payload: JSON.stringify({ recipient: { identifier: telefone }, content: conteudo }),
    muteHttpExceptions: true
  });

  // Não inclua a chave no erro.
  var codigo = resposta.getResponseCode();
  if (codigo >= 300) {
    var corpo = {};
    try { corpo = JSON.parse(resposta.getContentText()) || {}; } catch (err) { corpo = {}; }
    // A documentação cita o subCode (ex.: CHANNEL_DISCONNECTED no 422).
    throw new Error('Omni respondeu ' + codigo + (corpo.subCode ? ' (' + corpo.subCode + ')' : ''));
  }
}

/** Envia uma mensagem de texto pela Z-API por instância. Lança erro se a API recusar. */
function enviarTextoZapi_(telefone, mensagem) {
  if (!telefone) throw new Error('telefone do destinatário ausente');

  var props = PropertiesService.getScriptProperties();
  var base = props.getProperty('ZAPI_BASE_URL');
  var instancia = props.getProperty('ZAPI_INSTANCE_ID');
  var token = props.getProperty('ZAPI_TOKEN');
  if (!base || !instancia || !token) {
    throw new Error('Z-API sem configuração (ZAPI_BASE_URL, ZAPI_INSTANCE_ID ou ZAPI_TOKEN)');
  }

  var url = base.replace(/\/+$/, '') + '/instances/' + encodeURIComponent(instancia) +
    '/token/' + encodeURIComponent(token) + '/send-text';

  var cabecalhos = {};
  var tokenDaConta = props.getProperty('ZAPI_CLIENT_TOKEN');
  if (tokenDaConta) cabecalhos['Client-Token'] = tokenDaConta;

  var resposta = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: cabecalhos,
    payload: JSON.stringify({ phone: telefone, message: mensagem }),
    muteHttpExceptions: true
  });

  // Não inclua a URL no erro: ela leva o token da instância.
  var codigo = resposta.getResponseCode();
  if (codigo >= 300) throw new Error('Z-API respondeu ' + codigo);

  var corpo = {};
  try { corpo = JSON.parse(resposta.getContentText()) || {}; } catch (err) { corpo = {}; }
  if (corpo.error) throw new Error('Z-API: ' + corpo.error);
}

/** Envia um modelo aprovado pela Meta (provedor legado). Lança erro se a API recusar. */
function enviarModelo_(telefone, nomeModelo, parametros) {
  if (!telefone) throw new Error('telefone do destinatário ausente');
  if (!nomeModelo) throw new Error('nome do modelo não configurado');

  var props = PropertiesService.getScriptProperties();
  var url = 'https://graph.facebook.com/' + GRAPH_VERSION + '/' +
    props.getProperty('WA_PHONE_NUMBER_ID') + '/messages';

  var corpo = {
    messaging_product: 'whatsapp',
    to: telefone,
    type: 'template',
    template: {
      name: nomeModelo,
      language: { code: props.getProperty('TEMPLATE_LANG') || 'pt_BR' },
      components: [{
        type: 'body',
        parameters: parametros.map(function (texto) {
          return { type: 'text', text: String(texto) };
        })
      }]
    }
  };

  var resposta = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + props.getProperty('WA_TOKEN') },
    payload: JSON.stringify(corpo),
    muteHttpExceptions: true
  });

  if (resposta.getResponseCode() >= 300) {
    throw new Error('WhatsApp ' + resposta.getResponseCode());
  }
}

/**
 * Cria o webhook do canal na Omni, apontando para este Apps Script, no formato Z_API.
 * Rode UMA vez pelo editor: escolha criarWebhookOmni no menu e clique em Executar.
 * Rodar de novo cria outro webhook igual, e cada mensagem passa a ser gravada duas vezes.
 * Usa OMNI_CHANNEL_ID, OMNI_SECRET_KEY e WEBHOOK_URL. Com ZAPI_WEBHOOK_SECRET, a URL vai com ?secret=.
 */
function criarWebhookOmni() {
  var props = PropertiesService.getScriptProperties();
  var canal = props.getProperty('OMNI_CHANNEL_ID');
  var chave = props.getProperty('OMNI_SECRET_KEY');
  var url = props.getProperty('WEBHOOK_URL');

  var existente = props.getProperty('OMNI_WEBHOOK_ID');
  if (existente) {
    throw new Error('Webhook já criado (id ' + existente + '). Só apague OMNI_WEBHOOK_ID se removeu esse webhook no painel.');
  }

  if (!canal || !chave || !url) {
    throw new Error('Faltam propriedades: OMNI_CHANNEL_ID, OMNI_SECRET_KEY ou WEBHOOK_URL');
  }

  var segredo = props.getProperty('ZAPI_WEBHOOK_SECRET');
  if (segredo) url += (url.indexOf('?') === -1 ? '?' : '&') + 'secret=' + encodeURIComponent(segredo);

  var resposta = UrlFetchApp.fetch(OMNI_API + encodeURIComponent(canal) + '/webhooks', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + chave },
    payload: JSON.stringify({
      url: url,
      events: ['MESSAGE_RECEIVED'],
      payloadFormat: 'Z_API',
      description: 'Confirmação de presença - planilha'
    }),
    muteHttpExceptions: true
  });

  var corpo = {};
  try { corpo = JSON.parse(resposta.getContentText()) || {}; } catch (err) { corpo = {}; }
  var codigo = resposta.getResponseCode();
  if (codigo >= 300) {
    throw new Error('Omni respondeu ' + codigo + (corpo.error ? ': ' + corpo.error : ''));
  }

  if (!corpo.id) {
    Logger.log('A Omni respondeu ' + codigo + ' sem id. Confira no painel se o webhook foi criado ' +
      'e, se foi, preencha OMNI_WEBHOOK_ID com o id dele antes de rodar de novo.');
    return null;
  }

  props.setProperty('OMNI_WEBHOOK_ID', String(corpo.id));
  Logger.log('Webhook criado: id ' + corpo.id + ', formato ' + corpo.payloadFormat + ', status ' + corpo.status);
  return corpo.id;
}

/** ["Ana Silva", "João Silva", "Bia Silva"] -> "Ana Silva, João Silva e Bia Silva" */
function juntarNomes_(nomes) {
  if (nomes.length < 2) return nomes.join('');
  return nomes.slice(0, -1).join(', ') + ' e ' + nomes[nomes.length - 1];
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
