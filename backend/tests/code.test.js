// Testes do backend/Code.gs (Google Apps Script).
//
// Como rodar (na pasta rsvp-alice-15):
//   node --test backend/tests/code.test.js
//
// Não há dependências: usa o teste embutido do Node (node:test) e o módulo vm.
// O Code.gs é carregado num ambiente isolado, e as APIs do Google (planilha,
// propriedades, UrlFetchApp) são substituídas por simulações.

const { test, describe } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const fonte = fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8');

// O Code.gs roda em outro contexto (vm), então arrays e objetos vêm de outro "realm".
// Esta função compara o conteúdo, e não o tipo de origem.
function deepEq(real, esperado) {
  assert.deepEqual(JSON.parse(JSON.stringify(real)), esperado);
}

// Cria um ambiente com o Code.gs carregado.
// respostaApi(url, opcoes, numeroDaChamada) devolve { codigo, corpo } para simular a API.
function criarAmbiente({ props = {}, respostaApi = () => ({ codigo: 200, corpo: '{}' }) } = {}) {
  const propriedades = { ...props };
  const abas = {};
  const chamadas = [];
  const logs = [];

  function criarAba(nome) {
    const aba = {
      nome,
      linhas: [],
      appendRow(linha) { this.linhas.push(linha); },
      getLastRow() { return this.linhas.length; },
      getRange(linha, coluna) {
        return { setValue: (valor) => { aba.linhas[linha - 1][coluna - 1] = valor; } };
      }
    };
    abas[nome] = aba;
    return aba;
  }

  const sandbox = {
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (chave) => (chave in propriedades ? propriedades[chave] : null),
        setProperty: (chave, valor) => { propriedades[chave] = valor; }
      })
    },
    SpreadsheetApp: {
      openById: () => ({
        getSheetByName: (nome) => abas[nome] || null,
        insertSheet: (nome) => criarAba(nome)
      })
    },
    ContentService: {
      createTextOutput: (texto) => ({ conteudo: texto, setMimeType() { return this; } }),
      MimeType: { JSON: 'application/json' }
    },
    UrlFetchApp: {
      fetch: (url, opcoes) => {
        chamadas.push({ url, opcoes });
        const r = respostaApi(url, opcoes, chamadas.length);
        return { getResponseCode: () => r.codigo, getContentText: () => r.corpo };
      }
    },
    Logger: { log: (texto) => logs.push(texto) },
    // Tipos básicos do JavaScript, usados pelo Code.gs
    JSON, String, Number, Math, Array, Object, Error, Date, parseInt, isNaN, encodeURIComponent
  };

  vm.createContext(sandbox);
  vm.runInContext(fonte, sandbox);

  return { g: sandbox, props: propriedades, abas, criarAba, chamadas, logs };
}

// Ambiente padrão do formulário: planilha "Respostas", dois números e Omni configurada.
function ambienteFormulario(opcoes = {}) {
  const amb = criarAmbiente({
    ...opcoes,
    props: {
      SHEET_ID: 'PLANILHA',
      SHEET_NAME: 'Respostas',
      MAE_PHONE: '5521900000001',
      OMNI_CHANNEL_ID: 'CANAL',
      OMNI_SECRET_KEY: 'CHAVE123',
      ...(opcoes.props || {})
    }
  });
  amb.criarAba('Respostas');
  return amb;
}

// Envia um objeto (formulário) ao doPost e devolve a resposta em JSON.
function enviarFormulario(amb, dados) {
  return JSON.parse(amb.g.doPost({ postData: { contents: JSON.stringify(dados) } }).conteudo);
}

// Envia um texto bruto ao doPost (para testar JSON inválido e webhooks).
function enviarTexto(amb, texto, parametros) {
  const e = { postData: { contents: texto } };
  if (parametros) e.parameter = parametros;
  return JSON.parse(amb.g.doPost(e).conteudo);
}

const corpoDe = (chamada) => JSON.parse(chamada.opcoes.payload);

const mensagemZapi = (sobrescrever = {}) => JSON.stringify({
  type: 'ReceivedCallback',
  phone: '5521999990000',
  fromMe: false,
  isGroup: false,
  senderName: 'Maria',
  text: { message: 'Oi, vou sim!' },
  momment: 1700000000000,
  ...sobrescrever
});

// A primeira chamada (mãe) falha; a segunda (convidado) passa.
const maeFalha = (url, opcoes, n) => (n === 1 ? { codigo: 500, corpo: '' } : { codigo: 200, corpo: '{}' });

// ===================================================================
describe('doPost - formulário (nome e WhatsApp)', () => {
  test('grava data, nome e status Enviado', () => {
    const amb = ambienteFormulario();
    const r = enviarFormulario(amb, { nome: 'Ana', whatsapp: '(21) 90000-0002' });
    deepEq(r, { ok: true, status: 'Enviado' });
    const linha = amb.abas['Respostas'].linhas[0];
    deepEq(linha.slice(1), ['Ana', '5521900000002', 'Enviado']);
    assert.ok(linha[0] instanceof amb.g.Date);
  });

  test('avisa a organização e depois confirma ao convidado, cada um com seu texto', () => {
    const amb = ambienteFormulario();
    enviarFormulario(amb, { nome: 'Ana', whatsapp: '(21) 90000-0002' });
    assert.equal(amb.chamadas.length, 2);
    assert.equal(corpoDe(amb.chamadas[0]).recipient.identifier, '5521900000001');
    assert.equal(corpoDe(amb.chamadas[1]).recipient.identifier, '5521900000002');
    assert.equal(
      corpoDe(amb.chamadas[0]).content.body.message,
      'Presença confirmada na festa de 15 anos da Alice: Ana.'
    );
    assert.match(corpoDe(amb.chamadas[1]).content.body.message, /^Olá, Ana! Presença confirmada na festa de 15 anos da Alice para: Ana\./);
  });

  test('família: uma linha por pessoa, mesmo WhatsApp e status nas três', () => {
    const amb = ambienteFormulario();
    const r = enviarFormulario(amb, { nomes: ['Ana Silva', ' João Silva ', '', 'Bia Silva'], whatsapp: '(21) 90000-0002' });
    assert.equal(r.ok, true);
    const linhas = amb.abas['Respostas'].linhas;
    assert.equal(linhas.length, 3);
    deepEq(linhas.map((l) => l.slice(1)), [
      ['Ana Silva', '5521900000002', 'Enviado'],
      ['João Silva', '5521900000002', 'Enviado'],
      ['Bia Silva', '5521900000002', 'Enviado']
    ]);
    assert.equal(amb.chamadas.length, 2, 'uma mensagem para a organização e uma para o WhatsApp informado');
  });

  test('família: textos com a lista de nomes e o primeiro nome', () => {
    const amb = ambienteFormulario();
    enviarFormulario(amb, { nomes: ['Ana Silva', 'João Silva', 'Bia Silva'], whatsapp: '(21) 90000-0002' });
    assert.equal(corpoDe(amb.chamadas[0]).content.body.message,
      'Presença confirmada na festa de 15 anos da Alice: Ana Silva, João Silva e Bia Silva.');
    assert.match(corpoDe(amb.chamadas[1]).content.body.message,
      /^Olá, Ana! Presença confirmada na festa de 15 anos da Alice para: Ana Silva, João Silva e Bia Silva\./);
  });

  test('família: falha no envio vai para o status de todas as linhas', () => {
    const amb = ambienteFormulario({ respostaApi: maeFalha });
    enviarFormulario(amb, { nomes: ['Ana Silva', 'João Silva'], whatsapp: '(21) 90000-0002' });
    deepEq(amb.abas['Respostas'].linhas.map((l) => l[3]), [
      'Falha no envio: mae: Omni respondeu 500',
      'Falha no envio: mae: Omni respondeu 500'
    ]);
  });

  test('família: lista vazia é recusada e mais de 10 pessoas também', () => {
    const amb = ambienteFormulario();
    deepEq(enviarFormulario(amb, { nomes: ['', '  '], whatsapp: '21900000002' }), { ok: false, erro: 'Nome é obrigatório' });
    const onze = Array.from({ length: 11 }, (_, i) => 'Pessoa ' + i);
    deepEq(enviarFormulario(amb, { nomes: onze, whatsapp: '21900000002' }), { ok: false, erro: 'Máximo de 10 pessoas por confirmação' });
    assert.equal(amb.abas['Respostas'].linhas.length, 0);
    assert.equal(amb.chamadas.length, 0);
  });

  test('recusa nome vazio, sem gravar nem enviar', () => {
    const amb = ambienteFormulario();
    deepEq(enviarFormulario(amb, { nome: '   ' }), { ok: false, erro: 'Nome é obrigatório' });
    assert.equal(amb.abas['Respostas'].linhas.length, 0);
    assert.equal(amb.chamadas.length, 0);
  });

  test('corta o nome em 80 caracteres', () => {
    const amb = ambienteFormulario();
    enviarFormulario(amb, { nome: 'x'.repeat(120), whatsapp: '(21) 90000-0002' });
    assert.equal(amb.abas['Respostas'].linhas[0][1].length, 80);
  });

  test('preserva símbolos "$" do nome na mensagem', () => {
    const amb = ambienteFormulario();
    enviarFormulario(amb, { nome: 'Ana $& Bia', whatsapp: '(21) 90000-0002' });
    assert.equal(corpoDe(amb.chamadas[0]).content.body.message, 'Presença confirmada na festa de 15 anos da Alice: Ana $& Bia.');
  });

  test('ignora campos antigos do formulário (resposta, acompanhantes)', () => {
    const amb = ambienteFormulario();
    const r = enviarFormulario(amb, { nome: 'Ana', resposta: 'nao', whatsapp: '21999999999', acompanhantes: 'Bia' });
    assert.equal(r.ok, true);
    deepEq(amb.abas['Respostas'].linhas[0].slice(1), ['Ana', '5521999999999', 'Enviado']);
    assert.equal(amb.chamadas.length, 2);
  });

  test('sem MAE_PHONE, usa ORGANIZER_PHONE', () => {
    const amb = ambienteFormulario({ props: { MAE_PHONE: undefined, ORGANIZER_PHONE: '5521993103830' } });
    delete amb.props.MAE_PHONE;
    enviarFormulario(amb, { nome: 'Ana', whatsapp: '(21) 90000-0002' });
    assert.equal(corpoDe(amb.chamadas[0]).recipient.identifier, '5521993103830');
  });

  test('número com máscara nas propriedades é enviado só com dígitos', () => {
    const amb = ambienteFormulario({ props: { MAE_PHONE: '+55 (21) 90000-0001' } });
    enviarFormulario(amb, { nome: 'Ana', whatsapp: '(21) 90000-0002' });
    assert.equal(corpoDe(amb.chamadas[0]).recipient.identifier, '5521900000001');
  });

  test('recusa WhatsApp sem 11 números (vazio, incompleto ou com +55), sem gravar nem enviar', () => {
    const amb = ambienteFormulario();
    const erro = { ok: false, erro: 'Informe o WhatsApp com 11 números (DDD + celular)' };
    deepEq(enviarFormulario(amb, { nome: 'Ana' }), erro);
    deepEq(enviarFormulario(amb, { nome: 'Ana', whatsapp: '2190000000' }), erro);
    deepEq(enviarFormulario(amb, { nome: 'Ana', whatsapp: '+55 21 90000-0002' }), erro);
    assert.equal(amb.abas['Respostas'].linhas.length, 0);
    assert.equal(amb.chamadas.length, 0);
  });

  test('WhatsApp com máscara vira só números, com 55 na frente', () => {
    const amb = ambienteFormulario();
    enviarFormulario(amb, { nome: 'Ana', whatsapp: ' (21) 9 0000-0002 ' });
    assert.equal(amb.abas['Respostas'].linhas[0][2], '5521900000002');
  });

  test('falha no envio à mãe: a linha fica salva e o convidado ainda recebe', () => {
    const amb = ambienteFormulario({ respostaApi: maeFalha });
    const r = enviarFormulario(amb, { nome: 'Ana', whatsapp: '(21) 90000-0002' });
    deepEq(r, { ok: true, status: 'Falha no envio: mae: Omni respondeu 500' });
    assert.equal(amb.abas['Respostas'].linhas[0][3], 'Falha no envio: mae: Omni respondeu 500');
    assert.equal(amb.chamadas.length, 2);
  });

  test('falha nos dois envios aparece no status, separada por ";"', () => {
    const amb = ambienteFormulario({ respostaApi: () => ({ codigo: 401, corpo: '' }) });
    const r = enviarFormulario(amb, { nome: 'Ana', whatsapp: '(21) 90000-0002' });
    assert.equal(r.status, 'Falha no envio: mae: Omni respondeu 401; convidado: Omni respondeu 401');
  });

  test('usa a aba "Respostas" quando SHEET_NAME não está definido', () => {
    const amb = ambienteFormulario();
    delete amb.props.SHEET_NAME;
    enviarFormulario(amb, { nome: 'Ana', whatsapp: '(21) 90000-0002' });
    assert.equal(amb.abas['Respostas'].linhas.length, 1);
  });
});

// ===================================================================
describe('envio pela Omni Z-API (provedor padrão)', () => {
  test('POST no canal, com Bearer e corpo recipient/content do tipo TEXT', () => {
    const amb = ambienteFormulario();
    enviarFormulario(amb, { nome: 'Ana', whatsapp: '(21) 90000-0002' });
    const chamada = amb.chamadas[0];
    assert.equal(chamada.url, 'https://api.omni.z-api.io/v1/channels/CANAL/messages');
    assert.equal(chamada.opcoes.method, 'post');
    assert.equal(chamada.opcoes.contentType, 'application/json');
    assert.equal(chamada.opcoes.headers.Authorization, 'Bearer CHAVE123');
    deepEq(corpoDe(chamada), {
      recipient: { identifier: '5521900000001' },
      content: { type: 'TEXT', body: { message: 'Presença confirmada na festa de 15 anos da Alice: Ana.' } }
    });
  });

  test('422 com subCode mostra o subCode no status', () => {
    const amb = ambienteFormulario({
      respostaApi: () => ({ codigo: 422, corpo: '{"subCode":"CHANNEL_DISCONNECTED"}' })
    });
    const r = enviarFormulario(amb, { nome: 'Ana', whatsapp: '(21) 90000-0002' });
    assert.match(r.status, /mae: Omni respondeu 422 \(CHANNEL_DISCONNECTED\)/);
  });

  test('corpo de erro que não é JSON não derruba o envio', () => {
    const amb = ambienteFormulario({ respostaApi: () => ({ codigo: 502, corpo: '<html>erro</html>' }) });
    const r = enviarFormulario(amb, { nome: 'Ana', whatsapp: '(21) 90000-0002' });
    assert.equal(r.ok, true);
    assert.match(r.status, /mae: Omni respondeu 502/);
  });

  test('sem canal ou chave, o erro diz o que falta e nada é chamado', () => {
    const amb = ambienteFormulario();
    delete amb.props.OMNI_SECRET_KEY;
    const r = enviarFormulario(amb, { nome: 'Ana', whatsapp: '(21) 90000-0002' });
    assert.match(r.status, /Omni sem configuração \(OMNI_CHANNEL_ID ou OMNI_SECRET_KEY\)/);
    assert.equal(amb.chamadas.length, 0);
  });

  test('a Secret Key não aparece no status gravado na planilha', () => {
    const amb = ambienteFormulario({ respostaApi: () => ({ codigo: 500, corpo: '' }) });
    enviarFormulario(amb, { nome: 'Ana', whatsapp: '(21) 90000-0002' });
    assert.equal(amb.abas['Respostas'].linhas[0][3].includes('CHAVE123'), false);
  });
});

// ===================================================================
describe('modelos aprovados na Omni (sem a janela de 24 horas)', () => {
  test('com TEMPLATE_CONVIDADO e TEMPLATE_AVISO, envia modelos com as variáveis por nome', () => {
    const amb = ambienteFormulario({ props: { TEMPLATE_AVISO: 'aviso_alice', TEMPLATE_CONVIDADO: 'confirmacao_alice' } });
    const r = enviarFormulario(amb, { nomes: ['Ana Silva', 'João Silva'], whatsapp: '(21) 90000-0002' });
    assert.equal(r.status, 'Enviado');
    deepEq(corpoDe(amb.chamadas[0]), {
      recipient: { identifier: '5521900000001' },
      content: {
        type: 'TEMPLATE',
        attachments: [{
          template: {
            name: 'aviso_alice',
            language: { policy: 'deterministic', code: 'pt_BR' },
            components: [{ type: 'body', parameters: [{ type: 'text', parameter_name: 'nomes', text: 'Ana Silva e João Silva' }] }]
          }
        }]
      }
    });
    const conv = corpoDe(amb.chamadas[1]).content.attachments[0].template;
    assert.equal(conv.name, 'confirmacao_alice');
    deepEq(conv.components[0].parameters, [
      { type: 'text', parameter_name: 'primeiro_nome', text: 'Ana' },
      { type: 'text', parameter_name: 'nomes', text: 'Ana Silva e João Silva' }
    ]);
  });

  test('só um modelo configurado: ele vai como modelo e o outro como texto livre', () => {
    const amb = ambienteFormulario({ props: { TEMPLATE_CONVIDADO: 'confirmacao_alice' } });
    enviarFormulario(amb, { nomes: ['Ana Silva'], whatsapp: '(21) 90000-0002' });
    assert.equal(corpoDe(amb.chamadas[0]).content.type, 'TEXT');
    assert.equal(corpoDe(amb.chamadas[1]).content.type, 'TEMPLATE');
  });

  test('TEMPLATE_LANG muda o idioma do modelo', () => {
    const amb = ambienteFormulario({ props: { TEMPLATE_CONVIDADO: 'confirmacao_alice', TEMPLATE_LANG: 'pt_PT' } });
    enviarFormulario(amb, { nomes: ['Ana Silva'], whatsapp: '(21) 90000-0002' });
    assert.equal(corpoDe(amb.chamadas[1]).content.attachments[0].template.language.code, 'pt_PT');
  });

  test('erro no envio do modelo vai para o status como no texto', () => {
    const amb = ambienteFormulario({
      props: { TEMPLATE_CONVIDADO: 'confirmacao_alice' },
      respostaApi: (u, o, n) => (n === 2 ? { codigo: 422, corpo: '{"code":"BUSINESS_RULE_VIOLATION"}' } : { codigo: 200, corpo: '{}' })
    });
    const r = enviarFormulario(amb, { nomes: ['Ana Silva'], whatsapp: '(21) 90000-0002' });
    assert.equal(r.status, 'Falha no envio: convidado: Omni respondeu 422');
  });
});

// ===================================================================
describe('provedores alternativos', () => {
  test('zapi: URL da instância, Client-Token e corpo phone/message', () => {
    const amb = ambienteFormulario({
      props: {
        WHATSAPP_PROVIDER: 'zapi', ZAPI_BASE_URL: 'https://api.z-api.io/',
        ZAPI_INSTANCE_ID: 'INST', ZAPI_TOKEN: 'TOKEN123', ZAPI_CLIENT_TOKEN: 'CLIENTE'
      }
    });
    enviarFormulario(amb, { nome: 'Ana', whatsapp: '(21) 90000-0002' });
    const chamada = amb.chamadas[0];
    assert.equal(chamada.url, 'https://api.z-api.io/instances/INST/token/TOKEN123/send-text');
    assert.equal(chamada.opcoes.headers['Client-Token'], 'CLIENTE');
    deepEq(corpoDe(chamada), { phone: '5521900000001', message: 'Presença confirmada na festa de 15 anos da Alice: Ana.' });
  });

  test('zapi: erro no corpo da resposta conta como falha, sem vazar o token', () => {
    const amb = ambienteFormulario({
      props: { WHATSAPP_PROVIDER: 'zapi', ZAPI_BASE_URL: 'https://api.z-api.io', ZAPI_INSTANCE_ID: 'INST', ZAPI_TOKEN: 'TOKEN123' },
      respostaApi: () => ({ codigo: 200, corpo: '{"error":"instância desconectada"}' })
    });
    const r = enviarFormulario(amb, { nome: 'Ana', whatsapp: '(21) 90000-0002' });
    assert.match(r.status, /mae: Z-API: instância desconectada/);
    assert.equal(r.status.includes('TOKEN123'), false);
  });

  test('meta: envia o modelo TEMPLATE_ORGANIZER com o nome como parâmetro', () => {
    const amb = ambienteFormulario({
      props: { WHATSAPP_PROVIDER: 'meta', WA_PHONE_NUMBER_ID: '123', WA_TOKEN: 'WT', TEMPLATE_ORGANIZER: 'aviso' }
    });
    enviarFormulario(amb, { nome: 'Ana', whatsapp: '(21) 90000-0002' });
    const chamada = amb.chamadas[0];
    assert.match(chamada.url, /\/123\/messages$/);
    assert.equal(chamada.opcoes.headers.Authorization, 'Bearer WT');
    assert.equal(corpoDe(chamada).template.name, 'aviso');
    deepEq(corpoDe(chamada).template.components[0].parameters.map((p) => p.text), ['Ana']);
  });

  test('provedor desconhecido vira falha registrada, sem derrubar a resposta', () => {
    const amb = ambienteFormulario({ props: { WHATSAPP_PROVIDER: 'outro' } });
    const r = enviarFormulario(amb, { nome: 'Ana', whatsapp: '(21) 90000-0002' });
    assert.equal(r.ok, true);
    assert.equal(
      r.status,
      'Falha no envio: mae: WHATSAPP_PROVIDER desconhecido: outro; convidado: WHATSAPP_PROVIDER desconhecido: outro'
    );
  });
});

// ===================================================================
describe('doPost - webhook da Z-API (mensagens recebidas)', () => {
  test('sem segredo configurado, registra a mensagem na aba Mensagens', () => {
    const amb = criarAmbiente({ props: { SHEET_ID: 'P' } });
    deepEq(enviarTexto(amb, mensagemZapi()), { ok: true, status: 'Registrada' });
    const aba = amb.abas['Mensagens'];
    deepEq(aba.linhas[0], ['Recebida em', 'Telefone', 'Nome', 'Mensagem']);
    deepEq(aba.linhas[1].slice(1), ['5521999990000', 'Maria', 'Oi, vou sim!']);
    assert.equal(amb.chamadas.length, 0, 'o webhook não responde nem envia mensagens');
  });

  test('converte o campo momment em data', () => {
    const amb = criarAmbiente({ props: { SHEET_ID: 'P' } });
    enviarTexto(amb, mensagemZapi({ momment: 1700000000000 }));
    assert.equal(amb.abas['Mensagens'].linhas[1][0].getTime(), 1700000000000);
  });

  test('com segredo configurado, recusa a requisição sem ?secret', () => {
    const amb = criarAmbiente({ props: { SHEET_ID: 'P', ZAPI_WEBHOOK_SECRET: 'segredo' } });
    deepEq(enviarTexto(amb, mensagemZapi()), { ok: false, erro: 'Não autorizado' });
    assert.equal(amb.abas['Mensagens'], undefined);
  });

  test('com segredo configurado, recusa ?secret errado', () => {
    const amb = criarAmbiente({ props: { SHEET_ID: 'P', ZAPI_WEBHOOK_SECRET: 'segredo' } });
    deepEq(enviarTexto(amb, mensagemZapi(), { secret: 'errado' }), { ok: false, erro: 'Não autorizado' });
  });

  test('com segredo configurado, aceita ?secret correto', () => {
    const amb = criarAmbiente({ props: { SHEET_ID: 'P', ZAPI_WEBHOOK_SECRET: 'segredo' } });
    deepEq(enviarTexto(amb, mensagemZapi(), { secret: 'segredo' }), { ok: true, status: 'Registrada' });
  });

  test('ignora mensagens enviadas pela própria instância (fromMe)', () => {
    const amb = criarAmbiente({ props: { SHEET_ID: 'P' } });
    deepEq(enviarTexto(amb, mensagemZapi({ fromMe: true })), { ok: true, ignorado: true });
    assert.equal(amb.abas['Mensagens'], undefined);
  });

  test('ignora mensagens de grupo', () => {
    const amb = criarAmbiente({ props: { SHEET_ID: 'P' } });
    deepEq(enviarTexto(amb, mensagemZapi({ isGroup: true })), { ok: true, ignorado: true });
  });

  test('mensagem sem campo text é registrada com o tipo entre colchetes', () => {
    const amb = criarAmbiente({ props: { SHEET_ID: 'P' } });
    enviarTexto(amb, mensagemZapi({ text: undefined }));
    assert.equal(amb.abas['Mensagens'].linhas[1][3], '[ReceivedCallback]');
  });

  test('JSON inválido recebe "Dados inválidos"', () => {
    const amb = criarAmbiente({ props: { SHEET_ID: 'P' } });
    deepEq(enviarTexto(amb, '{não é json'), { ok: false, erro: 'Dados inválidos' });
  });

  test('o segredo do webhook não bloqueia o formulário', () => {
    const amb = ambienteFormulario({ props: { ZAPI_WEBHOOK_SECRET: 'segredo' } });
    assert.equal(enviarFormulario(amb, { nome: 'Ana', whatsapp: '(21) 90000-0002' }).ok, true);
  });

  // Comportamento atual, registrado para que uma mudança aqui seja percebida.
  // Mensagem de outro tipo, sem campo text, não é reconhecida como webhook e vai para o formulário.
  test('mensagem de outro tipo sem campo text cai no formulário (comportamento atual)', () => {
    const amb = ambienteFormulario();
    const r = enviarTexto(amb, JSON.stringify({ type: 'ImageCallback', phone: '5521999990000', fromMe: false }));
    deepEq(r, { ok: false, erro: 'Nome é obrigatório' });
    assert.equal(amb.chamadas.length, 0);
  });
});

// ===================================================================
describe('criarWebhookOmni (rodado uma vez pelo editor)', () => {
  const propsWebhook = {
    OMNI_CHANNEL_ID: 'CANAL',
    OMNI_SECRET_KEY: 'CHAVE123',
    WEBHOOK_URL: 'https://script.google.com/macros/s/ID/exec'
  };
  const criado = () => ({ codigo: 201, corpo: '{"id":"WH1","payloadFormat":"Z_API","status":"ENABLED"}' });

  test('POST em /webhooks do canal, com Bearer, MESSAGE_RECEIVED e formato Z_API', () => {
    const amb = criarAmbiente({ props: propsWebhook, respostaApi: criado });
    assert.equal(amb.g.criarWebhookOmni(), 'WH1');
    const chamada = amb.chamadas[0];
    assert.equal(chamada.url, 'https://api.omni.z-api.io/v1/channels/CANAL/webhooks');
    assert.equal(chamada.opcoes.method, 'post');
    assert.equal(chamada.opcoes.headers.Authorization, 'Bearer CHAVE123');
    deepEq(corpoDe(chamada), {
      url: 'https://script.google.com/macros/s/ID/exec',
      events: ['MESSAGE_RECEIVED'],
      payloadFormat: 'Z_API',
      description: 'Confirmação de presença - planilha'
    });
    assert.match(amb.logs[0], /Webhook criado: id WH1, formato Z_API, status ENABLED/);
  });

  test('com ZAPI_WEBHOOK_SECRET, a URL vai com ?secret=', () => {
    const amb = criarAmbiente({ props: { ...propsWebhook, ZAPI_WEBHOOK_SECRET: 'a b' }, respostaApi: criado });
    amb.g.criarWebhookOmni();
    assert.equal(corpoDe(amb.chamadas[0]).url, 'https://script.google.com/macros/s/ID/exec?secret=a%20b');
  });

  test('sem propriedades, avisa o que falta e não chama a API', () => {
    const amb = criarAmbiente({ props: { OMNI_CHANNEL_ID: 'CANAL', OMNI_SECRET_KEY: 'CHAVE123' } });
    assert.throws(() => amb.g.criarWebhookOmni(), /Faltam propriedades: OMNI_CHANNEL_ID, OMNI_SECRET_KEY ou WEBHOOK_URL/);
    assert.equal(amb.chamadas.length, 0);
  });

  test('recusa criar de novo quando OMNI_WEBHOOK_ID já existe, sem chamar a API', () => {
    const amb = criarAmbiente({ props: { ...propsWebhook, OMNI_WEBHOOK_ID: 'WH0' }, respostaApi: criado });
    assert.throws(() => amb.g.criarWebhookOmni(), /Webhook já criado \(id WH0\)/);
    assert.equal(amb.chamadas.length, 0);
  });

  test('grava o id do webhook em OMNI_WEBHOOK_ID depois de criar', () => {
    const amb = criarAmbiente({ props: propsWebhook, respostaApi: criado });
    amb.g.criarWebhookOmni();
    assert.equal(amb.props.OMNI_WEBHOOK_ID, 'WH1');
  });

  test('sucesso sem id: não grava OMNI_WEBHOOK_ID e avisa no registro', () => {
    const amb = criarAmbiente({ props: propsWebhook, respostaApi: () => ({ codigo: 201, corpo: 'não é json' }) });
    assert.equal(amb.g.criarWebhookOmni(), null);
    assert.equal('OMNI_WEBHOOK_ID' in amb.props, false);
    assert.match(amb.logs[0], /A Omni respondeu 201 sem id\. Confira no painel/);
  });

  test('erro da API aparece com o código e a mensagem', () => {
    const amb = criarAmbiente({
      props: propsWebhook,
      respostaApi: () => ({ codigo: 400, corpo: '{"error":"At least one event is required"}' })
    });
    assert.throws(() => amb.g.criarWebhookOmni(), /Omni respondeu 400: At least one event is required/);
  });
});
