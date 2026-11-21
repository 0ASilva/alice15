# Configurar o backend da confirmação de presença

O convidado digita os nomes da família (até 10) e um WhatsApp no site. O Apps Script grava uma
linha por pessoa na planilha, avisa a organização no WhatsApp e manda uma confirmação, com todos
os nomes, para o WhatsApp informado.

## 1. Planilha
1. Crie uma planilha no Google Sheets.
2. Renomeie a primeira aba para `Respostas`.
3. Na linha 1, escreva os cabeçalhos: `Data`, `Nome`, `WhatsApp`, `Status`.
4. Copie o ID da planilha (o trecho entre `/d/` e `/edit` na URL).

## 2. Apps Script
1. Abra o projeto do Apps Script que já está implantado.
2. Substitua todo o conteúdo de `Code.gs` pelo conteúdo de `backend/Code.gs`.
3. Em **Configurações do projeto > Propriedades do script**, adicione:

| Propriedade | Valor |
| --- | --- |
| `SHEET_ID` | ID da planilha |
| `SHEET_NAME` | `Respostas` |
| `MAE_PHONE` | WhatsApp que recebe o aviso de cada confirmação, só números, com 55 e DDD (ex.: `5521999999999`) |
| `OMNI_CHANNEL_ID` | ID do canal (painel da Omni > Conexões > o canal > id) |
| `OMNI_SECRET_KEY` | Secret Key (painel da Omni > Segurança) |

Não cole esses valores no código nem no site. Eles ficam só nas propriedades do script.

Se `MAE_PHONE` não existir, o sistema usa `ORGANIZER_PHONE`, se ela já estiver lá.

4. Publique uma **nova versão**: Implantar > Gerenciar implantações > lápis (editar) > Versão: Nova versão > Implantar.
   A URL continua a mesma, e ela já está em `script.js`.

## 3. Mensagens
- O envio usa a API da Omni Z-API (`/v1/channels/{canal}/messages`), com a Secret Key.
- Os textos ficam no topo de `Code.gs`: `TEXTO_CONFIRMACAO` (aviso) e `TEXTO_CONVIDADO` (confirmação ao convidado).
  `{nomes}` vira a lista de nomes e `{nome}` vira o primeiro nome da primeira pessoa.
- **Janela de 24 horas:** nesse canal da Meta, texto livre só chega se a pessoa mandou mensagem
  para o número do canal nas últimas 24 horas. Se não mandou, o envio falha e a coluna Status
  mostra o motivo. Para resolver de vez, é preciso um modelo de mensagem aprovado pela Meta.
- Outros provedores, só se precisar: `WHATSAPP_PROVIDER` = `zapi` (Z-API por instância:
  `ZAPI_BASE_URL`, `ZAPI_INSTANCE_ID`, `ZAPI_TOKEN`, `ZAPI_CLIENT_TOKEN`) ou `meta`
  (`WA_PHONE_NUMBER_ID`, `WA_TOKEN`, `TEMPLATE_ORGANIZER`, `TEMPLATE_LANG`).

## 3b. Modelos aprovados (para não precisar do "oi")
Sem modelo aprovado, a Meta só entrega a mensagem a quem escreveu para o número do canal nas últimas
24 horas. Com modelo aprovado, a mensagem chega para qualquer convidado.

1. No painel da Omni, em **Templates**, crie os modelos (categoria **Utility**, idioma **Português (BR)**).
   As variáveis têm nome. Para inserir uma: escreva o nome no campo `{{ }}` abaixo do texto e clique em **Inserir**.
   - `confirmacao_alice` (para o convidado), com as variáveis `primeiro_nome` e `nomes`:
     `Olá, {{primeiro_nome}}! Presença confirmada na festa de 15 anos da Alice para: {{nomes}}. Sexta, 11 de dezembro de 2026, às 20h30, na Rua Monte Cristo, s/n, Pau Grande. Traje: esporte fino. Até lá!`
     Exemplos para a Meta: `primeiro_nome` = `Ana`, `nomes` = `Ana Silva e João Silva`.
   - `aviso_alice` (para a organização), com a variável `nomes`:
     `Nova confirmação de presença na festa de 15 anos da Alice: {{nomes}}. A lista completa está na planilha.`
     Exemplo para a Meta: `nomes` = `Ana Silva e João Silva`.
   Os nomes das variáveis precisam ser exatamente esses: o script envia os valores por esses nomes.
2. Espere o status ficar **Aprovado** (a Meta revisa; pode levar de minutos a um dia).
3. Em **Propriedades do script**, crie `TEMPLATE_CONVIDADO` = `confirmacao_alice` e `TEMPLATE_AVISO` = `aviso_alice`.
   Se o idioma não for `pt_BR`, crie também `TEMPLATE_LANG`.
4. Não precisa publicar de novo: a partir do próximo envio, o script usa os modelos.

Se o nome do modelo ficar diferente no painel, use o nome que aparecer lá. Sem essas propriedades,
o script continua mandando texto livre.

## 3a. Receber mensagens (webhook, opcional)
Mensagens recebidas no WhatsApp são gravadas na aba `Mensagens` (Recebida em, Telefone, Nome, Mensagem).
O código entende o formato **Z_API**.

- Pela documentação da Omni, o formato do evento (`Z_API`) só pode ser escolhido em um webhook
  **de canal**. O webhook da conta usa o formato padrão, que o código não entende.
- Um canal com webhook próprio deixa de mandar eventos para o webhook da conta.

Para criar o webhook do canal (uma vez só):
1. Em **Propriedades do script**, adicione `WEBHOOK_URL` com a URL da implantação (a que termina em `/exec`).
   `OMNI_CHANNEL_ID` e `OMNI_SECRET_KEY` já devem estar preenchidas.
2. Opcional: `ZAPI_WEBHOOK_SECRET` (um texto longo). Se existir, a URL vai com `?secret=` sozinha.
3. No editor do Apps Script, escolha a função `criarWebhookOmni` no menu de cima e clique em **Executar**.
   Na primeira vez, o Google pede autorização: aceite com a sua conta.
4. No **Registro de execução**, deve aparecer `Webhook criado: id ..., formato Z_API, status ENABLED`.
5. Não rode de novo: cada execução cria outro webhook, e cada mensagem passa a ser gravada duas vezes.
   Se isso acontecer, apague o repetido no painel da Omni.
6. Mande uma mensagem de outro celular para o número do canal. Uma linha nova deve aparecer na aba
   `Mensagens`. No painel da Omni, **Webhooks > Entregas de webhook** mostra se a entrega deu certo.

`MESSAGES_SHEET_NAME` (padrão `Mensagens`) é opcional.

## 4. Testar
1. Sem os modelos da seção 3b: do celular de teste, mande antes um "oi" para o número do canal.
2. Abra o site e envie um nome e esse WhatsApp.
3. Veja a nova linha na planilha com `Status` = `Enviado`, e confira se as mensagens chegaram.
4. Se aparecer `Falha no envio: ...`, a linha continua salva. O texto depois de `mae:` ou `convidado:` diz o motivo.

Testes automáticos do backend (não enviam nada): `node --test backend/tests/code.test.js`.
