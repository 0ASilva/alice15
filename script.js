// Confirmação de presença: o convidado informa o nome e o WhatsApp, e a página
// envia ao backend (Google Apps Script, ver backend/SETUP.md), que grava na
// planilha, avisa a organização e manda a confirmação ao convidado no WhatsApp.
(function () {
  "use strict";

  // URL da implantação do Apps Script (Web App). Não é senha: a página precisa dela.
  var APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyRRODTOv2nm-BZKOVedgJsoCZg9nTOyf1UI5bqISA34RbGy8xgpYSTYDJroI1TD9aS/exec";

  var pagina = document.getElementById("pagina");
  var obrigado = document.getElementById("obrigado");
  var form = document.getElementById("form-rsvp");
  var erro = document.getElementById("erro");
  var botao = form.querySelector("button[type=submit]");

  // Envia ao backend. Resolve com { ok, erro }.
  function enviarAoBackend(dados) {
    if (!APPS_SCRIPT_URL) {
      return Promise.resolve({ ok: false, erro: "Sistema ainda não configurado. Avise a organizadora." });
    }
    return fetch(APPS_SCRIPT_URL, {
      method: "POST",
      // text/plain evita a checagem CORS (preflight) do Apps Script
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(dados)
    })
      .then(function (res) { return res.json(); })
      .then(function (r) { return { ok: !!(r && r.ok), erro: r && r.erro }; })
      .catch(function () {
        return { ok: false, erro: "Não foi possível enviar agora. Confira sua conexão e tente de novo." };
      });
  }

  var MAX_PESSOAS = 10;
  var listaPessoas = document.getElementById("lista-pessoas");
  var botaoAdicionar = document.getElementById("adicionar-pessoa");

  function camposPessoa() {
    return listaPessoas.querySelectorAll("input[name=pessoa]");
  }

  // Acrescenta mais um campo de nome, até MAX_PESSOAS.
  function adicionarPessoa() {
    var total = camposPessoa().length;
    if (total >= MAX_PESSOAS) return;
    var campo = document.createElement("input");
    campo.type = "text";
    campo.name = "pessoa";
    campo.maxLength = 80;
    campo.placeholder = "Nome e sobrenome";
    campo.setAttribute("aria-label", "Pessoa " + (total + 1));
    listaPessoas.appendChild(campo);
    campo.focus();
    botaoAdicionar.disabled = total + 1 >= MAX_PESSOAS;
  }

  botaoAdicionar.addEventListener("click", adicionarPessoa);

  // Volta a ter um só campo de nome, vazio.
  function limparPessoas() {
    var campos = camposPessoa();
    for (var i = campos.length - 1; i > 0; i--) campos[i].remove();
    campos[0].value = "";
    botaoAdicionar.disabled = false;
  }

  function lerNomes() {
    var nomes = [];
    camposPessoa().forEach(function (campo) {
      var nome = campo.value.trim();
      if (nome) nomes.push(nome);
    });
    return nomes;
  }

  function mostrarObrigado(nomes) {
    document.getElementById("obrigado-texto").textContent = nomes.length === 1
      ? nomes[0] + ", sua presença está confirmada. Até lá!"
      : "Presença confirmada para " + nomes.length + " pessoas: " + nomes.join(", ") + ". Até lá!";
    pagina.hidden = true;
    obrigado.hidden = false;
    window.scrollTo(0, 0);
  }

  form.addEventListener("submit", function (ev) {
    ev.preventDefault();
    erro.textContent = "";

    var nomes = lerNomes();
    var whatsapp = String(form.elements["whatsapp"].value || "").replace(/\D/g, "");
    if (!nomes.length) {
      erro.textContent = "Escreva pelo menos um nome para confirmar.";
      camposPessoa()[0].focus();
      return;
    }
    if (whatsapp.length !== 11) {
      erro.textContent = whatsapp.length === 0
        ? "Informe o WhatsApp: 11 números, com DDD. Ex.: (21) 99999-9999."
        : "O WhatsApp tem " + whatsapp.length + " números, e precisa ter 11: DDD + celular, sem +55. Ex.: (21) 99999-9999.";
      form.elements["whatsapp"].focus();
      return;
    }

    botao.disabled = true;
    botao.textContent = "Enviando...";

    enviarAoBackend({ nomes: nomes, whatsapp: whatsapp }).then(function (r) {
      botao.disabled = false;
      botao.textContent = "Enviar resposta";
      if (r.ok) {
        mostrarObrigado(nomes);
      } else {
        erro.textContent = r.erro || "Não foi possível enviar agora. Tente de novo.";
      }
    });
  });

  document.getElementById("refazer").addEventListener("click", function () {
    form.reset();
    limparPessoas();
    obrigado.hidden = true;
    pagina.hidden = false;
    window.scrollTo(0, 0);
  });
})();
