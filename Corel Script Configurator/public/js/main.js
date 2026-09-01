console.log("APP.JS CARREGOU");
//import { renderListaScripts } from "./renderer.js";

import { renderSidebar, setActiveTool } from "./sidebar.js";

import { state } from "./state.js";

import { renderToolbar } from "./toolbar.js";

import { renderScript } from "./renderer.js";

async function iniciar() {

    try {

        renderSidebar(selecionarFerramenta);

        selecionarFerramenta("inicio");

    }

    catch (erro) {

        console.error(erro);

    }

}

window.addEventListener("abrir-script", async (event) => {

    const nomeArquivo = event.detail;

    console.log("Abrindo:", nomeArquivo);

    try {

        const resposta = await fetch(
            "/api/scripts/" + encodeURIComponent(nomeArquivo)
        );

        const model = await resposta.json();

        state.setScript(model);

        renderScript(model);

    }

    catch (erro) {

        console.error(erro);

    }

});

function selecionarFerramenta(tool) {

    setActiveTool(tool);

    const app = document.getElementById("app");

    app.innerHTML = "";

    switch (tool) {

        case "inicio":

            renderInicio();
            break;

        case "adesivos":

            carregarAdesivos();
            break;

        case "qrcodes":

            renderQRCode();
            break;

    }

}

function renderInicio() {

    const app = document.getElementById("app");

    app.innerHTML = "";

    const titulo = document.createElement("h1");

    titulo.textContent = "Central de Ferramentas";

    app.appendChild(titulo);

    const texto = document.createElement("p");

    texto.textContent =
        "Escolha uma ferramenta no menu lateral.";

    app.appendChild(texto);

}

async function carregarAdesivos() {

    const app = document.getElementById("app");

    app.innerHTML = "";

    try {

        const resposta = await fetch("/api/scripts");

        const dados = await resposta.json();

        renderToolbar(dados.scripts);

    }

    catch (erro) {

        console.error(erro);

    }

}

function renderQRCode() {

    const app = document.getElementById("app");

    app.innerHTML = "";

    const titulo = document.createElement("h1");

    titulo.textContent = "Gerador de QR Codes";

    app.appendChild(titulo);

    const texto = document.createElement("p");

    texto.textContent =
        "Esta ferramenta será implementada em seguida.";

    app.appendChild(texto);

}

iniciar();