console.log("APP.JS CARREGOU");
//import { renderListaScripts } from "./renderer.js";

import { renderSidebar, setActiveTool } from "./sidebar.js";

import { state } from "./state.js";

import { renderToolbar } from "./toolbar.js";

import { renderScript } from "./renderer.js";

import { renderQRCode } from "./tools/qrcodes.js";

import { renderPDFs } from "./tools/pdfs.js";

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
        if (ferramentaAtual !== "adesivos") return;

        state.setScript(model);

        renderScript(model);

    }

    catch (erro) {

        console.error(erro);

    }

});

let encerrarFerramenta = null;
let ferramentaAtual = null;
let navegando = false;

async function selecionarFerramenta(tool) {

    if (navegando || tool === ferramentaAtual) return;
    navegando = true;
    try {
        if (encerrarFerramenta?.confirmarSaida && !await encerrarFerramenta.confirmarSaida()) return;
        await encerrarFerramenta?.();
        encerrarFerramenta = null;
        ferramentaAtual = tool;

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

            case "pdfs":

                encerrarFerramenta = renderPDFs();
                break;

        }

    } catch (erro) {
        console.error("Erro ao trocar de ferramenta:", erro);
    } finally {
        navegando = false;
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
        if (ferramentaAtual !== "adesivos") return;

        renderToolbar(dados.scripts);

    }

    catch (erro) {

        console.error(erro);

    }

}

iniciar();
