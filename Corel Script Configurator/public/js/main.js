console.log("APP.JS CARREGOU");
//import { renderListaScripts } from "./renderer.js";

import { state } from "./state.js";

import { renderToolbar } from "./toolbar.js";

import { renderScript } from "./renderer.js";

async function iniciar() {

    try {

        const resposta = await fetch("/api/scripts");

        const dados = await resposta.json();

        renderToolbar(dados.scripts);

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

iniciar();