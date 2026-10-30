import { renderToolbar } from "../toolbar.js";
import { renderScript } from "../renderer.js";
import { state } from "../state.js";

export function montarAdesivos(painel) {
    painel.classList.add("app-modulo-generico");
    const abort = new AbortController(); let encerrado = false, geracao = 0;
    const erro = e => { if (!encerrado && e.name !== "AbortError") painel.textContent = e.message; };
    async function abrir(nome) {
        const versao = ++geracao;
        try {
            const resposta = await fetch("/api/scripts/" + encodeURIComponent(nome), { signal: abort.signal });
            if (!resposta.ok) throw new Error("Não foi possível abrir o modelo.");
            const modelo = await resposta.json();
            if (encerrado || versao !== geracao) return;
            state.setScript(modelo); renderScript(modelo, painel);
        } catch (e) { erro(e); }
    }
    (async () => {
        const resposta = await fetch("/api/scripts", { signal: abort.signal });
        if (!resposta.ok) throw new Error("Não foi possível listar os modelos.");
        const dados = await resposta.json(); if (!encerrado) renderToolbar(dados.scripts, painel, abrir);
    })().catch(erro);
    return { destruir() { encerrado = true; ++geracao; abort.abort(); state.setScript(null); } };
}
