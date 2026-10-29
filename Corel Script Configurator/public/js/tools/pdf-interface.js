import { interfaceRedimensionar } from "./pdf-resize-interface.js";
// Estrutura e interação visual; as alterações do PDF continuam no PDFWorkspace.
export function renderPDFInterface(app, aoAlterarPainel = () => {}) {
    app.innerHTML = `
        <section id="pdf-ferramenta" class="pdf-editor" aria-label="Gerenciador de PDFs">
            <header class="pdf-topo">
                <h2>Gerenciador de PDFs</h2>
                <div class="pdf-acoes-principais">
                    <label class="pdf-importar">
                        Importar PDF
                        <input id="pdf-arquivo" type="file" accept=".pdf,application/pdf" hidden>
                    </label>
                    <button id="pdf-informacoes" type="button" aria-controls="pdf-detalhes" aria-expanded="false" disabled>Informações</button>
                    <div class="pdf-exportacao">
                        <button id="pdf-salvar-copia" type="button" disabled>Exportar / Salvar cópia</button>
                        <span id="pdf-alteracoes-pendentes" role="status" hidden>Alterações não exportadas</span>
                    </div>
                </div>
            </header>
            <p id="pdf-status" role="status">Selecione um PDF ou use Unir PDFs para começar.</p>
            <div id="pdf-conteudo" hidden>
                <h3 id="pdf-nome"></h3>
                <div class="pdf-resumo">
                    <div><small>Páginas</small><strong id="pdf-total">-</strong></div>
                    <div><small>Dimensões</small><strong id="pdf-tamanho">-</strong></div>
                    <div><small>Cores</small><strong>Não verificado</strong></div>
                </div>
            </div>
            <div id="pdf-detalhes" class="pdf-detalhes" hidden>
                <h3>Informações do arquivo</h3>
                <p id="pdf-alertas"></p>
                <p><strong>Cores:</strong> Análise RGB/CMYK ainda não implementada.</p>
                <p><strong>Sangria, fontes e resolução:</strong> Ainda não verificadas.</p>
                <h4>Dimensões por página</h4>
                <ul id="pdf-lista-paginas"></ul>
            </div>
            <div class="pdf-barra">
                <nav class="pdf-navegacao" aria-label="Navegação do PDF">
                    <button id="pdf-anterior" type="button" disabled>Anterior</button>
                    <span id="pdf-pagina-atual" aria-live="polite">Nenhum PDF aberto</span>
                    <button id="pdf-proxima" type="button" disabled>Próxima</button>
                </nav>
                <button id="pdf-painel-alternar" type="button" aria-controls="pdf-painel-ferramentas" aria-expanded="true">Recolher ferramentas</button>
            </div>
            <div id="pdf-area-trabalho" class="pdf-area-trabalho">
                <aside id="pdf-painel-miniaturas" class="pdf-painel-miniaturas" aria-label="Páginas do PDF" hidden>
                    <h3>Páginas</h3>
                    <div class="pdf-miniaturas-controles">
                        <button id="pdf-mover-cima" type="button" aria-label="Mover página atual para cima" title="Mover página atual para cima" disabled>↑</button>
                        <button id="pdf-mover-baixo" type="button" aria-label="Mover página atual para baixo" title="Mover página atual para baixo" disabled>↓</button>
                    </div>
                    <div class="pdf-mover-posicao">
                        <label for="pdf-mover-destino">Mover para posição</label>
                        <input id="pdf-mover-destino" type="number" min="1" step="1" inputmode="numeric" placeholder="Posição" disabled>
                        <button id="pdf-mover-aplicar" type="button" disabled>Mover</button>
                    </div>
                    <p class="pdf-miniaturas-ajuda">Selecione a página e informe o destino, use ↑/↓ ou arraste. Role a lista para ver outras páginas.</p>
                    <div id="pdf-miniaturas" class="pdf-miniaturas" tabindex="0" aria-label="Miniaturas navegáveis"></div>
                </aside>
                <div id="pdf-visualizacao" class="pdf-visualizacao" aria-label="Página do PDF">
                    <p class="pdf-vazio">Importe ou una PDFs para visualizar o documento.</p>
                </div>
                <aside id="pdf-painel-ferramentas" class="pdf-painel-ferramentas" aria-label="Ferramentas do PDF">
                    <h3>Ferramentas</h3>
                    <div class="pdf-ferramentas-menu" aria-label="Escolher ferramenta">
                        <button id="pdf-ferramenta-rotacionar" type="button" aria-controls="pdf-opcoes-rotacionar" aria-expanded="true">Rotacionar</button>
                        <button id="pdf-ferramenta-organizar" type="button" aria-controls="pdf-opcoes-organizar" aria-expanded="false">Organizar páginas</button>
                        <button id="pdf-ferramenta-unir" type="button" aria-controls="pdf-opcoes-unir" aria-expanded="false">Unir PDFs</button>
                        <button id="pdf-ferramenta-redimensionar" type="button" aria-controls="pdf-opcoes-redimensionar" aria-expanded="false">Redimensionar</button>
                    </div>
                    <section id="pdf-opcoes-rotacionar" class="pdf-ferramenta-opcoes" aria-label="Rotacionar">
                        <p>Girar a página exibida</p>
                        <button id="pdf-girar-esquerda" type="button" disabled>↶ Girar esquerda</button>
                        <button id="pdf-girar-direita" type="button" disabled>↷ Girar direita</button>
                    </section>
                    <section id="pdf-opcoes-organizar" class="pdf-ferramenta-opcoes pdf-selecao" aria-label="Organizar páginas" hidden>
                        <label for="pdf-selecao-paginas">Selecionar páginas</label>
                        <input id="pdf-selecao-paginas" type="text" placeholder="Ex.: 3, 7 ou 4-8">
                        <button id="pdf-extrair" type="button" disabled>Extrair páginas</button>
                        <button id="pdf-excluir" type="button" disabled>Excluir páginas</button>
                        <small>Use números ou intervalos. A extração baixa uma cópia; a exclusão atualiza o documento.</small>
                    </section>
                    <section id="pdf-opcoes-unir" class="pdf-ferramenta-opcoes" aria-label="Unir PDFs" hidden>
                        <p>Acrescente PDFs ao documento ou inicie um novo.</p>
                        <div id="pdf-uniao-opcoes"></div>
                    </section>
                    ${interfaceRedimensionar}
                </aside>
            </div>
            <dialog id="pdf-confirmar-descarte" class="pdf-confirmacao" aria-labelledby="pdf-confirmar-titulo" aria-describedby="pdf-confirmar-mensagem">
                <h3 id="pdf-confirmar-titulo">Alterações não exportadas</h3>
                <p id="pdf-confirmar-mensagem">Você possui alterações no documento atual que ainda não foram exportadas. Se sair agora, essas alterações serão perdidas.</p>
                <div class="pdf-confirmacao-acoes">
                    <button id="pdf-continuar-editando" type="button" autofocus>Continuar editando</button>
                    <button id="pdf-descartar-sair" type="button">Descartar e sair</button>
                </div>
            </dialog>
        </section>
    `;

    const raiz = app.querySelector("#pdf-ferramenta");
    const ferramentas = ["rotacionar", "organizar", "unir", "redimensionar"];
    for (const nome of ferramentas) {
        raiz.querySelector("#pdf-ferramenta-" + nome).addEventListener("click", () => {
            for (const opcao of ferramentas) {
                const ativa = opcao === nome;
                raiz.querySelector("#pdf-opcoes-" + opcao).hidden = !ativa;
                raiz.querySelector("#pdf-ferramenta-" + opcao).setAttribute("aria-expanded", String(ativa));
            }
        });
    }

    const alternar = raiz.querySelector("#pdf-painel-alternar");
    const painel = raiz.querySelector("#pdf-painel-ferramentas");
    alternar.addEventListener("click", () => {
        painel.hidden = !painel.hidden;
        raiz.classList.toggle("pdf-painel-recolhido", painel.hidden);
        alternar.setAttribute("aria-expanded", String(!painel.hidden));
        alternar.textContent = painel.hidden ? "Mostrar ferramentas" : "Recolher ferramentas";
        aoAlterarPainel();
    });

    return raiz;
}
