export const alertaAssinaturas = `
<div id="pdf-assinaturas-alerta" class="pdf-assinaturas-alerta" role="status" hidden>
    <div class="pdf-assinaturas-cabecalho">
        <div><strong id="pdf-assinaturas-titulo"></strong><p id="pdf-assinaturas-resumo"></p></div>
        <div class="pdf-assinaturas-acoes">
            <button id="pdf-assinaturas-resolver" type="button">Resolver</button>
            <button id="pdf-assinaturas-expandir" type="button" aria-expanded="false" aria-controls="pdf-assinaturas-expandido">Ver detalhes</button>
        </div>
    </div>
    <p id="pdf-assinaturas-expandido" class="pdf-assinaturas-detalhes" hidden></p>
</div>`;

export const janelaCopiaImpressao = `
<dialog id="pdf-impressao-dialogo" class="pdf-impressao-dialogo" aria-labelledby="pdf-impressao-titulo">
    <h3 id="pdf-impressao-titulo">Preparar cópia para impressão</h3>
    <p id="pdf-impressao-contexto"></p>
    <p>Confira a aparência com o original em um leitor confiável antes de aplicar.</p>
    <div class="pdf-impressao-config">
        <label for="pdf-impressao-modo">Estratégia</label>
        <select id="pdf-impressao-modo"><option value="necessarias">Rasterizar somente páginas necessárias</option><option value="inteiro">Rasterizar documento inteiro</option></select>
        <label for="pdf-impressao-dpi">Resolução</label>
        <select id="pdf-impressao-dpi"><option value="300">300 DPI — padrão</option><option value="600">600 DPI — detalhes finos (sujeito ao limite de memória)</option></select>
        <label for="pdf-impressao-manual">Páginas adicionais/manuais (ex.: 2, 4-8)</label>
        <input id="pdf-impressao-manual" type="text" placeholder="Opcional; obrigatório se não houver páginas determinadas">
    </div>
    <details id="pdf-impressao-detalhes" class="pdf-impressao-detalhes">
        <summary id="pdf-impressao-detalhes-alternar">Mostrar detalhes técnicos</summary>
        <p id="pdf-impressao-diagnostico" class="pdf-assinaturas-detalhes"></p>
        <p>Proposta baseada na renderização do PDF.js. Compare as assinaturas com o original em um leitor confiável, considerando as edições já aplicadas. Se houver partes faltando ou quadrados pretos, cancele: este fluxo não comprova a aparência em todos os documentos.</p>
        <p>As páginas necessárias incluem outras anotações imprimíveis. Nas páginas rasterizadas, texto e vetores tornam-se imagem RGB opaca sobre branco, sem texto pesquisável; cores e transparências podem diferir no RIP. Formulários, interatividade, navegação e acessibilidade não são preservados na cópia independente. As demais páginas permanecem vetoriais quando isso for seguro. Não há validação de certificado ou validade jurídica.</p>
        <p id="pdf-impressao-resultado-tecnico" class="pdf-assinaturas-detalhes"></p>
    </details>
    <button id="pdf-impressao-prever" type="button">Preparar e pré-visualizar</button>
    <p id="pdf-impressao-progresso" role="status"></p>
    <div id="pdf-impressao-controles-revisao" class="pdf-impressao-revisao" hidden>
        <button id="pdf-impressao-anterior" type="button" disabled>Anterior necessária</button>
        <label for="pdf-impressao-pagina">Conferir página</label><input id="pdf-impressao-pagina" type="number" min="1" step="1" value="1" disabled>
        <button id="pdf-impressao-ir" type="button" disabled>Abrir página</button>
        <button id="pdf-impressao-proxima" type="button" disabled>Próxima necessária</button>
        <label><input id="pdf-impressao-origem" type="checkbox" disabled> Comparar com documento antes do preparo (PDF.js)</label>
    </div>
    <p id="pdf-impressao-revisadas"></p>
    <div id="pdf-impressao-previa" class="pdf-impressao-previa" aria-label="Prévia da proposta para impressão"></div>
    <label id="pdf-impressao-confirmacao" hidden><input id="pdf-impressao-conferi" type="checkbox" disabled> Conferi as páginas rasterizadas e a aparência das assinaturas, comparando com o original em leitor confiável.</label>
    <div class="pdf-confirmacao-acoes">
        <button id="pdf-impressao-cancelar" type="button">Cancelar / fechar</button>
        <button id="pdf-impressao-aplicar" type="button" disabled hidden>Aplicar cópia ao documento</button>
    </div>
    <small>Não há download automático. Após aplicar, use Salvar cópia; o nome incluirá _IMPRESSAO.</small>
</dialog>`;
