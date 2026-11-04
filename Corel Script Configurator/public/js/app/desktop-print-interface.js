export const interfaceImpressao=`
    <header class="print-config-heading"><h2 id="print-config-title">Impressão</h2><p id="print-config-scope"></p></header>
    <div class="print-config-body">
        <section class="print-config-settings" aria-label="Configurações de impressão">
            <fieldset><legend>Impressora</legend>
                <label for="print-config-printer">Nome da impressora</label>
                <select id="print-config-printer"><option value="">Selecione uma impressora</option></select>
                <div class="print-config-fields"><button type="button" id="print-config-properties" disabled>Propriedades do driver</button><button type="button" id="print-config-refresh">Atualizar lista</button></div>
                <div class="print-config-driver-values"><span>Cópias <output id="print-config-copies">—</output></span><small>Definidas nas propriedades do driver</small></div>
            </fieldset>
            <fieldset><legend>Páginas a imprimir</legend><p id="print-config-pages">Todas as páginas do documento</p><small>A navegação da prévia não altera as páginas enviadas.</small></fieldset>
            <fieldset><legend>Dimensionamento e manuseio</legend>
                <div class="print-config-fields">
                    <label>Escala do aplicativo<select id="print-config-scale"><option value="100">Tamanho real (100%)</option><option value="120">120%</option><option value="custom">Escala personalizada</option></select></label>
                    <label id="print-config-custom-label" hidden>Percentual<input id="print-config-custom" type="text" inputmode="decimal" value="100" aria-label="Escala personalizada em porcentagem"><small>De 1% a 400%</small></label>
                </div>
                <label id="print-config-paper-label">Papel da simulação<select id="print-config-paper"><option value="a3">A3 — 297 × 420 mm</option><option value="driver">Papel informado pelo driver</option></select></label>
                <small>Escala do aplicativo e zoom/encaixe do driver são controles independentes.</small>
            </fieldset>
            <fieldset><legend>Orientação</legend><p id="print-config-orientation">Controlada pelo driver</p><small>Altere a orientação nas propriedades da impressora.</small></fieldset>
            <details><summary>Informações do documento e da impressora</summary><dl id="print-config-summary"></dl></details>
            <details><summary>Como interpretar a prévia</summary><p>A folha branca representa o papel físico. A linha tracejada indica a área imprimível conhecida. O vermelho mostra conteúdo fora do papel.</p><p>O documento de trabalho fornece a imagem da prévia. Cópias, cor, duplex, zoom, encaixe e acabamentos privados continuam sob controle do driver. A prévia geométrica não simula esses recursos.</p></details>
            <details id="print-config-diagnostics"><summary>Diagnóstico opcional Konica / Hold</summary>
                <p id="print-config-hold" role="status"></p>
                <div class="print-config-fields"><button type="button" id="print-config-verify-hold" disabled>Verificar Hold da sessão</button><button type="button" id="print-config-fixture" hidden>Abrir PDF sintético de teste</button><button type="button" id="print-config-send-hold" disabled>Enviar trabalho retido</button></div>
                <button type="button" id="print-config-new-operation" hidden>Nova operação</button>
                <details><summary>Configurações recuperadas do PostScript</summary><pre id="print-config-ps-details">Aguardando captura verificada.</pre></details>
            </details>
        </section>
        <section class="print-config-preview-pane" aria-label="Prévia da impressão">
            <div class="print-config-preview-heading"><h3>Pré-visualização</h3><nav aria-label="Página da prévia"><button type="button" id="print-config-previous" aria-label="Página anterior da prévia">‹</button><span id="print-config-page" aria-live="polite">—</span><button type="button" id="print-config-next" aria-label="Próxima página da prévia">›</button></nav></div>
            <svg id="print-config-preview" role="img" aria-label="Prévia do conteúdo na folha"></svg>
            <p id="print-config-measures"></p><p id="print-config-warning" role="status"></p>
            <div id="print-config-canvas" hidden></div>
        </section>
    </div>
    <footer class="print-config-footer"><p id="print-config-status" role="status" aria-live="polite"></p><div class="print-config-actions"><button type="button" id="print-config-close">Cancelar</button><button type="button" id="print-config-print" hidden>Imprimir</button></div></footer>`;
