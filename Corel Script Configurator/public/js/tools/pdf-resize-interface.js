export const interfaceRedimensionar = `
<section id="pdf-opcoes-redimensionar" class="pdf-ferramenta-opcoes pdf-redimensionar" aria-label="Redimensionar páginas" hidden>
    <label for="pdf-redim-formato">Formato</label>
    <select id="pdf-redim-formato" aria-label="Formato de página"></select>
    <div class="pdf-redim-dimensoes">
        <div><label for="pdf-redim-largura">Largura (mm)</label>
            <input id="pdf-redim-largura" type="text" inputmode="decimal" value="210"></div>
        <button id="pdf-redim-vincular" type="button" aria-pressed="false" aria-label="Proporção livre" aria-describedby="pdf-redim-vinculo-ajuda" title="Proporção livre — medidas independentes">
            <svg class="pdf-cadeado-fechado" data-icon="lock" aria-hidden="true" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="5" y="10" width="14" height="11" rx="2"></rect><path d="M8 10V7a4 4 0 0 1 8 0v3"></path></svg>
            <svg class="pdf-cadeado-aberto" data-icon="lock-open" aria-hidden="true" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="5" y="10" width="14" height="11" rx="2"></rect><path d="M8 10V7a4 4 0 0 1 7.5-2"></path></svg>
        </button>
        <div><label for="pdf-redim-altura">Altura (mm)</label>
            <input id="pdf-redim-altura" type="text" inputmode="decimal" value="297"></div>
    </div>
    <p id="pdf-redim-proporcao" aria-live="polite"></p>
    <small id="pdf-redim-vinculo-ajuda">O cadeado vincula as medidas da página final.</small>
    <label for="pdf-redim-orientacao">Orientação</label>
    <select id="pdf-redim-orientacao"><option value="retrato">Retrato</option><option value="paisagem">Paisagem</option></select>
    <label for="pdf-redim-alvo">Aplicar a</label>
    <select id="pdf-redim-alvo"><option value="atual">Página atual</option><option value="intervalo">Intervalo</option><option value="todas">Todas as páginas</option></select>
    <label id="pdf-redim-intervalo-label" for="pdf-redim-intervalo" hidden>Páginas (ex.: 2, 4-8)</label>
    <input id="pdf-redim-intervalo" type="text" placeholder="2, 4-8" hidden>
    <label for="pdf-redim-modo">Encaixe do conteúdo</label>
    <select id="pdf-redim-modo" aria-describedby="pdf-redim-modo-descricao"><option value="ajustar">Ajustar proporcionalmente</option><option value="centralizar">Centralizar sem alterar a escala</option><option value="preencher">Preencher com corte</option></select>
    <small id="pdf-redim-modo-descricao" aria-live="polite">Cabe inteiro; pode sobrar margem.</small>
    <small>Medidas de 0,1 a 5000 mm, com vírgula ou ponto decimal. A área visível original é preservada; não são criadas sangrias. Recursos interativos complexos podem impedir a operação.</small>
    <p id="pdf-redim-aviso" role="status"></p>
    <button id="pdf-redim-prever" type="button" disabled>Pré-visualizar</button>
    <button id="pdf-redim-impressao" type="button" hidden>Continuar em modo de impressão</button>
    <button id="pdf-redim-cancelar" type="button" disabled>Cancelar</button>
    <button id="pdf-redim-aplicar" type="button" disabled>Aplicar alterações</button>
    <p id="pdf-redim-resultado" role="status"></p>
    <details><summary>Formatos personalizados</summary>
        <label for="pdf-formato-nome">Nome do formato</label>
        <input id="pdf-formato-nome" type="text" maxlength="80" placeholder="Ex.: Formato 4">
        <small>Use a largura e altura acima. Cadastrar ou editar um formato não altera o PDF.</small>
        <button id="pdf-formato-salvar" type="button">Salvar formato personalizado</button>
        <button id="pdf-formato-atualizar" type="button" disabled>Atualizar / renomear formato</button>
        <button id="pdf-formato-excluir" type="button" disabled>Excluir formato</button>
        <p id="pdf-formatos-status" role="status"></p>
    </details>
</section>`;
