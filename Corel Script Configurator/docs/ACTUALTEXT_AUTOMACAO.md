# ActualText — interrupção da automação após a correção de autoscroll

Data: 08/11/2026. Base `d5a5bdf27836e37074c164a68cdd861e54b84342/main`. Os seis arquivos pendentes do autoscroll foram copiados e conferidos por SHA256 em `.desktop-test/actualtext-audit/before/` antes desta investigação. Nenhuma alteração foi descartada ou refeita.

## Teste e sequência exatos

As duas interrupções anteriores estavam em `tests/helpers/pdf-actualtext-browser.mjs:73`, chamado por `tests/pdf-browser-smoke.mjs:744` (linhas do código antes desta tarefa). A espera tinha o título **“ActualText validado, aguardando consentimento”** e exigia `#pdf-redim-impressao.hidden === false`.

1. Criar o PDF sintético `actualtext-redundante.pdf`, com ActualText redundante na árvore marcada e sem navegação adicional.
2. Conferir pelo validador a redundância e comparar a cópia sem geometria: texto igual e2.073.600 componentes RGBA sem diferenças. Essa etapa já passava.
3. Importar esse original na interface, abrir Redimensionar, selecionar Ajustar/todas/210×297mm e clicar Pré-visualizar, com `modoImpressao:false`.
4. Esperar a oferta de preparar uma cópia removendo a estrutura. O teste interpretava essa oferta como o único resultado correto.

## Reprodução isolada e causa confirmada

Antes de alterar o teste, a mesma sequência foi executada em Chrome com perfil isolado. Após22s, a mensagem era **“Prévia temporária … Estrutura marcada e ActualText preservados, sem rasterização ou descarte”**, Aplicar estava habilitado, nenhuma edição havia sido aplicada e o botão de descarte continuava oculto. O registro de diálogos estava vazio. O retorno da sessão indicava `marcacaoPreservada:true`, `preparadoParaImpressao:false` e `removidos:[]`.

Evidência: `.desktop-test/actualtext-audit/isolated-before/diagnostic.json`, captura e log correspondente. O bloqueio do campo de abrir outro arquivo durante a prévia é parte do ciclo transacional existente, não demonstra que o motor continuou ocupado indefinidamente.

**Causa: expectativa desatualizada do teste.** A experimental.29 acrescentou o caminho restrito de preservação da estrutura e do ActualText em `pdf-resize-engine.js`, função `preservarMarcacao`. Quando ele se aplica, nada dessa estrutura é descartado, portanto a interface não oferece nem pede consentimento para removê-la. O teste antigo ainda esperava o fluxo anterior em todos os documentos. A condição esperada não seria satisfeita aumentando o timeout.

| Hipótese | Resultado da investigação |
|---|---|
| Teste desatualizado | Confirmado: exigia oferta de descarte em uma prévia que preserva a estrutura desde a .29. |
| Diálogo não tratado | Não nesse ponto: nenhum diálogo estava aberto ou havia sido solicitado. A espera era pelo botão da oferta. |
| Corrida/espera indefinida de produção | Não reproduzida: a prévia concluiu, retornou resultado válido e habilitou Aplicar. A espera indefinida era da automação. |
| Defeito na interface/proteção ActualText | Não demonstrado. O comportamento coincide com o motor, a documentação .29 e as regressões de preservação já existentes. |

## Correção somente dos testes

`tests/helpers/pdf-actualtext-browser.mjs` agora separa oito sintéticos com expectativas explícitas:

- Quatro sem navegação: redundante, essencial, vetorial e misto. Devem preservar a marcação, oferecer prévia válida, permitir cancelar sem pendência e manter StructTreeRoot/valores ActualText após Aplicar/exportar/reabrir. Preservar texto essencial não é autorizar seu descarte.
- Quatro com navegação adicional, exigindo o caminho de preparação anterior. Só o redundante pode chegar ao descarte: mantém a prova visual/textual e o diálogo nativo, testa recusar, aceitar, cancelar a prévia e exigir novo consentimento. Essencial/vetorial/misto continuam recusados pela prova e pela interface, sem oferta de descarte ou alteração pendente.

As verificações originais de dimensões A4, streams intactos, igualdade prévia/aplicação, pixels exportados, texto extraível, cancelamento e reabertura foram mantidas. A asserção de ausência de StructTreeRoot passou a distinguir o descarte confirmado da preservação, com verificação adicional de que os valores ActualText permanecem intactos. O original real opcional pode seguir qualquer caminho legítimo; não foi fornecido nem executado nesta tarefa.

Novo `tests/pdf-actualtext-browser.mjs`: executa isoladamente o mesmo helper da suíte geral, sem instalar dependências. Perfil/pasta de saída próprios; diálogo nativo inspecionado por tipo e conteúdo e respondido explicitamente em cada cenário, sem substituir `window.confirm` ou aceitar diálogos desconhecidos. Captura o viewport real, como o roteiro geral, para comparar pixels sem deduzir escala das dimensões arredondadas do canvas. Os primeiros ajustes desse novo harness (worker PDF.js e viewport) foram somente da preparação do ensaio, sem relaxar tolerâncias.

Nenhum arquivo de produção foi modificado nesta investigação. Em especial, não houve alteração de `preservarMarcacao`, do validador de redundância, do diálogo de confirmação, dos motores de redimensionamento/imposição/impressão ou do Word privado.

## Verificações e reprodução

- Reprodução anterior isolada: confirmou prévia válida e ausência de diálogo após22s, mantendo impossível a condição antiga.
- `node tests/pdf-actualtext-browser.mjs`: **aprovado**, oito cenários; quatro preservações/exportações, um descarte com recusa/aceites nativos e três bloqueios. Originais reais opcionais não executados.
- `node --test tests/pdf-actualtext.test.mjs tests/pdf-tagged-resize.test.mjs tests/pdf-resize-ui.test.mjs`: **12 aprovados**, zero falhas/skips.
- `npm test`: **441 aprovados**, zero falhas/cancelamentos/skips. Não foram removidos testes ou marcados skips.
- `node tests/pdf-browser-smoke.mjs`: **suíte geral Chrome aprovada até o fim**, saída0, incluindo autoscroll inferior/superior, os oito cenários ActualText atualizados, biblioteca persistente, assinaturas/CMS sintéticos e cópias para impressão. Nenhuma outra falha permaneceu nessa execução. Original real opcional de ActualText não foi fornecido nem testado; a mensagem geral do harness não amplia esse alcance.

O isolado registrou três confirmações nativas: uma recusada e duas aceitas deliberadamente para o caso de descarte redundante, com novo consentimento após cancelar a prévia. Não houve exceção JavaScript capturada. Sintaxe dos dois scripts novos/alterados e `git diff --check` aprovados. O HEAD permanece `d5a5bdf`, staging vazio; nove caminhos pendentes no total (seis anteriores e três acrescentados nesta investigação).

```powershell
node tests/pdf-actualtext-browser.mjs
node --test tests/pdf-actualtext.test.mjs tests/pdf-tagged-resize.test.mjs tests/pdf-resize-ui.test.mjs
npm test
node tests/pdf-browser-smoke.mjs
```

Evidências ignoradas em `.desktop-test/actualtext-audit/`: `isolated-before/`, `chrome-focused-final.log`, `chrome-*/results.json`, `node-affected.log`, `node-full.log` e `browser-full.log`. O novo runner informa a pasta de cada execução e gera PDFs apenas sintéticos. O roteiro usa a expressão “modo de impressão” para preparar uma cópia PDF; não envia trabalhos a qualquer impressora.

## Preservação e limites

CSS, testes e relatório de autoscroll anteriores foram preservados. `ESTADO_PROJETO.md` recebe somente o novo registro de continuidade. Os seis arquivos anteriores permanecem pendentes; este trabalho acrescenta apenas automação e documentação de ActualText. Não houve correção adicional de autoscroll, caixas, zoom ou qualquer motor de produção.

Não foi gerada nem instalada a .33, executada impressão física/virtual, criado commit ou realizado push. Aprovação dos testes não substitui avaliação manual da futura versão nem homologação universal de PDF/UA, fontes, cor/RIP ou de documentos reais.
