# Autoscroll das miniaturas — correção sobre a experimental.32

Data: 08/11/2026. Base `d5a5bdf27836e37074c164a68cdd861e54b84342`, branch `main`, árvore inicialmente limpa. Escopo autorizado: diagnosticar e corrigir o autoscroll inferior, sem outra experimental, build, instalação, impressão, commit ou push. A instalação .32 permanece intacta; esta correção está apenas no checkout para avaliação.

## Reprodução e causa

Reproduzido antes de mudar produção em Chrome e Electron, com PDF sintético de100 páginas, viewport1440×1000 e arraste nativo interceptado pelo CDP. O painel tinha topo431,5625, fundo985, altura útil553px e conteúdo16600px. O ponteiro foi levado ao centro horizontal da borda inferior, Y981. Após3s, `scrollTop` continuava0, havia origem de arraste, mas nenhum `dragover` chegara ao painel e não havia quadro de animação.

`document.elementFromPoint` mostrou o alvo real: **`p#pdf-status` com “PDF carregado.”**, sobreposto à borda. A regra de apresentação compacta posiciona esse aviso absolutamente no rodapé do editor, com z-index4. O retângulo das miniaturas continua atrás dele. Portanto o gesto atingia um irmão do painel, que não recebia o evento necessário para iniciar a rolagem. O resultado explica exatamente a espera histórica “Autoscroll inferior nativo”, sem atribuir a falha ao cálculo de scroll ou ao indicador.

Evidências anteriores, ignoradas: `.desktop-test/autoscroll/before-chrome/diagnostic.json` e `before-electron/diagnostic.json`, capturas e logs correspondentes. A primeira repetição após o ajuste atingiu10056px em3s no Chrome, usando a mesma posição do ponteiro. Os limites e os extremos são conferidos pelo novo roteiro permanente.

## Correção mínima

Em `public/css/theme.css`, o status recebe `pointer-events:none` **somente enquanto existe uma miniatura com a classe de arraste no mesmo editor**. O aviso continua visível e anunciado; o gesto passa para o painel por baixo dele. Quando o arraste acaba, a regra deixa de se aplicar automaticamente e o status recupera a interação normal.

Não houve mudança de layout, dimensões, velocidade, região de ativação ou cálculo de inserção. Nenhum listener, timer ou estado de produção foi acrescentado. A borda superior e o ciclo de cancelamento existentes foram preservados e receberam regressões adicionais. Não foi necessário modificar `pdf-thumbnail-drag.js`, `pdf-thumbnails.js`, PDFWorkspace ou motores PDF.

## Testes e limites da evidência

- `node --test tests/pdf-thumbnail-drag.test.mjs tests/pdf-thumbnails.test.mjs`:20 aprovados. Casos8/100/350, limites inferior/superior, indicador/destino, parada no centro/saída, dragend com quadro ativo, remoção dos listeners, virtualização, pausa, substituição, encerramento e cancelamento de render.
- `npm test`: **441 aprovados**, zero falhas/cancelamentos/skips;437 anteriores+4 novos testes.
- Roteiro permanente `tests/pdf-thumbnail-drag-browser.mjs`, com helper em `tests/helpers/pdf-thumbnail-drag-browser.mjs`: editor real e arraste nativo, PDFs sintéticos8/100/350 em1024×600,1440×1000 e1366×768. Primeira página até o fim, última até o início, referências e ordem do PDF serializado, originais intactos, janela virtual de no máximo16 itens/canvases.
- O mesmo roteiro cobre parada no centro e retomada, cancelamento, saída/soltura fora da lista sem aplicar o último indicador, suspensão/retomada da aba, troca de documento e fechamento durante arraste. Confere ausência de quadro pendente e ausência dos listeners do painel destruído pela API de inspeção do navegador. Só utiliza perfis de teste isolados.

O CDP exige atualizar o destino antes de soltar após virtualização. Para cancelar o arraste interceptado, o roteiro o leva a uma região neutra antes de enviar `dragCancel`, como já faz a suíte geral: sem isso o navegador pode conservar seu autoscroll nativo mesmo com o quadro da aplicação cancelado. O cancelamento direto do quadro ativo e a retirada dos listeners também são verificados sem essa particularidade no Node. Ensaios exploratórios sem interceptação mostraram limitações de sincronização do CDP; não motivaram alteração do código de produção nem contam como homologação manual.

Resultados finais e capturas ficam em `.desktop-test/autoscroll/chrome-approved/`, `electron-approved/` e respectivos logs. Os primeiros ajustes dos próprios ensaios (captura da instância, atualização do alvo CDP após virtualização e inspeção do painel após fechamento) não representam alterações funcionais do editor.

**Chrome e Electron dedicados aprovados integralmente**, com as três quantidades/tamanhos de janela e todos os cenários de encerramento acima. Para reproduzir na raiz da aplicação:

```powershell
npm test
node tests/pdf-thumbnail-drag-browser.mjs
$env:THUMB_DRAG_ELECTRON='1'
node tests/pdf-thumbnail-drag-browser.mjs
Remove-Item Env:THUMB_DRAG_ELECTRON
```

`THUMB_DRAG_RESULTS` permite outra pasta de evidências; `THUMB_DRAG_LIFECYCLE_ONLY=1` restringe o ensaio ao ciclo de cancelamento/encerramento durante diagnóstico. Os testes completos acima não usaram essa restrição. Perfis temporários e artefatos ficam em `.desktop-test/autoscroll/`, sem afetar o perfil instalado.

## Suíte geral e falha fora do escopo

`node tests/pdf-browser-smoke.mjs` ultrapassou as verificações antigas de autoscroll inferior/superior e cancelamento e avançou até `tests/helpers/pdf-actualtext-browser.mjs`, espera **“ActualText validado, aguardando consentimento”**, chamada em `pdf-browser-smoke.mjs:744`. Duas execuções terminaram por timeout nessa etapa; a segunda foi executada sem outra bateria de navegador concorrente e repetiu o mesmo resultado. O roteiro não foi alterado para contornar o erro. Evidências: `.desktop-test/autoscroll/browser-full.log` e `browser-full-isolated.log`; os cenários posteriores à falha não contam como executados. Não se declara a suíte geral totalmente aprovada ou o caminho ActualText corrigido por esta tarefa. A causa desse timeout não foi investigada fora do escopo.

Sintaxe dos três novos JS/MJS, `git diff --check` e inventário final conferidos. Apenas CSS de produção e testes/documentação desta tarefa estão pendentes; staging vazio. Não ficaram processos dos roteiros dedicados identificados pelos seus perfis de teste. Os controladores de arraste/miniaturas foram comparados byte a byte com a cópia anterior e permanecem idênticos.

## Preservação e próximos passos

Permanecem inalterados Step & Repeat, N-up em células/compacto, livretos, impressão Windows, redimensionamento vetorial, Word/LibreOffice privado .32 e preparação raster. Não foram alterados drivers, preferências de impressoras, perfis reais ou documentos de clientes. TrimBox/BleedBox e zoom não foram tratados. Não foi gerada ou instalada a .33.

Arquivos da tarefa: `public/css/theme.css`; três testes/helpers acima; este relatório e `docs/ESTADO_PROJETO.md`. A correção elimina a interceptação reproduzida, sem declarar eliminada qualquer futura falha de arraste em cenários não testados. A avaliação manual permanece necessária antes de decidir empacotar.

## Roteiro manual curto

1. Usar o aplicativo em desenvolvimento com um PDF sintético ou cópia de trabalho de8/100/350 páginas. A instalação .32 existente ainda contém o CSS anterior.
2. Com o status visível, arrastar a primeira miniatura até a borda inferior e manter até alcançar a última; soltar após ela. Conferir sequência e total.
3. Arrastar a última até a borda superior e voltar ao início. Sair da faixa de ativação: a rolagem deve parar; aproximar novamente: deve retomar.
4. Cancelar com Escape e repetir soltando fora do painel. Não deve haver movimento aplicado nem rolagem persistente. Repetir alternando/fechando a aba e abrindo outro documento, conferindo as proteções de alterações pendentes.
5. Repetir com janela menor/maior e seleção múltipla habitual. Exportar uma cópia deliberadamente para conferir ordem, sem impressão física.
