# 4PrintShops — Estado do projeto e passagem de contexto

**Atualizado em:** 29/10/2026<br>
**Fonte:** histórico de desenvolvimento informado pelo usuário; detalhes técnicos devem ser conferidos nos arquivos reais.  
**Ambiente:** Windows, Node.js, servidor local em `http://localhost:3000`, iniciado por `node app.js` na raiz.

## 1. Objetivo geral
Criar uma central de ferramentas para gráficas (4PrintShops), inicialmente com três módulos:
- **Adesivos / CorelDRAW:** configurar scripts por controles visuais, com parser e geração de scripts.
- **QR Codes:** futuramente importar planilhas XLSX de contatos, mapear campos, gerar vCard em lote e exportar PNG/SVG/ZIP.
- **Gerenciador de PDFs:** prioridade atual, para reduzir dependência do Acrobat na gráfica.

O plugin da máquina de produção recém-chegada já realiza a montagem automática de adesivos; essa funcionalidade ficou em espera.

## 2. Estrutura conhecida
- Raiz `app.js`: entrada Node.js; **não confundir** com `public/js/main.js` (controlador da interface).
- `src/server.js`, `repository.js`, `parser.js`, `model.js`, `generator.js`: backend de scripts CorelDRAW.
- `public/js/main.js`, `sidebar.js`, `toolbar.js`, `renderer.js`, `state.js`: interface e estado.
- `public/js/tools/qrcodes.js`: módulo QR Codes, conectado ao menu; a importação XLSX ainda não foi implementada.
- `public/js/tools/pdfs.js`: módulo principal de PDFs; já ficou grande e teve problemas com código antigo duplicado.
- `public/js/tools/pdf-interface.js`: estrutura visual do editor, ferramentas contextuais e recolhimento do painel direito.
- `public/js/tools/pdf-preview.js`: tarefas de carregamento/renderização do PDF.js, canvas principal e encerramento idempotente.
- `public/js/tools/pdf-visual-document.js`: mapa imutável de identidades/rotações do workspace sobre uma fonte PDF.js; permite girar/mover sem recarregar essa fonte. Extração consulta os bytes atuais do workspace.
- `public/js/tools/pdf-document-info.js`: resumo técnico e atualização dos itens de dimensões que mudaram, preservando os elementos existentes.
- `public/js/tools/pdf-resize-engine.js`: transformação vetorial, caixas e anotações suportadas; trabalha numa cópia transacional fornecida pelo workspace.
- `public/js/tools/pdf-resize-interface.js`, `pdf-resize-ui.js`, `pdf-resize-session.js`: painel contextual, biblioteca de formatos e ciclo de prévia/cancelamento/aplicação.
- `public/js/tools/pdf-formats.js`, `pdf-format-storage.js`: formatos padrão/customizados e persistência isolada em localStorage, com fallback em memória.
- `public/js/tools/pdf-thumbnails.js`: miniaturas reais em uma janela virtual, seleção e liberação dos canvases fora da janela.
- `public/js/tools/pdf-thumbnail-drag.js`: arraste nativo, indicador de inserção e rolagem nas bordas; solicita o movimento ao controlador sem manipular bytes PDF.
- `public/js/tools/pdf-thumbnail-job.js`: cancelamento e prazo por tarefa de miniatura, sem depender da conclusão de uma promise antiga para liberar a fila.
- `public/js/tools/pdf-protection.js`: confirmação de descarte, indicador e listener condicional de `beforeunload`, consultando as pendências do workspace.
- `public/js/tools/pdf-workspace.js`: classe `PDFWorkspace`, usando pdf-lib; métodos `abrir`, `adicionarArquivos`, `getTotalPaginas`, `girarPagina`, `excluirPaginas`, `reordenarPaginas`, `moverPagina`, `obterBytes`.
- `public/js/tools/pdf-editor.js`: exportação de PDFs girados/extraídos e lógica de operações legadas.
- `public/js/tools/pdf-merge.js`: ferramenta modal com lista, subir/descer/remover e ação **Adicionar ao documento**, integrada ao workspace por callback.
- `public/vendor/pdfjs/` e `public/vendor/pdf-lib/`: cópias locais de bibliotecas de PDF.
- `modelos/`: scripts de origem CorelDRAW. Não sobrescrever arquivos de produção sem autorização.

## 3. Funcionalidades de PDF e situação de validação
O usuário informou que as funcionalidades anteriores e o redimensionamento passaram nos testes reais realizados e estão aprovados. Em 29/10/2026, confirmou testes com PDFs reais da gráfica, incluindo PDF/X-1a, PDF/X-3, PDF/X-4, Acrobat 9, textos, curvas e páginas com dimensões diferentes. Os testes automatizados Node/Chrome e os limites de compatibilidade permanecem registrados abaixo; a aprovação manual refere-se aos arquivos e cenários relatados pelo usuário.
- Importar e visualizar PDF e trocar de arquivo sem precisar recarregar a página.
- Navegar pelas páginas.
- Ver número de páginas e dimensão em milímetros; identificar diferentes dimensões e avisar pelo botão Informações.
- Girar páginas no `PDFWorkspace` e exportar PDF com alterações.
- Excluir páginas no `PDFWorkspace` e atualizar a pré-visualização **sem download automático**.
- Acumular rotação e exclusão até clicar em **Salvar cópia**.
- Extrair páginas por números/intervalos; **atualmente gera download imediato** (funcionamento legado).
- Unir/adicionar um ou mais arquivos, ordenando a lista: acrescenta páginas ao workspace aberto ou inicia um novo documento, **sem download automático**. Mantém rotações/exclusões existentes; exportação por **Exportar / Salvar cópia**. Testes Node e Chrome headless passaram; arquivos reais pendentes.
- Interface minimalista com importar, informações e exportar no topo, resumo técnico, ferramentas contextuais à direita e opção de recolher o painel para ampliar a visualização.
- Miniaturas reais à esquerda, navegáveis, com destaque da página exibida e atualização após giro/exclusão/união. Renderização sob demanda em baixa resolução e até 16 itens na janela virtual; documento sintético de 350 páginas testado.
- Reorganização pelas miniaturas: arrastar/soltar com linha de inserção ou usar ↑/↓ para mover a página atual. A operação atualiza workspace, numeração, prévia e miniaturas, mantendo a identidade da página exibida e sinalizando pendências somente quando a ordem muda. Sem download automático; teste nativo em Chrome headless com 350 páginas passou.
- Rolagem independente por roda/touchpad, autoscroll pelas bordas visíveis e **Mover para posição** com validação, inclusive para destinos fora da janela virtual. Erros de miniatura mostram **Tentar novamente**, sem interromper as demais. Roda e arraste foram testados no Chrome headless; touchpad físico e PDFs reais ainda precisam de validação.
- Giro e movimentação reutilizam a fonte PDF.js e a identidade da página; mover a página exibida conserva seu canvas. Miniaturas têm cache LRU de até 32 canvases, com até 16 itens na janela virtual. Exclusão mantém recarregamento completo, reaproveitando dimensões/prévias válidas; importação/união invalidam o cache. Medições e limites registrados na entrada de otimização abaixo.
- Redimensionar: A3/A4/A5/A6, medidas personalizadas com vírgula/ponto, orientação, página atual/intervalo/todas e três modos (ajustar, centralizar 100%, preencher com corte). Prévia real temporária, Cancelar e Aplicar; somente a aplicação efetiva gera pendência. Formatos personalizados persistentes, editáveis e excluíveis com confirmação. Passou em Node/Chrome e nos testes reais relatados pelo usuário; recursos bloqueados e roteiro permanecem em `docs/REDIMENSIONAR_PDFS.md`.
- Proteção contra perda de alterações não exportadas: confirmação ao sair para outra ferramenta ou substituir o PDF, indicador próximo à exportação e `beforeunload` apenas enquanto houver pendências. Abrir para visualização não gera aviso; exportar a versão atual limpa a pendência, e editar novamente a reativa. Validado em Node e Chrome headless; conferir os limites do download/navegador na entrada mais recente abaixo.
- Ferramenta "Unir PDFs" foi movida de painel expansível para botão compacto que abre uma janela/modal. O usuário confirmou que funciona.
- O usuário confirmou que o `pdfs.js` revisado está funcionando após correção de um trecho duplicado de rotação fora da função.

## 4. Bugs importantes já encontrados
- Erro de troca de arquivos por gerenciamento inadequado do objeto `.destroy()`: foi corrigido usando a tarefa de carregamento (`tarefaPDF`) e seu `.destroy()`.
- Erros de sintaxe por `try` sem `catch` e chaves sobrando: ocorreram em alterações incrementais de `pdfs.js`.
- Trecho antigo duplicado `workspace.girarPagina(paginaAtual, graus)` fora da função `girarPagina`: causava `graus` indefinido e impedia que os botões funcionassem. Foi removido em arquivo corrigido; usuário validou o funcionamento.
- Evitar retornar a dupla aplicação de rotações (`rotacoes` temporárias + workspace). A meta é que workspace seja fonte de verdade.

## 5. Funcionalidades ainda não feitas
- Ampliar a cobertura de arquivos e recursos ainda não contemplados nas validações realizadas, mantendo a aprovação dos cenários já testados.
- Seleção múltipla visual (reorganização individual por arraste ou botões já implementada).
- Ampliar a compatibilidade do redimensionamento para recursos atualmente bloqueados (formulários/aparências complexas, destinos internos, marcadores e estrutura marcada), somente com estratégia e testes de fidelidade específicos.
- Histórico de desfazer/refazer.
- Inspeção confiável de espaços de cor RGB/CMYK/cores especiais, perfis ICC, sangrias, fontes e resolução efetiva das imagens. Atualmente mostrar "Não verificado".
- Melhorar extração para abrir outro workspace/aba em vez de baixar automaticamente.
- QR Codes: XLSX, seleção de colunas/campos, validação, geração em lote e exportação.

## 6. Prioridade e próximo passo acordado
**Consolidar a versão de redimensionamento aprovada nos testes manuais relatados pelo usuário e aguardar a próxima melhoria solicitada.** Os 78 testes Node e o roteiro Chrome headless passaram; as limitações de compatibilidade permanecem. Fluxo disponível no código:
1. Importar PDF existente ou iniciar um documento com vários PDFs.
2. Selecionar "Unir PDFs" no painel direito e usar seu botão para abrir a janela.
3. Escolher/ordenar arquivos e clicar em **Adicionar ao documento** (sem exportação automática).
4. Conferir o total e as miniaturas; navegar por clique ou Anterior/Próxima. Reorganizar arrastando uma miniatura, usando ↑/↓ ou informando **Mover para posição** no painel esquerdo. Rolar a lista independentemente da prévia; durante o arraste, aproximar o ponteiro das bordas para percorrer páginas distantes. Escolher Rotacionar ou Organizar páginas no painel direito para continuar editando.
5. Em **Redimensionar**, definir formato/escopo/modo, Pré-visualizar e conferir avisos/dimensões. Cancelar preserva as edições anteriores; Aplicar atualiza o workspace. Continuar editando ou baixar por **Exportar/Salvar cópia**.
6. Recolher/reabrir o painel direito quando necessário. Imposição, seleção múltipla, criação de sangrias/marcas e análise de cores permanecem fora desta etapa.

**Cuidados nas próximas etapas:** preservar a rotina compartilhada de atualização do PDF.js, o bloqueio de operações simultâneas e os listeners ligados aos elementos da instância. A reutilização visual é explícita e limitada a giro/movimentação; novas operações de conteúdo devem usar o caminho completo, invalidando fontes/caches. `PDFDocumentoVisual.getData()` consulta o workspace, pois os bytes da fonte PDF.js podem ter a ordem/rotação anterior. `renderPDFs()` retorna uma função de encerramento com `confirmarSaida()`; `main.js` consulta a confirmação antes de limpar o módulo ou atualizar o menu ativo. Miniaturas devem parar antes da destruição dos documentos PDF.js. `PDFWorkspace.alterado` é derivado das revisões atual/exportada; novas operações devem chamar `registrarAlteracao()`, sem atribuir diretamente a `alterado`. A extração ainda baixa diretamente e não quita a exportação do documento completo. O nome informado na união é usado ao iniciar um documento; ao acrescentar páginas, o documento aberto mantém seu nome.

## 7. Fluxo de trabalho que o usuário prefere
- Mudanças pequenas, um passo por vez, explicações claras de qual arquivo foi alterado.
- Testar no navegador com PDFs reais, mantendo os arquivos originais intactos.
- Para testes de sintaxe, usar `node --check` em JS modificado.
- Priorizar funcionalidades de uso diário da gráfica e minimizar riscos de regressão.
- Arquivos JavaScript baixados pelo navegador já foram bloqueados; `.txt` funcionou como alternativa para transferência, com renomeação local.

## 8. Registro de trabalho futuro (Codex deve preencher)

### 28/10/2026 — Manutenção preventiva dos eventos de PDFs
- **Objetivo:** corrigir somente o acúmulo de listeners ao sair e voltar à ferramenta PDFs.
- **Arquivos alterados:** `public/js/tools/pdfs.js` e `docs/ESTADO_PROJETO.md`.
- **Causa confirmada:** `main.js` substitui os filhos de `#app`, mas preserva esse elemento. Cada chamada de `renderPDFs()` registrava nele um novo listener delegado de extração/exclusão, mantendo os handlers e os estados das entradas anteriores.
- **O que foi implementado:** o mesmo listener passou a ser registrado em `#pdf-conteudo`, elemento próprio da ferramenta, recriado a cada entrada. Adicionado comentário explicando o ciclo de vida. Sem alterações no HTML/CSS, nas operações de PDF, no estado legado de rotação ou nos outros módulos. União continua com download direto, sem integração ao workspace nesta etapa.
- **Testes executados e resultados:**
  - `node --check public/js/tools/pdfs.js`: passou.
  - Simulação em memória com Node.js/VM e DOM simplificado, sem PDF aberto: após cinco entradas, a versão anterior executava o handler cinco vezes por clique em Extrair/Excluir; a versão corrigida executou uma vez por clique. Clique em Informações não disparou o handler de seleção e manteve a alternância do painel. Imports/serviços PDF foram substituídos por stubs nessa simulação; ela verifica os eventos, não o processamento de PDFs.
  - `node teste-generator.js` e `node teste2.mjs`: concluíram com código 0; são scripts avulsos, não testes de PDFs.
  - `npm.cmd test`: terminou com código 1, pois `package.json` contém apenas o comando padrão `Error: no test specified`; não há suíte configurada.
  - `git diff --check`: passou. Nenhum reset/checkout ou descarte de alterações foi executado.
- **Pendências / limites:** não foi realizado teste em navegador nem com PDFs reais nesta etapa. Validar manualmente importação/troca de arquivos, navegação, rotação, extração, exclusão, união e Salvar cópia após alternar entre PDFs e outras ferramentas várias vezes; cada ação deve executar uma única vez e preservar os originais.
- **Próximo passo recomendado:** validar o fluxo no navegador antes de avançar para a integração da união ao workspace, que permanece planejada e não foi iniciada.

### 28/10/2026 — União integrada ao documento de trabalho
- **Objetivo:** acrescentar os PDFs escolhidos no modal ao `PDFWorkspace`, sem download automático, permitindo continuar editando até Salvar cópia.
- **Arquivos alterados:** `public/js/tools/pdf-workspace.js`, `public/js/tools/pdf-merge.js`, `public/js/tools/pdfs.js`, `tests/pdf-workspace.test.mjs`, `tests/pdf-integration.test.mjs` e este registro. `pdf-editor.js` foi inspecionado e preservado.
- **O que foi implementado:**
  - `adicionarArquivos()` prepara uma cópia do documento atual em memória e acrescenta os arquivos na ordem escolhida; somente substitui o workspace quando todos os arquivos forem válidos. Preserva páginas mantidas, rotações, nome e metadados do documento existente. Sem documento aberto, cria um novo e usa o nome informado no modal.
  - Modal mantém seleção, ordenação, remoção e fechamento, aceita um ou mais PDFs e usa **Adicionar ao documento**. Após sucesso, limpa a lista e fecha a janela para mostrar o resultado. Falhas mantêm a lista para correção/repetição.
  - Importação, rotação, exclusão e união reutilizam `atualizarVisualizacaoWorkspace()` para atualizar PDF.js, resumo, dimensões, alertas e navegação. A nova tarefa é preparada antes de destruir a anterior; falhas de carregamento na importação/união restauram o workspace anterior.
  - Um bloqueio compartilhado impede importar, girar, excluir, extrair, unir ou exportar simultaneamente. Navegações pendentes são invalidadas ao iniciar uma operação. Elementos são capturados por instância, mantendo a correção anterior de listeners e evitando que uma tarefa antiga modifique a nova tela ao sair e voltar.
  - Salvar cópia exporta o workspace completo por solicitação explícita. Extração mantém o download legado e usa a versão exibida, incluindo páginas adicionadas e edições anteriores. O mapa legado de rotações permanece; HTML/CSS da interface geral não foram alterados.
- **Testes executados e resultados:**
  - `node --test tests/pdf-workspace.test.mjs tests/pdf-integration.test.mjs`: 6 testes passaram. Verificados: início sem PDF aberto; ordem/remoção/fechamento no modal; acréscimo ao PDF existente; união após giro/exclusão; edição posterior; bytes exportados e extraídos com a sequência e rotação esperadas; ausência de download na união; preservação dos originais sintéticos; lote inválido sem adição parcial; falha simulada da prévia com recuperação sem duplicação; bloqueio de cliques repetidos/operações concorrentes; navegação pendente e isolamento ao recriar a ferramenta.
  - Os testes usam as bibliotecas locais reais pdf-lib e PDF.js para ler PDFs sintéticos, obter dimensões e conferir os resultados. O DOM é simplificado e a rasterização do canvas é simulada. O teste fornece o caminho local do worker e compatibilidade `Uint8Array.toHex` para o Node 24; essas adaptações não alteram os arquivos de produção. PDF.js emite aviso sobre preferência pelo build `legacy` no Node.
  - Sintaxe dos três módulos JS alterados e dos dois testes conferida com `node --check`; `git diff --check` passou.
  - `node teste-generator.js` e `node teste2.mjs`: concluíram com código 0 (scripts avulsos, fora do fluxo de PDFs).
  - `npm.cmd test`: código 1 porque o script continua sendo o padrão `Error: no test specified`. Executar os testes novos pelo comando `node --test` acima; nenhuma dependência foi instalada.
- **Pendências / limites:** navegador e renderização visual de PDFs reais não foram testados. Validar arquivos reais, inclusive maiores, e conferir o PDF exportado. A adição usa uma cópia temporária em memória. Análises de cores/sangria/fontes/DPI continuam não verificadas; extração ainda baixa diretamente.
- **Próximo passo recomendado:** validar no navegador o fluxo abaixo. Miniaturas e redimensionamento não foram iniciados.
- **Roteiro manual de aceitação:**
  1. Iniciar `node app.js`, abrir PDFs e, sem importar antes, usar Unir PDFs. Selecionar dois arquivos, ordenar/remover itens e clicar Adicionar ao documento. Conferir ausência de download, total, dimensões/alertas e navegação.
  2. Importar um PDF conhecido, girar uma página e excluir outra; adicionar novos arquivos pelo modal. Conferir que as páginas anteriores mantêm as edições e as novas aparecem ao final na ordem escolhida.
  3. Continuar girando/excluindo, extrair páginas e usar Salvar cópia. Reabrir o arquivo exportado e verificar sequência, total e rotações. Conferir que os originais permanecem intactos.
  4. Durante o processamento, conferir que as demais operações ficam bloqueadas. Depois de exportar, fechar/reabrir o modal e alternar entre ferramentas, verificando que cada clique executa uma única ação. O workspace continua sendo recriado ao reabrir a ferramenta, como antes; esta etapa não implementou persistência entre ferramentas.

### 28/10/2026 — Interface contextual e miniaturas reais, em três fases
- **Referência inicial:** Git limpo no commit `75c7ae2`. Nenhum commit novo, push ou comando destrutivo executado nesta tarefa.
- **Objetivo:** interface minimalista com painel contextual direito e miniaturas reais navegáveis à esquerda, preservando a edição acumulada no workspace.
- **Arquivos alterados:** `public/js/tools/pdfs.js`, `public/js/tools/pdf-merge.js`, `public/js/main.js`, `public/css/style.css`, `tests/pdf-integration.test.mjs` e este registro.
- **Módulos novos:** `public/js/tools/pdf-interface.js`, `public/js/tools/pdf-preview.js`, `public/js/tools/pdf-thumbnails.js`.
- **Testes/apoio novos:** `tests/pdf-thumbnails.test.mjs`, `tests/helpers/pdf-dom.mjs`, `tests/pdf-browser-smoke.mjs`.
- **Fase 1 — concluída e verificada antes de prosseguir:**
  - Importar, informações e exportação ficam no topo; resumo técnico preservado.
  - Painel direito com Rotacionar, Organizar páginas e Unir PDFs; só as opções escolhidas ficam visíveis. Modal existente permanece integrado ao workspace. Botão recolhe/reabre o painel e a área principal usa o espaço liberado.
  - Layout e estados visuais separados em `pdf-interface.js`; estilos novos ficam restritos ao editor e conservam o tema atual.
  - Sintaxe e `git diff --check` passaram; 7 testes Node passaram nesta fase.
- **Fase 2 — concluída e verificada antes de prosseguir:**
  - Miniaturas PDF.js em baixa resolução (encaixe de até 112 × 126 pixels), com número da página, clique para navegar e destaque por `aria-current`.
  - Janela virtual mantém no máximo 16 botões/canvases no DOM; fila renderiza uma miniatura por vez. Rolagem reutiliza os itens da janela e libera canvases removidos. Troca de documento invalida e cancela a fila anterior; giro/exclusão/união recriam as miniaturas a partir do PDF atualizado.
  - `PDFPreview` concentra tarefas PDF.js e renderização principal, cancela renders pendentes e memoiza `destroy()` por tarefa. A separação reduziu o controlador `pdfs.js`.
  - `renderPDFs()` fornece um encerramento idempotente; `main.js` aguarda a limpeza antes de montar a próxima ferramenta. Modal é fechado, listeners de scroll/clique e observador de tamanho das miniaturas são removidos, canvases são liberados e documentos são destruídos.
  - Sintaxe e `git diff --check` passaram; 11 testes Node passaram nesta fase.
- **Fase 3 — concluída:**
  - Comando final `node --test tests/pdf-workspace.test.mjs tests/pdf-integration.test.mjs tests/pdf-thumbnails.test.mjs`: **12 testes passaram**, incluindo painel contextual, operações anteriores, janela de 350 páginas, não duplicação de renderizações na mesma janela, seleção, orientação/quantidade após atualização, cancelamento pendente e limpeza mesmo após falha de rasterização.
  - `node tests/pdf-browser-smoke.mjs`: **passou no Chrome headless instalado**. Usa servidor local, perfil temporário e PDFs sintéticos, sem upload. Verificou canvases reais, clique/seleção de miniaturas, ampliação ao recolher o painel, giro, exclusão, união sem download automático, extração, download de exportação e conteúdo do PDF resultante. Navegou até a página 350, substituiu o documento e alternou três vezes entre PDFs e QR Codes, sem erros JavaScript capturados.
  - Capturas reais inspecionadas visualmente. Artefatos da execução final em `%TEMP%/4printshops-browser-u97FMr/` (`01-editor.png`, `02-edicoes.png`, `03-350-paginas.png` e PDFs sintéticos exportados); não fazem parte do repositório.
  - `node --check` nos JS/MJS alterados/novos e `git diff --check`: passaram.
  - `node teste-generator.js` e `node teste2.mjs`: código 0. `npm.cmd test` permanece com código 1 por ser o placeholder `Error: no test specified`; `package.json` não foi alterado e nenhuma dependência foi instalada.
- **Arquitetura / limitações documentadas:**
  - Miniaturas são sob demanda, mas a rotina existente de resumo ainda percorre todas as páginas para comparar dimensões, e cada edição serializa o workspace para atualizar PDF.js. Isso pode custar tempo/memória em PDFs grandes ou complexos; não foi feita uma refatoração desse fluxo nesta tarefa.
  - A janela virtual descarta os canvases fora da área próxima à rolagem; retornar a páginas distantes pode gerar sua miniatura novamente. Não há renderizações simultâneas duplicadas da mesma miniatura.
  - Ao sair da ferramenta, o documento de trabalho é descartado, como antes; exportar antes de trocar de ferramenta. Não foi implementada persistência entre módulos.
  - Browser testado: Chrome headless com PDFs sintéticos. Não houve validação manual com PDFs de clientes, nem cobertura de outros navegadores, documentos protegidos, arquivos muito pesados ou todas as resoluções de tela.
  - RGB/CMYK, sangria, fontes e DPI continuam não verificados. Redimensionamento, imposição, arraste, reordenação e seleção múltipla não foram implementados. `PDFWorkspace` e `pdf-editor.js` não precisaram de alterações.
- **Próximo passo:** validação manual abaixo; aguardar orientação do usuário antes de outras funcionalidades ou commits.
- **Roteiro manual:**
  1. Executar `node app.js` e recarregar a página (Ctrl+F5). Abrir PDFs: conferir ações no topo e alternar Rotacionar, Organizar páginas e Unir PDFs; apenas as opções selecionadas devem aparecer. Recolher/reabrir o painel e conferir o espaço da prévia.
  2. Importar um PDF real e clicar em diferentes miniaturas; conferir destaque, número e visualização principal. Usar Anterior/Próxima e conferir a sincronização. Abrir Informações e verificar dimensões/alertas e cores como não verificadas.
  3. Girar uma página, excluir outra e acrescentar arquivos pelo modal. Conferir os novos totais/miniaturas e a preservação das alterações. Extrair um intervalo e exportar o documento completo; reabrir os resultados para conferir ordem/rotações.
  4. Testar PDF com mais de 300 páginas, rolar até o fim e clicar em páginas distantes. A lista deve permanecer responsiva, carregando prévias próximas à rolagem. Conferir no DevTools que os canvases em `#pdf-miniaturas` permanecem limitados.
  5. Trocar de arquivo enquanto as prévias carregam. Após exportar, sair e voltar várias vezes à ferramenta; conferir ausência de erros no console, canvases antigos ou ações duplicadas. Repetir com os tamanhos de janela usados na produção.

### 28/10/2026 — Proteção contra perda de alterações não exportadas
- **Objetivo:** proteger o PDF em memória ao trocar de ferramenta, importar um substituto ou sair da página, preservando a interface e as operações atuais.
- **Situação inicial:** alterações locais da tarefa anterior ainda sem commit; preservadas. Nenhum commit, push ou comando destrutivo executado nesta etapa.
- **Arquivos alterados nesta etapa:** `public/js/main.js`, `public/js/tools/pdfs.js`, `public/js/tools/pdf-workspace.js`, `public/js/tools/pdf-interface.js`, `public/css/style.css`, `tests/pdf-workspace.test.mjs`, `tests/pdf-integration.test.mjs`, `tests/helpers/pdf-dom.mjs`, `tests/pdf-browser-smoke.mjs` e este registro.
- **Arquivos novos nesta etapa:** `public/js/tools/pdf-protection.js` e `tests/pdf-protection.test.mjs`.
- **Fase 1 — controle de pendências:**
  - O workspace mantém revisão atual e revisão exportada. `alterado` é um getter; rotações, exclusões e uniões chamam `registrarAlteracao()`. Abrir um arquivo para visualizar inicia sem pendências; criar um documento por união já gera pendência.
  - Exportação captura um token com documento/revisão antes de serializar e só marca a versão correspondente após gerar o PDF e solicitar seu download. Um token antigo não quita edições posteriores nem outro documento. Serialização para prévia ou extração não marca o documento completo como exportado.
  - Captura/restauração de estado preserva também revisões em falhas de importação/união. Os 7 testes de workspace passaram antes da integração da confirmação.
- **Fases 2 e 3 — confirmação e navegador:**
  - Modal com o título, mensagem e botões solicitados: **Alterações não exportadas**, **Continuar editando**, **Descartar e sair**. Cancelamento, Escape ou fechamento mantêm o documento; o foco inicial fica em Continuar editando.
  - O menu só muda de ferramenta após a aprovação e a limpeza; clicar novamente em PDFs mantém a instância atual. Requisições repetidas de navegação não criam múltiplos avisos. Respostas atrasadas de Adesivos não sobrescrevem outro módulo ativo.
  - Importar outro PDF usa a mesma confirmação, antes de alterar o workspace. Cancelar preserva nome, páginas, giro, canvas e miniaturas. Sair durante uma operação aguarda a conclusão para avaliar as pendências resultantes.
  - Indicador discreto próximo a Salvar cópia consulta apenas o estado do workspace. `beforeunload` é instalado uma vez enquanto houver pendências, removido ao exportar ou encerrar; as confirmações também são desmontadas no encerramento.
- **Fase 4 — testes executados:**
  - `node --test tests/pdf-workspace.test.mjs tests/pdf-integration.test.mjs tests/pdf-thumbnails.test.mjs tests/pdf-protection.test.mjs`: **23 testes passaram**. Incluem saída limpa/alterada, cancelamento, aprovação, importação protegida, versão exportada, nova edição, falha de exportação, token assíncrono antigo, extração sem quitar pendências, união em andamento e não acúmulo de listeners. Os testes anteriores de edição e miniaturas foram preservados.
  - `node tests/pdf-browser-smoke.mjs`: **passou no Chrome headless** com PDFs sintéticos. Testou os três destinos do menu (Início, Adesivos e QR Codes), menu ativo durante o aviso, preservação do canvas ao cancelar, importação cancelada/aprovada, indicador após exportação/nova edição e contagem real de listeners `beforeunload` em reentradas. Uma tentativa de navegação nativa para outra página acionou o diálogo `beforeunload` do Chrome; cancelar manteve o editor. Também repetiu operações, downloads e documento de 350 páginas.
  - Captura `04-confirmacao.png` inspecionada visualmente; artefatos em `%TEMP%/4printshops-browser-ZTRUiQ/`, fora do repositório, contendo apenas PDFs sintéticos e capturas do teste.
  - Sintaxe JS/MJS e `git diff --check`: passaram. Scripts avulsos `teste-generator.js` e `teste2.mjs`: código 0. `npm.cmd test` continua com o placeholder `Error: no test specified`, código 1; executar os testes pelo comando explícito acima.
- **Limitações:**
  - O navegador controla o texto e a exibição do aviso nativo de fechamento/recarregamento, exige interação prévia e pode suprimi-lo em certas situações. Esse aviso não usa o texto/botões do modal interno. Não há salvamento automático nem recuperação após descarte/encerramento.
  - A API de download por link não informa se o usuário salvou, cancelou ou bloqueou o arquivo no navegador. A versão é marcada após gerar os bytes e iniciar a solicitação de download sem erro; conferir o arquivo na pasta de downloads antes de abandonar o trabalho.
  - Testes de navegador usam Chrome headless e PDFs sintéticos. Fechamento/recarregamento manual, outros navegadores e PDFs reais de produção ainda precisam de conferência. As limitações anteriores de desempenho da análise de dimensões permanecem.
- **Próximo passo / roteiro manual:**
  1. Recarregar com Ctrl+F5, importar um PDF e apenas visualizar. Sair para outra ferramenta: não deve haver aviso.
  2. Girar/excluir uma página ou iniciar um documento pela união. Conferir o indicador e tentar Início, Adesivos e QR Codes: o menu deve permanecer em PDFs enquanto o aviso estiver aberto.
  3. Escolher Continuar editando (e testar Escape): conferir que página atual, miniaturas e edições permanecem. Repetir e escolher Descartar e sair: confirmar a navegação.
  4. Com pendências, importar outro PDF. Cancelar deve preservar o anterior; confirmar deve abrir o novo sem pendências. Testar também um arquivo inválido e verificar a preservação do documento anterior.
  5. Exportar e conferir o download completo. O indicador deve desaparecer; uma nova edição deve reativá-lo. Extrair páginas isoladas não deve quitar pendências do documento completo.
  6. Com pendências e após interagir com a página, tentar recarregar/fechar a aba e cancelar no aviso nativo. Após exportar, repetir: não deve haver aviso de pendência. Reentrar várias vezes no módulo e conferir ausência de avisos/ações duplicados.

### 28/10/2026 — Reorganização de páginas pelas miniaturas
- **Objetivo:** mover páginas por arraste nativo ou controles acessíveis, acumulando a alteração no workspace até Salvar cópia.
- **Situação inicial:** Git limpo no commit `dc9b803`. Nenhum commit, push, instalação de dependência ou comando destrutivo executado nesta etapa.
- **Arquivos alterados:** `public/js/tools/pdf-workspace.js`, `public/js/tools/pdf-thumbnails.js`, `public/js/tools/pdf-interface.js`, `public/js/tools/pdfs.js`, `public/js/tools/pdf-protection.js`, `public/css/style.css`, `tests/pdf-workspace.test.mjs`, `tests/pdf-thumbnails.test.mjs`, `tests/pdf-integration.test.mjs`, `tests/pdf-protection.test.mjs`, `tests/helpers/pdf-dom.mjs`, `tests/pdf-browser-smoke.mjs` e este registro.
- **Módulo novo:** `public/js/tools/pdf-thumbnail-drag.js`. `pdf-preview.js` foi inspecionado e reutilizado sem alteração.
- **Implementação:**
  - `reordenarPaginas()` valida uma permutação completa, rejeitando páginas inválidas/duplicadas/ausentes antes de modificar qualquer documento. Ordem idêntica retorna sem alterar revisão ou pendências. `moverPagina()` monta a permutação para um movimento individual.
  - Reorganização transacional: carrega uma cópia serializada do PDF completo, materializa Resources/MediaBox/CropBox/Rotate herdados e reinsere as mesmas folhas no mesmo contexto PDF. Não usa imagens nem recria o documento por `copyPages()`. Só substitui o workspace após concluir; mudanças concorrentes de revisão são rejeitadas.
  - A API permite inserir uma instância `PDFPage` existente: [PDFDocument.insertPage](https://pdf-lib.js.org/docs/api/classes/pdfdocument#insertpage). Também foi conferido o código local da versão instalada (pdf-lib 1.17.1), incluindo os atributos herdáveis de `PDFPageLeaf`.
  - Arraste mostra uma linha azul entre as páginas, mantém a origem no DOM durante a rolagem e solicita a posição final ao controlador. Fila e janela virtual continuam limitadas a 16 miniaturas; não renderiza todas as páginas para arrastar. Rolagem nas bordas usa quadros de animação cancelados ao encerrar/substituir o documento.
  - Botões ↑/↓ junto às miniaturas movem a página atual, respeitam limites e bloqueio de operações, e restauram o foco após o movimento. A página exibida é acompanhada pela permutação, mesmo quando outra página é movida através de sua posição.
  - O controlador reutiliza a rotina existente de atualização PDF.js/miniaturas, preserva rotações/exclusões/uniões e restaura o workspace/seleção caso a preparação da nova prévia falhe. A seleção textual de números para extração/exclusão é limpa após reordenar, evitando reutilizar índices da sequência anterior.
  - A revisão só muda em uma reorganização efetiva; a proteção de alterações não exportadas continua ativa. A exportação permanece exclusivamente por Salvar cópia.
  - Durante os testes apareceu uma corrida preexistente: o evento `close` de uma confirmação podia chegar depois de uma nova abertura e cancelá-la. `pdf-protection.js` agora ignora esse evento atrasado quando o diálogo está aberto; adicionado teste de regressão.
- **Testes executados:**
  - `node --test tests/pdf-workspace.test.mjs tests/pdf-integration.test.mjs tests/pdf-thumbnails.test.mjs tests/pdf-protection.test.mjs`: **34 testes passaram**. Verificados ordem inválida/sem efeito, falha na cópia, revisão concorrente, páginas giradas, exclusões anteriores, documentos unidos, metadados, identidade da seleção, recuperação de falha de prévia, arraste/posição final, fila/limite de canvases e cancelamento dos eventos/quadros.
  - Testes estruturais em PDFs sintéticos exportados preservaram referências das páginas, anotação de texto, link interno por referência de página, campo de formulário AcroForm simples e recursos/caixas/rotação herdados de um ramo da árvore.
  - `node tests/pdf-browser-smoke.mjs`: **passou no Chrome headless**. Incluiu início de arraste nativo por eventos de mouse, marcador de inserção, botões/foco, página selecionada, movimento sem efeito sem criar pendência, ordem e rotação do arquivo exportado. Moveu a página 1 para o fim de um PDF de 350 páginas, mantendo no máximo 16 canvases e conferindo a sequência de referências do PDF baixado. A rolagem para o destino distante foi feita programaticamente durante o gesto; a rolagem automática nas bordas também foi verificada em teste Node.
  - Captura `05-arraste.png` inspecionada visualmente. Artefatos da execução aprovada em `%TEMP%/4printshops-browser-j3NIGJ/`, contendo apenas capturas e PDFs sintéticos.
  - As primeiras tentativas de automação expuseram problemas de início do gesto/alvo fora do viewport. Ajustados a área interativa da miniatura (conteúdo interno não intercepta o ponteiro), a sequência de mouse e a visibilidade do destino no teste; a execução final usou arraste nativo, não `DragEvent` sintético disparado por JavaScript.
  - Sintaxe JS/MJS e `git diff --check`: passaram. Scripts avulsos `teste-generator.js` e `teste2.mjs`: código 0. `npm.cmd test` permanece como placeholder `Error: no test specified`, código 1.
- **Limitações e preservação de recursos:**
  - A reordenação mantém os objetos de página e o catálogo da cópia completa; os testes cobrem casos sintéticos específicos, não garantem fidelidade de todos os recursos de qualquer PDF. Conferir formulários/anotações reais em leitores apropriados.
  - Rótulos de página (`PageLabels`), destinos expressos por índices numéricos e a ordem lógica de leitura de PDFs marcados não são reindexados. Podem exigir revisão após mudar a sequência física. Referências diretas às páginas foram preservadas no teste de link interno.
  - A serialização reescreve o arquivo; não preserva a validade de assinaturas digitais. Não foi adicionada compatibilidade com documentos protegidos/criptografados, XFA ou outros recursos não suportados pela biblioteca. Recursos já perdidos em uma operação anterior, como cópias durante união, não são recuperados por reorganizar.
  - Cada movimento prepara uma cópia em memória e atualiza a análise de dimensões/visualização existente. PDFs muito grandes/complexos ainda podem consumir tempo e memória; só as miniaturas são virtualizadas. Touch/mobile e outros navegadores não foram validados.
- **Roteiro manual / próximo passo:**
  1. Iniciar `node app.js` e recarregar com Ctrl+F5. Importar uma cópia de teste e arrastar uma miniatura para antes/depois de outra; conferir a linha azul, numeração, seleção e ausência de download.
  2. Selecionar uma página, usar ↑/↓ por teclado (Tab e Enter/Espaço) e conferir foco, limites da primeira/última página e que a mesma página continua exibida ao mudar de posição. Arrastar outra página através da selecionada e repetir a conferência.
  3. Girar, excluir e unir arquivos antes de reorganizar. Exportar e reabrir o resultado, conferindo ordem, conteúdo, rotação e arquivos originais intactos. Se houver anotações, links ou formulários, conferir também esses recursos.
  4. Após exportar, soltar uma página na mesma posição: não deve aparecer pendência. Fazer um movimento real: o indicador e o aviso de saída devem voltar.
  5. Em documento com mais de 300 páginas, arrastar e manter o ponteiro perto das bordas para rolar. Conferir que as miniaturas carregam por proximidade e que o console não apresenta erros ao trocar o arquivo ou sair após exportar.
  6. Validar PDFs reais da gráfica antes de outras funcionalidades. Não foram implementados seleção múltipla, redimensionamento, imposição ou análise de cores.

### 28/10/2026 — Rolagem, movimento direto e recuperação das miniaturas
- **Objetivo:** corrigir dificuldades de navegação e bloqueios ao reorganizar repetidamente, sem alterar a manipulação de bytes do workspace nem iniciar downloads automáticos.
- **Situação inicial:** as alterações da etapa anterior de reorganização já estavam presentes e não commitadas. Foram preservadas integralmente. Nenhum commit, push, comando destrutivo ou instalação de dependência nesta correção.
- **Arquivos alterados nesta correção:** `public/js/tools/pdf-thumbnails.js`, `public/js/tools/pdf-thumbnail-drag.js`, `public/js/tools/pdf-interface.js`, `public/js/tools/pdfs.js`, `public/css/style.css`, `tests/pdf-thumbnails.test.mjs`, `tests/pdf-integration.test.mjs`, `tests/pdf-browser-smoke.mjs` e este registro.
- **Novo módulo:** `public/js/tools/pdf-thumbnail-job.js`. `pdf-preview.js` e `pdf-workspace.js` foram inspecionados e reutilizados sem mudanças nesta correção; as modificações locais anteriores no workspace permanecem.
- **Causas verificadas antes da correção:**
  - A limpeza aguardava diretamente o processamento serial das miniaturas. Uma promise de `getPage()`/render sem conclusão, mesmo após `cancel()`, retinha essa espera e o bloqueio de edição durante a atualização do documento. Reproduzido com uma tarefa controlada que não respondia ao cancelamento.
  - Existia uma janela entre a conclusão do executor e a liberação de sua referência: miniaturas criadas nessa janela podiam ficar pendentes sem um executor para retomá-las. A reprodução criou uma nova janela nesse ponto e encontrou itens que continuavam em carregamento.
  - O cálculo de autoscroll usava a borda física inferior do painel mesmo quando ela ficava abaixo da tela. Reproduzido com painel parcialmente fora do viewport; o ponteiro junto à borda visível não provocava rolagem.
  - Os três testes de regressão falharam no código anterior e passaram após a correção. Isso confirma os mecanismos de falha; sem os PDFs reais que originaram o relato, não é possível atribuir todos os sintomas do cliente a uma única causa específica.
- **Implementação:**
  - Tarefa local de miniatura com interrupção explícita e prazo de 10 segundos para `getPage`/render. Cancelar não depende da promise original terminar. Erros/prazo esgotado viram estado de erro com nova tentativa por clique; o executor prossegue para as próximas páginas.
  - Cada resultado confere documento, geração e identidade do item antes de publicar o canvas. Encerramento/substituição cancelam a tarefa e o arraste; canvases descartados são liberados. Retornos antigos não alteram a nova janela/documento.
  - A fila só é acionada quando há itens pendentes, verifica novamente as pendências ao terminar e cede ao navegador entre tarefas. Mantém a janela limitada a 16 itens, incluindo a origem fixada durante o arraste. Rolagem rápida cancela o trabalho que sai da janela.
  - Painel com rolagem nativa independente, altura ajustada à área visível e ancoragem automática desativada na janela virtual. Autoscroll usa as bordas visíveis e não é interrompido por uma simples troca de alvo causada pela virtualização. Quadros e listeners são removidos no encerramento.
  - **Mover para posição** valida inteiro entre 1 e o total, por botão ou Enter, e reutiliza o movimento transacional já existente. Funciona para destinos não visíveis, mantém a página selecionada e não gera download. ↑/↓ foram preservados.
  - O posicionamento inicial da lista ocorre após criar sua área rolável. Uma prévia principal atrasada respeita a versão da rolagem e não força a lista de volta após uma rolagem manual ou durante o arraste.
  - Workspace, revisões/aviso de perda de alterações, operações anteriores e fluxo de exportação foram preservados.
- **Testes executados e resultados:**
  - `node --test tests/pdf-workspace.test.mjs tests/pdf-integration.test.mjs tests/pdf-thumbnails.test.mjs tests/pdf-protection.test.mjs`: **42 testes passaram**, incluindo as três regressões, prazo esgotado, erro/retry isolado, resposta antiga após substituição, navegação atrasada, movimento direto/posições inválidas, reentradas sem acumular listeners e as operações/proteção anteriores.
  - `node tests/pdf-browser-smoke.mjs`: **passou no Chrome headless** com PDF.js real e PDFs sintéticos de 100 e 350 páginas. Moveu 5→50 fora da janela, conferiu seleção visível, roda do mouse sem rolar a página central, limite de canvases, autoscroll inferior/superior por arraste nativo e cancelamento. Fez quatro movimentos adicionais por documento com renders atrasados, seguidos de giro, exclusão, união e exportação; conferiu sequência de referências de páginas, total e rotações nos PDFs baixados.
  - No mesmo navegador, injetou falha de render em uma miniatura e verificou erro, continuidade das vizinhas e retry; depois injetou render sem conclusão/cancelamento e reorganizou/exportou sem reter o bloqueio. Ao estabilizar, todas as miniaturas presentes estavam prontas ou em erro, sem carregamento permanente. Não houve erros JavaScript capturados. Os fluxos anteriores de importação, extração, exportação e confirmação de saída também passaram.
  - A automação do arraste foi ajustada para atualizar a posição real do mouse ao inverter a direção e terminar fora das bordas antes do cancelamento via CDP. Durante a investigação, o Chrome mantinha seu autoscroll nativo após um cancelamento interceptado; o rastreamento distinguiu essa rolagem de escritas feitas pelo aplicativo. Não foram usados eventos `DragEvent` sintéticos de JavaScript para substituir o gesto nativo.
  - Captura `06-mover-posicao.png` inspecionada visualmente: página original 1 agora na posição 50, miniatura destacada e indicador de pendência. Artefatos finais em `%TEMP%/4printshops-browser-OlH2hg/`, apenas capturas e PDFs sintéticos fora do repositório.
  - Sintaxe JS/MJS e `git diff --check`: passaram. `node teste-generator.js` e `node teste2.mjs`: código 0. `npm.cmd test`: código 1 porque permanece o placeholder `Error: no test specified`; usar o comando explícito da suíte acima.
- **Limitações / pendências:**
  - Não foram disponibilizados os PDFs reais do incidente. Falhas/atrasos foram reproduzidos de forma controlada, combinados com renderização real de documentos sintéticos; repetir o roteiro nos arquivos de produção.
  - O prazo de 10 segundos pode indicar erro em uma página muito complexa; é possível tentar novamente ou navegar normalmente. A proteção é contra esperas assíncronas; não interrompe uma função síncrona que bloqueie a thread do navegador.
  - Touchpad físico, outros navegadores e cancelamento manual por Escape junto às bordas ainda precisam de conferência. O Chrome headless testou roda e arraste nativo com cancelamento no centro do painel.
  - Cada edição ainda serializa o documento e recalcula as dimensões de todas as páginas. As limitações de desempenho e preservação de recursos registradas na entrada anterior continuam válidas. Não foram implementados redimensionamento, análise de cores ou outras ferramentas.
- **Roteiro manual / próximo passo:**
  1. Iniciar `node app.js`, recarregar com Ctrl+F5 e abrir uma cópia de teste de 100 páginas. Selecionar a página 5, digitar 50 em **Mover para posição** e clicar Mover (repetir com Enter): conferir conteúdo, posição/numeração, miniatura destacada e pendência, sem download. Repetir com 350 páginas e destinos fora da janela.
  2. Testar destino vazio, zero, negativo, fracionário e maior que o total: deve mostrar orientação, sem alterar o documento. Testar ↑/↓ e primeira/última página.
  3. Rolar sobre a lista com roda/touchpad; a prévia central deve permanecer na página atual. Arrastar e manter o ponteiro junto às bordas inferior/superior: a lista deve avançar/recuar com linha de inserção. Soltar em posição distante; repetir e cancelar com Escape, inclusive junto à borda.
  4. Fazer vários movimentos enquanto as prévias ainda carregam, rolar entre regiões distantes e selecionar páginas dessas regiões. As miniaturas devem concluir ou indicar erro com **Tentar novamente**; a falha de uma não deve bloquear as outras nem manter os controles desabilitados.
  5. Girar, excluir e unir após reorganizar; exportar por Salvar cópia, reabrir o arquivo e conferir ordem, rotações, total e recursos importantes. Conferir que os originais permanecem intactos.
  6. Editar novamente, tentar sair/importar outro PDF e cancelar o aviso: conferir preservação. Exportar e alternar várias vezes entre ferramentas; não deve haver ações duplicadas. Registrar o arquivo e a sequência de ações caso algum sintoma real reapareça antes de avançar para novas ferramentas.

### 28/10/2026 — Otimização de atualização visual e cache de PDFs
- **Objetivo:** reduzir o trabalho repetido após giro/movimentação/exclusão, preservando as funções aprovadas e a fidelidade do workspace, sem alterações de interface.
- **Situação inicial:** Git limpo no commit `76eedd0`. Nenhum commit, push, comando destrutivo ou dependência instalada nesta tarefa.
- **Arquivos alterados:** `public/js/tools/pdf-workspace.js`, `pdfs.js`, `pdf-preview.js`, `pdf-thumbnails.js`; `tests/pdf-workspace.test.mjs`, `pdf-integration.test.mjs`, `pdf-thumbnails.test.mjs`, `pdf-browser-smoke.mjs`; este registro.
- **Novos arquivos:** `public/js/tools/pdf-visual-document.js`, `public/js/tools/pdf-document-info.js`, `tests/pdf-visual-document.test.mjs`, `tests/pdf-performance.mjs`, `docs/PERFORMANCE_PDFS_2026-10-28.json`. `pdf-thumbnail-job.js` e `pdf-thumbnail-drag.js` foram inspecionados e preservados; cancelamento, prazo, rolagem e arraste continuam sendo reutilizados.
- **Fase 1 — diagnóstico e linha de base:**
  - Giro: alterava a rotação no pdf-lib, chamava `obterBytes()`/`pdf.save()`, carregava outro PDF.js por `getDocument()`, lia as dimensões de todas as páginas, destruía a fonte anterior, reconstruía toda a lista técnica e descartava/refazia miniaturas.
  - Movimento: além desse ciclo, `reordenarPaginas()` salvava/carregava uma cópia integral para preservar objetos/recursos e validar a alteração antes de substituir o workspace. Portanto, eram dois saves e um novo carregamento PDF.js por movimento.
  - Exclusão: removia as folhas da árvore e executava o mesmo ciclo visual completo do giro. A medição confirmou que o custo crescia com o documento, mesmo para uma única página afetada.
  - A pdf-lib local 1.17.1 mantém `getPages()` em cache após `removePage()`. Confirmado no código executável da biblioteca e numa reprodução de duas páginas (depois de excluir: `getPages().length` continuava 2 e `getPageCount()` era 1). O novo mapa consulta a árvore real; giro localiza a folha pela referência correta, evitando índices obsoletos após exclusões.
- **Fase 2 — implementação:**
  - `PDFWorkspace.obterPaginas()` fornece identidades por referência e rotações absolutas a partir da árvore real. Continua sendo a fonte de verdade; não há uma segunda sequência editável mantida pela interface.
  - `PDFDocumentoVisual` cria snapshots imutáveis da correspondência posição → página original PDF.js. Giro/movimentação atualizam esse mapa e derivam dimensões do cache; não salvam/recarregam o PDF para atualizar o visual. Rotação é absoluta, aplicada uma única vez no viewport.
  - O canvas principal é reutilizado quando fonte, identidade, rotação e largura permanecem iguais. Mesmo nesse caso, invalida solicitações antigas que ainda aguardem `getPage()`.
  - Miniaturas são reaproveitadas por identidade + rotação, com cache LRU limitado a 32 canvases (aproximadamente 1,8 MB de pixels no máximo), incluindo os visíveis. Há até 16 itens no DOM; o giro invalida só a prévia afetada. Exclusões removem chaves inválidas; importação/união limpam o cache. Evicção e encerramento zeram os canvases.
  - O resumo conserva textos, alertas e disposição atuais. A lista de dimensões preserva seus elementos e só reescreve textos alterados. Ainda existe uma varredura leve de metadados em memória; não se afirma atualização O(1).
  - Extração e exportação usam os bytes atuais do workspace. A fonte PDF.js reutilizada nunca é usada como se já contivesse a nova ordem ou rotação nos seus bytes.
  - Exclusão mantém o save + carregamento PDF.js completos, reaproveitando dimensões/canvases por identidade. A cópia transacional do workspace por movimento foi mantida para não arriscar recursos do PDF. Importação, união e incompatibilidade do mapa usam o caminho completo.
- **Fase 3 — segurança:**
  - Operações continuam serializadas pelo bloqueio existente. Mapas anteriores não são alterados; callbacks conferem documento/geração, e miniaturas pendentes são canceladas antes de mudar a janela/fonte.
  - Uma falha incremental tenta o carregamento completo. Uma nova fonte só substitui/destrói a anterior depois de validar seu primeiro render, preservando o documento anterior nas falhas de importação/união. Falha persistente ao girar restaura a rotação/revisão anterior.
  - Se o recarregamento falhar após uma exclusão já aplicada, o mapa/resumo/miniaturas são sincronizados ao workspace e o canvas principal antigo é removido. O erro é informado e os controles voltam; navegar permite tentar novamente. Não é exibida uma página antiga como se fosse o documento novo.
- **Medições antes/depois:**
  - Script: `node tests/pdf-performance.mjs antes` no código base, depois `node tests/pdf-performance.mjs depois` no código otimizado. Chrome headless `145.0.7632.117`, Windows, viewport 1440×1000, mesma máquina, sem benchmarks concorrentes. Documentos sintéticos com texto/vetores, 20/100/350 páginas. Três amostras sucessivas por operação/tamanho; valores abaixo são medianas em ms. Importação e seleção da página 5 ficam fora da medição. Movimento: 5↔20 no PDF de 20 páginas; 5↔50 nos demais. Exclusões removem a última página, sucessivamente.

  | Páginas | Giro antes → depois | Movimento antes → depois | Exclusão antes → depois |
  | --- | --- | --- | --- |
  | 20 | 80,7 → 14,3 | 96,9 → 11,6 | 83,1 → 83,0 |
  | 100 | 165,3 → 16,6 | 199,7 → 72,0 | 132,6 → 116,0 |
  | 350 | 265,8 → 16,6 | 465,9 → 222,0 | 278,9 → 200,0 |

  - Tempo acima: clique até o término do render principal/desbloqueio dos controles; não inclui esperar todas as miniaturas. Em 350 páginas, redução medida de 93,8% no giro e 52,4% no movimento. Exclusão de 20 páginas não apresentou ganho relevante.
  - Medianas por fase em 350 páginas, também em ms (antes → depois):

  | Operação | Saves | Cópia pdf-lib | Carregamento PDF.js | Dimensões/prepare | Mostrar página | Demais trabalho/interface |
  | --- | --- | --- | --- | --- | --- | --- |
  | Giro | 141,3 → 0 | 0 → 0 | 64,8 → 0 | 24,8 → 0 | 18,0 → 13,8 | 16,6 → 2,6 |
  | Movimento | 277,7 → 146,3 | 55,8 → 60,5 | 64,3 → 0 | 23,9 → 0 | 19,7 → 0,1 | 22,8 → 9,4 |
  | Exclusão | 153,2 → 111,6 | 0 → 0 | 63,4 → 63,9 | 23,2 → 0,4 | 20,8 → 20,0 | 16,4 → 4,9 |

  - Contagens verificadas: giro passou de 1 save/1 getDocument para 0/0; movimento, de 2/1 para 1/0; exclusão permanece 1/1. Mostrar página no movimento retorna pelo cache, sem rasterização. A cópia transacional segue sendo o principal custo residual.
  - As fases são instrumentadas por wrappers nos módulos reais. “Dimensões/prepare” mede o intervalo de preparação técnica; “Demais trabalho/interface” é o resíduo por amostra após descontar fases reconhecidas, não uma medição exclusiva do DOM. Medianas independentes não precisam somar exatamente ao total. Fila de miniaturas é medida separadamente e pode se sobrepor ao render principal.
  - Na primeira movimentação 5→50, ainda sem cache no destino: 100 páginas, interação 216,4 → 71,8 ms e estabilização das miniaturas 316,2 → 168,9 ms; 350 páginas, interação 482,0 → 216,8 ms e estabilização 582,7 → 317,2 ms. Nas repetições, o LRU evita refazer as prévias; a estabilização mediana em 350 páginas caiu de 568,9 para 229,8 ms.
  - Dados completos, fases, contagens e 54 amostras: [PERFORMANCE_PDFS_2026-10-28.json](PERFORMANCE_PDFS_2026-10-28.json). Artefatos originais: `%TEMP%/4printshops-performance-RqvXgL/` e `%TEMP%/4printshops-performance-YTHVT1/`.
  - **Limites da comparação:** três amostras e PDFs sintéticos simples; não são garantias para arquivos de produção. O custo do save na exclusão variou entre as execuções, mas seu algoritmo não foi otimizado; não atribuir essa variação ao cache. O ganho consistente vem de evitar carregamentos/leituras/renders repetidos. A mediana de movimentos sucessivos inclui o benefício de cache aquecido.
- **Fase 4 — testes executados:**
  - Antes: os 42 testes existentes passaram. Depois: `node --test tests/pdf-workspace.test.mjs tests/pdf-integration.test.mjs tests/pdf-thumbnails.test.mjs tests/pdf-protection.test.mjs tests/pdf-visual-document.test.mjs`: **54 testes passaram**.
  - Novos casos: movimentos/giros sucessivos em 100/350 páginas com contagem de saves/carregamentos; extração antes de recarregar PDF.js; giro/exclusão/união/exportação subsequentes; identidade/dimensões; LRU/evicção/liberação; mapa antigo com resposta atrasada; getPage pendente ao reutilizar canvas; falha incremental com fallback; falha persistente de giro/união; falha de recarga na exclusão; cache de índices da pdf-lib após exclusões sucessivas.
  - `node tests/pdf-browser-smoke.mjs`: **passou no Chrome headless**. Roteiro anterior preservado: arraste nativo, ↑/↓, Mover para posição 5→50 em 100/350 páginas, roda/autoscroll, atrasos/erro/retry/cancelamento, giro/exclusão/união/exportação, proteção e reentradas. Acrescentadas preservação do canvas no movimento e comparação dos pixels da prévia com a página extraída girada e com o PDF final exportado, para ambos os tamanhos. Tolerância: menos de 0,1% dos componentes RGBA com diferença maior que 2 níveis. Nenhum erro JavaScript capturado.
  - Artefatos finais do Chrome: `%TEMP%/4printshops-browser-p8N635/`, apenas capturas e PDFs sintéticos. Sintaxe dos 12 JS/MJS alterados/novos e `git diff --check` passaram. Scripts avulsos `teste-generator.js`/`teste2.mjs`: código 0. `npm.cmd test` continua com o placeholder `Error: no test specified`, código 1; usar a suíte explícita acima.
- **Pendências e próximo passo:**
  - Repetir com PDFs reais da gráfica, inclusive conteúdo complexo, links, anotações e formulários. As limitações anteriores de preservação de recursos/assinaturas permanecem. Nenhum redimensionamento, análise de cores, ferramenta nova ou mudança visual nesta tarefa.
  - Roteiro: abrir 20/100/350 páginas; girar várias vezes, mover 5→50 e voltar, usar setas/arraste/autoscroll, excluir antes da página selecionada e girar novamente, unir, extrair e exportar. Reabrir a exportação e comparar ordem, conteúdo e orientação com a prévia. Conferir pendências/aviso ao sair, cancelamento de importação e reentradas.
  - Para futuras operações que alterem conteúdo/recursos, usar recarga completa e limpar o cache. Só ampliar o caminho incremental com invalidação e testes específicos. Não retirar a cópia transacional de reorganização sem outra estratégia comprovada de preservar o documento inteiro.

### 29/10/2026 — Redimensionar páginas e biblioteca de formatos
- **Objetivo:** acrescentar redimensionamento vetorial, prévia temporária e formatos reutilizáveis, sem refazer as operações de edição existentes.
- **Situação inicial:** Git limpo no commit `76f1ee4`. Todos os 54 testes existentes passaram antes da implementação. Nenhum commit, push, comando destrutivo ou dependência instalada nesta etapa.
- **Arquivos novos:** `public/js/tools/pdf-format-storage.js`, `pdf-formats.js`, `pdf-resize-interface.js`, `pdf-resize-ui.js`, `pdf-resize-session.js`, `pdf-resize-engine.js`; `tests/helpers/pdf-resize-modules.mjs`, `tests/pdf-formats.test.mjs`, `tests/pdf-resize.test.mjs`; `docs/REDIMENSIONAR_PDFS.md`.
- **Arquivos modificados:** `public/js/tools/pdf-interface.js`, `pdf-workspace.js`, `pdfs.js`, `pdf-preview.js`, `pdf-thumbnails.js`, `pdf-visual-document.js`; `public/css/style.css`; `tests/pdf-workspace.test.mjs`, `tests/pdf-integration.test.mjs`, `tests/pdf-browser-smoke.mjs`; este registro.
- **Interface e formatos (fases 1/2):**
  - Nova ferramenta no painel contextual direito, preservando o layout geral. A3 297×420, A4 210×297, A5 148×210, A6 105×148 mm e Personalizado; orientação retrato/paisagem; medidas de 0,1 a 5000 mm com vírgula ou ponto decimal, sem separador de milhar.
  - Escopo: atual, intervalo pela mesma interpretação existente de páginas, ou todas. Padrão: Ajustar proporcionalmente. Alternativas: Centralizar sem alterar a escala e Preencher com corte. Avisos de corte na configuração e na prévia real.
  - Biblioteca: nome/dimensões, seleção, cadastro, atualização/renomeação e exclusão com confirmação nativa. Padrões protegidos; valida nomes duplicados, medidas, campos incompletos, limite de 80 caracteres/200 formatos. Não presume dimensões de Formato 4/8 nem altera o PDF ao cadastrar/selecionar.
  - localStorage isolado no adaptador `ler`/`gravar`, dados versionados e recuperação de dados inválidos. Bloqueio/quota: aviso e retenção durante a sessão da página, incluindo reentradas no módulo; recarregar/fechar pode perder dados não persistidos. O serviço também aceita adaptadores assíncronos; interface aguarda inicialização/CRUD. Nenhum servidor, login, cobrança ou sincronização implementado.
- **Prévia, motor e integração (fases 3/4/5):**
  - Workspace prepara uma cópia completa do PDF, sem alterar o definitivo. O motor transforma streams vetoriais/textos/imagens por matriz afim e clipping, com escala uniforme e deslocamento central. Todas as páginas selecionadas são validadas/preparadas transacionalmente; versões antigas/canceladas não podem ser aplicadas.
  - Usa MediaBox∩CropBox como origem visível, respeita origem deslocada/negativa, UserUnit e rotações 0/90/180/270. Mantém o conteúdo antes oculto recortado. Nas páginas alteradas, incorpora a orientação ao conteúdo e normaliza Rotate=0/UserUnit=1, evitando dupla rotação.
  - MediaBox/CropBox recebem o destino; TrimBox/BleedBox/ArtBox existentes acompanham a transformação e são limitadas à área visível. Não cria caixas ausentes ou sangria. Caixa que ficaria vazia provoca erro antes de publicar o documento.
  - Links URI/GoToR e Text/Popup simples sem AP: transforma Rect, QuadPoints e bordas aplicáveis, preservando os demais dados/referências. Streams/recursos e operadores de cor anteriores são mantidos; não houve rasterização nem conversão RGB/CMYK.
  - Pré-visualizar carrega os bytes realmente transformados no PDF.js e mostra uma página-alvo, com dimensões originais/novas. Não gera pendência nem atualiza o resumo/miniaturas definitivos. Se a página atual não pertence ao intervalo, identifica a primeira página do intervalo como prévia.
  - Cancelar, mudar configurações ou selecionar outra ferramenta descarta a prévia e conserva as edições anteriores. Durante a prévia ficam bloqueadas operações que misturariam os dois estados. Aplicar confirma a revisão, conserva a seleção e atualiza resumo, alertas, navegação e miniaturas.
  - Aplicar reutiliza o PDF.js já preparado; não serializa/carrega novamente o candidato. Reaproveita dimensões/miniaturas não afetadas e invalida identidades transformadas. Giro/movimento otimizados continuam sem recarregar a fonte; cache LRU/janela virtual/autoscroll foram preservados.
  - Cancelamento usa AbortSignal/geração e revisão. O carregamento PDF.js pode ser interrompido, sem esperar sua promise concluir; `destroy()` continua memoizado por tarefa. Falha na aplicação restaura o workspace. Se a restauração visual também falhar, o canvas temporário é removido, o erro aparece e os controles são liberados, evitando mostrar uma prévia descartada como versão definitiva.
  - Dimensões já iguais não produzem nova transformação/revisão. O redimensionamento não dispara download; exportação pelo botão Salvar cópia. Extração continua com o comportamento anterior.
- **Limitações de compatibilidade (bloqueios deliberados):**
  - Páginas-alvo com Widget/campos, tipos de anotação além de Link/Text/Popup, aparências AP, ações AA ou ações de link não suportadas. Não achata nem remove esses recursos silenciosamente.
  - Documentos com destinos/links internos, marcadores Outlines, estrutura marcada StructTreeRoot, OpenAction, XFA ou SigFlags. O bloqueio é conservador para não corromper coordenadas, interatividade ou assinaturas; as outras ferramentas continuam disponíveis.
  - Ícones de notas dependem do leitor; fontes, links/anotações reais e variantes não cobertas precisam de validação. Não há garantia para qualquer PDF ou recurso avançado. Não afirma CMYK, DPI ou sangria suficiente.
  - O preparo ainda exige cópia/serialização completa. Algumas fases de parsing/save não são interrompidas imediatamente. A janela virtual e o cache são limitados, mas o documento/candidato consomem memória integralmente.
- **Testes e desempenho (fase 6):**
  - `node --test tests/pdf-workspace.test.mjs tests/pdf-integration.test.mjs tests/pdf-thumbnails.test.mjs tests/pdf-protection.test.mjs tests/pdf-visual-document.test.mjs tests/pdf-formats.test.mjs tests/pdf-resize.test.mjs`: **78 testes passaram**.
  - Cobertos A4→A5, A5→A4, A3→A4, personalizados decimais, paisagem, três modos, página/intervalo/todas, páginas mistas/giradas, caixas, origem/UserUnit, links/notas, recursos/streams, transação, cancelamento/versão vencida, não duplicação, operações antes/depois, CRUD/persistência, corrupção, quota e adaptador assíncrono.
  - Integração: prévia sem pendência, Cancelar mantendo giros anteriores, Aplicar sem novo getDocument, alteração de configuração durante render, troca de ferramenta, caso sem efeito, recursos bloqueados, cancelamento de getDocument sem conclusão e falhas de aplicação/restauração sem canvas obsoleto. Regressões anteriores preservadas.
  - `node tests/pdf-browser-smoke.mjs`: **passou no Chrome headless**. Testou três modos × quatro rotações com texto/vetores/PNG e CropBox deslocado; centroides de marcas coloridas confirmaram origem, orientação e centralização no ajuste. Pixels de prévia/aplicação coincidiram; exportações foram comparadas com a visualização pelo teste existente. Foram testadas persistência após reload, seleção/renomeação e exclusão de formato com cancelamento/confirmação nativa.
  - Em 100/350 páginas, prévia/aplicação ao documento inteiro, cancelamento durante preparo, eventos continuando a executar, miniaturas limitadas e operações posteriores de giro, mover para 50, exclusão, união e exportação passaram. Uma amostra final de preparo completo: **271 ms / 789 ms**; tentativa cancelada de 350 páginas: **182 ms**. PDFs sintéticos simples; não são garantias para arquivos reais. Artefatos em `%TEMP%/4printshops-browser-g2ocEb/`; captura `07-redimensionar.png` também inspecionada durante a validação.
  - `node tests/pdf-performance.mjs redimensionamento-regressao`: passou em 20/100/350 páginas. Em 350 páginas, medianas giro **17,1 ms**, movimento **234,3 ms**, exclusão **219,5 ms**; a medição anterior era 16,6 / 222,0 / 200,0 ms, sem evidência de grande regressão nessa amostra. Giro continua 0 save/0 getDocument; movimento 1/0; exclusão 1/1. Artefatos `%TEMP%/4printshops-performance-GHTJwR/`; há variação normal de ambiente/serialização e amostra pequena.
  - Sintaxe dos JS/MJS novos/alterados e `git diff --check`: passaram. Scripts avulsos `teste-generator.js`/`teste2.mjs`: código 0. `npm.cmd test` segue como placeholder `Error: no test specified`, código 1; usar a suíte explícita acima.
- **Documentação e próximo passo (fase 7):**
  - [REDIMENSIONAR_PDFS.md](REDIMENSIONAR_PDFS.md) detalha uso, contrato de persistência/futuro servidor, caixas, compatibilidade e roteiro com PDFs reais.
  - Validar cópias de arquivos da gráfica: formatos, três modos, intervalos/mistos/girados, Cancelar/Aplicar, exportar/reabrir, recursos suportados e mensagens dos bloqueados. Testar formatos após reiniciar o navegador e reentradas sem listeners duplicados.
  - Não implementados imposição, sangria artificial, marcas de corte, conversão de cor, contas/cobrança ou sincronização entre computadores. Não ampliar suporte a recursos bloqueados sem transformação e testes próprios.

### 29/10/2026 — Aprovação manual e preparação do commit de redimensionamento
- **Validação informada pelo usuário:** testes com PDFs reais da gráfica, incluindo PDF/X-1a, PDF/X-3, PDF/X-4, Acrobat 9, textos, curvas e páginas de dimensões diferentes. Todos os testes manuais realizados foram aprovados.
- **Escopo da aprovação:** arquivos/cenários efetivamente testados e relatados pelo usuário. Os bloqueios de recursos complexos e demais limites descritos em `REDIMENSIONAR_PDFS.md` continuam em vigor.
- **Commit local autorizado:** `Adiciona redimensionamento de PDFs e formatos personalizados`.
- **Revisão:** alterações de produção, testes e documentação relacionadas ao redimensionamento. Capturas, PDFs sintéticos e logs de teste ficaram no diretório temporário do sistema, fora do commit. Nenhuma alteração de credenciais, configuração remota ou push autorizado/executado.
- **Verificações antes do commit:** 78 testes automatizados passaram novamente; sintaxe dos 18 JS/MJS alterados/novos e `git diff --check` passaram. Não foram encontrados temporários ou arquivos não relacionados entre as alterações pendentes.
- **Próximo passo:** manter esta versão aprovada; novas funcionalidades dependem de nova solicitação.

### Próxima entrada
- **Data:**
- **Objetivo:**
- **Arquivos alterados:**
- **O que foi implementado:**
- **Testes executados e resultados:**
- **Problemas conhecidos / dúvidas:**
- **Próximo passo recomendado:**

## 9. Mensagem para sincronizar com o ChatGPT
Ao encerrar uma tarefa no VS Code, apresentar ao usuário um resumo curto com arquivos alterados, testes e próximo passo, e atualizar a seção 8 deste arquivo. O usuário poderá copiar esse resumo para sua conversa de planejamento no ChatGPT. Não há comunicação direta em tempo real entre os dois ambientes.
