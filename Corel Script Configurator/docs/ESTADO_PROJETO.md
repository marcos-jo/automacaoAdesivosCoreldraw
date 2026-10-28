# 4PrintShops — Estado do projeto e passagem de contexto

**Atualizado em:** 28/10/2026  
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
Os fluxos anteriores foram confirmados pelo usuário conforme o histórico abaixo. A união e a nova interface passaram em testes locais e em Chrome headless com PDFs sintéticos. A validação manual com arquivos reais da gráfica permanece pendente.
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
- Proteção contra perda de alterações não exportadas: confirmação ao sair para outra ferramenta ou substituir o PDF, indicador próximo à exportação e `beforeunload` apenas enquanto houver pendências. Abrir para visualização não gera aviso; exportar a versão atual limpa a pendência, e editar novamente a reativa. Validado em Node e Chrome headless; conferir os limites do download/navegador na entrada mais recente abaixo.
- Ferramenta "Unir PDFs" foi movida de painel expansível para botão compacto que abre uma janela/modal. O usuário confirmou que funciona.
- O usuário confirmou que o `pdfs.js` revisado está funcionando após correção de um trecho duplicado de rotação fora da função.

## 4. Bugs importantes já encontrados
- Erro de troca de arquivos por gerenciamento inadequado do objeto `.destroy()`: foi corrigido usando a tarefa de carregamento (`tarefaPDF`) e seu `.destroy()`.
- Erros de sintaxe por `try` sem `catch` e chaves sobrando: ocorreram em alterações incrementais de `pdfs.js`.
- Trecho antigo duplicado `workspace.girarPagina(paginaAtual, graus)` fora da função `girarPagina`: causava `graus` indefinido e impedia que os botões funcionassem. Foi removido em arquivo corrigido; usuário validou o funcionamento.
- Evitar retornar a dupla aplicação de rotações (`rotacoes` temporárias + workspace). A meta é que workspace seja fonte de verdade.

## 5. Funcionalidades ainda não feitas
- Validar manualmente a nova interface e as operações de PDFs com arquivos reais, inclusive documentos maiores e complexos.
- Seleção múltipla visual (reorganização individual por arraste ou botões já implementada).
- Redimensionamento para A4, A5 e dimensões personalizadas, com encaixe proporcional e escolha explícita para cortar/preencher.
- Histórico de desfazer/refazer.
- Inspeção confiável de espaços de cor RGB/CMYK/cores especiais, perfis ICC, sangrias, fontes e resolução efetiva das imagens. Atualmente mostrar "Não verificado".
- Melhorar extração para abrir outro workspace/aba em vez de baixar automaticamente.
- QR Codes: XLSX, seleção de colunas/campos, validação, geração em lote e exportação.

## 6. Prioridade e próximo passo acordado
**Repetir os testes reais de rolagem, reorganizações consecutivas e recuperação de miniaturas após esta correção, conferindo a preservação dos recursos dos PDFs.** Os testes Node e Chrome headless com documentos sintéticos passaram. Fluxo disponível no código:
1. Importar PDF existente ou iniciar um documento com vários PDFs.
2. Selecionar "Unir PDFs" no painel direito e usar seu botão para abrir a janela.
3. Escolher/ordenar arquivos e clicar em **Adicionar ao documento** (sem exportação automática).
4. Conferir o total e as miniaturas; navegar por clique ou Anterior/Próxima. Reorganizar arrastando uma miniatura, usando ↑/↓ ou informando **Mover para posição** no painel esquerdo. Rolar a lista independentemente da prévia; durante o arraste, aproximar o ponteiro das bordas para percorrer páginas distantes. Escolher Rotacionar ou Organizar páginas no painel direito para continuar editando.
5. Só baixar quando o usuário clicar em **Exportar/Salvar cópia**.
6. Recolher/reabrir o painel direito quando necessário. Redimensionamento, imposição, seleção múltipla e análise de cores permanecem fora desta etapa.

**Cuidados nas próximas etapas:** preservar a rotina compartilhada de atualização do PDF.js, o bloqueio de operações simultâneas e os listeners ligados aos elementos da instância. `renderPDFs()` retorna uma função de encerramento com `confirmarSaida()`; `main.js` consulta a confirmação antes de limpar o módulo ou atualizar o menu ativo. Miniaturas devem parar antes da destruição dos documentos PDF.js. `PDFWorkspace.alterado` é derivado das revisões atual/exportada; novas operações devem chamar `registrarAlteracao()`, sem atribuir diretamente a `alterado`. A extração ainda baixa diretamente e não quita a exportação do documento completo. O nome informado na união é usado ao iniciar um documento; ao acrescentar páginas, o documento aberto mantém seu nome.

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
