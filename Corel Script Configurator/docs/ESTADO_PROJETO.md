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
- `public/js/tools/pdf-workspace.js`: classe `PDFWorkspace`, usando pdf-lib; métodos `abrir`, `adicionarArquivos`, `getTotalPaginas`, `girarPagina`, `excluirPaginas`, `obterBytes`.
- `public/js/tools/pdf-editor.js`: exportação de PDFs girados/extraídos e lógica de operações legadas.
- `public/js/tools/pdf-merge.js`: ferramenta modal com lista, subir/descer/remover e ação **Adicionar ao documento**, integrada ao workspace por callback.
- `public/vendor/pdfjs/` e `public/vendor/pdf-lib/`: cópias locais de bibliotecas de PDF.
- `modelos/`: scripts de origem CorelDRAW. Não sobrescrever arquivos de produção sem autorização.

## 3. Funcionalidades de PDF e situação de validação
Os fluxos anteriores foram confirmados pelo usuário conforme o histórico abaixo. A integração da união, implementada nesta etapa, passou em testes locais e ainda depende de validação no navegador.
- Importar e visualizar PDF e trocar de arquivo sem precisar recarregar a página.
- Navegar pelas páginas.
- Ver número de páginas e dimensão em milímetros; identificar diferentes dimensões e avisar pelo botão Informações.
- Girar páginas no `PDFWorkspace` e exportar PDF com alterações.
- Excluir páginas no `PDFWorkspace` e atualizar a pré-visualização **sem download automático**.
- Acumular rotação e exclusão até clicar em **Salvar cópia**.
- Extrair páginas por números/intervalos; **atualmente gera download imediato** (funcionamento legado).
- Unir/adicionar um ou mais arquivos, ordenando a lista: agora acrescenta páginas ao workspace aberto ou inicia um novo documento, **sem download automático**. Mantém rotações/exclusões existentes e permite continuar editando; exportação por **Salvar cópia**. Implementação validada em testes locais; navegador pendente.
- Ferramenta "Unir PDFs" foi movida de painel expansível para botão compacto que abre uma janela/modal. O usuário confirmou que funciona.
- O usuário confirmou que o `pdfs.js` revisado está funcionando após correção de um trecho duplicado de rotação fora da função.

## 4. Bugs importantes já encontrados
- Erro de troca de arquivos por gerenciamento inadequado do objeto `.destroy()`: foi corrigido usando a tarefa de carregamento (`tarefaPDF`) e seu `.destroy()`.
- Erros de sintaxe por `try` sem `catch` e chaves sobrando: ocorreram em alterações incrementais de `pdfs.js`.
- Trecho antigo duplicado `workspace.girarPagina(paginaAtual, graus)` fora da função `girarPagina`: causava `graus` indefinido e impedia que os botões funcionassem. Foi removido em arquivo corrigido; usuário validou o funcionamento.
- Evitar retornar a dupla aplicação de rotações (`rotacoes` temporárias + workspace). A meta é que workspace seja fonte de verdade.

## 5. Funcionalidades ainda não feitas
- Validar manualmente no navegador a integração de **Unir PDFs** ao workspace, incluindo PDFs reais e arquivos maiores.
- Organizar páginas com **miniaturas** e seleção/arraste.
- Interface com ferramentas em **painel lateral contextual**; barra superior minimalista.
- Redimensionamento para A4, A5 e dimensões personalizadas, com encaixe proporcional e escolha explícita para cortar/preencher.
- Histórico de desfazer/refazer.
- Inspeção confiável de espaços de cor RGB/CMYK/cores especiais, perfis ICC, sangrias, fontes e resolução efetiva das imagens. Atualmente mostrar "Não verificado".
- Melhorar extração para abrir outro workspace/aba em vez de baixar automaticamente.
- QR Codes: XLSX, seleção de colunas/campos, validação, geração em lote e exportação.

## 6. Prioridade e próximo passo acordado
**Validar no navegador a integração da união ao `PDFWorkspace`, implementada em 28/10/2026.** Fluxo disponível no código:
1. Importar PDF existente ou iniciar um documento com vários PDFs.
2. Clicar no botão compacto "Unir PDFs" para abrir a janela.
3. Escolher/ordenar arquivos e clicar em **Adicionar ao documento** (sem exportação automática).
4. Mostrar o novo total e permitir navegar, girar e excluir páginas da versão combinada.
5. Só baixar quando o usuário clicar em **Exportar/Salvar cópia**.
6. Após validação do usuário, planejar miniaturas e painel lateral. Posteriormente redimensionar para A5; nenhuma dessas etapas foi iniciada nesta alteração.

**Cuidados nas próximas etapas:** preservar a rotina compartilhada de atualização do PDF.js, o bloqueio de operações simultâneas e os listeners ligados aos elementos da instância. A extração ainda baixa diretamente; a união só modifica o documento em memória. O nome informado no modal é usado ao iniciar um documento; ao acrescentar páginas, o documento aberto mantém seu nome.

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
