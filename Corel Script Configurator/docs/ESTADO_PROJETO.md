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
- `public/js/tools/pdf-workspace.js`: classe `PDFWorkspace`, usando pdf-lib; métodos conhecidos `abrir`, `getTotalPaginas`, `girarPagina`, `excluirPaginas`, `obterBytes`.
- `public/js/tools/pdf-editor.js`: exportação de PDFs girados/extraídos e lógica de operações legadas.
- `public/js/tools/pdf-merge.js`: ferramenta modal para unir PDFs; interface com lista, subir/descer/remover e gerar arquivo.
- `public/vendor/pdfjs/` e `public/vendor/pdf-lib/`: cópias locais de bibliotecas de PDF.
- `modelos/`: scripts de origem CorelDRAW. Não sobrescrever arquivos de produção sem autorização.

## 3. Funcionalidades de PDF confirmadas pelo usuário
- Importar e visualizar PDF e trocar de arquivo sem precisar recarregar a página.
- Navegar pelas páginas.
- Ver número de páginas e dimensão em milímetros; identificar diferentes dimensões e avisar pelo botão Informações.
- Girar páginas no `PDFWorkspace` e exportar PDF com alterações.
- Excluir páginas no `PDFWorkspace` e atualizar a pré-visualização **sem download automático**.
- Acumular rotação e exclusão até clicar em **Salvar cópia**.
- Extrair páginas por números/intervalos; **atualmente gera download imediato** (funcionamento legado).
- Unir vários arquivos, ordenando a lista e exportando PDF; **atualmente gera download imediato**, ainda não integrado ao workspace.
- Ferramenta "Unir PDFs" foi movida de painel expansível para botão compacto que abre uma janela/modal. O usuário confirmou que funciona.
- O usuário confirmou que o `pdfs.js` revisado está funcionando após correção de um trecho duplicado de rotação fora da função.

## 4. Bugs importantes já encontrados
- Erro de troca de arquivos por gerenciamento inadequado do objeto `.destroy()`: foi corrigido usando a tarefa de carregamento (`tarefaPDF`) e seu `.destroy()`.
- Erros de sintaxe por `try` sem `catch` e chaves sobrando: ocorreram em alterações incrementais de `pdfs.js`.
- Trecho antigo duplicado `workspace.girarPagina(paginaAtual, graus)` fora da função `girarPagina`: causava `graus` indefinido e impedia que os botões funcionassem. Foi removido em arquivo corrigido; usuário validou o funcionamento.
- Evitar retornar a dupla aplicação de rotações (`rotacoes` temporárias + workspace). A meta é que workspace seja fonte de verdade.

## 5. Funcionalidades ainda não feitas
- Integrar **Unir PDFs** ao workspace, adicionando páginas ao documento de trabalho, com pré-visualização atualizada e sem download imediato.
- Organizar páginas com **miniaturas** e seleção/arraste.
- Interface com ferramentas em **painel lateral contextual**; barra superior minimalista.
- Redimensionamento para A4, A5 e dimensões personalizadas, com encaixe proporcional e escolha explícita para cortar/preencher.
- Histórico de desfazer/refazer.
- Inspeção confiável de espaços de cor RGB/CMYK/cores especiais, perfis ICC, sangrias, fontes e resolução efetiva das imagens. Atualmente mostrar "Não verificado".
- Melhorar extração para abrir outro workspace/aba em vez de baixar automaticamente.
- QR Codes: XLSX, seleção de colunas/campos, validação, geração em lote e exportação.

## 6. Prioridade e próximo passo acordado
**Integrar a união de PDFs ao `PDFWorkspace`.** Comportamento desejado:
1. Importar PDF existente ou iniciar um documento com vários PDFs.
2. Clicar no botão compacto "Unir PDFs" para abrir a janela.
3. Escolher/ordenar arquivos e clicar em **Adicionar ao documento** (sem exportação automática).
4. Mostrar o novo total e permitir navegar, girar e excluir páginas da versão combinada.
5. Só baixar quando o usuário clicar em **Exportar/Salvar cópia**.
6. Em seguida, criar miniaturas e painel lateral. Posteriormente redimensionar para A5.

**Cuidados antes de implementar:** verificar `pdfs.js`, `pdf-merge.js` e `pdf-workspace.js` reais, pois alguns caminhos/funções podem ter evoluído; manter a função de união atual funcionando durante a migração; preservar operações do PDF.js e não salvar no disco sem ação explícita.

## 7. Fluxo de trabalho que o usuário prefere
- Mudanças pequenas, um passo por vez, explicações claras de qual arquivo foi alterado.
- Testar no navegador com PDFs reais, mantendo os arquivos originais intactos.
- Para testes de sintaxe, usar `node --check` em JS modificado.
- Priorizar funcionalidades de uso diário da gráfica e minimizar riscos de regressão.
- Arquivos JavaScript baixados pelo navegador já foram bloqueados; `.txt` funcionou como alternativa para transferência, com renomeação local.

## 8. Registro de trabalho futuro (Codex deve preencher)

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
