# AGENTS.md — 4PrintShops

## Objetivo
Este repositório contém o 4PrintShops, uma central de ferramentas locais para o fluxo de trabalho de uma gráfica. O usuário trabalha com o projeto no Windows e prefere desenvolver em etapas pequenas, validando cada mudança no navegador. Converse com o usuário em português do Brasil.

## Contexto obrigatório

Antes de iniciar qualquer implementação, correção ou investigação relevante, o agente deve consultar, nesta ordem:

1. `AGENTS.md` — regras permanentes do projeto.
2. `docs/PLANEJAMENTO.md` — tarefas, prioridades e decisões oficiais.
3. `docs/CONTEXTO_ATUAL.md` — checkpoint recente e passagem de trabalho entre computadores.
4. `docs/ESTADO_PROJETO.md` — histórico técnico detalhado, testes e limitações.

Antes de alterar código, conferir `git status -sb`, branch, HEAD, divergência com o remoto e alterações locais não publicadas.

O planejamento é a referência para prioridades aprovadas. O contexto atual identifica o último trabalho realizado. O histórico técnico fornece evidências, mas não substitui a inspeção do código.

Se houver divergência entre documentos, código local e GitHub, investigar e informar ao usuário antes de prosseguir. Não presumir que trabalho realizado em outro computador foi sincronizado.

Nenhum agente pode alterar prioridades aprovadas ou iniciar outra tarefa por conta própria. Novas sugestões devem ser registradas como propostas e submetidas ao usuário.

## Organização conhecida
- `app.js` (raiz): inicializa o servidor Node.js, executado com `node app.js`.
- `src/`: servidor HTTP, repositório de scripts, parser, modelo e gerador de scripts CorelDRAW.
- `public/`: interface web; `public/js/main.js` é o controlador do navegador.
- `public/js/tools/pdfs.js`: interface, leitura, visualização e operações de PDF.
- `public/js/tools/pdf-workspace.js`: documento de trabalho em memória via pdf-lib.
- `public/js/tools/pdf-editor.js`: operações/exportações PDF legadas.
- `public/js/tools/pdf-merge.js`: janela/modal de união de PDFs.
- `public/js/tools/qrcodes.js`: módulo inicial de QR Codes.
- `public/vendor/`: bibliotecas locais PDF.js e pdf-lib (confira nomes e versões instalados).
- `modelos/`: scripts originais do CorelDRAW; não alterar sem solicitação explícita.

## Princípios para alterações
1. **Preservar o que já funciona.** Investigue antes de refatorar. Evite reescrever arquivos grandes indiscriminadamente.
2. Antes de cada tarefa, explique brevemente quais arquivos precisam mudar e por quê. Implemente uma etapa pequena por vez.
3. Alterações no Gerenciador de PDFs devem atualizar o documento de trabalho e a pré-visualização; **não iniciar downloads automaticamente**. O usuário exportará o resultado final quando desejar. Algumas funções legadas ainda exportam diretamente; migre-as gradualmente.
4. PDFs originais e arquivos dos clientes não podem ser alterados diretamente. Processar localmente, sem upload a serviços externos sem permissão explícita.
5. Ao usar PDF.js, gerenciar corretamente tarefas de carregamento e renderização. Não voltar a introduzir destruição concorrente nem referências incorretas a `destroy()`.
6. Evitar estados duplicados: `PDFWorkspace` deve tornar-se a fonte principal das alterações. Atenção especial à rotação aplicada duas vezes.
7. Não afirmar que um PDF é CMYK, possui sangria correta ou imagens em DPI suficiente sem uma inspeção confiável. Mostrar "não verificado" quando ainda não há detector.
8. O projeto utiliza módulos ES no navegador e Node.js; preservar caminhos de importação e o funcionamento do servidor HTTP atual.
9. Não instalar dependências, remover funcionalidades ou modificar vários módulos não relacionados sem necessidade e sem explicação ao usuário.
10. Ao propor mudanças, fornecer explicações curtas e testes práticos de aceitação. O usuário usa VS Code/Codex e anteriormente Sublime Text.

## Verificações
- Confirmar a sintaxe dos arquivos JS alterados com `node --check caminho/do/arquivo.js`.
- Inspecionar `package.json` antes de presumir que existem scripts como `npm test` ou `npm run lint`.
- Se possível, validar o fluxo de abrir PDF, trocar arquivos, girar, excluir, navegar e exportar, sem alterar originais.
- Relatar o que foi efetivamente testado e o que depende de teste manual do usuário. Não declarar testes que não foram executados.
- Antes de grandes refatorações, pedir uma cópia de segurança ou orientar commit Git.

## Política de Git e GitHub

- O Codex pode modificar arquivos e executar testes normalmente.
- O Codex só deve criar commits locais após aprovação explícita do usuário.
- Antes do commit, verificar git status e git diff, incluindo apenas os arquivos relacionados à tarefa aprovada.
- Nunca executar git push, git push --force, criar ou modificar repositórios remotos ou publicar código sem autorização explícita.
- Nunca alterar as credenciais ou a configuração do GitHub.
- Não executar git reset --hard nem descartar alterações locais sem autorização.
- Após o commit, informar o hash, a mensagem e os arquivos incluídos.
- O usuário fará o git push manualmente pelo Git Bash.

## Interface planejada
- Interface minimalista, resumo técnico do PDF no topo e botão "Informações" com alertas amarelos/vermelhos quando pertinentes.
- Ferramentas agrupadas em um painel lateral contextual; miniaturas de páginas para seleção/organização; área principal de visualização ampla.
- União de PDFs em uma janela modal acessada por botão compacto (já implementada).
- Fluxo desejado: importar → unir/adicionar → girar/excluir/reorganizar/redimensionar → conferir → exportar.

## Registro de andamento e passagem de contexto

O desenvolvimento do 4PrintShops ocorre em diferentes ambientes:

- ChatGPT: planejamento, análise e decisões.
- Codex: implementação e testes no código.
- Work: operações complexas, investigações e validações no computador.
- GitHub: referência compartilhada do projeto.
- Computador dedicado: ambiente principal de desenvolvimento e testes.
- Computador de casa: desenvolvimento complementar.
- Estações da gráfica: utilização e homologação operacional.

Conversas de diferentes agentes não compartilham automaticamente todas as informações. A sincronização oficial depende de arquivos versionados no Git.

### Antes de começar

1. Ler os documentos obrigatórios.
2. Verificar o estado do Git.
3. Identificar a tarefa aprovada no planejamento.
4. Conferir se ela já foi implementada ou iniciada em outro computador.
5. Informar o escopo e eventuais riscos antes de alterar arquivos.

### Ao concluir uma etapa

1. Registrar no `docs/ESTADO_PROJETO.md` os detalhes técnicos relevantes, arquivos alterados, testes executados, limitações e resultados.
2. Atualizar `docs/CONTEXTO_ATUAL.md` com o último checkpoint, pendências, próximo passo e estado de sincronização.
3. Atualizar `docs/PLANEJAMENTO.md` quando houver uma decisão ou alteração de status devidamente confirmada.
4. Verificar o diff e preservar todo trabalho preexistente.
5. Criar commit local somente após autorização explícita.
6. Fazer push somente após autorização explícita.

Nunca apagar registros históricos para simplificar a documentação.

### Conflitos entre ambientes

Se houver alterações locais não publicadas, commits divergentes ou prioridades diferentes entre computadores:

- Interromper a implementação da tarefa conflitante.
- Identificar as diferenças.
- Preservar todos os arquivos.
- Apresentar as alternativas ao usuário.

Não utilizar `git reset --hard`, `git push --force` ou descartes automáticos para resolver divergências.

Nenhum computador deve ser considerado sincronizado apenas porque sua aplicação instalada possui a mesma versão. É necessário conferir os commits e o estado real do repositório.
