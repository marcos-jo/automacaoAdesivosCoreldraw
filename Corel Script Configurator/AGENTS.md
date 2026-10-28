# AGENTS.md — 4PrintShops

## Objetivo
Este repositório contém o 4PrintShops, uma central de ferramentas locais para o fluxo de trabalho de uma gráfica. O usuário trabalha com o projeto no Windows e prefere desenvolver em etapas pequenas, validando cada mudança no navegador. Converse com o usuário em português do Brasil.

## Contexto obrigatório
Antes de alterar código, leia `docs/ESTADO_PROJETO.md`. Esse arquivo registra o estado conhecido do projeto e o próximo objetivo. Sempre confira o código real do repositório: o registro é contexto, não substitui inspeção dos arquivos.

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
Após realizar uma alteração relevante, atualizar `docs/ESTADO_PROJETO.md` com: data, arquivos alterados, problema resolvido, testes executados, pendências e próximo passo. Não apagar o histórico recente nem afirmar que funções não testadas estão prontas.

Esta conversa no Codex e uma conversa separada no ChatGPT **não trocam mensagens automaticamente**. O usuário poderá compartilhar o texto de `docs/ESTADO_PROJETO.md` na outra conversa para sincronizar o contexto. Não alegar acesso direto ao outro assistente.
