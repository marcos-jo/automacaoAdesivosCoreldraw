# Consolidação local das experimentais .24–.26

Data: 05/11/2026. Commit local expressamente autorizado pelo usuário, mensagem `Consolida livretos e melhorias de imposicao ate experimental 26`. Branch `main`, base `b08c7ce007fc72e0b86942857438b2e1eaa7357b`. Origin conferido com `git remote -v`; consulta somente leitura confirmou o mesmo hash em `refs/heads/main`. Não foi executado push. O hash final do checkpoint será informado na entrega do commit.

## Estado para retomar na gráfica

- **Experimental.24:** livretos grampeados ao meio já implementados no commit base, com paginação, frente/verso, brancos finais e prévias de impressão/leitura. Nenhuma reimplementação nesta consolidação. **Ainda sem homologação física com os funcionários da gráfica.** Guia [LIVRETOS_EXPERIMENTAL24.md](LIVRETOS_EXPERIMENTAL24.md) e [auditoria doméstica](AUDITORIA_DOMESTICA_EXPERIMENTAL24.md).
- **Experimental.25:** N-up usa todas as páginas por padrão. Seleção específica avançada validada antes do gerador, sem acessar página inexistente; mantém a prévia válida identificada durante erro de seleção e bloqueia saída até corrigir. Última folha não repete páginas. [Diagnóstico e testes](MONTAGEM_IMPOSICAO_EXPERIMENTAL25.md).
- **Experimental.26:** N-up inicia em tamanho real de 100%, com Reduzir somente se necessário e Preencher células como escolhas explícitas. Resumo e confirmações mostram escala/dimensões; conflitos bloqueiam saída inadequada. **Validação manual completa da .26 ainda pendente na gráfica.** [Decisão de escala e roteiro](MONTAGEM_IMPOSICAO_EXPERIMENTAL26.md).
- **Pendência conhecida:** zoom do editor desloca o ponto sob o cursor aproximadamente 53 px, medidos 53,45 px mesmo após 500 ms. A suíte geral do navegador foi reprovada nesse ponto; não foi corrigido nem considerado simples atraso do teste. O usuário autorizou avaliação experimental com essa ressalva.
- A correção .25 cobre a seleção inválida reproduzida. Não demonstra que todas as combinações intermitentes de desaparecimento da prévia .23 estejam resolvidas.

## Auditoria do escopo e preservação

Inventário inicial: 28 caminhos pendentes (19 modificados e 9 novos), staging vazio. Este documento acrescenta um arquivo: **29 arquivos para o checkpoint**, todos código-fonte, testes, configuração ou documentação. Os caminhos pessoais dos novos relatórios foram substituídos por caminhos relativos e variáveis padrão do Windows; o conteúdo técnico e o histórico foram mantidos.

Diferenças de produção auditadas: seleção pura nova, controles/estado/atualização da janela, políticas de escala nos planejadores e nas propostas de ajuste, resumo e versão Hold Test .24 → .26. A identidade de instalação, chave NSIS, origem `app://4printshops`, perfil `%APPDATA%/4PrintShops Hold Test` e retenção de dados permanecem iguais. O `package.json` principal mantém .6 conforme a configuração existente; `extraMetadata` da variante declara .26.

Step & Repeat conserva seus caminhos e opções independentes. O posicionamento compacto aprovado permanece; a ampliação nova depende de Preencher explícito. Livretos, materializador vetorial, renderizador, PDFWorkspace/editor, impressão direta, drivers, DEVMODE, atualizadores e dependências não receberam mudanças de produção neste diff. Preservação dos planejadores anteriores verificada na etapa .26: 180 planos válidos idênticos e 12 recusas iguais, incluindo 96 configurações Step & Repeat.

Nenhuma mudança existente foi descartada. Os 2.212 arquivos históricos de node_modules já rastreados permanecem inalterados. Instaladores, builds, dependências geradas, PDFs sintéticos/de clientes, perfis, logs, capturas e temporários ficam fora do staging. Os testes versionados geram seus próprios sintéticos; não incluem arquivos de clientes. Não foram encontradas credenciais nos candidatos auditados. Nenhum instalador adicional foi gerado ou instalado neste checkpoint, nem impressão física ou virtual enviada.

## Verificações

- Nova execução `npm test`: **390 aprovados**, zero falhas, cancelamentos ou skips; inclui seleção, prévia, políticas/saída vetorial, livretos, workspace e regressões existentes.
- Novas execuções `node tests/pdf-imposition-scale-policy-browser.mjs` no Chrome e no Electron de desenvolvimento (`NUP_SELECTION_ELECTRON=1`): aprovadas, 23 PDFs por ambiente com matrizes verificadas, políticas/dimensões, mudanças rápidas, bloqueios e recusa de confirmação antes de salvar/encaminhar. Não imprime.
- Nova execução `node scripts/test-print-native.mjs`: **41 checks C# aprovados**, sem API de envio ou alteração do Windows.
- Smoke Electron, seleção .25 no Chrome e integração Chrome/livretos já aprovados na etapa .26; logs finais reconferidos em `.desktop-test/nup26/`. Não equivalem a homologação física nem à suíte geral do navegador.
- Pacote existente reconferido com `node scripts/check-desktop-package.mjs dist-desktop/hold-native-test/win-unpacked/resources/app.asar`: 455 entradas, fontes iguais ao checkout e exclusões aprovadas. SHA256 do Setup existente reconferido, sem gerar outro instalador.
- Sintaxe dos **22 JS/MJS/CJS** novos ou modificados, `git diff --check`, inventário/exclusões e revisão de staging verificados antes do commit; conferência final com `git diff --cached --name-status` e `git diff --cached --check`.
- A suíte geral de navegador continua registrada como reprovada no zoom. Não foi renomeada como aprovada por causa das baterias pertinentes que passaram.

Logs da nova rodada ficam em `.desktop-test/checkpoint26/`, ignorados pelo Git. Relatórios anteriores preservam suas contagens e resultados históricos. A documentação de “sem commit” em etapas anteriores descreve o momento em que foram executadas.

## Inventário de arquivos

Caminhos relativos à raiz do aplicativo `Corel Script Configurator`:

| Grupo | Arquivos |
|---|---|
| Documentação (7) | `docs/CHECKPOINT_EXPERIMENTAIS24_26.md`, `docs/ESTADO_PROJETO.md`, `docs/DISTRIBUICAO_DESKTOP.md`, `docs/LIVRETOS_EXPERIMENTAL24.md`, `docs/AUDITORIA_DOMESTICA_EXPERIMENTAL24.md`, `docs/MONTAGEM_IMPOSICAO_EXPERIMENTAL25.md`, `docs/MONTAGEM_IMPOSICAO_EXPERIMENTAL26.md` |
| Configuração (1) | `electron-builder.hold-native-test.cjs` |
| Produção (7) | `public/js/tools/pdf-imposition-selection.js`, `pdf-imposition-settings.js`, `pdf-imposition-ui.js`, `pdf-imposition-layout.js`, `pdf-imposition-compact.js`, `pdf-imposition-bounds.js`, `pdf-imposition-summary.js` (todos no mesmo diretório) |
| Testes Node (6) | `tests/pdf-imposition-selection.test.mjs`, `pdf-imposition-scale-policy.test.mjs`, `pdf-imposition-preview.test.mjs`, `pdf-imposition-scale.test.mjs`, `pdf-imposition-compatibility.test.mjs`, `pdf-imposition.test.mjs` (todos no diretório `tests/`) |
| Testes de navegador (8) | `tests/pdf-imposition-browser.mjs`, `tests/pdf-imposition-selection-browser.mjs`, `tests/pdf-imposition-scale-policy-browser.mjs`; em `tests/helpers/`: `pdf-imposition-basic-browser.mjs`, `pdf-imposition-compact-browser.mjs`, `pdf-imposition-overflow-browser.mjs`, `pdf-imposition-scale-browser.mjs`, `pdf-imposition-scale-policy-browser.mjs` |

Total: 7 documentos + 1 configuração + 7 módulos + 6 testes Node + 8 roteiros/helpers = **29**. O inventário do staging é a conferência final dos caminhos exatos.

## Continuidade amanhã

1. Após o envio manual pelo usuário, abrir o clone da gráfica, ler AGENTS.md/ESTADO_PROJETO, conferir branch/remoto/status e recuperar este checkpoint com a árvore limpa. Não sobrescrever alterações locais da gráfica; se houver, comparar antes de sincronizar.
2. Conferir a versão real no código e no pacote. O Setup .26 já gerado fica fora do Git, em `dist-desktop/hold-native-test/4PrintShops-Hold-Test-1.0.0-experimental.26-Windows-x64-Setup.exe`, SHA256 `55DEE87B64FCB95B3C5F8CA4C80659BFC86EFA70AB9F6A0B43A63C720ED51ABE`. Não é transportado por push/pull. A auditoria do pacote está no guia .26; conferir hash de qualquer cópia antes da avaliação.
3. Verificar destino e versão instalada, salvar trabalhos pendentes e conferir perfil/preferências antes de uma atualização deliberada. Não substituir instalações automaticamente.
4. Validar manualmente N-up com três A6 em A3/4-up, quarta posição vazia, 100%/105×148 mm em células e compacto; quatro A6/A4 com giro autorizado; mistos, marcas, margens, seleção inválida/recuperação e alternância rápida. Conferir PDF e medidas físicas sem encaixe adicional oculto do driver.
5. Homologar separadamente livretos A5/A4 e A4/A3 com os funcionários: pares frente/verso, topo/base, eixo duplex/alimentação, dobra/aninhamento, brancos, registro, refile, grampo, papel/bandeja, área imprimível e RIP. Apenas um caderno; sem creep, múltiplos cadernos ou Cut & Stack. Não presumir homologação universal de cor/ICC/PDF-X.
6. Investigar o zoom em tarefa própria. Registrar qualquer nova falha de prévia com sequência de controles e console.

O checkpoint prepara a continuidade. **Não homologa .24 nem .26 para produção.** Push permanece exclusivamente manual, a cargo do usuário.
