# 4PrintShops — Contexto Atual e Passagem de Trabalho

## Checkpoint autorizado — 13/11/2026, PDF-001 Em testes

- Usuário autorizou commit exclusivamente local da implementação PDF-001 e seus testes/documentos, com mensagem `Corrige atributo Tabs no recorte de PDFs`. Base antes do commit: main/118032e, sincronizada com origin/main. Hash será informado após a criação; sem push autorizado.
- Código funcional inalterado desde a implementação revisada. A versão de distribuição continua experimental.39; não houve novo pacote, instalação ou impressão.
- Validações anteriores: 84 testes de recorte (74 novos), regressões relacionadas, 28 testes transacionais e Chrome sintético aprovados. Não repetidos para este checkpoint.
- Auditoria local de 24 PDFs autorizados: sete documentos únicos com recorte aplicado/exportado/reaberto pelos módulos, seis com Tabs /S preservado; uma operação adicional pela interface no mesmo CASO-002 (Word), totalizando oito PDFs de recorte. Dois documentos PowerPoint (CASO-008/009) continuam corretamente bloqueados por StructTreeRoot; não homologados para recorte.
- Revisão final conferiu dimensões esperadas dos sete recortes, Tabs, reabertura e hashes dos 11 resultados. Os 24 originais e as cópias preservadas permaneceram intactos. Comparação visual completa em leitor independente e homologação operacional continuam pendentes; PDF-001 permanece Em testes.
- Biblioteca, inventários, PDFs e evidências ficam exclusivamente em .desktop-test/casos-reais, ignorados pelo Git. O commit inclui somente os seis arquivos autorizados, sem dados ou nomes de clientes.
- Risco separado e preexistente: inspeção e estrutura sintética em memória mostraram divergência entre detector de assinaturas e diagnóstico de montagem para Perms/DocMDP sem formulário/anotações. A montagem desse caso não foi executada. Rotação e união não possuem bloqueio universal; preservação criptográfica não demonstrada. Nenhuma correção de assinaturas iniciada, nenhuma prioridade alterada.
- Próximo passo: aguardar nova autorização do usuário. Checkpoint de implementação não significa homologação nem autorização de push ou de nova experimental. Os registros abaixo são históricos.

## Checkpoint — 12/11/2026, computador dedicado, PDF-001

- Revisão complementar: diff de código somente lido; nenhuma suíte repetida. Validação de PDFName e bloqueios preservados. A/W pertencem ao PDF 2.0 e às extensões Adobe 1.7; não há certificação de versão/extensão. Nenhuma amostra real encontrada no projeto. Pastas locais ignoradas preparadas em .desktop-test/pdf001-real/{originais,resultados,evidencias}; aguardando caminhos e autorização de amostras Word/PowerPoint/outra ferramenta. Riscos e roteiro no novo registro de revisão de ESTADO_PROJETO.md. Tarefa continua Em testes.

- HEAD e origin/main: 118032e; a8e25f8 (.39) presente no histórico publicado. As seções antigas abaixo registram o estado anterior à sincronização.
- PDF-001 autorizada pelo usuário e implementada localmente; em testes, aguardando revisão. Somente prepararRecorte ganhou validação explícita de Tabs R/C/S/A/W, com preservação e bloqueios anteriores mantidos.
- 84 testes de recorte aprovados; regressões relacionadas e 28 testes transacionais aprovados. Chrome headless aprovou recorte/exportação/reabertura com Tabs /S. Sem instalação de dependências; usadas bibliotecas existentes e dependências transacionais isoladas.
- Alterações ainda não publicadas: public/js/tools/pdf-workspace.js, tests/pdf-crop.test.mjs, tests/pdf-crop-browser.mjs e estes três documentos. Sem commit/push.
- Pendências: revisão do usuário e PDFs representativos Word/PowerPoint/comuns; estrutura marcada e anotações reais continuam bloqueadas. Não preparar nova experimental antes da revisão. Detalhes no checkpoint PDF-001 de ESTADO_PROJETO.md.


**Atualizado em:** 12/11/2026
**Ambiente de referência:** gráfica
**Projeto:** 4PrintShops Hold Test
**Branch principal:** `main`

## 1. Finalidade

Este documento registra a situação mais recente do projeto para permitir a continuidade do desenvolvimento entre a gráfica, o computador dedicado de testes e o computador de casa.

Consultar também:

- `AGENTS.md` — regras permanentes.
- `docs/PLANEJAMENTO.md` — prioridades e tarefas planejadas.
- `docs/ESTADO_PROJETO.md` — histórico técnico detalhado.

Este documento não substitui a inspeção do Git ou dos arquivos reais.

## 2. Situação do Git

Repositório:

`https://github.com/marcos-jo/automacaoAdesivosCoreldraw`

Pasta principal:

`Corel Script Configurator/`

**Último commit local informado:**

`a8e25f8 — Corrige instalador e registra experimental 39`

**Último commit remoto confirmado:**

`7558bf5 — Gera e valida experimental 38 com recorte de pagina`

Na última conferência, a branch local estava um commit à frente de `origin/main`.

O commit da experimental.39 ainda não havia sido enviado ao GitHub.

Após esse checkpoint, foi iniciado o trabalho de documentação:

- `docs/PLANEJAMENTO.md` — criado manualmente, aguardando commit.
- `docs/CONTEXTO_ATUAL.md` — documento em preparação.
- `AGENTS.md` — atualização das regras de continuidade ainda pendente.

Antes de alterar arquivos, conferir novamente `git status -sb` e `git log -1 --oneline`.

## 3. Versão atual

**Versão experimental mais recente na gráfica:** 1.0.0-experimental.39.

Principais alterações:

- Correção das regras de atualização entre versões experimentais.
- Suporte à reinstalação da própria versão.
- Tratamento seguro de determinados resíduos de desinstalação.
- Preservação de arquivos não gerenciados.
- Melhoria da mensagem de erro e do diagnóstico de rollback.

A documentação do Work registra 520 testes Node aprovados e uma instalação bem-sucedida da .39.

Em 12/11/2026, foram executados manualmente os 28 testes específicos de `tests/update-transaction.test.mjs`:

- 28 aprovados.
- 0 falhas.
- 0 ignorados.

A verificação de sintaxe passou.

Esses resultados não representam homologação completa do sistema em produção.

## 4. Situação dos computadores

### Computador de trabalho da gráfica

- Contém o checkout local com a experimental.39 registrada no commit `a8e25f8`.
- A aplicação .39 foi instalada e aberta com sucesso, conforme os relatos e registros existentes.
- Os arquivos de planejamento estão sendo preparados neste ambiente.
- Não realizar intervenções de instalação ou limpeza sem autorização.

### Computador dedicado ao desenvolvimento

- Foi preparado pelo usuário para desenvolvimento e testes.
- Possui aproximadamente 100 GB livres, conforme relato inicial.
- VS Code com Codex e aplicativo ChatGPT instalados e conectados.
- Ainda precisa receber o repositório sincronizado.
- Será utilizado para implementar, testar e gerar novas versões antes da distribuição.

### Computador de casa

- Utilizado para desenvolvimento complementar.
- O usuário informou que não realizou novas alterações no projeto na noite de 11/11/2026.
- O estado exato do Git local deve ser conferido antes de qualquer sincronização.
- Não presumir que arquivos locais ou evidências não versionadas foram transferidos pelo GitHub.

## 5. Problemas conhecidos

### Recortar página — atributo /Tabs

A experimental.39 apresenta o bloqueio:

`Recorte bloqueado: atributo Tabs ainda não suportado.`

Relatos de ocorrência:

- PDFs comuns.
- PDFs originados no Word.
- PDFs exportados do PowerPoint.

PDFs criados a partir de imagens funcionaram nos testes relatados.

A causa provável está nas verificações conservadoras de atributos de página em `pdf-workspace.js`, identificadas na análise da versão publicada .38.

A implementação local da .39 deverá ser inspecionada antes de modificar o código.

**Tarefa:** PDF-001 no planejamento.

### Espaço temporário do instalador

Houve uma tentativa de instalação com esgotamento do espaço disponível em C:.

A correção da .39 resolveu problemas de compatibilidade e resíduos, mas não foi demonstrada uma redução do pico de consumo de espaço.

**Tarefa:** INST-001.

### Homologações pendentes

- Step & Repeat por grade em cenários reais de produção.
- Frente e verso com impressora e RIP.
- Documentos Word complexos.
- Tratamento completo de caixas de corte e sangria.
- Compatibilidade com PDFs complexos.
- Validação mais ampla antes da distribuição.

Consultar o planejamento para os respectivos IDs.

## 6. Próxima tarefa autorizada

**SYS-001 — Organização e sincronização do desenvolvimento.**

Sequência:

1. Finalizar `PLANEJAMENTO.md` e `CONTEXTO_ATUAL.md`.
2. Atualizar `AGENTS.md` com as regras de sincronização.
3. Revisar as alterações documentais.
4. Criar um commit local separado do checkpoint da experimental.39, após autorização.
5. Conferir os commits e executar push autorizado para o GitHub.
6. Clonar o repositório atualizado no computador dedicado.
7. Confirmar que o computador dedicado possui a mesma versão e os mesmos documentos.
8. Iniciar a próxima tarefa aprovada no planejamento.

A correção do atributo `/Tabs` será tratada em uma etapa separada.

## 7. Regras para retomar o trabalho

Antes de implementar:

- Conferir o Git local e remoto.
- Identificar alterações não publicadas.
- Ler o planejamento e confirmar a prioridade aprovada.
- Verificar se a tarefa já foi iniciada em outro ambiente.
- Preservar alterações anteriores, originais, perfis e backups.
- Não fazer reset destrutivo, instalação, impressão ou push sem autorização.

Ao concluir uma tarefa:

- Informar arquivos modificados.
- Registrar testes e limitações.
- Atualizar o planejamento, quando houver mudança de status autorizada.
- Atualizar este documento com a nova situação.
- Registrar o histórico técnico em `ESTADO_PROJETO.md` quando apropriado.
- Informar explicitamente se o trabalho foi enviado ao GitHub.

## 8. Pendência de sincronização

**Situação no momento da criação deste documento: não sincronizado.**

O checkpoint local da experimental.39 e os arquivos novos de planejamento ainda precisam ser publicados.

Não considerar o computador dedicado ou o computador de casa atualizados até que o Git seja conferido nesses ambientes.

---

**Próximo passo:** finalizar a documentação de continuidade, revisar e preparar o segundo commit.
