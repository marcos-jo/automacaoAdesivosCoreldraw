# 4PrintShops — Planejamento Geral de Implementações

**Revisão:** 0.1 — Rascunho inicial
**Data:** 12/11/2026
**Responsável pelas prioridades:** usuário responsável pelo projeto
**Situação:** aguardando revisão e aprovação das prioridades

## 1. Objetivo

Este documento é a referência central para o planejamento do 4PrintShops, sistema de ferramentas para gráfica.

Seu objetivo é permitir que diferentes agentes de IA e computadores trabalhem no mesmo projeto, respeitando as decisões já tomadas e evitando retrabalho, implementações duplicadas ou mudanças não autorizadas de prioridade.

Documentos complementares:

- `AGENTS.md` — regras permanentes para os agentes.
- `docs/CONTEXTO_ATUAL.md` — situação recente, checkpoints e passagem de trabalho.
- `docs/ESTADO_PROJETO.md` — histórico técnico detalhado e evidências.

Este arquivo organiza o trabalho futuro; não substitui a inspeção do código.

## 2. Situação atual do projeto

**Última versão registrada localmente:** 1.0.0-experimental.39
**Commit local:** `a8e25f8` — Corrige instalador e registra experimental 39
**Último commit remoto confirmado:** `7558bf5` — experimental.38
**Sincronização:** o commit da .39 ainda aguarda push.

A experimental.39 corrigiu problemas de atualização e reinstalação do aplicativo.

A auditoria específica do instalador realizada em 12/11/2026 aprovou 28 testes, sem falhas. O histórico registra 520 testes Node aprovados anteriormente pelo Work; essa suíte completa não foi repetida nesta auditoria.

O aplicativo .39 foi instalado e abriu com sucesso na gráfica, conforme os registros disponíveis. Isso não representa homologação geral de todas as funções.

O problema de recorte envolvendo o atributo PDF `/Tabs` continua pendente.

## 3. Regras de planejamento

1. Este documento será a lista oficial de implementações após sua aprovação.
2. Cada tarefa deve possuir um identificador permanente.
3. Nenhum agente pode excluir, alterar prioridades aprovadas ou declarar uma tarefa concluída sem justificativa, evidências e concordância do usuário.
4. Novas ideias devem entrar como **Proposta**, sem substituir automaticamente as prioridades existentes.
5. Toda tarefa deve possuir objetivo, dependências e critérios de aceitação.
6. A implementação e a homologação operacional são etapas diferentes.
7. Antes de iniciar uma tarefa, conferir o Git, o contexto atual e eventuais mudanças locais não publicadas.
8. Alterações realizadas em casa e na gráfica só serão consideradas sincronizadas depois do envio e recebimento dos commits correspondentes.
9. Não fazer push, apagar alterações, instalar aplicativos ou executar impressão sem autorização específica.
10. Ao terminar uma etapa, atualizar a situação da tarefa e registrar os resultados técnicos.

### Estados permitidos

- **Proposta:** identificada, ainda não autorizada.
- **Aprovada:** autorizada, aguardando implementação.
- **Em desenvolvimento:** implementação iniciada.
- **Bloqueada:** depende de uma condição não atendida.
- **Em testes:** implementação existente, aguardando validação.
- **Concluída:** critérios de aceitação cumpridos e conclusão confirmada.

### Prioridades

- **P0:** bloqueio, segurança, integridade ou problema importante para continuar.
- **P1:** próxima implementação relevante.
- **P2:** melhoria planejada.
- **P3:** evolução futura.

As prioridades na primeira versão são sugestões, exceto quando houver uma decisão explícita registrada.

## 4. Lista principal de implementações

| ID | Prioridade sugerida | Situação | Tarefa |
|---|---|---|---|
| SYS-001 | P0 | Em preparação | Organizar planejamento e sincronização entre computadores |
| PDF-001 | P0 | Proposta | Corrigir bloqueio de recorte por atributo `/Tabs` |
| INST-001 | P1 | Proposta | Investigar e reduzir o pico de espaço em disco do instalador |
| QA-001 | P1 | Em testes | Homologar Step & Repeat por grade e frente e verso |
| PDF-002 | P1 | Proposta | Reconhecer e preservar corretamente caixas de corte e sangria |
| WORD-001 | P1 | Em testes | Ampliar homologação de importação DOC/DOCX |
| PDF-003 | P2 | Proposta | Permitir senha para abertura de PDFs protegidos, quando autorizado |
| QR-001 | P2 | Proposta | Importar planilhas XLSX e gerar QR Codes de contatos em lote |
| PDF-004 | P2 | Proposta | Ampliar compatibilidade de PDFs complexos |
| PDF-005 | P2 | Proposta | Implementar Desfazer/Refazer |
| PDF-006 | P2 | Proposta | Integrar extração de páginas ao fluxo do workspace |
| PDF-007 | P2 | Proposta | Inspeção técnica de cores, fontes, imagens e perfis |
| IMP-001 | P2 | Em testes | Homologar fisicamente a montagem de livretos |
| DESK-001 | P3 | Proposta | Evoluir distribuição e atualização assistida |
| IMP-002 | P3 | Proposta | Avaliar Cut & Stack, múltiplos cadernos e creep |
| LEG-001 | Sem prioridade | Preservado | Manter legado CorelDRAW separado da aplicação principal |

**Observação:** esta tabela não constitui autorização para iniciar as tarefas. A ordem definitiva será aprovada pelo usuário.

## 5. Detalhamento das próximas etapas

### SYS-001 — Continuidade e sincronização

**Objetivo:** centralizar as decisões e permitir alternar com segurança entre gráfica, computador dedicado e casa.

Critérios de aceitação:

- Criar `docs/PLANEJAMENTO.md`.
- Criar `docs/CONTEXTO_ATUAL.md`.
- Atualizar `AGENTS.md` com regras de leitura e sincronização.
- Preservar `docs/ESTADO_PROJETO.md`.
- Registrar os documentos em commit separado do código funcional.
- Enviar os commits autorizados ao GitHub.
- Clonar e conferir o repositório no computador dedicado.
- Confirmar que os agentes reconhecem as mesmas prioridades.

### PDF-001 — Correção do atributo /Tabs

**Problema:** PDFs comuns, PDFs gerados pelo Word e PDFs exportados do PowerPoint apresentaram a mensagem:

`Recorte bloqueado: atributo Tabs ainda não suportado.`

PDF criado a partir de imagem funcionou no teste relatado.

**Objetivo:** aceitar os casos válidos do atributo `/Tabs` sem enfraquecer as restrições de segurança do recorte.

Critérios de aceitação:

- Inspecionar o comportamento real na versão atual.
- Validar os valores e a semântica do atributo.
- Preservar os bloqueios de documentos que não possam ser processados com segurança.
- Acrescentar testes de regressão.
- Confirmar que outras páginas e os arquivos originais permanecem intactos.
- Validar PDFs representativos do Word, PowerPoint e PDFs comuns.

Não remover indiscriminadamente todas as verificações de atributos PDF.

### INST-001 — Consumo de espaço do instalador

**Problema:** uma tentativa de instalação da .39 encontrou falta de espaço em C: durante o processamento.

**Objetivo:** medir o espaço temporário necessário e avaliar alternativas que não comprometam a integridade ou a recuperação da instalação anterior.

Critérios de aceitação:

- Medir o consumo real das etapas de instalação.
- Identificar quais cópias temporárias são indispensáveis.
- Estimar o espaço mínimo exigido antes de iniciar.
- Impedir atualizações inseguras por falta de espaço.
- Preservar rollback, perfil e arquivos não gerenciados.
- Executar testes em ambiente isolado antes de qualquer instalação operacional.

### QA-001 — Step & Repeat e duplex

**Situação:** a montagem por linhas e colunas foi implementada na experimental.37.

Permanecem pendentes verificações operacionais e físicas.

Critérios de aceitação:

- Testar grade 3 × 3 e outras configurações.
- Conferir quantidades, marcas e posições.
- Verificar frente e verso, borda longa e borda curta.
- Comparar resultado com impressora/RIP da gráfica.
- Registrar eventuais diferenças de posicionamento.

Nenhuma impressão física sem autorização.

### PDF-002 — Corte e sangria

**Objetivo:** melhorar o tratamento de `MediaBox`, `CropBox`, `TrimBox` e `BleedBox`.

A primeira implementação de recorte modifica as caixas da página, mas não cria sangria nem apaga permanentemente conteúdo oculto.

Critérios de aceitação:

- Identificar caixas existentes com confiabilidade.
- Diferenciar corte visual de remoção definitiva de conteúdo.
- Preservar sangria válida quando tecnicamente possível.
- Informar claramente quando a sangria não estiver verificada.
- Validar exportação, montagem e comportamento no RIP.

### WORD-001 — Compatibilidade Word

**Situação:** importação de um documento Word simples foi aprovada em teste operacional. DOC e DOCX sintéticos também passaram em testes técnicos.

Pendente:

- Documentos com imagens e tabelas.
- Fontes diferentes e substituições.
- Documentos extensos.
- Paginação e layout.
- Desempenho em lotes.
- Compatibilidade de impressão.

Não considerar todos os documentos Word homologados com base no teste simples.

## 6. Outras implementações planejadas

### PDF-003 — PDFs protegidos

A conferência individual de PDFs protegidos e inválidos já foi implementada.

A próxima evolução é permitir informar senha, quando houver autorização, e preparar uma cópia adequada para processamento. Não contornar restrições nem comprometer os originais.

### QR-001 — QR Codes por planilhas

Importar planilhas XLSX de contatos e permitir:

- Mapear as colunas para campos de contato.
- Selecionar quais dados entram no QR Code.
- Gerar vCards compatíveis.
- Produzir vários QR Codes em lote.
- Exportar PNG, SVG e ZIP.
- Validar os resultados antes da exportação.

O módulo atual de QR Codes já existe, mas essa importação automatizada ainda é futura.

### PDF-004 — Compatibilidade avançada

Investigar separadamente documentos com formulários, anotações, estruturas marcadas, destinos internos e outros recursos que hoje possuem restrições de processamento.

Não remover bloqueios preventivos sem testes de fidelidade.

### PDF-005 — Desfazer e refazer

Planejar histórico de operações do PDFWorkspace, com recuperação segura de edições e controle de consumo de memória.

### PDF-006 — Extração integrada

Evoluir a extração de páginas, que ainda utiliza download direto no fluxo legado, para permitir abrir ou preparar outro documento de trabalho sem exportação automática.

### PDF-007 — Inspeção gráfica

Planejar inspeção confiável de:

- RGB, CMYK e cores especiais.
- Perfis ICC.
- Fontes incorporadas ou substituídas.
- Resolução efetiva das imagens.
- Caixas de corte e sangria.
- Alertas de pré-impressão.

Até existir análise confiável, apresentar **Não verificado**.

### IMP-001 e IMP-002 — Livretos e imposição avançada

Homologar primeiro a montagem de livreto grampeado e o duplex físico.

Somente depois avaliar funcionalidades adicionais, como Cut & Stack, múltiplos cadernos e compensação de lombada (creep).

### DESK-001 — Distribuição e atualização

Avaliar futuramente o fluxo de atualização assistida, assinatura e distribuição entre máquinas, sem reduzir proteções de integridade ou segurança.

### LEG-001 — CorelDRAW

A automação antiga de Adesivos/CorelDRAW foi retirada da aplicação principal em 30/10/2026.

Preservar código, modelos e documentação legados. Uma retomada ou integração futura dependerá de decisão separada.

## 7. Implementações importantes já existentes

Estas funcionalidades não devem ser recriadas por outro agente sem antes conferir o código:

- Gerenciador de PDFs e PDFWorkspace.
- Importação, união, exclusão e rotação de páginas.
- Miniaturas, reorganização e controles de visualização.
- Redimensionamento vetorial.
- Proteção contra perda de alterações.
- Exportação PDF e imagens.
- Montagem e imposição de PDFs.
- Step & Repeat por grade.
- Primeira etapa de Recortar página.
- Diagnóstico de PDFs protegidos em lotes mistos.
- Conversão Word com LibreOffice privado.
- Aplicativo Electron para Windows.
- Atualização transacional e rollback.
- Configurações e recursos de impressão Hold Test.
- Módulo inicial de QR Codes.

A existência no código não significa homologação universal.

## 8. Processo de execução

Antes de iniciar:

1. Consultar `AGENTS.md`, `PLANEJAMENTO.md`, `CONTEXTO_ATUAL.md` e `ESTADO_PROJETO.md`.
2. Conferir branch, HEAD, alterações locais e situação do remoto.
3. Confirmar o ID da tarefa e sua autorização.
4. Inspecionar a implementação existente para evitar refazer trabalho.
5. Definir escopo, arquivos afetados e testes de aceitação.

Durante a execução:

- Fazer alterações pequenas e isoladas.
- Preservar originais e recursos existentes.
- Executar testes proporcionais ao risco.
- Registrar impedimentos e descobertas relevantes.
- Não mudar a ordem das tarefas sem aprovação.

Ao encerrar:

1. Documentar o que mudou.
2. Registrar testes executados e não executados.
3. Atualizar o status da tarefa.
4. Atualizar o contexto de passagem de trabalho.
5. Preparar um checkpoint Git quando autorizado.
6. Informar se existem alterações não sincronizadas.

## 9. Registro de decisões

| Data | Decisão ou registro |
|---|---|
| 30/10/2026 | Legado CorelDRAW preservado e separado da aplicação principal |
| 10/11/2026 | Step & Repeat por grade registrado na experimental.37 |
| 10/11/2026 | Primeira etapa do recorte vetorial implementada |
| 11/11/2026 | Experimental.39 preparada com correções do instalador |
| 12/11/2026 | Auditoria específica: 28 testes do instalador aprovados |
| 12/11/2026 | Commit local `a8e25f8` criado para preservar a .39 |
| 12/11/2026 | Decidida a organização do trabalho entre chat de planejamento, Codex, Work e GitHub |

## 10. Mudanças futuras de prioridade

Toda alteração de prioridade aprovada deverá registrar:

- Data.
- ID da tarefa.
- Prioridade anterior e nova.
- Motivo da alteração.
- Aprovação do usuário.
- Commit que contém a mudança, quando disponível.

O histórico de decisões não deve ser apagado.

## 11. Próxima ação

Consolidar este planejamento com as listas que possam existir nas conversas do computador de casa.

Após aprovação, finalizar a documentação compartilhada, sincronizar os commits e preparar o computador dedicado.

**Nenhuma implementação nova é autorizada automaticamente por este documento.**