# Tamanhos — exclusões sincronizadas e gerenciador global

## Revisão final e checkpoints — 17/11/2026

**Lista oficial confirmada integralmente pelo responsável**, inclusive A4 210×297 mm. A3 297×420, A5 148×210, A6 105×148, A7 74×105, F4 328×478, F8 238×328, F6 318×318 e Faixa 320×700 mm também conferidos no código e nos testes. Todos editáveis/excluíveis. As pendências provisórias de A4 nos registros históricos abaixo estão superadas; não exigem nova confirmação da lista.

Revisados os 18 arquivos pendentes (10 modificados/oito novos), preservados previamente em auditoria externa; escopo restrito ao catálogo, seletores, menu, testes e documentação autorizados. Nenhuma funcionalidade extra, instalador, perfil, PDF privado ou dependência incorporada. Código funcional mantido; divisão pelo índice e snapshots não sobrescreveram alterações do checkout.

Checkpoints organizados na ordem **A→C→B**, para não criar seletores duplicados ou imports de módulos ausentes:

- **A — `ab483e623be36ee6af1cb6068782e09b6594f9f6`**, Sincroniza formatos das montagens com o catalogo compartilhado: `public/js/tools/pdf-imposition-ui.js` e trecho da regressão `tests/pdf-sheet-catalog-browser.mjs` que remove a suposição de SRA3 estático.
- **C — `0e80a687fb211bf5e9b18e6b91e713afdc2f8a2a`**, Define catalogo inicial e recuperacao controlada de tamanhos: `public/js/tools/pdf-formats.js`, `pdf-format-backup.js`; `tests/pdf-formats.test.mjs`, `pdf-format-catalog.test.mjs`, `pdf-format-deletion.test.mjs`, `pdf-format-backup.test.mjs` e restante da adaptação `pdf-sheet-catalog-browser.mjs`.
- **B — checkpoint do gerenciador com estes registros:** `public/js/main.js`, `public/js/app/menus.js`, `public/js/tools/pdf-format-manager.js`, `tests/pdf-size-manager-browser.mjs`, `tests/helpers/pdf-size-manager-scenarios.mjs`, `tests/helpers/pdf-size-manager-electron.cjs` e quatro documentos (este relatório, PLANEJAMENTO, CONTEXTO_ATUAL, ESTADO_PROJETO). Hash final e Git pós-checkpoints no relatório externo `C:\Dev\4PrintShops_Auditoria_Checkpoints_Catalogo_20261117-SUPhnT\RELATORIO_CHECKPOINTS.md`.

Validações **novas desta revisão**, separadas das contagens históricas:

- Snapshot do índice A sem C/B: Edge catálogo/orientação/exportações sintéticas/reabertura aprovados, `%TEMP%\4ps-sheet-catalog-VYzL8W`.
- Snapshot do índice C sem B: **22/22** unidades de catálogo/backup e Edge aprovado, `%TEMP%\4ps-sheet-catalog-cAxxs3`.
- Conteúdo final A/C/B: **108/108**, execução serial única dos 13 arquivos (os 11 de catálogo/PDF já relacionados abaixo e `pdf-dimension-groups.test.mjs`/`desktop-print-presentation.test.mjs`). Equivale a 94+14 nesta execução, sem repetir suíte completa nem ensaios de empacotamento.
- Novo Edge completo: `%TEMP%\4ps-size-manager-cst0Yp`, quatro processos/três reinícios, zero erros.
- Novo Electron 45.1.0: `%TEMP%\4ps-size-manager-3Fg23w`, PIDs 18448/10060/18004/15704, quatro saídas normais código 0 em **173/215/160/121 ms**, três reinícios no mesmo perfil isolado, zero console.error/exceções. Conferidos exclusão definitiva (inclusive SRA3), vazio persistente, sincronização imediata, CRUD/IDs/metadados, orientação/arte, redefinição e restauração por backups sintéticos salvos/relidos/SHA verificado.
- Git/whitespace/sintaxe e integridade da separação conferidos. Complemento externo registra a checagem posterior ao último commit. A versão instalada, perfis e PDFs reais não foram usados.

### Redefinição operacional confirmada como procedimento técnico

Não existe botão de redefinição na interface; Arquivo → Gerenciar tamanhos oferece somente CRUD. A API `redefinir(backup)` não abre diálogo e **não é chamada silenciosamente por inicialização, menu ou atualização**. A confirmação humana deve ser obtida pelo responsável técnico, com o operador, **antes** de executar a chamada explícita. Uma futura atualização da .41 conserva a lista já persistida; não aplica automaticamente as nove sugestões.

Com autorização específica: exportar backup privado, gravar arquivo, reler e verificar SHA-256/conteúdo/estado atual; obter confirmação explícita do operador; só então chamar `redefinir` com o backup verificado. Conferir os nove formatos (incluindo F4 328×478, F8 238×328, F6 318×318 e Faixa 320×700), as três ferramentas e reabertura. Em recuperação autorizada, exportar também o estado corrente e usar `restaurarBackup` com o arquivo anterior verificado: restaura somente o texto da chave `4printshops.pdf.formatos.v1`, sem mesclar/restaurar o perfil inteiro. A guarda de backup em disco e a confirmação humana fazem parte do procedimento assistido; a API valida integridade/atualidade, mas não comprova sozinha que o arquivo foi guardado fora do aplicativo.

**Nenhuma redefinição real foi executada.** Para preparar a experimental.42 ainda faltam revisão/autorização de geração, metadados da versão e preflight (recursos, armazenamento), INST-VAL e auditoria do novo pacote. Não há impedimento funcional identificado nesta revisão do catálogo. Não alterar versão, gerar Setup, instalar, imprimir ou executar push agora. PDF-001/008/004/009 seguem Em testes, PDF-010 não iniciado.

## Decisão posterior de 17/11/2026 — catálogo inicial único e simplificação

Esta seção substitui o mecanismo de inicialização/compatibilidade descrito no ensaio anterior abaixo, que permanece como histórico. Não há exigência de compatibilidade completa com os catálogos de cada experimental antiga. Trabalho A/B anterior aproveitado, sem recomeçar interfaces e sem checkpoint.

### Lista inicial e estados

`FORMATOS_INICIAIS` em `pdf-formats.js` contém exatamente **A3 297×420; A4 210×297; A5 148×210; A6 105×148; A7 74×105; F4 328×478; F8 238×328; F6 318×318; Faixa 320×700 mm**, nesta ordem. A4 é a interpretação provisória aprovada para preparação; **confirmar A4 antes de qualquer redefinição operacional**. SRA3 não é sugestão inicial, mas pode ser cadastrado, editado, excluído e recadastrado como qualquer outro.

- **Primeiro uso:** chave `4printshops.pdf.formatos.v1` ausente (`null`); serviço oferece os nove registros. A leitura não grava. A primeira mutação explícita persiste a lista integral, `versao:1`, `catalogoCompleto:true`. Não há listas adicionais nos seletores das ferramentas.
- **Já inicializado:** qualquer documento v1 válido existente é exatamente a lista cadastrada, com ou sem o marcador. Removido o merge que acrescentava A3/A4/A5/A6 aos documentos legados. Seus IDs, nomes, medidas e metadados são preservados sem completar formatos ausentes. Catálogos existentes não recebem automaticamente os nove tamanhos.
- **Intencionalmente vazio:** documento existente com `formatos:[]`; continua vazio após recarga/novo processo. Excluir todos é permitido e não remove a chave. Nenhum formato é recriado automaticamente. Chave/dados inválidos bloqueiam edição em vez de gerar uma lista substituta.
- As referências ISO `FORMATOS_PADRAO` permanecem apenas para reconhecer dimensões técnicas nos resumos e apresentação de impressão; não são outra lista administrável nem protegem cadastros. As três ferramentas e o gerenciador usam somente o catálogo do serviço.
- B — menu e CRUD anteriores preservados. Exclusão em uso mantém medidas avulsas e não altera documentos; orientação/Girar arte permanecem independentes. Não há sincronização entre computadores.

### Redefinição futura assistida — não executada no perfil real

O mecanismo foi preparado no serviço, **sem adicionar um botão de reset ao aplicativo instalado**. API: `exportarBackup()`, `redefinir(backup)` e `restaurarBackup(backup)`. O módulo `pdf-format-backup.js` produz envelope JSON com texto original exato, lista efetiva, IDs, nomes, medidas e metadados, chave de destino e SHA-256 do conteúdo. O hash detecta corrupção; não é assinatura/autenticação de origem. No primeiro uso, o backup registra `original:null` e a lista efetiva de nove sugestões.

Procedimento para uma etapa futura, após autorização específica:

1. Confirmar a lista, especialmente A4; identificar o perfil/origem corretos e deixar o documento ocioso. Não limpar localStorage nem substituir o perfil Electron.
2. Em sessão assistida com o código revisado, executar `catalogo.exportarBackup()`, onde `catalogo` é o singleton de `obterCatalogoFormatos()`. Salvar o JSON **fora do Git e dos relatórios compartilhados**, em pasta privada. Reler o arquivo, conferir seu SHA-256 interno com `lerBackupFormatos(texto)` e registrar também o SHA-256 do arquivo. Conservar o texto original, inclusive metadados desconhecidos. Não basta guardar uma variável na memória.
3. Somente após revisão/autorização, passar o backup relido a `catalogo.redefinir(textoDoBackup)`. O serviço exige backup íntegro e correspondente à lista atual; alterações posteriores ou por outra instância obrigam novo backup. A escrita modifica somente a chave do catálogo. Nenhuma migração ampla, exclusão de perfil ou reinstalação é necessária.
4. Conferir os nove registros nas três ferramentas e no gerenciador; fechar/reabrir normalmente e conferir persistência, demais preferências e ausência de duplicações. Erro/aviso de armazenamento interrompe a validação; não declarar sucesso por fallback de memória nem repetir automaticamente.
5. Para recuperação autorizada, antes exportar/conferir o catálogo corrente também e então chamar `catalogo.restaurarBackup(backupAnteriorRelido)`. Backup corrompido, catálogo inválido e mudança externa são recusados antes da escrita. Quando havia documento persistido, a restauração repõe **o texto original exato da mesma chave**, incluindo IDs/metadados. Backup de primeiro uso restaura os nove registros como catálogo persistido. Conferir texto/hash/listas após reabertura. Nenhum dado de outras chaves é restaurado ou mesclado.

Essas APIs são de manutenção assistida e não são chamadas na inicialização, abertura do menu ou atualização do aplicativo. A redefinição real da .41 permanece **não autorizada e não executada**. Não é necessário adicionar compatibilidade reversa complexa; retorno a leitores antigos pode exigir recuperar o backup anterior sob revisão.

### Execuções desta simplificação

- **94/94**, execução serial nova com Node 26: os dez arquivos da rodada 87/87 abaixo mais `tests/pdf-format-backup.test.mjs` (sete testes novos). Cobre os nove formatos, CRUD/recadastro, legado sem complementação, vazio, backup/restore exatos, checksum divergente, backup obsoleto, conflitos, falhas de escrita/leitura/quota, preservação de outra chave, orientação/prévia/exportação e proteções PDF-008. Não é uma nova execução da suíte PDF completa.
- **14/14 adicionais:** `node --test --test-concurrency=1 tests/pdf-dimension-groups.test.mjs tests/desktop-print-presentation.test.mjs`. Somente unidades/mocks, sem impressão ou acesso ao driver operacional. Conferem consumidores das referências ISO e apresentação existente.
- Roteiro Edge `pdf-sheet-catalog-browser.mjs` aprovado: `%TEMP%\4ps-sheet-catalog-3qMhQJ`. A3/A4 agora são registros explícitos da fixture; não supõe mais que venham automaticamente do legado. Exportações sintéticas com medidas corretas, edição/exclusão, Web Locks, corrupção e reabertura continuam cobertos.
- Novo roteiro ampliado `pdf-size-manager-browser.mjs`: Edge `%TEMP%\4ps-size-manager-uzIeoC`, quatro PIDs/saídas normais e três reinícios. CRUD A/B, exclusão SRA3, sincronização imediata, catálogo vazio após reabrir, redefinição sintética e restauração com backup em arquivo SHA-256 verificado; zero erros de console.
- Mesmo roteiro com `--electron`: `%TEMP%\4ps-size-manager-gyi6IX`, quatro processos reais (21024/20036/18476/12916), três reinícios e saídas normais em **107/171/219/138 ms**, código 0. Origem app://4printshops, sandbox/contextIsolation, cópia atual conferida e perfil exclusivo. `catalogo-antes-esvaziar.json` e `catalogo-vazio.json` foram gravados, relidos e conferidos antes das operações sintéticas. Vazio não repovoado, nove registros após redefinição, catálogo original/sentinela restaurados, sem duplicação/erros.
- Sintaxe dos arquivos de código/testes pendentes e whitespace/diff conferidos. O teste inicial de quota precisava de um adaptador novo após a falha de leitura, pois o aviso conservador permanece; corrigido somente o fixture, preservando o bloqueio. Não foram repetidos os ensaios de instalador ou a suíte PDF completa; resultados antigos abaixo são históricos.

### Arquivos e divisão A/B atualizados

- **A:** `public/js/tools/pdf-formats.js`, `pdf-format-backup.js`, `pdf-imposition-ui.js`; `tests/pdf-formats.test.mjs`, `pdf-format-catalog.test.mjs`, `pdf-format-deletion.test.mjs`, `pdf-format-backup.test.mjs` e ajuste do fixture `pdf-sheet-catalog-browser.mjs`.
- **B preservada:** `public/js/tools/pdf-format-manager.js`, `public/js/main.js`, `public/js/app/menus.js`; integração em `tests/pdf-size-manager-browser.mjs` e seus dois helpers. Cobertura A/B compartilhada mantida logicamente identificada, para revisão antes de separar checkpoints.
- Continuidade: PLANEJAMENTO, CONTEXTO_ATUAL, ESTADO_PROJETO e este relatório, por acréscimos, preservando todas as decisões e evidências anteriores.
- HEAD `a3dfdf7`, main ahead 8 da referência local. Nenhum staging/commit/push, Setup, instalação, impressão, perfil operacional ou PDF real alterado. Próxima etapa: revisão A/B e do procedimento; confirmar A4 antes da futura redefinição real. Somente depois decidir checkpoints/experimental e homologação operacional. PDF-001/008/004/009 Em testes.

## Escopo e ocorrência operacional

Base `a3dfdf78a11c54ddd8df3c0678574c01800d1e2b`, main, oito commits locais à frente da referência origin/main; sem fetch/push/commit nesta tarefa. Os três documentos de continuidade já estavam modificados e seus registros foram preservados.

A atualização real .40→.41 anterior concluiu com transação `.4ps-upgrade-oCOQJR` committed. O registro `C:\Dev\4PrintShops_Auditoria_Atualizacao41_20261117-oH5gXG\pos-instalacao.json` confirmou uma instalação .41, 19.516 arquivos instalados iguais ao build e 76 arquivos do perfil iguais ao backup **antes da primeira abertura**. Não é uma nova auditoria nesta tarefa. Operador abriu o PDF sintético, confirmou funcionamento e informou F4/F6/F8 com as medidas esperadas no Redimensionar. Durante a homologação prática, relatou a permanência de SRA3 nas montagens após excluí-lo no Redimensionar. A homologação geral permanece pendente.

O responsável esclareceu que A3/A4/A5/SRA3 e os demais cadastros são administráveis, inclusive F4/F6/F8. A exigência de preservação destes últimos refere-se aos registros reais durante desenvolvimento/testes. A decisão substitui a hipótese anterior de modelos imutáveis da proposta de formatos; não autoriza catálogo em rede, PDF-010, Livreto ou reformulação visual do N-up.

## A — causa e correção de sincronização

- Causa comprovada no código: `pdf-imposition-ui.js` mantinha `presets` A3/A4/SRA3 e reconstruía essas opções em toda atualização, além de consultar somente os registros `!padrao` do catálogo. Excluir `custom-…` SRA3 removia o registro compartilhado, mas não a opção estática SRA3. A orientação dessa opção usava ainda dimensões estáticas. As notificações de CRUD já existiam.
- `PDFFormatos` também anexava A3/A4/A5/A6 a cada `listar()` e proibia sua edição/exclusão. Agora fornece registros administráveis e o mesmo catálogo alimenta as três ferramentas; nenhuma lista paralela nas montagens.
- Selecionar um formato usa suas medidas cadastradas, sem constantes que sobreponham edições. Alterar ou excluir o registro em uso desassocia a seleção para Personalizada/Personalizado e conserva a geometria atual, sem aplicar operações, exportar ou imprimir. O tamanho de folha permanece compartilhado na troca Step & Repeat/N-up; as opções próprias de cada modo continuam separadas.
- A3 continua selecionado inicialmente quando seu ID existe, preservando a apresentação atual. Quando não existe, usa-se a folha avulsa; A3 não é recriado por isso.

### Armazenamento, preservação e retorno — revisar antes de distribuir

1. Mesma chave `4printshops.pdf.formatos.v1`, mesmo serviço singleton e adaptador; nenhuma segunda chave/catalogação. Abrir/listar **não grava**. Nenhum perfil real foi convertido ou acessado por estes testes.
2. Documento legado é lido preservando registros, IDs, nomes, representação das medidas, metadados desconhecidos e extensões da raiz. A3/A4/A5/A6 são oferecidos inicialmente somente quando não houver ID/nome equivalente. SRA3 não é acrescentado a um documento legado onde esteja ausente, pois pode já ter sido excluído. Apenas um armazenamento completamente novo recebe SRA3 inicial.
3. A primeira operação explícita de salvar/editar/excluir persiste a lista integral no mesmo array `formatos`, com `versao:1` e marcador `catalogoCompleto:true`. Os registros legados permanecem primeiro, intactos, exceto o registro explicitamente editado/excluído. Esse marcador impede recriar tamanhos excluídos/renomeados, inclusive quando o catálogo fica vazio.
4. IDs iniciais permanecem estáveis; IDs `custom-…` continuam aceitos. Limite total 205 preserva até 200 registros legados mais os tamanhos iniciais. Nomes únicos normalizados e medidas de 0,1 a 5000 mm continuam validados. Dados inválidos/versão desconhecida bloqueiam gravação, preservando o texto original.
5. Serialização, Web Locks e recusa de escrita quando outro catálogo modificou o armazenamento continuam ativos. Uma seleção obsoleta do gerenciador é conferida novamente dentro da fila antes de editar/excluir; não sobrescreve a alteração mais recente. Falha de escrita que lance erro não publica lista nova. O fallback anterior de sessão com aviso permanece; não garante persistência quando o armazenamento falha.
6. **Retorno à .41 antiga não é transparente após a primeira edição no novo catálogo:** seu leitor só aceita IDs `custom-…` e reinsere os modelos antigos. Antes de instalar uma futura versão, guardar backup privado verificado do perfil e texto original desta chave. Em retorno autorizado, fechar completamente o aplicativo e restaurar a chave/perfil anterior somente após revisão; registrar separadamente cadastros feitos na versão nova para não perdê-los. Não tentar limpar/reinicializar o armazenamento. Este plano ainda exige revisão antes de implantação real; nenhum retorno/restauração foi executado.

## B — Arquivo → Gerenciar tamanhos

- Novo `public/js/tools/pdf-format-manager.js`, integrado por `main.js` e `app/menus.js`. Disponível inclusive no Início sem PDF. Usa o estilo dos diálogos existentes, nome acessível, lista e campos nome/largura/altura em mm.
- Novo/Adicionar, Salvar alterações e Excluir com confirmação simples. Cancelar a confirmação preserva o cadastro; editar mantém ID e metadados; nomes duplicados e medidas inválidas são recusados pelo serviço compartilhado.
- Publicação imediata das mudanças nos seletores das três ferramentas, inclusive instâncias já abertas. Formulário recebido de outra ferramenta não é silenciosamente sobrescrito, e uma edição/exclusão baseada em seleção desatualizada é recusada.
- Cadastro rápido do Redimensionar e salvar folha nas montagens continuam usando o mesmo serviço. Não há aplicação de operação PDF pelo gerenciador, novo armazenamento, IPC, alteração de segurança ou dependência instalada.

## Testes efetivamente executados

Node isolado `C:\Dev\4PrintShops-Tools\node-v26.13.1-win-x64\node.exe`. Electron 45.1.0 e dependências isoladas previamente preparadas; Edge instalado. Nenhum npm install/ci, build, Setup ou impressão.

- **87/87 regressões dirigidas, execução serial:** `node --test --test-concurrency=1 tests/pdf-formats.test.mjs tests/pdf-format-catalog.test.mjs tests/pdf-format-deletion.test.mjs tests/pdf-sheet-orientation.test.mjs tests/pdf-imposition-preview.test.mjs tests/pdf-resize.test.mjs tests/pdf-imposition.test.mjs tests/pdf-imposition-compatibility.test.mjs tests/pdf-signature-policy.test.mjs tests/pdf-signature-presentation.test.mjs`. Inclui exclusão de todos os tamanhos iniciais, reabertura de catálogo vazio, legado sem SRA3, falha de gravação, concorrência e seleção obsoleta, prévias, exportação sintética e proteções PDF-008. As contagens 567/567 e outras suítes completas são referências anteriores, **não reexecutadas**.
- `node tests/pdf-sheet-catalog-browser.mjs`: passou orientação repetida/papel quadrado, PDF sintético com dimensões corretas, formatos legados preservados, edição/exclusão sem aplicar geometria, N-up, reabertura real, Web Locks/recusa de instância desatualizada e armazenamento corrompido não sobrescrito. Evidências finais `%TEMP%\4ps-sheet-catalog-fQRPXk`. Retirado apenas o teste da antiga opção estática SRA3 inexistente no legado; a exclusão SRA3 tem agora cobertura própria.
- `BROWSER_PATH=<Edge> node tests/pdf-repeat-interface-browser.mjs`: aprovado em `.desktop-test/repeat-interface-Y6WPyd`, incluindo interface básica existente, orientação, giro, grade, duplex, teclado, exportação e aceite de excesso; console vazio. Chrome padrão não estava disponível; usado Edge explicitamente.
- `node tests/pdf-size-manager-browser.mjs`: CRUD por eventos reais do menu/Redimensionar, duas instâncias das montagens, exclusão em uso, nome duplicado, confirmação cancelada/aceita, edição obsoleta, A3 editável/A5 excluível, orientação e arte 90° independentes, preservação de F4/F6/F8 sintéticos e sentinela. Três processos Edge, duas reaberturas, zero erros de console; `%TEMP%\4ps-size-manager-vqDcSs`.
- `node tests/pdf-size-manager-browser.mjs --electron`: mesma sequência na origem real `app://4printshops`, sandbox/contextIsolation e nodeIntegration=false. Cópia nova de desktop/public/package.json conferida por hash, dependências isoladas via junction, userData e temporários exclusivos. Três processos Electron e dois reinícios reais; catálogo bruto idêntico após reabertura, exclusões mantidas, sem duplicações ou console.error/exceções. Evidências finais `%TEMP%\4ps-size-manager-isFaWY`; PIDs 19052/13568/22216, saídas código 0 em 242/190/159 ms com app.quit/before-quit/will-quit/quit. Nenhuma saída forçada. O roteiro inicial precisou de correções de espera do carregamento/ready do main e encapsulamento de await; falhas intermediárias preservadas nas pastas de teste.
- Captura do gerenciador inspecionada visualmente: campos/botões legíveis, diálogo no estilo existente, sem sobreposição de controles. Sintaxe dos JS/CJS/MJS alterados e `git diff --check` também conferidos.

Os testes usaram formulários reais e instâncias reais de `JanelaMontagem` com snapshots sintéticos; não imprimiram, não exportaram PDFs reais e não usaram a instância instalada. Os cenários de catálogo não constituem nova homologação operacional da .41 nem auditoria de instalador.

## Separação para revisão e próxima etapa

- **A:** `public/js/tools/pdf-formats.js`, `pdf-imposition-ui.js`, `tests/pdf-formats.test.mjs`, `pdf-format-catalog.test.mjs`, `pdf-format-deletion.test.mjs` (exceto o caso B) e atualização do roteiro `pdf-sheet-catalog-browser.mjs`.
- **B:** novo gerenciador, `public/js/main.js`, `public/js/app/menus.js`, teste B de seleção obsoleta e novos `tests/pdf-size-manager-browser.mjs`, `tests/helpers/pdf-size-manager-scenarios.mjs`, `tests/helpers/pdf-size-manager-electron.cjs`. As evidências de integração cobrem A e B conjuntamente; não forçar separação que omita dependências.
- Documentação atualizada por acréscimos; nenhum registro anterior removido. Sem staging/commit/push. Aguardar revisão dos dois conjuntos e do plano de armazenamento/retorno antes de checkpoints ou preparação de outra experimental. Ainda não escolher/alterar versão.
- Não houve alteração do instalador/atualizador, proteções PDF, mecanismos de impressão/exportação, perfil real, formatos reais F4/F6/F8 ou biblioteca privada. A .41 instalada continua com o defeito até uma futura distribuição autorizada; sua homologação prática não é declarada concluída.
