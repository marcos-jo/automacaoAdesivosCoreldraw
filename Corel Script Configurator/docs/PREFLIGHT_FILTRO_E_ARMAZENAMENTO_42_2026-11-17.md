# Preflight .42 — filtro e armazenamento (17/11/2026)

## 18/11/2026 — Ambiente isolado .42 recuperado e validações pendentes aprovadas

- Responsável autorizou corrigir exclusivamente o ambiente isolado e concluir o Setup se todas as verificações passarem; instalar/publicar/push continuam proibidos. Causa confirmada: checkout ativo com 2.212 arquivos de dependências históricos versionados, enquanto validações anteriores usavam a árvore completa de 9.905 arquivos. A preparação .42 anterior ainda não tinha source/dependências; a chamada foi feita diretamente no checkout incompleto. Os alvos da limpeza não incluíam o checkout/preparo preservados, e o procedimento de 16/11 já documentava 7.693 arquivos adicionais no ambiente isolado. DESKTOP_TEST_MODULES não redireciona imports ESM gerais.
- builder-util-runtime 9.7.0 e @electron/asar 3.4.1 estão no lockfile, como dependências transitivas do conjunto Electron/builder/updater. Manifesto e lockfile do checkout/preparo/deps/.41 coincidem: lock SHA-256 0bcfa9c9d88373ae7f2e86c96f3ec4f93857b05107704dd834b4a8f016e0c0a9, 313 pacotes. Nenhuma adição arbitrária ao projeto, import de produção alterado, npm install ou alteração de lockfile/node_modules ativo.
- Restaurada a árvore completa já produzida pelo npm ci anterior em C:/Dev/4PrintShops-Builds/experimental.42-20261118-45a8d3bb/source: 9.905 arquivos conferidos contra inventário SHA-256 independente preservado da .40, antes e depois da cópia. 445 arquivos de fonte/documentação/configuração copiados e conferidos; 2.212 arquivos de dependências rastreados idênticos aos preservados. Restaurar essa origem comprovada manteve o resultado reproduzível do lockfile e os binários Electron previamente verificados, sem nova resolução/download. Node 26.13.1 para compatibilidade PDF, Electron 45.1.0, builder/app-builder-lib 26.18.0, updater 6.8.11, ASAR 3.4.1, builder-util-runtime 9.7.0, pdfjs-dist 6.4.299 e pdf-lib 1.17.1.
- Reexecução dos quatro arquivos antes não carregados: **56/56**. Em seguida, suíte completa transacional: **100/100**, zero ignorados; inclui esses 56, não somar como casos únicos. INST-VAL .41→.42: **7/7** sintéticos, sucesso/quarentena/lock/catálogo preservados, rollback, interrupção entre diário e movimentação, erro inesperado recuperável, instalação limpa, reinstalação e EPERM com handle Windows real. Sem executar Setup ou instalação operacional. Harness externo adaptado para imports da cópia isolada; erro de variável no controlador foi corrigido antes de iniciar os sete casos, sem alterar testes de produto/proteções.
- Electron isolado aprovado: PIDs 16676/17360/20028/21472, quatro saídas normais código 0 em 185/182/204/157 ms, três reaberturas, zero erros. Origem app://4printshops, sandbox/contextIsolation e nodeIntegration=false. Gerenciador global, nove sugestões, CRUD/IDs, três ferramentas sincronizadas, exclusões/vazio persistentes, backups sintéticos em disco SHA-256, redefinição/restauração sintéticas e orientação/Girar arte conferidos. Evidência 4ps-size-manager-dszkmh no TEMP exclusivo. Nenhuma abertura/redefinição do perfil real.
- **141/141 funcionais anteriores reutilizados**, não reexecutados nesta recuperação: fontes funcionais permanecem iguais; mudanças do checkout restritas a versão e documentos. O filtro já aprovado permanece e será conferido novamente no pacote real. Baseline novo somente de leitura do perfil operacional: 76 arquivos, preservando registro das 21 diferenças anteriores à recuperação; não restaurado. Conferência final deve comparar esse novo baseline.
- Logs/recibos/harnesses permanecem exclusivamente na pasta externa .42, fora do Git. Checkpoint autorizado reúne somente versão Hold Test 1.0.0-experimental.42 e quatro documentos revisados; identidade/appId/produto/exe/perfil/políticas NSIS mantidos. **Build ainda não executado neste checkpoint:** após HEAD limpo, validar fonte exata, LibreOffice/PDFium/fontes/NSIS e armazenamento; só então gerar novo Setup e auditar ASAR/carga integral. Sem PDF-010/PDF-011/abas/Livreto/servidor/impressão/instalação/push.


## 18/11/2026 — Preparação .42 interrompida por dependências antes do build

**Setup .42 não gerado; não instalado.** A condição de aprovação integral não foi satisfeita. A regressão transacional no checkout ativo não carregou quatro arquivos de teste por dependências ausentes. Nenhum defeito funcional foi demonstrado por essas falhas de carregamento, mas elas impedem aprovar o preflight. Interrompidos os próximos testes e o empacotamento, conforme instrução do responsável. Não corrigidos testes/produto para contornar a falha.

### Checkpoints locais e proveniência

- 4bec43c330245e1ec5c534c8c55f125dd268a309 — Corrige filtro de empacotamento do backup de tamanhos: somente electron-builder.config.cjs e tests/package-format-backup.test.mjs. Regra negativa mantida seguida de exceção exata; FileMatcher real verifica módulo ativo incluído, legado/outros backups/privados excluídos, grafo e identidade/segurança.
- 49b425a35369be9473af24b20699a2b4013aa25a — Documenta limpeza concluida e preparacao autorizada da experimental.42: somente CONTEXTO_ATUAL/ESTADO_PROJETO/PLANEJAMENTO/PREFLIGHT_FILTRO_E_ARMAZENAMENTO_42_2026-11-17, histórico preservado. Backups/logs/perfis/temporários não incluídos.
- HEAD atual 49b425a, main 14 commits à frente da referência local origin/main; sem fetch/push. **Não existe commit-base de instalador .42.** A única alteração de metadados pendente é extraMetadata.version .41→1.0.0-experimental.42 em electron-builder.hold-native-test.cjs. Identidade/appId/produto/exe/perfil e políticas de atualização não foram alterados. Checkpoint dos metadados não criado, pois depende da aprovação integral. Raiz package.json conserva sua versão base histórica; a versão Hold Test é definida pelo extraMetadata, como nas gerações anteriores.

### Testes efetivamente executados

| Execução | Resultado | Evidência e limite |
| --- | --- | --- |
| Filtro/backup/distribuição antes do checkpoint | 15/15 | Node 26.13.1 isolado, cwd da aplicação e DESKTOP_TEST_MODULES correto; sintaxe/whitespace aprovados |
| Regressões funcionais ampliadas em 18 arquivos | 141/141, zero ignorados | audit/functional.log e functional-result.json; inclui novamente os 15 acima, não somar como testes únicos |
| Regressões de atualização em 11 arquivos | 44 testes aprovados; quatro arquivos falharam ao carregar | audit/updates.log e updates-result.json; execução total não aprovada |
| Electron/Edge e reaberturas | Não executados nesta retomada | interrompidos após falha da regressão transacional; aprovações anteriores não substituem esta etapa |
| INST-VAL .41→.42, instalação/reinstalação sintéticas | Não executado | harness externo instval42.test.mjs preparado e conferido sintaticamente, sem aprovação atribuída |
| Recursos/dependências integrais, ASAR/Setup/carga instalada | Não executados para .42 | sem novo estágio, executável, pacote ou instalação |

Os 141 testes incluem catálogo único, nove sugestões editáveis/excluíveis, exclusões/vazio persistente em unidades, sincronização do serviço, backup/redefinição e restauração **sintéticos**, orientação/Girar arte, prévias/imposição/redimensionamento, assinaturas/proteções, filtro e validação de distribuição. Não são a suíte PDF completa nem a comprovação de reinício real do Electron.

A chamada transacional passou, entre outros, nos cinco testes de persistência/EPERM do diário e no snapshot/restore de registro de laboratório. Não atribuir aprovação aos cenários transacionais de arquivo Windows em uso/quarentena/recuperação/INST-VAL .41→.42 ainda faltantes.

### Falha e próximo procedimento

desktop-update-signature.test.mjs e desktop-updates.test.mjs: ERR_MODULE_NOT_FOUND de builder-util-runtime importado por desktop/update-signature.mjs. update-template.test.mjs: MODULE_NOT_FOUND do mesmo pacote; update-transaction.test.mjs: MODULE_NOT_FOUND de @electron/asar. O checkout contém dependências históricas parciais; DESKTOP_TEST_MODULES é consumido por testes específicos, mas não altera a resolução ESM dos módulos do produto. Esses dois pacotes existem no preparo isolado source/node_modules, sem garantia nova de integridade integral nesta rodada. **Não foi feita instalação npm nem modificação do node_modules ativo.**

Retomada proposta: preparar cópia isolada de todos os fontes atuais e das dependências completas já preservadas, conferir hashes/versões e diferenças admitidas de fim de linha e somente então executar as regressões nessa cópia. Repetir os testes transacionais, executar INST-VAL .41→.42/EPERM/quarentena/recuperação e Electron com perfil novo e reinícios, validar LibreOffice/PDFium/fontes/NSIS, reconferir espaço. Nenhum executável/pacote antigo poderá substituir o build novo. Só após aprovação integral criar checkpoint de metadados, exigir HEAD/árvore limpa, exportar fonte exata e gerar/auditar o Setup em pasta inédita.

A primeira chamada reduzida desta retomada também teve erro de invocação (cwd acima da aplicação e caminho errado de dependências), 10 aprovações/duas falhas de infraestrutura; foi corrigida e repetida com 15/15. A falha posterior do grupo transacional, registrada acima, é independente e mantém a geração bloqueada. Não silenciar ou omitir essas ocorrências.

### Preservação e armazenamento

- .41 instalada confirmada internamente como 1.0.0-experimental.41; **19.515 arquivos estáticos iguais aos hashes preservados**, zero divergências. debug.log é dinâmico, fora da comparação estática. Setup .41 preservado com SHA-256 253D672DDFB8C02346D540E767450A4D800395309AAB0F5655A55F9BDE7D16D3. Nove transações operacionais históricas em committed/rolled-back, nenhuma pendente; nenhum processo Electron/4PrintShops/Office/Setup identificado no encerramento.
- O perfil foi somente lido para hashes/metadados, sem abrir aplicativo instalado, restaurar, escrever ou redefinir dados. Comparação com o snapshot anterior à limpeza (não com um snapshot novo pré-teste) encontrou **21 diferenças entre 76 arquivos**, em caches/logs/armazenamento Chromium. Maior data de gravação observada: 18/11/2026 10:54:50 no fuso de São Paulo, anterior aos testes atuais das 11:42. Esses metadados não demonstram causa nem conteúdo alterado do catálogo. Portanto não afirmar igualdade byte a byte do perfil com o snapshot da limpeza nem atribuir as diferenças aos testes desta retomada. Evidências privadas: preservation.json e profile-differences-metadata.json; sem conteúdo do perfil nos documentos versionados.
- Livre no encerramento da conferência: **21720371200 bytes (20.229 GiB)**. Estimativa histórica de 9–12 GiB adicionais para fonte nova/build/auditoria/temporários ainda deve ser reconferida após preparo isolado. Não houve cópia integral de build, extração de carga ou novas exclusões; não declarar aprovação de capacidade para um build ainda não preparado.
- Pasta nova exclusiva de evidências: C:/Dev/4PrintShops-Builds/experimental.42-20261118-45a8d3bb; contém somente audit/harnesses/relatório, sem source/dist/Setup. Logs privados permanecem fora do Git. Setups .29/.36 continuam sem recuperação confirmada, como já registrado. Sem alterações de PDFs reais, modelos operacionais, F4/F6/F8, impressão, instalação/desinstalação, push ou novas funcionalidades.

### Situação final para revisão

Sem Setup .42: caminho, tamanho, SHA-256, auditoria ASAR e integridade da carga do novo instalador **não existem nesta etapa**. Git mantém os dois novos checkpoints e cinco arquivos pendentes: metadado Hold Test e quatro documentos desta ocorrência. Nenhum checkpoint adicional criado automaticamente após a interrupção. Próximo passo é revisar este diagnóstico e retomar a preparação isolada/validações pendentes; instalar .42 continua não autorizado.


## 18/11/2026 — Limpeza aprovada e retomada de preparação autorizada

Responsável aprovou os oito resultados e autorizou checkpoints separados, preparação da versão 1.0.0-experimental.42 e Setup em pasta inédita exclusivamente após testes, preflight e armazenamento inteiramente aprovados. **Instalar/desinstalar não está autorizado.** Preservar .41, perfil/catálogo reais e Setups; sem push/impressão/reset real/PDF-010/PDF-011/abas/servidor/Livreto.

Diff inicial integral revisado (quatro modificados/dois novos); filtro/teste registrados em `4bec43c`. Nova execução válida: **15/15** em package-format-backup.test.mjs, pdf-format-backup.test.mjs e distribution-validation.test.mjs, serial, Node 26.13.1 isolado. Inclusão exata do módulo ativo e exclusão do legado confirmadas. Sintaxe/whitespace aprovados.

A primeira invocação nesta retomada foi inválida por cwd acima da aplicação e variável de dependências apontando a caminho inexistente; ENOENT/MODULE_NOT_FOUND, 10 aprovações e duas falhas de infraestrutura. Corrigidos somente os parâmetros de execução; repetição completa 15/15 aprovada, sem modificar código/testes. Não houve tentativa de empacotamento durante essa ocorrência. Regressões funcionais, Electron, INST-VAL, integridade de recursos/dependências, espaço e pacote permanecem pendentes neste checkpoint. Histórico anterior mantido abaixo.

## 18/11/2026 — Limpeza autorizada concluída no checkout correto

A primeira operação desta retomada foi Get-Location sem parâmetros e sem workdir: C:\Dev\4PrintShops. O impedimento anterior de diretório padrão foi resolvido. Responsável reafirmou a autorização dos sete alvos originais mais a auxiliar; exclusão concluída exclusivamente nesses oito caminhos, listados individualmente abaixo. Nenhum Set-Location/CLI do VS Code foi utilizado para contornar o contexto.

Antes da remoção: backup ZIP no servidor relido com SHA-256 D874DAE29E1673A6C04C47C2CCE566E684C5F7F9426FFF9CF4EE371B3064231A; 86.087 registros de origem comparados por tamanho/hash, aproveitando a identidade física dos hardlinks para evitar leituras repetidas. Manifesto local/auxiliar iguais. Bundles ativo/antigo verificados pelo Git e SHA-256; patch/teste/Setup .41 no servidor e ZIP dos PDFs reais também relidos. Seis pendências atuais copiadas e verificadas em Alteracoes_ativas-COb5aw no complemento privado do servidor. Seis intermediários mantinham inventários. Projeto antigo continha uma referência técnica adicional de 41 bytes para a árvore já preservada, sem código exclusivo. Nenhuma mudança funcional nesse checkout antigo.

**Interrupção e correção registradas:** a primeira tentativa removeu o conteúdo do projeto antigo, mas o Windows bloqueou a raiz vazia por diretório corrente de processos. A sondagem inicial de abertura exclusiva e a inspeção de linhas de comando não detectaram esse vínculo. Os outros sete alvos não foram removidos naquela tentativa. Inspeção específica dos diretórios correntes identificou três auxiliares ociosos das ferramentas do próprio Codex, iniciados às 08:47:41 (PIDs 3576, 18376, 11752). Somente esses auxiliares foram encerrados, após confirmação de identidade/parentesco; IDE e aplicativos do usuário permaneceram abertos. Nova leitura dos diretórios correntes não encontrou alvos em uso. Sondagem exclusiva com direito de exclusão, inventários restantes, ancestrais e todos os descendentes sem reparse points foram reconferidos antes da retomada. A tentativa inicial e o resultado final estão preservados separadamente na auditoria privada. O diagnóstico inicial excessivamente lento foi interrompido e substituído pela sondagem sem leitura de conteúdo; a primeira comparação de hashes, interrompida por concorrência com esse diagnóstico, foi repetida integralmente e aprovada.

| Caminho autorizado exato | Resultado final | Variação observada no intervalo (GiB) |
| --- | --- | ---: |
| `C:\Users\SAGRADA02\Documents\scripts_Marcos\automacaoAdesivosCoreldraw` | Removido | 0,291 |
| `C:\Dev\4PrintShops-Builds\experimental.40-2026-11-16\source\dist-desktop\hold-native-test\win-unpacked` | Removido | 1,634 |
| `C:\Dev\4PrintShops-Builds\experimental.40-2026-11-16\source\node_modules` | Removido | 0,443 |
| `C:\Dev\4PrintShops-Builds\experimental.40-2026-11-16\source\build\libreoffice-cache` | Removido | 0,960 |
| `C:\Dev\4PrintShops-Builds\experimental.40-EPERM-20261116-164725\source\dist-desktop\hold-native-test\win-unpacked` | Removido | 1,649 |
| `C:\Dev\4PrintShops-Builds\experimental.40-EPERM-20261116-164725\source\node_modules` | Removido | 0,535 |
| `C:\Dev\4PrintShops-Builds\experimental.40-EPERM-20261116-164725\source\build\libreoffice-cache` | Removido | 0,960 |
| `C:\Dev\4PrintShops_Auditoria_Armazenamento42_20261117-jQkeqj\conteudo-unico-20261118-WrI4mT` | Removido | 3,553 |

As variações da tabela são medições do volume durante cada remoção, incluindo a tentativa inicial no primeiro alvo; não são tamanhos exclusivos atribuíveis sem interferência de outros processos. A auxiliar continha hardlinks dos mesmos dados e não representa ganho adicional ao volume lógico previsto. Os sete alvos originais somavam 10.439.353.394 bytes nesta retomada (incluindo a referência técnica adicional).

Espaço livre antes da primeira exclusão: **11561639936 bytes (10,768 GiB)**. Ao terminar: **22101618688 bytes (20,584 GiB)**. Aumento líquido observado: **10539978752 bytes (9,816 GiB)**. A soma das variações dentro dos intervalos de exclusão foi 10,026 GiB; atividade do volume entre os intervalos explica a diferença em relação ao saldo líquido. Conferência posterior: **22115475456 bytes (20,597 GiB)** livres. Medição efetiva, sem somar hardlinks duas vezes.

Pós-limpeza aprovado, antes destes registros documentais: oito caminhos ausentes; **2.660 arquivos do checkout iguais por SHA-256**, **19.515 arquivos estáticos instalados da .41 iguais ao inventário**, **76 arquivos do perfil iguais por SHA-256 ao snapshot imediatamente anterior**. debug.log permanece dinâmico e fora da comparação estática. Setup .41 local relido com hash esperado. PDFs reais, Setups preservados, preparo isolado/recursos para builds e restante da auditoria estão fora dos caminhos removidos. Git fsck --full aprovado (somente árvores históricas sem referência, sem corrupção), diff --check aprovado; main/HEAD bccce139eda34a396325baf854de3675ba876890 e os mesmos seis caminhos pendentes, 12 commits à frente da referência local origin/main. Não foi feita consulta nova ao remoto, commit ou push. Em seguida, somente os quatro documentos de continuidade receberam este checkpoint.

Evidências privadas, fora do Git: `C:/Dev/4PrintShops_Auditoria_Armazenamento42_20261117-jQkeqj`. Recibos principais: resultado-limpeza-autorizada-20261118.json (oito resultados e tentativa inicial), resultado-limpeza-primeira-tentativa-20261118.json, preflight-limpeza-atual-20261118.json, uso-e-vinculos-limpeza-20261118.json, processos-cwd-alvos-20261118.json, perfil-antes-limpeza-20261118.json e conferencia-pos-limpeza-20261118.json. As fontes antigas removidas continuam recuperáveis pelos bundles e ZIP/manifesto verificados no servidor.

**Experimental.42:** impedimento de armazenamento resolvido para revisão, com cerca de 20,60 GiB livres e margem de 20 GiB anteriormente solicitada atingida na medição. Filtro aprovado e 15/15 testes são resultados anteriores, não reexecutados nesta limpeza. Metadados permanecem .41. Retomada de metadados/preflight funcional/regressões/Electron/INST-VAL/auditoria do pacote continua pendente de revisão; reavaliar espaço e temporários ao retomar. Setups .29/.36 continuam indisponíveis, sem promessa de reconstrução exata. Nenhum build, Setup .42, instalação, impressão, redefinição operacional, alteração de F4/F6/F8, commit ou push. **Aguardar revisão do responsável antes de continuar o empacotamento.**


## 18/11/2026 — Recuperação independente da execução interrompida

Nova execução somente de diagnóstico, com diretório corrente explícito `C:\Dev\4PrintShops`; não retomadas sessões/comandos anteriores e nenhuma exclusão iniciada. Na execução anterior, comandos de troca do workspace pelo CLI do VS Code receberam resultado `aborted`; não há evidência de que um comando de exclusão tenha sido enviado. A relação causal com a interrupção da extensão é hipótese, não falha confirmada do sistema de arquivos.

Conferência inicial: oito caminhos presentes; seis intermediários com quantidades/tamanhos iguais aos inventários anteriores. Pasta antiga contém dois arquivos a mais, total adicional de 82 bytes, a identificar antes de qualquer exclusão. Auxiliar: 25.730 arquivos com hardlinks e um manifesto, sem junctions/links simbólicos detectados. ZIP do servidor com SHA-256 original, manifesto interno igual ao local, ambos os bundles verificados e seis arquivos pendentes iguais às cópias privadas. Nenhum processo de limpeza/Setup/Electron/LibreOffice identificado.

A conferência contra o inventário da primeira instalação .41 encontrou crescimento de `debug.log` e interrompeu a comparação estrita (`Installed file size mismatch debug.log`), sem alteração de arquivos. Esse erro de conferência foi preservado e investigado: o log passou de 332 para 664 bytes com mensagens Chromium `TransactNamedPipe`, enquanto **19.515 outros arquivos instalados corresponderam integralmente aos hashes anteriores**. Versão interna/registro únicos em .41; nenhum aplicativo foi iniciado. Os 76 arquivos atuais do perfil foram lidos com sucesso, sem modificação, mas não se afirma igualdade byte a byte com o backup anterior à primeira abertura do aplicativo. Nenhuma operação em F4/F6/F8.

Todos os arquivos versionados e arquivos locais arquivados do projeto antigo permanecem presentes, e os arquivos locais mantêm os tamanhos. As referências técnicas adicionais apontam à mesma árvore `5ff9719ed1328fb1914b9a0e6dda8b1c04f4b2fb` já preservada: nenhum código exclusivo identificado nesses 82 bytes adicionais. Nenhum sinal de exclusão parcial nos seis intermediários ou na auxiliar. Diários novamente normalizados por caminho: nove transações históricas, todas committed/rolled-back, nenhuma pendente. Sem processo anterior de limpeza/Setup/LibreOffice identificado.

**Impedimento de sessão confirmado:** apesar dos comandos de trabalho usarem explicitamente `C:\Dev\4PrintShops`, uma chamada nova de `Get-Location` sem `workdir` retornou `C:\Users\SAGRADA02\Documents\scripts_Marcos\automacaoAdesivosCoreldraw`. Portanto, não confirmar que o workspace padrão desta execução mudou. Não repetir o CLI de troca que interrompeu a execução anterior e não excluir a pasta que ainda serve de diretório padrão. A autorização de limpeza dos oito caminhos permanece válida; é necessário que a execução esteja efetivamente vinculada ao checkout ativo antes de realizá-la.

Livre observado ao final: cerca de **10,81 GiB**; nada foi excluído. Relatórios privados `recuperacao-estado-20261118.json`, `recuperacao-windows-20261118.json` e `recuperacao-conferencia-final-20261118.json` na auditoria de armazenamento. Sem geração .42, metadados novos, build, instalação, commit ou push. Resultados e diagnóstico persistentes neste arquivo para recuperar a tarefa sem retomar sessões anteriores.

## 18/11/2026 — Preparação final autorizada, sem exclusão

O responsável aprovou o filtro (15/15 já executados) e o plano de 9,72 GiB, condicionando qualquer exclusão à conferência final dos caminhos. Setups .29/.36 continuam sem cópias recuperáveis confirmadas; não presumir reconstrução exata.

O bundle ativo no servidor foi novamente verificado pelo Git: histórico completo, main em `bccce139eda34a396325baf854de3675ba876890`, SHA-256 `7163211F5A5C69F57231E06E2811B0219762BA417A97D83EAEC1A985A340816E`. Os 12 commits locais continuam preservados. PDFs reais e Setup .41 já verificados não foram duplicados.

Preparação adicional privada concluída em `\\OFICINA\Arquivos\000AAA 4PrintShops\Preservacao_SAGRADA02_20261117-iAwI6m\Complemento_20261118-PI2bfC`: arquivos ignorados do projeto antigo, configuração Git local e conteúdos dos seis intermediários. Formato: ZIP com objetos identificados por SHA-256 e manifesto de caminhos originais, tamanhos e grupos; conteúdos idênticos têm um único objeto.

`arquivos-locais-e-intermediarios.zip`: **2.004.989.317 bytes**, SHA-256 **D874DAE29E1673A6C04C47C2CCE566E684C5F7F9426FFF9CF4EE371B3064231A**. Leitura integral do ZIP no servidor verificou os 25.730 objetos e o manifesto, cobrindo 86.087 registros originais, sem divergências, ausências ou entradas duplicadas/inesperadas. Origens conferidas antes/depois da gravação e inventários dos sete alvos reconferidos. Recibo `recibo-verificado.json` e instruções `LEIAME-RECUPERACAO.txt` no mesmo diretório privado. A execução do arquivo .ps1 de verificação foi recusada pela política local; a mesma verificação foi executada como comandos inline, sem mudar a política do Windows.

Os sete alvos somam exatamente **10.439.353.353 bytes (9,722 GiB)**. Medição durante a conclusão: aproximadamente **10,97 GiB livres**, projetando **20,69 GiB**, sujeitos à atividade normal do disco e à liberação dos vínculos auxiliares. Recibos/inventário final em `plano-final-exclusao.json` na auditoria privada. A cópia separada dos seis arquivos pendentes usa o subdiretório `Alteracoes_ativas-*`, com recibo e SHA-256, sem incluir perfis ou PDFs.

Pasta auxiliar local do arquivamento: `C:\Dev\4PrintShops_Auditoria_Armazenamento42_20261117-jQkeqj\conteudo-unico-20261118-WrI4mT`. Contém vínculos físicos temporários, sem duplicar os bytes dos arquivos. **Também precisa constar da autorização final de remoção**, depois de confirmado o ZIP: enquanto esses vínculos existirem, a exclusão das origens não liberará todo o espaço previsto. Sua remoção não acrescenta outros 9,72 GiB ao total; são os mesmos dados. Não remover o restante da auditoria. Nenhuma pasta, inclusive esta auxiliar, foi excluída.

O projeto antigo foi novamente medido em 3.314.876.702 bytes (3,087 GiB), 35.922 arquivos, sem redirecionamento para o checkout ativo. Esta sessão/IDE ainda parte da pasta antiga: antes da futura exclusão, abrir o workspace `C:\Dev\4PrintShops` e deixar de usar a pasta antiga como diretório corrente.

Os sete alvos originais e dimensões permanecem na tabela abaixo. A correção do filtro e documentos pendentes serão preservados em cópia privada adicional. .41 instalada, perfis, catálogo operacional, recursos isolados, Setups e evidências fora dos alvos permanecem intocados. Não foram executados testes funcionais novos, build, instalação, commit, push ou redefinição operacional nesta etapa.

## Correção e verificações realizadas

Base `bccce139eda34a396325baf854de3675ba876890`, main com 12 commits locais ainda não enviados. Sem alteração funcional ou de versão: metadados continuam .41.

A regra `!public/js/tools/*backup*` excluía também o módulo ativo `pdf-format-backup.js`. Mantida a exclusão e acrescentada, imediatamente depois, somente a inclusão exata `public/js/tools/pdf-format-backup.js`. O legado `pdfs backup.js`, outros backups, testes e dados privados continuam excluídos.

Novo `tests/package-format-backup.test.mjs` usa o FileMatcher real do preparo isolado, nas configurações base e Hold Test. Confere inclusão, negativos, os quatro módulos do grafo do catálogo e identidade/configurações de segurança. O módulo de backup não requer pacote npm adicional.

Execução realizada com Node 26 isolado, DESKTOP_TEST_MODULES apontando ao preparo desktop, `--test --test-concurrency=1`: `package-format-backup.test.mjs`, `pdf-format-backup.test.mjs`, `distribution-validation.test.mjs`: **15/15 aprovados**. Antes da correção, os testes específicos demonstraram duas falhas de inclusão. Sintaxe da configuração e do novo teste, e `git diff --check`, aprovados. Não executados build, Electron, INST-VAL .42 ou suíte PDF completa nesta etapa.

## Política aprovada e preservação

- .29 em produção, .36 no beta tester e .41 no desenvolvimento. Validar novas versões primeiro no desenvolvimento.
- Exclusão futura do projeto redundante e intermediários foi autorizada condicionalmente, exigindo apresentação dos caminhos e itens únicos para última conferência. **Nada foi excluído ou movido.**
- C: é o único volume local montado, com aproximadamente **11,28 GiB livres** no último levantamento. Servidor acessível com cerca de 474 GiB disponíveis antes das cópias de preservação. Não pressupor que build em SMB é validado.
- 31/31 conteúdos dos casos reais conferidos por SHA-256 no servidor: 001–029 dentro do ZIP de migração e 030/031 nos arquivos adicionais do servidor. ZIP conferido: `E73183D9761BFE5424ADEFBCAE81C3933EEF42B5FF6CD8FA0176B01F75966843`. Sem extração ou alteração dos originais.
- Preservados e relidos por SHA-256 no servidor: bundles Git ativo/antigo, patch pendente, novo teste e Setup .41. Bundles verificados pelo Git; o ativo preserva os 12 commits locais. Setup .41 mantém SHA `253D672DDFB8C02346D540E767450A4D800395309AAB0F5655A55F9BDE7D16D3`.
- Instaladores .29 e .36 **não localizados** nas buscas limitadas realizadas. O responsável informou que provavelmente foram excluídos anteriormente. A preservação dessas versões permanece pendente; não declarar recuperabilidade a partir das instalações existentes nem gerar substitutos.

Recibos privados: `C:\Dev\4PrintShops_Auditoria_Armazenamento42_20261117-jQkeqj`. Preservação no servidor: `\\OFICINA\Arquivos\000AAA 4PrintShops\Preservacao_SAGRADA02_20261117-iAwI6m`. Evidências privadas não devem entrar no Git.

## Diretórios candidatos — não executar exclusão ainda

Plano revisto: sete alvos, **9,72 GiB lógicos**, elevando a disponibilidade estimada para **21,0 GiB**. Não somar este plano à proposta anterior de oito intermediários: existem sobreposições. Verificar o espaço efetivamente recuperado após eventual execução autorizada.

| Caminho completo | GiB | Condição anterior à remoção |
| --- | ---: | --- |
| `C:\Users\SAGRADA02\Documents\scripts_Marcos\automacaoAdesivosCoreldraw` | 3,087 | Arquivar dados ignorados/configurações locais; trocar a sessão de trabalho para o checkout ativo |
| `C:\Dev\4PrintShops-Builds\experimental.40-2026-11-16\source\dist-desktop\hold-native-test\win-unpacked` | 1,839 | Preservar geração histórica no servidor e verificar leitura/hashes |
| `C:\Dev\4PrintShops-Builds\experimental.40-2026-11-16\source\node_modules` | 0,518 | Mesma condição |
| `C:\Dev\4PrintShops-Builds\experimental.40-2026-11-16\source\build\libreoffice-cache` | 0,960 | Mesma condição |
| `C:\Dev\4PrintShops-Builds\experimental.40-EPERM-20261116-164725\source\dist-desktop\hold-native-test\win-unpacked` | 1,839 | Preservar geração histórica no servidor e verificar leitura/hashes |
| `C:\Dev\4PrintShops-Builds\experimental.40-EPERM-20261116-164725\source\node_modules` | 0,518 | Mesma condição |
| `C:\Dev\4PrintShops-Builds\experimental.40-EPERM-20261116-164725\source\build\libreoffice-cache` | 0,960 | Mesma condição |

O projeto antigo é distinto de `C:\Dev\4PrintShops`, sem links de redirecionamento. HEAD antigo `118032e`, árvore limpa, sem commits exclusivos da branch e sem arquivos não rastreados fora das exclusões. As referências técnicas de turn-diffs são árvores iguais à árvore desse HEAD. Não contém a biblioteca de casos reais.

Há **27.241 arquivos ignorados (2,934 GiB)** em build/node_modules: não recuperáveis somente por checkout. Preservar antes de remover, incluindo `build/distribution/identity.json`, `build/hold-native/manifest.json`, `build/native/manifest.json`, `build/print-lab/manifest.json`, `build/libreoffice/trust.json`, `build/libreoffice/extraction.log` e três arquivos gerados NSIS. Não foram encontrados hooks personalizados nem configurações Git locais de credenciais. O arquivo desses diretórios e das gerações ainda é uma pendência, não um backup já concluído.

## Outros volumes e grupos medidos

| Geração em C:\Dev\4PrintShops-Builds | GiB |
| --- | ---: |
| experimental.40-2026-11-16 | 6,191 |
| experimental.40-EPERM-20261116-164725 | 5,445 |
| experimental.40-INSTVAL-20261116-1794868567608 | 8,019 |
| experimental.41-FINAL-20261117-dE9Y3n | 8,018 |
| experimental.41-INSTVAL-20261117-CqiKQD | 3,118 |

Preservar preparo isolado em `C:\Dev\4PrintShops-Tools` (3,872 GiB), backups privados (0,515 GiB), auditorias externas (0,297 GiB), restauração verificada, ZIP de migração, biblioteca ativa e todos os Setups atuais. Os intermediários acima são regeneráveis, mas os relatórios, manifestos, fontes históricas e Setups devem continuar disponíveis. Não remover genericamente `build`, `audit` ou pastas temporárias.

Temporários selecionados somam cerca de 12,30 GiB, mas incluem evidências e estágio de instalações: `4ps-setup36-graph-pod2lI` (4,158), `nsrE63A.tmp` (2,430), `4ps-full-upgrade-PLtmLQ` (2,329) e `4ps-setup38-RzNGyz` (1,313). Não incluídos na proposta de remoção sem revisão específica.

O builder usa TEMP/TMP do usuário (`C:\Users\SAGRADA02\AppData\Local\Temp`), com `APP_BUILDER_TMP_DIR` atualmente ausente; cache isolado em `C:\Dev\4PrintShops-Tools\desktop-prep-2026-11-16\builder-cache`. Em outro volume local NTFS, direcionar fonte, saída e temporários por processo, sem mudar permanentemente o ambiente. Não há outro volume local disponível agora.

Próximo passo: última conferência do responsável, resolver ou aceitar explicitamente a indisponibilidade dos Setups .29/.36, arquivar e verificar os dados locais identificados e então executar somente a limpeza aprovada. Reavaliar margem de pelo menos 20 GiB para geração e temporários, considerando geração anterior de cerca de 8 GiB. Retomar metadados/preflight/testes .42 após armazenamento resolvido; **sem autorização atual para gerar Setup**. Instalação .41, perfis, F4/F6/F8 e arquivos de produção intactos.
