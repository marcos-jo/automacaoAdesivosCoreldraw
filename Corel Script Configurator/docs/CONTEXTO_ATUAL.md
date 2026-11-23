# 4PrintShops — Contexto Atual e Passagem de Trabalho

## 23/11/2026 — Checkpoint após MIG-005B2 e decisões atuais

Operador autorizou um único commit local de `scripts/dev-environment.ps1` e deste contexto, com a mensagem `chore(dev): registra validacoes e continuidade da migracao SSD`. Base inicial conferida: workspace/diretório efetivo/raiz Git `D:\Dev\4PrintShops`, branch `main`, HEAD `19a5930a092b669c534922c1d13fd137a0b996c7`, somente esses dois arquivos pendentes e stage vazio. Push não autorizado nesta etapa. Marcos MIG-004B2B2 (parâmetro `-SourcePath`, diagnósticos positivos/negativos e isolamento), MIG-004B2B3 (modo restrito `MigrationChecks`, 12 testes unitários aprovados) e MIG-005B2 (seis arquivos de instaladores copiados/verificados, sem liberação operacional) já detalhados abaixo; não houve nova execução desses diagnósticos/testes neste checkpoint.

### MIG-005A e MIG-005B1 — Preservação e destino conhecido

- MIG-005A inventariou o conjunto a preservar: Git e trabalho local não versionado, builds .40/.41/.42, Setups e integridade, ferramentas/dependências, auditorias e recursos externos de recuperação. Estimativa bruta de 42,981 GiB com fonte de validação/evidências e backups externos, ou 56,682 GiB incluindo transações de recuperação após revisão. Não é uma lista autorizada para cópia automática: casos reais e demais materiais privados precisam de triagem, sem copiar documentos de clientes. GitHub não substitui esses recursos locais.
- MIG-005B1 confirmou por leitura `\\oficina\ARQUIVOS\4PrintShops`, pasta comum acessível à sessão. Na consulta de 22/11, aproximadamente 471,33 GiB disponíveis; não é nova medição deste checkpoint. Existiam `arquivos de teste`, `instaladores` então vazio e um ZIP antigo de migração, preservados. ACL observável herdava FullControl para Todos e modificação para usuários autenticados; permissões efetivas combinadas do compartilhamento não foram integralmente verificadas. Portanto, não foi comprovada proteção adequada para backups privados; instaladores e backups privados exigem tratamentos distintos. Nenhum acesso ao servidor realizado neste checkpoint.

### Decisões atuais do operador

- Configuração definitiva de permissões e backups privados no servidor **adiada pelo operador**. Isso não aprova transferência de material privado para o compartilhamento atual nem libera os instaladores para distribuição operacional.
- Objetivo prioritário permanece **liberar espaço no C: do computador de desenvolvimento**. A limpeza de `C:\Dev` continua pendente e não autorizada; antes de excluir originais, é necessário preservar uma cópia **independente e verificada** dos recursos críticos, definir o escopo exato e obter autorização específica. As cópias de instaladores no servidor não equivalem a backup integral do desenvolvimento/recuperação.
- Os 12 testes unitários aprovados em D: não constituem homologação completa da Experimental.42, do Electron, do empacotamento ou da recuperação operacional. Nenhum build, instalador ou teste Electron foi executado nas validações MIG-004B2B2/MIG-004B2B3; esta revisão não inicia novas implementações, testes ou limpeza. Histórico anterior preservado integralmente.

## 23/11/2026 — MIG-005B2: instaladores .41/.42 preservados no servidor; distribuição não autorizada

- Abertura conferida: diretório efetivo e raiz Git `D:\Dev\4PrintShops`, `main`, HEAD `19a5930a092b669c534922c1d13fd137a0b996c7`; alterações locais do iniciador e deste contexto preservadas. Operador autorizou exclusivamente cópia dos Setups homologados e arquivos de integridade, sem distribuição ou instalação em outros computadores.
- Origens em D:\Dev\4PrintShops-Builds: `experimental.41-FINAL-20261117-dE9Y3n\source\dist-desktop\hold-native-test` e `experimental.42-20261118-45a8d3bb\source\dist-desktop\hold-native-test`. Setups conferidos em tamanho/SHA-256 contra os registros de homologação preservados na MIG-001; sidecars .sha256 conferidos contra os respectivos nomes/hashes e blockmaps originais presentes.
- Destino exclusivo: `\\oficina\ARQUIVOS\4PrintShops\instaladores\Experimentais\Experimental.41` e `Experimental.42`. Criadas somente `Experimentais` e essas duas subpastas, antes inexistentes. Copiados os seis arquivos: `4PrintShops-Hold-Test-1.0.0-experimental.41-Windows-x64-Setup.exe` e equivalentes `.exe.sha256`/`.exe.blockmap`; mesmo conjunto para `.42`. Nenhum conflito de nome encontrado; cópia com falha se destino já existir, sem sobrescrita, movimentação ou exclusão.
- Setup .41: **515.071.478 bytes**, SHA-256 **253D672DDFB8C02346D540E767450A4D800395309AAB0F5655A55F9BDE7D16D3**. Setup .42: **515.070.952 bytes**, SHA-256 **7AE0D35B919714481E8E64AA0E304038DAD0A6548E1A1944C974CF0ACBDEBC1C**. Sidecars: 132 bytes cada; blockmaps: .41 **535.291 bytes**, .42 **535.327 bytes**.
- SHA-256 dos complementos copiados: .41 `.exe.sha256` **60AE7247A4DB795BB6AE3C05F0C36AB35A7A8FEA7B0877DACAD66790CE10524F**, `.exe.blockmap` **CE80863665A428890B8CF964F25A6407BD29E9C663954ABC514C6B06084BF190**; .42 `.exe.sha256` **77F03337E4ABFDD52AB1E6AF877C84C16152313240ED86D3B4DDB6186C3C5413**, `.exe.blockmap` **C5E037E3FB3462D8ED68C7A27CFB4E9871B9F10CABE9AFF2C418C1CB61DE775D**.
- **6/6 arquivos relidos no servidor, tamanhos e SHA-256 iguais às origens confiáveis; zero falhas de rede/integridade observadas.** Total lógico **1.031.213.312 bytes (aproximadamente 0,960 GiB)**. Originais em D: relidos após a cópia, com tamanhos, datas de modificação e hashes preservados. Nenhum arquivo preexistente substituído; `arquivos de teste` e ZIP de backup existente não alterados.
- **Cópia verificada não é autorização de distribuição:** os instaladores no servidor NÃO estão liberados para distribuição operacional. ACLs/segurança efetiva do compartilhamento ainda exigem revisão; permissões amplas observadas na MIG-005B1 não foram alteradas. Não criar cópias em outros computadores nem executar Setup/atualização por este registro.
- Nenhuma cópia de fonte, perfil, PDF real ou backup privado; nenhum build, alteração de C:\Dev, limpeza, instalação, commit ou push. Próximas etapas dependem de autorização: definir/revisar proteção de leitura e gravação dos instaladores antes de eventual liberação e confirmar destino efetivamente restrito para backups privados/recuperação, com triagem e plano de integridade da MIG-005A. Não iniciar essas etapas automaticamente.

## 22/11/2026 — MIG-004B2B3: primeira validação técnica Node em D: aprovada

- Abertura conferida em `D:\Dev\4PrintShops`, workspace/raiz Git corretos, `main`, HEAD `19a5930a092b669c534922c1d13fd137a0b996c7`. Alterações pendentes do contexto e iniciador preservadas. Sem commit ou push nesta etapa.
- Inspecionados integralmente três testes da fonte independente e seu grafo de dependências: `pdf-formats.test.mjs` (5 casos), `pdf-format-catalog.test.mjs` (5) e `pdf-sheet-orientation.test.mjs` (2), com `pdf-formats.js`, `pdf-format-storage.js`, `pdf-format-backup.js` e `pdf-sheet-orientation.js`. Usam fixtures sintéticas/armazenamento em memória, sem imports de filesystem, processos, rede ou Electron. Não acessam dados ou perfis operacionais.
- Acrescentado ao iniciador o modo específico `MigrationChecks`, limitado a esses três arquivos, exigindo `-SourcePath` explícito. O modo padrão permanece `Diagnostic`; não há controlador genérico de comandos/builds. Node/cache/temporários reutilizam o mesmo ambiente de processo filho. NODE_COMPILE_CACHE/NODE_V8_COVERAGE são removidos somente do filho e NODE_DISABLE_COMPILE_CACHE=1 evita cache de compilação herdado. Proteções: não iniciar com menos de 10 GiB livres na unidade do Windows; após os testes, queda de pelo menos 256 MiB ou valor abaixo de 10 GiB interrompe novas execuções. Timeout de 60s e nenhuma repetição automática.
- Comando do iniciador: `powershell.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy RemoteSigned -File "D:\Dev\4PrintShops\Corel Script Configurator\scripts\dev-environment.ps1" -Mode MigrationChecks -SourcePath "D:\Dev\4PrintShops-Builds\MIG-004B2B-validacao-20261119-g5NpOU\source"`. Somente política do novo processo; nenhuma configuração persistente alterada.
- Comando Node efetivo: `D:\Dev\4PrintShops-Tools\node-v26.13.1-win-x64\node.exe --test --test-concurrency=1 --test-reporter=tap`, seguido dos caminhos absolutos dos três testes acima; cwd na fonte independente selecionada. Node v26.13.1 confirmado pelo diagnóstico. **12/12 aprovados, zero falhas/cancelados/ignorados, saída 0, stderr vazio**, duração informada pelo runner 320,0552 ms. Nenhuma falha técnica nesta rodada.
- Evidências em `D:\Temp\4PrintShops\dev-aa9918286284420db61f2a8ed5030ea0\results`: `stdout.tap`, `stderr.txt`, `execution.json`, com executável, argumentos, cwd, hashes dos testes, tempos, destinos e medições de espaço. TEMP/TMP exclusivos nessa execução; APP_BUILDER_TMP_DIR em subpasta e caches em D:\Cache\4PrintShops. Não houve uso dos caminhos antigos C:\Dev; referências textuais encontradas pelo diagnóstico foram apenas relatadas.
- Monitoramento externo da execução completa: C: **12.109.729.792 → 12.109.725.696 bytes** (redução observada de 4 KiB, cerca de 11,278 GiB livres); D: **72.601.726.976 → 72.601.718.784 bytes** (8 KiB, cerca de 67,616 GiB). Dentro do processo dos testes, C: permaneceu igual; controle de espaço não acionado. Variações observadas não constituem atribuição de causa exclusiva ao teste.
- Sonda do filho aprovada; snapshots externos Process/User/Machine e políticas iguais antes/depois. SHA-256 dos sete arquivos inspecionados permaneceu igual. Fonte histórica, instaladores, perfis operacionais, scripts históricos e configurações globais não foram modificados. Sem Electron, navegador, Office, compilador, build, instalação, download, limpeza ou acesso ao servidor.
- **Limites e pendências:** validação unitária pequena em memória; não homologa integralmente a Experimental.42, execução gráfica, LibreOffice/PDFium, empacotamento ou recuperação operacional. Há base para propor a preservação independente no servidor em etapa própria, com destino/autorização e conferência integral; servidor não acessado. Liberação de C: ainda não autorizada e depende de backup verificado, revisão dos caminhos/baselines pendentes e escopo explícito de limpeza. Não iniciar próximos testes ou etapas automaticamente.

## 22/11/2026 — MIG-004B2B2: seleção explícita da fonte e diagnóstico aprovados

- Ambiente inicial conferido: diretório efetivo, workspace e raiz Git `D:\Dev\4PrintShops`, `main`, HEAD `19a5930a092b669c534922c1d13fd137a0b996c7`; somente a atualização documental da MIG-004B2B1 pendente, preservada integralmente.
- `scripts/dev-environment.ps1` agora aceita `-SourcePath` com caminho completo, independente do nome da pasta. Resolve o caminho literal de filesystem, rejeita reparse points na ancestralidade, diretórios inexistentes, manifestos inválidos/ausentes e dependências principais ausentes ou divergentes do lockfile. Seleção explícita inválida encerra sem buscar outra fonte. Sem o parâmetro, mantém a seleção histórica por candidato único `experimental.42-*`; modo continua exclusivamente `Diagnostic`.
- Fonte explícita diagnosticada: `D:\Dev\4PrintShops-Builds\MIG-004B2B-validacao-20261119-g5NpOU\source`. Electron 45.1.0, builder/app-builder-lib 26.18.0, updater 6.8.11, ASAR 3.4.1, builder-util-runtime 9.7.0, PDF.js 6.4.299 e pdf-lib 1.17.1 conferidos por manifestos contra lockfile; presentes 9.905 arquivos de dependências e recursos LibreOffice/PDFium/auxiliares nativos. Não foi repetida a auditoria integral de hashes de dependências desta cópia.
- Validações leves: parser PowerShell aprovado; fonte explícita válida e modo padrão aprovados com saída 0; caminho inexistente e diretório existente sem estrutura necessária rejeitados com saída 1, sem fallback ou criação de temporários. Primeiro ensaio identificou incompatibilidade de ConvertFrom-Json do Windows PowerShell 5.1 com a chave vazia do lockfile; corrigida somente a leitura no iniciador para JavaScriptSerializer do .NET, sem alterar lockfiles/fontes. Depois, os quatro casos passaram.
- Diagnósticos executados em novos processos PowerShell com `RemoteSigned` exclusivamente no processo, sem Bypass/Unrestricted ou políticas persistentes alteradas. Node portátil v26.13.1 e sonda simples por stdin aprovados: PATH com Node somente no filho, TEMP/TMP exclusivos sob D:\Temp\4PrintShops, APP_BUILDER_TMP_DIR em subpasta e caches sob D:\Cache\4PrintShops; DESKTOP_TEST_MODULES/ELECTRON_PATH apontando para a fonte selecionada. ELECTRON_RUN_AS_NODE/NODE_OPTIONS filtrados somente no filho; APPDATA/LOCALAPPDATA preservados. Snapshots externos do ambiente Process/User/Machine e políticas permaneceram iguais; nenhuma variável herdada continha referência ao antigo C:\Dev na conferência.
- Fonte independente e fonte histórica preservadas nas comparações completas de inventário/metadados antes/depois dos quatro casos. Setups .41/.42 originais relidos por SHA-256 e iguais aos hashes conhecidos. Nenhuma alteração de perfis operacionais, scripts históricos, fontes ou configuração global; somente diretórios exclusivos de diagnóstico em D: criados, sem limpeza ou preenchimento de caches.
- Espaço livre observado nesta etapa: C: **12.582.121.472 bytes (11,718 GiB)**; D: **72.601.821.184 bytes (67,616 GiB)**, sujeito à atividade normal do sistema; sem atribuir variações a uma causa não verificada.
- **Próxima etapa depende de autorização:** definir e revisar um controlador específico para os testes isolados, com fonte/perfil/saídas explícitos e tratamento dos caminhos/baselines pendentes da MIG-004B2A. O iniciador ainda não executa comandos arbitrários. Nenhum Electron, teste funcional, instalação, download ou build iniciado; nenhuma homologação funcional realizada. Ativação completa, empacotamento, backup independente e limpeza de C: permanecem pendentes. Sem commit ou push.

## 19/11/2026 — MIG-004B2B1: fonte independente em D: copiada e verificada

- Abertura conferida: workspace, diretório efetivo e raiz Git `D:\Dev\4PrintShops`, branch `main`, HEAD `19a5930a092b669c534922c1d13fd137a0b996c7`, árvore limpa. O checkpoint anterior foi publicado; esta etapa não cria commit ou push.
- Origem histórica preservada: `D:\Dev\4PrintShops-Builds\experimental.42-20261118-45a8d3bb\source`. Novo ambiente: `D:\Dev\4PrintShops-Builds\MIG-004B2B-validacao-20261119-g5NpOU\source`, em pasta inédita criada exclusivamente para esta preparação, sem usar o padrão `experimental.42-*`.
- Preflight: **49.389 arquivos regulares**, **4.468 subdiretórios**, **5.737.333.089 bytes (5,343 GiB)**; aproximadamente 73,1 GiB livres em D:, com reserva adicional mínima de 4 GiB. Nenhum hardlink ou reparse point encontrado na origem. Cópia integral de `source`, incluindo dependências, recursos e saídas históricas já existentes em `dist-desktop`; não é uma nova geração de pacote. Auditoria externa histórica não duplicada.
- Transferência por cópias regulares independentes, com exclusão de sobrescrita (`COPYFILE_EXCL`), sem hardlinks, junctions, movimentação, espelhamento ou exclusões. Inventários de arquivos/diretórios e tamanhos iguais; SHA-256 de todos os arquivos conferido no destino e recalculado na origem contra o inventário anterior à cópia: **zero divergências, ausentes ou extras**. Conteúdo da origem preservado. Destino sem hardlinks/reparse points ou vínculos para C:\Dev; identificadores dos arquivos distintos da origem.
- Evidências novas em `D:\Dev\4PrintShops-Builds\MIG-004B2B-validacao-20261119-g5NpOU\audit`: `source-before.json`, `destination-verified.json`, `source-after.json` e `copy-verification.json` (`approved=true`, aprovação somente da cópia). Quatro arquivos/37.197.252 bytes; ambiente completo com evidências: **5.774.530.341 bytes lógicos (5,378 GiB)**. Livre após preparação: aproximadamente 67,6 GiB, sujeito à atividade normal do volume.
- Disponíveis no destino: **9.905 arquivos de dependências**, Electron 45.1.0, electron-builder/app-builder-lib 26.18.0, updater 6.8.11, ASAR 3.4.1, builder-util-runtime 9.7.0, PDF.js 6.4.299 e pdf-lib 1.17.1; executáveis/manifestos do LibreOffice privado, DLL PDFium e auxiliares nativos presentes e incluídos na comparação integral.
- Iniciador `scripts/dev-environment.ps1` inalterado por SHA-256; permanece exatamente um candidato histórico `experimental.42-*`. Scripts históricos, Setups .41/.42 originais, perfis operacionais e configurações globais preservados. Referências textuais antigas foram copiadas fielmente, sem adaptações; ausência de vínculos não significa ausência desses caminhos nos scripts.
- **Nenhuma homologação funcional realizada:** sem Electron, testes funcionais, builds, downloads ou reconstrução de dependências. Próxima etapa mínima proposta: autorizar seleção explícita desta fonte independente no iniciador e diagnóstico leve antes de qualquer teste. Roteiros históricos com caminhos absolutos e onze linhas de baselines pendentes continuam sujeitos à investigação registrada na MIG-004B2A; não executá-los diretamente. Ativação completa, testes/builds, backup independente e limpeza de C: continuam dependendo de autorização própria.

## 19/11/2026 — Checkpoint local após MIG-004B2A

Operador autorizou um único commit local de `AGENTS.md`, deste contexto e de `scripts/dev-environment.ps1`, com a mensagem `chore(dev): registra regras e ambiente de migracao para SSD`. Base inicial conferida: `main`, HEAD `ddfa97aecdb7770622f8ca17191cf8ab3205fa72`, somente esses três arquivos pendentes e stage vazio. Push não autorizado nesta etapa. Histórico abaixo preservado; as conclusões seguintes atualizam as pendências do diagnóstico inicialmente bloqueado na MIG-004B1.

### MIG-004B1A — Diagnóstico dinâmico aprovado

- Restrição PowerShell `Restricted` diagnosticada: todos os escopos retornaram `Undefined`, inclusive MachinePolicy/UserPolicy, sem política administrativa explícita indicada. Arquivo sem `Zone.Identifier`.
- Iniciador revisado e executado com `RemoteSigned` exclusivamente no novo processo PowerShell autorizado, sem Bypass/Unrestricted ou alteração persistente de políticas. Saída normal, código 0.
- Node portátil **26.13.1** confirmado; caminhos em D:, TEMP/TMP exclusivos, caches e variáveis do processo filho conferidos. PATH recebeu Node somente no filho; ELECTRON_RUN_AS_NODE/NODE_OPTIONS filtrados somente nele. APPDATA/LOCALAPPDATA preservados; snapshots do ambiente pai, User/Machine e políticas permaneceram iguais.
- Criadas somente as pastas de diagnóstico/cache previstas em D:, sem copiar ou sobrescrever caches históricos. Nenhuma configuração persistente do Windows, instalação Experimental.42 ou perfil operacional foi alterado. Sem downloads, instalação, build ou teste funcional.

### MIG-004B2A — Auditoria de caminhos e recursos, somente leitura

- Examinadas **49 linhas**: 36 correspondências anteriores, duas na fonte isolada, nove verificações de espaço em C:/ e dois exemplos sintéticos. Classificação: 23 caminhos executáveis que exigiriam adaptação se reutilizados, quatro substituíveis por ambiente, nove históricos, dois sintéticos e **onze que exigem investigação adicional**. Linhas podem conter mais de um caminho; não representam 49 arquivos ou 49 literais distintos.
- Onze linhas pendentes: `check-preservation.mjs` 2/6/10/14; `diagnose-profile42.mjs` 2; `final-protection42.mjs` 2/4/7/8/9; `preserve-operational.mjs` 2, todos na pasta externa histórica .42. Dependem de baselines antigos, recibos encadeados e instalação/perfil operacional; não substituir C: por D: nem aplicar automaticamente baselines da .41 à instalação atual .42.
- **9.905 arquivos de dependências** da fonte isolada .42 conferidos por SHA-256 contra inventário preservado, zero divergências e mesma quantidade de arquivos. Cinco executáveis principais de empacotamento, MSI e quatro fontes LibreOffice, manifesto do runtime LibreOffice, arquivo/DLL PDFium e Setups .41/.42 conferiram com hashes conhecidos. Não houve revalidação individual dos 19.410 arquivos do runtime LibreOffice nem de todos os plugins/recursos NSIS.
- Scripts históricos, recibos, evidências, fonte original e instaladores preservados. Caches centrais npm/builder/Electron em `D:\Cache\4PrintShops` vazios; cache npm histórico sem conferência integral e cache builder com validação parcial. Presença de cache não garante build offline.
- Recomendada preparação independente em pasta inédita de D:, por exemplo `D:\Dev\4PrintShops-Builds\MIG-004B2-validacao-<identificador>`, com cópias e inventários conferidos, perfis sintéticos e saídas novas, preservando origens e mecanismos de atualização/recuperação. Base do Setup .42 continua `4509c59baaef2f189b8f6959e0d799c85e909861`, distinta do HEAD documental.

**Próxima etapa proposta, ainda não autorizada:** seleção explícita da fonte no iniciador e preparador específico para pasta inédita, sem alterar roteiros históricos. O iniciador continua restrito ao diagnóstico; ferramentas não homologadas para builds em D:. Ativação completa pendente; builds e testes funcionais ainda não autorizados. Backup independente e limpeza de C: permanecem pendentes, sem autorização para limpeza. Este checkpoint não altera versões, código ou ambiente operacional da Experimental.42.

## 19/11/2026 — MIG-004B1: iniciador implementado; diagnóstico bloqueado pela política PowerShell

- Conferência inicial: `Get-Location` isolado e raiz Git em `D:\Dev\4PrintShops`, branch `main`, HEAD `ddfa97aecdb7770622f8ca17191cf8ab3205fa72`. Alterações documentais da MIG-003 em `AGENTS.md` e neste contexto preservadas.
- Criado `scripts/dev-environment.ps1`, restrito ao modo `Diagnostic` (padrão), sem controlador genérico de comandos/builds. Calcula aplicação/repositório e ferramentas pela localização do próprio script; temporários e caches pela unidade efetiva. Exige exatamente uma fonte preparada `experimental.42-*` com Electron; verifica presença de recursos e rejeita reparse points nos caminhos utilizados. Presença não equivale a auditoria de hashes ou homologação funcional.
- Ambiente planejado somente em `ProcessStartInfo.EnvironmentVariables`: Node portátil 26.13.1 no início do PATH do filho; `TEMP/TMP` em execução exclusiva sob `<unidade>:\Temp\4PrintShops`, `APP_BUILDER_TMP_DIR` em sua subpasta `builder`; caches npm/builder/Electron sob `<unidade>:\Cache\4PrintShops`. Indica dependências e Electron preparados por `DESKTOP_TEST_MODULES`/`ELECTRON_PATH`; remove `ELECTRON_RUN_AS_NODE` e `NODE_OPTIONS` somente do filho. Não atribui APPDATA/LOCALAPPDATA nem altera perfil operacional. Diretórios novos seriam criados sem apagar/copiar/sobrescrever caches históricos.
- Diagnóstico implementado para executar apenas Node `--version` e uma sonda simples por stdin; confere caminhos, cwd, ambiente herdado, APPDATA/LOCALAPPDATA e saída normal. Compara snapshots privados do ambiente do processo pai e dos escopos User/Machine; não imprime valores arbitrários de ambiente ou credenciais. Relata referências a `C:\Dev` somente por arquivo/linha em scripts selecionados do projeto e roteiros externos de primeiro nível; não é varredura integral dos arquivos históricos.
- **Validação realizada:** análise sintática pelo parser do PowerShell aprovada. **Tentativa de executar o diagnóstico bloqueada antes de iniciar**, com `PSSecurityException/UnauthorizedAccess`: política efetiva `Restricted`, escopos listados como `Undefined`. Nenhuma política alterada ou contornada. Portanto versão do Node, sonda do filho, criação dos diretórios e comparações internas de isolamento **não foram executadas nem declaradas aprovadas**.
- Windows, instalação Experimental.42, perfis operacionais, arquivos de C:\Dev e scripts históricos de build não foram modificados. Sem instalações, atualizações, downloads, builds, testes funcionais, exclusões, commit ou push.
- **Pendências:** operador decidir como disponibilizar execução permitida do iniciador em etapa autorizada, sem mudança automática de política; depois concluir diagnóstico e comprovar isolamento em execução. Ferramentas continuam não homologadas para builds em D:. Preparação/validação de fonte isolada, eventual aproveitamento conferido dos caches históricos, testes do aplicativo e empacotamento dependem de nova autorização. Backup independente e eventual limpeza de C: permanecem separados.

## 19/11/2026 — MIG-003: regra permanente e checkpoint da migração para D:

MIG-001 e MIG-002 concluídas; nova sessão vinculada a `D:\Dev\4PrintShops`. A ativação completa do ambiente de desenvolvimento permanece pendente. Esta MIG-003 registra somente a regra operacional em `AGENTS.md` e este checkpoint; não autoriza implementação, ativação, limpeza, commit ou push. Os registros anteriores abaixo são históricos e suas autorizações não ampliam o escopo desta etapa.

### MIG-001 — Cópia e conferência inicial

- Copiados **36,97 GiB** para o SSD D: (36,965 GiB no relatório original).
- **332.577 arquivos** conferidos sem divergências de SHA-256 entre origem e destino.
- Originais preservados em C:, sem limpeza.
- Instalação operacional Experimental.42 preservada, sem alteração de instalação ou perfil ativo.
- Relatório existente, externo ao checkout: [RELATORIO_MIG_001.md](D:/Dev/MIG-001-20261119-01/RELATORIO_MIG_001.md). O relatório permanece inalterado; resultados de integridade são os da MIG-001, sem nova auditoria integral nesta etapa.

### MIG-002 — Validação da nova sessão

- Nova instância do VS Code aberta no repositório em D: e nova sessão Codex iniciada.
- Diretório efetivo confirmado por `Get-Location` isolado: `D:\Dev\4PrintShops`; workspace e raiz Git conferidos no mesmo caminho.
- Git íntegro conforme a conferência da MIG-001 e verificações da MIG-002; branch `main`, HEAD esperado **ddfa97aecdb7770622f8ca17191cf8ab3205fa72**, árvore limpa e referência local `origin/main` igual, sem nova consulta ao remoto.
- Ferramentas copiadas e identificadas em `D:\Dev\4PrintShops-Tools`, incluindo Node portátil 26.13.1; dependências completas da .42 presentes no ambiente isolado `D:\Dev\4PrintShops-Builds\experimental.42-20261118-45a8d3bb\source` (9.905 arquivos). O checkout ativo mantém dependências parciais.
- Ativação completa ainda pendente: `node` resolve para a instalação global em C:, `TEMP/TMP` continuam em C: e há caminhos absolutos antigos no teste `tests/pdf-size-manager-browser.mjs` e em roteiros externos. `D:\Temp\4PrintShops` e `D:\Cache\4PrintShops` existem e estavam vazias na conferência.
- Nenhuma instalação, build, teste funcional ou alteração de ambiente executada na MIG-002.

### Regra operacional permanente

Registrada em `AGENTS.md`, sob o título **Migração de workspace — reinicialização obrigatória da sessão Codex**: após mudança de localização, concluir a execução anterior, encerrar a conversa antiga e a janela anterior do VS Code, abrir nova instância diretamente no novo repositório e iniciar nova conversa Codex. Antes de novas implementações, executar `Get-Location` isoladamente e conferir diretório efetivo, workspace, raiz Git, branch, HEAD, caminhos absolutos, links e dependências anteriores. Interromper se houver divergência de workspace ou diretório, sem correção automática.

O operador relatou comandos executados com caminhos incorretos, erros e travamentos em migrações anteriores. Reiniciar a sessão e o editor no repositório correto foi a solução operacional adotada; não se afirma uma causa técnica definitiva não comprovada.

### Próximas pendências — dependem de nova autorização

- Configurar caminhos ativos e ferramentas para D:.
- Configurar temporários de desenvolvimento por processo.
- Validar o ambiente isolado em D:.
- Preservar recursos críticos em armazenamento independente; destino de backup no servidor ainda pendente.
- Liberar espaço no C: somente após autorização e validação.

Na abertura da MIG-003: diretório efetivo e raiz Git em `D:\Dev\4PrintShops`, branch `main`, HEAD `ddfa97aecdb7770622f8ca17191cf8ab3205fa72`, sem alterações locais inesperadas. Conteúdo histórico preservado; código-fonte, ferramentas, configurações do Windows, arquivos de C: e Experimental.42 instalada permanecem fora do escopo. Sem testes, builds, migrações, commit ou push nesta etapa.

## 19/11/2026 — Homologação operacional da experimental.42 e checkpoint documental autorizado

Operador confirmou o funcionamento prático da experimental.42: atualização .41→.42 concluída com sucesso e gerenciamento geral de tamanhos aprovado no escopo planejado. Esta é a situação atual; os registros de 18/11 abaixo preservam as etapas e pendências existentes em cada momento. A aprovação abrange o catálogo e o gerenciador compartilhados entre Redimensionar, Step & Repeat e N-up, com os testes supervisionados e a preservação de dados já registrados. Não declara o aplicativo inteiro estável nem conclui outras homologações.

Base exata do instalador permanece **4509c59baaef2f189b8f6959e0d799c85e909861**; o novo checkpoint é exclusivamente documental. SHA-256 do Setup .42: **7AE0D35B919714481E8E64AA0E304038DAD0A6548E1A1944C974CF0ACBDEBC1C**. Última auditoria de 18/11: instalação única por usuário, 19.512 arquivos de carga íntegros, transação concluída, F4 328×478/F6 318×318/F8 238×328 mm e IDs preservados, cadastro sintético ausente e preferências iguais ao backup. Espaço livre então medido: 5,690 GiB; não é nova medição de hoje.

Pendências mantidas: Livreto (inclusive sua integração ao catálogo e homologação física) e PDF-011 (agrupamento seguro com OutputIntents, planejado antes da primeira estável). PDF-010, abas e atualizador pelo servidor não foram iniciados. Redefinição para nove tamanhos continua como operação separada, dependente de autorização específica e backup verificado; não foi executada.

Autorizados agora revisão dos quatro documentos, commit documental e push normal para origin/main após conferência de segurança/remoto. Na conferência inicial: main/4509c59, 15 commits locais à frente, nenhum atrás após fetch; origin/main em 2affffb2b99a5be54485f995f156b63e6c91dae8. Instalação e perfis operacionais não serão modificados nesta etapa. Não criar release/tag nem publicar instalador. Resultado do commit/push será informado ao operador após a execução; não pressupor sincronização somente por este registro.

## 18/11/2026 — Validação supervisionada concluída; aguardar homologação prática

Operador confirmou exclusão somente de TESTE42-EDIT, ausência nas três ferramentas e após reabertura, três formatos corretos e encerramento normal. Leitura final em clone privado aprovada: F4 328×478, F6 318×318 e F8 238×328 mm, mesmos IDs e registros completos do backup; sintético ausente; demais chaves persistentes idênticas. Marcador catalogoCompleto=true introduzido pelo CRUD autorizado é previsto na .42, sem reset. Perfil atual com 78 arquivos, estável durante a leitura; não se exige igualdade binária global após uso normal. ZIP pré-instalação e fixtures preservados por hash. Livre 6.109.941.760 bytes (5,690 GiB). Evidência privada catalogo-final42-confirmado.json.

Instalação real já auditada: .42 única por usuário, 19.512 arquivos de carga idênticos, transação committed sem pendências, 76 arquivos do perfil iguais antes da primeira abertura. Testes manuais confirmados: gerenciador, visualização dos três formatos, inclusão/edição/exclusão sintética, sincronização Redimensionar/Step & Repeat/N-up, persistência/reabertura, orientação e giro independentes e mensagem/Ver detalhes. PDF descartável do operador observado na prévia; fixtures próprias não foram abertas no teste manual. Livreto mantém opções fixas e ficou fora do escopo. HEAD 4509c59baaef2f189b8f6959e0d799c85e909861; main ahead 15, somente quatro documentos modificados. Sem novos commits/push/reset/impressão/limpeza; aplicativo fechado. Aguardar homologação prática do operador antes de declarar .42 aprovada; redefinição para nove tamanhos exige autorização separada.

## 18/11/2026 — Orientação e Girar arte confirmados pelo operador

Operador confirmou no PDF descartável que alternar F4 Retrato/Paisagem e Girar arte 0°/90° funcionou de forma independente, somente na prévia. Próximo passo autorizado: excluir exclusivamente TESTE42-EDIT 124×235 mm pelo gerenciador, preservar F4/F6/F8, conferir ausência nas três ferramentas e após fechamento/reabertura normal; então fechar aplicação para auditoria semântica final. Nenhum reset dos nove tamanhos/impressão/commit/push; homologação prática pendente.

## 18/11/2026 — Ver detalhes confirmado pelo operador

Operador confirmou que a explicação do aviso de corte abriu corretamente por Ver detalhes. Resultado manual confirmado, sem impressão ou aceite de corte informado. Próximo passo: conferir orientação da folha e Girar arte somente na prévia do PDF descartável, sem editar cadastros; depois remover apenas TESTE42-EDIT, verificar persistência da exclusão e auditar catálogo final com aplicação fechada. Homologação prática ainda pendente; sem reset/commit/push.

## 18/11/2026 — Sincronização nas três ferramentas confirmada pelo operador

Operador confirmou TESTE42-EDIT também em Redimensionar, completando a conferência manual nas três ferramentas com Step & Repeat e N-up já confirmados após reinício. A resposta não confirmou abertura de Ver detalhes; esse teste permanece pendente, assim como orientação/Girar arte, exclusão somente do sintético com persistência e auditoria semântica final. Homologação prática ainda pendente; sem reset/impressão/commit/push.

## 18/11/2026 — Persistência confirmada pelo operador; Livreto mantém opções próprias

Operador reiniciou normalmente o programa e confirmou TESTE42-EDIT disponível no Step & Repeat e N-up. Relatou ausência no Livreto, que mostra A3/A4: conferência de pdf-booklet-ui.js e pdf-booklet-layout.js confirma papel fixo A4/A3 e formato final A5/A4, sem integração ao catálogo, conforme exclusão explícita de alterações do Livreto na .42. Não é falha de sincronização das três ferramentas abrangidas; nenhuma mudança no código. Redimensionar, abertura de Ver detalhes, orientação/Girar arte e exclusão sintética persistente ainda aguardam confirmação, além da leitura semântica final com app fechado e homologação prática. Não excluir formatos operacionais/reset/imprimir/commit/push.

## 18/11/2026 — Cadastro sintético salvo e PDF descartável confirmados pelo operador

Operador confirmou inclusão de TESTE42-TEMP 123×234 mm e edição para TESTE42-EDIT 124×235 mm, mantendo F4/F6/F8, e classificou arquivos_unidos.pdf como sintético/descartável. Bloqueio por possível trabalho real esclarecido; nenhuma exportação/descarte executado pelo agente. Leitura atual da UI mostra Step & Repeat, folha personalizada 29×21 mm em Paisagem, Girar arte 0° e aviso de corte/margens com Ver detalhes disponível. Solicitada conferência manual do cadastro nas três ferramentas e abertura dos detalhes; não imprimir/excluir ainda. Persistência após reabertura, exclusão sintética, orientação/arte e conferência semântica final continuam pendentes. Recibo privado confirmacao-crud-e-pdf-sintetico.json; instalação auditada, homologação prática pendente. Sem reset/commit/push.

## 18/11/2026 — Preservar documento aberto antes de continuar testes de reabertura

Operador confirmou gerenciador aberto com F4/F6/F8 corretos; CRUD temporário TESTE42-TEMP→TESTE42-EDIT solicitado, ainda sem confirmação. Em leitura posterior, UI mostrou arquivos_unidos.pdf, uma página350,01×399,97mm, alterações não exportadas e prévia temporária de redimensionamento10×10mm, diferente da fixture de duas páginas A4 criada pelo agente. Nenhuma aplicação/exportação/descarte/impressão foi executada pelo agente. Perguntado ao operador se o documento é sintético/descartável ou trabalho real. **Não fechar/reabrir nem descartar o workspace enquanto essa condição não estiver esclarecida e eventual trabalho real preservado pelo operador.** Auditoria da instalação e preservação pré-primeira-abertura continuam aprovadas; testes restantes/homologação ainda pendentes. Evidência privada observacao-workspace-pendente.json na supervisão .42.

## 18/11/2026 — .42 instalada e auditada; testes supervisionados em andamento

- Operador informou término e clique em Concluir. Sem processos de Setup/aplicativo/Office ativos antes da auditoria. Versão interna e registro únicos HKCU em 1.0.0-experimental.42, mesmo GUID/destino por usuário. **19.512 arquivos de carga idênticos por SHA-256** ao payload auditado; 19.516 arquivos instalados no total, quatro auxiliares conhecidos (inventário, debug.log dinâmico, desinstalador e ícone).
- Nova transação única .4ps-upgrade-Mwsffu, **committed**, origem .41/destino .42, sem failure/recoveryError e sem pendências. **Perfil76 arquivos exatamente igual ao backup antes da primeira abertura**, preservando catálogo e F4 328×478/F6 318×318/F8 238×328 mm. Setup .41 e .42 preservados por hash. Livre observado 6.115.635.200 bytes (~5,70GiB); nenhum resíduo/diário/quarentena foi limpo.
- Primeira leitura do inventário pelo auditor falhou por tratá-lo como UTF-8; arquivo corretamente escrito em UTF-16LE com BOM pelo instalador. Corrigido somente o leitor externo, mantendo a falha registrada e exigindo BOM/identidade; repetição integral aprovada, sem alterar instalação/perfil. Evidências em supervisionada42-20261118/auditoria-pos-instalacao42.json e registro-pos-instalacao42.json.
- Depois da auditoria, aplicação instalada aberta normalmente; tela Início app://4printshops observada por UIA. Controle de janelas voltou a responder, mas captura/foco de entrada continuam sem confiabilidade. Operador solicitado a abrir Arquivo→Gerenciar tamanhos e confirmar os três registros, sem editar operacionais. Testes de CRUD sintético/sincronização/reabertura/orientação/arte/PDF sintético/Ver detalhes e homologação prática ainda pendentes. Fixtures próprias criadas na pasta privada de supervisão, sem PDF real. Sem reset para nove formatos/impressão/commit/push/novas funcionalidades.

## 18/11/2026 — Operador iniciou a instalação .42; conclusão ainda não conferida

- Destino final observado na UI é o existente por usuário em AppData/Local/Programs/4PrintShops Hold Test. Operador confirmou explicitamente iniciar .41→.42 nesse destino, sem documentos em trabalho/reset. Por falha de captura/geometria, o clique único Instalar foi realizado pelo operador, que confirmou o início. UIA observou Instalando, navegação desativada, barra4711→4836 e nenhum erro visível nessas leituras.
- Durante o acompanhamento, list_windows passou a retornar timeout. Feita recuperação prescrita: espera2s/repetição, reset somente do kernel JavaScript de controle (não do instalador), inicialização sky e nova tentativa; timeout persistiu. Solicitado ao operador o estado atual e, quando concluído, fechar apenas o assistente e manter a aplicação fechada.
- Nenhum arquivo de transação consultado durante a instalação, segunda execução, processo encerrado à força, restauração/reset/impressão/commit/push. Resultado real ainda desconhecido; não declarar sucesso ou falha do Setup pelo erro da ferramenta de acompanhamento. Após conclusão confirmada e processos finalizados, executar auditoria de .42/única instalação/carga19512/diário concluído/perfil76 e formatos, antes de abrir a aplicação. Auditor externo preparado com gate ui-conclusao.json em pre-atualizacao42-20261118-bb06eeb8/supervisionada42-20261118; ainda não executado. Testes funcionais pós-instalação continuam pendentes e homologação depende do operador.

## 18/11/2026 — Setup .42 aberto; aguardando navegação e confirmação do operador

- Procedimento supervisionado autorizado, com confirmação final antes de Instalar ainda obrigatória. Reconferidos Setup .42 SHA esperado, Setup .41 preservado, 19.515 estáticos da .41 e 76 arquivos do perfil/backup idênticos. Antes da abertura: sem processos conflitantes, nove diários finais, espaço 9,63 GiB e margem 3,47 GiB além da reserva; destino/identidade HKCU por usuário preservados.
- Aberto somente o Setup autorizado, PID inicial12420, versão .42 confirmada na UI. Última tela observada: escolha de modalidade, opção Apenas para mim (SAGRADA02) disponível e todos usuários desabilitado. Nenhum clique em Instalar. Captura falhou (FrameArrived/window capture timed out), clique recusado por coordinate input geometry is unavailable; tentativa Tab sem mudança de foco confirmável. Leitura UIA textual funciona; não realizar cliques/teclas de ativação sem alvo/foco confiável.
- Solicitado ao operador selecionar Apenas para mim e avançar com Próximo somente até o destino, sem Instalar, para nova conferência e confirmação explícita. Nenhum arquivo de transação consultado após abrir o Setup, nenhuma segunda execução/força de processos. Instalação e testes pós-instalação ainda não realizados; não declarar .42 instalada/aprovada. Evidências: pasta privada pre-atualizacao42-20261118-bb06eeb8/supervisionada42-20261118. Sem reset/restauração/impressão/commit/push.

## 18/11/2026 — Preflight de instalação .42 e backup aprovados; aguardar operador

- Aplicativo fechado, nove diários finais/nenhum pendente; consulta final sem Electron/4PrintShops/Office/Setup/MSI, serviço MSI parado. .41 confirmada no ASAR/registro HKCU e 19.515 estáticos íntegros; funcionamento apoiado no uso relatado, sem nova abertura normal nesta etapa. Identidade/destino por usuário preservados.
- As 19 diferenças foram classificadas: 10 caches, 6 logs diagnósticos, DIPS interno Chromium, journal Local Storage operacional e journal Session Storage. CRC32C dos journals/manifestos e SQLite integrity_check aprovados. Preferences JSON válido e hash igual ao baseline da recuperação. Nenhuma restauração; causas individuais não atribuídas. Registro das 21 diferenças históricas mantido.
- Leitura exclusivamente em clone privado confirmou catálogo v1 com **três formatos**, texto completo idêntico ao backup histórico verificado: **F4 328×478, F6 318×318, F8 238×328 mm**. Tema claro, painéis320/230 e marcador histórico preservados. Nenhuma alteração operacional suspeita detectada; sem preencher nove sugestões/reset. Perfil original/backup76 arquivos reconferidos iguais após a leitura.
- Backup privado atualizado: C:/Users/SAGRADA02/AppData/Local/4PrintShops-Backups/pre-atualizacao42-20261118-bb06eeb8; 76 arquivos/8463943 bytes. ZIP 214751 bytes, leitura integral/76 hashes aprovada, SHA-256 **D73F6C73CAEB1B9234E2CADEB4CE4F4F5882FE880C28CC3E07FB9F6CE349054F**. Inventário/ZIP/catálogo/recibos privados fora do Git. Setup .42 SHA 7AE0D35B…BDEBC1C e Setup .41 SHA 253D672D…7D16D3 relidos e corretos.
- Livre 9,796 GiB; reserva conservadora 6,164 GiB para cópias/Setup/temporários/rollback, margem 3,632 GiB (mínimo extra2GiB atendido), sem exclusão. Relatório e recibo completos na pasta privada. **Podemos iniciar somente após confirmação explícita do operador; Setup não executado.** Preservar diários/quarentena; em falha não repetir automaticamente. Retorno após commit exige procedimento/autorização separados, sem downgrade direto. Nenhuma instalação/desinstalação/impressão/commit/push.


## 18/11/2026 — Esclarecimento sobre o perfil operacional

Responsável confirmou uso da .41 no início do turno das 13h, sem lembrar os minutos. Relato compatível com as gravações 13:31–13:39, sem comprovar causa exata das 19 diferenças novas. Preservar perfil atual e registro das 21 diferenças históricas; sem restauração/reset. Setup/hash/auditoria .42 inalterados; instalação continua aguardando revisão e autorização específica. Nenhum teste ou build repetido, commit ou push nesta atualização documental.

## 18/11/2026 — Setup .42 auditado e aguardando revisão, sem instalação

- Ambiente isolado completo restaurado com 9.905 hashes conferidos, mesmo lockfile/versões preservados; checkout/lockfiles/node_modules ativos não alterados. Ausência anterior era preparação incompleta/execução no checkout parcial, não remoção dos pacotes pela limpeza.
- Novos: 56/56 nos quatro arquivos antes não carregados, 100/100 atualização (inclui os 56), 7/7 INST-VAL sintético .41→.42, Electron com quatro saídas normais/três reaberturas e zero erros. 141 funcionais anteriores reutilizados, não reexecutados. Catálogo vazio persistente, três ferramentas, gerenciador/nove sugestões, backups sintéticos e orientação/arte aprovados.
- **Commit-base do instalador 4509c59baaef2f189b8f6959e0d799c85e909861** (checkpoint novo de versão e quatro documentos); fonte exata de HEAD limpo e recursos/espaço aprovados antes do build. ASAR 411 arquivos/116 desktop/public idênticos; backup ativo incluído/legado excluído, LibreOffice 19410, PDFium/manifests/fuses e carga integral 19512 hashes aprovados, sem privados. --verify-distribution em perfil exclusivo passou.
- Setup: C:\Dev\4PrintShops-Builds\experimental.42-20261118-45a8d3bb\source\dist-desktop\hold-native-test\4PrintShops-Hold-Test-1.0.0-experimental.42-Windows-x64-Setup.exe; **515070952 bytes**, SHA-256 **7AE0D35B919714481E8E64AA0E304038DAD0A6548E1A1944C974CF0ACBDEBC1C**, não assinado. Livre final ~11,298 GiB, sem limpeza adicional. .41: 19515 estáticos iguais. **Perfil não aprovado por igualdade: 57/76 arquivos iguais, 19 diferenças novas**, em caches/logs/armazenamento Chromium, gravações 13:31–13:39 (depois dos testes isolados e antes do build). Causa não confirmada; nenhuma restauração das diferenças novas ou das 21 históricas. Sem instalação/push/impressão/reset. Interrompidas ações além de diagnóstico/documentação ao detectar a diferença.
- Relatório completo: C:/Dev/4PrintShops-Builds/experimental.42-20261118-45a8d3bb/RELATORIO_EXPERIMENTAL_42.md; evidências privadas fora do Git. Quatro documentos finais pendentes para revisão, sem commit adicional. **Aguardar revisão das diferenças do perfil e autorização específica antes da atualização supervisionada .41→.42.** Manter Setup .41, backup atual e conferir formatos/reabertura posteriormente; não redefinir automaticamente o catálogo. Setups .29/.36 seguem indisponíveis.


## 18/11/2026 — Ambiente isolado .42 recuperado e validações pendentes aprovadas

- Responsável autorizou corrigir exclusivamente o ambiente isolado e concluir o Setup se todas as verificações passarem; instalar/publicar/push continuam proibidos. Causa confirmada: checkout ativo com 2.212 arquivos de dependências históricos versionados, enquanto validações anteriores usavam a árvore completa de 9.905 arquivos. A preparação .42 anterior ainda não tinha source/dependências; a chamada foi feita diretamente no checkout incompleto. Os alvos da limpeza não incluíam o checkout/preparo preservados, e o procedimento de 16/11 já documentava 7.693 arquivos adicionais no ambiente isolado. DESKTOP_TEST_MODULES não redireciona imports ESM gerais.
- builder-util-runtime 9.7.0 e @electron/asar 3.4.1 estão no lockfile, como dependências transitivas do conjunto Electron/builder/updater. Manifesto e lockfile do checkout/preparo/deps/.41 coincidem: lock SHA-256 0bcfa9c9d88373ae7f2e86c96f3ec4f93857b05107704dd834b4a8f016e0c0a9, 313 pacotes. Nenhuma adição arbitrária ao projeto, import de produção alterado, npm install ou alteração de lockfile/node_modules ativo.
- Restaurada a árvore completa já produzida pelo npm ci anterior em C:/Dev/4PrintShops-Builds/experimental.42-20261118-45a8d3bb/source: 9.905 arquivos conferidos contra inventário SHA-256 independente preservado da .40, antes e depois da cópia. 445 arquivos de fonte/documentação/configuração copiados e conferidos; 2.212 arquivos de dependências rastreados idênticos aos preservados. Restaurar essa origem comprovada manteve o resultado reproduzível do lockfile e os binários Electron previamente verificados, sem nova resolução/download. Node 26.13.1 para compatibilidade PDF, Electron 45.1.0, builder/app-builder-lib 26.18.0, updater 6.8.11, ASAR 3.4.1, builder-util-runtime 9.7.0, pdfjs-dist 6.4.299 e pdf-lib 1.17.1.
- Reexecução dos quatro arquivos antes não carregados: **56/56**. Em seguida, suíte completa transacional: **100/100**, zero ignorados; inclui esses 56, não somar como casos únicos. INST-VAL .41→.42: **7/7** sintéticos, sucesso/quarentena/lock/catálogo preservados, rollback, interrupção entre diário e movimentação, erro inesperado recuperável, instalação limpa, reinstalação e EPERM com handle Windows real. Sem executar Setup ou instalação operacional. Harness externo adaptado para imports da cópia isolada; erro de variável no controlador foi corrigido antes de iniciar os sete casos, sem alterar testes de produto/proteções.
- Electron isolado aprovado: PIDs 16676/17360/20028/21472, quatro saídas normais código 0 em 185/182/204/157 ms, três reaberturas, zero erros. Origem app://4printshops, sandbox/contextIsolation e nodeIntegration=false. Gerenciador global, nove sugestões, CRUD/IDs, três ferramentas sincronizadas, exclusões/vazio persistentes, backups sintéticos em disco SHA-256, redefinição/restauração sintéticas e orientação/Girar arte conferidos. Evidência 4ps-size-manager-dszkmh no TEMP exclusivo. Nenhuma abertura/redefinição do perfil real.
- **141/141 funcionais anteriores reutilizados**, não reexecutados nesta recuperação: fontes funcionais permanecem iguais; mudanças do checkout restritas a versão e documentos. O filtro já aprovado permanece e será conferido novamente no pacote real. Baseline novo somente de leitura do perfil operacional: 76 arquivos, preservando registro das 21 diferenças anteriores à recuperação; não restaurado. Conferência final deve comparar esse novo baseline.
- Logs/recibos/harnesses permanecem exclusivamente na pasta externa .42, fora do Git. Checkpoint autorizado reúne somente versão Hold Test 1.0.0-experimental.42 e quatro documentos revisados; identidade/appId/produto/exe/perfil/políticas NSIS mantidos. **Build ainda não executado neste checkpoint:** após HEAD limpo, validar fonte exata, LibreOffice/PDFium/fontes/NSIS e armazenamento; só então gerar novo Setup e auditar ASAR/carga integral. Sem PDF-010/PDF-011/abas/Livreto/servidor/impressão/instalação/push.


## 18/11/2026 — .42 interrompida antes do build por dependências ausentes

- Checkpoints autorizados: filtro/teste 4bec43c330245e1ec5c534c8c55f125dd268a309 e documentação 49b425a35369be9473af24b20699a2b4013aa25a. HEAD 49b425a/main ahead 14 da referência local; sem push. ExtraMetadata .42 preparado, identidade preservada, sem checkpoint dos metadados porque o preflight não foi aprovado.
- Nova execução funcional: **141/141** (inclui filtro/backup/distribuição novamente; 15/15 separado antes do primeiro checkpoint). Atualização: **44 aprovados e quatro arquivos não carregados** por builder-util-runtime/@electron/asar ausentes no checkout ativo. DESKTOP_TEST_MODULES não redireciona imports ESM gerais. Interrompido conforme instrução; sem Electron/INST-VAL .41→.42/validação integral de recursos/build/Setup/auditoria/instalação. Harness INST-VAL preparado, não executado.
- Próxima etapa após revisão: cópia isolada da fonte atual com dependências completas preservadas e hashes conferidos, regressões transacionais/INST-VAL/Electron isolado, recursos e armazenamento; somente depois commit de metadados/HEAD limpo e novo Setup. Não alterar node_modules ativo nem instalar pacotes para contornar a falha. Nenhuma correção funcional foi necessária ou feita.
- .41: 19.515 arquivos estáticos iguais e Setup preservado. Perfil somente lido: 21 diferenças frente ao snapshot anterior à limpeza, com gravações até 10:54:50, anteriores aos testes das 11:42; não afirmar igualdade de perfil nem causa dessas diferenças. Nenhum reset/restauração/perfil operacional aberto. Nove diários finais, nenhum pendente; sem processo conflitante identificado. Livre ~20.23 GiB; nenhuma limpeza adicional.
- Resultados e limites em PREFLIGHT_FILTRO_E_ARMAZENAMENTO_42_2026-11-17.md e relatório externo C:/Dev/4PrintShops-Builds/experimental.42-20261118-45a8d3bb/RELATORIO_PREPARACAO_EXPERIMENTAL_42.md. Logs/perfil/harnesses privados fora do Git. Quatro documentos e metadado ficam pendentes para revisão, sem novos commits após a interrupção. **Não instalar .42.**


## 18/11/2026 — Retomada autorizada da experimental.42

- Limpeza aprovada pelo responsável; preparação/testes e Setup .42 autorizados condicionalmente. Instalação/desinstalação não autorizadas. Preservar .41 instalada, perfil/catálogo reais, PDFs e Setups; sem impressão/push/novas funcionalidades.
- Checkpoint `4bec43c` contém somente electron-builder.config.cjs e tests/package-format-backup.test.mjs. Revisão do diff e nova execução válida: 15/15 filtro/backup/distribuição, sintaxe e whitespace aprovados. Primeira invocação não validou produto: cwd acima da aplicação e caminho de dependências incorreto; correção restrita ao ambiente de execução, seguida de aprovação completa.
- Próxima etapa autorizada: metadados .42, regressões de catálogo/orientação/proteções, Electron/reinícios isolados, INST-VAL sintético .41→.42/EPERM/quarentena/recuperação, integridade de recursos/dependências e margem de armazenamento. Somente com tudo aprovado: checkpoint dos metadados, fonte de HEAD limpo, novo Setup e auditoria integral, sem instalar.
- Manter catálogo persistido inclusive vazio; não executar reset real. PDF-010/PDF-011/abas/atualizador pelo servidor/Livreto fora do escopo. Documentação da limpeza aprovada segue abaixo; nenhum backup/log privado será incluído no Git.

## 18/11/2026 — Oito alvos removidos; revisão antes de empacotar .42

- Diretório padrão confirmado por Get-Location sem parâmetros/workdir: C:\Dev\4PrintShops. Backups no servidor e 86.087 registros de origem revalidados. Removidos exclusivamente os sete alvos mais a auxiliar de hardlinks; restante da auditoria preservado. A primeira tentativa esvaziou o projeto antigo, mas a raiz foi bloqueada por três auxiliares ociosos do Codex com cwd antigo; somente eles foram encerrados, a verificação foi ampliada para cwd/direito de exclusão e a retomada concluiu os oito alvos. Histórico completo no preflight.
- Livre: 10,768 → 20,584 GiB durante a operação; aumento líquido observado 9,816 GiB, sujeito à atividade normal do volume. Pós-conferência: 20,597 GiB. Não somar a auxiliar aos dados originais. Resultados individuais e bytes nos recibos privados da auditoria de armazenamento.
- Pós-limpeza: 2.660 arquivos do checkout íntegros, 19.515 estáticos da .41 iguais, 76 arquivos do perfil idênticos ao snapshot pré-limpeza; Setup .41 e Git íntegros. HEAD bccce139/main, 12 commits locais; seis caminhos pendentes preservados. Apenas quatro documentos de continuidade atualizados depois da conferência. Sem consulta nova ao remoto, commit ou push.
- Armazenamento resolvido para revisão; filtro e 15/15 são resultados anteriores. Metadados .41, sem build/Setup .42/instalação/reset/impressão. Setups .29/.36 seguem indisponíveis. **Aguardar revisão antes de retomar metadados, testes/preflight funcional e empacotamento .42.** Lista exata, bloqueio/resolução e medições em PREFLIGHT_FILTRO_E_ARMAZENAMENTO_42_2026-11-17.md. Instalação/perfis/PDFs reais/Setups preservados/recursos futuros fora da limpeza.


## 18/11/2026 — Recuperação somente leitura: diretório padrão ainda antigo

- Oito alvos presentes, sem remoção parcial identificada; seis intermediários iguais ao inventário e auxiliar com 25.730 hardlinks, sem junctions/links simbólicos. ZIP/manifesto/bundles e cópias pendentes revalidados. Dois arquivos adicionais são referências técnicas para árvore já preservada; nenhum código novo.
- .41: 19.515 arquivos correspondem aos hashes anteriores; somente `debug.log` cresceu com registros de execução. Perfil: 76 arquivos legíveis, sem alteração pelo diagnóstico. Nenhum processo de limpeza ou transação pendente. Livre ~10,81 GiB.
- **Não executar limpeza nesta sessão:** chamada sem workdir ainda retorna a pasta antiga, embora ferramentas do projeto tenham sido executadas explicitamente em `C:\Dev\4PrintShops`. Não repetir o CLI de troca que retornou aborted na execução anterior. Prosseguir somente em execução efetivamente vinculada ao checkout ativo; autorização dos oito caminhos já existe. Diagnóstico persistido em `PREFLIGHT_FILTRO_E_ARMAZENAMENTO_42_2026-11-17.md`. Sem exclusão, Setup, commit ou push.

## 18/11/2026 — Arquivamento anterior à limpeza

- HEAD `bccce139`, main ahead 12, sem commit/push. Filtro aprovado preservado; documentos anteriores mantidos. Bundle no servidor novamente verificado: histórico completo até o HEAD e SHA-256 inalterado.
- Autorizado preparar backups privados dos dados locais e seis intermediários, sem duplicar PDFs/Setup .41. Backup usa objetos SHA-256 deduplicados e manifesto recuperável. Detalhes e conclusão da verificação integral no documento de preflight/armazenamento.
- Não houve exclusão, alteração operacional, build ou redefinição. Aguardam revisão os sete caminhos originais e a pasta auxiliar de hardlinks criada pelo arquivamento; remover também esses vínculos será necessário para liberar os bytes. Antes de excluir o projeto antigo, trocar o workspace desta sessão para `C:\Dev\4PrintShops`.
- Setups .29/.36 continuam indisponíveis; não presumir reconstrução. Próximo passo permanece autorização final de limpeza, medição do espaço e retomada do preflight .42, sem geração de Setup nesta etapa.

## 17/11/2026 - Filtro .42 corrigido e plano de armazenamento em revisao

- Regra de backup mantida com inclusao exata de pdf-format-backup.js; legado excluido. FileMatcher real e regressao do backup/distribuicao: 15/15 aprovados, sintaxe e diff-check aprovados; sem build ou novas suites completas.
- Politica: .29 producao, .36 beta, .41 desenvolvimento. Projeto antigo e intermediarios candidatos a remocao somente apos ultima conferencia e preservacao verificada dos itens locais ignorados. Plano de sete caminhos recupera aproximadamente 9,72 GiB; C: tem cerca de 11,28 GiB livres. Nenhuma exclusao executada.
- 31 casos reais verificados por hash no servidor; bundles dos repositorios (incluindo 12 commits locais), alteracoes atuais e Setup .41 preservados com verificacao. Setups .29/.36 nao localizados; responsavel informou provavel exclusao anterior, preservacao dessas versoes continua pendente.
- Detalhes, caminhos completos, itens locais e condicoes em PREFLIGHT_FILTRO_E_ARMAZENAMENTO_42_2026-11-17.md. Arquivamento dos dados ignorados/geracoes e decisao sobre instaladores ausentes precedem limpeza. Metadados ainda .41; .42 aguarda armazenamento, preflight e autorizacao de empacotamento. Perfis, formatos reais e instalacoes intactos. Sem commit/push. Historico abaixo preservado.

## 17/11/2026 — Encerramento OutputIntents e impedimento prévio da .42

- Diagnóstico CASO-030/031 revisado pelo responsável e encerrado sem correção: abertura individual aprovada, ambos com ICC sRGB, união bloqueada intencionalmente por OutputIntents; dez ensaios anteriores (oito recusas/dois controles aprovados) e integridade dos originais mantidos. Não remover guarda/perfil para forçar união.
- Criado PDF-011 no planejamento após conferir IDs/escopos: agrupamento seguro com OutputIntents antes da primeira estável, independente de catálogo/PDF-010/Livreto e excluído da experimental.42. Testar ICC, compatibilidade de perfis, referências globais, saída íntegra e bloqueios negativos; implementação não iniciada.
- .42 autorizada para os tamanhos já implementados na base funcional `44a69da` mais documentos revisados, sem reset de catálogo operacional. **Preparação interrompida no preflight:** regra `!public/js/tools/*backup*` exclui o novo módulo obrigatório `pdf-format-backup.js`; resultado confirmado pelo FileMatcher real do builder, sem gerar pacote. Deixar a regra produziria pacote incompleto. Corrigir o filtro preservando a exclusão do legado `pdfs backup.js` requer revisão antes de prosseguir. Configurações/metadados ainda não alterados, versão permanece .41.
- Espaço ~11,34 GiB, abaixo da reserva de 12 GiB da última geração; recursos/dist/auditoria anteriores medidos em ~7,91 GiB, fora consumo temporário/margem. Nenhuma limpeza. Não executar novo build até rever filtro e capacidade. Novos testes funcionais, Electron isolado, INST-VAL .41→.42 e auditoria de pacote não foram iniciados; sem Setup .42/SHA. .41 instalada, Setups, perfis e formatos reais intactos, sem push.
- Checkpoint documental separado autorizado, somente quatro documentos revisados, sem PDFs/capturas/perfis privados. Aguardar revisão dos impedimentos antes de retomar a geração; autorização de instalação continua ausente. Histórico abaixo preservado.

## 17/11/2026 — Diagnóstico separado: CASO-030/031 não combinam

- Checkpoints do catálogo concluídos em `44a69da`, main inicialmente limpa/ahead 11. Dois novos PDFs da entrada local registrados incrementalmente como CASO-030 (cinco páginas)/031 (uma), categoria união; cópias por hash, nomes/recebidos intactos, IDs anteriores preservados. Nenhum PDF no Git.
- Interface Edge isolada: quatro aberturas (novos e controles), todas as páginas dos novos renderizadas. Dez uniões: oito recusas com novos (par nos dois sentidos, isolados e cruzados com controle), duas uniões de controles aprovadas com duas páginas. Recusas preservam objeto/revisão/nome/quantidade de páginas do workspace. Hashes finais iguais, zero erros inesperados.
- Causa confirmada: guarda `/OutputIntents` em `PDFCombinacao.add:51`; ambos têm mesma intenção/perfil ICC sRGB. Não é assinatura/XRef no diagnóstico atual. Erro sem código específico produz mensagem genérica, com motivo técnico em Ver detalhes. Onze módulos conferidos iguais à fonte .41 auditada; não executado aplicativo instalado, somente código equivalente no navegador.
- Detalhes públicos sem conteúdo privado em `DIAGNOSTICO_UNIAO_CASOS_030_031_2026-11-17.md`; relatório/nomes/hashes/capturas privados em `.desktop-test/casos-reais/evidencias/uniao-030-031-20261117/`, execução `execucao-F4GAzU`. Proposta: diagnóstico específico e estudo de preservação criteriosa de perfis/estruturas globais, sem remover guarda ou perfil. Aguardar autorização, separado de catálogo/PDF-010/Livreto. Nenhuma correção/commit/push/Setup/impressão; documentos ficam pendentes.

## 17/11/2026 — Revisão do catálogo, A4 oficial e checkpoints separados

- Lista inicial e medidas oficialmente confirmadas pelo responsável, inclusive **A4 210×297 mm**; não há mais pendência de interpretação de A4. Todos os nove formatos são administráveis. Registros provisórios anteriores mantidos como histórico.
- Revisados os 10 modificados e oito novos, com preservação externa dos 18 arquivos e conferência de escopo. Checkpoint A `ab483e623be36ee6af1cb6068782e09b6594f9f6` e C `0e80a687fb211bf5e9b18e6b91e713afdc2f8a2a` criados; B reúne gerenciador/menu/integração e documentação. Ordem A→C→B necessária à consistência, sem descartar ou substituir o checkout. O roteiro Edge compartilhado foi separado por trechos e seu conteúdo final preservado.
- A e C testados em cópias do índice: Edge catálogo/orientação/exportação aprovados; C também 22/22 unidades. Conteúdo funcional final: nova execução **108/108** serial, Edge `%TEMP%\4ps-size-manager-cst0Yp` e Electron `%TEMP%\4ps-size-manager-3Fg23w`, quatro processos/três reinícios cada, catálogo vazio persistente, sincronização, giro e backup/reset/restore sintéticos aprovados. Electron código 0 em 173/215/160/121 ms, sem console.error/exceções. As suítes completas anteriores não foram reexecutadas.
- Redefinição disponível somente como manutenção técnica assistida, não como botão. Antes da chamada explícita exige-se confirmação humana do operador e backup privado salvo, relido e verificado; o método valida backup/estado, mas não abre diálogo. Não há reset na inicialização/atualização. Escreve/restaura apenas `4printshops.pdf.formatos.v1`, sem perfil completo, e entrega F4 328×478/F8 238×328/F6 318×318/Faixa 320×700 na redefinição autorizada. Nenhum dado real tocado.
- Relatório A/B atualizado; complemento final externo em `C:\Dev\4PrintShops_Auditoria_Checkpoints_Catalogo_20261117-SUPhnT\RELATORIO_CHECKPOINTS.md` registra hashes finais e conferências após commits. Sem push/build/Setup/instalação/impressão. .41 instalada intacta; preparar .42 só após revisão/autorização, preflight/metadados/INST-VAL/auditoria/espaço. PDF-001/008/004/009 Em testes, PDF-010 não iniciado.

## 17/11/2026 — Simplificação posterior do catálogo geral, sem redefinir dados reais

- Decisão do responsável substitui a preservação de compatibilidade completa com catálogos antigos. A/B anteriores reaproveitadas, não descartadas. Lista inicial única: A3 297×420, A4 210×297, A5 148×210, A6 105×148, A7 74×105, F4 328×478, F8 238×328, F6 318×318 e Faixa 320×700 mm; **confirmar interpretação A4 antes de reset operacional**.
- Somente chave ausente oferece sugestões; v1 existente é lista integral, inclusive legado vazio sem marcador. Removido o merge de iniciais no legado; exclusões/renomes não recriam itens. A primeira mutação de primeiro uso persiste a lista, e todos os registros são administráveis. Nenhum dado existente foi redefinido automaticamente.
- Preparados `exportarBackup`, `redefinir(backupAtual)` e `restaurarBackup`, com SHA-256, original exato/IDs/metadados, bloqueio de backup obsoleto, falhas/conflitos e escrita restrita à chave. Procedimento assistido no relatório A/B, não adicionado ao menu instalado nem executado no perfil real.
- Novas validações: 94/94 dirigidas seriais e 14/14 adicionais; Edge catálogo/exportação aprovado. Roteiro A/B ampliado Edge/Electron: quatro processos, três reinícios, catálogo vazio persistente, nove iniciais, reset/restore por arquivos de backup verificados, zero erros. Electron final `%TEMP%\4ps-size-manager-gyi6IX`, saídas normais 107/171/219/138 ms; dados e sentinela sintéticos preservados.
- Mesma branch/HEAD `a3dfdf7`, ahead 8 da referência local, sem staging/commit/push/build/instalação/impressão. Metadados de versão/atualizador e .41 instalada não modificados. Perfis/F4/F6/F8 reais e biblioteca preservados. Revisar A/B e procedimento antes de checkpoints/experimental; reset operacional separado e homologação futura. Histórico abaixo mantido.

## 17/11/2026 — .41 instalada; correção do catálogo e gerenciador global aguardam revisão

- A atualização operacional .40→.41 autorizada concluiu com transação `.4ps-upgrade-oCOQJR` committed, uma instalação, 19.516 hashes instalados e 76 arquivos do perfil/backup iguais antes da primeira abertura. Evidência anterior: `C:\Dev\4PrintShops_Auditoria_Atualizacao41_20261117-oH5gXG\pos-instalacao.json`. Operador abriu PDF sintético e confirmou F4/F6/F8 com dimensões esperadas. Não repetir esses dados como nova auditoria nesta tarefa.
- Durante a avaliação prática, operador relatou SRA3 ainda disponível nas montagens após exclusão no Redimensionar. Causa no código: opção estática e medidas próprias A3/A4/SRA3, além dos modelos não editáveis reinseridos pelo serviço. Nova autorização permite administrar todos os tamanhos; não proteger F4/F6/F8 artificialmente, mas preservar os registros reais durante testes.
- A: fonte única de opções/medidas, exclusão/edição também dos iniciais, desassociação segura da seleção em uso. B: Arquivo → Gerenciar tamanhos, CRUD com confirmação, ID/metadados preservados e recusa de edição obsoleta. Redimensionar/Step & Repeat/N-up compartilham o mesmo serviço e armazenamento. Resultado no checkout, ainda fora da .41 instalada.
- Testes: 87 regressões dirigidas seriais, roteiros Edge de catálogo/exportação e interface; novo roteiro Edge/Electron com três processos/dois reinícios, exclusões persistentes, catálogo íntegro e zero erros de console. Somente dados sintéticos e perfil exclusivo; sem PDFs reais, perfil operacional, instalação, build, impressão ou alterações do atualizador/proteções.
- **Revisão necessária antes de distribuir:** extensão da mesma chave v1 com `catalogoCompleto:true` na primeira alteração explícita. Sem gravação ao carregar e sem reinicialização; originais/metadados preservados. Voltar à .41 antiga após essa gravação pode exigir restauração revisada do backup da chave/perfil. Nenhuma conversão real foi realizada. Plano e evidências no relatório `SINCRONIZACAO_E_GERENCIAMENTO_TAMANHOS_2026-11-17.md`.
- HEAD continua `a3dfdf7`, main ahead 8 da referência local origin/main. Os três documentos já pendentes foram preservados por acréscimos; A/B e documentação ficam sem staging/commit/push, para revisão. Não gerar outra experimental automaticamente. Homologação operacional .41 e PDF-001/008/004/009 permanecem pendentes/Em testes. Registros anteriores de geração abaixo são históricos, não o estado atual de instalação.

## 17/11/2026 — Experimental.41 gerada e auditada, NÃO instalada

- Checkpoints autorizados: `fd8cd5aec2c9ca30b339358069c53734fc908437` (somente metadado .41) e `a3dfdf78a11c54ddd8df3c0678574c01800d1e2b` (três documentos de validação). **Base exata do build: a3dfdf7**, exportada do HEAD limpo, 2.646 arquivos conferidos; nenhuma fonte de snapshot anterior reutilizada.
- Pasta inédita `C:\Dev\4PrintShops-Builds\experimental.41-FINAL-20261117-dE9Y3n`. Setup em `source\dist-desktop\hold-native-test\4PrintShops-Hold-Test-1.0.0-experimental.41-Windows-x64-Setup.exe`, **515.071.478 bytes**, SHA-256 **253D672DDFB8C02346D540E767450A4D800395309AAB0F5655A55F9BDE7D16D3**, sidecar ao lado. Relatório `RELATORIO_EXPERIMENTAL_41.md` nessa pasta.
- Preflight: Electron 45.1.0/builder 26.18.0/updater 6.8.11, 9.905 arquivos de dependências e ferramentas verificados por hash; LibreOffice 19.410 arquivos/fontes correspondentes e PDFium fixados conferidos. Identidade/GUID/NSIS/perfil iguais à .40, sem processos ou transações pendentes no preflight.
- Geração local concluída com saída 0, sem publicação/instalação. Auditoria: ASAR 409 arquivos, 114 desktop/public iguais à fonte, correções/mensagens/interface presentes; fuses e manifestos aprovados; carga integral 19.512 arquivos idênticos por SHA, sem ausentes/extras ou biblioteca/perfis privados. Diagnóstico empacotado --verify-distribution aprovado com userData/TEMP exclusivos; nenhum perfil operacional aberto.
- Consulta por arquivo .ps1 bloqueada pela política do Windows; metadados Windows e SHA conferidos por cmdlets nativos, sem alterar política. Aviso conhecido de cauda NSIS no 7z delimitado registrado; CRC próprio e comparação integral aprovados. Setup **NotSigned**, como os anteriores.
- Nesta etapa não se repetiram 81/100/4/2 testes, ciclos Electron ou suíte PDF 567/567 já aprovados; executados preflight, geração e verificações de pacote acima. Setups .39/.40 e EXE/ASAR .40 instalados preservados nos hashes. F4/F6/F8/perfis/PDFs reais não alterados; sem impressão/push/limpeza.
- Livre após auditoria: **16.438.538.240 bytes (~15,31 GiB)**. main em a3dfdf7, ahead 8; estes novos registros documentais ficam pendentes, sem terceiro commit. **Aguardar autorização para instalar** e renovar backup privado/conferir formatos e ausência de trabalho ativo antes da futura atualização. PDF-001/008/004/009 Em testes; homologação operacional .41 pendente.

## 17/11/2026 — Encerramento instrumentado e INST-VAL .41 aprovados

- Base mantida em `3774468`, main ahead 6. Metadado .41 preexistente preservado separadamente; nenhuma mudança funcional/commit/push. Novos acréscimos nestes três documentos mantêm os registros anteriores, inclusive a falha original.
- Harness externo agora aguarda fechamento da montagem e ociosidade do main, mantém CDP até saída, registra app.quit/before-quit/close/will-quit/quit/diálogos e conserva controle até will-quit. Limite de saída permanece 15s, sem força ou mudança de proteções. A falha antiga não foi reproduzida; não há prova retrospectiva de app.quit na execução QaMbb1 nem causa funcional confirmada.
- Electron isolado: **três saídas normais**, 296/241/186 ms, PIDs 18140/10148/12180; duas reaberturas no mesmo perfil/origem com catálogo/IDs intactos, três ferramentas, orientação e arte 90° independentes, zero erros de console. Eventos confirmam app.quit efetivo e nenhuma recusa por busy nas execuções novas. Perfil operacional e .40 não usados.
- Execuções novas: **100/100** regressões de atualização/EPERM/LibreOffice/recuperação; **4/4** INST-VAL .40→.41 (sucesso/quarentena/dados, rollback, interrupção, erro inesperado); **2/2** instalação limpa .41 e reinstalação .41→.41 sintéticas. As 81 regressões anteriores não foram repetidas.
- Armazenamento: ~23,92 GiB livres, preparo .41 ~3,12 GiB e build .40 completo ~8,02 GiB. Reservar mais 6–8 GiB usando o preparo existente, ou 9–12 GiB em outra pasta integral. Cópias derivadas de auditoria ~2,57 GiB identificadas, **não apagadas**; não há necessidade de limpeza.
- Evidências: `C:\Dev\4PrintShops-Builds\experimental.41-INSTVAL-20261117-CqiKQD\audit\encerramento\RELATORIO_ENCERRAMENTO_E_INSTVAL.md`. Testes liberam retomada do empacotamento após revisão, com preflight/conferência integral dos recursos do preparo interrompido antes do build. Nenhum instalador gerado/instalado; homologação operacional .41 futura. PDF-001/008/004/009 Em testes.

## 17/11/2026 — Preparação .41 interrompida antes do empacotamento

- Checkpoint documental autorizado criado: `377446849c9d67d958ef89af2961707d457f0304`, Registra validacao Electron dos formatos compartilhados, somente os três documentos anteriores. main ahead 6, sem push. Metadado extraMetadata.version .40→.41 preparado em electron-builder.hold-native-test.cjs e permanece pendente, identidade/perfil inalterados.
- Fonte em `C:\Dev\4PrintShops-Builds\experimental.41-INSTVAL-20261117-CqiKQD`, base 3774468, 2.646 arquivos conferidos com override somente da versão. **81/81** testes funcionais/proteções PDF dirigidos aprovados nesta execução; 567/567 continua referência anterior.
- Repetição Electron em perfil isolado passou cadastro/catálogo nas três ferramentas e orientação, mas **falhou no primeiro encerramento**, Error: Fechamento não concluiu (15s), sem erro de console. Não houve reabertura nesta repetição; não substituir esse resultado pelo ensaio anterior aprovado. Roteiro encerrou o processo isolado no finally, sem tocar na instalação.
- Conforme instrução, interrompido o preparo antes do pacote: sem build/Setup/instalação .41. INST-VAL/EPERM/quarentena/rollback e preflight integral ainda não executados/concluídos nesta rodada. Preparação privada e logs preservados; não tratá-la como distribuição validada.
- Causa ainda não confirmada. Harness desativa seu controle antes de app.quit e não registra eventual recusa por estado ocupado; falta evidência de eventos/diálogos/estado no fechamento. Próximo passo proposto: instrumentar somente o harness, aguardar ociosidade/fechamento da montagem e manter diagnóstico até saída normal, sem aumentar timeout ou mudar proteções, antes de repetir.
- Relatório: pasta nova acima, `RELATORIO_PREPARACAO_INTERROMPIDA.md`. Setups .39/.40 funcionais reconferidos nos hashes conhecidos, preservados. .40/perfil/F4/F6/F8 reais intocados. Sem outro commit/push, impressão, limpeza ou nova implementação. Metadados e três documentos ficam pendentes de revisão.

## 17/11/2026 — Electron isolado: formatos e reabertura aprovados

- Base `2bf25e15d075f72f3238640f7ba53b611556419c`, main inicialmente limpa/ahead 5. Checkpoints A/B aprovados pelo responsável. Ele também informou avaliação positiva da interface/textos da .40 instalada, sem que isso homologue as novas melhorias ainda ausentes dela.
- Executado `startDesktop` real em Electron 45.1.0, origem app://4printshops, sandbox/contextIsolation e perfil temporário exclusivo. Checkout direto carecia de builder-util-runtime; usada cópia de 133 arquivos desktop/public/package.json idênticos por SHA, com dependências isoladas já preparadas, sem instalação ou mudança funcional.
- Cadastro/reutilização pelo Redimensionar; mesmos quatro formatos sintéticos em Step & Repeat/N-up. Orientação livre 297×420 ↔ 420×297/retorno e salvo 140×180 ↔ 180×140, arte 90° independente e quadrado 318×318 conferidos com prévia pronta. N-up usou grade compatível 1×1, sem contornar proteções.
- **Reabertura real aprovada:** encerramento normal e novo PID 16648→15760, mesmo userData/origem, três ferramentas conferidas; quatro IDs e JSON bruto idênticos, sem duplicação ou alteração. Nenhum erro de console na execução final. Perfil real, F4/F6/F8 reais e instalação não foram utilizados/modificados.
- Evidências e diagnóstico em `C:\Dev\4PrintShops_Auditoria_Electron_Formatos_2026-11-17\RELATORIO.md`, execução final `execucao-JZ2Lmv`. Impedimentos intermediários do roteiro foram registrados e corrigidos somente no harness externo.
- Código apto à preparação da .41 **após revisão/autorização**; ainda sem gerar/instalar, testes INST-VAL/pacote e homologação operacional futuros. Só documentos ficam pendentes; nenhum commit/push, instalação, impressão ou PDF real. PDF-001/008/004/009 Em testes. Registros seguintes mantêm o histórico anterior do ensaio pendente.

## 17/11/2026 — Checkpoints A/B autorizados; validação Electron somente planejada

- Checkpoint A: `d1ad3f0a93012503d1cba5c1d68881d9622f4b46` — Corrige orientacao de folhas personalizadas no Step Repeat. Cinco arquivos; apenas orientação e testes/adaptador. Index materializado em cópia isolada sem B: 14/14 testes e roteiro Edge aprovados.
- Checkpoint B: este registro acompanha **Compartilha formatos personalizados nas montagens PDF**, separado e filho direto de A. Inclui o catálogo local, sua integração e os cinco documentos revistos, preservando todos os registros anteriores de instalação/recuperação. O hash de B deve ser obtido no Git pela mensagem, sem embutir um hash autorreferente.
- Revalidação A+B: 37/37 testes direcionados; roteiros Edge de interface e catálogo aprovados; sintaxe dos 11 arquivos e diff/staging conferidos. A suíte serial 567/567 continua referência da rodada anterior, **não reexecutada neste checkpoint**.
- Código/testes preservados byte a byte na separação; nenhuma alteração funcional adicional. Após B, main fica cinco commits à frente da referência local origin/main, sem fetch/push. Sem build/instalador, impressão ou acesso ao perfil operacional.
- Próxima etapa: revisar o plano Electron no relatório A/B. Primeiro ensaio isolado com `userData` explícito e catálogo sintético; depois, somente com nova autorização de versão/instalação, verificar os formatos reais nas três ferramentas e após reabertura sem editar os cadastros. **Electron não homologado**; o app instalado continua sem estas melhorias. PDF-001/008/004/009 Em testes.
- Os tópicos seguintes são registros históricos anteriores aos checkpoints e foram mantidos integralmente.

## 17/11/2026 — Orientação e formatos nas montagens: implementação em revisão

- Autorização posterior ao diagnóstico abaixo: **A**, corrigir orientação personalizada; **B**, compartilhar com segurança o catálogo local já existente. Partes documentadas separadamente em [ORIENTACAO_E_FORMATOS_MONTAGENS_2026-11-17.md](ORIENTACAO_E_FORMATOS_MONTAGENS_2026-11-17.md).
- A: medidas efetivas 297×420 → 420×297 → 297×420, com arte 90° preservada; prévia e exportação sintética verificadas. Folhas quadradas e presets cobertos.
- B: PDFFormatos compartilhado por Redimensionar/Step & Repeat/N-up; salvar/reutilizar formatos na montagem. Chave `4printshops.pdf.formatos.v1` e versão 1 mantidas; IDs, nomes, medidas e metadados legados preservados. Sem migração, importação ou alteração do perfil real/F4/F6/F8.
- Verificações: suíte PDF serial **567/567**, prévia direcionada **12/12**, roteiro de interface Step & Repeat e roteiro A/B em Edge isolado aprovados, incluindo encerramento/reabertura reais, exportações e dados sintéticos F4/F6/F8. Ensaio adicional Electron isolado não iniciou o endpoint de depuração: **persistência em reinício do Electron ainda pendente**, sem usar o app instalado.
- HEAD permanece `6c5a7f8024a5130eb4ee5178037473655a8b04f2`; main três commits à frente de origin/main, sem fetch/push nesta rodada. Alterações anteriores dos dois documentos foram preservadas e acrescidas destes registros. Sem staging ou commit.
- Instalação experimental.40 inalterada; a correção só existe no checkout. Aguardar revisão/checkpoints e futura autorização de empacotamento/homologação. PDF-001/PDF-008/PDF-004/PDF-009 continuam Em testes; PDF-010 e demais implementações maiores não iniciadas.

## 17/11/2026 — Validação visual pausada: orientação de folha personalizada

- Inicialização/PDF e F4/F6/F8 na interface confirmados. Step & Repeat observado com prévia pronta, controles Orientação da folha/Girar arte separados; arte 90° manteve folha Retrato. Aviso resumido de corte/margens excedidas bloqueou Gerar/Imprimir; Ver detalhes expandiu folha 1/bordas esquerda e direita e ofereceu Ocultar detalhes, sem confirmar corte nas bordas.
- Problema confirmado pelo operador e UIA: com folha **Personalizada 297×420**, selecionar **Paisagem** manteve largura/altura e resumo 297×420; prévia não girou, arte continuou 90°. Causa em leitura de pdf-imposition-ui.js:83: troca de orientação só atualiza medidas quando há preset; custom fica sem tratamento, e config() usa as medidas diretamente. **Não considerar orientação/Step & Repeat homologados.**
- Testes interrompidos e evidências preservadas em `OCORRENCIA_ORIENTACAO_PERSONALIZADA_2026-11-17.md` e ui-ocorrencia-orientacao-paisagem.txt na auditoria Terceira40. Proposta mínima: tratar a orientação custom sem mudar rotação da arte, sujeita a autorização; nenhuma implementação, build/Setup, impressão, exportação, commit ou push. Dois PDFs sintéticos mantêm hashes originais. Atualização/perfil/quarentena aprovados anteriormente permanecem resultados distintos; reinstalação adicional tem origem não esclarecida.

## 17/11/2026 — Reinstalação adicional observada; origem não esclarecida

- Formatos agora confirmados também na interface instalada: lista Formato de página contém F4 — 328×478 mm, F6 — 318×318 mm e F8 — 238×328 mm; evidência ui-formatos-confirmados.txt. Nenhum cadastro editado/excluído pela automação. Menu Editar observado com Montagem / Imposição; solicitada abertura dessa ação para conferir Step & Repeat, ainda pendente na última leitura.

- Conferência seguinte: operador confirmou a abertura da fixture e a interface PDF foi observada carregada, com 1 página e status de alterações ainda não exportadas (redimensionamento para A4 em memória). Abertura/visualização básica aprovada; evidência ui-pdf-aberto.txt. Solicitada expansão de Formatos personalizados para conferir os cadastros na interface; lista ainda não observada expandida. Step & Repeat/Girar arte e Ver detalhes continuam pendentes, sem impressão/exportação pela automação.

- Atualização posterior: operador abriu o atalho verificado e a interface normal do executável instalado foi observada em UIA (`app://4printshops/index.html`, janela 4PrintShops), sem nova chamada de abertura pela automação. Inicialização normal confirmada; evidência `ui-aplicativo-aberto.txt`. Importação do PDF sintético ainda não confirmada: leitura posterior mostrava seletor nativo Importar arquivos, campo Nome vazio; solicitada seleção explícita do arquivo, sem impressão/exportação. Step & Repeat, formatos na interface e Ver detalhes permanecem pendentes.

- Após a primeira conferência, foram observados Setup/PID 13092 e tela Instalando, sem aplicativo aberto. Já encerrado, diário `.4ps-upgrade-o53zun` confirmou **.40→.40 committed**, iniciado às 01:54:04 de 17/11 (São Paulo), sem failure/recoveryError. O operador declara ter aberto o Setup uma única vez quando solicitado e depois dormido; **não atribuir a ele nova execução**. Não foi enviado novo comando explícito de Setup, mas chamadas de abertura do app via sky.launch_app retornaram 0x80070486; origem permanece desconhecida, inclusive sem excluir a ferramenta como causa. Abertura automática suspensa.
- Nova conferência pós-encerramento aprovada: 19.516 arquivos/hashes, perfil e backup de 76 arquivos idênticos, primeira quarentena e diário MSjeKu íntegros, .39 preservada e cache .40 correto. Atalhos Desktop/menu Iniciar conferidos somente em leitura: destino é o executável instalado em AppData/Local/Programs/4PrintShops Hold Test, sem argumentos extras. Solicitado ao operador abrir esse atalho, em vez da pasta do Setup que havia informado; testes de interface ainda pendentes.
- Evidências e retificação de atribuição em `OCORRENCIA_SETUP_ADICIONAL_2026-11-17.md`, `post-reinstall-observed.json/log` e `atalhos-somente-leitura.json` na auditoria Terceira40 já indicada. Sem reparo, rollback manual, nova execução de Setup pela automação, mudança funcional, impressão, commit ou push. O resultado anterior descreve a primeira atualização; esta observação posterior não altera seu histórico.

## 17/11/2026 — Terceira tentativa: experimental.40 instalada, conferência visual pendente

- Atualização real autorizada e iniciada em 16/11 pelo operador, usando somente o terceiro Setup (SHA EF30BE5D9AE3F33BB7F96EC1C94C2D5D1474463157094A03CD27107E55EC92F6, base 6c5a7f8), por usuário no mesmo destino da .39. Interface informou conclusão; Setup encerrou. Sem leitura do diário/diretórios transacionais durante execução. Controle de cliques da ferramenta bloqueado por RoInitialize/0x80070486; cliques feitos pelo operador com leitura UIA supervisionada.
- Conferência de 17/11, antes da primeira abertura: transação `.4ps-upgrade-MSjeKu` **committed**, sem failure/recoveryError; instalada **1.0.0-experimental.40**, uma única entrada/GUID. **19.516 arquivos** com hashes esperados, ASAR igual ao build aprovado. **76 arquivos do perfil e backup privado idênticos por SHA**, preservando F4 328×478, F6 318×318 e F8 238×328 mm e preferências. Setup .39 preservado íntegro; cache atualizado para o Setup .40 autorizado. Sem rollback/reinstalação manual.
- Lock original de 120 bytes preservado em `.4ps-upgrade-MSjeKu/quarantine/libreoffice-opencl-lock`, fora de lo; SHA 66d5eb3cc48cbd7131af4e677865f0f86f2bfbe35712d8c85862f5cc45f6b166 igual ao anterior, original ausente do runtime. Reter diário/quarentena/old, sem limpeza. Relatório e evidências em `C:\Dev\4PrintShops_Auditoria_Terceira40_2026-11-16-1794880110213\RELATORIO_TERCEIRA_TENTATIVA_40.md`; inventários/perfil completos somente no backup privado.
- INST-VAL-002 real aprovado e 005 aprovado para o resíduo real; 004 integridade antes da primeira abertura aprovada. INST-VAL-001 confirma registro único, mas não houve instalação limpa/reinstalação nesta tentativa; INST-VAL-003 sem falha induzida/rollback agora. Abertura automática do app também impedida pela ferramenta 0x80070486; aguardando abertura manual e verificações de PDF sintético, mensagens/Ver detalhes, Step & Repeat/Girar arte e cadastros/preferências em interface. **Não homologado para produção.** PDF-001/008/004/009 Em testes, PDF-010 não iniciado.
- Espaço livre pós-instalação 32.810.860.544 bytes (~30,6 GiB). HEAD 6c5a7f8/main ahead 3; somente estes documentos permanecem pendentes, sem commit/push, impressão, mudança funcional, novo build ou nova execução de Setup. Histórico anterior preservado.

## 16/11/2026 — Nova experimental.40 com EPERM/quarentena gerada e auditada; não instalada

- Checkpoint documental autorizado: `6c5a7f8024a5130eb4ee5178037473655a8b04f2`, Registra validacoes obrigatorias de instalacao, exclusivamente PLANEJAMENTO/CONTEXTO_ATUAL/ESTADO_PROJETO. É também o commit-base exato do novo build, exportado e comparado ao checkout (2.640 arquivos; diferenças somente CRLF/LF). Versão .40 já commitada, sem novos metadados ou .41.
- Pasta inédita: `C:\Dev\4PrintShops-Builds\experimental.40-INSTVAL-20261116-1794868567608`. Setup em `source\dist-desktop\hold-native-test\4PrintShops-Hold-Test-1.0.0-experimental.40-Windows-x64-Setup.exe`: **515.069.135 bytes**, SHA-256 **EF30BE5D9AE3F33BB7F96EC1C94C2D5D1474463157094A03CD27107E55EC92F6**, sidecar ao lado. Relatório completo `RELATORIO_EXPERIMENTAL_40_INSTVAL.md` e evidências em audit. Ambos Setups anteriores preservados e hashes reconferidos.
- Nesta execução: 100 regressões direcionadas aprovadas, três casos INST-VAL adicionais aprovados e desktop-smoke completo aprovado. INST-VAL-001/002: instalação/reinstalação e atualização .39→.40 sintéticas; INST-VAL-003/005: falhas, interrupções antes/depois do movimento, rollback, resíduo, desconhecidos, colisões, links e arquivos/processos em uso; INST-VAL-004: dados e preferências/F4/F6/F8 sintéticos preservados. Instalação Windows, atualização operacional e conferência do perfil real após instalar continuam pendentes; falta de espaço não foi injetada. Suíte PDF completa não reexecutada (referência anterior 560/560).
- Dependências isoladas reutilizadas: 9.905 arquivos comparados por SHA ao preparo/build auditado, Electron 45.1.0/builder 26.18.0/updater 6.8.11. Pacote --verify-distribution aprovado, 19.410 arquivos LibreOffice e startup privado; PDFium/manifestos/fuses conferidos. ASAR com 407 arquivos e 112 arquivos desktop/public idênticos ao staging, incluindo EPERM, quarentena, mensagens/Ver detalhes e Step & Repeat/Girar arte. Carga integral do Setup: 19.512 arquivos iguais por SHA, nenhum ausente/extra; sem biblioteca privada/perfis. Identidade igual à .39, canal automático ausente, assinatura NotSigned como no preparo anterior.
- C: livre **36.101.267.456 bytes (~33,6 GiB)** após auditoria. Fonte pronto para revisão da próxima tentativa, sem instalar, desinstalar, imprimir, publicar ou push. .39/perfil/F4/F6/F8 e biblioteca real não foram modificados. Antes de instalar: autorização, hash/versão, backup atualizado, Setup .39 para retorno, espaço, aplicativos fechados e ausência de transação pendente; preservar diário/quarentena. Estes registros finais ficam pendentes, sem segundo commit. PDF-001/008/004/009 continuam Em testes, homologação visual .40 pendente.

## 16/11/2026 — Validação de instalação passa a requisito permanente

- Decisão do responsável registrada em PLANEJAMENTO.md: novas versões exigem validação de instalação, atualização entre versões e recuperação após falhas, priorizando resíduos reais, preservação de perfis/F4/F6/F8 e rollback. Cenários identificados como INST-VAL-001 a INST-VAL-005, com versão/commit, hash do Setup, ambiente, resultados, evidências e pendências, seguindo a rastreabilidade das tarefas PDF. Testes isolados não substituem homologação operacional.
- Checkpoints anteriores: correção `5899f75123a710caa0d7565ee8f157035fd62259` e documentação `f2ed6a20b7ddcd95890587018ca22f07448797be`. Esta etapa altera somente documentação, sem novos testes, commit/push, build ou Setup. Escopo atual continua limitado ao necessário para gerar e validar a experimental.40, aguardando as autorizações correspondentes; .39 e dados operacionais preservados.

## 16/11/2026 — Revisão final e checkpoint da quarentena

- Criado checkpoint local autorizado `5899f75123a710caa0d7565ee8f157035fd62259`, Preserva residuos LibreOffice na atualizacao transacional, exclusivamente os três módulos desktop e o teste de transação. Revisão confirmou registro antes do movimento, recuperação nos estados anterior/posterior ao movimento, preservação em colisões, bloqueio de arquivos em uso e validador inalterado. A quarentena é irmã da instalação, fora de lo e de PLUGINSDIR; nenhuma limpeza de sucesso examinada a exclui. Sem risco de exclusão identificado nesse ciclo.
- Conferido log anterior de 72/72 testes aprovados, sem repetir a suíte; sintaxe dos quatro JS e diff staged --check novamente aprovados. Registros históricos abaixo mantidos integralmente, somente acréscimos documentais. Fonte pronto para incorporar em novo instalador após autorização: usar checkout deste checkpoint, verificar inclusão dos três módulos, pacote/versão/hashes e ausência de dados privados; renovar backup/perfil/F4/F6/F8, conferir Setup .39 para retorno, espaço, aplicativos fechados e ausência de transação pendente antes de instalar. Sem geração, Setup, limpeza operacional ou push; .39 preservada. Este registro e as tentativas anteriores seguem no checkpoint documental separado autorizado.

## 16/11/2026 — Quarentena transacional restrita do lock OpenCL (aguarda revisão)

- Base `2affffb2b99a5be54485f995f156b63e6c91dae8`; correção local sem commit/push. Relatos pendentes das duas tentativas abaixo preservados. Instalação operacional .39, perfil/F4/F6/F8, PDFs reais e ambos Setups não foram modificados nesta tarefa.
- `desktop/update-quarantine.mjs` reconhece exclusivamente o não gerenciado `lo/program/opencl/.~lock.cl-test.ods#`: 1–4096 bytes, UTF-8 válido, cinco campos CSV, primeiro vazio, proprietário não vazio, host local, data no formato do lock e URI file de `profile` ausente em `TEMP/4printshops-word-<alfanumérico>/profile`. TEMP é comparado por caminho real, inclusive aliases 8.3. Nome sem conteúdo reconhecido, perfil existente, links ou qualquer outro extra em lo causam recusa.
- `update-transaction.mjs` registra caminho original, tamanho e SHA-256 no diário antes de mover. Destino preservado: `<pasta da instalação>/../.4ps-upgrade-<id>/quarantine/libreoffice-opencl-lock`, fora de lo e do diretório temporário do Setup. Movimento nativo sem substituição, hash conferido com handle que recusa leitores/escritores concorrentes. Rollback reverte e confere o hash; interrupção após mover é recuperável pelo diário. Colisão/adulteração impede sobrescrita e pode exigir recuperação assistida (`recovery-required`).
- `update-registry.mjs` e o helper recusam processos ativos soffice/soffice.bin, opencltest e oosplash; fechar LibreOffice antes da atualização. Verificador/manifesto do runtime permanecem intactos e rigorosos. Retry EPERM de save() inalterado.
- 72 testes aprovados, zero falhas/ignorados: transação (incluindo 14 novos), EPERM (5), runtime LibreOffice, registro isolado, assinatura de atualização e atualizador. Sintaxe dos quatro arquivos JS e `git diff --check` aprovados. Node 26 isolado; dependências existentes externas resolvidas somente no processo de teste, sem instalação. Log: `C:\Dev\4PrintShops_Auditoria_Instalacao40_2026-11-16\quarentena-final.log`. Tentativa intermediária das duas suítes ESM não iniciou por dependência fora do checkout; execução final com resolução isolada aprovou ambas.
- Em sucesso, reter diário/quarentena e backups para revisão; não restaurar o lock em lo nem apagá-lo automaticamente. Liberação/descarte futuro depende de revisão. Nenhum Setup/build/atualização executado. Apto à revisão para checkpoint; geração e instalação posteriores exigem autorização. Homologação visual .40 pendente; PDF-001/008/004/009 continuam Em testes.

## 16/11/2026 — Segunda tentativa .40: recuperação e origem do lock conferidas

- Segunda transação `.4ps-upgrade-m8zN4A` rolled-back sem erro de recuperação; .39 instalada. Conferidos 19.515 arquivos anteriores e 76 arquivos do perfil/backup, zero divergências; F4/F6/F8 preservados, sem processo pendente.
- Lock OpenCL já estava na .39 antes da primeira tentativa, como arquivo não gerenciado; não está em nenhum dos dois Setups, recursos preparados ou win-unpacked. Conteúdo aponta a perfil temporário de conversão Word de 16/11 às 10:50; processo criador exato não identificado. Preservação do lock antigo no destino conflita com o verificador estrito de lo; filtro de build não resolveria.
- Sem alteração de código/política: aguardar revisão de tratamento transacional com preservação externa/quarentena antes de implementar. 15 testes existentes do runtime/EPERM aprovados; sem build/Setup/commit/push. Setup 78C10BE97F84DE19A64B9EF0AD76527483B2A9258994886E02CD27A0AD61F7FA preservado, reprovado para esta atualização. Relatório: `C:\Dev\4PrintShops_Auditoria_Instalacao40_2026-11-16\SEGUNDA_FALHA_LIBREOFFICE.md`.


## 16/11/2026 — Tentativa .40 abortada; .39 recuperada pelo próprio Setup

- Checkpoint dos três metadados/documentos aprovado e criado: `da04ca52e0eba49e889d801ac4ea69a15fbd6275`, Registra versao experimental.40 e seu empacotamento. Árvore limpa após o commit, main ahead 3 de origin/main, sem push. Registros desta tentativa ficam pendentes, sem outro commit.
- Todos os pré-requisitos passaram: Setup .40 no hash autorizado, .39 preservada, backup/perfil de 76 arquivos íntegros, F4/F6/F8 presentes, aplicativo fechado, sem instalador ativo/recuperação pendente e ~49,1 GiB livres. Tentativa única por usuário no mesmo destino, com ajuda do operador na UI por falhas de captura/entrada do Computer Use; sem pedido administrativo observado.
- Instalação abortada por EPERM ao renomear `journal.json.tmp` para `journal.json` em `%LOCALAPPDATA%\Programs\.4ps-upgrade-bQ4TGq`, fase copying, Exit=2. O Setup executou sua recuperação transacional, estado final **rolled-back**, sem recoveryError. Não houve desinstalação/reinstalação manual, nova tentativa, alteração de permissões/código ou novo pacote. Motivo do bloqueio do arquivo ainda indeterminado; não atribuir a antivírus/administração sem evidência.
- Conferência independente: 19.513 arquivos gerenciados e dois não gerenciados iguais aos hashes anteriores; 76 arquivos do perfil e backup iguais, atalhos e Setup/cache .39 íntegros. Apenas um registro instalado do Hold Test, versão **1.0.0-experimental.39**. Tela abortada encerrada, sem processo ativo desta tentativa. Não foi aberto o aplicativo após a falha; testes visuais/abertura de PDF na .40 permanecem **pendentes**.
- Relatório/evidências: `C:\Dev\4PrintShops_Auditoria_Instalacao40_2026-11-16\RELATORIO_TENTATIVA_EXPERIMENTAL_40.md`. Preservar diário/backups, aguardar revisão da ocorrência antes de nova tentativa/correção. Sem push, impressão, homologação ou implementação de PDF-010/catálogo/N-up. Config/fonte continuam .40 no checkpoint; instalação efetiva permanece .39.


## 16/11/2026 — Checkpoint criado e experimental.40 gerada; não instalada

- Checkpoint autorizado concluído: `be9e8e572d87abca5913f22990451b53f5aa9777`, mensagem Atualiza smoke desktop e planejamento da proxima experimental, exatamente os três documentos e três arquivos de teste/fixture aprovados. Sem push. main agora ahead 2 de origin/main; após o commit, somente metadado .40 e registros desta geração ficam pendentes, sem segundo commit.
- `electron-builder.hold-native-test.cjs`: única alteração de configuração é extraMetadata.version .39 → **1.0.0-experimental.40**. AppId/GUID/nome/perfil/atalhos/cache/políticas preservados; nenhum módulo funcional modificado. Build a partir do checkpoint + esse diff, em pasta exclusiva `C:\Dev\4PrintShops-Builds\experimental.40-2026-11-16`, usando dependências isoladas e publish never.
- Instalador: `source\dist-desktop\hold-native-test\4PrintShops-Hold-Test-1.0.0-experimental.40-Windows-x64-Setup.exe`, relativo à pasta de build; **515.067.551 bytes**, SHA-256 **47B96CE6D446EEF64B44285819F4456E29DDBA2E2017FD7473B5BBFA2FF16846**. Sidecar .sha256 junto do Setup. Config, PE e ASAR identificam .40; sem assinatura Authenticode, release ou canal publicado.
- Validação atual: 36 regressões de assinaturas/fixture, 79 testes de segurança/empacotamento/simulados e 1 transação sintética .39→.40 aprovados; desktop-smoke aprovado no novo staging. Executável empacotado --verify-distribution aprovou os 19.410 arquivos e inicialização do LibreOffice privado em perfil temporário. ASAR/recursos/fuses e carga do Setup conferidos; 111 arquivos próprios iguais ao staging, mudanças visuais presentes, sem biblioteca/perfis reais no pacote. Referência PDF anterior 560/560 serial, não repetida.
- .39 continua instalada; perfil e backup (76 arquivos), F4/F6/F8/preferências e Setup .39 reconferidos íntegros. Não houve instalação/desinstalação, impressão, limpeza manual ou alteração das pastas antigas/restauração. Atualização futura única com mesma identidade/perfil; retorno Windows desinstalar nova/reinstalar .39 preservada. Aguardar autorização para instalar e revisão de metadados/documentação antes de outro commit.
- Relatório completo e evidências: `C:\Dev\4PrintShops-Builds\experimental.40-2026-11-16\RELATORIO_EXPERIMENTAL_40.md`. PDF-001/008/004/009 Em testes; PDF-010 Planejado, sem catálogo/N-up ou novas funcionalidades. Avaliação visual/operacional e homologação continuam pendentes.


## 16/11/2026 — Desktop-smoke corrigido; aguardando revisão/checkpoint

- Base main/8081175, três documentos locais anteriores preservados. Corrigidos somente `tests/desktop-smoke.mjs`, nova fixture `tests/helpers/pdf-print-prepared-fixture.mjs` e seu contrato `tests/desktop-prepared-fixture.test.mjs`; nenhum módulo do aplicativo alterado. Roteiro antigo tentava reconstruir assinatura aplicada, operação corretamente bloqueada pela PDF-008, e clicava Prever sem abrir a sessão.
- Desktop agora reutiliza recusas/snapshots e console estrito do helper já aprovado no Edge; exportação legítima usa PDF sintético já preparado, sem assinatura e sem histórico interno forjado. Três recusas reconhecidas, bytes/objetos preservados; eventos adicionais/exceções continuam fatais. Único beforeunload esperado é validado na tentativa deliberada de fechar/cancelar; nenhum diálogo HTML pendente. Sem timeout aumentado.
- Desktop-smoke final aprovado no Electron 45.1.0; 36 regressões direcionadas aprovadas; controle negativo externo com erro inesperado reprovou com saída 1, como exigido. Sintaxe/diff --check aprovados. Suíte PDF anterior 560/560 serial preservada como referência, não repetida nem somada aos resultados atuais. Perfil original/backup (76 arquivos) iguais por SHA-256.
- Pronto para revisão/checkpoint e autorização de geração experimental; proposta inequívoca **Hold Test 1.0.0-experimental.40**, ainda sem mudar metadados .39 do config ou gerar pacote. Mesma identidade/perfil, atualização única da instalação .39; retorno por desinstalar nova no Windows/reinstalar .39 preservada. Instalação só após verificar pacote/autorização. Relatório: `C:\Dev\4PrintShops_Auditoria_DesktopSmoke_2026-11-16\RELATORIO.md`.
- GitHub main reconferido por consulta somente leitura em 5c547d2; main local 8081175 ahead 1. Quatro modificados (três docs anteriores e smoke), dois novos testes/fixtures, nada staged; sem commit/push/instalador/impressão/atualização. PDF-001/008/004/009 Em testes; PDF-010 Planejado, catálogo/N-up não iniciados. Experimental.39, formatos, biblioteca, checkout antigo e restauração intactos.


## 16/11/2026 — Dependências desktop isoladas e reversão simplificada

- Reversão adotada conforme experiência prática confirmada pelo responsável: desinstalar a nova versão pelo Windows e reinstalar normalmente a .39, com backup e conferência de F4/F6/F8/preferências. Bloqueio de downgrade direto não impede atualização; sem exigência de ensaio complexo na ausência de outro impedimento concreto.
- Setup .39 e perfil (76 arquivos) preservados fora do cache em `%LOCALAPPDATA%\4PrintShops-Backups\pre-atualizacao-2026-11-16-65bd9755-c29e-4678-8345-3dd88f1889be`, com hashes conferidos e dados originais inalterados. F4 328×478, F6 318×318 e F8 238×328 mm confirmados na cópia privada; nenhum perfil incluído no Git.
- Dependências preparadas em `C:\Dev\4PrintShops-Tools\desktop-prep-2026-11-16`: Electron 45.1.0, builder 26.18.0, updater 6.8.11, recursos privados LibreOffice/PDFium e ferramentas do builder verificados. Fonte de staging do commit 8081175, sem misturar os três documentos locais pendentes nem alterar node_modules do checkout. Canvas aprovado em Node 26 e Electron, sem rebuild; 75 testes direcionados aprovados.
- Pendência: desktop-smoke reprovou por esperar preparo seletivo de assinatura aplicada; botão atual mostra Ver restrições e não abre esse diálogo. Não modificar a proteção para atender o roteiro antigo. Relatório em `C:\Dev\4PrintShops_Auditoria_2026-11-16\RESULTADO_PREPARACAO_DESKTOP.md`; revisar teste de forma localizada antes de recomendar liberação da geração/instalação. Não houve correção de teste/código nesta preparação, instalador, atualização, desinstalação, impressão, commit ou push. .39, biblioteca e formatos preservados; tarefas PDF mantêm estados anteriores.


## 16/11/2026 — Planejamento da atualização única e Ctrl+A

- Checkpoint Edge concluído: `8081175ec099b1c2513008632ea6f4e19d730322`, main um commit à frente do origin/main observado (`5c547d2`), sem push. Árvore limpa antes desta atualização exclusivamente documental.
- Responsável substituiu a preferência lado a lado por atualizar a experimental.39 existente, preservando perfil/F4/F6/F8 e instalador anterior. Procedimento e diagnóstico em `C:\Dev\4PrintShops_Auditoria_2026-11-16\PROCEDIMENTO_ATUALIZACAO_HOLD_TEST.md`; preparação final e instalação aguardam revisão. Nenhuma dependência instalada, pacote gerado ou aplicativo/perfil modificado.
- PDF-010 (Ctrl+A no documento ativo) registrado **Planejado**: depois da validação da nova versão instalada e antes das implementações maiores do catálogo compartilhado e N-up. Sem implementação. PDF-001/PDF-008/PDF-004/PDF-009 continuam Em testes.


## 16/11/2026 — Validação Edge de assinaturas ajustada; aguardando revisão

- Base main/5c547d2, HEAD e origin/main iguais, árvore inicialmente limpa. Ambiente aprovado em C:\Dev\4PrintShops; Node 26.13.1 isolado já validado na etapa anterior com 560 testes PDF em execução serial. Essa suíte e o harness Electron 45 não foram repetidos neste ajuste exclusivo de testes.
- ConsoleAssinaturas, em tests/helpers/pdf-signature-console.mjs, exige exatamente um console.error por recusa de giro nos cenários aplicada, multipla e perms: janela temporal da ação, origem girarPagina em pdfs.js, frames da proteção/workspace, mensagem exata, codigo PDF_ASSINATURA_PROTEGIDA e assinaturaDetectada=true. Validação final exige os três cenários em ordem e nenhum registro adicional; Runtime.exceptionThrown nunca é aceito como recusa esperada. Console da aplicação não foi alterado ou silenciado.
- Integrados o roteiro pdf-browser-smoke e o chamador pdf-signature-policy-browser; removido desse último o filtro genérico por texto. Helper compartilhado compara SHA-256 dos bytes de entrada/visualização, revisões, páginas e manifesto dos objetos antes/depois de cada giro recusado. Espião de capturarEstado somente observa a instância e delega sem mudar retorno/estado; restaurado inclusive no controle negativo. Nenhuma proteção ou implementação de rotação modificada.
- Validação: cinco arquivos de teste com sintaxe aprovada; 13 testes Node do verificador aprovados. Edge 157.0.4341.60, perfil isolado, ambos os roteiros aprovados com três recusas esperadas/zero inesperados e documentos sintéticos preservados. Controle negativo PDF_TESTE_CONSOLE_INJETAR_ERRO=1 introduziu um console.error adicional dentro da primeira recusa: roteiro reprovou com saída 1 e dois registros, como exigido. Duplicatas idênticas, exceções não tratadas, origem/código/mensagem/momento divergentes e registros fora da ação também são recusados pelos testes do verificador.
- Evidências novas fora do checkout: C:\Dev\4PrintShops_Auditoria_Edge_2026-11-16. Sem biblioteca real nos cenários, dependência desktop, instalador, impressão, commit ou push. PDF-001/PDF-008/PDF-004/PDF-009 permanecem Em testes, prioridades e experimental.39 preservadas. Próximo passo: revisão do usuário antes de checkpoint; homologação instalada/operacional continua pendente.

## 16/11/2026 — Retomada do Step & Repeat; interface preparada, revisão pendente

- Recuperação após desligamento: main/9877601052aafe3415e55a1c5aa5814a2a8bb4a2, cinco commits locais à frente de origin/main. O checkpoint das mensagens já existe; não havia alterações pendentes daquela tarefa. Encontrados cinco arquivos modificados e um teste novo do Step & Repeat, todos reconhecidos na conversa. Estado recuperado preservado em .desktop-test/repeat-retomada-20261116-srmIhG/; nada descartado ou refeito.
- Parte A já implementada antes da interrupção: único seletor de rotação realocado para a área básica somente no Step & Repeat, com rótulo Girar arte, ao lado da orientação da folha; Colunas/Linhas e frente/verso mantidos. Resumo curto de folha/quantidade/escala; explicações e diagnósticos completos em Ver detalhes técnicos. Avisos geométricos, aceites e guardas permanecem ativos. CSS responsivo exclusivo de repeat; N-up restaura os mesmos controles e a apresentação anterior. Nenhum cálculo, algoritmo de imposição, motor PDF ou Livreto modificado.
- Evidências anteriores recuperadas: 560 testes PDF/41 arquivos aprovados; Chrome específico com antes/depois, orientação/rotação independentes, grades, duplex, teclado, navegação/guias e exportação; Chrome de assinaturas aprovado. Não repetida a suíte Node, pois os dois arquivos de produção permanecem idênticos ao estado recuperado validado. A regressão Chrome ampliada havia ficado sem resultado final; foi concluída agora, com aprovação. O teste específico foi complementado para conferir o plano antes/depois de abrir detalhes, rodapé na resolução menor, origem intacta e reabertura da exportação também no PDF.js.
- Resultados atuais: Chrome ampliado aprovou apresentação básica/avançada, escalas, geometria, marcas, erros/retry, cancelamento e regressões N-up/Livreto. Chrome específico aprovado em 1366×900, 1024×768 e 720×640, com exportação sintética de duas páginas e reabertura em pdf-lib/PDF.js. N-up teve estado completo idêntico ao baseline; Step & Repeat manteve plano/configuração iniciais idênticos. Nenhuma diferença de cálculo encontrada. Testes são do fonte atual no Chrome, não do aplicativo instalado.
- Capturas anteriores mantidas em .desktop-test/repeat-interface-byi02q/ (antes), .desktop-test/repeat-interface-iJZ2Dj/ (depois) e nova execução .desktop-test/repeat-interface-uTV7Ni/. Regressão retomada em .desktop-test/repeat-retomada-20261116-montagem/ e respectivo log. Não usar datas internas do host dos artefatos como data de homologação; a retomada foi em 16/11/2026.
- Parte B somente investigada: já existem PDFFormatos e adaptador localStorage na chave 4printshops.pdf.formatos.v1, com IDs/nome/largura/altura em mm; fallback em memória não persiste após encerrar. Proposta separada em docs/PROPOSTA_FORMATOS_COMPARTILHADOS.md. Nenhum catálogo, menu, migração, cadastro ou seletor compartilhado implementado/alterado. Perfis reais não foram abertos para inspecionar cadastros.
- Próximo passo: revisão do usuário antes do checkpoint. Homologação no aplicativo instalado, uso pelos operadores e papel/impressão permanecem pendentes. PDF-001/PDF-008/PDF-004/PDF-009 e as homologações anteriores continuam Em testes; prioridades, experimental.39, biblioteca e checkpoints preservados. Sem commit/push, dependência, instalador ou impressão nesta retomada. Registros abaixo são históricos.

## 13/11/2026 — Mensagens de balcão: ajustes autorizados, aguardando revisão

- Base main/23c1e39a4fcf7dfcf790c9ffa3c4b79122c52b3a, quatro commits locais à frente de origin/main, árvore inicialmente limpa. Esse é o checkpoint local já criado da PDF-009; os registros abaixo descrevem etapas anteriores. PDF-001/PDF-008/PDF-004/PDF-009 permanecem Em testes; prioridades e experimental.39 preservadas.
- Novo componente de apresentação pdf-feedback.js: título/frase simples e Ver detalhes recolhido, diagnóstico original, operação e código quando disponível. Copiar código envia somente o código; se a API não estiver disponível ou recusar, oferece seleção/cópia manual. Sem novos modais, classificação por texto ou alteração do documento ao consultar detalhes. Erros desconhecidos ficam genéricos no resumo, sem expor caminhos ou termos técnicos; não são classificados como PDF restrito.
- Reutilizados status do editor, conferência de união e painéis existentes de assinatura/montagem. Camadas e recorte marcado receberam apenas códigos nos pontos de recusa existentes; condições mantidas. Metadado de assinatura detectada acompanha os erros pelo worker para distinguir indícios inconclusivos. Nomes duplicados nos erros de união removidos; nome original segue no item e na origem. Ver restrições abre o diagnóstico de assinatura, sem iniciar preparo. Campo vazio continua distinto e mantém os caminhos permitidos.
- Recorte mantém aceite explícito para remover ajustes de corte/sangria; aviso curto explica que ocultar não apaga conteúdo. Limitações de acessibilidade permanecem em detalhes. A confirmação de cópia para impressão não afirma invalidar uma assinatura inexistente. Na interface de montagem, somente resumo de anotação e apresentação dos códigos foram alterados; nenhum motor, geometria ou fluxo de Step & Repeat/N-up/Livreto modificado.
- Validação atual: **560 testes PDF em 41 arquivos aprovados**, zero falhas/cancelados/ignorados (547 anteriores + 13 novos). Chrome específico: sete casos reais CASO-003/005/006/011/012/013/029 e dois sintéticos (comum/assinado), com detalhes, cópia do código, preservação do estado, aceite de sangria obrigatório e cancelamento sem aplicação. Área de transferência simulada no roteiro; nenhum conteúdo copiado para o Windows. Chrome de assinaturas, XRef/CASO-029/comum e recorte comum/marcado também aprovados. Não equivale a aplicativo instalado ou leitor independente.
- Evidências novas ignoradas: .desktop-test/mensagens/, .desktop-test/mensagens-chrome-EFjd3A/ e .desktop-test/xref-chrome-Wz9b0S/. Nenhuma origem, inventário ou evidência anterior da biblioteca alterada. Pendentes: diagnóstico específico na abertura individual de PDF protegido (sem inferir restrição pelo erro genérico da biblioteca), padronização dos demais erros sem código e homologação com operadores/no aplicativo instalado. Sem novo suporte de senha, camadas ou serialização XRef. Sem commit/push, dependência, impressão ou instalador. Aguardar revisão antes de checkpoint.

## PDF-009 — 13/11/2026, concorrência da abertura corrigida, Em testes

- Autorização específica após revisão: corrigir publicação obsoleta em abrir(), sem ampliar compatibilidade XRef. Base main/9eb6a7f, três commits locais à frente de origin/main; preservados os 15 modificados/quatro novos anteriores e índice vazio. Agora 16 modificados/quatro novos, pois pdf-preview.js também precisa impedir a publicação de canvas atrasado. Experimental.39 e prioridades inalteradas; PDF-001/PDF-008/PDF-004/PDF-009 permanecem Em testes.
- Cada tentativa incrementa geracaoAbertura antes da leitura, inclusive quando falha. Identidade/revisão e geração são conferidas após os awaits e imediatamente antes da única publicação; bytes, páginas e diagnósticos são preparados localmente. Erros antigos recebem PDF_ABERTURA_OBSOLETA. Geração não integra o snapshot restaurável. A tentativa recente inválida conserva o documento corrente e não revalida a abertura anterior.
- pdfs.js verifica geração e versão efetivamente publicada, descarta respostas/erros obsoletos sem rollback, mudança de mensagens ou navegação; ocupação continua pertencendo à tentativa da interface. pdf-preview.js recebe verificação opcional antes de publicar o canvas/adotar a fonte ou restaurar uma prévia; demais chamadores mantêm o padrão existente. A guarda de união continua independente e uma tentativa de abertura a invalida desde o início.
- Regressões antes: 15 casos, um aprovado/14 falhas. Guarda inicial: 15 aprovados; dois casos adicionais de prévia expuseram um canvas tardio (um aprovado/uma falha), corrigido antes da regressão final. Acrescentados também dois casos de load tardio. Resultado final: 19 novas regressões aprovadas; **547 testes PDF em 40 arquivos aprovados, sem falhas/cancelados/ignorados**. Contagens intermediárias são subconjuntos, não somar.
- Chrome atual: CASO-029/sintético XRef visualizados, origens idênticas, bloqueios sem downloads e worker recusado; comum girado/exportado/reaberto. Assinaturas PDF-008 e recorte PDF-001/PDF-004 sintético aprovados (12 viewports e 1.285.504 componentes RGBA iguais no marcado). CASO-029 também reaberto no PDF.js Node: uma página/116 operadores, zero saves; SHA256 c816a5d2208cc8860f86af0814d0d430ccd25b994847531bf0de4239b229f66b preservado. Biblioteca não modificada.
- Evidências incrementais ignoradas em .desktop-test/pdf009-abertura/ e .desktop-test/xref-chrome-RhL356/. Limite: operações obsoletas podem terminar em memória; não há cancelamento global ou isolamento/timeout rígido. Sem nova capacidade de reescrita XRef; aplicativo instalado, leitor independente e homologação operacional pendentes. Aguardar revisão antes de checkpoint; sem commit/push, instalação, impressão ou pacote. Registros abaixo são históricos.

## PDF-009 — 13/11/2026, CASO-029: abertura segura, reescrita bloqueada, Em testes

- Base main/9eb6a7f77161ff5c72ec93f3e11aa0ff959fe64b, três commits à frente de origin/main; Git limpo no início. Checkpoint PDF-004 já criado, conforme autorização anterior. ID PDF-009 disponível, sem prioridade definida e sem alterar prioridades/ordem anteriores. PDF-001/PDF-008/PDF-004 continuam Em testes.
- Diagnóstico: o arquivo recebido mantém revisões incrementais com Root 3 0 R e 3 1 R. pdf-lib carrega ambas no contexto. O aplicativo descartava os bytes de comuns e serializava antes da primeira prévia: a XRef gerada fazia a entrada do objeto 4 apontar a 3 1 obj, causando a recusa pelo PDF.js. Não era necessário reescrever só para visualizar a origem.
- Correção limitada: guarda de gerações concorrentes em pdf-xref-policy.js; cache seletivo dos bytes originais para assinatura protegida ou contexto incompatível, associado a PDF/revisão e incluído no snapshot/rollback. Prévia recebe cópia dos bytes enquanto identidade/revisão correspondem. Comuns sem esse risco mantêm o fluxo existente, sem cache extra; edições suportadas continuam vindo do workspace atualizado. Bloqueios de assinatura mantidos e anteriores ao novo diagnóstico.
- Gerações concorrentes são recusadas antes de editar/exportar, unir/extrair ou preparar cópias nos módulos internos pertinentes, incluindo worker via rascunho de união. Nada de reparo ou seleção de maior geração. Interface informa “PDF aberto para visualização” com reescrita bloqueada; ações não publicam sucesso ou download ao recusar. Step & Repeat, N-up e Livreto não alterados.
- Testes: antes da correção, fixture final com 12 testes, dois aprovados e dez falhas que demonstravam falta de cache/guardas. Suíte final: **528 testes PDF em 40 arquivos, todos aprovados**, incluindo 15 novos e os 513 anteriores. Chrome: sintético XRef, CASO-029 e comum; preservação de bytes, bloqueios sem downloads, worker e giro/exportação/reabertura do comum. Chrome de PDF-008 e de recorte PDF-001/PDF-004 aprovados. Registros em .desktop-test/xref-fix/ e evidências do novo browser em .desktop-test/xref-chrome-WDoszA/.
- CASO-029: visualização usa os mesmos 378.402 bytes e SHA256 c816a5d2208cc8860f86af0814d0d430ccd25b994847531bf0de4239b229f66b, zero saves; PDF.js Node lê uma página/116 operadores e Chrome renderiza. Giro/recorte/exportação recusados com PDF_XREF_INCOMPATIVEL sem mutação. Salvar sem object streams em ensaio isolado abriu com avisos/recuperação de XRef, sem comprovação de integridade; alternativa não adotada nem publicada.
- Limites: diagnóstico detecta a colisão de números/gerações conhecida, não certifica toda XRef ou conformidade PDF. Edição e exportação desse perfil continuam bloqueadas, mesmo sem alterações; não se adicionou “salvar original” como novo fluxo. .39 instalada existe, mas não foi executada; testes são do fonte atual em Chrome/Node. Sem instalador, impressão, dependência, versão nova, commit ou push. CASO-025–028 e seus seis PNGs não foram testados/corrigidos. Biblioteca e inventários anteriores preservados.
- Próximo passo: revisão do usuário antes de checkpoint; eventual suporte de reescrita exigirá proposta e regressões próprias. PDF-009 permanece Em testes.

## PDF-004 — 13/11/2026, validação lexical de MCIDs corrigida, Em testes

- Base main/a9071ac, dois commits locais à frente de origin/main. Preservados os nove arquivos modificados e quatro novos da implementação anterior; índice vazio. Autorização específica para corrigir a validação lexical, sem ampliar o perfil marcado ou mudar prioridades.
- Lexer mantém o lexema numérico, exige `[+-]?\d+` para MCID antes de converter e depois confere inteiro seguro, não negativo, menor que 10.000 e não duplicado. ParentTree e demais vínculos continuam obrigatórios. Inteiros com sinal/zeros à esquerda, inclusive -0 (valor zero), mantidos; reais com ponto, arredondamento ou underflow recusados. Não alterada a interpretação global de números pela pdf-lib.
- Novos 35 casos com tokens literais nos streams, conferidos após serialização para excluir normalização prévia: antes 27 aprovados e oito falhas reproduzidas; após correção 35 aprovados. Direcionados de recorte: 175 aprovados. Suíte PDF atual: **513 testes em 39 arquivos, todos aprovados, sem falhas/cancelados/ignorados**, incluindo as regressões dos 478 anteriores, PDF-001, PDF-008 e concorrência. Contagens direcionadas são subconjuntos, não somar à suíte.
- CASO-008 conserva 99 elementos/47 MCIDs; CASO-009, 343/204. Novo recorte com margens 10 mm/exportação/reabertura em memória aprovado; manifestos e bytes de streams conferidos. SHA256 das duas saídas idênticos aos resultados Node anteriores, sem criar outros PDFs nem sobrescrever evidências. Os 24 originais e 24 cópias continuam íntegros por SHA256. Comparação Chrome anterior permanece evidência do mesmo resultado binário, sem nova execução visual nesta correção.
- Alterados nesta correção somente pdf-crop-structure.js, tests/pdf-crop-structure.test.mjs e os três documentos de continuidade. Evidências novas ignoradas em .desktop-test/pdf004-mcid-lexical/{antes.tap,depois.tap,direcionados.tap,regressao-pdf.tap,reais.json}. Código das demais proteções preservado, experimental.39/dependências sem mudanças. Nenhum commit/push, instalação, impressão ou instalador.
- Limites continuam cooperativos, sem isolamento de memória ou prazo rígido. Mantidas as pendências de leitor independente, ferramenta assistiva e aplicativo instalado; não certificar PDF/UA nem homologação de produção. PDF-004, PDF-001 e PDF-008 permanecem Em testes. Aguardar revisão antes do checkpoint. Registros abaixo são históricos.

## PDF-004 — 13/11/2026, recorte marcado limitado, Em testes

- Base confirmada: main/a9071acf523b939ccfd5f807dcb427fe6b2bddde, árvore limpa no início, dois commits locais à frente de origin/main/118032e. a9071ac é o checkpoint local autorizado da PDF-008, incluindo concorrência da união; os registros abaixo descrevem etapas anteriores. Nenhum push efetuado nesta etapa.
- Autorização específica para compatibilidade de recorte com estrutura marcada verificável. Novo diagnóstico limitado a uma página, ParentTree Nums, elementos indiretos e MCIDs no Contents; valida referências, pais, página, índices, vínculos reversos, duplicações e ciclos, com limites de recursos. Não repara nem remove marcação. Suporte independente de nome/produtor.
- Cópia integral e recorte por caixas mantidos; manifesto de objetos/streams antes/depois da preparação e da exportação. RoleMap, MarkInfo, idioma, ActualText, Alt e Tabs preservados. Bloqueios da PDF-008 e validação de Tabs da PDF-001 mantidos. Aviso de conteúdo oculto ainda extraível/acessível e ausência de certificação PDF/UA ampliado; sangria continua exigindo aceite quando aplicável.
- Validação final: 478 testes PDF em 39 arquivos, todos aprovados, sem ignorados; inclui 56 novos casos e as regressões de assinaturas, Tabs e concorrência. Chrome: 12 viewports existentes, botão integrado e sintético marcado + CASO-008/009 com cancelar/aplicar/exportar/reabrir e aviso. CASO-008 conserva 99 elementos/47 MCIDs; CASO-009 conserva 343/204. Streams/manifestos conferidos; regiões renderizadas com zero diferenças. Hashes dos originais e cópias preservados.
- Evidências locais ignoradas: .desktop-test/pdf004-recorte/ (relatórios Node, PDFs derivados, imagens antes/região original/depois e JSON em chrome/). Não incorporar biblioteca ou dados reais ao Git. Nenhuma instalação, impressão, dependência ou mudança da experimental.39.
- Limitações: somente perfil estrutural restrito; múltiplas páginas marcadas, MCR/OBJR/XObjects, atributos de layout e demais estruturas não reconhecidas ficam bloqueados. O conteúdo oculto não é apagado nem filtrado da leitura lógica. Comparação no mesmo PDF.js e inspeção das imagens não homologam leitor independente, PDF/UA, tecnologia assistiva, cor/RIP ou aplicativo instalado.
- Próximo passo: revisão do usuário e homologação complementar; checkpoint requer nova autorização. PDF-004 nesta etapa, PDF-001 e PDF-008 permanecem Em testes, sem mudança de prioridades. Alterações desta etapa ainda sem staging/commit.

## PDF-008 — 13/11/2026, concorrência da união corrigida, Em testes

- Autorização adicional limitada à publicação de união obsoleta, testes e comentário do roteiro de assinaturas. Base permanece main/f910315, um commit local à frente de origin/main; alterações anteriores da PDF-008 preservadas, sem commit ou push.
- Reprodução permanente criada antes da correção: união suspensa na serialização, nova abertura comum/assinada ou edição, depois liberação. Antes: oito testes, um aprovado e sete falhas. Depois: geração monotônica por operação, identidade do PDF e revisão verificadas antes da publicação; erros tardios são classificados como PDF_UNIAO_OBSOLETA. A interface não restaura snapshot antigo nem publica sucesso quando a operação perdeu validade.
- Validação atual: 141 testes direcionados aprovados; depois, **422 testes PDF em 38 arquivos aprovados**, sem falhas/ignorados. Os 141 são subconjunto, não somar aos 422. Incluídos oito casos de concorrência e três de interface; regressões de assinaturas PDF-008 e recorte Tabs PDF-001 preservadas. Chrome de assinaturas e conferência progressiva aprovados. Cinco arquivos JS/MJS desta etapa com sintaxe aprovada e git diff --check aprovado.
- Relatórios locais em .desktop-test/pdf008-concorrencia/antes.tap, depois-direcionados.tap e regressao-pdf.tap. Biblioteca de PDFs reais não alterada; nenhum instalador, dependência, impressão ou mudança da experimental.39.
- Limites: geração da união não é um novo mecanismo global de cancelamento. Preparações obsoletas podem terminar em memória, mas não publicam resultados; temporários são locais e liberados pelo coletor, sem destruir documentos/tarefas alheios. Outras operações assíncronas fora da união não foram refatoradas. Electron/app instalado e homologação operacional continuam pendentes.
- Aguardar nova revisão antes do checkpoint. PDF-001 e PDF-008 permanecem Em testes; prioridades e planejamento preservados.

## PDF-008 — 13/11/2026, implementação local Em testes

- Base: main/f910315 (checkpoint PDF-001), um commit local à frente de origin/main/118032e. Checkout limpo no início. PDF-001 permanece Em testes; seu commit não foi alterado nem enviado.
- Usuário autorizou PDF-008: política conservadora central no detector existente, aplicada antes das operações de alteração/reescrita. Abrir/visualizar assinados usa cópia dos bytes originais sem save(); somente documentos protegidos mantêm esse cache extra. Campo vazio não é assinatura aplicada; restrições de formulários permanecem.
- Cobertura: recorte e adoção de cópia preparada; união no workspace/rascunho/worker; giro, exclusão, reordenação, movimento, duplicação; redimensionamento; Step & Repeat, N-up, livreto; exportação PDF/extração legada; preparos de impressão/vetorial/imagem. O antigo Resolver não remove mais assinatura aplicada ou inconclusiva. Interface de Step & Repeat/N-up não alterada.
- Regressões Node e Chrome executadas com dependências existentes. Relatórios desta tarefa em .desktop-test/pdf008; detalhes, contagens finais e limitações no registro PDF-008 de ESTADO_PROJETO.md. Biblioteca anterior: 24 originais e cópias/11 resultados conferidos por SHA256, sem alterações; leitura dos seis sintéticos confirma cinco bloqueados para reescrita e um campo vazio tratado separadamente.
- Limitações: detecção estrutural, sem verificação de confiança/certificados; não se afirma segurança absoluta ou que um PDF nunca foi assinado. Electron/instalado e documentos arbitrários ainda não homologados. Nada impresso, instalado ou empacotado; experimental.39 inalterada.
- Alterações ainda sem commit; não houve push, instalação de dependências nem mudança de prioridade de outra tarefa. Próximo passo: revisão do usuário e autorização específica para checkpoint. Registros abaixo são históricos.

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
