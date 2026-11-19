# Preflight .42 — filtro e armazenamento (17/11/2026)

## 19/11/2026 — Encerramento: atualização concluída e escopo de tamanhos homologado

Operador confirmou funcionamento prático da experimental.42 e sucesso da atualização .41→.42, aprovando o gerenciamento geral de tamanhos no escopo planejado. Esta conclusão sucede as fases históricas de preflight/build/instalação registradas abaixo; não significa estabilidade integral do aplicativo.

Instalação real auditada em 18/11: única por usuário, mesma identidade/destino, carga de 19.512 arquivos idêntica ao payload autorizado e transação committed sem pendências. Os 76 arquivos do perfil estavam iguais ao backup antes da primeira abertura. Após testes supervisionados e remoção somente do cadastro sintético, leitura final em clone privado confirmou exatamente F4 328×478, F6 318×318 e F8 238×328 mm, mesmos IDs e registros completos do backup, demais preferências iguais e ZIP de recuperação íntegro. Perfil após uso normal com 78 arquivos, estável durante a leitura; marcador catalogoCompleto=true esperado pelo CRUD .42, sem reset. Não houve restauração automática nem redefinição para nove tamanhos.

Operador confirmou gerenciador, cadastro/edição/exclusão sintética, sincronização Redimensionar/Step & Repeat/N-up, persistência após reabertura, orientação da folha e Girar arte independentes e abertura de Ver detalhes. PDF descartável do operador foi visualizado; fixtures próprias não foram abertas manualmente. Último espaço livre observado em 18/11: **6.109.941.760 bytes (5,690 GiB)**, sem nova medição na etapa documental de 19/11. Nenhum teste funcional foi reexecutado hoje.

Commit-base exato do instalador: **4509c59baaef2f189b8f6959e0d799c85e909861**. Setup autorizado: **515070952 bytes**, SHA-256 **7AE0D35B919714481E8E64AA0E304038DAD0A6548E1A1944C974CF0ACBDEBC1C**. Backup, ZIP, inventários, PDFs e recibos completos permanecem fora do Git; referências técnicas neste relatório não incorporam seus conteúdos. Setup .41 e evidências de recuperação preservados conforme auditoria anterior, sem nova operação de recuperação.

Livreto e PDF-011 continuam pendentes; demais funcionalidades/compatibilidades não recebem aprovação geral por esta homologação. Redefinição dos nove tamanhos continua separada e depende de autorização específica. Autorizado agora checkpoint apenas documental e push origin/main após verificações, sem release/tag/publicação do instalador e sem modificar instalação ou perfis. Resultado efetivo de Git será informado após a execução.

## 18/11/2026 — Auditoria da instalação real .42 aprovada, antes da primeira abertura

Operador informou conclusão do assistente; processos Setup/.42/Office ausentes antes da auditoria. Registro único HKCU e ASAR em 1.0.0-experimental.42, identidade/GUID/destino por usuário mantidos. Carga integral19512 arquivos igual por SHA-256 ao payload extraído/auditado do Setup autorizado; total19516 arquivos com quatro auxiliares esperados. ASAR instalado SHA6675f23020d0fafbf3408ea11505583f94f226f6227fefb21a6dbb65b0281b5b. Nova transação única .4ps-upgrade-Mwsffu committed, previous.41/identity.version.42, sem failure/recoveryError ou pendência.

Perfil76 arquivos exatamente igual por nomes/hashes ao backup atualizado antes da primeira abertura, preservando o catálogo e F4 328×478, F6 318×318, F8 238×328 mm. Setup .41 e .42 relidos sem alteração. Livre6.115.635.200 bytes (~5,70GiB), sem limpeza de pastas antigas/rollback/quarentena. Auditor inicial falhou por decodificar inventário UTF-16LE como UTF-8; leitor externo ajustado para exigir BOM UTF-16LE/identidade correta, falha conservada em auditoria-pos-instalacao42-primeira-leitura-utf8.json e comparação integral repetida/aprovada. Nenhuma alteração de produto ou perfil para aprovar o resultado.

Aplicação instalada aberta normalmente apenas após a auditoria: Início real lido em app://4printshops. Captura/entrada da ferramenta continuam sem confiabilidade, apesar de leitura UIA restaurada. Pedido ao operador para abrir Arquivo→Gerenciar tamanhos e confirmar os três existentes, sem editá-los. Demais testes supervisionados/homologação ainda pendentes. Fixtures sintéticas válidas e inválida para mensagens criadas na pasta privada de supervisão, sem arquivos de cliente ou impressão. Sem reset/commit/push/novas funcionalidades. Recibos completos em pre-atualizacao42-20261118-bb06eeb8/supervisionada42-20261118, fora do Git.

## 18/11/2026 — Confirmação explícita recebida e Instalar acionado pelo operador

Tela de destino lida: C:\Users\SAGRADA02\AppData\Local\Programs\4PrintShops Hold Test, mesma instalação por usuário, versão experimental.42, botão Instalar. Operador confirmou explicitamente início nesse destino e depois informou que acionou Instalar. UIA confirmou Instalando, botões desativados, avanço4711→4836 e nenhum erro visível. A captura visual permanecia indisponível; posteriormente list_windows também excedeu prazo. Recuperação do controle (espera2s/repetição/reset do kernel JS/inicialização/repetição) não recuperou leitura leve; erro da ferramenta não comprova falha nem conclusão do Setup. Solicitado estado atual ao operador, com aplicação fechada após finalizar assistente.

Durante esse período nenhum diário/arquivo de transação foi consultado, nenhuma segunda instalação foi executada e nenhum processo foi encerrado à força. Só registros da interface e preparação do auditor pós-instalação na pasta privada. Resultado da atualização, espaço final, perfil pós-instalação e testes funcionais ainda pendentes. Não restaurar/resetar/imprimir/criar commit/push. Após conclusão real confirmada, auditoria antecede abertura do aplicativo e homologação prática permanece do operador.

## 18/11/2026 — Reconferência e abertura supervisionada do Setup, sem iniciar instalação

Setup .42 relido com SHA-256 7AE0D35B919714481E8E64AA0E304038DAD0A6548E1A1944C974CF0ACBDEBC1C; .41 atual com 19.515 estáticos íntegros, Setup .41 preservado, perfil e backup atual76 arquivos iguais por SHA. Antes da abertura: nenhum processo conflitante/diário pendente, nove históricos finais, identidade/destino HKCU por usuário confirmados. Livre10.344.693.760 bytes (~9,63GiB), reserva6.618.191.166 bytes, margem3.726.502.594 bytes (~3,47GiB). Recibo privado supervisionada42-20261118/antes-setup.json.

Aberto somente Setup autorizado, PID inicial12420. Leitura textual confirma experimental.42 e a tela Escolha uma opção de instalação; Apenas para mim (SAGRADA02) disponível, todos usuários desabilitado. Tentativa de captura/recuperação retornou FrameArrived/window capture timed out; clique por índice recusado por coordinate input geometry is unavailable. Tentativa Tab sem mudança confirmável de foco. Não foram acionados Instalar nem outra execução do Setup. Solicitado ao operador selecionar a modalidade por usuário e avançar apenas até a tela do destino, sem instalar; destino final ainda precisa ser observado e confirmado explicitamente antes de Instalar. Após a abertura não foram consultados arquivos de transação. Nenhum teste pós-instalação executado. Evidências somente na pasta privada; sem restauração/reset/impressão/commit/push. Não declarar .42 instalada/aprovada.

## 18/11/2026 — Preflight final .41→.42 e backup aprovados, sem executar Setup

Etapa autorizada exclusivamente para inspeção e backup. **Podemos iniciar a instalação supervisionada somente após confirmação explícita do operador e nova conferência de aplicativos fechados. Nenhum Setup foi executado, aplicativo instalado/desinstalado, catálogo redefinido, arquivo de produção alterado, impressão, commit ou push.**

### Classificação das 19 diferenças de perfil

| Arquivo | Classe |
| --- | --- |
| DawnGraphiteCache\data_1 | cache |
| DawnGraphiteCache\index | cache |
| DawnWebGPUCache\data_1 | cache |
| DawnWebGPUCache\index | cache |
| DIPS | chromium-state |
| Extension State\LOG | log |
| Extension State\LOG.old | log |
| GPUCache\data_1 | cache |
| GPUCache\index | cache |
| GrShaderCache\data_1 | cache |
| GrShaderCache\index | cache |
| Local Storage\leveldb\000003.log | operational-data |
| Local Storage\leveldb\LOG | log |
| Local Storage\leveldb\LOG.old | log |
| Session Storage\000003.log | session-data |
| Session Storage\LOG | log |
| Session Storage\LOG.old | log |
| ShaderCache\data_1 | cache |
| ShaderCache\index | cache |

Resumo: **10 caches gráficos**, **6 logs diagnósticos LevelDB** (LOG/LOG.old), **1 banco interno Chromium DIPS**, **1 journal de Local Storage com dados operacionais** e **1 journal de Session Storage**. Os arquivos numerados .log não foram tratados como simples logs descartáveis. Causa temporal individual não comprovada pelo uso relatado do operador; não houve restauração baseada em hash diferente. As 21 diferenças históricas continuam preservadas nos registros anteriores.

Os journals numerados e manifestos LevelDB passaram na conferência de estrutura, fragmentação e CRC32C: Local Storage 198 registros e Session Storage 95 registros, mais um registro em cada manifesto. DIPS passou em PRAGMA integrity_check em cópia SQLite aberta read-only. Arquivo Preferences é JSON válido e mantém o SHA-256 do baseline da recuperação.

### Dados operacionais

Leitura sem executar código do aplicativo, em clone privado de Local Storage via Electron genérico de hash verificado, com userData/sessionData explícitos, origem app://4printshops, sandbox/contextIsolation e rede externa bloqueada. Backup imutável e perfil original não foram abertos pelo Chromium. Ambos reconferidos por SHA-256 após a leitura; 76 arquivos iguais ao snapshot inicial desta etapa. O validador de catálogo recebeu adaptador somente de leitura, sem acesso ao storage original nem chamadas de redefinição.

**Catálogo v1 válido com exatamente três registros; texto completo byte a byte igual ao catálogo do backup histórico verificado anterior à .41**, SHA-256 dc30457bd407d6d32a5dca923d0b1e23e0f7424334f842119ad7748a6a5f0180. IDs, nomes, dimensões e demais campos mantidos. **F4 328×478 mm, F6 318×318 mm e F8 238×328 mm.** Nenhum preenchimento automático para nove tamanhos. Tema light/claro, painéis esquerda 320/direita 230 e marcador histórico print-config-check=preservado conferidos; nenhuma outra chave de aplicação inesperada encontrada. Nenhuma alteração operacional suspeita detectada nas verificações efetuadas; isso não atribui causa individual aos caches/logs que mudaram.

### Backup atualizado e recuperação

Backup privado local, fora do Git: **C:/Users/SAGRADA02/AppData/Local/4PrintShops-Backups/pre-atualizacao42-20261118-bb06eeb8**. Perfil: **76 arquivos/8463943 bytes**, copiados e comparados por tamanho/SHA-256, com origem reconferida antes/depois e após o leitor do clone. Manifesto perfil-inventario-privado.json; ZIP perfil-operacional.zip, **214751 bytes**, aberto integralmente com comparação dos 76 conteúdos, sem extras/duplicatas/ausentes.

**SHA-256 do ZIP: D73F6C73CAEB1B9234E2CADEB4CE4F4F5882FE880C28CC3E07FB9F6CE349054F.** SHA-256 do manifesto: CABAB203B3ABB59C78D893700F13AABEC09E316CDB2158229F7BAB484628A078. Catálogo bruto, preferências e recibos completos somente na pasta privada. Leitura-clone é descartável e distinta do perfil preservado e do ZIP; nenhum clone é fonte autoritativa de recuperação.

Conservar ZIP/manifesto/recibos, Setup .41 e diários/quarentenas históricos. Se a futura atualização falhar antes do commit, permitir a recuperação transacional e preservar evidências. Em recovery-required ou rollback incompleto, parar para revisão, sem repetir Setup ou executar reparos automaticamente. Após commit, a guarda de versão não permite usar .41 diretamente como upgrade sobre .42: eventual retorno exige autorização separada e procedimento Windows de desinstalação/reinstalação com Setup .41 preservado, e restauração de perfil/chave somente se necessária, revisada e autorizada. Nenhuma recuperação foi executada agora.

### Aplicativo, identidade, instaladores e espaço

.41 instalada confirmada por versão interna/registro único HKCU, GUID 90ff2ff0-f013-571c-b109-7edb265dacc8; **19.515 arquivos estáticos idênticos ao inventário**, debug.log dinâmico explicitamente excluído da comparação estática. Funcionalidade apoiada no uso operacional confirmado e integridade do pacote; **não foi iniciada interface normal para novo smoke test**, pois esta etapa exigiu aplicativo fechado. Destino existente: C:\Users\SAGRADA02\AppData\Local\Programs\4PrintShops Hold Test. UninstallString aponta a esse destino com /currentuser; nenhuma entrada equivalente de instalação por máquina. Identidade/appId/produto/exe/atalhos/cache iguais entre .41 e .42, somente versão alterada.

Setup .42 conferido novamente, **515070952 bytes**, SHA-256 **7AE0D35B919714481E8E64AA0E304038DAD0A6548E1A1944C974CF0ACBDEBC1C**, commit-base 4509c59baaef2f189b8f6959e0d799c85e909861; não assinado, conforme auditoria existente. Setup .41 local disponível e relido com SHA-256 **253D672DDFB8C02346D540E767450A4D800395309AAB0F5655A55F9BDE7D16D3**. Não executados para esta conferência.

Instalação atual mede 1977818283 bytes. Reserva conservadora livre para atualização/temporários/rollback: 2×maior(instalação atual,carga nova)+Setup novo+2 GiB = **6618191166 bytes (6,164 GiB)**. Exigida margem adicional de pelo menos 2 GiB. Livre final **10518503424 bytes (9,796 GiB)**, margem além da reserva **3900312258 bytes (3,632 GiB)**; suficiente nesta medição. Não eliminar arquivos automaticamente e reconferir se a instalação for adiada.

Consulta final: aplicativos/conversores/Setup/MSI ausentes, serviço MSI parado, chave MSI InProgress ausente. Um msiexec de serviço observado no início terminou naturalmente; não foi encerrado por esta operação. Nove diários históricos committed/rolled-back, nenhum pendente. Sem transação preparada por este preflight.

### Handoff

Recibo completo: C:/Users/SAGRADA02/AppData/Local/4PrintShops-Backups/pre-atualizacao42-20261118-bb06eeb8/preflight-final-aprovado.json. Aguardar confirmação do operador antes de abrir o Setup .42. Na futura operação autorizada, manter identidade/destino/perfil atuais, sem reset do catálogo; depois conferir versão, três ferramentas e persistência dos formatos/preferências em reabertura. Preservar perfil atual e evidências, sem restauração automática. Os quatro documentos de continuidade receberam apenas resultados de preflight; sem commit/push.


## 18/11/2026 — Relato de uso operacional no turno das 13h

Após a auditoria, o responsável informou uso da .41 no início do turno das 13h, sem lembrar os minutos exatos. Isso confirma uso operacional no período e é compatível com as gravações observadas entre 13:31 e 13:39, sem comprovar correspondência exata ou causa de cada diferença. Permanecem registradas as 19 diferenças frente ao baseline desta recuperação e as 21 históricas; não concluir alteração dos formatos pelos hashes nem restaurar/redefinir dados. Nenhuma validação foi repetida ou proteção dispensada. Setup .42, SHA-256 e auditoria permanecem inalterados; aguardar revisão e autorização específica antes de instalar.

## 18/11/2026 — Setup experimental.42 gerado e auditado, não instalado

**Pacote gerado e auditado, sem instalação/desinstalação, impressão, publicação ou push.** Binários da .41 preservados. A conferência de igualdade do perfil operacional falhou: 19 diferenças novas frente ao baseline desta recuperação. Nenhuma redefinição/restauração foi feita; as 21 diferenças históricas frente ao snapshot da limpeza continuam registradas. Não declarar preservação integral do perfil. A divergência encerrou ações além de diagnóstico somente de leitura e documentação.

### Causa e correção do ambiente

O checkout ativo continha somente 2.212 arquivos históricos de dependências versionados; o ambiente completo de testes era separado. A preparação .42 anterior ainda não tinha source/node_modules, e os testes transacionais foram invocados diretamente no checkout incompleto. builder-util-runtime 9.7.0 e @electron/asar 3.4.1 são transitivos previstos no lockfile (updater/builder), não pacotes novos arbitrários. O preparo de 16/11 já registrava 7.693 arquivos adicionais; nenhum alvo autorizado da limpeza incluía o checkout ativo ou a preparação preservada. Portanto a causa foi o uso de ambiente incompleto, não a exclusão desses pacotes pela limpeza.

Restaurados na cópia isolada os 9.905 arquivos do npm ci anteriormente preservado, comparados antes e depois com inventário SHA-256 independente anterior à limpeza. Manifesto/lockfile idênticos entre checkout, deps e preparos .40/.41; lock com 313 pacotes, origem registry.npmjs.org e integridades registradas. Nenhum npm install, resolução nova, import de produção, manifesto, lockfile ou node_modules ativo foi alterado. Node 26.13.1 para suítes PDF; Electron 45.1.0, electron-builder/app-builder-lib 26.18.0, electron-updater 6.8.11, ASAR 3.4.1, builder-util-runtime 9.7.0, pdfjs-dist 6.4.299, pdf-lib 1.17.1. Recibo audit/recovered-environment.json.

### Checkpoint e proveniência exata

- Checkpoints aprovados anteriores: 4bec43c330245e1ec5c534c8c55f125dd268a309 (filtro/teste) e 49b425a35369be9473af24b20699a2b4013aa25a (documentação).
- Novo checkpoint autorizado: **4509c59baaef2f189b8f6959e0d799c85e909861**, Prepara experimental.42 e registra recuperacao das validacoes isoladas. Somente electron-builder.hold-native-test.cjs e quatro documentos de continuidade revisados. Única alteração de produto/configuração: extraMetadata.version .41→1.0.0-experimental.42.
- **Commit-base exato do instalador: 4509c59baaef2f189b8f6959e0d799c85e909861**. Árvore limpa confirmada antes da fonte final e novamente imediatamente antes do build. 2.657 arquivos rastreados conferidos no preparo; os 445 fora de node_modules foram escritos diretamente dos blobs desse commit. Dependências completas têm os hashes preservados. 395 diferenças de quebra de linha contra o checkout foram verificadas sem diferença de conteúdo.
- Lock de trabalho CRLF: 0bcfa9c9d88373ae7f2e86c96f3ec4f93857b05107704dd834b4a8f016e0c0a9; blob Git LF exportado: 28d22c2e7ee4c4d0f21bfe639f08521bfc2a01e240b0a947ff438387812fd17d. Igualdade integral do texto após normalizar somente CRLF/LF comprovada. Um primeiro preflight interrompeu antes da cópia por comparar hash bruto desses dois formatos; o preparador passou a exigir hash conhecido do checkout e equivalência integral normalizada, sem dispensar verificação nem alterar lock. Ocorrência em audit/git-export-line-endings.json.
- Escopo: catálogo único/gerenciador/nove sugestões editáveis e excluíveis/backup e manutenção controlada, preservando .41 e suas proteções. Sem PDF-010/PDF-011/abas/Livreto/atualizador pelo servidor. Nenhum reset automático na inicialização/instalação; catálogo existente inclusive vazio preservado.

### Testes desta recuperação

| Grupo | Resultado novo | Limites |
| --- | --- | --- |
| Quatro arquivos antes não carregados | **56/56**, zero ignorados | assinatura/updates/template/transação; incluídos novamente na suíte completa seguinte |
| Suíte completa de atualização, 11 arquivos | **100/100**, zero ignorados | EPERM/temporários, quarentena LibreOffice, processo Office simulado, handle Windows real, rollback/recuperação, caminhos e registro de laboratório |
| INST-VAL .41→.42 | **7/7**, zero ignorados | sucesso preserva resíduo/hash/catálogo sintéticos; rollback restaura .41; interrupção e erro inesperado recuperáveis; instalação limpa e reinstalação sintéticas; EPERM Windows .41→.42 |
| Electron isolado | **Aprovado**, quatro processos e três reaberturas | PIDs 16676/17360/20028/21472; saídas normais código 0 em 185/182/204/157 ms; zero erros |
| Funcionais ampliados | **141/141 anteriores reutilizados** | não reexecutados nesta recuperação; fontes funcionais inalteradas, versão/documentos apenas |

Os 56 não são casos adicionais aos 100. Não somar execuções repetidas como testes únicos. Não foi repetida a suíte PDF completa nem impressão física. Os sete INST-VAL não executam o Setup Windows nem atualizam a instalação operacional. O controlador externo teve uma referência de variável corrigida antes de iniciar os sete casos; nenhum teste/proteção foi desabilitado ou alterado. Logs antigos de falha preservados, separados dos recovered-*.log aprovados.

Electron: origem app://4printshops, sandbox/contextIsolation e nodeIntegration=false, cópia atual por hash e perfil exclusivo. Conferidos Arquivo → Gerenciar tamanhos, nove sugestões em primeiro uso, três ferramentas sincronizadas, CRUD/IDs/metadados, exclusões definitivas, catálogo vazio após reinício, backup/redefinição/restauração sintéticos com JSON salvo/relido/SHA-256, orientação de folha e Girar arte independentes. Todos os fechamentos foram normais, sem kill como prova de sucesso. Evidência 4ps-size-manager-dszkmh em TEMP; resultado em recovered-electron.log.

### Recursos, armazenamento e build

LibreOffice 26.8.2.2: 19410 arquivos, manifesto 01f0286083b702da3c085ec7219c22b98db2d86aa8ae3f6c3633f6707537247e, fontes correspondentes e MSI verificados. PDFium arquivo fixado 9a964261760cff52570aabf038ab346d3a5034dd8b30d0d5fe8861235d356b8b; DLL distribuída **4b0dec92555d69437bef0df37be630181ab3e83a6823ca264ede355a4f2ef994**, idêntica à extração do arquivo fixado. NSIS/7-Zip/rcedit conferidos por hashes anteriores; template NSIS aprovado. Nove diários operacionais committed/rolled-back, nenhuma transação ou processo conflitante antes do build. Nenhuma exclusão adicional.

Livre após preparo e antes do build: **17682845696 bytes (16,468 GiB)**; reserva estimada adicional de 8 GiB para build/auditoria/temporários, com margem. Geração completa histórica estimada em 9–12 GiB desde antes da nova cópia. Builder 26.18.0, configuração electron-builder.hold-native-test.cjs, Windows x64/NSIS, publish never, npmRebuild=false, Electron preparado fixado; fontes C# recompiladas nesta geração. TEMP/TMP e APP_BUILDER_TMP_DIR restritos ao processo e à pasta exclusiva .42, cache de ferramentas preservado, sem alterar ambiente global. Comando/saída em audit/build-command.json, build42.log e build-exit.json.

### Auditoria do pacote

- ASAR: **411 arquivos**, **116 arquivos desktop/public idênticos à fonte do commit**; identidade, manifestos e fuses aprovados. pdf-format-backup.js, pdf-format-manager.js e grafo do catálogo incluídos; **pdfs backup.js excluído**. Menu/guardas e código transacional/quarentena presentes.
- appId br.com.4printshops.holdtest; GUID 90ff2ff0-f013-571c-b109-7edb265dacc8; nome 4printshops-hold-test, produto/exe/atalho 4PrintShops Hold Test; cache/perfil/identidade iguais à .41, somente versão alterada. perMachine=false, allowElevation=false, deleteAppDataOnUninstall=false, runAfterFinish=false, publish=null.
- Carga integral do Setup: **19512 arquivos iguais por SHA-256 ao win-unpacked**, sem ausentes/extras. CRC do 7z delimitado e comparação integral aprovados. Aviso conhecido de cauda do contêiner NSIS registrado; não é ignorado nem usado como aprovação isolada. Extração somente na auditoria privada, sem instalar.
- LibreOffice/PDFium e auxiliares conferidos com manifestos/pins; fixture PDF externa ao fornecedor limitada ao sintético autorizado. Nenhuma biblioteca de clientes, perfil, sessão, credencial, log privado, teste/documento de auditoria ou canal app-update.yml encontrado no distribuível. Ghostscript não configurado, como antes; recursos existentes preservados.
- Exe novo do pacote executado somente com --verify-distribution e TEMP/userData exclusivos: código 0, verified=true, 19410 arquivos, inicialização privada do LibreOffice aprovada. Não abriu aplicativo operacional nem perfil real. Versão FileVersion do Setup/exe e versão interna ASAR confirmadas como 1.0.0-experimental.42. Authenticode **NotSigned**, como nas experimentais anteriores; não publicar automaticamente.

### Artefato e preservação final

- Caminho: C:\Dev\4PrintShops-Builds\experimental.42-20261118-45a8d3bb\source\dist-desktop\hold-native-test\4PrintShops-Hold-Test-1.0.0-experimental.42-Windows-x64-Setup.exe
- Tamanho: **515070952 bytes**.
- **SHA-256: 7AE0D35B919714481E8E64AA0E304038DAD0A6548E1A1944C974CF0ACBDEBC1C**. Sidecar .sha256 ao lado.
- Livre final observado após geração/auditoria/provas: **12130627584 bytes (11,298 GiB)**, sujeito à atividade normal do volume. Sem limpeza adicional.
- .41 instalada: **19515 arquivos estáticos iguais ao inventário**, exe/ASAR preservados e Setup .41 com hash anterior. **Perfil: 57/76 arquivos iguais e 19 diferenças novas** frente ao baseline desta recuperação, em caches/logs/armazenamento Chromium. Gravações novas entre 13:31 e 13:39, depois dos testes Electron isolados (13:28) e antes do build (13:41). Causa não confirmada; não afirmar que os formatos foram alterados nem que permaneceram iguais apenas com esses hashes. O assert de igualdade falhou e foi mantido como ocorrência, sem restaurar dados, dispensar proteções ou modificar testes. As 21 diferenças históricas continuam registradas. Evidência privada somente de leitura em operational-after-recovery.json. Lockfile ativo inalterado. Git fsck/diff-check aprovados, sem corrupção; árvores históricas sem referência preservadas.
- Git no build: main/4509c59, ahead 15 da referência local origin/main, árvore limpa. Após esta auditoria, somente quatro documentos receberam registros finais e ficam para revisão, sem commit adicional automático. Nenhum push/release/instalação/desinstalação/impressão/reset real.

### Pendências para atualização supervisionada

Revisar este relatório, esclarecer as 19 diferenças novas do perfil e autorizar separadamente a instalação. A igualdade integral do perfil não foi aprovada; não apresentar o resultado como aprovação integral da preservação. A divergência foi detectada após a geração/auditoria do pacote. Antes da atualização: renovar backup privado do perfil/catalogo atual, reler SHA/versão do Setup .42, conferir espaço, formatos/preferências existentes, apps/Office fechados e diários finais. Atualizar uma única instalação no destino/identidade existentes, sem apagar perfil, lock ou quarentena. Não redefinir tamanhos automaticamente. Após futura instalação autorizada, conferir versão/abertura sintética, gerenciador/três ferramentas, persistência de catálogo/orientação após reabertura, sem reset e sem impressão automática. Homologação operacional continua posterior. Preservar Setup .41 para retorno autorizado; Setups .29/.36 seguem indisponíveis, sem garantia de reconstrução.


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
