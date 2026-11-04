# Hold Test — correções .11 e histórico .10

## Resultado físico confirmado pelo usuário

Posteriormente o usuário testou e aprovou fisicamente a experimental.11 na C6100/IC-604, confirmando escala, envios consecutivos, Hold e preservação de todas as propriedades configuradas. A pendência sobre propriedades nesse cenário foi resolvida pela conferência física. A .12 conserva esse diagnóstico especializado como opção e acrescenta um fluxo geral independente de Hold.

Em 04/11/2026, o usuário confirmou o envio direto do sintético à AccurioPress C6100/IC-604 e seu armazenamento em Hold, sem início de impressão automática. É o primeiro cenário físico positivo. Foram identificadas restrições de escala fixa em 120%, geometria excessivamente bloqueada e tentativa única por execução. O checkpoint está em `0af565f`. A continuação .11 preserva a geração, verificação e entrega do mesmo PS como RAW.

## Experimental.11 — escala, corte e novas operações

Continuação sobre `0af565f`, sem nova auditoria dos 48 arquivos e sem commit. Foram preservados PDFium/HDC, o DEVMODE completo, a interpretação obrigatória do PS e a entrega RAW dos mesmos bytes congelados. O laço de escrita na fila de produção não foi reescrito. Durante esta etapa, todas as capturas ficaram exclusivamente na fila FILE: e as confirmações finais foram canceladas.

### Correções realizadas

- A .10 forçava 120% na UI e no auxiliar e exigia A3. A .11 inicia em 100%, oferece 120% e percentual personalizado entre 1% e 400%, com até duas casas decimais e vírgula. O valor selecionado segue pelo workspace/IPC até as dimensões usadas no desenho PDFium. O zoom público e o zoom efetivo do PS continuam obrigatoriamente em 100%.
- O original permanece sintético A4, 210 × 297 mm. A interface separa papel lógico do driver, folha física retornada pelo contexto, escala do aplicativo e zoom do driver. A folha física é confrontada com o PS interpretado. Neste laboratório são aceitas A4 e A3 retrato. Uma diferença de papel que não seja comprovada ou um zoom conflitante continua bloqueando o envio.
- O excesso da arte sobre o papel ou área imprimível virou `cutRisk`, sem suprimir os demais bloqueios. A prévia mostra posicionamento e excedentes. Depois da prova de Hold, o processo principal pede confirmação separada de possível corte. Mesmo aceitando o corte, ainda é necessário confirmar o envio individual. A análise usa a caixa inteira do A4 e as métricas informadas pelo driver, com a tolerância de arredondamento anterior de 0,1 mm.
- A tentativa única da .10 durava toda a execução nos três níveis, UI, processo principal e auxiliar. Na .11 ela consome somente a operação atual. **Nova operação** exige decisão explícita, libera o token anterior e cria outra sessão. A UI exige verificar Hold novamente e faz nova captura e interpretação. Isso funciona após envio, cancelamento ou resultado incerto. Não existe repetição automática. Em caso de incerteza, conferir primeiro o trabalho anterior na Konica.
- A seção recolhível de configurações PS mostra somente valores públicos recuperados da execução do arquivo. A geração e o envio não foram substituídos para tentar fazer os campos reaparecerem no painel da máquina.

### Propriedades presentes e limites da investigação

| Propriedade | Evidência no arquivo executado | Situação na IC-604 |
| --- | --- | --- |
| Hold | `SHKMOutputMethod=/Hold`, modo Normal | Validado fisicamente pelo usuário na .10 |
| Papel | `HiResPageSize` e `currentpagedevice/PageSize` conferidos em A3 e A4 | A3 confirmado no relato físico; novos casos .11 aguardam conferência |
| Orientação | `PageOrient=(Port)` no sintético retrato | Aparência e edição no painel ainda precisam de conferência |
| Escala do aplicativo | Incorporada às coordenadas vetoriais, com medidas de 100%, 120% e 87,5% registradas | O painel não recebe um campo separado de zoom do aplicativo |
| Zoom do driver | `Zoom=100` exigido na interpretação | Não representa a ampliação já incorporada à arte |
| Cor/PB | Captura observada com `SHKMColorMode=/Auto`. O PPD mapeia Cor Total para `/CMYK` e Escala de Cinza para `/Gray` | PB direto ainda exige captura escolhida no driver e teste físico; o teste anterior via Edge não homologa esse caso |
| Cópias | Estado público exige 1; interpretação exige 1 em `#copies` e `NumCopies`, usando o padrão PS quando ausente | Outras quantidades continuam bloqueadas neste laboratório |
| Bandeja | A captura automática observada não forneceu `SHKMInputSlot`; a UI informa ausência | Não presumir que uma bandeja específica chegou |
| Duplex e grampo | Na captura observada, `SHKMDuplex=false` e `SHKMStaple=/None` | Opções diferentes ainda não homologadas fisicamente |

O transporte RAW não reaplica o DEVMODE. O DEVMODE é utilizado antes, para o driver gerar o PS que será entregue. O arquivo também não carrega todas as propriedades originais do PDF como um documento PDF editável. Escala já incorporada às coordenadas e recursos transformados em contornos não se tornam parâmetros editáveis de origem no painel.

Não foi comprovada a causa de todos os campos ausentes na IC-604. O manual oficial descreve limitações de edição conforme o tipo/PDL do trabalho, incluindo trabalhos PPD. A fila FILE: possui sua própria configuração de opções instaladas. Esses pontos precisam ser confrontados com o tipo do trabalho e com os nomes exatos dos campos ausentes na máquina. Não foram modificadas configurações da controladora nem sincronizados dados privados entre filas. A documentação do transporte RAW exige que as configurações já estejam descritas na linguagem enviada. [Microsoft — WritePrinter](https://learn.microsoft.com/en-us/windows/win32/printdocs/writeprinter), [Konica — edição de Hold Job](https://manuals.konicaminolta.com.au/phpmanuals/C6100_C6085/contents/JBC_102639647.html)

### Validação direcionada

228 testes Node, 30 verificações C#, sintaxe de 12 JS/MJS/CJS e dois PowerShell, `git diff --check` e Chrome headless curto aprovados. A rodada final após os ajustes de folha física também passou. Foram feitas capturas reais somente na fila FILE: para A3/120%, A4/100%, A4/120% e A4/87,5%. Os resultados e medidas estão em `.desktop-test/hold11/prepare-result.json`. Hold e zoom 100% foram comprovados em todos os casos. A4/120% retornou risco de corte sem dispensar a prova de Hold.

O instalador atualizou a .10 para a .11 no mesmo diretório, registro NSIS e perfil. A validação instalada aprovou duas novas operações, cada uma com nova verificação Hold, confirmação geométrica separada, percentual 87,5% com vírgula e três confirmações finais canceladas. O PDF girado permaneceu pendente e foi exportado posteriormente com uma página a 180°. Preferência anterior, DEVMODEs globais/per-user e impressora padrão foram preservados. Relatórios em `.desktop-test/hold11/installed-validation/`, incluindo a evidência do primeiro upgrade .10 → .11 e a revalidação do pacote final.

Instalador final `dist-desktop/hold-native-test/4PrintShops-Hold-Test-1.0.0-experimental.11-Windows-x64-Setup.exe`, SHA256 `38bdde63b22f1eb6ee2aee0eb77c53a6d464fc6a50627438c91c37e1f1457bb3`. A auditoria do ASAR instalado aprovou 430 entradas. As versões anteriores foram preservadas.

### Roteiro físico .11

1. Abrir **4PrintShops Hold Test** e seu sintético. Selecionar a fila Konica e conferir o papel físico informado após as propriedades.
2. Em A4 retrato, escolher escala do aplicativo **100%**, zoom do driver **100%**, uma cópia, nenhuma imposição e Hold/Normal. Verificar Hold e conferir o PS recuperado. Confirmar pessoalmente um envio e conferir retenção, dimensões e aparência na IC-604.
3. Clicar **Nova operação** e confirmar conscientemente. Configurar A3 retrato, verificar Hold novamente e escolher **120%**. Repetir o envio individual e a conferência física.
4. Em outra operação, testar **Personalizada**, por exemplo **87,5%**. Conferir a medida da arte na prévia e na confirmação.
5. Em A4/120%, conferir o aviso de possível corte. Cancelar deve preservar o PDF. Se decidir continuar, confirmar também o envio final e verificar os cortes previstos na máquina.
6. Comparar os valores PS de cor, duplex, bandeja e acabamento com a IC-604. Registrar o tipo/PDL do trabalho e quais campos ficaram ausentes. Não deduzir sua aplicação apenas pelo resumo.

As verificações de Hold/Normal, vínculo da sessão, identidade do driver, cópia única e ausência de imposição continuam obrigatórias. Sintético exclusivo, A4/A3 retrato e Ghostscript local com hash fixado permanecem como limites do laboratório. Edge, exportação e demais edições continuam disponíveis.

## Histórico experimental.10

O registro abaixo descreve o estado anterior ao teste físico informado pelo usuário. As pendências de primeiro envio e reinício do aplicativo nele citadas foram atualizadas pelas seções .11 acima.

Etapa em 04/11/2026, sobre o commit `5d07e81` e as 33 alterações ainda não registradas da experimental.9. Uma cópia de segurança desses 33 arquivos foi preservada em `.desktop-test/hold-native/baseline-20261104-105153/`, ignorada pelo Git. Nenhuma alteração anterior foi descartada.

O usuário aprovou manualmente a experimental.9, incluindo confirmação de Hold na sessão e impressão física pelo Edge. Relatou que preto e branco e outros ajustes avançados escolhidos no driver foram respeitados. Esse resultado valida o fluxo externo relatado pelo usuário. O primeiro envio direto pelo 4PrintShops continua aguardando sua ação e conferência na máquina.

Conforme autorização posterior do usuário, a experimental.10 conserva a identidade da experimental.9, **4PrintShops Hold Test**, appId `br.com.4printshops.holdtest`, executável/atalhos iguais e perfil `%APPDATA%/4PrintShops Hold Test`. A versão principal permanece separada. O novo pacote atualiza o laboratório existente. A versão da raiz continua `.6`, com `.10` aplicada somente no metadado do laboratório.

## Referência e restrições

- Konica AccurioPress C6100, IC-604, fila `KONICA MINOLTA C6100PS`, driver PS V3 `KONICA MINOLTA C6100/C6085PS`, OEM 5.0.504.0.
- Sintético fornecido pelo aplicativo, uma página A4 retrato, contendo dois textos e um retângulo vetorial azul. O workspace atual é serializado, incluindo giro de 180° ainda não exportado.
- Papel lógico/de saída A3 retrato, escala do conteúdo aplicada uma vez em 120%, zoom do driver 100%, uma cópia, Hold, modo Normal e ausência de imposição/booklet.
- Print, HoldPrint, campos inconclusivos, documento diferente, alterações de sessão/revisão/driver e falhas de interpretação bloqueiam o envio.
- O ensaio é exclusivo desta fila e deste laboratório. O produto principal e a versão web permanecem com seus fluxos anteriores.

## Captura autorizada e proteção da impressora padrão

O usuário autorizou a criação da fila local **4PrintShops Konica PS Capture Lab**, usando o driver instalado, sem compartilhamento e com a porta **FILE:**. A fila de produção conservou `IP_192.168.88.100`. Não foram alterados driver, porta ou DEVMODE global da produção.

Durante os primeiros ensaios, o Windows estava com gerenciamento automático da impressora padrão ativo, `LegacyDefaultPrinterMode=0`. Após as capturas, o sistema selecionou a fila de laboratório como padrão. A consulta feita antes da primeira captura havia identificado a Konica como padrão. O efeito foi identificado na auditoria final e informado ao usuário.

O usuário autorizou explicitamente restaurar a Konica e desativar a gestão automática **somente no usuário atual**. Foi restaurada `KONICA MINOLTA C6100PS`, e `LegacyDefaultPrinterMode=1` foi aplicado exclusivamente em HKCU de `MARCOS\SAGRADA02`. A fila de laboratório foi preservada. Registro local em `.desktop-test/hold-native/default-restoration.json`. Não houve alteração em outros computadores ou preferências do driver.

`CaptureEnvironmentPolicy` consulta essa preferência, sem modificá-la. Tanto a captura Konica quanto o ensaio virtual Microsoft agora bloqueiam antes de iniciar um trabalho se o gerenciamento automático estiver ativo, ausente ou inconclusivo. Um teste real confirmou a recusa enquanto o valor ainda era 0. O aplicativo não executa SetDefaultPrinter nem altera essa preferência. O teste instalado compara também a impressora padrão antes/depois. O comportamento automático está descrito pela Microsoft. [Configuração da impressora padrão](https://support.microsoft.com/en-au/windows/hardware/printer/set-a-default-printer-in-windows)

## Caminho de geração e envio

1. O auxiliar mantém o DEVMODE integral confirmado, incluindo dados privados. A decisão Hold continua vinculada à fila, sessão, revisão, bytes completos e arquivos/porta/processador/tipo de dados do driver.
2. PDFium confere a identidade visual e os tipos dos três objetos do sintético, com ausência de anotações. A comparação em bitmap serve somente para reconhecer o documento de teste. A impressão usa o PDF original em HDC, sem imprimir a imagem dessa comparação.
3. O mesmo DEVMODE é passado ao contexto do driver Konica da fila FILE:. O buffer é conferido para detectar alterações. PDFium desenha os objetos no HDC, centralizando o A4 a 120% sobre A3. O driver produz um PS novo, com nome `4PS HOLD EXP xxxxxxxx`.
4. O PS completo é interpretado localmente antes da confirmação. No momento de showpage, a verificação exige exatamente uma página e os valores efetivos de Hold, Normal, zoom 100%, folha A3, uma cópia e N-up/booklet ausentes. O marcador por execução é aleatório. O teste não depende de localizar a palavra Hold no arquivo.
5. Os bytes PS são guardados em memória depois da conferência de integridade. A confirmação nativa mostra nome do trabalho, fila, papel, escala, cópias e Hold. Cancelar é a opção padrão.
6. Após confirmação explícita, a identidade da fila e o vínculo com o DEVMODE/revisão são conferidos novamente no handle de produção. O auxiliar encaminha **os mesmos bytes PS verificados** como um trabalho RAW, com nome igual ao gravado no PS.

A escolha RAW nesta etapa evita gerar outra representação depois da prova. As configurações são produzidas pelo driver a partir do DEVMODE completo e já estão descritas no PS. O aplicativo não monta comandos OEM nem reconstrói propriedades privadas. A Microsoft exige que dados RAW descrevam as configurações na linguagem do equipamento. Este caminho fica limitado ao PS dessa fila de referência. [WritePrinter](https://learn.microsoft.com/en-us/windows/win32/printdocs/writeprinter), [StartDocPrinter](https://learn.microsoft.com/en-us/windows/win32/printdocs/startdocprinter)

O token é consumido antes de iniciar o trabalho. Erro ou timeout produz estado de resultado incerto e bloqueia outra tentativa automática. O processo principal também conserva o bloqueio após uma tentativa. Conferir a fila e a máquina antes de reiniciar o aplicativo para qualquer novo ensaio. O workspace não é alterado nem marcado como exportado.

## Evidências de saída

Foram capturados e interpretados três controles na fila FILE:. No momento de produzir a página, os valores foram respectivamente `Print`, `Hold` e `HoldPrint`, com `SHKMPrintMode=Print` significando Normal. O prólogo proprietário usa rotinas protegidas. A execução do PS demonstrou que a rotina de configuração foi aplicada.

No caso A3/120%, a interpretação retornou Hold, zoom 100%, `HiResPageSize=[297000 420000]` e `PageSize=[842 1191]`. O contexto reportou aproximadamente 297,032 × 420,158 mm e conteúdo de 252,010 × 356,404 mm. A tolerância de 0,1 mm no enquadramento trata o arredondamento físico observado no PScript, sem liberar cortes maiores.

A renderização local foi inspecionada em 72 e 300 dpi. O retângulo esperado com 63,5 × 21,167 mm foi medido em aproximadamente 63,585 × 21,251 mm a 300 dpi, diferença de um pixel por dimensão. A amostra azul passou de RGB 51/102/204 para 51/101/204. A saída de análise continha zero imagens e zero fontes, com textos convertidos em contornos vetoriais. A aparência do sintético ficou legível, com escala e posicionamento compatíveis. Conservação de fontes pesquisáveis, CMYK/ICC, transparências complexas e fidelidade no RIP permanecem pendentes.

Essa prova valida a representação de retenção gerada pelo driver e interpretada localmente. O Ghostscript não substitui a controladora IC-604. Armazenamento real sem impressão física automática ainda precisa da conferência pessoal do usuário no primeiro envio.

## Dependências locais e arquivos gerados

Instalador final em `dist-desktop/hold-native-test/4PrintShops-Hold-Test-1.0.0-experimental.10-Windows-x64-Setup.exe`, SHA256 `3da4dbbeb43ee2950c7fc4fadef83a39c04b06ab200a9fe0f72f9159b81229ea`. O registro NSIS conservou o identificador `90ff2ff0-f013-571c-b109-7edb265dacc8`, com a mesma pasta da Hold Test .9. A identidade provisória Hold Native Test usada no início desta investigação foi desinstalada somente da pasta de teste criada nesta tarefa, preservando os perfis. A entrega utiliza exclusivamente a identidade Hold Test.

O laboratório utiliza o Ghostscript **10.01.2 já instalado** em `C:\Program Files\gs\gs10.01.2\bin\gswin64c.exe`, SHA256 `6d9721296781e9f8af1b94db878a7adf96b8bcf744b675d00b810a19a038945f`. O caminho/hash foram incorporados à variante de laboratório. Ausência ou alteração bloqueia o preparo. Nenhum Ghostscript foi instalado ou incluído no Setup. Um pacote oficial 10.08.0 foi baixado para a pasta ignorada durante a investigação, sem instalação ou uso na prova.

Essa dependência antiga está limitada a PS produzido pelo driver para o sintético reconhecido. A solução comercial ainda exige revisão do interpretador, licenças e homologação, sem promover este laboratório a distribuição geral. PDFium e suas licenças permanecem fixados pelo lock anterior.

PS/PDFs sintéticos, relatórios e capturas ficam em `.desktop-test/hold-native/` e `%TEMP%/4PrintShops-hold-experimental/`. DEVMODEs privados não são gravados. Os temporários devem ser tratados como cópias locais do ensaio e removidos somente após fechar os programas que os utilizam. Os dados originais e os documentos de clientes não foram utilizados nos testes.

## Roteiro para o primeiro envio pessoal

1. Fechar conscientemente a **Hold Test experimental.9**, exportando qualquer documento pendente. Instalar a experimental.10 pela mesma identidade. A instalação principal permanece separada.
2. Conferir no Windows que `KONICA MINOLTA C6100PS` é a padrão e que a fila de laboratório continua FILE:. O gerenciamento automático deve estar desativado. O aplicativo bloqueia capturas se esse requisito não estiver atendido.
3. Abrir PDFs → Arquivo → Configuração de impressão. Selecionar `KONICA MINOLTA C6100PS` e clicar **Abrir PDF sintético de teste**. Se já houver outro documento, respeitar a confirmação de descarte. O documento continua A4 no workspace.
4. Em Propriedades, configurar tamanho original/lógico A3 e tamanho de saída A3, retrato, zoom 100%, uma cópia e nenhuma imposição/booklet. O aplicativo posiciona o conteúdo A4 a 120% sobre essa página lógica A3. O uso de A3 no driver evita um segundo ajuste automático por diferença entre tamanho original e saída.
5. Em Workflow, selecionar **Output Method = Hold** e **Print Mode = Normal**. Confirmar e clicar **Verificar Hold da sessão**.
6. Clicar **Enviar trabalho retido — Experimental**. Aguardar a captura e interpretação locais. Qualquer falha deve bloquear a confirmação de envio. Na primeira inspeção, pode escolher Cancelar e conferir que o documento permanece aberto.
7. Para o teste real, repetir o preparo, conferir todos os dados da confirmação e anotar `4PS HOLD EXP xxxxxxxx`. **Somente o usuário deve confirmar Enviar trabalho retido.**
8. Conferir na IC-604 que esse trabalho ficou armazenado como Hold, sem saída física automática. Conferir papel A3, escala/posição e uma cópia antes de liberar qualquer impressão na máquina. Se não estiver retido, interromper a homologação e registrar o resultado.
9. Em caso de erro ou timeout, verificar a fila Windows e a Konica. Não reenviar por suposição. Cópias adicionais, preto e branco, bandejas e acabamentos serão avaliados em etapas posteriores.

## Reproduzir a validação sem envio físico

```powershell
node scripts/build-konica-capture.mjs
node tests/desktop-konica-capture.mjs print
node tests/desktop-konica-capture.mjs holdprint
node tests/desktop-konica-capture.mjs hold 120
node scripts/test-print-native.mjs
npm test
node node_modules/electron-builder/cli.js --config electron-builder.hold-native-test.cjs --win nsis --x64 --publish never
node scripts/check-desktop-package.mjs dist-desktop/hold-native-test/win-unpacked/resources/app.asar
node tests/helpers/seed-hold-preference.mjs
node tests/desktop-hold-native-installed.mjs
```

Os três comandos de captura criam trabalhos somente na fila FILE:. O teste instalado exige o laboratório .9 conhecido e fechado, preserva o perfil, testa a confirmação **Cancelando**, compara preferências e impressora padrão e exporta o sintético posteriormente. Nenhum harness contém uma ação para confirmar o envio físico. A autorização de envio dos testes unitários usa um cliente simulado.

O primeiro upgrade alterou o registro da .9 para a .10. Nas revalidações, o Setup .10 foi reaplicado à mesma instalação .10. A preferência foi gravada pelo executável .9 preservado, aguardando a interface e o fechamento normal, e foi lida pelo executável .10 no mesmo perfil. O teste anterior supunha a existência de um marcador não persistido e falhou nessa suposição. O procedimento foi corrigido antes de registrar a aprovação.

## Resultados executados

| Verificação | Resultado |
| --- | --- |
| Node completo | 223 aprovados, zero falhas/skips na rodada final sequencial, incluindo cancelamento, prova inválida, confirmação simulada única, timeout sem reenvio e validação do canal |
| C# | 30 verificações aprovadas, com 6 anteriores de DEVMODE, 17 de Hold e 7 da política de impressora padrão |
| Captura PS real isolada | Print, Hold e HoldPrint diferenciados por execução do PS; Hold/A3/120% medido e inspecionado |
| Sintético e workspace | Original e giro de 180° serializado aprovados; outro documento rejeitado |
| Gestão automática ativa | Recusa real antes de iniciar trabalho FILE:; aplicativo não altera a preferência |
| Pacote instalado | Mesma identidade e diretório da .9, preferência gravada pela .9 preservada na .10 |
| UI instalada | Botão do sintético, propriedades, verificação Hold, captura, interpretação e confirmação Cancelar aprovados |
| Proteções | Documento girado permaneceu pendente depois de cancelar; DEVMODEs globais/per-user e impressora padrão preservados |
| Exportação posterior | Uma página e giro de 180° reabertos corretamente; pendência quitada somente após salvar |
| Chrome headless curto | Versão web, abas/menus, PDFs de 100/350 páginas, exportação, proteção e QR aprovados |
| Pacote | 430 entradas ASAR, fontes atuais, hashes dos auxiliares/PDFium/sintético e exclusões conferidos |
| Sintaxe e espaços | 27 arquivos JS/MJS/CJS, três PowerShell e compilações C# aprovados; `git diff --check` aprovado |
| Envio direto real | Não executado pelo agente; aguarda confirmação e conferência pessoal do usuário |

Uma rodada Node concorrente com o teste instalado excedeu prazos de um segundo em dois testes simulados anteriores do auxiliar. A repetição sequencial passou com 223 testes. Os prazos funcionais do atualizador e do envio não foram reduzidos ou contornados.

## Arquivos envolvidos nesta continuação

- Novos módulos nativos `CaptureEnvironmentPolicy.cs`, `KonicaPostScriptCapture.cs`, `PdfiumSyntheticGuard.cs` e `HoldExperimentalJob.cs` em `desktop/native/`. O host compila os comandos de envio somente com `HOLD_EXPERIMENTAL`.
- `desktop/hold-test.mjs` coordena prova e confirmação. `app.mjs`, `preload.cjs` e `print-config.mjs` integram a variante restrita. O canal comum de propriedades continua recusando comandos de envio.
- `public/js/app/desktop-print-ui.js`, `public/js/main.js` e `public/js/tools/pdfs.js` ligam os controles de ensaio ao workspace atual, sem marcar exportação.
- `electron-builder.hold-native-test.cjs`, `scripts/build-hold-native.mjs`, `build-konica-capture.mjs`, auditoria do pacote e exclusões em `.gitignore` organizam o laboratório. A geração virtual anterior e seu build receberam a proteção da impressora padrão.
- Testes novos `desktop-hold-test.test.mjs`, `desktop-konica-capture.mjs`, `desktop-hold-prepare.mjs`, `desktop-hold-native-installed.mjs`, `helpers/seed-hold-preference.mjs` e `native/CaptureEnvironmentPolicyTests.cs`. Helpers Windows e o executor C# foram atualizados. Documentação anterior preservada e complementada.

Permanecem pendentes o primeiro envio direto à IC-604, retenção real, fidelidade física, demais configurações avançadas, outros PDFs/impressoras e distribuição comercial. Também permanecem os limites anteriores de atualizador automático, servidor/piloto assistido e SMB/UNC. Sem commit ou push nesta etapa.
