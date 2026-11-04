# Configuração de impressão Windows — protótipo sem envio

Data: 03/11/2026. Base limpa conferida: `77e3ca8`, após `c539206`. Esta etapa implementa somente configuração por trabalho e uma simulação geométrica; não implementa impressão física.

## Referência e investigação aprovada

- AccurioPress C6100, controladora **IC-604**, sem Fiery.
- Fila: `KONICA MINOLTA C6100PS`; driver `KONICA MINOLTA C6100/C6085PS`.
- Modelo V3. O número `1407374916583424`, informado pelo usuário, resulta em **5.0.504.0** ao separar quatro componentes de 16 bits. `MajorVersion=3` é outra informação.
- Porta cadastrada `IP_192.168.88.100`: o nome não comprova o protocolo configurado. O protótipo não abre conexão direta com esse endereço.
- Consulta anterior somente de leitura identificou `PSCRIPT5.DLL` e `PS5UI.DLL` em `10.0.19041.6456`, além de `KOI6100_.ppd`, sem versão de arquivo. Esses componentes não substituem a versão do pacote OEM.
- C368SeriesPCL e 6136PS podem aparecer na enumeração do Windows; não foram selecionadas nos testes.

A investigação foi aprovada na conversa antes deste desenvolvimento. Electron oferece parâmetros comuns, mas não um contrato público para transportar integralmente o DEVMODE privado. Envio RAW exigiria produzir os comandos corretos de configuração e não foi adotado. A arquitetura escolhida usa um auxiliar externo Win32, parâmetros completos por sessão e propriedades nativas; o motor gráfico futuro ainda precisa ser homologado.

Referências: [DocumentProperties](https://learn.microsoft.com/en-us/windows/win32/printdocs/documentproperties), [PRINTER_INFO_9](https://learn.microsoft.com/en-us/windows/win32/printdocs/printer-info-9), [EnumPrinters](https://learn.microsoft.com/en-us/windows/win32/printdocs/enumprinters), [CreateICW](https://learn.microsoft.com/en-us/windows/win32/api/wingdi/nf-wingdi-createicw), [DeviceCapabilitiesW](https://learn.microsoft.com/en-us/windows/win32/api/wingdi/nf-wingdi-devicecapabilitiesw), [GetDeviceCaps](https://learn.microsoft.com/en-us/windows/win32/api/wingdi/nf-wingdi-getdevicecaps), [Konica — Layout PS Plug-in](https://manuals.konicaminolta.eu/AccurioPress-C6100-C6085/EN/contents/PRT_102642550.html).

## Organização e segurança

- `desktop/native/DevModeState.cs`: buffer completo, revisão, substituição transacional e limpeza em memória.
- `desktop/native/PrintConfigNative.cs`: consultas Win32 e janela nativa. Enumeração nível 4 com LOCAL/CONNECTIONS; não procura impressoras pela rede nem abre todas as filas para obter propriedades.
- `desktop/native/PrintConfigHost.cs`: processo x64 separado, protocolo JSON limitado por stdin/stdout, sessão única e encerramento quando o processo principal termina. Código C# usa P/Invoke; não carrega o driver dentro do Electron. Reutiliza o .NET Framework do Windows, sem pacote npm ou instalação de dependência.
- `desktop/print-config.mjs`: caminho fixo do auxiliar, SHA256 conferido contra manifesto incorporado ao ASAR, comandos permitidos, exclusão mútua, ids e geração, limites e timeout. Não usa shell nem aceita executável, caminho, handle ou comando arbitrários vindos do renderer. stderr do driver não é registrado.
- `desktop/app.mjs` e `preload.cjs`: canal único `print:config`, validação de origem/frame e janela proprietária fornecida pelo processo principal. O diálogo nativo torna a janela proprietária temporariamente indisponível; erros liberam a janela. Operações em andamento impedem encerramento/instalação pelos fluxos existentes.
- `public/js/app/desktop-print-ui.js`, `desktop-print-layout.js` e `public/css/print-config.css`: janela técnica acessível em **PDFs → Arquivo → Configuração de impressão (protótipo)**, somente no Desktop. A versão web não ganha um comando que não possa executar.

O auxiliar não possui APIs de envio de trabalhos nem de gravação de preferências. Utiliza **CreateIC**, um contexto de informação que não aceita desenho GDI, para consultar geometria. Não recebe PDF, bytes do workspace ou caminho de documento.

O DEVMODE fica somente na memória do auxiliar, incluindo `dmSize + dmDriverExtra`, com limites e validação de tamanho. Não atravessa IPC, não vira JSON/base64, não é salvo e não é registrado em logs. O driver fornece os dados iniciais, usando preferências do usuário quando existentes; saída e entrada são buffers separados. O diálogo usa exclusivamente OUT_BUFFER, IN_BUFFER e IN_PROMPT, sem atualização persistente.

Cancelar descarta o buffer de trabalho; confirmar só substitui o estado após validar o resultado. Cada confirmação avança a revisão. Pedidos de sessão/revisão antigos são recusados; resultados de outro processo ou geração são ignorados. Uma solicitação concorrente é recusada, não intercalada. Fechar/reabrir aguarda a liberação da sessão anterior. Reiniciar o auxiliar exige selecionar novamente a impressora.

O driver continua responsável por interpretar seus dados privados. Não há offsets OEM, reconstrução por JSON ou simulação de grampeamento. Campos públicos ausentes aparecem como **Controlado pelo driver / não informado**. Um resumo válido não comprova acabamentos ou saída física: o Windows/driver pode normalizar ajustes e métricas públicas não revelam todas as transformações privadas.

## Prévia técnica

A página de origem é um retângulo **A4 sintético**, sempre 210 × 297 mm, sem consultar o PDF aberto:

| Papel simulado | Escala | Conteúdo | Margens geométricas centralizadas |
| --- | --- | --- | --- |
| A3 297 × 420 mm | 100% | 210 × 297 mm | 43,5 mm laterais; 61,5 mm superior/inferior |
| A3 297 × 420 mm | 120% | 252 × 356,4 mm | 22,5 mm laterais; 31,8 mm superior/inferior |

Também é possível usar o papel nominal devolvido pelo driver. Papel branco, conteúdo azul, área imprimível tracejada e excedente vermelho distinguem margens e cortes. A área imprimível só é sobreposta quando suas dimensões correspondem ao papel da simulação; informações ausentes/incompatíveis não são inventadas.

Escolher A3 ou 120% nesta janela **não configura o driver**. O operador altera papel e orientação nas propriedades nativas. O resumo informa a escala pública, quando disponível; zoom OEM, Fit, tamanho original/saída, imposição e acabamento não são interpretados. A prévia não promete reproduzir essas transformações.

Na impressão futura, escala/posicionamento precisam ter um único responsável. A4 a 120% em A3 exige conteúdo 252 × 356,4 mm sobre uma página lógica e uma folha A3, sem segundo zoom/ajuste no driver. Cópias não devem ser multiplicadas tanto pelo aplicativo quanto pelo driver. Comparar GDI/EMF e PS suportado, fontes, vetores, transparências, CMYK/ICC e sobreimpressão com Acrobat/RIP antes de escolher o motor. Nenhum desses caminhos de envio foi implementado aqui.

## Compilar e instalar o laboratório

No Windows x64, usar o .NET Framework com `Framework64/v4.0.30319/csc.exe` disponível. Não instalar runtimes ou modificar políticas automaticamente.

```powershell
node scripts/build-print-helper.mjs
npm run desktop
```

O compilador gera somente `build/native/4PrintShops.PrintConfig.exe` e `manifest.json`, ambos ignorados no Git. Os scripts/fontes são versionáveis. O builder compila o auxiliar antes de empacotar; o manifesto fica no ASAR e o executável em `resources/native/`. O hash detecta binário incompatível/adulterado em relação ao manifesto incorporado; não substitui Authenticode ou a autenticação da distribuição. A futura distribuição comercial precisa assinar/homologar seus componentes. O atualizador existente não teve suas verificações reduzidas.

```powershell
node node_modules/electron-builder/cli.js --config electron-builder.print-config-test.cjs --win nsis --x64 --publish never
node scripts/check-desktop-package.mjs dist-desktop/print-config-test/win-unpacked/resources/app.asar
node tests/desktop-print-config-installed.mjs
```

Instalador isolado: `dist-desktop/print-config-test/4PrintShops-Print-Config-Test-1.0.0-experimental.8-Windows-x64-Setup.exe`. Produto, appId, atalhos e perfil **4PrintShops Print Config Test** separados. Não substitui produto principal, feed ou outros laboratórios. Mantém a associação própria/Open With, sem mudar o leitor padrão. A raiz do projeto continua em `.6`; `.8` é metadado exclusivo do laboratório.

O teste instalado usa somente PDF sintético e deixa a instalação em `.desktop-test/print-config/installed-validation/installed/`. O perfil fica em `%APPDATA%/4PrintShops Print Config Test`. Instaladores, executáveis gerados, capturas e logs não pertencem ao commit. A automação usa a porta de depuração somente na inicialização explícita do teste; a execução normal não a ativa.

## Validação e limites

| Verificação executada | Resultado |
| --- | --- |
| Suíte Node completa | **213 aprovados**, zero falhas/skips; inclui 11 testes novos de comandos, geometria, revisões, reabertura, hash, timeout e falha do auxiliar |
| Estado C# compilado/executado | **6 verificações aprovadas**, incluindo 8192 bytes privados sintéticos, cancelamento, confirmação, cópia independente e rejeição transacional |
| Chrome headless, modo de abas/menus | Aprovado: versão web, PDFs 100/350 páginas, exportação, proteção e fechamento/reabertura QR. Não repetiu toda a bateria de redimensionamento/assinaturas neste modo |
| NSIS experimental.8 realmente instalado | Aprovado, produto/perfil próprios, leitor padrão do Windows preservado |
| Propriedades reais da Konica | Abertura, Cancelar e confirmação de mudança para paisagem aprovados; resumo atualizado |
| Microsoft Print to PDF | Abertura, Cancelar e confirmação de paisagem aprovados. Mudança A3 pela janela avançada não automatizada |
| Preferências globais e per-user | DEVMODEs antes/depois idênticos na Konica e na impressora virtual, comparados somente em memória |
| PDF modificado durante configuração | Preservado e pendente de exportação; fechar/reabrir não duplicou a janela e voltou aos defaults intactos |
| Falha real do auxiliar instalado | Encerrado somente o processo auxiliar do laboratório; erro exibido, seleção refeita com sucesso, Electron/PDF preservados |
| Exportação nativa após o protótipo | PDF sintético salvo/reaberto com duas páginas e giro de 90° preservado; pendência quitada somente após salvar |
| Pacote | **426 entradas ASAR**, fontes de runtime/ícones iguais ao código e auxiliar consistente com manifesto; sem PDFs, chaves ou testes |
| Sintaxe/espaços | 15 arquivos JS/MJS/CJS e dois PowerShell aprovados; fontes C# compiladas; `git diff --check` aprovado |

O teste de estado C# não interpreta ou homologa o significado de todos os recursos OEM. As confirmações reais comprovam o fluxo de configuração e o transporte do estado, não o resultado de impressão ou acabamento. Não foram enviados trabalhos em nenhuma dessas verificações.

SHA256 do instalador experimental.8: `1f1c3c7d8c4fd62d2904e74e6bb50dd1d7cc63a46e7edf1f3de54e8bb24f9d66`. Relatório, capturas e logs em `.desktop-test/print-config/`; não versionar os gerados. O teste instalado final está em `installed-validation/result.json`.

A consulta Konica inicial retornou A4 retrato, uma cópia, escala pública 100%, seleção automática de bandeja e contexto de 1200 dpi com área imprimível aproximada de **204,26 × 291,40 mm**. Duplex não estava indicado no campo público: a interface não o presumiu. Esses são ajustes observados nesta estação e sessão, não padrões universais da máquina.

A política local impede executar `.ps1` diretamente. Os testes executam seu código de instrumentação pelo modo de comando do Windows PowerShell 5.1, sem alterar políticas ou usar ExecutionPolicy Bypass. O monitor compara DEVMODEs globais/per-user somente em memória e devolve apenas igualdade; não grava bytes privados.

### Roteiro manual sem imprimir

1. Abrir **4PrintShops Print Config Test**, sem parâmetros de depuração. Abrir o módulo PDFs; não é obrigatório importar um arquivo.
2. Em **Arquivo**, abrir **Configuração de impressão (protótipo)**. Conferir a lista e selecionar somente `KONICA MINOLTA C6100PS` para a referência real. Não testar as filas desativadas.
3. Anotar o resumo. Abrir **Propriedades do driver**, alterar um ajuste local e pressionar **Cancelar**. Conferir que o resumo anterior permanece.
4. Reabrir Propriedades, escolher orientação/papel e confirmar em **OK**. Conferir os campos públicos recuperados. Não usar comandos do driver para salvar padrões, presets, autenticação ou configurações do equipamento.
5. Para papel A3 real, configurar A3 no driver; depois escolher **Papel informado pelo driver**. Comparar tamanho nominal, orientação e área imprimível. Se o driver representar tamanhos original/saída em campos privados, tratar a divergência como limitação, sem presumir um resultado físico.
6. Na simulação A3, alternar 100% e 120%; conferir medidas e margens da tabela. Em A4 a 120%, conferir a área vermelha e o aviso de corte.
7. Fechar e reabrir o protótipo: ele inicia das preferências existentes, sem persistir as escolhas da sessão. Confirmar que um PDF modificado continua aberto e pendente de exportação.

Não existe botão Imprimir. Fechar esta janela não executa, exporta ou envia nada. Se o driver falhar, consultar a mensagem e selecionar novamente a impressora; o documento do workspace permanece independente. O timeout de uma janela nativa travada encerra o auxiliar e libera a janela principal, mas descarta a sessão local.

Continuam pendentes: homologação física IC-604, fidelidade de conteúdo/cor/fontes, bandejas/duplex/acabamentos reais e ajustes privados de tamanho/escala. Também permanecem pendentes as etapas anteriores de servidor real/piloto assistido, instalação/reinício automáticos pelo atualizador e SMB/UNC. Este protótipo não resolve nem altera essas homologações.

## Auditoria para checkpoint — 03/11/2026

**Base:** `77e3ca8` (ícones), precedida por `c539206` (atualizações assistidas) e `c450fed` (abertura Windows). **27 candidatos: 12 modificados e 15 novos; staging vazio; nenhuma remoção.** Todos pertencem ao primeiro protótipo de configuração de impressão. Não houve alteração funcional nesta auditoria: foram atualizados somente este documento e `ESTADO_PROJETO.md`, já incluídos nos 27 candidatos.

O usuário relatou que realizou os testes manuais e que as propriedades do driver e o resumo aparentemente funcionam corretamente. Isso complementa os testes instalados registrados acima; **impressão física não foi implementada nem homologada**. A arquitetura de leitura de Hold/PrintTicket foi aprovada para uma etapa posterior; **a etapa 2A não está implementada neste checkpoint**.

M = modificado desde a base; N = novo:

| Arquivo | Estado | Finalidade |
| --- | --- | --- |
| `.gitignore` | M | Excluir auxiliar compilado e manifesto gerado |
| `desktop/app.mjs` | M | Integração IPC, auxiliar e bloqueios de operações |
| `desktop/preload.cjs` | M | Ponte restrita de configuração |
| `desktop/security.mjs` | M | Permitir somente os novos assets de interface |
| `desktop/native/DevModeState.cs` | N | Preservar buffer completo e revisão |
| `desktop/native/PrintConfigHost.cs` | N | Processo isolado e protocolo de comandos |
| `desktop/native/PrintConfigNative.cs` | N | Consultas Win32 e propriedades nativas |
| `desktop/print-config.mjs` | N | Cliente do auxiliar, integridade, erros e concorrência |
| `docs/DISTRIBUICAO_DESKTOP.md` | M | Distribuição e identidade do laboratório |
| `docs/ESTADO_PROJETO.md` | M | Estado, resultados, limites e checkpoint pendente |
| `docs/IMPRESSAO_WINDOWS.md` | N | Arquitetura, manual e esta auditoria |
| `electron-builder.config.cjs` | M | Compilar/empacotar auxiliar e interface |
| `electron-builder.print-config-test.cjs` | N | Identidade isolada experimental.8 |
| `public/index.html` | M | Carregar CSS do protótipo |
| `public/js/main.js` | M | Instanciar janela somente com ponte Desktop |
| `public/js/app/menus.js` | M | Comando de configuração no módulo PDFs |
| `public/js/app/desktop-print-layout.js` | N | Cálculos técnicos e rejeição de resultados antigos |
| `public/js/app/desktop-print-ui.js` | N | Janela de configuração e prévia sintética |
| `public/css/print-config.css` | N | Estilo da janela técnica |
| `scripts/build-print-helper.mjs` | N | Compilação local e manifesto de integridade |
| `scripts/authenticode-lab/prepare-kit.mjs` | M | Incluir fontes necessárias para reproduzir futuros kits |
| `scripts/check-desktop-package.mjs` | M | Conferir auxiliar no pacote contra manifesto ASAR |
| `tests/desktop-print-config.test.mjs` | N | Comandos, geometria, sessões e falhas |
| `tests/desktop-print-config-installed.mjs` | N | Teste real instalado, propriedades e regressões |
| `tests/helpers/windows-print-defaults.ps1` | N | Comparar preferências somente em memória |
| `tests/helpers/windows-print-properties.ps1` | N | Automatizar somente o diálogo do auxiliar de teste |
| `tests/native/PrintConfigStateTests.cs` | N | Verificar conservação transacional do DEVMODE |

### Segurança conferida nesta auditoria

- Fontes C# revisadas e recompiladas em pasta ignorada. Os binários recompilado e realmente instalado foram inspecionados por metadados, sem executar operações de impressora: exatamente **10 importações Win32**, todas de consulta, propriedades ou liberação: EnumPrintersW, OpenPrinterW, ClosePrinter, DocumentPropertiesW, GetPrinterW, DeviceCapabilitiesW, CreateICW, GetDeviceCaps, DeleteDC e GetWindowThreadProcessId.
- Sem importações/chamadas de StartDoc, StartPage, WritePrinter, AddJob, StartXpsPrintJob, SetPrinter ou SetDefaultPrinter no auxiliar. O protocolo aceita apenas list/select/properties/summary/release; não recebe documentos. A interface não contém ação de envio ou impressão. CreateIC consulta informações, sem desenhar páginas.
- O DEVMODE completo permanece em memória; entrada/saída separadas, cópias independentes, confirmação transacional e descarte no cancelamento. Não é serializado em respostas ou logs. O diálogo não usa flags de persistência. A inspeção não executou callbacks de propriedades nem enviou trabalhos.
- Nenhum candidato contém instalador, pacote, PDF, certificado, chave, credencial ou dado DPAPI. `.gitignore` conferido com artefatos reais e caminhos de teste dessas classes; todos permaneceram ignorados. Não foram excluídos arquivos existentes.
- Motores PDF/QR, PDFWorkspace, NSIS de associação, ícones e módulos dos atualizadores não foram modificados diretamente. As alterações no kit/builder/auditor de pacote são necessárias para compilar e distribuir o novo auxiliar. Authenticode, origem/frame/IPC, identidade principal e preferências foram preservados.

### Verificações executadas agora e evidências anteriores

- **Executados novamente:** 213 testes Node aprovados, zero falhas/skips; compilação das três fontes do auxiliar; seis verificações C# aprovadas; sintaxe de 15 JS/MJS/CJS e dois PowerShell; auditoria do pacote realmente instalado com 426 entradas e hash do auxiliar; hash do Setup igual ao registrado; `git diff --check`.
- **Revalidados por leitura, sem repetir a automação:** relatório/logs da instalação experimental.8 e testes reais de propriedades, preferências, recuperação e exportação. O pacote instalado continua correspondente às fontes de runtime atuais. Chrome/diálogos nativos e reinstalação não foram repetidos nesta auditoria; permanecem resultados da rodada de implementação, complementados pelo relato manual do usuário.
- A inspeção dos dois assemblies com a mesma identidade exigiu processos de reflexão separados, por restrição do .NET Framework. A repetição isolada confirmou o binário instalado. Não houve correção de código motivada por esse ajuste de instrumentação.
- Inventário e SHA256 finais preservados em `.desktop-test/print-config/auditoria-20261103/auditoria-checkpoint.json`, ignorado, para conferir o futuro staging. Fontes finais e histórico foram preservados.

**Recomendação:** registrar um único commit com os 27 arquivos acima, mensagem `Adiciona prototipo de configuracao de impressao Windows`. **Aguardar autorização explícita; nenhum commit ou push foi executado.** Não iniciar 2A nem geração/envio até concluir o checkpoint. Permanecem os limites de A3 real/bandejas/duplex/acabamentos e todas as homologações físicas e do atualizador descritas anteriormente.
