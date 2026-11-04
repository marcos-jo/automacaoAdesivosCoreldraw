# Impressão Windows — verificação de Hold e laboratório PDFium

## Continuação experimental.10

O usuário aprovou manualmente a experimental.9, incluindo Hold na sessão e impressão pelo Edge com os ajustes do driver respeitados. A etapa seguinte acrescenta captura PS real em FILE:, interpretação do estado de Hold e envio experimental condicionado à confirmação. A experimental.10 conserva a identidade/perfil da **4PrintShops Hold Test** e atualiza a .9. O primeiro envio direto à IC-604 ainda depende da ação e homologação pessoal do usuário. A gestão automática da impressora padrão passou a bloquear capturas, após o efeito identificado e corrigido com autorização. Procedimento e resultados em [IMPRESSAO_HOLD_EXPERIMENTAL.md](IMPRESSAO_HOLD_EXPERIMENTAL.md).

As seções seguintes permanecem como histórico da experimental.9, na qual o botão de envio ficou bloqueado.

## Situação em 04/11/2026

Base confirmada `5d07e81`, protótipo anterior aprovado manualmente pelo usuário. O Git estava limpo antes desta etapa. O laboratório atual é **4PrintShops Hold Test 1.0.0-experimental.9**, separado do produto principal e dos laboratórios anteriores.

**O botão Enviar trabalho retido permanece bloqueado.** Foi comprovada a leitura de Hold a partir do DEVMODE da sessão real. A saída PostScript produzida pelo driver Konica ainda precisa ser capturada e examinada em ambiente isolado. A Microsoft Print to PDF permitiu validar a geração do PDFium, mas não possui o recurso proprietário de retenção da IC-604.

Nenhum trabalho foi enviado à Konica. Somente a Microsoft Print to PDF recebeu páginas sintéticas nos ensaios explicitamente executados. O aplicativo oferece a abertura da cópia atual em leitor externo enquanto a homologação nativa permanece pendente.

## Validação da sessão

- Referência exata `KONICA MINOLTA C6100PS`, driver `KONICA MINOLTA C6100/C6085PS`, IC-604, V3. Os arquivos OEM consultados retornam versão textual `5.0.504`, compatível com a identificação informada `5.0.504.0`.
- `PrintTicketConverter.ConvertDevModeToPrintTicket` recebe uma cópia completa do DEVMODE confirmado, em JobScope. O conversor não substitui o estado do driver por um PrintTicket reconstruído. A leitura não equivale a validar a saída física. [Documentação Microsoft](https://learn.microsoft.com/en-us/dotnet/api/system.printing.interop.printticketconverter.convertdevmodetoprintticket?view=windowsdesktop-10.0)
- `DocumentKOOutputMethod=Hold` é exigido junto de `DocumentKOPrintMode=Print`, que significa Normal. Print, HoldPrint, modos incompatíveis, campos ausentes, duplicados, namespaces inesperados e XML inconclusivo mantêm o envio bloqueado. O PPD local também lista restrições entre Hold e Proof/Wait/Spool.
- O namespace OEM é conferido contra o driver de referência e a versão textual do PScript instalado. Nesta estação, `PSCRIPT5.DLL` informa textualmente `10.0.19041.6456`, embora seus componentes fixos de versão sejam diferentes. Não usar esses componentes fixos para montar o namespace.
- A decisão está vinculada à fila, sessão, revisão, digest dos bytes completos do DEVMODE e identidade dos arquivos/porta do driver. Os hashes e bytes privados permanecem na memória. Confirmar propriedades invalida a decisão anterior. Cancelar conserva os ajustes e a validação anterior, sujeito à conferência da identidade atual.
- O parser proíbe DTD e resolução externa e limita o XML. O renderer recebe somente campos públicos e motivo. Respostas antigas continuam protegidas pela geração e revisão existentes.
- `canSend` e `outputValidated` permanecem falsos. O canal IPC e o auxiliar de configuração recusam comandos de envio. Habilitar o botão pelo DOM não cria um caminho de impressão.

O manual distingue Hold de Print/Hold e Print. Conferir a retenção na máquina continua sendo etapa posterior à inspeção do arquivo gerado. [Manual oficial Konica](https://manuals.konicaminolta.eu/AccurioPress-C6100-C6085/EN/contents/PRT_102642662c.html)

## Geração virtual com PDFium

O executável separado `4PrintShops.PrintLab.exe` utiliza PDFium nativo, FPDF_RenderPage em HDC, modo EMF e flags de impressão. Não utiliza canvas, captura de tela ou rasterização prévia da página. Executa em processo próprio, recebe uma única página e não possui parâmetro para selecionar uma fila física.

O destino permitido é exclusivamente `Microsoft Print to PDF`, com driver `Microsoft Print To PDF` e porta `PORTPROMPT:` conferidos antes de criar o contexto. O DEVMODE completo dessa impressora virtual configura apenas o trabalho local, em A4 retrato e uma cópia. O driver precisa confirmar esses valores. A saída é um PDF novo com nome único. Não há lote, repetição automática ou integração do executável virtual ao canal de envio do Electron.

As APIs StartDoc/StartPage/EndPage/EndDoc existem **somente nesse executável virtual separado**. Falhas abortam o trabalho virtual e liberam os recursos. O auxiliar de propriedades continua usando CreateIC e não inclui essas APIs. O build gera o PDF sintético, sem iniciar qualquer trabalho.

PDFium fixado em `chromium/8125`, pacote `pdfium-win-x64.tgz`, SHA256 `9a964261760cff52570aabf038ab346d3a5034dd8b30d0d5fe8861235d356b8b`. O lock versionado e a conferência precedem a extração e o build. O binário, a fonte de ensaio e as licenças são recursos gerados ignorados. O pacote de laboratório inclui as licenças do PDFium, dependências e Liberation. A licença permite redistribuição sujeita às condições documentadas, cuja revisão deve acompanhar a distribuição futura. [Licença PDFium](https://pdfium.googlesource.com/pdfium/+/refs/heads/main/LICENSE), [API oficial](https://pdfium.googlesource.com/pdfium/+/refs/heads/main/public/fpdfview.h)

### Resultado observado

| Medida do PDF sintético com TrueType incorporada | Origem | Saída virtual instalada |
| --- | --- | --- |
| Páginas | 1 | 1 |
| Dimensões em pontos | 595,2756 × 841,8898 | 595,32001 × 841,92004 |
| Recursos de fonte | 1 | 0 |
| Imagens | 0 | 0 |
| Caracteres extraídos pelo pypdf | 82 | 0 |

Os textos visíveis e o retângulo continuam desenhados como vetores. O caminho EMF/Microsoft Print to PDF converteu os glifos em contornos, inclusive no ensaio com TrueType incorporada. **Conservação das fontes e texto pesquisável não foi demonstrada nesse caminho.** O ensaio inicial com fonte-base também produziu contornos.

As duas renderizações a 72 dpi foram inspecionadas. O título ocupa a mesma largura, com deslocamento vertical de um pixel. A amostra azul permaneceu RGB 51/102/204. Diferença média absoluta RGB das imagens completas de 0,723 em 255, sem equivalência pixel a pixel. As medidas registradas referem-se ao sintético inspecionado. CMYK, ICC, cores especiais, transparências complexas, sobreimpressão, fontes variadas e o RIP real permanecem sem homologação.

## Cópia atual em leitor externo

Em PDFs, **Arquivo → Abrir cópia atual em leitor externo** serializa o PDFWorkspace atual para uma pasta exclusiva em `%TEMP%/4PrintShops-external-copies/{UUID}/`. O operador escolhe Microsoft Edge ou outro executável local de leitor. O arquivo original permanece preservado. O leitor recebe um argumento de caminho separado, sem shell ou mudança de associação padrão.

Cancelar não cria o PDF temporário. Abrir a cópia não marca o workspace como exportado e não quita a proteção contra perda de alterações. O indicador continua pendente até Salvar cópia. Erros liberam os controles e são apresentados. A exclusão automática desses temporários foi adiada para evitar remover um arquivo ainda em uso pelo leitor. O operador pode removê-los depois de fechar o leitor. Eles podem conter dados do documento atual e permanecem locais.

No teste instalado, cancelar preservou o diretório e abrir no Edge recebeu uma cópia com duas páginas e o giro de 90° aplicado. O workspace continuou pendente, a exportação posterior passou e o leitor padrão permaneceu igual. A impressão pelo leitor externo depende das capacidades e configurações desse leitor e deve ser conferida pelo operador.

## Reproduzir o laboratório

Com as dependências do projeto já preparadas no Windows x64, executar os comandos a partir da raiz do projeto. A preparação baixa somente o pacote PDFium fixado quando ausente. Não modifica a política do PowerShell.

```powershell
& ([scriptblock]::Create((Get-Content scripts/prepare-pdfium.ps1 -Raw))) -ProjectRoot (Get-Location).Path
node scripts/build-print-helper.mjs
node scripts/build-print-lab.mjs
node scripts/test-print-native.mjs
npm test
node tests/desktop-hold-native.mjs
node node_modules/electron-builder/cli.js --config electron-builder.hold-test.cjs --win nsis --x64 --publish never
node scripts/check-desktop-package.mjs dist-desktop/hold-test/win-unpacked/resources/app.asar
$env:PRINT_CONFIG_HOLD_LAB='1'
node tests/desktop-print-config-installed.mjs
node scripts/run-print-lab.mjs .desktop-test/hold-lab/installed-validation/installed/resources/print-lab
```

O último comando cria um trabalho **virtual**, com saída em `.desktop-test/hold-lab/`. Repeti-lo cria um novo ensaio virtual, nunca um envio à Konica. O teste de propriedades abre e confirma diálogos locais do driver, sem persistir preferências. O teste instalado usa um PDF sintético de duas páginas para verificar o fluxo de edição/exportação separado do ensaio virtual de uma página.

O instalador gerado é `dist-desktop/hold-test/4PrintShops-Hold-Test-1.0.0-experimental.9-Windows-x64-Setup.exe`, SHA256 `994fa033d32f9f4abdf208698bf9db96d0a477293dbf93c5741a79930f2cc2e0`. O produto, appId, atalhos e perfil `%APPDATA%/4PrintShops Hold Test` são exclusivos. A raiz continua em experimental.6. Recursos PDFium e o PDF sintético acompanham somente o laboratório `.9`. O ASAR continua excluindo PDFs de usuário, certificados, chaves e testes.

## Testes concluídos

| Verificação | Resultado |
| --- | --- |
| Node completo | 218 aprovados, zero falhas/skips, incluindo cinco testes novos de leitor externo, validação IPC e bloqueio do botão |
| C# | 6 verificações anteriores de DEVMODE e 17 de Hold aprovadas |
| Driver real, somente propriedades | Print, Hold, HoldPrint e Normal identificados corretamente no PrintTicket |
| Cancelar e confirmar | Cancelar conservou Hold e revisão; confirmar invalidou a validação anterior |
| Preferências | DEVMODEs globais/per-user idênticos antes/depois na Konica e na impressora virtual, comparação somente em memória |
| NSIS `.9` realmente instalado | Produto isolado, abertura Windows/PDF.js, Hold e botão bloqueado, propriedades, simulação e cancelamento aprovados |
| Recuperação instalada | Falha do auxiliar isolada, seleção refeita, PDF modificado preservado |
| Leitor externo instalado | Cancelar e abertura pelo Edge aprovados, cópia com giro preservado e workspace ainda pendente |
| Exportação instalada | Duas páginas e rotação de 90° preservadas, pendência quitada após salvar |
| PDFium instalado | Uma página virtual gerada, zero imagens, vetores preservados no sintético, texto convertido em contornos |
| Chrome headless curto | Abas/menus web, PDFs de 100/350 páginas, exportação, proteção e QR aprovados |
| Pacote | 427 entradas ASAR, correspondência com fontes, hashes dos auxiliares/DLL/sintético e licenças conferidos |
| Sintaxe e espaços | 18 arquivos JS/MJS/CJS, três PowerShell e compilação C# aprovados, `git diff --check` aprovado |

O primeiro ensaio automatizado de seleção de Hold falhou porque o teste enviava notificações incompletas ao combo OEM. Foi corrigido o destino, o identificador de 16 bits e a notificação de confirmação do controle. Os testes nativo e instalado foram repetidos e passaram. O mecanismo de leitura não foi relaxado para acomodar esse problema da automação.

Relatórios e capturas ficam em `.desktop-test/hold-lab/`, incluindo `hold-native-result.json`, `installed-validation/result.json` e `virtual-latest.json`. Instaladores, DLLs baixadas, executáveis, PDFs e relatórios gerados estão ignorados no Git. Nenhum PDF de cliente foi utilizado.

## Validação manual disponível

1. Instalar e abrir **4PrintShops Hold Test experimental.9**. Conferir o nome separado do aplicativo principal.
2. Em PDFs → Arquivo → Configuração de impressão, selecionar `KONICA MINOLTA C6100PS`.
3. Abrir Propriedades → Workflow. Escolher **Output Method = Hold**, **Print Mode = Normal**, confirmar em OK e clicar **Verificar Hold da sessão**. Deve aparecer Hold confirmado e a pendência de comprovação da saída. Enviar trabalho retido continua desabilitado.
4. Alterar para Print ou Print/Hold, confirmar e verificar novamente. O envio continua bloqueado. Alterar e cancelar deve conservar a configuração anterior. Fechar/reabrir inicia uma nova sessão a partir das preferências existentes.
5. Importar um PDF sintético, girar uma página e usar **Arquivo → Abrir cópia atual em leitor externo**. Cancelar, repetir e escolher Edge. Conferir a cópia atual e o indicador de alterações ainda pendentes no 4PrintShops. Nenhuma impressão é iniciada por esse comando.
6. Salvar cópia e reabrir para conferir o documento. A configuração de impressão não altera o PDFWorkspace.

## Bloqueio e próxima evidência necessária

Preparar um laboratório isolado com o driver Konica correspondente e um destino local sem comunicação com a máquina real, mediante aprovação prévia para sua configuração. Capturar a saída PS do mesmo motor, contexto e DEVMODE integral que seriam usados no envio. Comparar os três métodos de saída e comprovar que a saída Hold contém o comando efetivo de retenção, sem Print/Hold ou substituição por defaults. O PPD local descreve `SHKMFeatures/SHKMOutputMethod/Hold put`, mas sua presença no PPD não comprova emissão pelo pipeline.

Somente depois dessa evidência, homologar a fidelidade necessária, a confirmação única do operador e o primeiro trabalho sintético retido na IC-604. Nenhum endpoint de envio à Konica foi entregue nesta etapa. A ausência dessa prova mantém o botão bloqueado conforme o requisito do usuário. A3 real completo, bandejas, cópias, duplex, acabamentos e cores continuam pendentes.

Também permanecem pendentes SMB/UNC com FQDN/IPv4, implantação assistida no servidor real/piloto em dois PCs e instalação/reinício automáticos pelo updater. Authenticode, confiança Ed25519, perfis e preferências existentes foram preservados. Não houve commit ou push nesta etapa.
