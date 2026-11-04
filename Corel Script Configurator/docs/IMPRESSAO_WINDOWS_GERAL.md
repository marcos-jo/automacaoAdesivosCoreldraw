# Impressão Windows geral — experimental.12

## Estado em 04/11/2026

O usuário aprovou fisicamente a experimental.11 na AccurioPress C6100/IC-604. Confirmou escala, trabalhos consecutivos, Hold e todas as propriedades configuradas. Essa aprovação refere-se ao caminho .11 PDFium/HDC → captura Konica FILE: → PS interpretado → mesmos bytes RAW.

A .12 acrescenta impressão geral PDFium/Windows GDI para o documento completo do workspace. Hold é um diagnóstico opcional. O clique em **Imprimir** autoriza o trabalho, incluindo saída física conforme o método escolhido no driver. Print, HoldPrint e métodos de outros fabricantes são controlados pelas propriedades nativas. Não existe confirmação experimental adicional no fluxo normal. Depois da entrega, o usuário informou testes práticos satisfatórios da .12 em outras impressoras da gráfica. A matriz por modelo/opção continua sendo ampliada. O ambiente para continuar em casa está em [DESENVOLVIMENTO_WINDOWS.md](DESENVOLVIMENTO_WINDOWS.md).

## Preservação e auditoria da .11

Base `0af565f`, staging vazio, 20 arquivos modificados e nenhum arquivo inesperado no início. Cópia local verificada por SHA256 em `.desktop-test/print12/baseline-20261104-141941/`, com inventário `manifest.json`. A cópia é ignorada pelo Git. Não foram descartadas alterações nem criado commit.

Os 20 caminhos correspondem à .11 aprovada, integração em `desktop/app.mjs`, `hold-test.mjs`, `print-config.mjs`, três fontes nativos de Hold/captura/host, builder Hold Test, UI/layout, `public/js/main.js`, `tools/pdfs.js`, três documentos e seis testes/helpers. Os detalhes e limites originais continuam em [IMPRESSAO_HOLD_EXPERIMENTAL.md](IMPRESSAO_HOLD_EXPERIMENTAL.md).

## Driver Epson identificado localmente

| Item | Resultado somente por leitura |
| --- | --- |
| Fila | `WF-C21000 Series(Rede)` |
| Driver | `EPSON WF-C21000 Series`, V3 |
| Porta | Porta de rede Epson, consultada localmente |
| Pacote INF | `e_wfbfhe.inf`, `DriverVer=08/01/2025,3.60.00.00` |
| Motor | `E_4MAPFHE.DLL`, versão de arquivo `0.3.0.17` |
| Propriedades | `E_4UICFHE.DLL`, versão `0.3.100.13` |
| Dados | `E_4VIFFHE.VIF` |
| Processador / tipo padrão | `EPSON WF-C21000 Series` / `RAW` |
| Papel observado na sessão | A5 retrato, cerca de 148 × 210 mm, área imprimível cerca de 142,07 × 204,05 mm |

Trata-se da instalação do driver padrão Epson, com motor/configuração próprios. Não é a instalação PS baseada em PScript/PPD nem o driver Windows PCL padrão. `RAW` é o tipo do spooler e não identifica uma linguagem PostScript. A documentação oficial do equipamento lista ESC/P-R como linguagem e também emulações PCL/PS/PDF, mas isso não comprova qual PDL esta instalação emite. A consulta Win32 usada não expôs essa PDL e nenhum fluxo de bytes da Epson foi capturado. A UI informa essa limitação. [Epson — especificação oficial WF-C21000](https://www.epson.com.ph/For-Work/Printers/Business-Inkjet-Printers/WorkForce-Enterprise-WF-C21000-A3-Multifunction-Printer/p/C11CH88502)

O caminho usado é PDFium → HDC Windows criado para a fila selecionada com o DEVMODE completo → driver Epson. O driver produz sua linguagem e administra cópias, cor, duplex e acabamento. Nenhum comando proprietário Konica é enviado à Epson.

## Caminhos disponíveis

| Caminho | Impressoras / uso | Implementação |
| --- | --- | --- |
| Imprimir | Filas Windows que aceitem contexto GDI válido, incluindo Epson, Konica PS e impressora virtual | `WindowsPrintJob.cs`, PDFium EMF, CreateDC com DEVMODE confirmado |
| Diagnóstico opcional Konica / Hold | Konica de referência e sintético reconhecido | Caminho .11 FILE:/Ghostscript/RAW, com suas verificações e confirmações preservadas |
| Abrir cópia atual em leitor externo | Alternativa Edge/outro leitor | Fluxo anterior preservado |

O caminho geral não monta PS, PCL ou ESC/P-R diretamente. O Windows e o driver escolhido fazem essa tradução. A Microsoft documenta o fornecimento do DEVMODE ao CreateDC para configurar o contexto específico. O modo EMF do PDFium é mantido, com bibliotecas e licenças já fixadas no projeto. [Microsoft — contexto da impressora](https://learn.microsoft.com/en-us/windows/win32/printdocs/retrieving-a-printer-device-context), [PDFium — APIs de impressão](https://pdfium.googlesource.com/pdfium/+/main/public/fpdfview.h)

## Configurações e consistência

- DEVMODE completo, inclusive dados privados, fica no auxiliar. Não é reconstruído a partir do resumo. Cancelar propriedades mantém a configuração anterior. O retorno confirmado fornece o resumo e uma nova revisão.
- A cópia atual do workspace é serializada localmente, com giros, união, exclusões, reordenação e redimensionamento acumulados. Imprimir não marca a versão como exportada. O PDF original permanece preservado.
- Cada clique deliberado gera identificação própria, arquivo temporário exclusivo e token de uso único. A preparação abre e valida todas as páginas antes de StartDoc. Fila, porta, driver, processador, tipo de dados, hashes de arquivos do driver, sessão, revisão e bytes do DEVMODE são confrontados antes de iniciar o trabalho.
- CreateDC recebe exatamente os bytes confirmados. Se o driver alterar esse buffer, a operação falha antes de StartDoc. Não há fallback silencioso para padrões. O diálogo usa somente flags de buffer/prompt, sem persistência.
- Uma chamada StartDoc representa o documento inteiro. Não há laço de cópias no aplicativo. Quantidade, collation, duplex, bandeja, cor/PB, orientação, métodos de saída e acabamento permanecem sob controle do driver.
- Escala do aplicativo entre 1% e 400%, até duas casas decimais e vírgula. Quando o campo público de zoom do driver informa valor positivo diferente de 100%, ele prevalece e a escala do aplicativo fica suspensa. Ajustes privados de encaixe/N-up/zoom não são interpretados. A UI informa que esses ajustes podem transformar a saída além da prévia. Para testar exclusivamente a escala do aplicativo, configurar zoom/encaixe do driver em 100%/sem ajuste automático.
- Cada página usa suas próprias dimensões e rotação visual. Ela é centralizada na folha do driver, sem encaixe automático ou seleção automática de outro papel. A prévia técnica representa a página atual. O trabalho inclui todas as páginas na ordem do workspace, inclusive dimensões mistas.
- Excesso sobre papel/área imprimível é aviso visual. Não bloqueia Imprimir nem exige nova confirmação. Papel/posição inadequados podem produzir corte intencional.
- PDFium usa renderização de impressão sobre HDC, sem captura de tela ou rasterização integral solicitada pelo aplicativo. Aparências imprimíveis de anotações/formulários são incorporadas vetorialmente somente no documento temporário nativo. O workspace e o arquivo original não são achatados. A API de flatten oferece esse caminho para impressão. [PDFium — flatten para impressão](https://pdfium.googlesource.com/pdfium/+/main/public/fpdf_flatten.h)
- Uma falha após a autorização de envio é reportada como resultado incerto. Não há retry automático. Conferir fila/máquina antes de outro clique. O auxiliar tem até dez minutos por preparação/envio; timeout encerra o auxiliar, invalida a sessão e exige nova seleção. Não presume que um timeout desfez páginas já entregues.
- Não são modificadas fila, porta, driver, preferências globais/per-user, impressora padrão ou gestão automática do padrão. No laboratório atual essa gestão permanece desativada por autorização anterior. Em outro Windows com gestão automática, o próprio sistema pode alterar a impressora padrão ao imprimir.

## Testes efetivamente executados

- 236 testes Node e 41 verificações C# aprovados. Incluem escala única e precedência do driver, conservação de 8192 bytes privados, cancelamento, revisão, tokens, repetição deliberada, erro sem retry, caminhos IPC e regressões anteriores.
- Epson e Konica reais foram listadas, tiveram resumos consultados e PDFs de três páginas mistas preparados e cancelados no auxiliar. Nenhum StartDoc em fila física foi solicitado. Token cancelado e revisão antiga foram recusados. Preferências comparadas por monitor somente de leitura permaneceram iguais.
- A Microsoft Print to PDF gerou três páginas nas escalas 100%, 120% e 87,5%, incluindo fonte incorporada, texto vetorial colorido, giro de 180°, página 300 × 120 mm e página 100 × 150 mm. Também gerou uma página com campo de formulário e aparência pronta. PDF corrompido foi recusado.
- Instalação real .11 → .12 no diretório/registro NSIS/perfil existentes. Preferência anterior preservada. No Electron instalado, propriedades Epson canceladas e confirmadas, botão Imprimir habilitado sem Hold, diagnóstico específico escondido na Epson e disponível na Konica. Três cliques Imprimir exclusivamente na fila virtual geraram três documentos de três páginas. Novos trabalhos funcionaram sem reiniciar e sem confirmação experimental. O workspace girado continuou pendente e exportou três páginas com a primeira a 270°.
- Inspeção Poppler das saídas virtuais confirmou visualmente texto legível, retângulo colorido e campo de formulário. A análise de objetos das saídas sintéticas encontrou vetores e zero imagens. Os textos saíram em contornos, sem caracteres extraíveis. Retângulo a 100% medido em cerca de 52,959 × 17,653 mm, para 52,917 × 17,639 mm esperados. A diferença é compatível com arredondamento do contexto a 600 dpi. Isso não homologa todos os PDFs ou RIPs.
- Chrome headless curto aprovou abas PDF 100/350, navegação/menus, exportação, proteção de alterações e QR Codes. Electron smoke aprovou isolamento, PDF.js/miniaturas, giro, salvamento/cancelamento/falha, proteção de fechamento, persistência e cópia seletiva de assinaturas.
- Sintaxe de 20 JS/MJS/CJS e dois PowerShell, compilação nativa, `git diff --check` e auditoria do ASAR instalado com 431 entradas aprovados. Regressão da preparação Hold antiga passou em FILE: no pacote final. O primeiro teste fez upgrade .11 → .12 e a revalidação reaplicou a .12 na mesma instalação. Relatórios em `.desktop-test/print12/`, ignorados. Cópias temporárias criadas pelo fluxo normal são apagadas ao concluir/cancelar. O material anterior foi preservado.

## Limitações conhecidas

O novo caminho normal ainda não foi testado fisicamente na Epson ou Konica pelo agente. A homologação física .11 continua válida para o caminho especializado preservado. Compatibilidade com todo driver Windows não foi demonstrada. Uma fila incapaz de fornecer DEVMODE/contexto válido apresenta erro, sem suporte simulado. Fax, drivers especiais e equipamentos indisponíveis não foram ensaiados.

PDFs XFA e widgets imprimíveis sem aparência pronta são recusados por risco de omissão visual; usar o leitor externo. Não se executa JavaScript do PDF para gerar conteúdo ausente. Fidelidade CMYK/ICC, sobreimpressão, transparências complexas, fontes pesquisáveis e acabamentos em outras filas ainda exigem homologação. O GDI/driver pode converter fontes em contornos e rasterizar efeitos que não suporte. A prévia é geométrica, sem simular cor, imposição ou comandos privados do fabricante. Não foi implementado intervalo de páginas ou seleção de cópias própria; o documento completo e as cópias configuradas no driver são enviados.

O instalador é experimental sem assinatura Authenticode comercial. Mantém a identidade Hold Test; não promove a variante à instalação principal. Authenticode dos atualizadores, associação PDF, preferências e demais motores continuam preservados. Instalação automática do atualizador, SMB/UNC e piloto de distribuição interna continuam com suas pendências anteriores.

## Instalação e teste pessoal

Instalador `dist-desktop/hold-native-test/4PrintShops-Hold-Test-1.0.0-experimental.12-Windows-x64-Setup.exe`, SHA256 `2890e38799f458e425cfc384978777b6ffb4dad84599d455a8170c32c0b85e53`. Mesmo appId `br.com.4printshops.holdtest`, executável, atalho, registro NSIS e perfil `%APPDATA%/4PrintShops Hold Test`. A versão principal e os instaladores anteriores foram preservados. PDFium é incluído com suas licenças; o fluxo normal dispensa Ghostscript. O diagnóstico antigo continua exigindo o interpretador local de hash fixado.

1. Exportar trabalhos pendentes, fechar a Hold Test e instalar a experimental.12 sobre ela. Conferir a versão e os atalhos existentes.
2. Abrir um PDF sintético e acessar **Arquivo → Impressão**. Selecionar `WF-C21000 Series(Rede)`. Em Propriedades, escolher conscientemente papel, orientação, cópias, cor/PB, duplex e bandeja. A sessão observada tinha A5, cinco cópias e duplex; conferir os valores antes de clicar.
3. Para o primeiro ensaio Epson, configurar uma cópia e papel A4, zoom/encaixe do driver em 100%/sem ajuste. Selecionar escala 100% no aplicativo. Conferir a prévia. **Imprimir inicia o envio**, conforme suas escolhas.
4. Conferir o papel físico, orientação, conteúdo, cores e quantidade. Repetir deliberadamente com 120% em A4 para verificar corte, depois com 87,5% e com PB escolhido no driver. Não há bloqueio geométrico.
5. Na Konica, escolher as propriedades habituais. Para o primeiro teste do caminho novo, usar Hold para poder conferir na IC-604 antes da saída física. Print e HoldPrint são permitidos e têm os efeitos definidos pelo driver. Não esperar retenção se selecionar Print.
6. Fazer outro clique Imprimir somente quando quiser outro trabalho. Conferir cópias e nomes na fila. Em erro/timeout, verificar a fila e a máquina antes de outro envio.
7. Testar posteriormente um PDF com páginas mistas, edição pendente, giro, união e exportação. Imprimir deve conservar o indicador de alterações pendentes.
8. Para comparar com a .11, expandir **Diagnóstico opcional Konica / Hold**, abrir seu sintético e usar as verificações/captura FILE: anteriores. Esse diagnóstico conserva seus limites e confirmações individuais. O leitor externo continua no menu Arquivo.

## Arquivos da etapa

Novos módulos `desktop/windows-print.mjs` e `desktop/native/WindowsPrintJob.cs`. Novos testes `tests/desktop-windows-print.test.mjs`, `desktop-windows-print-native.mjs`, `desktop-windows-print-installed.mjs` e `tests/native/WindowsPrintPolicyTests.cs`.

Integração em `desktop/app.mjs`, `preload.cjs`, `print-config.mjs`, `native/PrintConfigHost.cs`, `native/PrintConfigNative.cs`, `public/js/main.js`, `app/desktop-print-ui.js`, `app/desktop-print-layout.js`, `app/menus.js`, `tools/pdfs.js`, `scripts/build-hold-native.mjs`, `scripts/test-print-native.mjs` e `electron-builder.hold-native-test.cjs`. O teste de preparação Hold aceita destino de relatório separado para preservar evidências da .11. Os demais arquivos pendentes da .11 foram mantidos. Documentação atualizada neste guia, `ESTADO_PROJETO.md`, `DISTRIBUICAO_DESKTOP.md` e `IMPRESSAO_HOLD_EXPERIMENTAL.md`.
