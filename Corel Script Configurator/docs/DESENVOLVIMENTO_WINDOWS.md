# Desenvolvimento entre a gráfica e casa

## Motor Word privado — experimental.32 (08/11/2026)

Os builders que herdam `electron-builder.config.cjs` agora exigem o LibreOffice privado preparado, além das dependências anteriores. Depois de `npm ci`, executar `node scripts/prepare-libreoffice.mjs --download` no Windows x64. O preparador verifica o MSI oficial e quatro arquivos de código-fonte, faz extração administrativa sem instalar Office e mantém runtime/cache ignorados no Git. Um clone novo não inclui esses binários. Reservar vários GB livres. O aplicativo distribuído usa sua própria cópia; desenvolvimento só procura Office externo se a pasta privada estiver ausente. Teste do pacote sem imprimir: `node scripts/test-private-office.mjs dist-desktop/hold-native-test/win-unpacked`. Detalhes em [LIBREOFFICE_PRIVADO_EXPERIMENTAL32.md](LIBREOFFICE_PRIVADO_EXPERIMENTAL32.md); as seções .12 abaixo são históricas e preservam os requisitos nativos.


## Checkpoint experimental.12

Em 04/11/2026, o usuário informou resultados práticos satisfatórios da .12 em outras impressoras da gráfica. A .11 já havia sido aprovada na Konica C6100/IC-604. Preservar esses resultados e ampliar a matriz de homologação por driver/opção nos próximos trabalhos.

A auditoria começou sobre `0af565f`, branch `main`, staging vazio e 32 arquivos alterados/novos relacionados à .11/.12. Cópia SHA256 preservada em `.desktop-test/handoff/baseline-20261104-145309/`. O remote de leitura e envio é `https://github.com/marcos-jo/automacaoAdesivosCoreldraw.git`. A API GitHub confirmou `visibility=public`, `private=false` e branch padrão `main`. A consulta foi somente de leitura. Nenhuma configuração ou credencial GitHub foi alterada.

O commit desta preparação contém os fontes de impressão .11/.12, testes, geradores e documentação. Instaladores, PDFs, spool/PS, binários de drivers, certificados, chaves e dados DPAPI permanecem fora do Git. As dependências e ferramentas são preparadas novamente em cada máquina.

## Clonar e abrir a pasta correta

No Git Bash ou PowerShell, depois de o usuário realizar o push na gráfica, executar em uma pasta de trabalho escolhida em casa.

```sh
git clone https://github.com/marcos-jo/automacaoAdesivosCoreldraw.git
cd "automacaoAdesivosCoreldraw/Corel Script Configurator"
git status --short
git log -1 --oneline
```

A raiz Git é `automacaoAdesivosCoreldraw`. O aplicativo Node e seu `AGENTS.md` ficam na subpasta **Corel Script Configurator**. Os comandos de npm, scripts e testes abaixo partem dessa subpasta. Conferir o hash do checkpoint informado na entrega.

## Ambiente validado

| Componente | Versão/requisito |
| --- | --- |
| Sistema | Windows x64 com spooler disponível para testes nativos |
| Node.js | 24.18.0 validado |
| npm | 11.16.0 validado |
| Git | 2.53.0.windows.2 na gráfica |
| Electron | 45.1.0 fixado no projeto |
| electron-builder | 26.18.0 fixado |
| electron-updater | 6.8.11 fixado |
| pdf-lib / PDF.js | Instalar as versões do `package-lock.json` com npm ci |
| .NET Framework | 4.8.1 na gráfica, Release 533325, versão de runtime 4.8.09037 |
| Compilador C# | Framework x64, versão 4.8.9232.0, linguagem C# 5 |
| OpenSSL | 3.5.5 validado para os testes CMS reais |
| Chrome ou Edge | Instalado para Chrome headless, com `BROWSER_PATH` quando necessário |
| PDFium | `chromium/8125`, Windows x64, URL/SHA256 em `build/pdfium-lock.json` |

Os scripts C# usam `%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\csc.exe`, `System.Web.Extensions`, `System.Xml.Linq` e assemblies WPF `System.Printing`/`ReachFramework`. O SDK .NET moderno não participa desse build. O Framework e seus assemblies devem estar disponíveis. A chave Release 533325 corresponde ao Framework 4.8.1. [Microsoft — versões do Framework](https://learn.microsoft.com/en-us/dotnet/framework/install/versions-and-dependencies)

Consultas PowerShell para conferir o ambiente, sem modificar o Windows.

```powershell
node --version
npm --version
git --version
Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -Name Release,Version
Test-Path "$env:SystemRoot\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
Test-Path "$env:SystemRoot\Microsoft.NET\Framework64\v4.0.30319\WPF\System.Printing.dll"
```

## Instalar dependências e executar

```sh
npm ci
node scripts/build-print-helper.mjs
npm start
```

A versão web abre em `http://localhost:3000`. Parar o servidor com Ctrl+C quando terminar. Em outro terminal, executar o Desktop em desenvolvimento.

```sh
npm run desktop
```

`npm ci` usa o lock existente. A primeira preparação precisa de acesso às distribuições npm/Electron, ou caches previamente preparados pelo operador. Electron e suas dependências são locais do projeto. Se o download do runtime Electron tiver sido impedido, verificar a causa e executar `node node_modules/electron/install.js` após resolvê-la.

O `package.json` principal permanece na base experimental.6. As flags da impressão geral pertencem à variante **Hold Test .12**. Para abrir exatamente essa variante sem instalar, gerar o pacote de diretório conforme a seção de build abaixo. `npm run desktop` inicia a configuração principal e conserva essa separação de identidades.

## Testes que funcionam sem Konica ou Epson

```sh
npm test
node scripts/test-print-native.mjs
node tests/desktop-smoke.mjs
```

`npm test` inclui assinaturas CMS sintéticas com OpenSSL. Configure `OPENSSL_PATH` caso o executável não seja encontrado nos locais usuais do Git for Windows ou no PATH. As chaves/certificados são efêmeros de teste, criados em pasta temporária, sem instalação de confiança. O gerador remove sua chave privada ao finalizar. Nenhum desses artefatos deve ser versionado.

```powershell
$env:OPENSSL_PATH = 'C:\Program Files\Git\mingw64\bin\openssl.exe'
& $env:OPENSSL_PATH version
npm test
```

Os 41 checks C# exercitam estado DEVMODE, 8192 bytes privados, cancelamento, revisão, Hold/PrintTicket, política da impressora padrão e escala, sem abrir filas ou enviar trabalhos. O Electron smoke usa diálogos de teste e PDFs sintéticos, sem impressão física.

Para o teste curto do navegador, executar no PowerShell.

```powershell
$env:APP_TESTE_ABAS_APENAS = '1'
node tests/pdf-browser-smoke.mjs
Remove-Item Env:APP_TESTE_ABAS_APENAS
```

`npm run test:browser` executa a suíte de navegador completa. `BROWSER_PATH` aceita o caminho do Chrome/Edge local. Os testes que incorporam fontes usam Arial do próprio Windows ou DejaVu no Linux. Essas fontes não são copiadas para o Git. O fluxo de impressão nativa e sua compilação requerem Windows.

## Gerar os PDFs de teste

```sh
node scripts/generate-print-fixtures.mjs
node --test tests/print-fixtures.test.mjs
```

Os fontes seguros ficam em `tests/helpers/windows-print-fixtures.mjs`. O comando gera em `.desktop-test/fixtures/windows-print/` cinco arquivos e um manifesto com SHA256.

| Arquivo gerado | Conteúdo |
| --- | --- |
| `a4.pdf` | A4 com textos e retângulo vetorial de 50 × 20 mm |
| `a3.pdf` | A3 com os mesmos elementos de teste |
| `a4-rotated-90.pdf` | A4 com rotação de 90° |
| `mixed-pages.pdf` | A4 girado a 180°, 300 × 120 mm e 100 × 150 mm |
| `form-appearance.pdf` | Campo de formulário com aparência pronta |

Datas/metadados dos geradores Node são fixos. Duas execuções produzem os mesmos bytes, conferidos também em uma pasta independente com espaços/acentos e sem node_modules. Um arquivo existente com conteúdo diferente é preservado e causa erro. Para outra pasta, informar seu caminho como argumento.

As escalas 100%, 120%, 87,5% e demais percentuais são aplicadas na impressão ou na prévia. O PDF original da fixture mantém suas dimensões. Testes unitários cobrem valores decimais e limites 1–400%, precedência do driver e áreas imprimíveis.

O sintético exclusivo do diagnóstico Hold continua sendo produzido pelo **gerador C# versionado `desktop/native/PdfiumFixture.cs`**, compilado por `build-print-lab.mjs`. Ele incorpora Liberation Sans proveniente da dependência pdfjs-dist, com sua licença copiada para o laboratório. Seus objetos e aparência são os mesmos usados na gráfica. IDs internos podem variar entre gerações. O reconhecimento de segurança compara aparência/objetos e usa o manifesto de cada build. Não substituir esse sintético por um PDF genérico para passar pelo diagnóstico Hold.

## Preparar PDFium e testar impressão somente virtual

Executar no PowerShell. O download PDFium é conferido pelo SHA256 fixado antes da extração. As licenças acompanham os recursos gerados.

```powershell
.\scripts\prepare-pdfium.ps1
node scripts/build-print-lab.mjs
$env:PRINT_GHOSTSCRIPT = 'none'
node scripts/build-hold-native.mjs
$env:PRINT_VIRTUAL_ONLY = '1'
node tests/desktop-windows-print-native.mjs
Remove-Item Env:PRINT_VIRTUAL_ONLY
Remove-Item Env:PRINT_GHOSTSCRIPT
```

Esse modo consulta exclusivamente **Microsoft Print to PDF** e grava PDFs em `.desktop-test/print12/native/`. Não precisa das filas Epson/Konica. Verifica três páginas mistas, giros, escalas 100/120/87,5, campo de formulário, PDF inválido, revisão antiga, token cancelado e preferências. Trabalhos consecutivos são enviados apenas à fila virtual, com destino de arquivo explícito.

A fila virtual precisa existir previamente. Consultar `Get-Printer` para conferir. O ensaio recusa captura quando o gerenciamento automático da impressora padrão estiver ativo ou inconclusivo. É possível consultar a configuração abaixo, sem modificá-la.

```powershell
Get-Printer | Select-Object Name,DriverName,PortName
Get-ItemProperty 'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows' -Name LegacyDefaultPrinterMode
```

Se a fila estiver ausente ou a proteção do padrão bloquear, interromper o teste e decidir manualmente a preparação do Windows. Os scripts não criam impressoras, alteram portas, instalam drivers ou modificam essa preferência. O teste seguro exige `LegacyDefaultPrinterMode=1`. Não foi automatizada sua aplicação na casa.

## Build e instaladores

Para a identidade principal, com auxiliar de configuração e recursos próprios.

```sh
npm run desktop:pack
npm run desktop:dist
```

Para a identidade Hold Test com impressão geral .12, preparar antes PDFium e o laboratório nativo. No PowerShell, escolher explicitamente `PRINT_GHOSTSCRIPT=none` quando desejar somente impressão geral e testes virtuais.

```powershell
$env:PRINT_GHOSTSCRIPT = 'none'
npx --no-install electron-builder --config electron-builder.hold-native-test.cjs --win --x64 --dir --publish never
& '.\dist-desktop\hold-native-test\win-unpacked\4PrintShops Hold Test.exe'
```

O comando `--dir` gera o aplicativo executável sem instalar. Para gerar o Setup NSIS, quando necessário.

```powershell
npx --no-install electron-builder --config electron-builder.hold-native-test.cjs --win nsis --x64 --publish never
Remove-Item Env:PRINT_GHOSTSCRIPT
```

A variante mantém appId `br.com.4printshops.holdtest` e perfil `%APPDATA%/4PrintShops Hold Test`. O Setup atualiza essa identidade quando ela já estiver instalada. A identidade principal permanece separada. Os builders nunca enviam trabalhos à impressora. Compilações locais sem certificado geram pacotes experimentais sem assinatura comercial. Não publicar esses pacotes como atualizações comerciais.

O ensaio `tests/desktop-windows-print-installed.mjs` reproduz a instalação/upgrade executada na gráfica, exige a instalação de laboratório registrada no caminho relativo conhecido, a preferência `4printshops.print-config-check=preservado`, as filas da gráfica e o Setup .12 gerado. Não é um instalador inicial para casa. Recusa um diretório registrado diferente. Os ensaios históricos .9/.10/.11 também exigem seus pacotes e contexto de laboratório. Continuam versionados para uso controlado, sem copiar os instaladores antigos para o Git. Em casa, começar com Node/C#, geradores, Electron smoke, modo nativo somente virtual e o executável `--dir`.

## Ghostscript e diagnóstico específico Konica

A impressão Windows geral dispensa Ghostscript. Com `PRINT_GHOSTSCRIPT=none` ou com o caminho legado ausente, o build compila com confiança vazia para o interpretador e o diagnóstico PS/Hold permanece bloqueado. A verificação de hash em runtime foi preservada.

Para estudar PS, obter manualmente um executável Windows x64 oficial após conferir sua origem e licença, e informar seu caminho antes do build.

```powershell
$env:PRINT_GHOSTSCRIPT = 'C:\Ferramentas\Ghostscript\bin\gswin64c.exe'
node scripts/build-hold-native.mjs
```

O caminho deve ser absoluto. Um caminho explícito inválido causa erro. O build incorpora o SHA256 exato desse executável e o runtime o exige. Trocar o interpretador requer novo build e nova validação dos resultados. A descoberta legada de 10.01.2 permanece para reproduzir a estação da gráfica. Essa versão antiga foi usada somente nos sintéticos da investigação. Não é requisito para o ambiente doméstico nem uma nova homologação de versões do interpretador.

Ghostscript oferece distribuições AGPL e comerciais. Downloads, licença de uso e eventual redistribuição precisam ser avaliados antes de incluí-lo em um produto comercial. O projeto não inclui o binário ou uma licença comercial de Ghostscript. PDFium possui licença BSD e avisos das dependências, copiados pelo preparador. Não versionar binários Epson/Konica nem arquivos PS de driver. [Ghostscript — releases oficiais](https://www.ghostscript.com/releases/index.html), [Ghostscript — licenciamento](https://www.ghostscript.com/licensing/index.html)

Os ensaios Konica `desktop-hold-native.mjs`, `desktop-hold-prepare.mjs` e `desktop-konica-capture.mjs` dependem do driver de referência, seus recursos OEM, fila Konica e/ou fila adicional FILE: autorizada, além de Ghostscript quando há interpretação. O builder de captura isolada é `scripts/build-konica-capture.mjs`. Os scripts não criam essas filas. A ausência desses recursos em casa é uma limitação conhecida, sem simulação de prova física de Hold. Manter o diagnóstico bloqueado enquanto não houver o ambiente adequado.

## Matriz de reprodução

| Grupo | Em casa sem impressoras físicas | Dependências adicionais |
| --- | --- | --- |
| Node / workspace / escala / prévia / coordenador | Sim | npm ci, OpenSSL e fontes locais para os testes CMS/fontes |
| C# estado DEVMODE / Hold / política | Sim | Framework x64 e assemblies |
| Geradores A4/A3/mistos/formulário | Sim | Node e pdf-lib versionado |
| Navegador / Electron smoke | Sim | Chrome/Edge ou Electron baixado |
| PDFium/HDC somente virtual | Sim | PDFium preparado, Microsoft Print to PDF e proteção de padrão satisfeita |
| Propriedades Epson/Konica | Depende de drivers e filas | Instalação manual autorizada e configuração equivalente |
| Captura PS/Hold Konica | Depende do laboratório de driver | FILE: autorizada e interpretador fixado |
| Instalação/upgrade e preferências | Exige laboratório correspondente | Setup gerado, registro/perfil de referência e consentimento |
| Papel físico, cor/PB, bandeja, grampo, retenção IC-604 | Não | Equipamento e operador acompanhando cada teste |

## O que não é transportado pelo Git

- node_modules, cache npm/Electron/PDFium, builds, instaladores, `.desktop-test`, PS/spool e relatórios locais. Regenerar pelos scripts.
- Drivers, filas, portas, status de opcionais, DEVMODE privado e preferências de impressora. São específicos do Windows/equipamento.
- Perfis em APPDATA, localStorage, formatos personalizados salvos, documentos e alterações ainda em memória. Exportar os PDFs antes de mudar de computador. Se quiser migrar preferências, planejar essa cópia separadamente e fechar os aplicativos.
- Certificados Authenticode, chaves privadas, senhas e material de confiança local. Não copiar para o repositório.
- A chave Ed25519 do publicador assistido, protegida por DPAPI, é vinculada ao usuário/Windows que a protegeu. O código e as chaves públicas são versionados. A migração/recuperação de identidade de publicação exige procedimento próprio. Os testes unitários usam identidades efêmeras. Não regenerar a chave de produção para tentar substituir a chave existente.
- Conversas, skills, plugins, credenciais e configurações pessoais do Codex. O contexto reproduzível do projeto está nos arquivos versionados.

## Continuidade no VS Code com Codex

Instalar o VS Code e a extensão oficial Codex, autenticar-se e abrir a pasta `Corel Script Configurator`. Abrir o painel Codex e começar uma conversa com o contexto local do repositório. A documentação oficial descreve a instalação e a abertura do painel no editor. [OpenAI Docs — extensão Codex](https://learn.chatgpt.com/docs/codex/ide)

Pedir primeiro a leitura de `AGENTS.md`, `docs/ESTADO_PROJETO.md`, deste guia e de `docs/IMPRESSAO_WINDOWS_GERAL.md`. Conferir Git e executar os testes locais antes de mudar funcionalidades. Informar quais drivers e filas existem nessa máquina. A aprovação física da gráfica continua registrada, e os testes domésticos devem relatar seu próprio alcance.

Antes de alternar de computador, concluir o trabalho em uma máquina, criar o checkpoint autorizado e realizar o push manual. Na outra máquina, com a árvore de trabalho limpa, atualizar usando `git pull --ff-only origin main`. Se houver alterações locais ou divergência, revisar antes de prosseguir. Evitar editar a mesma etapa nos dois computadores sem sincronizar.

O envio deste checkpoint será feito pelo usuário, a partir de qualquer pasta dentro do repositório.

```sh
git status --short
git log -1 --oneline
git push origin main
```

O agente não executou esse push. As regras de aprovação de commits e de proteção de trabalhos continuam em `AGENTS.md`.

## Inventário do checkpoint

Validação desta preparação aprovada com 240 testes Node, 41 verificações C#, Chrome headless curto, Electron smoke e auditoria do pacote .12 instalado. O teste nativo somente virtual passou com o auxiliar compilado sem Ghostscript. A fixture também foi reproduzida em outra pasta sem node_modules, com os mesmos cinco hashes. Nenhum instalador novo foi gerado. O agente não enviou trabalhos físicos nem alterou configurações do Windows.

O conjunto auditado contém 37 caminhos. Os geradores e os fontes são versionados. Os PDFs resultantes permanecem ignorados.

- `.gitignore`
- `desktop/app.mjs`
- `desktop/hold-test.mjs`
- `desktop/native/HoldExperimentalJob.cs`
- `desktop/native/KonicaPostScriptCapture.cs`
- `desktop/native/PrintConfigHost.cs`
- `desktop/native/PrintConfigNative.cs`
- `desktop/native/WindowsPrintJob.cs`
- `desktop/preload.cjs`
- `desktop/print-config.mjs`
- `desktop/windows-print.mjs`
- `docs/DESENVOLVIMENTO_WINDOWS.md`
- `docs/DISTRIBUICAO_DESKTOP.md`
- `docs/ESTADO_PROJETO.md`
- `docs/IMPRESSAO_HOLD_EXPERIMENTAL.md`
- `docs/IMPRESSAO_WINDOWS_GERAL.md`
- `electron-builder.hold-native-test.cjs`
- `public/js/app/desktop-print-layout.js`
- `public/js/app/desktop-print-ui.js`
- `public/js/app/menus.js`
- `public/js/main.js`
- `public/js/tools/pdfs.js`
- `scripts/build-hold-native.mjs`
- `scripts/generate-print-fixtures.mjs`
- `scripts/test-print-native.mjs`
- `tests/desktop-hold-native-installed.mjs`
- `tests/desktop-hold-prepare.mjs`
- `tests/desktop-hold-test.test.mjs`
- `tests/desktop-print-config.test.mjs`
- `tests/desktop-windows-print-installed.mjs`
- `tests/desktop-windows-print-native.mjs`
- `tests/desktop-windows-print.test.mjs`
- `tests/helpers/windows-print-fixtures.mjs`
- `tests/helpers/windows-print-properties.ps1`
- `tests/helpers/windows-save-dialog.ps1`
- `tests/native/WindowsPrintPolicyTests.cs`
- `tests/print-fixtures.test.mjs`
