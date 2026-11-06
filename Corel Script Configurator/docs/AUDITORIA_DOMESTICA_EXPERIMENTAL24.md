# Auditoria doméstica da experimental.24

Data: 05/11/2026. Base `b08c7ce007fc72e0b86942857438b2e1eaa7357b`, branch `main`, inicialmente limpa. As .21–.24 já estavam no clone. Nenhuma funcionalidade foi reimplementada ou corrigida nesta auditoria. O usuário autorizou gerar o instalador **para avaliação com a ressalva do zoom**, após receber o resultado da suíte geral. **Não homologada para produção.**

## Ambiente e preparação

Windows x64, versão reportada pelo builder `10.0.26200`; Node `26.13.0`, npm `11.20.0`, Chrome disponível, .NET Framework x64 e OpenSSL do Git disponíveis. O guia anterior validava Node 24: os resultados abaixo foram obtidos neste ambiente doméstico com Node 26. Electron `45.1.0`, builder `26.18.0`, updater `6.8.11`, pdf-lib `1.17.1` e PDF.js `6.4.299` conforme lockfile.

Executados `npm ci`, `node node_modules/electron/install.js` e `scripts/prepare-pdfium.ps1`. O PDFium `chromium/8125` foi conferido pelo SHA256 fixado. Sem atualização do lockfile ou instalação de drivers. O npm reportou seis vulnerabilidades moderadas; não foi executado `npm audit fix` nem feita atualização de dependências. O download explícito do Electron foi necessário porque o runtime não estava presente após o npm ci.

Os 2.212 arquivos node_modules historicamente rastreados foram copiados antes do npm ci e repostos ao estado anterior após os ensaios; novas dependências locais continuam ignoradas. Não houve descarte de alterações do usuário: a árvore inicial estava limpa. Antes das mudanças documentais finais, o Git voltou a ficar limpo.

## Resultados desta execução

| Verificação | Resultado e alcance |
|---|---|
| `npm test` | **362 aprovados**, zero falhas, cancelamentos ou skips. Inclui os testes existentes de livretos e regressões .21–.23. |
| `node scripts/test-print-native.mjs` | **41 checks C# aprovados**: 6 DEVMODE, 17 Hold, 7 política de padrão, 11 impressão Windows. Nenhum envio. |
| `IMPOSITION_DESKTOP=0 node tests/pdf-imposition-browser.mjs` | **Aprovado integralmente**, incluindo Step & Repeat, células/compacto, excesso, escalas, preparação segura e livreto. |
| Auditoria adicional do motor | **56 combinações aprovadas**: A5/A4 e A4/A3 × 4/8/12/16/5/7/10 páginas × retrato/paisagem × verso normal/180°. |
| `node tests/desktop-smoke.mjs` | **Aprovado**, Electron real com perfil temporário, sandbox, PDF.js/worker, edição, salvamento/cancelamento/falha, proteção e persistência. |
| Livreto no Electron | **Aprovado**, reutilizando o roteiro existente com bootstrap de desenvolvimento e perfil isolado em `.desktop-test/home24/electron-booklet/profile`. Adicional A4/A3 com impressão/leitura/retrato aprovado. Nenhuma impressão ou instalação. Não equivale a teste de aplicativo instalado. |
| `node tests/pdf-browser-smoke.mjs` | **Reprovado duas vezes**, na asserção de ancoragem do zoom do editor. A suíte parou nessa etapa; cenários posteriores não contam como executados nessas rodadas. |
| `PDF_TESTE_ASSINATURAS_APENAS=1 node tests/pdf-browser-smoke.mjs` | **Aprovado**, teste dedicado de assinaturas/CMS e cópias. Não substitui a suíte geral reprovada. |
| Exemplos | **42 PDFs gerados e renderizados integralmente**, mais mapas JSON. Inspeção visual de amostras impostas A5/A4 e A4/A3, com e sem brancos, e das duas prévias da interface. |
| Pacote/Setup | **Aprovados na auditoria estática**; 454 entradas ASAR e conteúdo extraído do Setup igual aos fontes atuais. Instalador não executado. |
| Microsoft Print to PDF | **Não executado**: gerenciamento automático do padrão ativo, `LegacyDefaultPrinterMode=0`, condição bloqueada pelo projeto. |

### Paginação, integridade e fidelidade

| Originais | Brancos finais | Total lógico | Folhas físicas | Faces PDF |
|---:|---|---:|---:|---:|
| 4 | nenhum | 4 | 1 | 2 |
| 8 | nenhum | 8 | 2 | 4 |
| 12 | nenhum | 12 | 3 | 6 |
| 16 | nenhum | 16 | 4 | 8 |
| 5 | 6, 7, 8 | 8 | 2 | 4 |
| 7 | 8 | 8 | 2 | 4 |
| 10 | 11, 12 | 12 | 3 | 6 |

A auditoria adicional verificou os bytes/SHA256 da origem antes/depois, dimensões físicas, escala 100%, número de faces/folhas, brancos, ordem de leitura, pares de frente/verso e correspondência geométrica das metades para os dois eixos de virada. Matrix/BBox realmente serializados foram conferidos. O texto extraído de cada face corresponde às páginas de origem esperadas. Nas fixtures vetoriais, cada página de origem permanece como Form com fontes, sem imagens de página inteira. Não é promessa de que imagens de originais reais se tornem vetores.

Oito páginas: frente externa `[8 | 1]`, verso externo `[2 | 7]`, frente interna `[6 | 3]`, verso interno `[4 | 5]`. Quatro páginas: `[4 | 1]` e `[2 | 3]`.

Na prova de sete páginas, tanto Chrome quanto Electron compararam **18 pares / 10.953.600 componentes RGBA, diferença máxima zero**, entre origem, faces, cópia de inspeção e leitura. Branca 8 sem conteúdo; sete números de página extraíveis. Original/workspace pendente preservado. Trocas de visualização, papel/orientação, marcas/escala, aceite de brancos, cancelamento/reentrada, preparação compatível, falha sintética/retry e atualizações rápidas passaram. Livreto: cinco renders/quatro cancelamentos; montagem existente: nove/oito. Não ocorreu desaparecimento espontâneo da prévia nesses roteiros. A falha intermitente da .23 continua aberta.

### Problema novo observado: zoom do editor

As duas rodadas gerais falharam em `tests/helpers/pdf-ux-browser.mjs`, asserção “Zoom conserva o ponto sob o cursor, salvo arredondamento do canvas”. Reprodução: documento sintético com páginas 600×900 e 900×600 pontos; navegar à segunda, selecionar 100%, roda `deltaY=-180` no centro da área. Tolerância do teste: menos de 3 pixels.

Uma terceira rodada instrumentada em cópia temporária do teste mediu **53,45 pixels** de deslocamento, inclusive após aguardar mais 500 ms. Canvas final 1290×860, scrollLeft 162. Isso não demonstra que seja apenas leitura antecipada do teste. Não foi estabelecida causa definitiva, nem relação com o desaparecimento da prévia .23. O sintoma foi observado fora da janela de livreto.

**Correção proposta para etapa separada:** reproduzir isoladamente com posições X/Y, limites de rolagem e gerações do render/ResizeObserver instrumentados; verificar o cálculo de âncora em `pdf-view-controls.js` e seu retorno após atualizar a prévia; adicionar uma regressão que falhe antes da correção e conferir zoom/pan/encaixe. Nenhuma alteração de produção ou refatoração foi realizada. O usuário autorizou o pacote para avaliação com esta ressalva; a suíte geral continua registrada como reprovada.

## Exemplos e evidências locais

Raiz dos comandos: `Corel Script Configurator`.

- A5/A4: `.desktop-test/booklet24/examples/`. Para cada quantidade, `Origem N paginas A5.pdf`, `Livreto N paginas A4 paisagem.pdf`, `Conferencia leitura N paginas.pdf`, `Paginacao N.json`. Gerador original `node scripts/generate-booklet-examples.mjs`.
- A4/A3: `.desktop-test/home24/examples-a4-a3/`. Mesmos sete totais: `Origem N paginas A4.pdf`, `Livreto N paginas A3 paisagem.pdf`, conferência e mapa. Auditoria/gerador local `.desktop-test/home24/audit-engine.mjs`.
- Renderizações de todos os PDFs: `.desktop-test/home24/renders/`; inventário `renders.json`. Renderizador PDF.js com canvas nativo, em vez de Poppler, não disponível neste ambiente. Adaptação somente do ensaio Node para fornecer Path2D/DOMMatrix/ImageData, sem mudança do aplicativo.
- Resultados: `.desktop-test/home24/engine-audit.json`, `chrome-integrated/livreto-provas.json`, `chrome-integrated/livreto-fidelidade.json`, `electron-booklet/livreto-provas.json`, `electron-booklet/electron-booklet-extra.json`.
- Logs: `.desktop-test/home24/{node-tests,chrome-integrated,electron-smoke,electron-booklet,browser-full,browser-full-recheck,zoom-diagnostic,browser-signatures,installer-build,installer-extract}.log`.
- A conferência em ordem de leitura **não é o PDF de produção**. Para avaliação de impressão futura, usar os arquivos com prefixo `Livreto`.

## Instalador gerado nesta máquina

`dist-desktop\hold-native-test\4PrintShops-Hold-Test-1.0.0-experimental.24-Windows-x64-Setup.exe`

Tamanho: **126.418.429 bytes**. SHA256: `4E62F15ADA01DDFFAC99EE0337FA894CA96ABD6022A501C86092B938A1EC2115`.

Build existente, sem mudar a configuração:

```powershell
$env:PRINT_GHOSTSCRIPT='none'
npx --no-install electron-builder --config electron-builder.hold-native-test.cjs --win nsis --x64 --publish never
node scripts/check-desktop-package.mjs 'dist-desktop/hold-native-test/win-unpacked/resources/app.asar'
```

Mantidos appId `br.com.4printshops.holdtest`, produto/executável `4PrintShops Hold Test`, origem `app://4printshops`, perfil `%APPDATA%/4PrintShops Hold Test`, instalação por usuário e `deleteAppDataOnUninstall:false`. O ASAR contém `version:1.0.0-experimental.24`, `name:4printshops-hold-test` e as flags Hold Test/impressão Windows. O manifest principal permanece .6 por desenho da variante. Ghostscript não acompanha o pacote; diagnóstico PS/Hold sem interpretador permanece bloqueado.

Conferidos o pacote de diretório e **o conteúdo extraído do próprio Setup**, sem executá-lo: ASAR com 454 entradas, fontes de produção idênticos ao clone, auxiliares/PDFium iguais aos manifestos, ícones/licenças, ausência de testes/documentos/credenciais. ASAR extraído e empacotado: SHA256 `CE98C7E6A2BB12715AB95E32FB5AE54C57EA2F78F45315354F1FC16D4F82A95C`; executável do aplicativo extraído/empacotado também idêntico. O 7-Zip sinalizou dados após o fim do arquivo ao ler o contêiner autoextraível, mas concluiu a extração com “Everything is Ok”; o conteúdo foi conferido separadamente. Prova adicional em `.desktop-test/home24/package-proof.json`.

Setup **sem assinatura Authenticode**, conforme build experimental. Não comparar seu hash com o histórico gerado na gráfica como se fossem o mesmo binário. A identidade/configuração foi preservada; **não foi executada instalação/upgrade nesta máquina nem testada migração real de um perfil anterior**. A chave NSIS Hold Test consultada e a pasta de perfil correspondente não existiam. O aplicativo principal e a impressora padrão Adobe PDF foram preservados. Não houve publicação ou oferta de atualização.

## Limitações e próximo passo

Somente um caderno grampo canoa, leitura esquerda→direita e brancos finais até múltiplo de quatro. Sem múltiplos cadernos, Cut & Stack, lombada colada ou creep. Limite técnico 2.000 páginas e PDF 256 MiB não é capacidade recomendada de acabamento. A5 nominal 148×210 requer considerar 148,5×210 após dobrar A4 antes do refile. Marcas/sangrias podem exigir papel maior; a área imprimível é não verificada. Recursos incompatíveis continuam bloqueados ou exigem cópia com consentimento e prova de fidelidade. Não há nova homologação universal de PDF/X, ICC, cores especiais, sobreimpressão ou RIP.

Não foi enviada nenhuma impressão, nem mesmo virtual. Microsoft Print to PDF existe, mas `LegacyDefaultPrinterMode=0` bloqueia os ensaios seguros do projeto; não se alterou o registro. Não foram instalados drivers/fila/portas da gráfica. Ausentes Konica/Epson, sem tentativa de reproduzi-las. Nenhuma instalação automática, commit ou push.

Próximo passo: avaliação doméstica deliberada do Setup, com a ressalva de zoom conhecida; depois prova física de quatro páginas na gráfica e homologação separada A5/A4 e A4/A3: topo/base, eixo duplex, alimentação manual quando aplicável, registro, ordem/aninhamento, dobra, refile, grampo, gramatura, bandeja, margens/área imprimível e cores/RIP. Iniciar investigação do zoom em etapa própria. A .24 permanece experimental, **não homologada para produção**.
