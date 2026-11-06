# Interface e estabilidade — experimental.27

**Aceitação manual registrada em06/11/2026:** o usuário confirmou que os avisos compactos da .27 foram aprovados visualmente na gráfica e autorizou o checkpoint conjunto .27/.28. Essa aceitação não homologa impressão física, cor/ICC/RIP ou todos os cenários de livreto. Autoscroll inferior continua pendente. Os limites de zoom próximos às bordas permanecem sob observação: o diagnóstico abaixo explica a restrição física de rolagem reproduzida, sem afirmar que todos os possíveis problemas de zoom foram eliminados. Nenhuma correção adicional foi feita neste registro.

Data: 06/11/2026. Base `c9384f532a0f61c14749c5f849860ddaa45a2525`, branch `main`. Já havia alterações locais em `ESTADO_PROJETO.md` e `MONTAGEM_IMPOSICAO_EXPERIMENTAL26.md`, referentes à preparação/instalação .26; preservadas, com cópias e patch em `.desktop-test/ux27/before/`. Não houve reset, descarte, commit ou push.

O usuário relatou testes **visuais positivos do N-up .26** na gráfica. Livretos continuam aguardando avaliação especializada; a impressora está em manutenção e **não houve impressão física nesta etapa**. Esse relato não representa homologação de todos os tamanhos/duplex/drivers.

## Etapa 1 — Avisos compactos e enquadramento

Causa da caixa extensa: `geometryState()` montava até três mensagens por folha, concatenava as primeiras12 e exibia esse texto permanentemente. Havia repetição e as ocorrências seguintes não eram detalhadas. O novo módulo de apresentação recebe o diagnóstico pronto, sem recalcular a geometria:

- Resumo com quantidade de folhas afetadas e tipos de problema. No caso Letter215,9 ×279,4 mm em A4/100%,24 folhas: aviso de possível corte fora do papel e margens excedidas, sem reduzir a arte.
- **Ver detalhes / Ocultar detalhes**, usando o padrão de botão nativo, `aria-expanded`, painel identificado e rolagem que já existe nos avisos principais. Os detalhes ficam **na coluna esquerda**, sem ocupar a área da prévia ou cobrir navegação/ações.
- Agrupamento por tipo e conjunto de bordas, com intervalos exatos de folhas. “Folhas1–24 — esquerda, direita” conserva quais folhas têm cada problema. O modelo também mantém todas as ocorrências individuais; não há truncamento em12.
- Enter abre; Escape no painel recolhe e devolve foco ao botão, sem fechar a montagem. Instâncias possuem identificadores próprios; não foram acrescentados listeners globais.
- Avisos desaparecem ao resolver o conflito e são limpos em falha/fechamento. Alternar detalhes não inicia render nem altera resultado, escala ou aceite.
- Corte deliberado continua exigindo confirmação no Step & Repeat, agora com as ocorrências agrupadas completas também no texto de confirmação. N-up continua bloqueando conflitos e não aceita corte como liberação. Não foram modificados `canOutput`, a política de escala ou os cálculos de ajuste.

**Segundo problema de apresentação confirmado:** o enquadramento era calculado antes de os avisos/resumos ocuparem seu espaço final. Em1366×768, o canvas ficava com411px de altura, mas restavam299,36px de área visível. Foi acrescentada uma classe CSS **somente para Ver tudo / zoom relativo1**, que mantém a imagem proporcional dentro do espaço final. O canvas enquadrado passou a281,36px, com as margens, sem recorte visual. Zoom ampliado continua com rolagem. Não há segunda renderização, ResizeObserver novo, mudança de bytes, transformações PDF ou ciclo assíncrono para fazer esse ajuste.

Resultados em Chrome/Electron: aviso de63,55px para1/3/24 folhas; área de prévia299,36px no cenário exercitado; abrir detalhes não muda essa altura. Temas claro/escuro, vários padrões/bordas e rolagem nos detalhes conferidos. O tamanho real do PDF não é o zoom da imagem de tela.

## Etapa 2 — Diagnóstico exato do zoom do editor

Foi reproduzido o deslocamento histórico de **53,45px**, sem alterar `pdf-view-controls.js` ou `pdf-preview.js`:

1. Janela1440×1000; PDF sintético600×900/900×600pt, segunda página a100%.
2. Roda `deltaY=-180` no centro da visualização. Canvas final1290×860.
3. Preservar aquele ponto exigiria `scrollTop=164,45`. A extensão do conteúdo permite no máximo **111** (`scrollHeight-clientHeight`). O navegador aplica111.
4. Residual: **164,45−111=53,45px**. A posição horizontal ficou a0,17px do ideal, por arredondamento.

A causa desse cenário é **a fronteira da rolagem**, não uma altura fixa de toolbar nem destruição concorrente. O comportamento solicitado admite esses limites. Na janela1366×768, com extensão rolável suficiente, o erro foi0,30px em X/0,18px em Y. Não foi adicionado deslocamento constante, padding artificial, mudança de layout do editor ou patch ao zoom.

O teste antigo exigia erro abaixo de3px mesmo quando a posição era inalcançável. A asserção agora projeta a rolagem ideal no intervalo físico `[0, máximo]` e compara a rolagem obtida; **continua falhando se a âncora variar fora dos limites permitidos**. Testes Node reproduzem os valores geométricos e verificam caso sem limitação e botões com scroll prévio.

Chrome e Electron passaram **102 operações cada**:1366×768,1440×1000,1920×1080,1366×768 com densidade1,25 e1024×600; quatro páginas/dimensões/giros; roda nos dois sentidos, botões, seis eventos rápidos, rolagem já existente, giro no workspace e recolhimento de painéis. Conferidos também os números das páginas realmente exibidas. Erro máximo em relação à rolagem alcançável: **0,6834px em cada ambiente**, abaixo da tolerância de1,5px. Essa prova esclarece o cenário histórico; não demonstra inexistência de qualquer outro defeito de zoom em cenários não testados.

## Etapa 3 — Estabilidade preservada

- Seleção .25 no Electron:50 PDFs/planos,1/4/5/8/10 páginas, limites inválidos, preservação do último canvas, recuperação,30 mudanças rápidas, exclusão/união/substituição do workspace e Step & Repeat.
- Políticas .26 no Chrome:23 PDFs,100%, Reduzir e Preencher explícitos, dimensões/matrizes, conflitos sem liberação por corte e confirmações recusadas antes de salvar/encaminhar.
- Integração Chrome de montagem/livretos passou: erro sintético/retry, nove renderizações/oito cancelamentos da montagem, cinco/quatro no livreto, prévias de leitura/impressão, branco autorizado e cópia compatível. Nenhuma condição nova de prévia vazia indevida ou estado ocupado permanente foi reproduzida nesses cenários.
- O ciclo de seleção/cancelamento/publicação da .25 foi mantido. A apresentação não permite exportar uma prévia anterior como se fosse a configuração atual.

## Resultados e ressalvas dos testes

| Verificação | Resultado |
|---|---|
| Node completo | **399 aprovados**, zero falhas/skips/cancelamentos:390 anteriores+6 agrupamento+3 ancoragem |
| Avisos Chrome e Electron | Aprovados:1/3/24 folhas, agrupamento, teclado, rolagem, abrir/recolher repetidamente, enquadramento, zoom da montagem, temas, papel/orientação/margens/escala e segurança |
| Zoom Chrome e Electron |102 casos em cada, respeitando fronteiras; implementação de produção preservada |
| Seleção/estabilidade Electron |50 PDFs aprovados |
| Políticas .26 Chrome |23 PDFs aprovados |
| Montagem/livreto Chrome integrado | Aprovado; sem instalação/impressão |
| Electron smoke | Aprovado: abertura, PDF.js/miniaturas, giro, abas, exportação/cancelamento/falha de escrita, cópia seletiva, fechamento protegido e preferências em perfil isolado |
| Suíte geral Chrome | **Não totalmente aprovada**: após avançar além do zoom, parou na espera de autoscroll inferior nativo das miniaturas; repetiu na execução isolada |

Na suíte geral também foi corrigida uma preparação de teste: a checagem do amarelo agora recolhe Informações antes de medir sua cor. O estado expandido usa o azul de seleção já existente. Não foi alterada a interface de alertas principal para fazer o teste passar.

**Pendência explícita:** `tests/pdf-browser-smoke.mjs`, espera “Autoscroll inferior nativo”. O cenário não atingiu `scrollTop>700` dentro do prazo; os módulos de arraste/miniaturas não foram modificados. Não foi confirmada a causa e não se presume que esteja ligado aos avisos ou ao zoom. Cenários posteriores da suíte não contam como executados nessas rodadas. Não foram retiradas asserções nem aplicado contorno ao autoscroll. Os testes específicos aprovados não tornam a suíte geral inteiramente verde.

## Arquivos e preservação

Produção: novo `public/js/tools/pdf-imposition-warnings.js`; apresentação em `pdf-imposition-ui.js`; regras CSS restritas à montagem em `public/css/theme.css`; versão27 em `electron-builder.hold-native-test.cjs`.

Testes: novos `tests/pdf-imposition-warnings.test.mjs`, `tests/pdf-ui-maintenance-browser.mjs`, auxiliares `tests/helpers/pdf-imposition-warnings-browser.mjs` e `pdf-zoom-anchor-browser.mjs`; adaptados `pdf-imposition-preview.test.mjs`, `pdf-view-controls.test.mjs`, `helpers/pdf-ux-browser.mjs` e `helpers/pdf-orientation-browser.mjs`. Novos testes usam PDFs sintéticos, sem documentos de clientes.

Hashes do backup confirmaram11 módulos/motores inalterados byte a byte: planejadores/materializador/marcas de imposição, ambos os módulos de livreto, PDFWorkspace, interface de impressão e fontes nativos de geração/envio. O controle/renderer do zoom e os módulos de miniaturas/drag também permanecem sem alteração. Documentação local anterior preservada. Nenhuma dependência instalada ou removida; node_modules históricos não foram limpos.

## Reproduzir e conferir manualmente

```powershell
node --test tests/*-*.test.mjs
node tests/pdf-ui-maintenance-browser.mjs
$env:UX27_PHASE='zoom'
node tests/pdf-ui-maintenance-browser.mjs
$env:UX27_ELECTRON='1'
node tests/pdf-ui-maintenance-browser.mjs
$env:UX27_PHASE='avisos'
node tests/pdf-ui-maintenance-browser.mjs
```

As variáveis acima são apenas do processo de teste. Para mudar a pasta das provas, usar `UX27_RESULTS`. `UX27_ZOOM_DIAG=1` grava a reprodução simples sem rodar a matriz; o modo normal executa as asserções de fronteira. Perfis de desenvolvimento são isolados. `node tests/pdf-browser-smoke.mjs` mantém a falha de autoscroll registrada acima; não declarar homologação geral.

Na avaliação manual: usar Step & Repeat/Letter215,9×279,4 em A4/100%/grade1×1/24 peças, conferir resumo e todas as folhas nos detalhes; abrir/recolher pelo mouse e teclado; conferir folha inteira em Ver tudo e ampliação livre; testar papel maior para eliminar avisos. Em N-up, manter100% e verificar que conflitos continuam bloqueando saída; escolher redução/preenchimento explicitamente. Conferir prévias do livreto, sem impressão enquanto o equipamento estiver em manutenção. Observar separadamente arraste/autoscroll das miniaturas e registrar sequência/arquivo sintético se ocorrer falha.

## Instalador gerado — registro da etapa anterior à avaliação local

`dist-desktop/hold-native-test/4PrintShops-Hold-Test-1.0.0-experimental.27-Windows-x64-Setup.exe`

Tamanho:126.420.172 bytes. SHA256: `826D77F233B960F77F61191FCD4D9432EE639A00D1B02DD0F05B5CE97D8ECC09`.

Gerado pelo electron-builder26.18.0 com Electron45.1.0 e a configuração Hold Test existente. Pacote de diretório e conteúdo extraído do próprio Setup aprovados no auditor: **456 entradas ASAR**, código de produção igual ao checkout, ícones/auxiliares/PDFium/manifestos de integridade/licenças conferidos. ASAR e executável extraídos também comparados byte a byte com o pacote. O aviso do7-Zip de dados após o fim do contêiner NSIS não impediu a extração; a igualdade dos arquivos foi conferida separadamente. Sintaxe dos11 arquivos JS/CJS/MJS pertinentes e `git diff --check` aprovados.

Mesmos appId `br.com.4printshops.holdtest`, produto/executável/atalhos, chave NSIS, perfil `%APPDATA%/4PrintShops Hold Test`, origem e `deleteAppDataOnUninstall:false`. Somente o número experimental mudou. **A .27 não foi instalada**. Registro Windows reconferido: **1.0.0-experimental.26 continua instalada**. Instalador sem Authenticode comercial; nada publicado nem oferecido como atualização automática homologada. Proteções existentes não foram desativadas.

A .27 é para **avaliação controlada com a ressalva da suíte geral**, não homologação total. A instalação exige nova autorização expressa do usuário e fechamento normal/exportação dos trabalhos da .26. Sem impressão física/virtual, alterações de drivers, commit, push ou publicação.

### Evidências locais e passagem de contexto

Tudo em `.desktop-test/ux27/`, ignorado pelo Git: `before/` (cópias/patch/hash), `node-final.log`, `chrome-warnings-final.log`, `electron-warnings-final.log`, `zoom-1440-before/zoom-diagnostic.json`, `zoom-matrix/zoom-matrix.json`, `electron-zoom/zoom-matrix.json`, `electron-selection/results.json`, `policies-chrome/scale-policy-results.json`, `chrome-integrated/preview-stress.json`, provas de livreto e `electron-smoke.log`. Falhas gerais em `chrome-full.log` (premissa do indicador expandido), `chrome-full-final.log` e `chrome-full-serial.log` (autoscroll). Build/extrator em `build.log`, `extract.log` e `setup-extracted/`; auditoria consolidada em `package-proof.json`. Somente sintéticos; nenhum PDF de cliente versionado.

Próximo passo: avaliação visual da .27 após autorização para instalar, acompanhando N-up/Step & Repeat; investigação separada do autoscroll com reprodução reduzida e conferência do gesto nativo. Livretos e impressão física permanecem para avaliação especializada quando a manutenção permitir. Não acrescentar compensações de zoom, mexer em motores ou declarar o autoscroll resolvido a partir desta etapa.

Estado final do Git:16 caminhos pendentes (10 modificados/6 novos), staging vazio, `main` sobre `c9384f5`. Inclui o documento local .26 já modificado na entrada da tarefa e preservado sem mudanças adicionais nesta etapa. O servidor temporário dos testes foi encerrado após as baterias. Nenhum commit/push ou instalação .27.

## Avaliação da instalação existente — 06/11/2026

Após autorização para atualizar a .26, a inspeção encontrou a **.27 já instalada**, sem processo 4PrintShops aberto. O Setup não foi executado nesta avaliação; não se atribui a instalação anterior a esta tarefa. O SHA256 do instalador acima foi reconferido. Registro único NSIS `90ff2ff0-f013-571c-b109-7edb265dacc8`, FileVersion e versão em execução confirmaram `1.0.0-experimental.27`. Executável, ASAR, ícone PDF, auxiliares nativos e PDFium instalados são iguais aos extraídos do Setup auditado; auditoria do ASAR instalado aprovada com456 entradas.

Diretório preservado: `.desktop-test/hold-lab/installed-validation/installed/4PrintShops Hold Test`, dentro deste checkout. Mesmos appId `br.com.4printshops.holdtest`, perfil `%APPDATA%/4PrintShops Hold Test` e origem `app://4printshops`. Não há segunda instalação Hold Test registrada. As três preferências históricas da .26 continuam iguais: painéis `{"versao":1,"esquerda":312,"direita":230}`, marcador de configuração `preservado` e tema `light`. Também permaneceram iguais antes/depois dos testes atuais.

Testes na aplicação realmente instalada, com PDF sintético Letter e janela1366×768:

- Início, acesso ao módulo PDFs e importação sem alterar a origem: aprovados.
- Step & Repeat24 folhas: aviso compacto63,55px, com todas as folhas agrupadas nos detalhes. Três aberturas/recolhimentos não recriaram o canvas.
- Navegação1/24→2/24→1/24 aprovada. Zoom livre seguido de Ver tudo reenquadrou a imagem completa: canvas281,36px dentro de299,36px disponíveis; ações inferiores visíveis.
- N-up: aviso de conflito e bloqueio de saída preservados; papel maior eliminou o conflito e liberou a ação de gerar, sem acioná-la.
- Cancelamento, documento sem alterações pendentes e fechamento normal aprovados. Nenhuma exceção de execução capturada.

Diferenças observadas em relação à apresentação anterior .26: diagnóstico resumido no aviso, detalhes completos recolhíveis na coluna esquerda e enquadramento de Ver tudo preservado mesmo com o aviso visível. Não foi executada nova comparação simultânea com a .26 nem revalidada impressão nesta etapa.

Evidências ignoradas: `.desktop-test/ux27/installed-evaluation/installation-proof.json`, `interface-proof.json`, `interface.log`, capturas e inventário de hashes anterior. **Sem alteração de código, salvamento, impressão física/virtual, impressoras, autoscroll, commit ou push.** O autoscroll continua pendente conforme o diagnóstico anterior. Somente documentação atualizada;16 caminhos continuam pendentes, sem staging. Os registros de “não instalada” acima são históricos da geração e foram superados pela inspeção atual.

Para conferir manualmente: abrir Montagem com arquivo de teste, provocar aviso com papel menor, alternar Ver detalhes, percorrer folhas, ampliar e retornar a Ver tudo. Cancelar para retornar à origem. A avaliação não exige imprimir ou gerar arquivos.
