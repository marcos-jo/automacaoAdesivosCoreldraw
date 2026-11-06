# Política de escala explícita no N-up — experimental.26

Data: 05/11/2026. Continuidade sobre `b08c7ce/main` com o trabalho .25 e a auditoria .24 ainda locais, integralmente preservados. Sem commit, push, instalação ou impressão nesta etapa.

## Causa confirmada

A janela .25 iniciava N-up com `escalaModo=ajustar`. Em células, o planejador calculava `min(larguraDaCelula/larguraDaArte, alturaDaCelula/alturaDaArte)`, sem teto de 100%. No compacto, o ajuste antigo já usava um teto de 100%. Portanto a mudança de dimensão vinha da política de encaixe e de seus padrões distintos, não de um deslocamento ou deformação das peças.

O caso relatado foi reproduzido exatamente usando três páginas 105×148 mm, A3 retrato 297×420, grade 2×2, margens zero, marcas externas + divisões com afastamento 3 mm, comprimento 3 mm e traço 0,1 mm. A reserva externa de 6,05 mm por borda deixa célula de 142,45×203,95 mm: o encaixe escolhe **135,6666667%**, produzindo **142,45×200,7866667 mm**. Essa é uma configuração que reproduz os valores, sem presumir outros ajustes do documento real. Sem marcas, o encaixe anterior podia ampliar ainda mais. Cópias dos fontes anteriores e reprodução em `.desktop-test/nup26/before/`, especialmente `exact-reproduction.json`.

## Decisão de escala e operação

| Política do N-up | Comportamento |
|---|---|
| **Manter tamanho real (100%) — padrão** | Mantém as dimensões físicas de cada origem. Papel, grade, margens, marcas e distribuição não escolhem outra escala. Conflitos são exibidos e bloqueiam a saída. |
| **Reduzir somente se necessário** | Ajuste proporcional limitado a 100%, inclusive ao propor Ajustar montagem para caber. Nunca amplia. Em células, limita o encaixe de cada página; no compacto, conserva o fator comum e as relações físicas do bloco. |
| **Preencher células** | Escolha explícita que permite ampliar/reduzir. Em células usa o encaixe proporcional por posição. No compacto, sem células uniformes, preenche a área disponível com o bloco proporcional; essa diferença é explicada junto ao controle. |

As políticas ficam identificadas em Configurações avançadas → Política de escala. O resumo básico sempre mostra política, medidas físicas finais e escala efetiva, além da origem/corte. Escala personalizada continua disponível como escolha avançada explícita. A opção Definir tamanho final permanece exclusiva de Step & Repeat.

Em tamanho real, as três A6 em A3 permanecem **105×148 mm a 100%**, com três peças/uma folha e quarta posição vazia, em células e compacto. Giro autorizado pode trocar os eixos das medidas, sem alterar a escala. Páginas mistas conservam as próprias dimensões e proporções. Zoom da prévia e escala do driver continuam separados da escala física do PDF.

Ampliações recebem **“AMPLIAÇÃO INTENCIONAL: confira as medidas finais antes de gerar ou imprimir”** no resumo e nas confirmações obrigatórias de ambas as saídas. As confirmações usam as escalas/dimensões do mesmo resultado que contém os bytes de produção. Recusar impede salvar ou encaminhar os bytes. Escolher Preencher não inicia download nem impressão.

O ajuste legado `ajustar` permanece aceito explicitamente na API pura para compatibilidade dos cenários anteriores: células conserva a fórmula antiga, compacto conserva seu teto antigo. **Não é o padrão e não é oferecido no seletor da janela.** Os novos controles usam `real`, `reduzir` ou `preencher`. Configurações com fator adicional explícito da API legada continuam compatíveis; na janela, um ajuste autorizado partindo de tamanho real passa para Escala personalizada, sem continuar rotulado como 100%.

## Conflitos, marcas e ajuste autorizado

- Em 100%, falta de espaço não provoca redução automática. A prévia de inspeção pode mostrar excesso com hachura e aviso, sem habilitar um PDF de produção inadequado.
- **N-up não libera conflitos de papel/margens/marcas/área imprimível consultada pelo aceite de corte.** O botão de corte fica oculto e o método também recusa essa autorização. A saída exige resolver o conflito: papel maior, revisão de margens/marcas ou ajuste explícito. O comportamento de corte deliberado do Step & Repeat foi preservado.
- Marcas individuais que atravessariam outra arte continuam bloqueadas pelo verificador existente. Compacto mantém o posicionamento aprovado e suas reservas: Reduzir/preencher do bloco não inventa espaço para marcas externas. Se as marcas ainda não couberem, o bloqueio permanece, oferecendo o ajuste explícito existente.
- Ajustar montagem para caber apresenta o percentual antes de aplicar. Quando a política é tamanho real, aceitar converte o resultado para Escala personalizada explícita, com fator adicional 1. Remover ajuste restaura a política/grade anterior e 100%. O percentual manual autorizado permanece identificado.
- A seleção automática de todas as páginas, a seleção específica validada e a preservação da prévia da .25 continuam ativas. A última folha não repete páginas. Erros de seleção não deixam estado ocupado permanente nem liberam resultado anterior para saída.

## Arquivos desta etapa

- `public/js/tools/pdf-imposition-layout.js`: padrão 100% para N-up, políticas reduzir/preencher no cálculo de escala por célula; posicionamento, rotações, cortes e Step & Repeat preservados.
- `pdf-imposition-compact.js`: mantém os caminhos anteriores; acrescenta teto de redução e ampliação somente para Preencher explícito. A busca de giro só recebe teto diferente de 100% quando essa política é escolhida.
- `pdf-imposition-bounds.js`: proposta de ajuste respeita o teto de Reduzir; demais modos mantêm os cálculos existentes.
- `pdf-imposition-settings.js`: padrão da janela N-up passa a real; estados dos modos continuam separados.
- `pdf-imposition-summary.js`: política, dimensões/percentual efetivos e alerta de ampliação na conferência.
- `pdf-imposition-ui.js`: opções/explicações, bloqueio de conflito no N-up e ajuste explícito com retorno a tamanho real.
- `electron-builder.hold-native-test.cjs`: versão .26, mesma identidade/perfil.
- Novos `tests/pdf-imposition-scale-policy.test.mjs`, `tests/pdf-imposition-scale-policy-browser.mjs` e `tests/helpers/pdf-imposition-scale-policy-browser.mjs`. Testes anteriores de encaixe passaram a solicitar redução/preenchimento explicitamente quando precisam dessas operações; não se retiraram as asserções de geometria. Testes de corte no N-up agora verificam bloqueio em vez de aceite, conforme a nova regra.
- Adaptados testes/helpers de imposição, escala, compatibilidade, seleção, interface básica e excesso; documentação atualizada. Os novos testes de seleção da .25 permanecem e foram reexecutados.

Não foram alterados o materializador vetorial `pdf-imposition-engine.js`, marcas de corte, renderizador PDF.js, PDFWorkspace, módulos de livreto .24, impressão direta, drivers/DEVMODE, associação, atualizadores ou dependências. A raiz package.json continua .6, conforme a variante Hold Test.

## Verificações executadas

- **390 testes Node aprovados**, zero falhas/cancelamentos/skips: 377 anteriores + 12 testes de políticas/PDFs + 1 proteção de saída no N-up. Casos: três A6/A3, quatro A6/A4 retrato/paisagem com giro, políticas, fontes mistas, papel insuficiente, margens, marcas e matrizes/BBox realmente serializados. Originais sintéticos preservados.
- **Chrome e Electron de desenvolvimento aprovados**, cada um com **23 PDFs** de políticas/dimensões conferidos por matrizes. Incluem 30 mudanças rápidas, ajuste explícito/remoção, conflitos sem corte, recusa das confirmações de ampliação antes de salvar/encaminhar e páginas mistas.
- Em cada ambiente, duas comparações físicas prévia/saída (células e compacto), com **16.045.152 componentes RGBA por comparação, diferença máxima zero**. Guias de tela não fazem parte da comparação. Capturas de tamanho real e páginas mistas inspecionadas visualmente.
- **Regressão de seleção .25 no Chrome aprovada:** 50 PDFs/planos com 1/4/5/8/10 páginas, 2/4/6/8-up, inválidos/recuperação, canvas preservado, ordem, edição/substituição do workspace e Step & Repeat. A bateria escolhe redução explicitamente quando seu objetivo é testar todas as páginas em grades que não comportam 100%.
- **Roteiro integrado Chrome aprovado:** montagem, excesso, escala, preparação compatível, compacto e livretos. As recusas de corte em N-up e a liberação somente após resolver o conflito foram conferidas. Livretos permanecem com paginação/prévias anteriores.
- **Electron smoke aprovado** e **41 checks C# aprovados**, sem envio de trabalhos.
- Comparação independente com os planejadores .25 copiados antes de editar: **180 planos válidos idênticos** e 12 recusas iguais. Inclui 96 configurações de Step & Repeat, 48 N-up em células e 36 compactas, com políticas explícitas anteriormente suportadas, marcas/margens/giros. Mudança do padrão e novas políticas são testadas separadamente.
- Sintaxe de 20 JS/CJS/MJS e `git diff --check` aprovados.

As primeiras asserções antigas que presumiam encaixe automático foram adaptadas para pedir essa escala explicitamente. Um erro inicial do novo ensaio de navegador ao atribuir `value` a checkbox foi corrigido para atribuir `checked`; não se alterou o giro do aplicativo para contorná-lo. Os resultados aprovados acima são das execuções finais.

Comandos, na raiz do aplicativo:

```powershell
npm test
node tests/pdf-imposition-scale-policy-browser.mjs
$env:NUP_SELECTION_ELECTRON='1'
node tests/pdf-imposition-scale-policy-browser.mjs
Remove-Item Env:NUP_SELECTION_ELECTRON
$env:NUP_SELECTION_RESULTS='.desktop-test/nup26/selection-chrome'
node tests/pdf-imposition-selection-browser.mjs
$env:IMPOSITION_DESKTOP='0'
$env:IMPOSITION_RESULTS='.desktop-test/nup26/chrome-integrated'
node tests/pdf-imposition-browser.mjs
node tests/desktop-smoke.mjs
node scripts/test-print-native.mjs
```

Evidências ignoradas em `.desktop-test/nup26/`: `node-final.log`, `browser-policy.log`, `electron-policy.log`, `selection-chrome.log`, `chrome-integrated.log`, `electron-smoke.log`, `native.log`, `preservation.json`; PDFs/capturas/provas em `chrome/`, `electron/`, `selection-chrome/` e `chrome-integrated/`. Em `chrome/scale-policy-results.json` e `electron/scale-policy-results.json` constam as configurações, medidas, matrizes e comparações de pixels. São sintéticos locais, sem documentos de cliente.

## Instalador gerado e auditado

`dist-desktop\hold-native-test\4PrintShops-Hold-Test-1.0.0-experimental.26-Windows-x64-Setup.exe`

Tamanho: **126.421.325 bytes**. SHA256: `55DEE87B64FCB95B3C5F8CA4C80659BFC86EFA70AB9F6A0B43A63C720ED51ABE`.

```powershell
$env:PRINT_GHOSTSCRIPT='none'
npx --no-install electron-builder --config electron-builder.hold-native-test.cjs --win nsis --x64 --publish never
node scripts/check-desktop-package.mjs 'dist-desktop/hold-native-test/win-unpacked/resources/app.asar'
```

Pacote de diretório e conteúdo extraído do próprio Setup aprovados: **455 entradas ASAR**, fontes iguais ao checkout, versão interna .26 e auxiliares conferidos contra manifestos. ASAR extraído/empacotado iguais, SHA256 `AE97D99489B24E7D471DC64D2A098D33B819C39DCD1D4E4CEA813505B7872077`; executáveis também iguais. Exclusões de testes/documentos/credenciais aprovadas. Prova em `.desktop-test/nup26/package-proof.json`. O 7-Zip sinaliza dados após o fim do contêiner NSIS como nas auditorias anteriores; a extração concluiu e os bytes foram auditados independentemente.

Mantidos appId `br.com.4printshops.holdtest`, produto/executável/atalhos, chave NSIS, origem `app://4printshops`, perfil `%APPDATA%/4PrintShops Hold Test` e `deleteAppDataOnUninstall:false`. O registro encontrou **.25 instalada em `%LOCALAPPDATA%\Programs\4PrintShops Hold Test`**. A .26 não foi instalada e essa instalação não foi substituída. Upgrade e conferência valor a valor das preferências aguardam execução manual. Setup experimental sem Authenticode comercial; nada publicado, oferecido ou ativado por atualizador.

## Conferência na gráfica e pendências

**Validação manual completa da experimental.26 ainda pendente na gráfica.** Este registro e os testes automatizados não representam homologação para produção. A consolidação local posterior das .24–.26 está documentada em [CHECKPOINT_EXPERIMENTAIS24_26.md](CHECKPOINT_EXPERIMENTAIS24_26.md); as referências anteriores a “sem commit” registram o momento de desenvolvimento.

Exportar trabalhos pendentes e fechar a Hold Test antes de aplicar deliberadamente o Setup. Conferir versão .26 e preferências usuais. Abrir três A6, N-up 4-up/A3: o resumo deve indicar Manter tamanho real, 100%, 105×148 mm e três peças. Trocar células/compacto e orientação; conferir medidas e quarta vaga. Em A4 com quatro A6, conferir giro autorizado, marcas e bloqueio quando faltarem reservas. Escolher Reduzir ou Preencher conscientemente, verificar percentuais/dimensões e recusar a confirmação de ampliação uma vez antes de uma eventual saída decidida pelo operador.

Não houve nenhuma impressão, física ou virtual, nem configuração de driver ou impressora alterada. Os testes de encaminhamento usam somente callback em memória. A falha de zoom do editor identificada na auditoria .24 permanece separada e não foi corrigida nem declarada resolvida; não foi repetida a suíte geral já reprovada naquele ponto como se fosse homologação.

**Os livretos da experimental.24 continuam aguardando homologação física com os funcionários da gráfica**, incluindo dobra, aninhamento, duplex, registro, refile, grampo e conferência A5/A4 e A4/A3 por equipamento/papel. Fidelidade de cor/ICC/PDF-X/RIP não recebe garantia nova. A .26 é experimental para avaliação, não homologada universalmente para produção. Nenhum commit ou push.
