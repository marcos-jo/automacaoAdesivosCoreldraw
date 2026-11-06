# N-up com todas as páginas e seleção validada — experimental.25

Data: 05/11/2026. Base `b08c7ce`, branch `main`. As três alterações documentais da auditoria doméstica .24 já estavam pendentes e foram preservadas. Nenhum commit, push, instalação ou impressão foi executado nesta etapa. A experimental.24 dos livretos **continua aguardando homologação física com os funcionários da gráfica**; os testes abaixo não substituem essa avaliação.

## Diagnóstico reproduzido antes da correção

Reproduzido no Chrome com documentos sintéticos de **1, 4, 5, 8 e 10 páginas**: gerar uma prévia válida no N-up e informar `1-(total+1)` no campo de origem. Nos cinco casos, a prévia válida desaparecia, `result` ficava nulo e `busy` permanecia verdadeiro após a mensagem de seleção inválida. A correção posterior da entrada ainda podia recuperar a janela, mas o estado aparentava travamento.

A causa confirmada está na janela: `schedule()` descartava o resultado e zerava o canvas **antes** da validação; a faixa era expandida sem conhecer o total do PDF; o motor corretamente rejeitava a seleção fora dos limites; o tratamento do erro mantinha os controles no estado ocupado. Portanto não era falta de proteção de índices no motor. A reprodução não detectou exceções não tratadas. Prova anterior e cópias dos fontes em `.desktop-test/nup25/before/reproduction.json`, `ui-before.js` e `settings-before.js`.

Esse é um mecanismo concreto para o sintoma relatado. Não se conclui que todas as possíveis combinações responsáveis pelo desaparecimento ocasional da prévia .23 foram identificadas.

## Operação da .25

- **N-up padrão: todas as páginas reais**, na ordem original. A contagem vem dos bytes atuais capturados do PDFWorkspace, não de um número digitado pelo operador. O Básico mostra o total, papel/orientação, grade 2/4/6/8-up ou personalizada e distribuição tradicional/compacta. Não exige quantidade nem números de páginas.
- Cada página é distribuída uma vez. A última folha permanece incompleta quando necessário; não duplica a última página nem volta ao início para preencher vagas.
- **Configurações avançadas → Páginas do N-up** permite escolher Todas ou Selecionar páginas específicas. O campo só aparece na segunda opção. Aceita `1,3-5`, mantendo a ordem explicitamente indicada, inclusive repetições se o operador as escrever. Não ordena, elimina ou substitui páginas silenciosamente.
- Página inexistente, zero, negativo, texto não numérico, decimal, intervalo invertido, segmento vazio e seleção vazia são recusados antes de chamar o gerador. Mensagens identificam o limite atual quando pertinente. Zero não é usado para reservar vagas nessa interface; o suporte legado a posições `null` da API geométrica não foi alterado.
- Uma seleção inválida preserva o último canvas e resultado válidos **da mesma captura**, identificados como prévia anterior. A janela sai do estado ocupado e permite corrigir os campos. Gerar/imprimir e ações dependentes do plano ficam bloqueadas até concluir uma nova prévia válida. Não é possível exportar a seleção antiga como se fosse a atual.
- Atualizações válidas mantêm a prévia anterior identificada durante o processamento e só publicam o novo resultado após o render correspondente. Cancelamento, versão da requisição, worker isolado, prazos e retry continuam ativos. Resultados atrasados não substituem a seleção corrente.
- Ao fechar e reabrir após excluir/adicionar/substituir páginas, a janela recaptura os bytes e o total atuais. A janela é modal; não há edição simultânea do workspace por trás dela. Todas usa o novo total. Uma seleção específica que deixou de existir é mantida no campo e recusada, sem remapeamento automático. Nenhuma prévia de outro documento é reapresentada como válida após nova captura.
- **Step & Repeat** mantém arte única, quantidade de repetição, escala, grade e opções próprias. Seus controles e valores não herdam a seleção do N-up. Livretos permanecem em sua janela independente.

Em 4-up, os cenários novos conferiram:

| Páginas reais | Folhas PDF | Peças na última folha | Posições finais sem arte |
|---:|---:|---:|---:|
| 1 | 1 | 1 | 3 |
| 4 | 1 | 4 | 0 |
| 5 | 2 | 1 | 3 |
| 8 | 2 | 4 | 0 |
| 10 | 3 | 2 | 2 |

São vagas na folha imposta, não páginas adicionadas ao documento de origem. O planejador compacto conserva seu comportamento de usar o bloco realmente ocupado na folha parcial.

## Arquivos da correção

Produção:

- Novo `public/js/tools/pdf-imposition-selection.js`: resolução/validação pura com total real e orçamento de 10.000 posições, antes de expandir intervalos.
- `public/js/tools/pdf-imposition-ui.js`: contagem na captura, controles básicos/avançados, validação antecipada, preservação da prévia e publicação do resultado após render.
- `public/js/tools/pdf-imposition-settings.js`: opção Todas/específicas no estado independente do N-up.
- `electron-builder.hold-native-test.cjs`: somente incremento de versão para .25, mesma identidade.

Testes:

- Novo `tests/pdf-imposition-selection.test.mjs`: seleção, limites, ordem, recaptura/estado e PDFs reais gerados nos dois planejadores.
- Novo `tests/pdf-imposition-selection-browser.mjs`: Chrome ou Electron de desenvolvimento, sem instalação/impressão, com perfis isolados e workspace real.
- `tests/pdf-imposition-preview.test.mjs`: dois casos de regressão para seleção inválida sem chamar motor e cancelamento de resultado atrasado; adaptação do import do novo módulo.
- `tests/helpers/pdf-imposition-basic-browser.mjs`: apresentação básica esperada, agora com grade e sem seleção manual no N-up.
- `tests/pdf-imposition-browser.mjs`: roteiro instalado obtém a versão do builder atual, evitando instalar/verificar a .24 por engano. O modo instalado não foi executado nesta etapa.

Não foram alterados `pdf-imposition-engine.js`, planejadores tradicional/compacto, bounds, marcas, resumo de escala, renderizador PDF.js, `PDFWorkspace`, módulos de livreto, `main.js`, motores Windows, DEVMODE, associação ou atualizadores. A versão principal do package.json permanece .6, conforme a separação existente da variante Hold Test.

## Testes executados

- **377 testes Node aprovados**, zero falhas/cancelamentos/skips: 362 anteriores + 13 de seleção/geração + 2 de prévia. Incluem livretos e regressões geométricas .21–.23. Matrizes/BBox dos PDFs efetivamente gerados conferidas; originais sintéticos preservados.
- **Chrome dedicado aprovado:** 50 PDFs/planos, 1/4/5/8/10 páginas, células e compacto, 2/4/6/8-up, texto extraído na ordem esperada, última folha sem repetição, correção de entradas inválidas, canvas/pixels/resultados anteriores idênticos durante o erro, sem estado ocupado permanente ou saída indevida.
- **Electron dedicado aprovado:** a mesma bateria de 50 PDFs/planos, com perfil `.desktop-test/nup25/electron/profile`, sem usar o perfil da instalação .24. Nenhuma exceção não tratada em ambos os ambientes.
- Os dois roteiros dedicados também conferiram 30 mudanças rápidas de distribuição/orientação/grade por documento, seleção acima do total durante atualização, e reabertura após exclusão, união e substituição no PDFWorkspace real. Seleção específica inválida após excluir páginas não foi alterada silenciosamente. Step & Repeat continuou repetindo somente a arte explicitamente escolhida.
- **Roteiro integrado Chrome aprovado**: montagem básica/avançada, escala, excesso/corte, preparação compatível, Step & Repeat, N-up tradicional/compacto e livretos. Montagem: 9 renders/8 cancelamentos; livreto: 5/4, falha injetada/retry, frente/verso/leitura, brancos e reentrada. Prova do livreto: **18 pares / 10.953.600 componentes RGBA idênticos**, sete números de página extraíveis.
- **Electron smoke aprovado**: sandbox/isolamento, PDF.js/worker, giro, abas, exportação/cancelamento/falha, proteção e persistência/origem.
- **41 checks C# aprovados**: DEVMODE/Hold/política de padrão/escala Windows; sem enviar trabalhos.
- Sintaxe dos 9 JS/CJS/MJS novos/alterados e `git diff --check` aprovados.
- A falha de zoom do editor medida na auditoria .24 permanece separada, sem correção. A suíte geral que falhou naquele ponto não foi repetida como se a presente alteração a resolvesse; foram executadas as baterias pertinentes acima.

Comandos na raiz `Corel Script Configurator`:

```powershell
npm test
node tests/pdf-imposition-selection-browser.mjs
$env:NUP_SELECTION_ELECTRON='1'
node tests/pdf-imposition-selection-browser.mjs
Remove-Item Env:NUP_SELECTION_ELECTRON
$env:IMPOSITION_DESKTOP='0'
$env:IMPOSITION_RESULTS='.desktop-test/nup25/chrome-integrated'
node tests/pdf-imposition-browser.mjs
node tests/desktop-smoke.mjs
node scripts/test-print-native.mjs
```

Evidências ignoradas: `.desktop-test/nup25/{node-final,browser-selection,electron-selection,chrome-integrated,electron-smoke,native-checks,build,installer-extract}.log`; PDFs, capturas e resultados em `chrome/`, `electron/` e `chrome-integrated/`. Os roteiros novos usam apenas sintéticos; não contêm PDFs de clientes. Falhas de configuração dos próprios testes durante o desenvolvimento foram corrigidas nos ensaios (regex de extração e seleção explícita de Sem marcas para o compacto), sem modificar a geometria para contorná-las.

## Instalador atualizável preparado

`dist-desktop\hold-native-test\4PrintShops-Hold-Test-1.0.0-experimental.25-Windows-x64-Setup.exe`

Tamanho: **126.420.593 bytes**. SHA256: `81363AB7A98AC02E11849DE8C069C9FD514EB1B6DF02E4C5B66CDFB927A25688`.

```powershell
$env:PRINT_GHOSTSCRIPT='none'
npx --no-install electron-builder --config electron-builder.hold-native-test.cjs --win nsis --x64 --publish never
node scripts/check-desktop-package.mjs 'dist-desktop/hold-native-test/win-unpacked/resources/app.asar'
```

Build e extração/auditoria do próprio Setup aprovados: **455 entradas ASAR**, incluindo o novo módulo de seleção, fontes iguais ao checkout e auxiliares iguais aos manifestos. ASAR extraído: SHA256 `F30BC10B6622268DEB254D694C8C7FF30659BCA30FA07E97453D790A94AFD39C`; igual ao empacotado, assim como o executável. Versão interna `.25`, `name:4printshops-hold-test`, flags Hold Test/Windows conferidas. Auditoria de exclusões aprovada, sem testes/clientes/segredos no pacote. Prova em `.desktop-test/nup25/package-proof.json`. O aviso do 7-Zip sobre dados após o fim do arquivo do contêiner NSIS é acompanhado de extração concluída e conferência independente dos bytes, como na auditoria .24.

Preservados appId `br.com.4printshops.holdtest`, produto/executável/atalhos, chave NSIS, origem `app://4printshops` e perfil `%APPDATA%/4PrintShops Hold Test`, com `deleteAppDataOnUninstall:false`. A leitura do registro encontrou **Hold Test .24 instalada em `%LOCALAPPDATA%\Programs\4PrintShops Hold Test`**, e perfil existente. Nenhum processo dessa instalação foi iniciado ou substituído. O novo instalador foi somente gerado/auditado; upgrade real e preservação valor a valor das preferências nesta máquina aguardam execução manual. O Setup .24 anterior permaneceu com SHA256 `4E62F15ADA01DDFFAC99EE0337FA894CA96ABD6022A501C86092B938A1EC2115`.

O pacote é experimental e sem Authenticode comercial. Ghostscript não foi incorporado; o diagnóstico PS/Hold sem interpretador continua bloqueado. Atualizador automático e distribuição assistida não foram ativados nem publicados.

## Conferência manual e pendências

1. Exportar trabalhos pendentes e fechar a Hold Test antes de aplicar deliberadamente o novo Setup. Conferir que atualiza a instalação .24 existente, sem outra identidade, e comparar preferências usuais.
2. Abrir PDFs de 1/4/5/8/10 páginas. Escolher N-up: total automático, grade no Básico, todas as páginas em sequência e última folha incompleta. Conferir células e compacto.
3. Em Avançado, selecionar páginas específicas; informar número acima do total, zero, negativo, texto, intervalo invertido e campo vazio. A prévia anterior deve permanecer identificada, a mensagem deve orientar e a saída ficar bloqueada. Corrigir e gerar o PDF para medir/reabrir.
4. Fechar a montagem, excluir/adicionar páginas ou abrir outro PDF e reabrir. Conferir novo total; seleção específica que deixou de existir deve ser recusada, sem trocar por outras páginas.
5. Conferir Step & Repeat na arte/quantidade habituais e alternar rapidamente opções. Registrar sequência/console se aparecer outro desaparecimento de prévia.

Não houve nenhuma impressão, física ou virtual. A homologação física de **livretos .24 com os funcionários da gráfica continua pendente**, incluindo dobra, ordem/aninhamento, duplex, alimentação, registro, papel/bandeja, refile, grampo e fidelidade no RIP para A5/A4 e A4/A3. O zoom do editor e eventuais outros caminhos intermitentes de prévia continuam pendentes separadamente. Sem commit/push.
