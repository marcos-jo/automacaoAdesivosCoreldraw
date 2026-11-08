# Experimental.33 — autoscroll e automação de ActualText

## Aprovação manual e consolidação local — 08/11/2026

O operador confirmou os dois testes manuais pendentes no aplicativo .33 instalado:

- Importação pelo seletor nativo dos cinco arquivos sintéticos PDF, JPG, PNG,
  DOCX e DOC, acrescentando um arquivo real: todos abriram e funcionaram.
  Nenhum nome, caminho, conteúdo ou dado do arquivo real integra este registro.
- Cancelamento do arraste de miniaturas com Escape funcionando corretamente.

Esses resultados são relato de teste manual do operador, não nova execução
automatizada. Superam as pendências de importação mista e Escape nas entradas
históricas abaixo. A espera interrompida da automação do seletor permanece
registrada como limitação da automação, sem invalidar a aprovação manual.
Não se deduz quantidade total de páginas do lote manual nem compatibilidade
universal a partir dos documentos efetivamente testados.

Checkpoint local autorizado com a mensagem
`Consolida estabilidade experimental 33 e auditoria de armazenamento`.
Inventário auditado de17 caminhos (sem arquivos gerados ou dados de clientes):

- Documentação: `docs/DISTRIBUICAO_DESKTOP.md`, `docs/ESTADO_PROJETO.md`,
  `docs/ACTUALTEXT_AUTOMACAO.md`, `docs/AUTOSCROLL_MINIATURAS.md`,
  `docs/EXPERIMENTAL33_AUTOSCROLL_ACTUALTEXT.md`, `docs/MANUTENCAO_ARMAZENAMENTO.md`.
- Configuração e correção existentes: `electron-builder.hold-native-test.cjs`,
  `package.json`, `public/css/theme.css`.
- Ferramenta somente simulada: `scripts/storage-cleanup.mjs`.
- Testes/automação: `tests/desktop-word-smoke.mjs`,
  `tests/helpers/pdf-actualtext-browser.mjs`,
  `tests/helpers/pdf-thumbnail-drag-browser.mjs`, `tests/pdf-actualtext-browser.mjs`,
  `tests/pdf-thumbnail-drag-browser.mjs`, `tests/pdf-thumbnail-drag.test.mjs`,
  `tests/storage-cleanup.test.mjs`.

Nesta consolidação são repetidas a auditoria, a sintaxe e as conferências de diff
e staging. Aproveitam-se os458 testes Node e os testes instalados já aprovados,
sem repeti-los nem apresentar o relato manual como automação. Nenhuma alteração
nova de produção, proteção ActualText ou motor PDF. Cleanup continua sem operação
destrutiva; seu potencial anterior foi5,170GB e nada foi limpo. Instaladores,
builds, LibreOffice, dependências, perfis, backups e evidências ficam fora do commit.
Sem novo instalador, impressão ou push. Homologação física de livretos, Word em
Windows limpo/em massa, fontes/paginação/ICC/RIP, zoom, caixas, otimização de
Step & Repeat e compatibilidade vetorial OCG/OutputIntents continuam pendentes.

## Atualização instalada e validação do checkpoint — 08/11/2026

Esta entrada supera os registros históricos abaixo de pacote não instalado.
Usuário autorizou instalação e commit local condicionado à aprovação dos testes;
push e limpeza real não autorizados. Base main/d5a5bdf,17 caminhos pendentes e
staging vazio. Inventário/hashes anteriores preservados em
`.desktop-test/checkpoint33-1794162139516/before/source-manifest.json`.

- Setup .33 existente conferido pelo SHA256 acima, sem reconstrução. .32
  inicialmente fechada, sem documentos abertos. Backup de77 arquivos do perfil,
  com comparação SHA256 de cada cópia, em `before/profile-backup` nessa pasta
  de checkpoint, fora de todos os caminhos candidatos do cleanup.
- Atualização silenciosa retornou0. Mesmo diretório
  `%LOCALAPPDATA%/Programs/4PrintShops Hold Test`, perfil APPDATA, appId e chave
  NSIS anteriores. Registro/FileVersion confirmam1.0.0-experimental.33.
- ASAR instalado465 entradas e LibreOffice19.410 arquivos íntegros, avisos e
  licenças presentes. Executável, ASAR, ícone, auxiliar native, auxiliar hold-native
  e PDFium iguais por SHA256 ao payload auditado (installation-proof-final.json).
  Uma primeira conferência tentou dois nomes hipotéticos incorretos; o inventário
  real e a conferência corrigida não encontraram recursos ausentes.
- localStorage antes/depois idêntico (`4printshops.theme=light`); Chromium
  Preferences também idêntico ao backup. Nenhuma configuração de driver alterada.
- Testes CDP no próprio executável instalado:8/100/350 páginas, três tamanhos
  de interface, primeira→fim/última→início, autoscroll superior/inferior,
  serialização/ordem/original, dragCancel, soltar fora, parar, substituir/fechar
  documento e remover quadros/listeners aprovados. Escape físico separado
  continua sem confirmação nesta rodada; não equiparar dragCancel a essa prova.
- No mesmo instalado: Step & Repeat, três A6 em A3/4-up em células/compacto,
  política real e escala efetiva1 para cada peça, três posições ocupadas;
  prévias de impressão/leitura do livreto e cancelamento aprovados.
- Módulos reais do ASAR instalado converteram DOCX/DOC de duas páginas cada
  usando exclusivamente o LibreOffice distribuído, descoberta externa
  desativada, cancelamento/ausência recusados e originais preservados.
- Importação mista pela interface instalada **pendente**: seletor nativo aberto,
  mas automação de foco/seleção não se efetivou. Roteiro atingiu prazo esperando
  lote, antes de selecionar arquivos; não é prova de defeito Word/produção.
  Foi solicitada ajuda manual para selecionar apenas `synthetics/input.*` na
  pasta deste checkpoint. Não declarar sete páginas importadas nesta rodada
  enquanto não houver prova. Rodadas anteriores aprovadas continuam históricas.
- Nova execução458 testes Node aprovada, sem falhas/skips/cancelamentos. Sintaxe
  dos JS/MJS/CJS pendentes e git diff --check aprovados.
- `npm run cleanup` executado exatamente no modo padrão: saída0,5,170GB
  potenciais, fotografia de toda a árvore antes/depois sem alterações. Sem
  --apply/exclusão/movimento implementados. Builds, Setups, backups, runtime,
  código e perfis importantes permanecem protegidos. Ferramenta de desenvolvimento
  acrescentada ao repositório; o Setup já gerado não foi reconstruído para isso.

Evidências ignoradas: results.json, installed-results.json, before/after-preferences,
installation-proof-final.json, cleanup-proof.json/log e capturas nessa pasta.
O primeiro roteiro de observação .32 registrou preferências e fechou normalmente,
mas a espera CDP de Browser.close não recebeu resposta; controlador posterior
usa espera limitada. Isso não foi corrigido em produção.

Inventário permanece17 caminhos, com atualizações somente documentais nesta
avaliação. Commit/staging aguardam as validações restantes; nenhum push, impressão
ou limpeza. A .33 não é homologada fisicamente. Preservadas as pendências de
Word em Windows limpo e em massa, fontes/paginação/ICC/RIP, livreto físico,
zoom/bordas (~53px), caixas e preparação vetorial OCG/OutputIntents.

Data: 08/11/2026. Base `d5a5bdf27836e37074c164a68cdd861e54b84342/main`, com nove caminhos pendentes previamente aprovados nos testes. Todos foram preservados em `.desktop-test/release33/before/` com SHA256 antes desta preparação. Sem novas funcionalidades, instalação, impressão, commit ou push.

## Diferenças em relação à .32

**Produção:** somente a regra em `public/css/theme.css` que permite ao arraste atravessar o aviso sobreposto `#pdf-status` enquanto uma miniatura está sendo arrastada. O aviso permanece visível, o painel recebe o gesto na borda inferior e a interação normal retorna ao terminar. Não muda dimensões, zoom, escala física, velocidade ou cálculos de inserção. Diagnóstico em [AUTOSCROLL_MINIATURAS.md](AUTOSCROLL_MINIATURAS.md).

**Automação ActualText:** o helper distingue preservação da estrutura (caminho existente desde a .29) de descarte com prova de redundância e confirmação nativa. Oito cenários, incluindo recusas, cancelamento e bloqueios; runner isolado disponível. Nenhuma mudança da proteção ou do código de produção ActualText. Diagnóstico em [ACTUALTEXT_AUTOMACAO.md](ACTUALTEXT_AUTOMACAO.md).

**Empacotamento:** `electron-builder.hold-native-test.cjs` passa de `1.0.0-experimental.32` para `1.0.0-experimental.33`. Mesmo appId `br.com.4printshops.holdtest`, produto/executável/atalhos Hold Test, saída `dist-desktop/hold-native-test`, chave NSIS derivada da mesma identidade, origem `app://4printshops` e perfil `%APPDATA%/4PrintShops Hold Test`. `deleteAppDataOnUninstall:false` e `runAfterFinish:false` preservados. O package.json principal continua .6 por desenho da variante; o pacote Hold Test declara .33 em extraMetadata.

**Ajuste adicional somente de teste:** a primeira regressão Word concluiu os fluxos funcionais, mas falhou no pedido final de fechar porque `wait` tratava a saída do Electron como erro antes de observar a confirmação do comando. `tests/desktop-word-smoke.mjs` agora aceita término com código0 apenas quando o próprio teste pediu fechamento e também exige esse código ao concluir. Comandos normais, saídas inesperadas e códigos diferentes de0 continuam reprovando. A repetição completa passou. Não houve mudança do conversor Word ou do aplicativo para contornar a falha.

## Preservação e identidade instalada

A .32 foi identificada no executável/ASAR instalado em `%LOCALAPPDATA%/Programs/4PrintShops Hold Test`. Seis arquivos principais foram registrados antes do build para conferência posterior. Esta etapa não abre ou substitui essa instalação, não cria uma instalação paralela e usa somente perfis de teste isolados.

A comparação entre o ASAR .33 e o ASAR da .32 instalada confirmou **apenas `public/css/theme.css` diferente nos arquivos de produção desktop/public**. Motores de edição, redimensionamento, Step & Repeat, N-up em células/compacto, livreto, preparação raster, impressão Windows e Word/LibreOffice permanecem idênticos. O LibreOffice privado26.8.2.2 e seus19.410 arquivos, confiança, fontes/licenças e referências aos fontes continuam os da .32.

Os testes são mantidos no repositório, excluídos do instalador pela configuração existente. Não foram acrescentados clientes, perfis, evidências ou dependências de desenvolvimento ao pacote. Runtime/cache, compilados e Setup continuam ignorados no Git.

## Verificações desta preparação

| Verificação | Resultado |
|---|---|
| Node | Nova execução441/441 aprovados, zero falhas/skips/cancelamentos. |
| Electron smoke | Aprovado: sandbox/PDF.js/miniaturas, giro/abas, exportação/cancelamento/falha, proteção/persistência e encerramento. |
| Miniaturas Electron curta | Aprovada com `THUMB_DRAG_LIFECYCLE_ONLY=1`: autoscroll, parada/cancelamento, soltar fora, alternância/substituição/fechamento e limpeza. A matriz completa8/100/350 Chrome/Electron anterior foi aproveitada sem refazê-la. |
| Montagem integrada Chrome | Aprovada: Step & Repeat, N-up tradicional/compacto, escalas, excesso, compatibilidade, retry/cancelamentos, salvamento e livretos. Os callbacks de encaminhamento são de teste, sem fila física. |
| Fidelidade do livreto |18 pares,10.953.600 componentes RGBA, diferença máxima0; sete textos/páginas conferidos. |
| Word Electron | Repetição final aprovada: PDF+PNG+JPG+DOCX+DOC em7 páginas, cancelamentos/proteções/exportação, cinco originais intactos e saída normal0. A falha inicial de encerramento está registrada acima. |
| Word pelos módulos reais do ASAR .33 | Aprovado: DOCX/DOC produziram duas páginas cada, descoberta externa desativada, originais preservados, cancelamento/ausência recusados e19.410 arquivos íntegros após uso. |
| PDFs convertidos nos motores anteriores | Seis cenários aprovados: dois PDFs Word × Step & Repeat/N-up/livreto, sem modificar os bytes de origem. |
| ASAR de diretório | Aprovado:465 entradas, fontes atuais, auxiliares/PDFium e19.410 arquivos privados íntegros, licenças presentes, exclusões conferidas. Versão .33 e nome `4printshops-hold-test`. |
| Setup extraído | Aprovado: extração estática NSIS e7z sem executar o instalador,19510 arquivos no payload, auditor de pacote aprovado. ASAR/executável/ícone/auxiliares/PDFium idênticos aos do pacote de diretório. |

A suíte geral Chrome e os oito cenários ActualText isolados, aprovados na etapa imediatamente anterior, foram aproveitados; não são declarados como nova execução nesta preparação. Não houve modificação desses scripts durante o build. Evidências anteriores permanecem em `.desktop-test/actualtext-audit/` e `.desktop-test/autoscroll/`.

Comandos executados na raiz da aplicação, sem impressão:

```powershell
npm test
node tests/desktop-smoke.mjs
$env:THUMB_DRAG_ELECTRON='1'
$env:THUMB_DRAG_LIFECYCLE_ONLY='1'
$env:THUMB_DRAG_RESULTS='.desktop-test/release33/thumbnails-electron'
node tests/pdf-thumbnail-drag-browser.mjs
Remove-Item Env:THUMB_DRAG_ELECTRON
Remove-Item Env:THUMB_DRAG_LIFECYCLE_ONLY
$env:IMPOSITION_DESKTOP='0'
$env:IMPOSITION_RESULTS='.desktop-test/release33/imposition-chrome'
node tests/pdf-imposition-browser.mjs
node tests/desktop-word-smoke.mjs
$env:PRINT_GHOSTSCRIPT='none'
npx --no-install electron-builder --config electron-builder.hold-native-test.cjs --win nsis --x64 --publish never
node scripts/check-desktop-package.mjs dist-desktop/hold-native-test/win-unpacked/resources/app.asar
node scripts/test-private-office.mjs dist-desktop/hold-native-test/win-unpacked
node scripts/check-desktop-package.mjs .desktop-test/release33/payload/resources/app.asar
```

Os builders compilam recursos nativos sem enviar trabalhos. Ghostscript não foi incorporado; as proteções existentes do diagnóstico PS/Hold permanecem.

## Artefato

Setup: `dist-desktop/hold-native-test/4PrintShops-Hold-Test-1.0.0-experimental.33-Windows-x64-Setup.exe`.

Tamanho: **514.962.560 bytes** (aproximadamente515MB). SHA256: **`3C3E7B43A43E6B7C99B14E0A6329D000DF2CE30200DA51782739C71451BA2759`**.

Build NSIS Windows x64 concluído com retorno0, publicação desativada. Setup **sem Authenticode comercial** (`NotSigned`), conforme a configuração experimental. Não foi executado. A extração estática do contêiner NSIS e do `app-64.7z` terminou com integridade de descompressão aprovada. O próprio conteúdo extraído passou no auditor:465 entradas ASAR e19.410 arquivos privados.

ASAR extraído igual ao de diretório: SHA256 `37C3555C910B812FA8973DC202BAACDD6BACA620FD81A2579A60D02B548413DD`. Os seis arquivos principais foram comparados por SHA256 entre pacote e payload; FileVersion do executável e metadata ASAR confirmam **1.0.0-experimental.33**. Os seis arquivos da instalação .32 anterior continuam intactos, e o Setup .32 preserva seu SHA256 original. Provas em `.desktop-test/release33/package-proof.json` e `asar-proof.json`; logs em `build.log`, `asar-audit.log`, `setup-audit.log`, `extract-setup.log`, `extract-payload.log`, `packaged-word.log` e nos logs dos testes. Nada disso é versionado.

Não houve instalação/upgrade .32→.33, teste físico, publicação ou oferta de atualização. Identidade e configuração de perfil foram mantidas; isso não substitui a conferência manual de upgrade e preferências após uma instalação autorizada.

## Git e inventário final

Branch `main`, HEAD `d5a5bdf27836e37074c164a68cdd861e54b84342`, staging vazio. Treze caminhos pendentes: os nove anteriores preservados, builder .33, ajuste do teste de encerramento Word, este guia e atualização da distribuição (ESTADO já fazia parte dos nove). Apenas entradas novas foram acrescentadas ao estado/distribuição; registros históricos mantidos.

- `docs/ESTADO_PROJETO.md`, `docs/DISTRIBUICAO_DESKTOP.md`, `docs/AUTOSCROLL_MINIATURAS.md`, `docs/ACTUALTEXT_AUTOMACAO.md`, `docs/EXPERIMENTAL33_AUTOSCROLL_ACTUALTEXT.md`.
- `electron-builder.hold-native-test.cjs`, `public/css/theme.css`.
- `tests/desktop-word-smoke.mjs`, `tests/helpers/pdf-actualtext-browser.mjs`, `tests/helpers/pdf-thumbnail-drag-browser.mjs`, `tests/pdf-actualtext-browser.mjs`, `tests/pdf-thumbnail-drag-browser.mjs`, `tests/pdf-thumbnail-drag.test.mjs`.

Sintaxe dos sete JS/MJS/CJS pendentes e `git diff --check` conferidos. O conteúdo anterior dos nove arquivos foi preservado (o estado recebe somente a entrada desta preparação). Nenhum instalador, runtime/cache, PDF sintético/cliente, perfil, evidência ou binário gerado foi acrescentado ao Git. Sem commit/push.

## Roteiro manual e pendências

Após decisão explícita do usuário de instalar: exportar trabalhos pendentes, fechar o aplicativo, conferir hash/destino e confirmar atualização da mesma Hold Test. Comparar versão .33 e preferências usuais; upgrade .32→.33 ainda não executado por esta preparação.

1. Usar PDFs de8/100/350 páginas. Com o status visível, arrastar primeira→fim e última→início; conferir indicador, ordem e total. Parar no centro, cancelar com Escape, soltar fora e repetir alternando/fechando a aba.
2. Conferir redimensionamento com ActualText preservado; quando houver oferta de descarte, recusar e depois confirmar conscientemente, cancelar a prévia e verificar novo consentimento. Original e edições anteriores devem permanecer protegidos.
3. Conferir Step & Repeat habitual e N-up em células/compacto a100%, sem ampliação implícita ou repetição da última página. Abrir prévias do livreto em impressão/leitura; não presumir homologação física.
4. Importar seleção mista PDF/imagens/Word, conferir sequência/fontes/paginação e cancelamento sem perder o documento anterior. Usar cópias ou sintéticos e exportação deliberada, sem impressão física.

Permanecem pendentes Word em Windows limpo sem Office e validação em massa na gráfica; fontes/paginação/ICC/RIP universais não homologados. Caixas de corte/sangria, zoom nas bordas, Step & Repeat otimizado, preparação vetorial OCG/OutputIntents e homologação física de livretos continuam em etapas próprias. A .33 será para avaliação; nenhuma homologação física ou comercial nova é inferida dos testes automatizados.
