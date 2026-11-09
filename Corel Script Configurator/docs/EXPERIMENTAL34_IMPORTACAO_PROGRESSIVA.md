# Experimental.34 — importação progressiva para avaliação manual

## Diagnóstico posterior da distribuição — 09/11/2026

A investigação autorizada reproduziu exatamente os1108 ausentes com o processo
NSIS: extração curta íntegra, falha de CopyFiles no destino longo e fallback
Nsis7z incompleto. O pacote .34 contém tudo e o validador recusa corretamente o
runtime instalado. Não houve reparo manual, reinstalação ou alteração do perfil.
A correção isolada .35 e os ensaios em ambiente próprio estão documentados em
[DISTRIBUICAO_CAMINHOS_EXPERIMENTAL35.md](DISTRIBUICAO_CAMINHOS_EXPERIMENTAL35.md).
A entrega integral revelou outra limitação do próprio Office em caminhos longos;
o layout .35 usa `lo` junto ao executável e exige inicialização após os hashes.
No comprimento real da estação,19510 arquivos chegaram íntegros, DOCX/DOC
converteram texto e dimensões A4 e os testes de recusa passaram.472 Node e smoke
Electron aprovados. Setup .35 pronto em laboratório, **não instalado**; SHA256
`FC3199052123BB5A64E6CBD3EF058626024A65310BD97CC01113F74E432AB108`.
Perfil75 arquivos, registro e arquivos principais da .34 permanecem iguais.
A .34 ativa continua incompleta, sem conversão Word; o histórico abaixo permanece
inalterado e sua hipótese inicial é superada pelo diagnóstico reproduzido.

## Instalação na estação da gráfica — 09/11/2026

**Instalação na gráfica .29 → .34 — 09/11/2026: parcial, Word NÃO liberado.** Setup SHA256 B029B4934AF3B7F5DF12DB4BA0728B40FFAD44FA729CD336A006803D948A129E reconferido, aplicativo/auxiliares fechados, sem documento aberto. Backup integral verificado de75 arquivos/8.450.439 bytes do perfil em .desktop-test/install34-graph-FnMZoF/before/profile-backup. Instalador retornou0; registro, FileVersion e IPC da aplicação confirmam1.0.0-experimental.34 no mesmo diretório/perfil/origem Hold Test, com um único registro NSIS. Abertura/fechamento normal e comandos DOC/DOCX aprovados. Preferências antes/depois iguais: painéis312/230, tema light e marcador preservado. Associações PDF/UserChoice iguais por leitura; seis arquivos principais instalados iguais ao Setup. **Validação do Office instalado falhou por arquivo ausente antes de iniciar conversão:**1108 dos19410 arquivos esperados faltam (18302 presentes, mais o manifesto). O Setup/payload auditados contêm todos. Caminhos ausentes medem255–294 caracteres; diretório extenso é hipótese relevante, ainda sem causa exata comprovada. Não é timeout de conversão: proteção recusou runtime incompleto, sem fallback global. Não foram copiados arquivos à mão, relaxadas verificações, alterados caminhos ou feita reinstalação/rollback automático. LibreOffice instalado não homologado; disponibilidade dos comandos não implica capacidade de converter. Investigação/correção do instalador depende de nova autorização. Somente três documentos atualizados, código e alterações anteriores preservados, staging vazio; sem impressão física/virtual, drivers/impressoras, política de scripts, commit ou push. Evidências locais em .desktop-test/install34-graph-FnMZoF/. Registros de geração/instalação em casa abaixo são históricos e não superam esta falha local.

### Resultado verificado

| Item | Resultado |
|---|---|
| Setup autorizado |Hash local B029B4934AF3B7F5DF12DB4BA0728B40FFAD44FA729CD336A006803D948A129E confirmado antes da execução; retorno0 |
| Instalação/identidade |Mesmo diretório e perfil da .29; registro NSIS90ff2ff0-f013-571c-b109-7edb265dacc8 único; FileVersion e IPC1.0.0-experimental.34 |
| Abertura |Tela inicial/PDFs e importação unificada DOC/DOCX disponíveis; nenhuma exceção; fechamento normal0 |
| Preferências |localStorage integralmente igual: painéis312/230, tema light, marcador preservado |
| Associações |Cinco árvores de registro de PDF/capacidades/UserChoice idênticas antes/depois; não alterado o leitor padrão |
| Arquivos principais |Executável, ASAR, ícone PDF, dois auxiliares e PDFium iguais ao payload auditado |
| Word instalado |REPROVADO na conferência de integridade, antes de converter; não declarar DOCX/DOC aprovados nesta instalação |
| Perfil |Backup integral75 arquivos,8.450.439 bytes, cópias verificadas por hash, fora do Git |

### Falha do runtime instalado — sem correção automática

O teste existente contra o diretório realmente instalado falhou com ENOENT em
`resources/libreoffice/help/media/screenshots/filter/ui/pdfuserinterfacepage/PdfUserInterfacePage.png`.
O inventário encontrou18303 arquivos incluindo o manifesto; faltam1108 dos19410
arquivos de conteúdo. O arquivo de exemplo tem caminho absoluto de273 caracteres.
Os ausentes variam de255 a294; os presentes chegam a276. Diretórios dos ausentes
variam de225 a271 caracteres, enquanto os pais existentes chegam a247. Portanto
não se deve concluir uma regra simples de que todo caminho acima de260 falha.
Há indícios de limite de caminhos/diretórios na instalação NSIS neste destino
extenso; a chamada exata/causa ainda não foi investigada ou corrigida.

O Setup extraído antes da instalação possuía todos os19510 arquivos do pacote,
com hashes iguais, e o teste Word de win-unpacked passou. O retorno0 do NSIS não
é prova de cópia integral. A proteção do resolvedor privado corretamente recusou
o runtime incompleto; não tentou Office global nem descartou o manifesto.
Não confundir com o timeout120s da primeira rodada do pacote, registrado abaixo.
Não foi iniciada conversão no teste instalado que falhou. Nenhum arquivo de
cliente foi usado ou alterado.

A .34 permanece instalada e abre, porém **Word não deve ser homologado**. Não
houve reparo manual, cópia dos arquivos faltantes, alteração de política Windows,
redirecionamento do runtime, reinstalação ou retorno à .29. Próxima etapa requer
autorização para investigar/corrigir a instalação preservando identidade e perfil.
PDFs/imagens/impressão não receberam nova homologação física nesta verificação.
Provas: installation.json, before/profile-proof.json, before/after-preferences.json,
before/after-associations.json, installed-proof.json, installed-private-office.log
e missing-runtime.json na pasta local ignorada acima.

## Geração nesta estação da gráfica — 09/11/2026

**Geração local na gráfica — 09/11/2026:** base main/c85efe2903b70bcee11da272328e957cecd52ad5 inicialmente limpa. Código de produção, testes e builders preservados; somente documentação atualizada. Preparador existente obteve MSI e quatro fontes oficiais do LibreOffice26.8.2.2, verificou os SHA256 fixados e extraiu runtime privado sem instalação global. Runtime19410 arquivos/1.581.539.537 bytes; manifesto01f0286083b702da3c085ec7219c22b98db2d86aa8ae3f6c3633f6707537247e. Diferença de29 bytes versus casa restrita ao NOTICE com CRLF; normalizar somente esse arquivo reproduz o manifesto doméstico ae42c73b898efeb80813b024634d91cfd0e64503702b3c67e1588ea4b31d5781. Setup local .34 com retorno0,514.963.606 bytes e SHA256 B029B4934AF3B7F5DF12DB4BA0728B40FFAD44FA729CD336A006803D948A129E. ASAR466 e19510 arquivos do payload (1.974.774.815 bytes) iguais por SHA256 ao pacote.35 testes rápidos, importação progressiva Electron e DOCX/DOC pelos módulos do ASAR/runtime de win-unpacked aprovados,2 páginas cada, busca externa desativada, cancelamento e originais preservados. Primeiro teste Word concorrente ao build excedeu120s; repetição isolada passou sem alterar limite/código, causa inicial não isolada. Mesma Hold Test; recursos Windows/PDFium presentes; PRINT_GHOSTSCRIPT=none mantém diagnóstico experimental PS/Hold bloqueado sem Ghostscript. NotSigned experimental, sem enfraquecer atualizadores. Esta geração não instalou; .29 e cinco arquivos principais permaneceram intactos naquele momento. Sem impressão, mudança de drivers/impressoras, limpeza, commit ou push. Evidências/gerados ignorados em .desktop-test/graph34/ e build/dist-desktop. A instalação autorizada posterior nesta estação tem resultado separado acima; a integridade do pacote não comprova integridade da cópia instalada.

### Artefato e preparação local

- Setup: `dist-desktop/hold-native-test/4PrintShops-Hold-Test-1.0.0-experimental.34-Windows-x64-Setup.exe`,514.963.606 bytes.
- SHA256: `B029B4934AF3B7F5DF12DB4BA0728B40FFAD44FA729CD336A006803D948A129E`.
- ASAR466 entradas: `4d6392090d6723df94365c575bb07bb67c6a8ff6bfce12b390dd0813c022e165`; versão .34, módulos Word/progressivos presentes.
- Node24.18.0, Electron45.1.0, electron-builder26.18.0, electron-updater6.8.11 e compilador C# existentes; não reinstaladas dependências.
- MSI375.934.976 bytes e quatro fontes obtidos de download.documentfoundation.org, com redirecionamento oficial para mirrors e SHA256 do lock conferidos. Fonte/runtime/cache ignorados; licenças, fontes e SOURCE-CODE.json incluídos. Nenhum Office global instalado pelo preparador.

### Validação da geração, sem instalação naquele momento

| Verificação | Resultado |
|---|---|
| Node focado |35 aprovados: runtime, Word, importação unificada/progressiva;468 anteriores não repetidos |
| Electron progressivo |PDF/imagens1/100/350, estados/contagem/bloqueio, falha/cancelamento e edições anteriores aprovados, perfil isolado |
| Word de win-unpacked |DOCX/DOC sintéticos2 páginas cada, executável privado, busca externa desativada, originais intactos, cancelamento e motor ausente recusados |
| Runtime após conversão |19410 arquivos íntegros |
| Setup extraído |NSIS/7z íntegros;19510 arquivos iguais ao pacote por SHA256 |
| Instalação anterior |Na geração, .29 e cinco arquivos principais intactos, sem abrir a instalação |

A primeira execução Word durante o build excedeu120s. O teste isolado passou
com o limite original; causa inicial não isolada, logs preservados. Não houve
correção de produto para obter aprovação. Essa aprovação de win-unpacked não
substitui a validação da instalação, que falhou posteriormente como registrado acima.

Manifesto local01f0286083b702da3c085ec7219c22b98db2d86aa8ae3f6c3633f6707537247e:
normalizar apenas o NOTICE reproduz ae42c73b898efeb80813b024634d91cfd0e64503702b3c67e1588ea4b31d5781
do runtime doméstico. Os hashes de Setup/ASAR são da geração local; não se exige
reprodução binária idêntica à de casa. Sem assinatura comercial; verificações do
atualizador preservadas. PRINT_GHOSTSCRIPT=none mantém diagnóstico PS/Hold
experimental bloqueado sem Ghostscript; Windows/PDFium incluídos, sem testar envio.
Nenhuma limpeza, impressão, alteração de driver, commit ou push. Evidências em
.desktop-test/graph34/. Pendentes seletor nativo/PDF único, lotes Word reais, fontes,
paginação/ICC/RIP, livreto físico, caixas, zoom e OCG/OutputIntents.


## Consolidação local autorizada para continuar na gráfica — 09/11/2026

Usuário autorizou o checkpoint **sem esperar a avaliação manual pendente**.
Isso não constitui homologação da .34. Seletor nativo real, percepção visual dos
estados/miniaturas, conveniência da conferência de PDF único e fechamento pelo X
continuam pendentes, conforme o roteiro abaixo. Os resultados automatizados e
instalados anteriores foram aproveitados, sem nova execução nem atribuição de
aprovação manual ao operador.

Base main/1845fc1, staging inicialmente vazio. Auditados21 caminhos; somente
documentação mudou desde a última conferência instalada. Sintaxe de15 arquivos
JS/MJS/CJS e whitespace conferidos antes do staging/commit local autorizado.
Mensagem: `Consolida importacao progressiva e avaliacao da experimental 34`.

Inventário completo:

- `desktop/app.mjs`, `desktop/import-files.mjs`, `desktop/preload.cjs`.
- `public/css/theme.css`, `public/js/app/desktop-client.js`,
  `public/js/tools/pdf-import-review.js`, `public/js/tools/pdf-merge.js`,
  `public/js/tools/pdfs.js`.
- `electron-builder.hold-native-test.cjs`.
- `tests/desktop-import-files.test.mjs`, `tests/desktop-word-smoke.mjs`,
  `tests/helpers/pdf-combine-modules.mjs`, `tests/helpers/pdf-dom.mjs`,
  `tests/pdf-import-progress.test.mjs`, `tests/pdf-import-progress-browser.mjs`,
  `tests/pdf-integration.test.mjs`.
- `docs/DISTRIBUICAO_DESKTOP.md`, `docs/ESTADO_PROJETO.md`,
  `docs/IMPORTACAO_WORD_LOCAL.md`, `docs/IMPORTACAO_PROGRESSIVA.md`, este relatório.

A ferramenta `npm run cleanup` e seus testes/guia já foram registrados na .33,
commit1845fc1, e permanecem nessa história sem alteração. Nenhuma exclusão ou
modo destrutivo foi acrescentado. Não executar limpeza real como parte da retomada.

Instaladores, binários, builds, node_modules, LibreOffice/cache, perfis, backups,
evidências e arquivos de clientes não integram o staging. Não foram lidos ou
copiados documentos de clientes nesta consolidação. Não houve novo pacote,
instalação, impressão, limpeza ou push, nem alteração de credenciais/configuração
global do Git. Histórico e mudanças existentes preservados.

Para continuar na gráfica, o usuário deve enviar o commit manualmente e verificar
o Git de destino antes de sincronizar, preservando eventuais alterações daquela
máquina. O Setup .34 e os recursos privados de build não viajam pelo Git; o caminho,
hash e os procedimentos de preparação estão neste relatório e na documentação de
distribuição. Não presumir que a instalação da outra máquina foi atualizada.

## Atualização controlada .33 → .34 — 08–09/11/2026

**.34 instalada, sem alteração de funcionalidades.** Usuário autorizou a atualização
condicionada à segurança; não autorizou commit/push ou limpeza. Branch main,
HEAD1845fc1,21 caminhos pendentes e staging vazio conferidos. Esses21 arquivos
foram copiados e verificados por SHA256 antes da instalação. Nenhuma alteração
neles ocorreu durante os testes; esta etapa acrescenta somente estes registros
documentais. As indicações de “não instalado” abaixo são históricas, da preparação.

### Segurança e instalação

- Setup existente com SHA256 `6E9C7AD7CC1DD41939D044FB03BD891C1C8E68E6230BCEB48865F8EC8F9FD3C4`
  confirmado novamente. Nenhum pacote novo gerado.
- Inicialmente não havia processo Hold Test; versão anterior1.0.0-experimental.33.
  Uma sessão de coleta confirmou ausência de documento e registrou preferências.
  Antes de executar o Setup confirmou-se novamente ausência de processos.
- Perfil integral copiado para
  `.desktop-test/install34-1794192823707/before/profile-backup`:81 arquivos,
  8.564.612 bytes, cada cópia conferida por hash. Pasta fora da allowlist de limpeza.
- Setup autorizado executado com /S, retorno0. Mesmo diretório
  `%LOCALAPPDATA%/Programs/4PrintShops Hold Test`, perfil APPDATA, appId
  br.com.4printshops.holdtest e chave NSIS90ff2ff0-f013-571c-b109-7edb265dacc8.
  Nenhuma segunda instalação. Registro e FileVersion confirmam .34.
- Executável, ASAR, ícone, auxiliar native, auxiliar hold-native e PDFium instalados
  iguais por SHA256 ao payload auditado. Auditor do instalado aprovado antes e
  depois dos ensaios:466 entradas, código atual e19.410 arquivos LibreOffice íntegros.

### Preferências e resultados realmente executados

Preferências da aplicação (localStorage) iguais antes/depois, uma chave registrada,
e arquivo Chromium Preferences idêntico ao backup. Caches de execução não são
confundidos com preferências. Nenhum driver, fila, perfil de impressora ou original
de cliente foi alterado.

| Verificação | Resultado desta instalação |
|---|---|
| Binário realmente instalado | Teste CDP com sintéticos PDF/imagens aprovado: conferência/miniaturas progressivas, contagem,1/100/350 páginas, bloqueio/liberação, falha intermediária, cancelar/reabrir, ausência de resultados tardios e recusa/aceite da substituição com edições pendentes. Arquivos fornecidos por eventos de teste do renderer; não é seleção nativa pelo operador. |
| PDF individual no fluxo Importar | Aprovado pelo código do ASAR instalado com escolha de arquivo simulada no adaptador; exige conferência/Combinar antes de abrir, conforme o comportamento da .34. Seletor nativo real permanece para avaliação manual. |
| Lote PDF/JPG/PNG/DOCX/DOC | Código real do ASAR instalado, recursos/LibreOffice da instalação e perfil de teste isolado; apenas diálogo simulado. Sete páginas, ordem/relatórios, miniatura enquanto outro arquivo converte Word, batimento de interface, bloqueio/liberação e cinco originais intactos aprovados. Não declarar esse ensaio como operador usando o seletor nativo. |
| Cancelar durante Word | Aprovado no ensaio ASAR instalado: presença do processo soffice.com privado confirmada antes de clicar Cancelar importação; documento anterior de sete páginas preservado, sem resultado tardio. |
| Edições/documento anterior | Falha de arquivo após PDF válido, recusa de substituição e cancelamentos conservaram o documento e as edições sintéticas anteriores; não houve adoção parcial. |
| Step & Repeat/N-up/Livreto | Regressão curta no executável instalado aprovada: três A6 sintéticos, Step & Repeat, N-up em células/compacto com escala1 e prévias de impressão/leitura de livreto; cancelamento dos diálogos sem modificar a origem. Nenhum botão de impressão acionado. |
| Preservação do checkout |21 arquivos iniciais com hashes iguais antes das adições documentais. Nenhuma funcionalidade, builder, teste permanente ou dependência modificada nesta instalação. |

A suíte468 Node e as regressões mais amplas da preparação foram aproveitadas,
não repetidas nem declaradas como nova execução aqui. Nesta etapa foram executados
os ensaios instalados descritos acima e a auditoria do pacote após o uso.

O roteiro rápido de coleta de preferências conseguiu registrá-las e confirmar
igualdade, mas sua espera de encerramento automático não confirmou saída0 no
prazo, tanto antes quanto depois. Isso é registrado como limitação desse roteiro,
sem atribuir causa ao produto. Confirmou-se ausência de processos entre os ensaios;
a regressão curta posterior terminou normalmente com saída0. Nenhuma proteção de
fechamento foi removida. Incluir fechamento pelo X na avaliação manual.

Evidências ignoradas em `.desktop-test/install34-1794192823707/`:
before/source-proof.json, before/profile-proof.json, installation-proof.json,
preservation-proof.json, before/after-preferences.json, installed-binary/results.json,
installed-results.json e capturas. Word sintético final em
`%TEMP%/4printshops-import31-ieiex7/`. Adaptadores de teste ignorados não integram
o aplicativo nem foram inseridos no ASAR. Nenhuma automação prolongada de diálogo
nativo foi tentada nesta atualização.

### Verificações manuais restantes

Os cinco arquivos em `.desktop-test/install34-1794192823707/synthetics/` são
exclusivamente sintéticos, copiados e verificados, com nomes input.pdf, input.jpg,
input.png, input.docx e input.doc. O lote completo produz sete páginas; o Word
pode ter metadados de contagem/fontes deliberadamente diferentes para exercitar
os avisos existentes. Não representa prova de fidelidade universal.

1. Pelo atalho da .34 instalada, selecionar só input.pdf no seletor nativo e
   conferir a janela adicional antes de Combinar.
2. Selecionar os cinco input.* pelo seletor Windows e observar pessoalmente
   nomes/tipos, estados, primeiras miniaturas e contagem. Confirmar bloqueio de
   ordenação/combinação até o fim e liberação posterior.
3. Girar uma página sintética e tentar outro lote: cancelar durante Word e
   verificar que o anterior não muda. Reabrir e conferir ausência de miniatura tardia.
4. Para avaliar o lote de sete páginas isolado, escolher Criar novo PDF combinado;
   recusar primeiro a substituição para conferir a preservação, depois aceitar
   somente o descarte consciente das alterações sintéticas do teste.
5. Conferir a comodidade do fluxo de PDF único e fechar normalmente pelo X.

Não foi alegada aprovação manual desses passos. Permanecem também Word em Windows
limpo/em massa na gráfica, fidelidade fontes/paginação/ICC/RIP, homologação física
de livretos, zoom/bordas, caixas e OCG/OutputIntents. Sem impressão física/virtual,
limpeza, novo instalador, commit ou push. A .34 continua experimental.

Data:08/11/2026. Branch main, HEAD/checkpoint .33 preservado:
`1845fc1f4cbee8e75793b4875b8ee6c385d17d52`.
Setup .34 gerado e auditado, **não instalado**. Não há homologação física nova.

## Artefato e identidade

- Caminho: `C:\Users\Marcos\Documents\automacaoAdesivosCoreldraw\Corel Script Configurator\dist-desktop\hold-native-test\4PrintShops-Hold-Test-1.0.0-experimental.34-Windows-x64-Setup.exe`.
- Tamanho: **514.965.183 bytes**, aproximadamente515MB.
- SHA256: **`6E9C7AD7CC1DD41939D044FB03BD891C1C8E68E6230BCEB48865F8EC8F9FD3C4`**.
- Versão FileVersion/ASAR: **1.0.0-experimental.34**, Windows x64/NSIS.
- Mesmos appId `br.com.4printshops.holdtest`, produto/executável/atalhos Hold Test,
  diretório de distribuição, origem app://4printshops, chave NSIS e perfil APPDATA
  anteriores. Builder mudou somente .33→.34; deleteAppDataOnUninstall:false e
  runAfterFinish:false preservados. Nenhuma instalação paralela criada.
- Sem assinatura Authenticode comercial: status NotSigned, conforme a variante
  experimental anterior. Publish desativado, build com retorno0.

ASAR SHA256: `E19102A504250DDA3451E75311BB74FB4DD13A1DC2692CA15F85AB85BD632FD1`.
Contém466 entradas, incluindo o novo módulo pdf-import-review.js. A auditoria
comparou arquivos desktop/public com o checkout e verificou exclusões de testes,
documentos, clientes, perfis e dependências de desenvolvimento.

O Setup foi somente extraído estaticamente: NSIS e app-64.7z com “Everything is Ok”.
Os19.510 arquivos/1.974.893.246 bytes do payload foram comparados individualmente
por SHA256 com win-unpacked: **zero divergências**. O próprio ASAR extraído passou
novamente no auditor. LibreOffice privado26.8.2.2:19.410 arquivos, bibliotecas,
fontes, confiança e licenças íntegros, inclusive após os testes Word.

## Produção em relação à .33 instalada

Comparação direta dos ASARs identificou somente estes oito arquivos desktop/public:

- `desktop/app.mjs`: eventos de seleção, etapa e arquivo com token, na importação.
- `desktop/import-files.mjs`: progresso incremental, preservando preflight/limites.
- `desktop/preload.cjs`: observador restrito dos eventos da importação.
- `public/js/app/desktop-client.js`: entrega progressiva, barreira de conclusão,
  fila de validação e descarte de resultados de seleções canceladas.
- `public/js/tools/pdf-import-review.js`: nova área provisória de conferência.
- `public/js/tools/pdf-merge.js`: placeholders, estados/contagem, primeiras miniaturas
  e bloqueio dos controles; adoção somente quando o lote estiver completo.
- `public/js/tools/pdfs.js`: ligação da conferência ao fluxo existente.
- `public/css/theme.css`: apresentação/indicador de preparação e controles bloqueados.

Metadados de pacote/builder passam à .34. A regra de autoscroll aprovada permanece.
Motores de edição, redimensionamento, Step & Repeat, N-up tradicional/compacto,
livreto, exportação/impressão e conversor/runtime privado Word permanecem os
aprovados. Não há mudança de driver ou configuração de impressora.

No Desktop, um PDF único selecionado por Importar também passa pela conferência
e pelo botão Combinar. **Esse comportamento não foi alterado nesta preparação**;
é uma mudança de usabilidade da implementação já existente, para avaliação manual.
O navegador mantém a abertura direta de PDF único. Os detalhes transacionais e
limites estão em [IMPORTACAO_PROGRESSIVA.md](IMPORTACAO_PROGRESSIVA.md).

## Testes desta preparação

| Verificação | Resultado |
|---|---|
| Node | Nova execução468 aprovados, zero falhas/skips/cancelamentos. Avisos PDF.js/Math.sumPrecise conhecidos, sem falha. |
| Chrome/Electron de desenvolvimento | Novas rodadas dedicadas aprovadas: placeholders, primeira miniatura antes do fim, ordem, contagem real, bloqueio/liberação,1/100/350 páginas, falha intermediária, cancelar/reabrir e proteção de edições. |
| Binário real .34 de win-unpacked | Executado sem instalar: mesma matriz sintética PDF/imagens,1/100/350 páginas, falhas, cancelamento/resultados tardios, recusa/aceite de substituição aprovados. Captura binary-smoke/progressivo.png. |
| Smoke do ASAR .34 | Electron45.1.0 carregou o app real do ASAR, com recursos do pacote e perfil isolado: sandbox/PDF.js/miniaturas, edição, abas, giro, exportação/salvamento, cancelamentos/falhas/proteção/fechamento aprovados. |
| Lote misto no ASAR .34 | Diálogo simulado somente no adaptador de teste; código do app/ASAR não alterado. PDF/PNG/JPG/DOCX/DOC em7 páginas, relatórios/ordem, primeira miniatura durante Word, avanço do batimento da interface, bloqueio até concluir, falha no meio, cancelamento e5 originais intactos aprovados. |
| Word privado | Caminho usado pertence a win-unpacked/resources/libreoffice; descoberta externa desativada. Integração completa acima aprovada e integridade verificada após uso. |
| PDFs Word nos motores aprovados | DOCX e DOC convertidos × Step & Repeat/N-up/livreto: seis cenários aprovados, bytes de origem preservados. |
| ASAR/Setup |466 entradas atuais, motor íntegro, extração íntegra e19.510 hashes iguais; versão .34 confirmada. |
| Sintaxe/Git | JS/MJS/CJS pendentes e git diff --check conferidos. Staging vazio; nenhum commit/push. |

Os testes anteriores de montagem Chrome completos foram aproveitados; não se
declara nova rodada integral de montagem nesta preparação. Nenhum trabalho foi
enviado à fila de impressão, física ou virtual.

Duas ocorrências foram restritas à automação: o primeiro adaptador tentou atribuir
resourcesPath com atribuição simples e falhou antes de carregar o app; corrigido
somente no adaptador ignorado de testes. O primeiro ensaio misto exigia cinco
batimentos já na primeira observação de Word; falhou por depender da velocidade
da execução. `tests/desktop-word-smoke.mjs` agora mede avanço do contador a partir
da observação, sem afrouxar os bloqueios/proteções. Repetição completa aprovada.
Nenhuma correção de produção foi necessária durante o empacotamento.

## Preservação e metodologia

Os18 caminhos pendentes iniciais foram copiados e registrados por SHA256 em
`.desktop-test/release34-1794168709616/before/`. A produção foi preservada byte a
byte. Nesta etapa só se acrescentaram versão, documentação e o ajuste de espera
no teste descrito acima.

O binário normal usa o perfil Hold Test definido pelo app; uma sondagem confirmou
que alterar APPDATA no ambiente não o isola. Por isso sua execução ocorreu apenas
após confirmar ausência de processo Hold Test, sem documento aberto, e com backup
integral verificado de79 arquivos/8.557.800 bytes do perfil. Foram usados somente
sintéticos. Os testes extensivos do ASAR usaram perfis isolados via adaptador,
sem inserir modo de teste no pacote ou modificar o ASAR. Não confundir esses
adaptadores com instalação ou teste manual do seletor nativo da .34.

Instalação .33 existente: seis arquivos principais, Setup .33 e Preferences
Chromium conferidos por SHA256 antes/depois, sem diferenças. O backup e as provas
ficam fora da lista de candidatos do cleanup. Nenhuma limpeza foi executada;
extrações, logs e evidências desta auditoria foram preservados e ignorados pelo Git.

Provas: build.log/build-result.json, asar-proof.json, setup-proof.json,
payload-equality.json, binary-smoke/results.json nessa pasta. Testes Word finais
em `%TEMP%/4printshops-import31-nnido4/`; smoke ASAR em
`%TEMP%/4printshops-desktop-BfgcPT/`. São somente sintéticos, fora do commit.

## Avaliação manual antes de decidir instalar

Após decisão explícita de instalação, fechar trabalhos e conferir hash/destino.
Nesta tarefa não foi executado o Setup.

1. Selecionar um PDF pequeno e conferir a etapa extra de conferência/Combinar.
   Avaliar se esse passo é conveniente para o operador; não foi removido na .34.
2. Selecionar PDF, JPG, PNG, DOCX e DOC juntos. A conferência deve abrir após o
   seletor, com nomes/tipos, indicadores e contagem, sem porcentagens inventadas.
3. Observar primeiras miniaturas enquanto outros arquivos ainda convertem;
   organizar/remover/combinar deve permanecer bloqueado e liberar só ao final.
4. Cancelar durante Word e selecionar outro lote: nada atrasado deve reaparecer.
5. Falhar um arquivo no meio; confirmar ausência de importação parcial e correção
   por Cancelar/selecionar novamente. Conferir o documento já editado.
6. Criar novo PDF com alterações pendentes, recusar a substituição e confirmar
   que as edições anteriores permanecem; depois testar aceite consciente.
7. Conferir sequência, páginas/dimensões, Word/fontes e prévias habituais de
   Step & Repeat, N-up100% e livreto. Não presumir homologação física.

Permanecem: validação manual da experiência .34, Word em Windows limpo e em massa
na gráfica, fidelidade universal fontes/paginação/ICC/RIP, livreto físico,
zoom/bordas (~53px), caixas TrimBox/BleedBox, Step & Repeat otimizado e preparação
vetorial OCG/OutputIntents. A UI não acelera o LibreOffice nem renderiza todas as
páginas antecipadamente; PDFs complexos ainda podem consumir CPU no renderizador.
Sem assinatura comercial e sem nova homologação de produção.
