# 4PrintShops Desktop — distribuição experimental Windows

**Auditoria conjunta de 03/11/2026 — checkpoints registrados em `c539206` e `77e3ca8`:** atualizações assistidas e ícones foram organizados para dois checkpoints consistentes. O usuário aprovou visualmente os ícones. Os **202 testes Node** do conjunto atual passaram; uma cópia correspondente somente ao primeiro commit passou nos **199 testes**. Inventário e divisão exata em [AUDITORIA_ASSISTIDAS_ICONES.md](AUDITORIA_ASSISTIDAS_ICONES.md); recursos/configuração em [ICONES_DESKTOP.md](ICONES_DESKTOP.md). Abrir com/segunda abertura no **Icon Test** continuam pendentes de validação funcional, sem retirar a aprovação visual recebida. O novo pacote Assisted Lab com ícones foi verificado por publicação/download criptográficos locais, sem instalar/executar o Setup. Não equivale a homologar atualização automática ou implantação na gráfica.

Base instalada/aprovada: `1.0.0-experimental.1` (commit `e7b0ae3`); checkpoint do atualizador `90cc0ce` e abertura Windows `c450fed`. Código atual: `1.0.0-experimental.6`, Windows x64, com atualização **assistida** autenticada em laboratório isolado `.5`/`.6`. Os laboratórios anteriores `.2`/`.3` e PDF Test `.4` foram preservados. Não é publicação comercial. Nada foi enviado a releases, servidor da gráfica ou serviços de hospedagem.

**Etapa atual — atualização assistida:** publicação Ed25519, raiz pública incorporada e download SHA512 verificado, separado do electron-updater. A base Assisted Lab `.5` foi realmente instalada e detectou/adiou/baixou/preparou a `.6`; falhas criptográficas/downgrade/Setup adulterado foram recusados e PDF alterado em aba oculta impediu o preparo. Confirmação somente localiza o instalador, sem executá-lo. A `.6` oferecida **não foi executada**. Caddy portátil serviu arquivos reais em loopback, sem Node no PATH, serviço, firewall ou certificado. **199 testes Node**, Chrome curto e Electron aprovados. Nenhuma base principal nova foi instalada; piloto em dois PCs/servidor real e instalação/reinício automáticos continuam pendentes. Guia completo, comandos, hashes, guarda de chaves, configuração da origem e roteiro de implantação em [ATUALIZACOES_ASSISTIDAS.md](ATUALIZACOES_ASSISTIDAS.md).

**Checkpoint anterior de desenvolvimento (03/11/2026): instalação e reinício automáticos ainda NÃO homologados.** O usuário decidiu adiar a configuração da VM e o teste com assinatura confiável para uma etapa futura. A lista de arquivos e os testes daquele checkpoint estão em [AUDITORIA_ATUALIZADOR.md](AUDITORIA_ATUALIZADOR.md). A associação de PDFs, então planejada como próximo passo, foi implementada na etapa posterior descrita abaixo.

**Etapa posterior ao checkpoint:** associação e fila de abertura foram implementadas na experimental.4. O include NSIS registra capacidade/Open With sem alterar default/UserChoice ou associações de outros leitores; desinstalação remove somente suas entradas. Instalação real, abertura fria e segunda instância pelo Shell, múltiplos PDFs, proteção e exportação passaram no PDF Test. **O usuário concluiu pessoalmente os quatro testes manuais, informou que todos passaram e aprovou a funcionalidade para checkpoint.** A interrupção anterior da automação por Esc permanece somente como histórico; não se confunde com essa aprovação posterior. Teste SMB real continua pendente, incluindo correção/validação futura dos nomes de servidor com pontos (FQDN/IPv4), rejeitados atualmente. Relatório, inventário final dos 21 candidatos, hash, limites e roteiro: [ABRIR_PDFS_WINDOWS.md](ABRIR_PDFS_WINDOWS.md). O atualizador permanece sem homologação de instalação/reinício automáticos.

Instalador da etapa de abertura Windows: `dist-desktop/pdf-open-test/4PrintShops-1.0.0-experimental.4-Windows-x64-Setup.exe`, produto/perfil/appId isolados; construído com `electron-builder.open-test.cjs`. O instalador comum pode ser gerado com `npm run desktop:dist`, mas não foi gerado nesta rodada. Os laboratórios de atualização não herdam registro de associação. A produção preserva origem, preferências, proteção e motores atuais. Para incorporar a confiança assistida na identidade principal, usar a configuração e a nova chave principal descritas no guia assistido.

## Protótipo de configuração de impressão — sem envio

Autorizado após os dois checkpoints, em base limpa `77e3ca8`. Menu exclusivo do Desktop: **PDFs → Arquivo → Configuração de impressão (protótipo)**. Auxiliar Win32 externo, DEVMODE completo somente em memória, propriedades nativas e simulação A4/A3 a 100%/120%. Não contém APIs de envio nem alteração de preferências. Compile em desenvolvimento com `node scripts/build-print-helper.mjs`; o builder compila automaticamente antes do pacote. Requer compilador/runtime .NET Framework disponível no Windows x64, sem instalação automática de dependências.

Laboratório **4PrintShops Print Config Test experimental.8**, `electron-builder.print-config-test.cjs`, em `dist-desktop/print-config-test/`, com produto/appId/perfil próprios. Não substitui base principal, Icon Test ou pacotes dos atualizadores. A versão da raiz permanece `.6`. Fontes, limitações, verificações reais e roteiro sem imprimir em [IMPRESSAO_WINDOWS.md](IMPRESSAO_WINDOWS.md). Manifesto de integridade do auxiliar vai no ASAR e EXE em `resources/native/`; ambos são gerados/ignorados no Git. As verificações Authenticode do updater permanecem inalteradas e instalação/reinício automáticos continuam não homologados.

## Hold Test experimental.9

O protótipo anterior foi registrado em `5d07e81`. O novo pacote isolado **4PrintShops Hold Test** usa `electron-builder.hold-test.cjs`, appId `br.com.4printshops.holdtest`, perfil `%APPDATA%/4PrintShops Hold Test` e versão `1.0.0-experimental.9`. A raiz permanece `.6`. A instalação principal e os laboratórios anteriores foram preservados.

Instalador gerado e realmente instalado em teste local em `dist-desktop/hold-test/4PrintShops-Hold-Test-1.0.0-experimental.9-Windows-x64-Setup.exe`, SHA256 `994fa033d32f9f4abdf208698bf9db96d0a477293dbf93c5741a79930f2cc2e0`. Verificação de Hold utiliza o DEVMODE completo da sessão. **Enviar trabalho retido permanece bloqueado**, pois a retenção na saída PS da Konica não foi comprovada. A geração com PDFium foi executada exclusivamente na Microsoft Print to PDF. A saída sintética conservou vetores e converteu textos em contornos. O auxiliar de configuração continua sem APIs de envio.

PDFium e seu pacote fixado são preparados por `scripts/prepare-pdfium.ps1` e `build/pdfium-lock.json`. O beforePack desse laboratório compila o auxiliar virtual e gera o sintético sem criar trabalhos. O pacote inclui `resources/print-lab/`, com DLL, executável, manifesto, sintético e licenças. A configuração principal não inclui esse diretório. Todos os gerados são ignorados. O aplicativo instalado dispensa Node e PDFium externos.

A cópia atual pode ser aberta em leitor externo pelo menu Arquivo do módulo PDFs, com escolha explícita de Edge ou outro leitor local. O workspace continua pendente de exportação. Temporários ficam em `%TEMP%/4PrintShops-external-copies/` até remoção posterior pelo operador, depois de fechar o leitor.

218 testes Node, 23 verificações C#, Chrome curto, instalação NSIS, propriedades reais, cancelamento, falha recuperável do auxiliar, cópia editada no Edge e exportação passaram. Relatórios locais em `.desktop-test/hold-lab/`. A prova de retenção, A3 integral, bandejas/acabamentos e fidelidade de fontes/cores ainda exigem homologação. Authenticode, atualizadores e associação padrão foram preservados. Instalação/reinício automáticos permanecem não homologados. Procedimento completo e limites em [IMPRESSAO_HOLD.md](IMPRESSAO_HOLD.md).

## Hold Test experimental.10 — atualização da identidade .9

Conforme autorização do usuário, `electron-builder.hold-native-test.cjs` gera **4PrintShops Hold Test experimental.10** com o mesmo appId `br.com.4printshops.holdtest`, nome, executável, atalhos e perfil da .9. O novo instalador atualiza a instalação existente e conserva a instalação principal separada. Não criar uma identidade paralela para esta entrega.

Pacote em `dist-desktop/hold-native-test/4PrintShops-Hold-Test-1.0.0-experimental.10-Windows-x64-Setup.exe`. Inclui a variante nativa experimental em `resources/hold-native/` e seu manifesto no ASAR. Ghostscript não acompanha o Setup. O laboratório utiliza a instalação local 10.01.2 com hash fixado e bloqueia o preparo se estiver ausente ou diferente. O fluxo comum mantém o auxiliar anterior sem APIs de envio.

A captura requer a fila local autorizada **4PrintShops Konica PS Capture Lab**, com o driver de referência e exclusivamente FILE:, sem compartilhamento. O programa consulta o gerenciamento da impressora padrão e bloqueia capturas quando estiver automático ou inconclusivo. Após o efeito automático observado nos primeiros testes, o usuário autorizou restaurar a Konica e desativar essa gestão somente no HKCU atual. Nenhum driver, porta ou preferência global da Konica foi alterado. A fila FILE: foi preservada.

O ensaio é restrito ao sintético A4 do workspace, A3 retrato, 120%, uma cópia, zoom do driver 100% e Hold. O PS gerado pelo driver é interpretado, congelado e encaminhado como RAW somente após a confirmação nativa do operador. Cancelar é a opção padrão. A retenção real ainda precisa de homologação pessoal na IC-604. Roteiro, evidências, dependências e limitações em [IMPRESSAO_HOLD_EXPERIMENTAL.md](IMPRESSAO_HOLD_EXPERIMENTAL.md). Atualizadores, Authenticode, servidor/piloto assistido e SMB/UNC conservam suas pendências anteriores.

## Hold Test experimental.11

Atualiza a Hold Test .10 com o mesmo appId, executável, registro NSIS, diretório e perfil. Instalador em `dist-desktop/hold-native-test/4PrintShops-Hold-Test-1.0.0-experimental.11-Windows-x64-Setup.exe`. A instalação principal permanece separada. O resultado físico do usuário confirmou Hold na .10, sem impressão automática na IC-604.

SHA256 do Setup final `38bdde63b22f1eb6ee2aee0eb77c53a6d464fc6a50627438c91c37e1f1457bb3`. Upgrade .10 → .11 e revalidação do pacote final aprovados na instalação de laboratório existente.

A .11 remove a escala fixa, permite 100%, 120% ou 1% a 400% personalizados, separa papel lógico/folha física e mantém zoom do driver 100%. O corte passa por confirmação própria, após validação do PS. Nova operação cria outra sessão e exige Hold/captura/confirmação novamente. Não existe reenvio automático. PS e RAW permanecem no mesmo caminho validado.

Foram executadas capturas FILE: e validação instalada, cancelando todas as confirmações finais. O pacote preservou a preferência anterior e a impressora padrão Konica. Limites atuais, campos recuperados do PS e roteiro físico estão em [IMPRESSAO_HOLD_EXPERIMENTAL.md](IMPRESSAO_HOLD_EXPERIMENTAL.md). A apresentação de todos os campos no painel da IC-604 continua pendente. Não foram alterados drivers, filas, portas, certificados ou preferências Windows nesta etapa. Nenhum commit/push.

## Hold Test experimental.12 — impressão Windows geral

A .11 foi aprovada fisicamente pelo usuário na C6100/IC-604, incluindo escala, envios consecutivos, Hold e propriedades configuradas. Os 20 arquivos pendentes foram copiados com hashes antes da .12, sem commit.

`electron-builder.hold-native-test.cjs` gera a .12 com identidade, diretório e perfil Hold Test preservados. Pacote `dist-desktop/hold-native-test/4PrintShops-Hold-Test-1.0.0-experimental.12-Windows-x64-Setup.exe`, SHA256 `2890e38799f458e425cfc384978777b6ffb4dad84599d455a8170c32c0b85e53`. O teste instalado aprovou upgrade .11 → .12, preferência anterior e impressão virtual com alterações pendentes preservadas.

Menu **Arquivo → Impressão**, impressão normal do workspace via PDFium/HDC/driver escolhido, sem requisito Hold e sem confirmação experimental. A fila Epson real `WF-C21000 Series(Rede)` utiliza o driver padrão `EPSON WF-C21000 Series` V3, pacote 3.60.00.00. O driver administra sua linguagem de saída, cópias e recursos privados. Não se envia PS Konica à Epson. O caminho especializado .11 permanece no diagnóstico opcional Konica / Hold. O fluxo normal dispensa Ghostscript.

Testes Node/C#, Chrome, Electron e pacote instalado foram executados. Envios automatizados ocorreram exclusivamente na Microsoft Print to PDF. O usuário posteriormente informou testes práticos satisfatórios da .12 em outras impressoras da gráfica. A matriz detalhada de opções continua em ampliação. Assinatura comercial, instalação automática do atualizador e distribuição assistida na rede continuam com as pendências anteriores. Detalhes, limitações e roteiro em [IMPRESSAO_WINDOWS_GERAL.md](IMPRESSAO_WINDOWS_GERAL.md). Clone, ambiente doméstico, geradores, testes virtuais e build reproduzível em [DESENVOLVIMENTO_WINDOWS.md](DESENVOLVIMENTO_WINDOWS.md).

## Hold Test experimental.13 — reorganização visual

`electron-builder.hold-native-test.cjs` gera a .13 mantendo appId `br.com.4printshops.holdtest`, produto/executável/atalho Hold Test, perfil e registro NSIS anteriores. Atualização .12 → .13 realizada no laboratório e pacote final reaplicado no mesmo diretório, com preferências preservadas. A identidade principal permanece separada. A versão da raiz continua `.6`.

Setup final em `dist-desktop/hold-native-test/4PrintShops-Hold-Test-1.0.0-experimental.13-Windows-x64-Setup.exe`, SHA256 `ada07a7904c258aca4daa0b8a112e2af51059d89e094dfda84b3e05d5a5c0178`. O pacote instalado foi auditado com 433 entradas ASAR. Os dois novos módulos de apresentação constam no builder e na lista exata de recursos permitidos pelo protocolo Electron. Motores nativos, PDFium, DEVMODE, HDC, PS RAW e atualizadores preservados.

246 testes Node, 41 checks C#, Chrome completo, Electron smoke, propriedades Konica/Epson e impressão exclusivamente Microsoft Print to PDF aprovados. Cancelar/reabrir, proteção de pendências e perfil anterior aprovados. Layout em 1366 × 768 e 1920 × 1080 aprovado. A avaliação de 125% foi emulada no Electron, com conferência nativa do Windows ainda pendente. Nenhum envio físico nesta etapa. Instalação automática via atualizador continua não homologada. Referências, comparação visual, detalhes e roteiro em [UX_INTERFACE_EXPERIMENTAL13.md](UX_INTERFACE_EXPERIMENTAL13.md).

## Hold Test experimental.14 — PDFs e imagens

O builder Hold Test passa a .14 com appId, diretório, perfil e registro NSIS preservados. Setup `dist-desktop/hold-native-test/4PrintShops-Hold-Test-1.0.0-experimental.14-Windows-x64-Setup.exe`, SHA256 `b8524db89eaf8acf55f7773dd733d0e1447ca9647b2c797845d83b5455c95e4c`. Os dois módulos de combinação pertencem ao grupo `public/js/tools/*.js` já permitido no pacote/protocolo. A seleção nativa ampliada reutiliza o IPC de abertura, com validação e limite de 512 MiB. Não foram adicionados leitores padrão de imagens nem o menu de contexto do Explorer.

258 testes Node, 41 checks C#, Chrome completo e Electron smoke aprovados. Upgrade instalado .13 → .14 e pacote final revalidado no mesmo diretório/perfil. A seleção nativa, lote misto, A4, pendências/exportação e trabalho virtual único de oito páginas passaram. Padrão Konica e preferência anterior preservados. 435 entradas ASAR auditadas. Fontes nativas e motores de impressão preservados. Manual, limites de cor/perfis, fixtures, resultados do fluxo instalado, capturas e roteiro em [COMBINAR_PDFS_IMAGENS.md](COMBINAR_PDFS_IMAGENS.md). Instalação automática do atualizador, servidor assistido, piloto de rede e SMB/UNC continuam pendentes.

## Executar em desenvolvimento

Na raiz do projeto, com Node.js/npm instalados (ambiente validado: Node 24):

```powershell
npm ci
npm run desktop
```

Se a política de scripts do npm impedir o download do runtime Electron, executar explicitamente `node node_modules/electron/install.js`. Não é necessário aprovar scripts de pacotes de outros alvos, como Squirrel. O launcher de desenvolvimento retira do processo filho `ELECTRON_RUN_AS_NODE`, variável que pode ser herdada de um IDE Electron.

A versão de navegador continua disponível com `node app.js` ou `npm start`, em `http://localhost:3000`. O servidor escuta somente `127.0.0.1`. O aplicativo instalado não exige Node.js, npm, VS Code ou servidor externo.

## Arquitetura e dados

| Execução | Entrada e origem | Salvamento |
| --- | --- | --- |
| Navegador | `app.js` → HTTP local, `http://localhost:3000` | Mesmo download Blob do fluxo existente. |
| Desktop | `desktop/main.mjs` → `app://4printshops/index.html` | Diálogo nativo seguido de gravação temporária e renomeação no destino escolhido. |

O Desktop serve os assets empacotados pelo protocolo próprio, sem iniciar servidor HTTP, reservar porta ou carregar o backend CorelDRAW. A origem permanece fixa entre reinícios e versões. PDF.js, worker, pdf-lib, HTML/CSS e os módulos PDF/QR são os mesmos da versão web.

- `desktop/app.mjs`: instância única, janela, protocolo, restrições e encerramento. Uma segunda abertura restaura/focaliza a janela existente.
- `desktop/security.mjs`: validação de origem/frame/IPC, caminhos permitidos e dados de exportação.
- `desktop/preload.cjs`: somente abrir PDFs escolhidos pelo operador, salvar PDF, informar estado do documento e solicitar a verificação opcional de atualizações. Não expõe `require`, caminhos arbitrários ou IPC genérico.
- `desktop/files.mjs`: recebe os caminhos exclusivamente dos diálogos do processo principal. A importação devolve nome/bytes, não o caminho; a validação do conteúdo continua nos parsers existentes. O salvamento valida PDF e nome simples; falhas não truncam o destino existente. Exportações ou seleções de importação acima de **512 MiB** são recusadas nesta ponte experimental, sem rasterização para reduzir o arquivo.
- `public/js/app/save-file.js`: escolhe a ponte nativa quando disponível; no navegador mantém Blob/download. Cancelamento/falha preserva pendências; somente a gravação nativa concluída marca a revisão capturada como exportada. `desktop-client.js` integra o seletor nativo aos mesmos inputs/eventos de importação e união, com listener removido ao desmontar PDFs.

Preferências, formatos personalizados, armazenamento Chromium e configurações Desktop ficam em **`%APPDATA%\4PrintShops`**, separados da instalação. Reinstalar/atualizar conserva essa pasta. A desinstalação também a conserva (`deleteAppDataOnUninstall: false`). Preferências do navegador e do Desktop têm origens/perfis diferentes e não são migradas automaticamente. Documentos abertos e alterações em memória não são recuperados após reiniciar.

A janela conserva título, Início, abas, menus HTML e painéis; não recebe uma segunda barra nativa de menus. Pode ser redimensionada, minimizada e maximizada. O ícone experimental usa `public/img/logo.png`, convertido pelo builder para o executável. Uma identidade compacta definitiva pode substituir `win.icon` e o ícone da janela; o logo original não foi alterado.

## Gerar o instalador

```powershell
npm run desktop:pack
npm run desktop:dist
node scripts/check-desktop-package.mjs
```

`desktop:pack` prepara a pasta sem gerar o instalador. `desktop:dist` gera NSIS x64 com `--publish never`. Saídas:

- **`dist-desktop/4PrintShops-${version}-Windows-x64-Setup.exe`** — instalador da versão do `package.json`. O instalador `.1` aprovado permanece como artefato histórico, sem ser sobrescrito por esta etapa.
- `dist-desktop/4PrintShops-${version}-Windows-x64-Setup.exe.blockmap` — metadado do builder; sua existência não habilita atualizações.
- `dist-desktop/win-unpacked/4PrintShops.exe` — aplicativo empacotado antes da instalação.

A auditoria verifica assets essenciais, exclusões e igualdade dos arquivos de runtime no ASAR com o código atual. O pacote inclui apenas o frontend ativo, bibliotecas locais, processo Desktop e dependências de produção da atualização. Exclui testes, modelos CorelDRAW, backend legado, documentos PDF, chaves, logs, backups e dependências de desenvolvimento. As pastas geradas são ignoradas pelo Git. O repositório já continha `node_modules` versionado; esta tarefa não remove esse histórico. A fonte reproduzível de dependências é o `package-lock.json` da raiz.

## Instalar e desinstalar

1. Executar o `Setup.exe`, escolher a pasta e concluir. A instalação é **por usuário**, sem instalação de Node e sem solicitar elevação para instalar para todos os usuários.
2. Abrir pelo atalho da Área de Trabalho ou do menu Iniciar. O instalador não abre o programa automaticamente ao finalizar.
3. Para remover, usar Aplicativos do Windows ou `Uninstall 4PrintShops.exe` na pasta escolhida. A pasta de preferências é conservada.

O instalador desta fase **não possui assinatura Authenticode**. O Windows pode pedir confirmação de procedência/SmartScreen. Não foi desativada nenhuma proteção do Windows. A validação foi feita no Windows 10 x64, build 19045; instalação em outros computadores e Windows 11 precisa de aceitação complementar.

## Segurança e encerramento

`contextIsolation: true`, `nodeIntegration: false`, sandbox, `webSecurity`, CSP e validação do frame principal permanecem ativos. O renderer não navega para sites/arquivos externos, abre popups ou incorpora webviews; permissões adicionais são negadas. O protocolo só lê caminhos permitidos dentro do pacote. Os fuses desabilitam execução como Node, opções de Node por ambiente e inspeção do processo principal; exigem ASAR e sua integridade.

O estado de pendência continua vindo do PDFWorkspace. `beforeunload` aciona a confirmação nativa ao fechar, inclusive com PDF em aba oculta. Cancelar conserva o documento. Operações em curso/gravação devem terminar antes de fechar. Fechar o aplicativo termina seus processos Electron; não existe backend filho para deixar órfão. As portas de depuração usadas nos harnesses são somente opções explícitas dos testes, não da inicialização normal.

Cancelar o diálogo nativo de destino impede a gravação. Se o operador já confirmou o caminho e a escrita começou, cancelar depois a janela de exportação conserva a pendência, mas uma cópia que terminou de gravar pode permanecer no destino. Encerramentos forçados pelo sistema ou reinstalações externas não substituem o fechamento protegido; exportar/concluir o trabalho antes de executar manualmente outro instalador.

## Atualizações opcionais

Sem configuração, o aplicativo continua sem consultar a internet. Quando o provedor está configurado, a consulta ocorre uma vez ao abrir, sem aguardar a rede para mostrar a interface. Não abre modal, baixa ou instala automaticamente. Arquivo → **Verificar atualizações** permanece disponível só no Desktop.

`desktop/updates.mjs` mantém `autoDownload: false`, `autoInstallOnAppQuit: false` e `allowDowngrade: false`. O canal experimental é explícito. A compatibilidade foi conferida no código instalado de Electron **45.1.0**, electron-updater **6.8.11** e builder **26.18.0**, sem trocar dependências. O updater não declara um peerDependency que limite a versão de Electron; usa as APIs Electron net/session/app verificadas também nos testes reais. A versão de Electron usada pelo upstream em seus testes de desenvolvimento é diferente, portanto isso não promete compatibilidade com todas as versões futuras.

- Aviso discreto pelo módulo `public/js/app/desktop-updates.js`, com versão, **Ver novidades** e **Mais tarde**. Ocultar o aviso não perde a atualização; o menu permite retomá-la.
- Novidades em texto simples, em diálogo nativo. **Atualizar agora** autoriza o download. Progresso aparece no aviso e na barra de tarefas.
- Depois do download validado, uma segunda confirmação **Instalar e reiniciar / Mais tarde** é obrigatória. Fechar normalmente nunca agenda instalação.
- Antes e depois dessa confirmação, `desktop/work-state.mjs` solicita um estado novo de **todas as abas**, inclusive ocultas; falha/timeout/resposta inválida bloqueiam. O PDF expõe só leitura de `workspace.alterado` e ocupação/modal, sem mudar os motores. Futuras ferramentas podem fornecer `estadoTrabalho()` no seu controle.
- A interface fica temporariamente inerte durante essa decisão, e abertura/gravação nativas são recusadas nesse intervalo. Cancelar/falhar libera a interface. Pendências exigem exportar ou fechar a aba conscientemente pela confirmação existente; o atualizador não descarta nem exporta documentos.
- Erros de conexão, download, integridade e assinatura deixam o aplicativo aberto, com detalhes e possibilidade de tentar novamente. Nenhum caminho de documento do usuário é usado como destino do updater.

### Integridade e assinatura

O SHA-512 do instalador vem no metadado gerado pelo builder e é verificado no download. **Hash não autentica um editor:** o instalador também precisa de Authenticode incorporado, válido/confiável e correspondente a `publisherName` do pacote instalado. HTTPS e a confiança do Windows continuam exigidos para distribuição normal.

Foram corrigidos dois problemas da implementação anterior: em 6.8.11, `verifyUpdateCodeSignature` é uma **função**, não um booleano; além disso, seu verificador original pode aceitar quando não consegue executar PowerShell/ConvertTo-Json. O ensaio real com PATH mínimo demonstrou essa segunda condição. `desktop/update-signature.mjs` acrescenta uma prova que **falha fechada**, executando o PowerShell do Windows por caminho absoluto, sem shell e com timeout. Ausência/erro de PowerShell, resultado inválido, assinatura ausente/não confiável, catálogo, caminho ou editor divergentes bloqueiam. O verificador original é chamado somente depois dessa prova; não foi desligado. Nenhum certificado foi instalado no Windows nem proteção do sistema alterada.

### Configuração para distribuição assinada futura

`electron-builder.updates.cjs` exige provedor HTTPS e identidade de editor por `PRINTSHOPS_UPDATE_URL` e `PRINTSHOPS_PUBLISHER_NAME`, define `forceCodeSigning: true` e mantém a verificação de assinatura ativa. A chave/certificado deve ser configurada pelo operador conforme o builder, fora do Git. Gerar com:

```powershell
& .\node_modules\.bin\electron-builder.cmd --config electron-builder.updates.cjs --win nsis --x64 --publish never
```

Sem certificado/identidade/provedor esse caminho deve falhar; não é o build unsigned normal. Para habilitar a consulta na instalação assinada, criar `%APPDATA%\4PrintShops\updates.json` com `enabled: true`, `url` HTTPS e `channel: "experimental"` do provedor autorizado. `enabled: false` desativa a consulta. Não colocar credenciais na URL. O pacote deve conter seu `app-update.yml` com publisher. Nada dessa distribuição foi publicado ou configurado na instalação aprovada `.1` nesta etapa.

### Laboratório local `.2` → `.3`

`electron-builder.update-lab.cjs` compila **4PrintShops Update Lab**, appId, cache do updater, atalhos e perfil separados (`%APPDATA%\4PrintShops Update Lab`). Não substitui a instalação aprovada nem seus formatos. Somente esse pacote explicitamente marcado, de nome `4printshops-update-lab`, aceita HTTP em **127.0.0.1**; configurações do usuário não habilitam essa exceção no app normal. A verificação de assinatura permanece ativa, com editor de teste declarado.

```powershell
$env:PRINTSHOPS_LAB_VERSION='1.0.0-experimental.2'
& .\node_modules\.bin\electron-builder.cmd --config electron-builder.update-lab.cjs --win nsis --x64 --publish never
$env:PRINTSHOPS_LAB_VERSION='1.0.0-experimental.3'
& .\node_modules\.bin\electron-builder.cmd --config electron-builder.update-lab.cjs --win nsis --x64 --publish never
node scripts/serve-update-lab.mjs
```

O servidor lê apenas nomes permitidos dos artefatos, com suporte a Range, e escuta `127.0.0.1:18083`. Serve a pasta de saída da `.3`, sem upload/publicação. Saídas reais:

- `dist-desktop/update-lab/1.0.0-experimental.2/4PrintShops-Update-Lab-1.0.0-experimental.2-Setup.exe`
- `dist-desktop/update-lab/1.0.0-experimental.3/4PrintShops-Update-Lab-1.0.0-experimental.3-Setup.exe`
- Cada pasta inclui `experimental.yml`, `.exe.blockmap` e `win-unpacked/resources/app-update.yml` gerados pelo builder.

Teste automatizado técnico: `node tests/desktop-update-real.mjs`. Usa os pacotes reais em uma pasta temporária e diálogos Windows, sem mocks no pacote, sem Node no PATH do app e com limite de banda local para observar progresso. Somente reutiliza perfil marcado como pertencente ao harness. A configuração de produção não recebe hooks de teste. O teste não equivale a homologação de assinatura comercial.

**Limite obrigatório:** sem certificado Authenticode confiável, a instalação/reinício via updater não podem ser aprovados. O laboratório deve rejeitar o instalador unsigned. Instalar `.3` manualmente após encerrar conscientemente o trabalho é uma verificação complementar de versão/preferências, não uma atualização automática bem-sucedida. Resultados efetivamente obtidos estão no registro de 03/11/2026 de `ESTADO_PROJETO.md`.

Ensaio final de 03/11/2026 aprovado até esse limite: `.2` instalada e `.3` detectada; novidades/adiamento, progresso intermediário e download reais; rejeição de Authenticode ausente e SHA-512 adulterado; erro de conexão/recuperação; PDF alterado preservado. `.3` foi instalada **manualmente**, reaberta e identificada, conservando preferências e um formato personalizado que a biblioteca real recarregou. Resultados em `%TEMP%/4printshops-update-real-sJD9us/result.json`, com `automaticInstall: false` e `manualInstall: true`. Log `%TEMP%/4printshops-update-real-complete.log`. A instalação Lab de inspeção permanece na subpasta `installed/`; aplicativo/servidor terminaram. Repetir o harness só reutiliza perfis marcados como pertencentes a ele; não apontá-lo para um perfil de atendimento.

Roteiro manual: instalar o **Lab `.2`**, iniciar o servidor da `.3`, abrir o Lab e conferir aviso; Ver novidades → Mais tarde, continuar em PDF/QR; menu → Atualizar agora, acompanhar progresso e conferir rejeição segura de assinatura. Desligar/religar o servidor e repetir a consulta. Com pacotes devidamente assinados, repetir também pronto → Mais tarde, pendência em PDF oculto, exportar/fechar conscientemente a aba, cancelar/confirmar instalação, reiniciar, confirmar versão `.3` e formatos. Esses últimos passos via updater aguardam certificado confiável. Não usar o laboratório como distribuição para clientes.

### Certificado de teste em VM descartável — preparação em 03/11/2026

O uso de identidade autoassinada é viável **se confiada somente dentro do guest**, preservando o verificador Authenticode do aplicativo. Foi preparado o kit `dist-desktop/authenticode-vm-kit/` (66 fontes selecionadas, sem dependências, chaves ou documentos), com fingerprint do host e hashes. A configuração separada `electron-builder.authenticode-lab.cjs` exige guest reconhecido e confirmação explícita, certificado por thumbprint e `forceCodeSigning: true`; mantém `verifyUpdateCodeSignature: true`. Não modifica os builds normais ou seu mecanismo de segurança.

O host atual não tem um ambiente virtual acessível identificado. **Nenhuma confiança/certificado foi criado/importado, nenhum build assinado ou update automático foi executado nesta preparação.** Há 176 testes Node aprovados e o bloqueio do host foi verificado em leitura. Falta disponibilizar o guest e executar o procedimento, sem alterar políticas do host. Instruções de provisionamento, criação de chaves não exportáveis só no guest, builds `.2`/`.3` com a mesma identidade, feeds negativos, evidências e descarte: [LABORATORIO_AUTHENTICODE.md](LABORATORIO_AUTHENTICODE.md). Isso não substitui o certificado e a homologação comercial.

## Testes e aceitação

```powershell
npm test
npm run test:browser
npm run test:desktop
```

`test:desktop` usa Electron real, perfil temporário e diálogos simulados no bootstrap exclusivo dos testes, que não é empacotado. Verifica ponte, gravação real, cancelamento/falha, proteção, sandbox, origem, persistência e encerramento.

`tests/desktop-installed.mjs` é o harness técnico da instalação temporária identificada em `%TEMP%/4printshops-instalacao-teste.txt`; usa o executável instalado, sem mocks no pacote e sem Node no PATH do aplicativo. A automação Windows restringe-se ao PID/título/controles da instância de teste. Seus resultados e artefatos estão registrados em `ESTADO_PROJETO.md`; não devem ser confundidos com a aprovação manual em uma máquina de atendimento.

Aceitação manual: abrir o programa e tentar uma segunda abertura; importar PDF real; girar/reordenar/excluir/unir/redimensionar; conferir a cópia preparada de impressão; salvar em pasta escolhida e cancelar outra tentativa; alternar Início/PDFs/QR; fechar e cancelar com trabalho pendente; repetir com janela pequena/maximizada; reiniciar e conferir preferências; desinstalar/reinstalar e conferir dados. Comparar PDFs complexos no Acrobat/RIP antes de liberar para produção.

Para distribuição comercial futura ainda faltam assinatura de código e identidade final, revisão de licenças/avisos das bibliotecas, homologação em outras máquinas, canal de atualizações autorizado e seu teste entre versões, política de suporte/recuperação e aceitação de desempenho/fidelidade em arquivos reais. Não há autenticação, serviço pago ou novas ferramentas de edição nesta implementação.

Referências técnicas: [segurança do Electron](https://www.electronjs.org/docs/latest/tutorial/security), [protocolos locais](https://www.electronjs.org/docs/latest/api/protocol), [instalador NSIS](https://www.electron.build/nsis/), [atualizações na documentação v26](https://www.electron.build/v26/docs/features/auto-update/), [assinatura Windows v26](https://www.electron.build/v26/docs/win/). Documentação online mais recente pode descrever APIs de builder/updater 27/7; a implementação desta etapa foi conferida nas versões locais 26.18.0/6.8.11.
