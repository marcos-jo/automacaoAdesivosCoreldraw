# Combinar PDFs e imagens — experimental.14

## Validação operacional informada pelo usuário

Em 04/11/2026, o usuário informou que instalou a experimental.14 nos computadores da gráfica, testou as funções rotineiras e orientou os funcionários a registrar problemas e sugestões. As limitações técnicas abaixo continuam documentadas. O checkpoint conjunto .13/.14 foi autorizado para commit local e sincronização manual, sem nova implementação.

## Preservação e escopo

Implementação sobre `be77d0f`, branch `main`. Os 20 arquivos pendentes da experimental.13 foram auditados e copiados integralmente, com inventário SHA256, em `.desktop-test/combine14/baseline13-20261104-162344/`. O staging permaneceu vazio. A interface .13 foi ampliada pela janela de combinação existente.

A seleção dentro do aplicativo aceita PDF, JPG/JPEG e PNG em uma operação. Um único PDF continua abrindo diretamente. Vários arquivos ou uma imagem abrem o rascunho de combinação. No Desktop, o diálogo nativo usa os mesmos tipos e seleção múltipla, pelo canal IPC de abertura já existente. Os caminhos não são devolvidos ao renderer. Arquivos originais permanecem intactos.

## Uso no balcão

1. Em PDFs, escolher **Abrir arquivos** e selecionar os PDFs e imagens. O botão Abrir PDF do Início também utiliza essa seleção ampliada. Para acrescentar depois, usar **Combinar arquivos** na ferramenta contextual.
2. Conferir a lista por página. Ela mostra nome, tipo, número da página na origem, total de páginas da origem e dimensões físicas. Arrastar páginas ou usar ↑/↓. O botão × remove a página. Remover arquivo retira todas as páginas daquela origem.
3. Usar Adicionar mais arquivos quando necessário. A seleção inválida é recusada sem apagar o lote já analisado. Uma falha de miniatura oferece tentativa novamente e não impede as outras páginas.
4. Escolher Criar novo PDF combinado ou Acrescentar ao PDF aberto. Substituir um documento editado exige a confirmação de descarte existente. Cancelar conserva o documento e o rascunho.
5. Escolher Manter tamanho original, Caber em A4 proporcionalmente ou Preencher A4 com possível corte. Preencher exige marcar a aceitação do corte. A4 acompanha a orientação visual de cada página. Ao acrescentar, a opção A4 também padroniza as páginas do documento anterior.
6. Clicar Combinar. O resultado passa ao PDFWorkspace, com miniaturas, navegação, organização, giro, redimensionamento e alterações pendentes. Salvar cópia exporta somente por solicitação do operador.
7. Abrir Arquivo → Impressão, selecionar a fila e suas propriedades. O fluxo atual recebe o documento de trabalho inteiro em uma operação. O primeiro ensaio deve usar Microsoft Print to PDF. Impressões físicas dependem de decisão e conferência pessoal do operador.

## Imagens, dimensões e cores

JPEG RGB/cinza é incorporado com os bytes JPEG comprimidos originais. As oito orientações EXIF são representadas por matrizes, incluindo espelhamentos. PNG suporta transparência através de máscara alfa PDF. Páginas PDF são copiadas com seus recursos vetoriais e fontes, sem rasterização integral para combinar.

A dimensão inicial das imagens usa **150 ppi**, uma convenção explícita de colocação. O DPI declarado no arquivo não decide a escala. A lista mostra as dimensões resultantes e o operador pode escolher A4. Essa convenção não certifica resolução suficiente para impressão. Orientação EXIF também é lida em PNG quando presente.

Limites de admissão são 512 MiB no lote, 40 megapixels e 20.000 pixels por lado de imagem. O limite de bytes é verificado antes da leitura quando o tamanho está disponível e novamente com os bytes efetivos. O PDF combinado também é limitado a 512 MiB. O ensaio de tamanho incluiu um PDF de 350 páginas e PNG de aproximadamente 3 megapixels. O teto de 512 MiB não foi submetido a teste de estresse nesta rodada.

Casos bloqueados explicitamente incluem JPEG CMYK, perfil ICC em JPEG/PNG, PNG de 16 bits, animado, gama não padrão ou cromaticidade personalizada sem declaração sRGB. PDFs com intenção de saída global, como alguns PDF/X, ou camadas opcionais são bloqueados porque sua aparência/cor não está homologada nessa combinação. A interface explica a incompatibilidade. Não há conversão automática de cor, perfis ou fontes.

Formulários comuns e widgets sem suporte seguro exigem uma cópia para impressão já conferida. PDFs assinados mantêm o diagnóstico da origem e o mapeamento das páginas afetadas, inclusive após remoção/reordenação no rascunho. A combinação reescreve o PDF e pode invalidar assinaturas. O fluxo de preparar cópia para impressão permanece disponível no workspace. A combinação não certifica validade criptográfica.

Marcadores, destinos, árvore de acessibilidade e outros metadados globais não têm preservação integral garantida pelo `copyPages` de pdf-lib 1.17.1. O aviso está nos detalhes da janela. Ajustes A4 reutilizam o motor de redimensionamento e seus bloqueios de compatibilidade. Uma falha restaura o documento anterior. Não se contornam esses bloqueios nessa janela.

## Organização dos módulos

- `pdf-image-import.js` analisa formatos, limites, EXIF e incompatibilidades de cor, gerando uma página PDF por imagem.
- `pdf-combine-session.js` mantém origens e sequência do rascunho, valida lotes atomicamente e monta o documento com cópia vetorial. Guarda bytes de prévia independentes para evitar serializações concorrentes ao combinar.
- `pdf-merge.js` mantém os IDs existentes e oferece a interface de combinação. Miniaturas usam PDFPreview, fila serial, carregamento por visibilidade e cache limitado a 48 imagens. O fechamento cancela tarefas e libera recursos. As linhas da lista são leves, porém não são virtualizadas integralmente.
- `pdfs.js` reutiliza o bloqueio operacional, confirmação de descarte, atualização visual e redimensionador existentes. `PDFWorkspace.adicionarArquivos` ganhou opções explícitas de substituição e diagnóstico das origens. O comportamento padrão de acrescentar PDFs continua disponível.
- `desktop/files.mjs`, `preload.cjs`, `app.mjs` e `desktop-client.js` ampliam somente a seleção pelo diálogo autorizado. PDFium, fontes C#, DEVMODE, HDC, PS RAW e atualizadores não receberam alterações funcionais.

## Testes e evidências locais

- 258 testes Node aprovados, incluindo EXIF 1–8, JPEG sem recompressão, PNG alfa, bloqueios ICC/CMYK/perfis PDF/camadas/formulários, rollback, ordem, exclusão por página/arquivo, A4 nos dois modos, diagnóstico de assinatura, giro anterior, pendências e 350 páginas com imagem grande.
- 41 verificações C# aprovadas. Os testes de política e DEVMODE não enviam trabalhos.
- Chrome completo aprovou as regressões de PDF, QR, abas, 100/350 páginas, organização, autoscroll, edição, exportação, redimensionamento, ActualText sintético e cópias de impressão/assinaturas.
- O teste específico Chrome aprovou seleção mista, miniaturas, movimentos por controles e eventos de drag-and-drop, confirmação de corte, A4, exportação, cancelamento da substituição e acréscimo posterior. Comparação de amostras RGBA com o decoder do navegador apresentou diferença máxima de um nível no JPEG EXIF 6 e zero no PNG transparente. Essas amostras não equivalem a homologação completa de gerenciamento de cor.
- Electron smoke aprovou salvamento nativo, falhas/cancelamentos, proteção, preferências e cópia para impressão. O relatório do fluxo instalado está em `.desktop-test/combine14/installed/`, com o log final separado das tentativas inconclusivas.
- Hold Test instalada aprovou upgrade .13 → .14 e reaplicação do pacote final, no mesmo diretório/registro/perfil. Preferência anterior preservada. Importação de PNG pelo diálogo nativo, acréscimo dos PDFs/JPEGs, ordenação, A4, exportação, giro, cancelamento de substituição, acréscimo e impressão exclusivamente virtual aprovados. O trabalho único gerou oito páginas reabertas no parser PDF. O workspace permaneceu pendente e a impressora padrão Konica foi preservada. Resultado final em `installed-approved.log` e `installed/result.json`. O teste aceita o destino padrão novo e identificável em Downloads e conserva cópia da saída em `installed/combinado-virtual-validado.pdf`.
- Pacote instalado auditado com 435 entradas ASAR. SHA256 do Setup final conferido. Sintaxe dos JS/CJS/MJS alterados e `git diff --check` aprovados. Nenhuma alteração nos fontes nativos ou motores PDFium/DEVMODE/HDC/PS RAW.

Fixtures reproduzíveis em `tests/helpers/pdf-combine-fixtures.mjs`, sem PDFs de clientes nem certificados. Artefatos e logs em `.desktop-test/combine14/`, ignorados no Git. Capturas da combinação e do resultado A4 em `browser/01-combinar-pdfs-imagens.png` e `browser/02-combinado-a4.png`, além das capturas do aplicativo instalado.

Os testes revelaram e corrigiram o diálogo invisível quando o painel contextual estava recolhido. Ele passou à raiz do editor. Os testes antigos foram adaptados à análise assíncrona e à lista por página, mantendo as verificações de rollback, concorrência e proteção. A automação de impressão virtual aguarda a gravação pelo spooler após o retorno do trabalho. Tentativas anteriores procuraram o arquivo na pasta de teste, enquanto o diálogo salvou em Downloads. A automação foi ajustada para aceitar o destino padrão único. A comparação detalhada de amostras de imagem foi concluída no Chrome. Uma tentativa de repetir essa comparação auxiliar no Electron teve timeout e não conta como aprovação adicional de pixels. O fluxo instalado final exercita as imagens no documento, exportação, prévia e impressão virtual. As tentativas inconclusivas ficam registradas nos logs. Nenhum envio físico foi autorizado ou executado durante o desenvolvimento.

## Explorador do Windows — próxima etapa

A combinação interna foi priorizada. O menu **Combinar em PDF com 4PrintShops** ainda não foi registrado no Explorer. As associações existentes de PDF foram preservadas e JPEG/PNG não receberam novas associações.

Uma próxima implementação precisa receber a seleção completa com identidade de lote e encaminhá-la à instância única. Um comando estático baseado apenas em `%1` não demonstra entrega atômica de todos os itens. A escolha depende do modelo de seleção e dos limites do Shell. A Microsoft documenta `Document` e `Player`, com diferenças por implementação. [Modelo de seleção do Shell](https://learn.microsoft.com/en-us/windows/win32/shell/how-to-employ-the-verb-selection-model).

No Windows 11, o menu clássico continua acessível em Mostrar mais opções. A integração ao menu moderno utiliza IExplorerCommand e identidade de aplicativo, exigindo trabalho e validação adicionais. Não foram registrados COM, verbos ou mudanças globais do Explorer nesta etapa. [Documentação Microsoft sobre o menu Windows 11](https://blogs.windows.com/blog/2021/07/19/extending-the-context-menu-and-share-dialog-in-windows-11/).

O planejamento inclui adesão explícita, remoção apenas das entradas próprias, teste de múltiplos arquivos e nomes acentuados, entrega antes/depois da interface pronta e filas de lotes distintas. SMB real e os limites UNC/FQDN/IPv4 do fluxo anterior permanecem pendentes. Seleção via diálogo nativo pode acessar locais autorizados pelo Windows, mas compartilhamento SMB real não foi ensaiado aqui.

## Distribuição

Instalador `dist-desktop/hold-native-test/4PrintShops-Hold-Test-1.0.0-experimental.14-Windows-x64-Setup.exe`.

SHA256 `b8524db89eaf8acf55f7773dd733d0e1447ca9647b2c797845d83b5455c95e4c`.

Mesma identidade `br.com.4printshops.holdtest`, registro NSIS, executável e perfil `%APPDATA%/4PrintShops Hold Test`. A versão da raiz permanece `.6`. O pacote é experimental e sem assinatura comercial. Authenticode do atualizador integrado permanece obrigatório e instalação/reinício automáticos continuam não homologados. Nenhum commit ou push nesta etapa.
