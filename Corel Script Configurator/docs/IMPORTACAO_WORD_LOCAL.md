# Importação local — PDFs, imagens e Word

Atualizado em 08/11/2026, experimental.32 com motor LibreOffice privado, instalada sobre a .31 com preferências preservadas. Base de desenvolvimento: `main`, checkpoint `e350f12`. Consolidação local da .32 autorizada e documentada em `ESTADO_PROJETO.md`; sem push ou impressão.

## Como usar

Clique **Importar**, na tela inicial, na barra do editor ou em **Arquivo → Importar (Ctrl+O)**. Existe uma única entrada para **PDF, JPG/JPEG, PNG, DOCX e DOC**. O botão permanece acessível quando já há um documento aberto. É possível selecionar vários tipos juntos; não é necessário escolher um conversor ou clicar em um botão separado de Word.

- **Um PDF:** segue diretamente para o editor, usando a proteção de descarte existente.
- **Imagens:** são preparadas pelo importador já existente, respeitando suas verificações e dimensões iniciais a 150 ppi, sem confiar no DPI declarado.
- **Word:** é convertido automaticamente para PDF no próprio computador, usando LibreOffice. O nome preparado termina em `_WORD.pdf`.
- **Imagens, Word ou seleção múltipla:** abrem **Conferir e combinar arquivos**, com miniaturas, sequência, dimensões e opção de remover/reordenar páginas. Escolha acrescentar ao documento aberto ou criar um novo, confira e clique **Combinar**. Os detalhes de Word mostram páginas e fontes, com os limites da conversão.

**Adicionar mais arquivos** nessa janela utiliza a mesma identificação/conversão automática. Um erro na preparação de um lote não importa silenciosamente apenas parte da seleção. Cancelamento/falha preservam o documento anterior; criar um novo documento com edições pendentes continua exigindo confirmação de descarte. A seleção no diálogo nativo tem sua ordem preservada na preparação; a ordem exibida inicialmente pelo Windows depende do seletor. Confira a sequência antes de combinar.

Não há download automático. Use a exportação existente depois de revisar o documento. PDFs não são rasterizados pelo simples ato de importar. Motores de edição, redimensionamento, montagem, Step & Repeat, N-up, livreto e impressão permanecem inalterados.

## Desktop e navegador

A conversão Word está disponível no **Desktop**. No navegador, PDFs e imagens mantêm o fluxo anterior; escolher DOCX/DOC explica a necessidade do Desktop e recusa a seleção sem alterar parcialmente o documento. Não existe uma API HTTP de conversão nem envio a conversores online.

No Desktop, a preparação tem aviso de andamento e **Cancelar importação**. Conversões Word são sequenciais e o editor permanece protegido durante a preparação. O seletor do Windows também pode ser cancelado. A leitura/execução ocorre no processo principal, com caminhos fornecidos exclusivamente pelo diálogo nativo; o renderer não fornece caminhos, executáveis ou argumentos de conversão.

## LibreOffice incluído

A partir da **experimental.32**, o Desktop leva uma cópia privada do LibreOffice **26.8.2.2 x64** em `resources/libreoffice`. Basta instalar o 4PrintShops; não é necessário instalar um Office separadamente para importar Word. O LibreOffice global previamente autorizado permanece nesta máquina e não é utilizado pelo pacote .32.

O aplicativo confere a integridade do motor antes do primeiro uso em cada execução. Motor ausente, incompleto ou alterado bloqueia a conversão e orienta reinstalar o 4PrintShops, sem recorrer silenciosamente a outro Office. A cópia privada mantém suas bibliotecas e fontes na própria pasta e não registra atalhos/associações Word ou fontes no Windows. Perfil e cache da conversão ficam em pasta temporária exclusiva.

Em desenvolvimento, preparar os recursos com `node scripts/prepare-libreoffice.mjs --download`. Somente quando a pasta privada está ausente o código de desenvolvimento mantém a descoberta em Program Files ou `PRINTSHOPS_LIBREOFFICE` (caminho absoluto de soffice.com/soffice.exe). Essa variável não redireciona o aplicativo empacotado. Não há download automático durante o uso. Guia de build, licenças, fontes, tamanho e validação: [LIBREOFFICE_PRIVADO_EXPERIMENTAL32.md](LIBREOFFICE_PRIVADO_EXPERIMENTAL32.md).

## Segurança e limites

- Até **128 arquivos** e **512 MiB** de entrada/preparação agregada. Word mantém limite de **64 MiB**, PDF convertido de **128 MiB** e **1–2.000 páginas**. O limite de saída é conferido novamente depois das conversões.
- Arquivos originais são somente lidos. Cada Word é copiado para uma pasta exclusiva com nome controlado; saída nunca é gravada ao lado do cliente. O hash do Word original é comparado novamente após converter; mudanças externas causam recusa. Temporários são removidos em sucesso, falha e cancelamento; falha de limpeza é informada.
- Perfil LibreOffice separado por conversão, sem usar/alterar o perfil pessoal. Configurações bloqueiam macros, conteúdo ativo OLE/DDE e atualização de vínculos/campos/gráficos. Sem shell, servidor UNO ou opções vindas do documento/renderer.
- DOCX tem inspeção limitada de ZIP/XML, limites descompactados e recusa de pacotes com macros, ActiveX, objetos incorporados ou relações externas, inclusive hyperlinks. ZIP64 e pacotes divididos não são suportados. DOC exige cabeçalho OLE/WordDocument, com inspeção binária limitada.
- Word aceita caminhos absolutos de disco local; UNC, URLs, ADS e dispositivos são recusados nesse conversor. PDFs/imagens seguem os caminhos escolhidos pelo diálogo nativo e as verificações existentes de conteúdo.
- Prazo do processo Word: **120 segundos por arquivo**. Cancelar/deadline no Windows encerra somente a árvore criada para aquela conversão, sem terminar outras sessões do LibreOffice. O bloqueio permanece até o processo terminar. Operações de arquivo/rede já bloqueadas no sistema podem demorar; a rotina não usa credenciais nem modifica a configuração de rede.
- A saída é analisada pela biblioteca PDF local antes de chegar ao renderer. A combinação e suas miniaturas usam os parsers e guardas anteriores. Arquivos/recursos não suportados continuam recusados.
- Perfil separado e configurações de segurança não constituem isolamento completo do sistema operacional. Manter o 4PrintShops e seu motor privado atualizados; a versão é fixada no build. Queda abrupta do Windows/aplicativo pode deixar temporários; não há limpeza automática de sessões antigas.

## Fontes, paginação e conferência

O número de páginas do PDF é medido. No DOCX, a contagem salva nos metadados do Word é mostrada como referência possivelmente desatualizada. No DOC, a contagem original permanece **não verificada**.

A lista de fontes do DOCX pode incluir fontes declaradas sem uso. A inspeção do PDF informa BaseFont e presença de arquivos de fonte em seus descritores, incluindo fontes compostas. Isso não comprova correspondência de glifos, licença de incorporação ou ausência de substituição. Fontes Type3/diretas/incomuns ou texto em curvas podem não aparecer. Para DOC, fontes solicitadas ficam **não verificadas**.

Compare o PDF com o Word ou uma referência, especialmente quebras, tabelas, imagens, cabeçalhos/rodapés e página final. Depois de combinar, use a visualização e o zoom do editor para conferir o resultado antes de exportar/produzir. Não há garantia automática de fidelidade ao Microsoft Word, ICC/CMYK/RIP, sangria ou resolução suficiente das imagens.

## Arquivos envolvidos

- `desktop/import-files.mjs`: diálogo unificado, preflight do lote, identificação de extensão, conversão sequencial de Word, limites e resultado transacional.
- `desktop/word-import.mjs`: conversor local e leitura limitada compartilhada, sem mudar sua política de conversão.
- `desktop/app.mjs` / `preload.cjs`: canais restritos de importar/cancelar, bloqueio e cancelamento na perda do renderer. Canais anteriores de PDF/Word conservados por compatibilidade.
- `public/js/app/desktop-client.js`: seleção, andamento/cancelamento e entrega de Files/relatórios; shell/menus: rótulo Importar.
- `public/js/tools/pdf-interface.js`, `pdfs.js`, `pdf-merge.js`: entrada única, identificação automática e conferência do lote com informações Word. A classe de revisão Word anterior permanece no código, sem botão separado na interface atual.
- Testes de importação/lote, integração do editor e ensaio real `tests/desktop-word-smoke.mjs`. Config Hold Test identifica a versão experimental.32; resolvedor privado em `desktop/libreoffice-runtime.mjs` e preparo em `scripts/prepare-libreoffice.mjs`.

## Testes e reprodução

**437 testes Node aprovados**, zero falhas/skips. Cobrem seleção mista, ordem, leitura preservada, falha sem resultado parcial, limites antes/depois da conversão, cancelamento/deadline, origem identificada no erro, diálogos, proteção do editor e as regressões anteriores.

Na etapa .31, Chrome headless curto passou em abas PDF 100/350, menus, exportação, proteção e QR. Na etapa .32, ensaio Electron real com motor privado e descoberta externa desativada passou com **PDF + PNG + JPG + DOCX + DOC → sete páginas**, além de Word sozinho, adicionar Word na combinação, cancelamento da seleção/conversão, Word inválido no lote, recusa/aceite do descarte e exportação PDF reaberta. Ordem e bytes de todos os cinco originais foram preservados. Os relatórios também exercitam fonte inexistente e metadados de nove páginas para um PDF de duas.

```powershell
npm test
node tests/desktop-word-smoke.mjs
node scripts/test-private-office.mjs dist-desktop/hold-native-test/win-unpacked
```

Preparar o motor privado antes do ensaio. Os módulos reais do ASAR e seu runtime também converteram DOCX/DOC com caminhos externos inválidos, preservação dos originais e integridade após uso. Falta validar em Windows limpo sem Office. O ensaio utiliza somente documentos sintéticos e perfil Electron temporário. A pasta de evidências é informada ao finalizar. Não imprime. Para conferir os PDFs Word nos motores existentes:

```powershell
node tests/desktop-word-workflows.mjs 'C:\pasta-do-ensaio\docx.pdf' 'C:\pasta-do-ensaio\doc.pdf'
```

Permanece pendente a validação em massa com Word reais na gráfica e a homologação de cor/RIP/produção. Não se presume suporte universal a documentos Word. Não houve commit, push, publicação ou impressão física/virtual nesta etapa.

## Referências oficiais

- [Parâmetros de inicialização e perfil separado](https://help.libreoffice.org/latest/en-US/text/shared/guide/start_parameters.html).
- [Opções de exportação PDF](https://help.libreoffice.org/latest/en-US/text/shared/guide/pdf_params.html).
- [Schema de segurança Common](https://github.com/LibreOffice/core/blob/master/officecfg/registry/schema/org/openoffice/Office/Common.xcs).
- [Schema Writer, atualização de vínculos](https://github.com/LibreOffice/core/blob/master/officecfg/registry/schema/org/openoffice/Office/Writer.xcs).
