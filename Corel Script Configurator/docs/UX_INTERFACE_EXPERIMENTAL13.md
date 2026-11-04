# Interface clássica — experimental.13

## Estado em 04/11/2026

A experimental.13 reorganiza Início, área de PDFs e janela de impressão. A base é o checkpoint `be77d0f`, em `main`, com staging inicialmente vazio. O mecanismo de impressão aprovado na .12 foi preservado. Nenhum trabalho foi enviado a impressoras físicas nesta etapa.

A organização clássica do Acrobat de aproximadamente 2020–2021 foi usada como referência ergonômica. A identidade visual continua sendo do 4PrintShops. As referências consultadas incluem a tela inicial do Acrobat 2020 publicada em junho de 2020, a disposição clássica da janela de impressão ilustrada no manual Konica e a documentação Adobe de escala. As imagens de terceiros não foram incorporadas ao aplicativo. [Acrobat 2020 na B&H](https://www.bhphotovideo.com/explora/photography/news/adobe-announces-acrobat-standard-and-pro-2020), [janela clássica ilustrada pela Konica](https://manuals.konicaminolta.com.au/phpmanuals/2100/contents/guide01_03_05.html), [escala na documentação Adobe](https://helpx.adobe.com/in/acrobat/kb/scale-or-resize-printed-pages.html).

## Alterações concluídas

- Início compacto, menu superior tradicional e ação Abrir PDF ligada à importação existente. Os módulos PDF e QR Codes continuam acessíveis. A aplicação ainda não possui histórico de documentos recentes.
- Área de PDFs clara, resumo técnico compacto, ferramentas frequentes e rotação no topo, miniaturas à esquerda recolhíveis e painel contextual à direita. Zoom, navegação, abas e indicação de alterações pendentes foram preservados.
- Impressão em duas colunas. A esquerda reúne impressora, propriedades, cópias, páginas, escala, orientação e detalhes recolhíveis. A direita mostra o conteúdo real do PDF sobre o papel físico, área imprimível conhecida, avisos de corte e navegação independente entre páginas.
- Rodapé fixo com Cancelar e Imprimir. A coluna de configurações tem rolagem própria. Os botões ficaram visíveis nos três tamanhos avaliados, incluindo os cenários em que ficavam fora da área visível na .12.
- Escalas 100%, 120% e personalizada usam o fluxo existente. Cópias e orientação são apresentadas a partir do driver e alteradas nas propriedades nativas. A impressão continua abrangendo o documento completo. Navegar na prévia não seleciona um intervalo de impressão. Ajustar ao papel e intervalos de páginas permanecem fora desta etapa.
- O diagnóstico opcional Konica/Hold continua disponível em seção recolhível. A abertura da cópia atual no leitor externo permanece no menu existente.

## Separação entre prévia e impressão

`desktop-print-preview.js` reutiliza o ciclo de vida de `PDFPreview` para uma cópia em memória dos bytes atuais do workspace. Renderiza uma página por vez, limita a resolução da prévia e libera documento e tarefas ao fechar. A página inicialmente mostrada corresponde à seleção atual. A navegação dentro da janela é independente da seleção do editor.

Os pixels da prévia são usados somente na interface. O envio continua obtendo os bytes atuais do PDFWorkspace pelo fluxo aprovado. PDFium, auxiliares C#, DEVMODE completo, HDC e captura/validação/envio PS RAW foram preservados. Os 16 arquivos do inventário local `engine-baseline.json` permaneceram idênticos por SHA256.

A prévia representa a geometria calculada pelo aplicativo. Transformações privadas do driver, imposição, gerenciamento de cor e acabamentos não são simulados. Quando o zoom público do driver prevalece, a limitação continua explicitada. A inspeção física dessas opções permanece vinculada ao driver/equipamento.

## Arquivos

Novos módulos de interface em `public/js/app/desktop-print-interface.js` e `desktop-print-preview.js`. Integração em `desktop-print-ui.js`, `main.js` e no callback de snapshot de `pdfs.js`. Início/menu em `shell.js` e `menus.js`. Área PDF em `pdf-interface.js`, `style.css` e `print-config.css`.

`electron-builder.config.cjs` inclui os dois novos módulos e `desktop/security.mjs` permite exatamente esses recursos no protocolo local. As demais restrições e a CSP foram preservadas. O builder Hold Test altera somente sua versão para .13. Novos testes em `desktop-print-presentation.test.mjs` e `desktop-classic-ui-installed.mjs`. Os testes de orientação e de comparação de pixels foram ajustados à apresentação atual.

## Validação executada

| Verificação | Resultado |
| --- | --- |
| Testes Node | 246 aprovados, zero falhas |
| Verificações C# | 41 aprovadas, sem APIs de envio executadas nesses checks |
| Chrome curto | PDF/QR, abas, 100/350 páginas, exportação e proteção aprovados |
| Chrome completo final | Reorganização, miniaturas, rolagem/autoscroll, giro, exclusão, união, extração, exportação, redimensionamento nos três modos/quatro rotações, formatos, ActualText sintético e assinaturas aprovados |
| Electron smoke | Origem segura, PDF.js, miniaturas, giro, abas, salvamento nativo, cancelamento/falha, proteção de fechamento, preferências e cópia para impressão aprovados |
| Aplicativo instalado | Início/Abrir PDF, painéis/zoom, janela de impressão, navegação da prévia, cancelamento/reabertura e documento editado pendente aprovados |
| Propriedades nativas | Konica e Epson consultadas, confirmação/cancelamento e atualização da prévia exercitados sem envio físico |
| Impressão virtual | Botão Imprimir executado exclusivamente na Microsoft Print to PDF, com saída sintética de duas páginas |
| Layout | 1366 × 768 e 1920 × 1080 aprovados. Emulação Electron de 125% em 1092 × 614 pixels lógicos também aprovada |
| Preferências | Perfil anterior, impressora padrão e DEVMODEs globais/por usuário monitorados e preservados |
| Pacote instalado | Auditoria de 433 entradas ASAR aprovada, com recursos previstos e exclusões de documentos/credenciais/testes |
| Sintaxe e Git | Arquivos JS/CJS/MJS alterados verificados e `git diff --check` aprovado |

As capturas anteriores registram a .12 instalada. O primeiro upgrade passou para .13 no mesmo registro NSIS, diretório e perfil. Durante a validação, a primeira compilação revelou a ausência dos dois módulos na lista de recursos permitidos pelo protocolo Electron. A lista exata foi corrigida, ganhou teste e o pacote final foi reaplicado e validado na instalação .13 existente. O relatório final registra essa reaplicação, enquanto `before/result.json` conserva a origem .12.

`chrome-full-recheck.log` registra uma execução inconclusiva por download que permaneceu temporário. A execução isolada final em `chrome-full-final.log` passou integralmente. Uma tentativa anterior também revelou que o teste de comparação de pixels guardava a escala global da última renderização. O teste passou a consultar a escala do próprio canvas comparado, conservando o limite original de diferença de 0,001. O teste de orientação passou a esperar a nova cor amarela clara, conservando as verificações de gravidade e conteúdo.

Os logs e capturas são artefatos locais ignorados em `.desktop-test/ux13/`. O teste ActualText com documento real de cliente não foi executado nesta rodada. Os cenários sintéticos passaram. Não foram usados documentos de clientes.

## Instalador e comparação visual

Setup final em `dist-desktop/hold-native-test/4PrintShops-Hold-Test-1.0.0-experimental.13-Windows-x64-Setup.exe`.

SHA256 `ada07a7904c258aca4daa0b8a112e2af51059d89e094dfda84b3e05d5a5c0178`.

Mesmo appId `br.com.4printshops.holdtest`, produto, executável e atalho **4PrintShops Hold Test**, perfil `%APPDATA%/4PrintShops Hold Test` e registro NSIS `90ff2ff0-f013-571c-b109-7edb265dacc8`. Diretório local validado em `.desktop-test/hold-lab/installed-validation/installed/`. A instalação principal permanece separada. Os instaladores anteriores foram preservados.

Comparação lado a lado em `.desktop-test/ux13/comparacao.html`, com imagens de Início, PDF, Konica A4 em A3 a 120%, Epson e tamanhos de janela. Capturas completas em `before/` e `after/`. A galeria e os binários permanecem fora do Git.

## Roteiro para os funcionários

1. Exportar trabalhos pendentes e fechar a Hold Test. Executar o Setup .13 para atualizar a identidade Hold Test existente. Conferir versão, diretório, atalhos e formatos personalizados anteriores.
2. Na tela inicial, usar Abrir PDF com um sintético. Conferir menus e alternância PDF/QR, miniaturas, zoom e recolhimento independente dos painéis.
3. Girar uma página e conferir o indicador de alteração pendente. Abrir Arquivo → Impressão. Conferir a mesma rotação na prévia. Navegar entre páginas e verificar que o editor mantém sua seleção anterior.
4. Selecionar Konica ou Epson e abrir Propriedades. Alterar papel/orientação/cópias somente na sessão. Confirmar e verificar o resumo. Reabrir, cancelar e conferir a preservação do estado confirmado.
5. Conferir PDF A4 sobre papel A3 a 100% e 120%, escala personalizada com vírgula, área imprimível e avisos de corte. Usar rolagem na coluna esquerda e verificar os botões do rodapé.
6. Fazer o primeiro teste de Imprimir na Microsoft Print to PDF, conferindo o nome da fila antes do clique. Reabrir o PDF virtual e conferir as duas páginas e suas orientações. O workspace editado deve continuar pendente de exportação.
7. Cancelar e reabrir a janela, exportar pelo fluxo normal e testar o aviso de saída com nova edição. Conferir a abertura no Edge pelo menu existente.
8. Repetir a inspeção em 1366 × 768, 1920 × 1080 e escala nativa do Windows em 125%, incluindo as janelas de propriedades dos drivers. Essa última conferência nativa permanece pendente.
9. A impressão física da .13 depende de decisão do operador e validação manual posterior. Os testes automatizados desta etapa usaram exclusivamente saída virtual.

## Pendências

Aprovação visual e operacional dos funcionários, escala nativa Windows 125% e regressão física da .13 nas filas da gráfica. A prévia não homologa cor/ICC, ajustes privados ou acabamento. Instalação/reinício automáticos do atualizador, implantação assistida no servidor real/piloto em dois PCs e SMB/UNC conservam suas pendências anteriores. Nenhum commit ou push foi executado nesta etapa.
