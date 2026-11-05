# Barra compacta e temas — experimental.15

## Aprovação operacional

Em 05/11/2026, o usuário informou que os funcionários que utilizam o 4PrintShops aprovaram a atualização visual. O checkpoint da experimental.15 foi preparado para registro mediante aprovação explícita do commit. As limitações técnicas abaixo permanecem documentadas.

## Estado em 05/11/2026

Base `16e9187`, branch `main`, árvore inicialmente limpa. A .14 instalada na gráfica é o checkpoint anterior. Esta etapa altera apresentação, organização dos controles e preferência local de tema.

O título Gerenciador de PDFs e seu subtítulo foram retirados da área superior. Abrir arquivos, Informações e Exportar / Salvar cópia agora compartilham a barra de navegação/zoom. Abrir arquivos fica visível nessa barra somente enquanto não há documento na instância PDF. O menu Arquivo mantém a abertura disponível com documento já aberto. Fechar a aba PDF e reabri-la restaura o estado vazio. Combinar arquivos continua acessível na barra de ferramentas.

O resumo técnico permanece compacto. Mensagens operacionais e erros continuam em `#pdf-status`, em indicação discreta no rodapé da área de trabalho, sem reservar outra faixa no topo. A proteção de alterações e seus avisos continuam no fluxo existente.

## Medição e aparência

Em Chrome a 1366 × 768, com o mesmo PDF sintético e painéis abertos, a altura de `#pdf-visualizacao` passou de **476 para 559 pixels**. Ganho final de **83 pixels**, aproximadamente 17%. A primeira medição intermediária foi 561 pixels. O ajuste final de bordas dos controles resultou em 559 pixels.

Capturas anteriores em `.desktop-test/theme15/before/` e posteriores em `after/`. Incluem `vazio-claro.png`, `pdf-claro.png`, `pdf-escuro.png` e `inicio-escuro-persistido.png`. A janela de impressão escura e os tamanhos de tela foram registrados em `installed/after/`.

## Tema

Exibir → Tema oferece Claro e Escuro, com indicação da escolha atual. Claro é o padrão sem preferência válida. A opção é aplicada imediatamente e salva na chave `4printshops.theme` do localStorage da origem atual, dentro do perfil existente do aplicativo. O navegador utiliza seu próprio perfil. Falhas de armazenamento mantêm o tema utilizável na sessão.

`public/css/theme.css` define a paleta por variáveis. As cores de interface em `style.css` e `print-config.css` passaram a consultar essa paleta. Cores da página PDF, canvases e folha da simulação de impressão permanecem independentes da paleta. Não há filtro/inversão sobre documentos. Uma amostra RGBA de página branca foi conferida como `[255,255,255,255]` também no tema escuro. As propriedades nativas dos drivers continuam sob controle do Windows.

Novos `public/js/app/theme.js`, `public/css/theme.css`, `tests/app-theme.test.mjs`, `tests/theme-browser.mjs` e `tests/desktop-theme-installed.mjs`. Integração em main/menu/index, estrutura PDF e ocultação do botão de importação. Builder e protocolo Electron incluem exatamente o novo módulo. Motores de edição/PDFWorkspace, C#, PDFium, DEVMODE, HDC, Konica RAW e atualizadores não foram alterados nesta etapa.

## Verificações

- 260 testes Node aprovados, incluindo preferência padrão, persistência, mudança imediata, valor inválido e falha do armazenamento. Electron smoke aprovou PDF.js/miniaturas, giro, abas, exportação nativa, cancelamento/falha, proteção de fechamento e preferências.
- Chrome específico aprovou a barra vazia/com documento, Informações, acesso à abertura pelo menu, mudança claro/escuro, persistência ao recarregar e página branca. A regressão Chrome de abas/menu aprovou documentos de 100/350 páginas, fechamento/reabertura, QR, exportação e proteção.
- O submenu Tema revelou foco de teclado em botões recolhidos na primeira regressão. A navegação foi corrigida para considerar somente botões do nível atual do menu. A regressão passou depois da correção.
- O teste instalado inicial aprovou layouts em 1366 × 768, 1920 × 1080 e 1092 × 614 com densidade emulada de 125%, com Imprimir visível. Propriedades Konica/Epson, cancelamento, prévia navegável, giro e pendências foram exercitados sem enviar trabalhos. Logs finais e o teste de reinício ficam em `.desktop-test/theme15/`.
- O pacote final .15 foi reaplicado no mesmo diretório e perfil. A abertura pelo menu com documento editado mostrou a proteção e o cancelamento preservou as pendências. Após encerrar essa instância, outro processo restaurou o tema escuro. `installed-approved.log` e `restart.log` registram as aprovações. Uma tentativa intermediária teve timeout de renderização em segundo plano; a automação passou a manter seu renderer ativo e aguardar cada navegação antes de trocar a fila. Nenhum envio foi realizado.
- Sintaxe, pacote instalado com 437 entradas e `git diff --check` conferidos. Instaladores, capturas e PDFs sintéticos permanecem ignorados. Nenhum commit ou push nesta etapa.

## Distribuição e roteiro

Instalador `dist-desktop/hold-native-test/4PrintShops-Hold-Test-1.0.0-experimental.15-Windows-x64-Setup.exe`.

SHA256 `74b74561033240469887ff27ac10c9685e3aa6f06753e8d1c0f303f2f803d1a8`.

Mesma identidade `br.com.4printshops.holdtest`, registro NSIS e perfil Hold Test. O teste identifica a pasta instalada pelo registro, incluindo a subpasta Hold Test existente nesta estação. A versão principal da raiz continua `.6`. Assinatura comercial e instalação automática do atualizador conservam suas pendências anteriores.

1. Fechar a Hold Test após exportar trabalhos pendentes e executar o Setup .15. Conferir a mesma instalação e preferências.
2. Abrir PDFs sem documento e conferir Abrir arquivos junto dos controles de zoom. Importar PDF ou lote misto, verificando que esse botão sai da barra e permanece no menu Arquivo.
3. Usar Informações, Exportar, zoom, giro, miniaturas e os painéis recolhíveis. Testar uma nova abertura pelo menu com documento editado e cancelar o descarte.
4. Alternar Exibir → Tema → Claro/Escuro. Conferir Início, QR, abas, ferramentas, combinação, redimensionamento e exportação. As páginas devem conservar as cores originais.
5. Abrir Impressão no tema escuro, consultar propriedades e cancelar. Conferir folha branca, conteúdo, área imprimível e botões. Esta etapa automatizada não envia trabalhos.
6. Fechar e reabrir o aplicativo, conferindo a preferência. Repetir em 1366 × 768 e 1920 × 1080. A conferência com escala nativa Windows 125%, incluindo janelas do driver, continua manual. A validação automatizada de densidade foi emulada no Electron.

Os testes instalados dependem da instalação de laboratório e do marcador de preferência existente. `PRINT_UI_INSTALL_DIR` pode indicar o diretório apropriado, que precisa corresponder ao registro. A avaliação de contraste/ergonomia pelos funcionários e a ampliação de opções de impressoras continuam em uso real. As pendências de perfis de cor na combinação, Explorer, SMB/UNC e distribuição interna permanecem documentadas nos guias anteriores.
