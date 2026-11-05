# Exportação de imagens e atalhos — experimental.16

Data: 05/11/2026. Base `6b8f63a`, experimental.15 aprovada pelos funcionários. Alterações desta etapa ainda sem commit.

## Operação

- Ctrl+O: Abrir arquivos pelo fluxo existente, com confirmação para preservar documento alterado.
- Ctrl+S: abre Salvar como PDF; grava somente após confirmação, preservando o original.
- Ctrl+E: abre Exportar, com PDF, PNG ou JPEG.
- Ctrl+P: abre a configuração de impressão do próprio aplicativo. Não envia trabalho automaticamente. No navegador informa a necessidade do Desktop, sem abrir impressão do Chromium.

Os atalhos são tratados por um único listener no renderer, com remoção disponível; o menu nativo continua desativado. Repetição de tecla, edição de campos, diálogo aberto e operações incompatíveis impedem novas ações. Os menus mostram os atalhos. Não há nova barra permanente.

Em Exportar, escolher todas as páginas, página atual, seleção nas miniaturas ou intervalo (exemplo `1,3-5`). PNG é sem perdas; JPEG oferece qualidade de 10 a 100%, inicialmente 92%. Resoluções reais: 150, 300 ou 600 DPI. O padrão ao abrir a janela continua PDF.

O destino é escolhido pelo operador. Dentro da pasta escolhida é criada uma **subpasta nova por lote**, com nome base, formato e identificador aleatório. As imagens recebem o número da página do documento (`nome_p0001.png`, `nome_p0007.jpg`). Essa organização evita misturar lotes e substituir arquivos existentes. No Desktop, criação exclusiva e publicação exclusiva também recusam conflitos surgidos durante a operação. Nenhuma substituição é oferecida nesta etapa.

## Documento, resolução e cor

Os bytes são uma cópia do PDFWorkspace atual, incluindo rotação, ordem, exclusões, união, redimensionamento e imagens já incorporadas. Exportar imagens não modifica o workspace, não rasteriza seu PDF e não limpa a indicação de alterações não exportadas. O usuário continua salvando o PDF pelo fluxo habitual.

PDF.js renderiza cada página em canvas próprio, sem capturar a interface. O tamanho é calculado pelo viewport efetivo, incluindo rotação e UserUnit, a `DPI/72`, arredondando cada lado para cima com tolerância numérica. O mesmo DPI é registrado em PNG pHYs ou JPEG JFIF, sem recodificar os pixels comprimidos. Os metadados não substituem o cálculo da resolução real.

Ambos os formatos usam **RGB de 8 bits sobre fundo branco**. PNG preserva os pixels sem compressão com perdas, mas nesta versão não mantém transparência alfa. JPEG é com perdas. Não há promessa de preservar um perfil ICC de saída, valores CMYK originais, cores especiais, sobreimpressão ou equivalência com Acrobat/RIP. Conferência de cor na gráfica permanece necessária. Fontes/imagens/vetores são interpretados pelo PDF.js para produzir a imagem; o PDF original conserva sua representação.

## Limites e cancelamento

Uma página por vez; máximo 12.000 pixels por lado e 36 megapixels por página. Pico estimado de buffers: 16 bytes por pixel, sem contar todo o custo do PDF.js/documento. A4 em 600 DPI cabe; A3 em 600 DPI ultrapassa o limite e exige DPI menor. Não há redução silenciosa da resolução.

Máximo 1.000 páginas e 1,5 bilhão de pixels por lote. Desktop limita cada imagem a 192 MiB e a saída do lote a 2 GiB. Documentos complexos ainda podem consumir recursos elevados; uma renderização tem timeout de 60 segundos. Progresso informa página, posição no lote, pixels e estimativa de memória. Cancelar aborta as tarefas e limpa arquivos próprios; mantém lotes anteriores e arquivos alheios. Controles permanecem bloqueados até finalizar ou cancelar a operação.

No Desktop, os arquivos são temporários `.part` até a publicação final do lote. Cancelamento, erros e encerramento do renderer têm caminhos de limpeza. Foram testados cancelamento normal e gravação abortada; não foi feito ensaio de queda de energia. Encerramento forçado do processo/Windows pode deixar pasta ou `.part` a remover manualmente após conferir que não há exportação ativa.

No navegador, o destino depende da File System Access API de Chrome/Edge em origem permitida, com autorização da pasta. Não há fallback silencioso para centenas de downloads. A API do navegador não oferece a mesma criação exclusiva atômica do sistema de arquivos Node; subpasta UUID e verificações prévias reduzem conflitos. Não usar a mesma subpasta simultaneamente em outro programa. O seletor real de pasta foi validado no Desktop; o navegador foi testado com destino em memória e handles simulados para conflitos/limpeza, sem automatizar sua permissão de diretório.

## Validação realizada

- `npm test`: **271 testes Node aprovados**, incluindo os 260 anteriores e 11 novos. Plano/nomes, modos de página, resolução/memória, metadados, conflitos, gravação/cancelamento e os quatro atalhos cobertos. Uma execução anterior sob compilação concorrente excedeu timeout de dois testes HTTP do atualizador assistido; a execução final, sem compilação concorrente, passou integralmente sem alterar o atualizador.
- Chrome real: PNG 150/300/600, JPEG 40%/95%, documento sintético misto de sete páginas e giro com alterações pendentes, decodificação dos arquivos e dimensões verificadas. Lote de 350 páginas sintéticas pequenas concluído sequencialmente; não é ensaio de 350 páginas A3 complexas em 600 DPI.
- A4 girado: 1.754 × 1.241 px em 150 DPI, 3.508 × 2.481 em 300, 7.016 × 4.961 em 600. Amostras de JPEG confirmaram tamanho menor com qualidade 40% que 95%.
- Imagens de exemplo inspecionadas visualmente; não se afirma igualdade pixel a pixel com Acrobat ou com o RIP.
- Regressão Chrome de abas/menus: aprovada, incluindo documentos de 100/350 páginas e proteção de alterações. Electron smoke: aprovado, incluindo abrir/girar/salvar, cancelamento/falha nativa, proteção e persistência.
- Aplicativo realmente instalado .16: Ctrl+S/E/P/O, pasta nativa do Windows, sete PNGs em 150 DPI, dimensões da primeira página, cancelamento em 600 DPI sem resíduos e documento ainda pendente aprovados. Ctrl+O respeitou confirmação de descarte e cancelamento preservou o documento.
- Atualização da instalação Hold Test existente e reaplicação do pacote final no mesmo executável/diretório; versão .16 confirmada. Mesmo appId, NSIS e perfil `%APPDATA%/4PrintShops Hold Test`; não foi criada identidade paralela.
- ASAR instalado: **440 entradas**, allowlist, igualdade dos fontes e exclusões aprovadas. Sintaxe dos JS/MJS/CJS alterados e `git diff --check` aprovados.
- Nenhum trabalho físico enviado. Fontes C#, PDFium/HDC, DEVMODE e Konica RAW não foram alterados. Sem commit, push ou publicação.

Registros locais ignorados em `.desktop-test/images16/`: `node-complete.log`, `browser-release.log`, `electron-release.log`, `regression.log`, `installed-delivery.log`, `package-complete.log`. Capturas claras/escuras e exemplos em `browser/` e `installed/`; imagens realmente gravadas pelo Windows constam em `installed/native-output.json`.

## Instalador

`dist-desktop/hold-native-test/4PrintShops-Hold-Test-1.0.0-experimental.16-Windows-x64-Setup.exe`

SHA256: `02F2208583F92566222CA855328CB9DBE20C55F0B66FC2BC5B4E7EDB14A8F85D`.

Gerado com `electron-builder.hold-native-test.cjs`, identidade `br.com.4printshops.holdtest`. Sem assinatura Authenticode comercial; atualização automática com instalação/reinício permanece não homologada. Instalador e amostras são ignorados no Git.

## Arquivos desta etapa

Novos: `desktop/image-exports.mjs`, `public/js/app/shortcuts.js`, `public/js/tools/pdf-image-export.js`, `tests/image-export.test.mjs`, `tests/image-export-browser.mjs` e este guia.

Alterados: `desktop/app.mjs`, `desktop/preload.cjs`, `desktop/security.mjs`, `electron-builder.config.cjs`, `electron-builder.hold-native-test.cjs`, `public/css/theme.css`, `public/js/app/export-dialog.js`, `public/js/app/menus.js`, `public/js/main.js`, `public/js/tools/pdfs.js`, `tests/pdf-integration.test.mjs`, `docs/ESTADO_PROJETO.md` e `docs/DISTRIBUICAO_DESKTOP.md`.

## Aceitação manual na gráfica

1. Exportar qualquer PDF pendente, fechar Hold Test e instalar .16 no destino existente. Conferir preferências e temas.
2. Ctrl+O: abrir um PDF de teste; girar, reorganizar, redimensionar ou combinar. Conferir a página exibida.
3. Ctrl+E: PNG, página atual, 150/300/600 DPI; escolher pasta e conferir a nova subpasta e pixels no editor de imagens. Comparar proporção, giro, textos, transparências e cores.
4. Exportar seleção/intervalo e todas as páginas; testar JPEG 92% e qualidade menor. Conferir a sequência e os números. Repetir o lote sem substituir o anterior.
5. Cancelar uma exportação e conferir que controles voltam a responder e que não restam arquivos daquele lote. O indicador de alterações do PDF deve permanecer.
6. Ctrl+S: salvar PDF, reabrir e conferir as edições. Ctrl+P: apenas abrir e cancelar a janela; não é necessário imprimir para validar atalhos.
7. Conferir claro/escuro e campos em edição (atalhos não devem iniciar operações). Relatar diferenças de renderização, consumo de memória ou cor para ajuste posterior.

Pendentes: aceitação dos operadores; fidelidade de cor/ICC na gráfica; ensaios de PDFs complexos próximos dos limites. Permanecem os limites anteriores de SMB/UNC e homologação de instalação automática do atualizador.
