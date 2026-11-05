# Montagem / Imposição — experimental.17
**Validação física posterior informada pelo usuário em 05/11/2026:** original A3 reduzido a A6, seguido de Step & Repeat e impressão real, com dimensões conferidas e corretas. Aprovação geométrica somente nesse cenário; demais tamanhos, materiais e modos não ficam automaticamente homologados. A correção do ciclo de prévia e as marcas externas da etapa seguinte estão em [MONTAGEM_IMPOSICAO_EXPERIMENTAL18.md](MONTAGEM_IMPOSICAO_EXPERIMENTAL18.md).

## Estado e preservação

Implementação de 05/11/2026 sobre `6b8f63a`, sem commit/push. Os 19 caminhos pendentes da experimental.16 foram auditados e copiados antes de qualquer alteração em `.desktop-test/checkpoint16-before-imposition-20261105-100429/`, com patch e inventário. O staging estava vazio; `git diff --check` passou. Instaladores, PDFs sintéticos e logs continuam ignorados. A .16 permanece implementada.

Conforme decisão explícita do usuário, não foi ampliada a arquitetura para múltiplos PDFs. **Editar → Montagem / Imposição** abre uma janela própria. Ela captura os bytes atuais do PDFWorkspace uma vez ao abrir. Configurar, visualizar, cancelar, salvar ou imprimir a montagem não substitui a origem nem marca suas edições como exportadas. Reabrir a janela captura novamente o workspace. Não há opção de substituição automática ou abertura da montagem no editor nesta versão; salve o arquivo e use Abrir arquivos, com a confirmação de descarte existente quando aplicável.

## Modos e medidas

Motor geométrico puro, em milímetros; conversão para PDF por `72/25,4`. A precisão interna não depende dos dois decimais do resumo.

- **Step & Repeat:** exatamente uma página de origem; grade automática ou definida, tamanho final proporcional, escala 100% ou percentual, giro adicional e quantidade total. Folhas adicionais são calculadas por teto da quantidade/capacidade. A última folha contém apenas as cópias pedidas.
- **N-up:** páginas/intervalos em sequência explícita, horizontal ou vertical, grades 2/4/6/8-up ou personalizada. `0` reserva posição vazia; não é página 1. Há ajuste proporcional, 100% e escala personalizada, giro manual e giro automático para aumentar o aproveitamento. Uma página só se repete quando o operador a informa novamente na seleção.
- Folhas A4, A3, SRA3 ou personalizadas; retrato/paisagem. Margens independentes nos quatro lados, intervalos horizontal/vertical, centralização. A grade automática usa a rotação escolhida; não é um otimizador de empacotamento misto.
- Tamanho final manual em Step & Repeat prevalece sobre o percentual de escala; uma combinação incompatível de largura/altura é recusada para evitar deformação. Se nenhuma medida final for informada, usa-se a arte na escala escolhida. A escala percentual só é usada no modo Personalizada; 100% mantém as medidas atuais.
- N-up reserva células iguais e ajusta cada página individualmente sem distorcer. Escala 100%/personalizada que não caiba é recusada, sem recorte silencioso. A montagem não altera o papel do driver.

## Corte e sangria

TrimBox explícita válida define a peça. Na ausência, usa CropBox com aviso. CropBox, TrimBox e BleedBox devem ser retângulos válidos e coerentes dentro da MediaBox; casos incompatíveis são bloqueados. Giro original, giro adicional e UserUnit são considerados.

Sangria não é criada. Só é aceita quando existe área suficiente em BleedBox e após a escala. A reserva de sangria é simétrica e conservadora: considera a menor disponibilidade nos quatro lados. CropBox que esconda a sangria solicitada exige preparação do arquivo fora desta ferramenta.

Marcas são linhas vetoriais pretas RGB, com comprimento, afastamento e espessura em mm. Não são cor de registro em todas as chapas. O afastamento deve alcançar o limite da sangria; os cálculos reservam comprimento completo e meia espessura do traço. Cada posição recebe sua própria reserva. O intervalo informado é **entre reservas**, não entre as linhas de corte; isso pode gerar espaço conservador entre peças. Marcas não podem invadir a peça ou sair da área reservada da folha. Grades incompatíveis são recusadas.

Guias azuis de corte usam as coordenadas do plano sobre a prévia e **não são exportadas**. A folha com várias peças não recebe uma única TrimBox que fingiria representar todos os cortes; cada peça é identificada pelas marcas.

## Conteúdo PDF e prévia

O motor gera PDF novo com Form XObjects reutilizados; não rasteriza páginas para montar. Fontes, streams, imagens, vetores, espaços de cor e ExtGState referenciados são copiados. A prévia é uma renderização PDF.js dos **mesmos bytes** exportados/impressos, com navegação pelas folhas. Ela não é um segundo simulador geométrico.

Sondagem da pdf-lib instalada confirmou que `embedPage` não transfere automaticamente `/Group` da página. A integração copia explicitamente o grupo e suas referências para o contexto de saída. Teste estrutural e comparação sintética de renderização validaram essa correção. Páginas vazias são normalizadas na cópia, sem alterar o workspace.

Bloqueios conservadores atuais: AcroForm (inclusive campo vazio), anotações, StructTreeRoot, camadas opcionais OCProperties, OutputIntents, ações AA/OpenAction. Não se descartam silenciosamente aparências, semântica ou gerenciamento de cor. PDFs marcados podem primeiro passar pelo modo de impressão existente, incluindo sua proteção ActualText. Documentos com recursos complexos ainda exigem homologação específica. Não se afirma conformidade PDF/X, equivalência de ICC/CMYK/cores especiais/sobreimpressão com o RIP, nem fidelidade universal de fontes. A montagem não executa conversão RGB/CMYK.

Limites: até 10.000 posições por lote/grade, 1.000 folhas e saída de 256 MiB; a verificação de bytes ocorre após serializar, não é garantia rígida do pico de memória. Renderiza uma folha de prévia por vez; mudanças invalidam resultados antigos e a fila aguarda liberação antes da próxima tarefa. Cancelar aborta a preparação e destrói o carregamento PDF.js. Não houve ensaio de estresse no teto.

## Salvar e imprimir

**Gerar PDF de montagem** usa Salvar cópia nativo no Desktop ou download explícito no navegador. Não precisa carregar a montagem no editor. O original permanece pendente quando já estava pendente.

**Imprimir montagem**, somente no Desktop, exige confirmação do aviso de transformação dupla e abre a janela Windows existente com os bytes da montagem, começando em 100% no aplicativo. O clique final de Imprimir permanece necessário. Configurações privadas DEVMODE não são reescritas. A interface não consegue provar que todo driver está com N-up/livreto/zoom desativados: o operador deve conferir as propriedades. O mecanismo de impressão, PDFium/HDC, Hold e Konica RAW não foi modificado. Nenhum teste desta etapa enviou trabalho físico.

## Verificações realizadas

- 284 testes Node aprovados: os 271 da .16 e 13 novos, incluindo cartão 90×50 em A3, A6, N-up A4/misto A3, quatro páginas distintas, posições vazias, ordens, margens, percentuais100/120/87,5, giro90, rotação automática, marcas, sangria, grades impossíveis, cancelamento, preservação de fontes/Group e alterações acumuladas no workspace (união, giro, exclusão, ordem e redimensionamento).
- Chrome real: abertura pelo menu, prévia, temas claro/escuro, botões dentro de1366×768, N-up, erro de grade, recuperação após mudanças sucessivas, cancelar e reabrir preservando origem pendente. Extração do PDF imposto confirmou12 ocorrências do texto vetorial na primeira folha do cartão.
- Comparação de renderização PDF.js da fixture com texto e duas transparências sobrepostas, montada a100% em posição alinhada: **82.944 componentes RGBA comparados, diferença máxima0; nenhum componente diferente**. É uma prova sintética delimitada, não certificação de todos os PDFs.
- Instalador realmente aplicado à Hold Test existente, mesma identidade/diretório/perfil; exportação pela janela nativa do Windows, reabertura e origem pendente aprovadas. Microsoft Print to PDF recebeu exclusivamente a montagem; o PDF virtual de uma folha foi conferido, distinto das sete páginas da origem. Impressora física não acionada. Papel/geometria específicos do driver virtual não equivalem à homologação Konica.
- Regressão Electron smoke aprovada: leitura/miniaturas/giro, salvamento/cancelamento/falha, cópia seletiva, abas, proteção, persistência e encerramento. Auditoria do pacote instalado:443 entradas. Sintaxe e `git diff --check` verificados.
- Primeira tentativa Chrome falhou por automação que acionava um item desabilitado sem abrir o menu. Roteiro corrigido e reexecutado com sucesso; não houve remoção de proteção da aplicação.

Registros e amostras locais em `.desktop-test/imposition17/`: logs Node/Chrome/instalado/Electron, `browser/fidelity.json`, `browser/Cartoes 90x50 A3.pdf`, `installed/Montagem A3.pdf`, `installed/Montagem virtual.pdf`, capturas `montagem-claro.png`, `montagem-escuro.png` e `nup-4.png`. A sondagem `incorporacao-sem-correcao-nao-producao.pdf` é somente evidência de diagnóstico, não um arquivo para produção.

## Arquivos

Novos módulos: `public/js/tools/pdf-imposition-layout.js`, `pdf-imposition-engine.js`, `pdf-imposition-ui.js`.
Novos testes: `tests/pdf-imposition.test.mjs`, `tests/pdf-imposition-browser.mjs`.
Integração compartilhada: `public/js/tools/pdfs.js`, `public/js/main.js`, `public/css/theme.css`, `electron-builder.hold-native-test.cjs`, `tests/pdf-integration.test.mjs` e documentação. As demais alterações pendentes da .16 foram preservadas; não significam novos módulos da .17.

## Aceitação manual na gráfica

1. Exportar qualquer trabalho pendente e fechar Hold Test antes de instalar a .17 no destino existente. Conferir tema e preferências.
2. Abrir cartão90×50 de teste, fazer uma alteração sem exportar e acessar Editar → Montagem / Imposição. Selecionar A3, repetição, marcas e quantidade25; conferir capacidade, três folhas no caso padrão de12 posições, última folha parcial e medidas dos cortes.
3. Alterar margens, intervalos e percentual; conferir o resultado e os avisos. Testar uma grade impossível e recuperar. Em tamanho manual, manter proporção entre largura e altura.
4. Testar PDF com quatro páginas distintas em4-up; alternar horizontal/vertical, giro automático e uma vaga0. Testar A4/A3 mistos e grade personalizada.
5. Testar TrimBox/BleedBox conhecidas e sangria real; arquivo sem sangria deve recusar pedido positivo, sem inventar bordas. Conferir no editor profissional os tamanhos e distância/traço das marcas.
6. Salvar montagem e reabri-la em leitor para conferir texto, vetores, imagens, transparências e corte. Cancelar a janela deve retornar à origem exatamente como estava e ainda pendente. Salvar PDF original é uma ação separada.
7. Primeiro imprimir apenas na Microsoft Print to PDF. Conferir papel/escala100 e desligar N-up/livreto no driver. Só enviar à Konica posteriormente por decisão pessoal, com conferência das propriedades e do resultado.

Pendentes: aprovação operacional da gráfica, PDFs complexos/bloqueados, fidelidade de cor no RIP, fontes incomuns, estresse nos limites e teste físico de montagem. Não foram implementados livretos, cadernos, lombada, frente/verso complexo ou múltiplas abas de PDF. Atualizador automático e SMB/UNC mantêm as pendências anteriores.

## Instalador de entrega

`dist-desktop/hold-native-test/4PrintShops-Hold-Test-1.0.0-experimental.17-Windows-x64-Setup.exe`

SHA256: `E5204EB5F650A851A721F7CACAD78F6ADB953105F5AC26931C448F2F34277FC1`.

Mesmo appId `br.com.4printshops.holdtest`, NSIS, executável, diretório e perfil `%APPDATA%/4PrintShops Hold Test` da .16. Instalado e conferido localmente; não substitui a identidade principal. Sem assinatura Authenticode comercial; instalação automática pelo atualizador continua não homologada. Este pacote é experimental e depende da validação operacional da montagem na gráfica.

Conferência final adicional: os operadores vetoriais do PDF de cartões foram decodificados; a primeira folha contém96 segmentos de marcas (12 peças ×8), e as coordenadas em pontos conferem com o plano em mm dentro de1e-6pt. A regressão Chrome de imagens/atalhos da .16 passou novamente. Consulta somente por leitura confirmou `KONICA MINOLTA C6100PS` como impressora padrão. Inventário comparativo `preservation16.json` confirmou que somente os arquivos compartilhados esperados foram ampliados; demais arquivos da .16 conservaram seus hashes. Staging vazio,25 caminhos pendentes no total, incluindo os19 da .16.
