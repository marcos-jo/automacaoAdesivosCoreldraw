# Escala previsível na montagem — experimental.22

Data: 05/11/2026. Base `529a144`, branch main. As 19 alterações pendentes da .21/interface básica foram preservadas em `.desktop-test/imposition22/before-1793904268588/`, com inventário/hash, cópia dos arquivos e patch. Sem commit/push ou uso de PDF de cliente. O exemplo dimensional abaixo foi reproduzido com PDF sintético.

## Diagnóstico reproduzido antes das alterações

Para uma página de 395,82 × 494,59 mm, A3 retrato, margens de 10 mm, marcas individuais (3 mm de afastamento, 3 mm de comprimento e 0,1 mm de traço):

| Situação anterior | Peça final | Escala efetiva |
|---|---|---|
| Step & Repeat, ajuste proporcional selecionado | 395,82 × 494,59 mm | 100% |
| N-up 4-up, ajuste proporcional | 126,40 × 157,94092264 mm | 31,93370724% |

O N-up descontava a reserva das marcas da célula e aplicava corretamente a menor razão largura/altura. Com uma página de origem, somente uma posição era preenchida. Isso é ajuste proporcional esperado, não perda de escala no PDF original.

Havia três problemas de estado/apresentação:

1. O campo manual permanecia com 120, desativado, enquanto `escalaModo=ajustar` ignorava esse percentual. Não se imprimia a 120% nesse caso, mas o número visível confundia o operador.
2. Os mesmos campos eram usados pelos dois modos. Entrar em N-up trocava a grade automática por 4-up; voltar a Step & Repeat mantinha essa grade. Não existia configuração independente por modo.
3. `fatorMontagem`, criado por Ajustar montagem para caber, era único e permanecia ativo ao alternar modos ou escolher tamanho real. Num exemplo com fator 0,8, a escala N-up passava de 0,3193370724 para 0,2554696579. Era um fator adicional real, compartilhado indevidamente.

No Step & Repeat, a opção antiga Ajustar proporcionalmente não tinha uma célula independente à qual ajustar: a célula é dimensionada a partir da peça. Por isso o caminho calculava 100% quando não havia medidas finais/percentual personalizado. A matemática não foi reescrita; a nova interface restringe o ajuste à posição ao N-up e oferece ajuste da montagem inteira pelo botão existente.

O formulário Básico/Avançado anterior estava **somente no código**, conforme registrado no estado do projeto. A instalação consultada era a .21, com a apresentação anterior. As margens zero ainda não estavam aplicadas: os quatro valores da janela continuavam em 10 mm. A .22 reúne a apresentação simplificada, margens iniciais zero e as correções abaixo; não se considera a instalação antiga como validação dessa apresentação.

## Configurações independentes e retorno a tamanho real

- Cada modo guarda origem/intervalos, grade, ordem, giro, escala, medidas desejadas, quantidade aplicável, margens, espaços, marcas, sangria, centralização e seu próprio ajuste adicional. Papel/orientação da folha e zoom visual são compartilhados e explicitamente mostrados.
- Step & Repeat começa em **100% / tamanho real**, grade automática. N-up começa em **Ajustar proporcionalmente à posição**, 4-up. A grade 4-up é um padrão, não um otimizador novo. N-up não repete a página automaticamente para preencher posições vazias.
- Escala manual aparece somente no modo **Escala personalizada**. O valor anteriormente informado é conservado para reutilização, mas fica oculto/desativado quando não participa do cálculo. A área básica informa que a escala manual não está aplicada no ajuste automático.
- **Definir tamanho final da peça** é uma escolha explícita em Step & Repeat. As medidas prevalecem sobre percentual; essa precedência fica escrita no resumo. Não se multiplica escondidamente medida desejada por percentual manual.
- **Voltar a 100% / tamanho real** remove o fator adicional, limpa medidas finais que poderiam substituir 100% e seleciona escala real. A peça pode então ultrapassar a folha: o aviso de corte continua obrigatório, sem reduzir silenciosamente o conteúdo.
- Ajustar para caber continua usando a geometria/reserva existentes. Guarda a grade anterior à aplicação. **Remover ajuste de escala** retira o fator e restaura essa grade. Alterar percentual, método de escala, medidas ou grade também elimina o ajuste anterior. Mudanças de outros parâmetros mantêm o fator explícito no resumo; o operador pode recalcular ou remover.
- Alternar modos restaura o ajuste próprio do destino, sem multiplicá-lo pelo ajuste do outro modo. Repetir o ajuste é idempotente dentro da precisão numérica testada. Capturar novamente o workspace ao reabrir a janela limpa ajustes calculados para o snapshot anterior, mantendo os demais valores configurados.

## Resumo confiável e confirmação de saída

A área básica, sempre acessível com a seção avançada recolhida, informa:

- Origem física da página (MediaBox, considerando giro/UserUnit) e corte utilizado pelo motor (TrimBox ou CropBox).
- Dimensões finais das peças e escala efetiva calculadas a partir dos posicionamentos do plano gerado.
- Escala manual aplicável, encaixe automático e fator adicional de Ajustar para caber, separados.
- Grade/capacidade e folhas. Páginas com tamanhos diferentes recebem indicação de variação e **Medidas e escala por página**.

Antes de gerar ou imprimir, uma confirmação usa o mesmo objeto de resultado que fornece os bytes da saída. Informa folha, grade, quantidade real de peças/folhas, medidas, escala e avisos de corte. Cancelar preserva resultado/origem. Alterar parâmetros invalida o resultado anterior e o aceite de excesso, como na .21. Não há geração automática ou impressão física durante os testes.

Medidas são formatadas com até duas casas decimais; percentuais com até quatro, de forma uniforme no resumo, prévia e confirmação. A precisão interna das matrizes permanece integral. O detalhe textual lista até 100 combinações distintas de página/tamanho/escala e avisa se houver mais; as faixas de escala do resumo consideram todas. Conferir também as demais folhas na prévia em documentos extensos.

O zoom é identificado como **Zoom da prévia**, relativo ao enquadramento visual. Não altera os bytes nem dimensões de produção. O driver é uma etapa posterior: a escala efetiva apresentada é a do PDF montado, sem incorporar zoom do driver. Para imprimir o PDF já imposto, conferir papel físico e normalmente escala 100%, sem impor N-up/livreto outra vez no driver.

## Precisão preservada e separação de responsabilidades

`pdf-imposition-layout.js`, `pdf-imposition-bounds.js`, `pdf-imposition-preview.js`, `pdf-imposition-view.js`, `pdf-imposition-compatibility.js`, `pdf-imposition-compare.js` e `pdf-workspace.js` foram comparados byte a byte à .21 instalada: sem diferenças. O gerador recebe somente metadados adicionais de MediaBox/corte para o resumo; matrizes, inserção vetorial, marcas, papel de saída e conteúdo gráfico não foram reescritos. O padrão legado de margens da API pura permanece; a janela passa valores explícitos de 0 mm.

Continuam as proteções de ActualText essencial/formulários/aparências, preparação validada de cópia, cancelamento e worker isolado. O documento do workspace e seu estado pendente não são alterados pela troca de modo, prévia, salvamento ou impressão da montagem. Nenhuma configuração global, fila física, associação, motor Windows ou atualizador foi modificado.

Novos módulos de produção: `public/js/tools/pdf-imposition-settings.js` (opções por modo) e `pdf-imposition-summary.js` (medidas/escalas a partir do resultado). Integração em `pdf-imposition-ui.js`, metadados em `pdf-imposition-engine.js` e estilo em `public/css/theme.css`. Build Hold Test mantém identidade/perfil e passa à .22.

Novos testes: `tests/pdf-imposition-scale.test.mjs`, `tests/helpers/pdf-imposition-scale-browser.mjs` e `tests/helpers/pdf-imposition-proof.mjs`. Adaptação dos testes de interface, prévia e confirmação de salvamento. No auxiliar `tests/helpers/windows-save-dialog.ps1`, foi acrescentada espera limitada pelos controles do diálogo, sempre revalidando PID/título e destino, sem teclado global. A primeira tentativa instalada parou por controles não reconhecidos no diálogo virtual; a nova execução passou. Isso foi uma falha de automação observada, sem atribuí-la a um defeito do motor de impressão ou afirmar uma causa interna do Windows não comprovada.

## Testes e entrega

Resultados de referência, sem arredondamento nos cálculos:

| Configuração sintética de referência | Peça final (mm) | Efetiva |
|---|---|---|
| Step & Repeat 100% | 395,82 × 494,59 | 100% |
| Step & Repeat manual 120% | 474,984 × 593,508 | 120% |
| N-up 4-up, encaixe, margens 10 mm e marcas individuais | 126,4 × 157,94092264 | 31,93370724% |
| N-up 4-up, encaixe, margens zero e mesmas marcas | 136,4 × 170,43624880 | 34,46010813% |

Alterar a margem padrão muda o espaço disponível e, portanto, o percentual de encaixe esperado. Não representa uma mudança da proporção ou do motor.

- **320 testes Node aprovados**, incluindo quatro novos testes de configurações independentes, caso relatado, escalas/matrizes com páginas diferentes e distinção MediaBox/TrimBox. Os testes novos verificam 63 posicionamentos realmente serializados; origem byte a byte intacta.
- **Chrome headless:** 26 estados capturados com uma e quatro páginas diferentes, 44 posicionamentos conferidos pela Matrix/BBox dos Forms no PDF. Inclui 100%/120%, encaixe, ajuste repetido, remover ajuste, grade automática/4-up, margens zero/10 mm, retorno de modo e tamanho final informado. Percentuais do resumo básico e da prévia devem coincidir literalmente. Não há dependência apenas dos textos da interface.
- Confirmações antes de gerar/imprimir, recusa preservando o resultado, ações com os bytes físicos corretos e original pendente preservado. Modos de marcas, excesso/corte, recuperação após falha, nove prévias/oito cancelamentos e preparação sintética com ActualText continuam na regressão completa.
- Electron smoke aprovado: isolamento, PDF.js/miniaturas, edição, abas, salvar/cancelar/falhar, persistência e proteção de fechamento.
- **Pacote final .22 realmente instalado:** roteiro completo aprovado, novamente com 26 estados/44 posicionamentos e 16 confirmações no ensaio de escalas. Inclui apresentação básica/margens zero, restauração dos modos, percentuais idênticos entre resumos, salvar nativo, preparação segura de cópia e duas impressões exclusivamente virtuais. A primeira tentativa instalada havia parado no reconhecimento do diálogo virtual; os rechecks passaram com a espera limitada do auxiliar, sem mudar o motor.
- ASAR final instalado auditado: **450 entradas**, arquivos de produção iguais ao código, sem testes, documentos, credenciais ou artefatos indevidos. Sintaxe de 25 scripts JS e do auxiliar PowerShell aprovadas; `git diff --check` aprovado. Logs finais em `.desktop-test/imposition22-{node-final,chrome-final,electron,build-final,installed-final}.log`; provas/sintéticos/capturas em `.desktop-test/imposition22/{browser-final,installed-final}/`, todos ignorados.

Instalador final:

`dist-desktop/hold-native-test/4PrintShops-Hold-Test-1.0.0-experimental.22-Windows-x64-Setup.exe`

SHA256: `2BB8E7B2CFD55178E1B63502862EA038D30F136140F06F2E26D8D8C72F8A5861`.

Mesma identidade `br.com.4printshops.holdtest`, diretório e perfil `%APPDATA%\4PrintShops Hold Test`, atualizando a Hold Test .21. Não substitui a identidade principal nem instala certificado. Sem publicação. O instalador .21 anterior não continha esta nova apresentação.

O ensaio virtual usa exclusivamente **Microsoft Print to PDF**. Nessa estação o papel consultado nessa fila estava em A4: os dois arquivos virtuais medidos tinham aproximadamente 210,016 × 297,011 mm. O arquivo salvo diretamente da montagem foi confirmado com 297 × 420 mm e três folhas. O teste virtual comprova encaminhamento e conclusão do fluxo, **não homologa o encaixe físico de A3 nessa fila configurada em A4**. O tamanho final depende do contexto/papel/escala escolhidos nas propriedades do driver; essa etapa não força A3 ou muda preferências globais. A prévia Windows é a referência para essa etapa posterior.

Após os ensaios concluídos, a leitura do Windows confirmou **KONICA MINOLTA C6100PS** como padrão e nenhum trabalho pendente na fila virtual. Não houve envio à Konica ou a outra fila física. Não se homologa cor, PDF/X, ICC ou comportamento de todo RIP por estes testes.

## Roteiro manual para a gráfica

1. Exportar qualquer trabalho pendente, fechar a Hold Test e aplicar o instalador .22. Conferir a versão. Deve atualizar a instalação existente, sem outra identidade ou perfil.
2. Abrir um PDF de teste, acessar Editar → Montagem / Imposição e conferir a área básica, seção avançada recolhida e margens iniciais zero. Ler origem, corte, peça final e escala efetiva.
3. Em Step & Repeat, conferir 100%. Nas opções avançadas, escolher Escala personalizada e 120%; as dimensões devem aumentar exatamente 20%. Escolher Definir tamanho final e informar uma largura; conferir que esse tamanho, sem um percentual oculto, determina o resultado.
4. Mudar para N-up: conferir a grade padrão 4-up, quantidade de repetição oculta, encaixe calculado e ausência de campo manual ativo. Uma página preenche uma posição. Com quatro origens, conferir as quatro peças e suas medidas individuais.
5. Voltar ao Step & Repeat: escala, grade, medidas e origem próprias devem retornar. Ajustar para caber no N-up, alternar várias vezes e verificar que esse fator não altera o outro modo. Remover ajuste deve restaurar escala-base/grade anterior. Voltar a 100% deve retirar fator e medidas finais substitutivas.
6. Testar margens zero/10 mm, marcas, giro e espaços; conferir dimensão final e avisos de excesso. Cancelar um aceite de corte e depois aceitá-lo conscientemente, sem alterar tamanho/folha silenciosamente.
7. Gerar o PDF, ler o resumo de conferência e reabri-lo para medir folha e peças. Comparar com a prévia. O original e suas alterações pendentes devem permanecer intactos.
8. Para imprimir, selecionar o papel físico correto nas propriedades nativas e conferir a prévia Windows. Usar escala 100% para reproduzir o PDF já imposto, salvo decisão explícita de ampliação posterior; evitar N-up/livreto adicional no driver. Começar pela fila virtual. Qualquer teste físico na Konica depende da ação consciente do operador e da configuração Hold já aprovada; não foi executado nesta etapa.

Pendentes: aceitação manual desta apresentação pelos funcionários e conferência física das novas escolhas de escala/papel. O cenário físico A3 → A6 → Step & Repeat aprovado anteriormente permanece preservado, sem homologação automática dos demais. Atualizador automático/assinatura, piloto assistido e SMB mantêm as pendências anteriores.

Git ao concluir: 27 caminhos pendentes (15 modificados/12 novos), incluindo os 19 anteriores, staging vazio, base `529a144`. Os 2.212 arquivos históricos de node_modules rastreados não foram removidos ou alterados. Instaladores, backups, PDFs sintéticos, perfis, capturas e logs continuam ignorados. Nenhum PDF de cliente, certificado ou segredo foi utilizado/adicionado. Sem commit/push.
