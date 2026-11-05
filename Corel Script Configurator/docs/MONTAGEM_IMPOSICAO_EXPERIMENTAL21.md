# Montagem com excesso geométrico — experimental.21

Data: 05/11/2026. Base limpa `529a144` (checkpoint .16–.20 concluído). Sem commit, push, publicação ou impressão física nesta etapa. Não houve uso de PDF de cliente nestes novos testes.

## Diagnóstico e mudança controlada

O planejador recusava a grade antes de gerar o PDF quando ela ultrapassava a área útil; N-up a 100% também recusava páginas maiores que a célula. A janela descartava o resultado e zerava o canvas. Por isso o operador recebia o erro sem enxergar a posição das peças.

O comportamento estrito continua sendo o padrão da API. A janela solicita explicitamente um plano de inspeção que admite excesso geométrico. Em Step & Repeat, medidas/escala/transformações existentes são mantidas. Em N-up a 100% ou percentual, células podem crescer para comportar as peças sem sobreposição ou redução silenciosa. A grade continua centralizada, salvo opção contrária do operador. Grades automáticas que não comportem sequer uma peça mostram ao menos uma posição para inspeção.

Não se retiraram validações de números, proporções, caixas de página, sangria disponível, limites de lote ou compatibilidade PDF. O cenário físico A3 → A6 → Step & Repeat aprovado anteriormente permanece uma regressão explícita. Uma comparação independente com o código do checkpoint `529a144` encontrou **192 planos válidos idênticos**, incluindo os quatro modos de marcas, giros, ordens e escalas.

## Prévia e autorização

- A prévia tem área de trabalho maior que a folha, com as mesmas matrizes, peças e marcas do PDF de produção. Uma cópia interna muda somente MediaBox/CropBox para incluir o excesso. **Nunca é essa cópia ampliada que se salva ou imprime.** O PDF de produção permanece no tamanho físico escolhido.
- Contorno indica o papel; ocre tracejado indica margens; azul indica peças. Fora do papel há hachura vermelha, inclusive quando as marcas ultrapassam os limites. A prévia vazia de referência conserva papel/margens durante atualização ou erro, sem reapresentar arte antiga como resultado atual. Medidas de papel inválidas precisam ser corrigidas antes de desenhar uma folha válida.
- **Ver tudo**, **Folha inteira** e −/+ alteram somente o enquadramento, com rolagem para detalhes. Avisos/resumo entram no layout antes de calcular o encaixe do canvas. Cancelamento, worker isolado, fila, prazos, versão e **Tentar novamente** da .18 continuam ativos.
- O aviso identifica bordas esquerda/direita/superior/inferior, separando excesso físico, invasão das margens e eventual área imprimível consultada. Duas A4 a 100% numa A3 paisagem podem caber no papel e ainda invadir margens de 10 mm: isso não significa corte físico inevitável.
- **Continuar com corte nas bordas** exige confirmação e libera gerar/imprimir apenas o resultado atual. Qualquer regeneração de geometria, mudança de fonte/preparação ou recuperação que produza novo resultado revoga o aceite. Zoom, guias e navegação não mudam a geometria nem revogam o aceite. Recusar preserva os dados e mantém as saídas bloqueadas.
- O recorte usa os limites físicos do PDF, sem mudar papel ou escala. Marcas externas não serão incluídas na área impressa. É um recorte por caixas de página, não uma operação de expurgo/redação de conteúdo sensível.

## Aproveitamento e ajuste para caber

Resumo com folha, área útil, medidas finais das peças, escala efetiva, peças por folha, cm² das peças dentro do papel, percentual ocupado e área restante. São áreas dos **retângulos de corte**, não uma medição de cobertura de tinta. A área restante pode conter reservas de sangria/marcas. As medidas internas não são arredondadas; somente os rótulos recebem formatação.

**Ajustar montagem para caber** calcula o maior fator proporcional para a grade atual, descontando margens, espaços fixos, reservas de marcas (inclusive metade do traço), sangria e giro. Mostra o percentual final antes de aplicar; para páginas diferentes pode mostrar um intervalo de escalas. O aceite fixa linhas/colunas da grade atual e aplica um fator adicional explícito. Papel e proporções permanecem iguais. **Remover ajuste de escala** restaura as medidas configuradas. Alterar outros campos mantém o fator adicional sinalizado, evitando uma troca silenciosa da escala; para recalcular, use novamente o ajuste.

Se espaços/marcas já consumirem toda a área ou a redução deixar a sangria abaixo da solicitada, o ajuste informa o impedimento e não muda a configuração. Não inventa sangria. O cálculo considera papel/margens da montagem, não substitui as restrições do driver: mesmo após caber na folha, a impressora pode exigir outro ajuste ou aviso.

## Área imprimível e compatibilidade

Quando a janela de impressão Windows já tiver fornecido geometria pública, a montagem pode mostrar a **última consulta**, somente se as dimensões físicas corresponderem ao papel selecionado (tolerância de correspondência de 0,1 mm). A origem da área é identificada; não se envia DEVMODE/sessão ao motor geométrico. Sem consulta correspondente, inclusive no navegador, mostra **não verificada**. Não há tamanho fictício ou margem universal de impressora. A impressão mantém a sessão real e suas verificações existentes.

O diagnóstico geométrico usa tolerância de 0,001 mm para eliminar resíduos numéricos de medidas PDF; não arredonda posições/escala. A área de trabalho fica limitada a 10.000 mm por lado e a renderização a 36 milhões de pixels/12.000 px por lado. Ampliação excessiva recebe erro recuperável; reduzir zoom ou grade resolve o limite. Geração/serialização PDF-lib ainda têm trechos síncronos não preemptivos; a cópia de prévia exige memória adicional.

**Preparar cópia para montagem** da .20 permanece independente do aceite de corte. ActualText essencial, formulários, aparências não comprovadas e demais incompatibilidades continuam bloqueados. Os testes com cliente real da .20 não foram repetidos nesta etapa geométrica; sua aprovação não é uma homologação universal. Perfis ICC, CMYK/cores especiais, sobreimpressão, PDF/X e fidelidade de RIP continuam sujeitos às limitações anteriores. Nenhum motor Windows, configuração global ou fila de impressora foi modificado.

## Arquivos

- Novos `public/js/tools/pdf-imposition-bounds.js` (excesso, áreas e ajuste) e `pdf-imposition-view.js` (sobreposições e folha de referência).
- `pdf-imposition-layout.js`, `pdf-imposition-engine.js`, `pdf-imposition-preview.js`, `pdf-imposition-ui.js`: inspeção optativa, cópia de prévia, enquadramento, consentimento e integração.
- `public/js/main.js`, `public/js/tools/pdfs.js`: repasse somente da última geometria pública consultada; nenhum IPC novo ou alteração de DEVMODE.
- `public/css/theme.css`: rolagem/legendas/avisos compactos nos dois temas.
- `tests/pdf-imposition-overflow.test.mjs`, `tests/helpers/pdf-imposition-overflow-browser.mjs` novos; adaptação do carregador Node, teste de prévia e roteiro Chrome/instalado.
- `electron-builder.hold-native-test.cjs`: somente versão .21. Estado do projeto e guia de distribuição atualizados.

## Validação executada

- **316 testes Node aprovados**, incluindo oito novos casos geométricos/saída e um novo teste de consentimento/invalidação.
- Duas A4 em A3 paisagem sem margens: encaixe exato a 100%, sem sobreposição/corte. Com margens de 10 mm: aviso de invasão de margem, separado de excesso físico; ajuste máximo sem distorção.
- Marcas individuais/grupo/divisões, Step & Repeat a 120%, giros 0/90/270°, espaçamento zero, grades excedentes e sangria insuficiente. PDF de produção físico e prévia expandida com streams de conteúdo idênticos e Forms/fontes vetoriais conservados.
- **Chrome headless e .21 realmente instalada:** recusa/aceite do corte, mudança invalidando aceite, zoom preservando aceite, escala apresentada antes do ajuste, área imprimível sintética, retorno dos bytes físicos ao fluxo de impressão e original/estado pendente intactos.
- Prévia expandida recortada à folha versus PDF de produção: **16.045.152 componentes RGBA, diferença máxima 0**, em ambos os ambientes. Comparação anterior à sobreposição das guias/hachuras. Não é homologação de impressão física ou cor de todo PDF.
- Estabilidade: nove prévias iniciadas/oito cancelamentos; falha injetada e Tentar novamente; saídas voltam a responder. Preparação sintética com ActualText redundante/links continua validando texto e **4.080.928 componentes**; ActualText essencial e formulários seguem bloqueados.
- Salvamento nativo e duas impressões **exclusivamente Microsoft Print to PDF**, incluindo sintético preparado, no aplicativo instalado. Nenhum trabalho à Konica/Epson/outro equipamento físico.
- Electron smoke passou: isolamento, PDF.js/miniaturas, giro, abas, salvar/cancelar/falhar, proteção de fechamento e persistência.
- Sintaxe e `git diff --check` aprovados; ASAR instalado auditado com **448 entradas**. Artefatos e logs apenas em `.desktop-test/imposition21/` e `.desktop-test/imposition21-*.log`, ignorados pelo Git.

## Instalador e conferência manual

`dist-desktop/hold-native-test/4PrintShops-Hold-Test-1.0.0-experimental.21-Windows-x64-Setup.exe`

SHA256: `C248145617124BFCF1B0591A8530F0B73AB507DF46B99B607CC3E6B08179542E`.

Aplicado sobre a .20 de laboratório, mesma identidade `br.com.4printshops.holdtest`, diretório e perfil `%APPDATA%\4PrintShops Hold Test`; não cria nova linha de instalação. Continua experimental e sem certificado Authenticode comercial. O atualizador automático com instalação/reinício continua não homologado; nada foi publicado.

1. Abrir um PDF sintético de duas páginas A4. Fazer uma edição pendente e abrir **Editar → Montagem / Imposição**.
2. Escolher N-up, páginas 1–2, 2-up, A3 paisagem, escala 100%, margens/intervalos zero e Sem marcas. Conferir duas peças 210 × 297 mm na folha 420 × 297 mm, sem aviso geométrico.
3. Definir margens de 10 mm. Conferir o aviso das margens, mantendo a escala. Recusar o corte primeiro: saídas continuam bloqueadas; depois aceitar para conferir a liberação.
4. Selecionar Contorno externo + divisões e 120%. Em Ver tudo, conferir papel, artes externas, hachura e bordas indicadas. Testar Folha inteira/zoom/rolagem. Salvar somente após o aceite e reabrir o PDF salvo: tamanho A3, corte nos limites da folha, sem guias/hachuras.
5. Alterar algum parâmetro: conferir que o aceite anterior foi revogado. Usar Ajustar montagem para caber, conferir o percentual proposto e aceitar. Verificar peças menores e marcas dentro das margens. Remover o ajuste e repetir com marcas Individuais e Apenas no grupo.
6. Repetir com Step & Repeat, giro de 90° e espaçamento zero. Alterar parâmetros rapidamente, navegar e reabrir. Conferir resposta da prévia e medidas.
7. Se desejar conferir área imprimível, consultar previamente o papel correspondente nas propriedades Windows; a montagem identifica a última consulta. Imprimir em uma fila virtual e conferir o PDF resultante, com escala 100% e sem imposição adicional no driver.
8. Cancelar a montagem: documento original e suas alterações pendentes devem permanecer intactos. Para PDFs estruturados, testar separadamente Preparar cópia; aceitar corte não contorna incompatibilidade.

Pendente: validação manual desta .21 na rotina da gráfica e conferência física de cortes, novas escalas/materiais/grades. A aprovação física anterior A3 → A6 → Step & Repeat é preservada, sem extrapolação. Não há envio físico automático.

## Refinamento posterior da interface básica e avançada

Em 05/11/2026, após os testes/pacote acima, o formulário foi reorganizado sem alterar o motor:

- Inicialmente ficam visíveis **Tipo de montagem**, **Origem**, **Tamanho da folha**, **Orientação** e **Quantidade de peças desejadas** somente em Step & Repeat. N-up oculta também os tamanhos de peça exclusivos de repetição, sem apagar seus valores.
- Selecionar **Personalizada** revela largura/altura da folha na área básica. Os formatos predefinidos mantêm essas medidas acessíveis ao código, sem campos extras na tela inicial.
- **Configurações avançadas** usa disclosure HTML nativo, inicialmente recolhido, com grupos de distribuição, escala/peças, margens/espaços, marcas/sangria e posicionamento/guias. Todos os controles anteriores continuam disponíveis. Expandir/recolher não recalcula nem invalida o PDF ou o aceite de corte.
- Resumo básico informa linhas/colunas, capacidade por folha e quantidade de folhas. Step & Repeat identifica a distribuição automática quando ativa. N-up mantém a grade existente: ao mudar da grade automática de repetição, assume 4-up como antes; outras grades permanecem nas opções avançadas. Não foi implementado um novo otimizador de N-up nem mudado o comportamento de escala.
- Prévia, aproveitamento, avisos, preparação segura, confirmação de corte, exportação e impressão Desktop permanecem nos mesmos fluxos. As ações finais continuam fora da área rolável. No navegador, a impressão direta mantém sua disponibilidade anterior, restrita ao Desktop.

Arquivos deste refinamento: `public/js/tools/pdf-imposition-ui.js`, `public/css/theme.css`, `tests/pdf-imposition-browser.mjs`, novo `tests/helpers/pdf-imposition-basic-browser.mjs`, este guia e `docs/ESTADO_PROJETO.md`. As 18 alterações pendentes da etapa .21 foram preservadas; agora há 19 caminhos pendentes no total, sem staging/commit.

Validação: **316 testes Node aprovados**; sintaxe dos três scripts deste refinamento e `git diff --check` aprovados. Chrome headless verificou cinco campos básicos na repetição/quatro no N-up, campos personalizados condicionais, todos os ajustes avançados acessíveis, acionamento por Espaço, valores mantidos na troca de modo, recolher sem mudar o resultado, original pendente preservado e regressões completas de montagem/corte/ajuste/ActualText. Capturas claro/escuro, resultados e sintéticos ficam ignorados em `.desktop-test/imposition-basic/`, com log `.desktop-test/imposition-basic-chrome.log`. Não houve impressão física; o ensaio de encaminhamento usa callback de teste, sem fila real.

Os oito módulos `pdf-imposition-{engine,layout,preview,bounds,view,compatibility,compare}.js` e `pdf-workspace.js` foram comparados byte a byte ao ASAR instalado da .21: idênticos. A validação instalada e as impressões virtuais descritas nas seções anteriores pertencem ao pacote .21 anterior a este refinamento. **Não foi gerado novo instalador nem testada instalada esta nova apresentação.** Para conferir agora, usar `node app.js`, abrir um PDF e acessar Editar → Montagem / Imposição; alternar os modos, escolher folha personalizada, expandir/recolher as opções e conferir que a prévia/resultado permanecem coerentes. Aguardar avaliação dos funcionários antes do próximo pacote. Sem commit/push ou novas ferramentas.
