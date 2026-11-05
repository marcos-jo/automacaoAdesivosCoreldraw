# Montagem / Imposição — experimental.18

Data: 05/11/2026. Continuação da .17, sem commit/push. As alterações pendentes foram copiadas antes da edição em `.desktop-test/imposition18/baseline17-20261105-111219/` (25 caminhos, inventário e patch). Nenhum trabalho anterior foi descartado. Referência funcional e limites gerais: [MONTAGEM_IMPOSICAO_EXPERIMENTAL17.md](MONTAGEM_IMPOSICAO_EXPERIMENTAL17.md).

## Aprovação física informada pelo usuário

O usuário reabriu a .17, utilizou original A3, reduziu a A6, executou Step & Repeat e conferiu medidas na impressão física. Relatou dimensões corretas da folha/montagem e das peças. **Aprovação geométrica limitada a esse cenário**; não homologa automaticamente N-up, outros papéis/materiais, cor ou recursos complexos. Nenhuma impressão física foi enviada pelo desenvolvimento da .18.

## Diagnóstico e correção da prévia

Foram identificados dois defeitos no ciclo assíncrono anterior:

1. `render()` voltava a ler `this.result.plan` e `this.page` depois de `await`. Uma alteração de parâmetros podia limpar/substituir o resultado durante essa espera. A reprodução sobre a cópia da .17 produziu `Cannot read properties of null (reading 'plan')`.
2. A fila de atualizações dependia da conclusão de `loadingTask.promise`, renderização e `destroy()`, sem uma promessa própria de interrupção nem timeout. Apenas chamar `destroy()` no evento de abort não garantia liberar a espera que segurava a fila. Em ensaio controlado com carregamento pendente, o destroy era chamado e a espera antiga continuava pendente após o cancelamento. Uma única tarefa assim impedia as posteriores, conservando os botões bloqueados.

Evidência local: `.desktop-test/imposition18/legacy-race.json`. O carregamento pendente foi provocado controladamente; não se afirma ter determinado por que um worker deixou de responder em um PDF específico do atendimento sem esse arquivo/trace. Os defeitos da aplicação são reproduzíveis e foram tratados.

A nova implementação:

- Captura resultado e número da folha por tarefa; não consulta estado mutável depois para desenhar guias ou publicar o canvas.
- Renderiza em canvas isolado; a área visível é invalidada imediatamente. Só adota resultado vigente, sem permitir que renderizações antigas sobrescrevam estados novos.
- Usa worker PDF.js próprio por prévia, `RenderTask.cancel()`, carregamento/renderização com disputa explícita contra cancelamento e prazo de30 segundos.
- Encerramento memoizado; aguarda render/documento com limite de2 segundos por etapa e encerra o worker próprio ao final. Não destrói workers de outras ferramentas. A tarefa antiga pode levar até aproximadamente4 segundos adicionais de limpeza no caso de falha de comunicação; a fila não depende de confirmação indefinida.
- Geração também tem interrupção e prazo de60 segundos. O PDF-lib executa trechos síncronos: esses prazos são atendidos quando o JavaScript volta a processar eventos, não são preempção rígida de CPU/memória em PDFs extremos.
- Mudanças rápidas são agrupadas pelo debounce e número de versão. Cancelar/reabrir aguarda a limpeza da janela anterior. Falha exibe texto e **Tentar novamente**. A tentativa ou alteração válida libera novamente gerar/imprimir, sem reaproveitar erro antigo.

## Modos de marcas

O seletor **Marcas de corte** oferece:

- **Individuais:** padrão e comportamento da .17. Reservas, escala e posições existentes foram mantidas.
- **Apenas no grupo / contorno externo:** oito segmentos vetoriais nos quatro cantos do retângulo delimitador das peças presentes na folha; nenhuma marca interna por peça. A reserva das marcas é feita só ao redor do bloco.
- **Sem marcas:** preserva a possibilidade anterior de desativar marcas.

No modo grupo, espaçamento0 e sangria0 encostam peças regulares (Step & Repeat ou N-up com peças que ocupam igualmente as células). Se houver sangria, ela continua reservada e não é sobreposta: espaçamento0 não elimina a área necessária para as bordas. N-up de proporções/tamanhos diferentes continua centralizando cada peça na célula; isso pode deixar branco interno, mesmo com intervalo0. Vagas são preservadas. Folhas totalmente vazias não recebem marcas; folhas parciais recebem contorno das peças existentes, mantendo a grade/posições, sem recentralização inesperada.

O contorno é um retângulo envolvente, não uma faca de formato irregular e não fornece marcas para cada corte interno. As guias azuis por peça continuam opcionais, apenas na prévia e nunca no PDF exportado. Comprimento, espessura, afastamento e dimensões continuam em mm. Marcas pretas RGB e limites de sangria/cor/compatibilidade da .17 permanecem.

Não foram reescritas as matrizes de transformação, escala, incorporação vetorial, fontes ou grupo de transparência da origem. A geração apenas desenha as marcas externas da folha quando esse modo é escolhido. PDFWorkspace, exportação/impressão independentes, Hold/DEVMODE/PDFium/HDC e os temas foram preservados.

## Testes

- **296 testes Node aprovados**, incluindo todos os284 anteriores e12 novos. Novos testes cobrem carregamento pendente, abort explícito do render, timeout e limpeza sem resposta, erro/retry, resultado antigo,30 mudanças coalescidas e fechamento; marcas individuais/grupo/sem, peças encostadas, N-up, vagas/folhas parciais e coordenadas reais dos operadores do PDF; regressão A3 → A6 → Step & Repeat com origem ainda pendente.
- Comparação direta com motor geométrico preservado da .17: **144 configurações individuais/sem marcas** com modos, giros, tamanhos e espaçamentos diferentes produziram posições/escala e capacidade idênticas. Registro local `baseline-geometry.json`.
- **Chrome real e Electron realmente instalado:** nove prévias iniciadas e oito cancelamentos durante mudanças sucessivas, última concluída, gerar/imprimir liberados. Falha injetada seguida de Tentar novamente recuperou a janela. PDF gerado depois das alterações contém oito segmentos externos. Fonte conservou seus bytes. Relatórios `browser/preview-stress.json` e `installed/preview-stress.json`.
- Temas claro/escuro, controles dentro de1366×768, exportação com diálogo nativo, cancelar/reabrir e proteção da origem aprovados. Impressão executada **exclusivamente na Microsoft Print to PDF**, com fila explicitamente conferida; nenhum envio físico à Konica ou demais impressoras.
- Comparação sintética de texto e transparências da .17 reexecutada no Chrome:82.944 componentes RGBA iguais, diferença máxima0.
- Electron smoke aprovado: leitura/miniaturas/giro, abas, salvar/cancelar/falhar, proteção de fechamento e persistência. Pacote instalado auditado com **444 entradas**. Sintaxe e `git diff --check` verificados.

A primeira execução dos testes novos apontou rejeição não tratada de uma promise auxiliar do mock, sem uso no render concluído. A fixture foi corrigida e a bateria final passou; não houve supressão de erros de produção para passar testes.

## Arquivos desta etapa

- `public/js/tools/pdf-imposition-preview.js` (novo): ciclo isolado, cancelamento, prazos e limpeza da prévia.
- `public/js/tools/pdf-imposition-ui.js`: resultado estável, invalidação, erro/retry, seletor e reabertura segura.
- `public/js/tools/pdf-imposition-layout.js`: reserva externa e segmentos do grupo; ramo individual conservado.
- `public/js/tools/pdf-imposition-engine.js`: desenho único das marcas externas, sem alterar transformações de página.
- `tests/pdf-imposition-preview.test.mjs` (novo), `tests/pdf-imposition.test.mjs`, `tests/pdf-imposition-browser.mjs`.
- `electron-builder.hold-native-test.cjs`: versão experimental.18, mesma identidade.
- Documentação: este guia, `ESTADO_PROJETO.md`, `DISTRIBUICAO_DESKTOP.md` e registro da aprovação física no guia .17.

## Instalador e validação manual

Instalador: `dist-desktop/hold-native-test/4PrintShops-Hold-Test-1.0.0-experimental.18-Windows-x64-Setup.exe`.

SHA256: `6A77BB8162C25EC6D44E29938B3856FA6ECAEC899AFD0C35CA565015F1456877`.

Foi aplicado sobre a Hold Test existente no mesmo diretório, appId e perfil, sem criar instalação paralela. Sem assinatura Authenticode comercial; atualização automática com instalação/reinício continua não homologada.

1. Exporte trabalhos pendentes e feche Hold Test antes de instalar a .18 no destino existente.
2. Repita o cenário aprovado A3 → A6 → Step & Repeat, com **Individuais**, conferindo medidas. O teste físico complementar é decisão do operador.
3. Mude margens, quantidade, escala, grade e folha diversas vezes, inclusive antes da prévia terminar. A última configuração deve aparecer; gerar PDF deve voltar a habilitar. Configuração inválida deve mostrar erro; corrigir ou Tentar novamente deve recuperar.
4. Escolha **Apenas no grupo / contorno externo**, sangria0 e espaços0. Confira peças encostadas e somente marcas pretas nos cantos externos. Desative Guias azuis se quiser ver apenas as marcas exportadas.
5. Gere o PDF e reabra para conferir geometria e ausência de marcas internas. Teste N-up com páginas iguais e diferentes e última folha parcial. As páginas de proporções diferentes podem deixar branco interno.
6. Cancele e reabra a janela; a origem e suas pendências devem permanecer. Para testar impressão primeiro use Microsoft Print to PDF e confira papel/escala/N-up nas propriedades, sem acionar fila física.

Capturas/amostras/logs ignorados no Git em `.desktop-test/imposition18/`. Exemplo gerado após mudanças: `installed/Grupo apos alteracoes.pdf`; captura `installed/grupo-externo.png`. Limitações de PDFs complexos, cor/RIP, stress nos tetos e implantação/atualizador/SMB anteriores permanecem. A homologação física do novo modo de marcas ainda depende do usuário.
