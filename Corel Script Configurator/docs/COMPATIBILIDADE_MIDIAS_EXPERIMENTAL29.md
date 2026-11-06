# Compatibilidade do redimensionamento — experimental.29

**Checkpoint aprovado na gráfica — 06/11/2026:** usuário confirmou teste manual de redimensionamento seguido de montagem Step & Repeat aprovado na experimental.29, limitado ao cenário efetivamente testado. Correção vetorial preserva ActualText e conteúdo no ajuste do PowerPoint de 254 × 190,5 mm para 200 × 150 mm. Exportação PDF/PNG/JPEG aprovada nos testes documentados; a falha histórica não foi reproduzida isoladamente. Instalação .29 sobre .28 concluída com identidade/perfil/preferências preservados. Impressão física desta correção e fidelidade ICC/RIP não homologadas. Permanecem pendentes TrimBox/BleedBox e demais caixas de acabamento, autoscroll inferior, zoom nas bordas, homologação física dos livretos e preparação vetorial OCG. Checkpoint local autorizado, sem novas funcionalidades ou empacotamento; push exclusivamente manual pelo usuário.

## Diagnóstico em 06/11/2026

Base limpa `b36ff4d`, branch `main`. O original real autorizado `8) MIDIAS_Canais.pdf` foi processado exclusivamente em arquivos locais ignorados, sem upload, alteração do original ou inclusão de conteúdo de cliente nos testes permanentes.

**Redimensionamento reproduzido:** página720×540pt (254×190,5mm), estrutura marcada, ParentTree/StructParents e26 entradas ActualText. A implementação anterior tentava provar que o texto substituto poderia ser descartado junto com a acessibilidade. A primeira entrada não demonstrou redundância e a operação foi corretamente recusada por esse caminho. O botão Aplicar ficava desabilitado; os campos e Pré-visualizar continuavam disponíveis. Não foi constatado travamento do workspace ou falha silenciosa de renderização.

**Exportação histórica não reproduzida:** o original renderiza e exportou PNG na base .28, inclusive pelo fluxo completo Electron/IPC após a recusa do redimensionamento. Não foi encontrada dependência da exportação de imagens em relação à compatibilidade de edição. O operador esclareceu que utilizou outro software para converter o original antes de reimportar e imprimir. Isso confirma o problema operacional, mas não identifica o estado ou mensagem da tentativa antiga de exportação. Não declaramos essa causa histórica corrigida. Uma prévia válida ainda não aplicada bloqueia intencionalmente a exportação, para evitar exportar bytes diferentes dos exibidos.

## Correção restrita

- Mantém a árvore marcada e ActualText no próprio documento transacional; não exige descartar texto essencial para aplicar a transformação vetorial existente. Nenhuma rasterização, conversão de fontes ou mudança de escala/posicionamento do motor.
- Caminho permitido somente quando ActualText/acessibilidade são os únicos impedimentos e não existem atributos de layout/classes, referências a conteúdo externo ou outras estruturas sem tratamento. Varredura limitada e conferência da árvore, catálogo e identificadores de página antes de devolver a preparação.
- Formulários, ações, anotações/recursos incompatíveis e o descarte de ActualText essencial continuam protegidos. O modo explícito que remove estruturas mantém a prova anterior; a nova opção não relaxa sua validação.
- Prévia e cancelamento não alteram o workspace. Aplicar registra alteração pendente; a prévia temporária informa que é necessário aplicar ou cancelar antes de exportar.
- Quando houver bloqueio técnico, **Exportar PNG/JPEG como alternativa** pede confirmação com aviso de pixels, perda de vetores/texto pesquisável e diferenças de cor; abre a janela compartilhada em PNG. Não exporta automaticamente nem aplica um redimensionamento pendente. A conversão já existente usa os bytes atuais do workspace, DPI e limites de memória. Para trabalhar com uma imagem, exporte e importe-a conscientemente pelo fluxo existente; não há substituição automática do documento.
- Preparação raster da montagem/livreto .28, detecção OCG/OutputIntents, edição, motores Windows, associação e atualizadores não foram alterados.

Preservar a marcação não equivale a revalidar acessibilidade/PDF/UA: essa limitação aparece na prévia. Documentos com atributos estruturais geométricos continuam exigindo outro tratamento.

## Validações executadas

| Verificação | Resultado |
|---|---|
| Node completo |410 testes aprovados, zero falhas |
| Sintaxe |63 arquivos aprovados, incluindo módulos PDF e testes/configuração alterados |
| C# existente |41 verificações aprovadas; sem envio físico |
| Chrome headless e Electron |3 casos por ambiente: sintético com fonte incorporada/símbolos/transparências/ActualText, original autorizado, estrutura com atributos não suportados |
| Original 254×190,5 →200×150mm |Escala78,74015748%; dimensão exata na MediaBox; árvore e ActualText mantidos |
| Integridade vetorial/texto |Streams originais de conteúdo/recursos preservados byte a byte; texto extraído idêntico pelo PDF.js e pypdf (29 caracteres neste extrator) |
| Prévia / saída |Bytes da preparação iguais aos bytes exportáveis após Aplicar; cancelar e reabrir prévia preservaram estado |
| PNG e JPEG |150/300/600DPI no original redimensionado e sintético;1182×886,2363×1772 e4725×3544px; exportação PNG real por IPC no Electron |
| Estado/recuperação |Alterações permanecem pendentes após imagens; falha forçada de codificação cancela o destino e permite nova exportação; alternativa exige confirmação; proteções de substituição preservadas |
| Regressão imagens/atalhos Chrome |Giro, PNG150/300/600, qualidade JPEG, cancelamento com limpeza do destino e lote sintético de350 páginas aprovados |
| Montagem/livreto Chrome |Regressão existente aprovada: N-up, Step & Repeat, livreto, prévias, cancelamentos, falha/retry e exportação |
| Impressão virtual |Somente Microsoft Print to PDF: três páginas mistas, escalas100/120/87,5%, formulário, rejeição de PDF inválido/token cancelado e preferências preservadas |
| Pacote / Git |ASAR458 entradas e conteúdo extraído do Setup conferidos; fontes, ícones, auxiliares e PDFium correspondentes; `git diff --check` aprovado |

Comparação visual normalizada do original antes/depois: no Chrome220 de1.555.200 componentes RGBA diferiram mais de2 níveis (máximo9); no Electron171 componentes (máximo47). São diferenças residuais de renderização, sem perda de elementos observada. Não é identidade matemática dos pixels nem prova de equivalência colorimétrica com RIP. A comparação separa a redução geométrica intencional da conservação do conteúdo. O arquivo original foi reconferido por SHA256 e permaneceu intacto.

Evidências e derivados locais em `.desktop-test/midias29/`, ignorados. Testes permanentes usam apenas conteúdo sintético; o original é opcional por variável de ambiente e nunca é incorporado ao código.

## Arquivos desta etapa

- Código: `public/js/tools/pdf-resize-engine.js`, `pdf-resize-interface.js`, `pdf-resize-ui.js`, `pdfs.js`; `public/js/app/export-dialog.js`.
- Configuração: `electron-builder.hold-native-test.cjs`.
- Testes: `tests/pdf-actualtext.test.mjs`, `app-export-dialog.test.mjs`, novo `pdf-tagged-resize.test.mjs`, novo `pdf-tagged-resize-browser.mjs`, novos auxiliares `tests/helpers/pdf-tagged-resize-browser.mjs` e `desktop-tagged-resize-entry.mjs`.
- Documentação: este relatório, `ESTADO_PROJETO.md` e `DISTRIBUICAO_DESKTOP.md`.

## Distribuição e conferência manual

**Atualização de instalação em06/11/2026:** após autorização explícita, Setup/hash reconferidos e aplicativo/auxiliares fechados. As preferências foram capturadas na .28, sem documento aberto, e o aplicativo fechado normalmente antes da atualização. Setup terminou com código0 no mesmo diretório, chave NSIS e perfil Hold Test. Registro/FileVersion confirmam1.0.0-experimental.29 e seis arquivos instalados correspondem ao Setup auditado. Nenhuma instalação duplicada. Preferências anteriores/posteriores integralmente iguais: painéis312/230, tema light e marcador preservado. Na versão instalada, PDF sintético renderizou, campos/Pré-visualizar de redimensionamento estão disponíveis e Exportar abriu com PDF/PNG/JPEG; cancelado sem salvar ou imprimir. Aplicativo fechado normalmente ao finalizar. Primeira tentativa automática não conseguiu conexão ao diagnóstico dentro de60s; nova abertura passou sem reinstalação ou alteração de código. Causa inconclusiva, logs preservados. Somente documentação alterada nesta instalação;15 caminhos anteriores preservados. Avaliação manual do arquivo de produção continua pendente. A nota de não instalação a seguir registra o momento da geração, anterior a esta autorização.

Setup: `dist-desktop/hold-native-test/4PrintShops-Hold-Test-1.0.0-experimental.29-Windows-x64-Setup.exe`.

126.424.116 bytes; SHA256 `FCAA3F9DBC9909F16D44EBC761A4404A3DED59237B1D6E125F6AD8FB79EF5E92`.

Mesmos appId `br.com.4printshops.holdtest`, produto/executável e perfil Hold Test, NSIS por usuário e associação existente. Preparado para atualizar .28, **não instalado**: a estação permanece .28. Preservação real de preferências após instalação da .29 e aceite manual da gráfica ainda pendentes. Instalador experimental sem Authenticode comercial; verificações do atualizador não foram enfraquecidas e instalação/reinício automáticos permanecem não homologados.

1. Após autorização para instalar, exportar trabalhos pendentes e fechar Hold Test normalmente. Conferir hash e atualizar a instalação existente.
2. Abrir o original MÍDIAS; informar200×150mm e Ajustar proporcionalmente. Conferir prévia, cancelar uma vez e verificar retorno ao original; gerar novamente e Aplicar.
3. Conferir200×150mm, texto/logos/imagens, salvar cópia PDF com outro nome e reabri-la. Nunca sobrescrever o original.
4. Exportar PNG/JPEG em150/300/600DPI. Conferir dimensão, orientação e conteúdo; o indicador pendente do PDF deve continuar até salvar PDF.
5. Testar um PDF que continue bloqueado: alternativa de imagem deve explicar perdas, respeitar recusa e abrir exportação somente após aceite.
6. Conferir edição, N-up, Step & Repeat e livreto com sintéticos. Nenhuma impressão física foi feita nesta etapa.

Permanecem separados: seletor TrimBox/BleedBox, autoscroll inferior, limites de zoom, preparação vetorial OCG, homologação física completa de livretos e fidelidade ICC/RIP. Não há novos controles para esses assuntos, instalação, commit ou push nesta etapa.
