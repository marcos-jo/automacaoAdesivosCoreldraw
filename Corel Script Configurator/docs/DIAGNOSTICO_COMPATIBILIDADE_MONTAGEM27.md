# Diagnóstico de compatibilidade de montagem/livreto — experimental.27

Data: 06/11/2026. Base `c9384f5`, branch `main`, com16 caminhos pendentes da .27 preservados. Esta etapa é de diagnóstico e proposta: **nenhum código de produção foi alterado, bloqueio removido, instalador gerado/instalado ou trabalho de impressão enviado**.

O original e a captura autorizados foram examinados somente localmente. PDFs de cliente, imagens, dados do conteúdo, metadados pessoais e derivados não acompanham este documento nem os testes permanentes. Evidências e scripts específicos ficam em `.desktop-test/carmecita-diagnostic/`, ignorada pelo Git.

## Resultado confirmado no original

- PDF1.7 de uma página,87,5 ×50 mm, sem rotação,187.745 bytes.
- Uma camada OCG ativa, declarada em `/OCProperties/OCGs` e na configuração `/D/ON`; `/OFF` vazio. O conteúdo gráfico da página está dentro de `/OC /Pr12 BDC … EMC`. Não é uma entrada de catálogo sem uso.
- Uma imagem689×393,8bits, `/ICCBased`, perfil RGB com3 componentes,3.144 bytes, idêntico ao sRGB padrão presente no Windows. Resolução efetiva aproximada200ppi no tamanho atual. Aumentar a resolução de uma cópia não recupera detalhes inexistentes na imagem original.
- Nenhuma fonte ou texto extraível: a escrita visível já faz parte da imagem do original. Não há ActualText, árvore marcada, formulário ou anotação no arquivo examinado.
- Não foram encontrados Group/SMask ou parâmetros de opacidade/blend que indiquem transparência. Os estados gráficos encontrados definem RelativeColorimetric e sobreimpressão desligada. Não foi realizada prova colorimétrica no RIP.
- Há marcadores e destinos de navegação, removíveis pelo preparador existente; eles não são o impedimento decisivo.
- **Não existe `/OutputIntents` no catálogo deste original.** ICCBased e OutputIntents são mecanismos distintos. O perfil ICC usado pela imagem não disparou o bloqueio observado.

## Captura e tentativa de redimensionamento

A captura local mostra bloqueio por `OutputIntents`. O operador esclareceu que ela foi feita **após uma tentativa de redimensionar para montar com marcas**. Esse estado intermediário não deve ser confundido com o catálogo do original examinado.

Foram feitas provas somente em memória do redimensionador atual, com ajuste para90×50 mm e A4, modo de impressão confirmado apenas nas cópias. Ambas preservaram OCProperties, removeram somente marcadores/destinos e **não criaram OutputIntents**; a montagem continuou recusada por OCG. Portanto, não foi reproduzida a introdução de OutputIntents por esse caminho. A origem exata da entrada mostrada na captura continua dependendo do arquivo intermediário e da sequência exata de operações. Não se presume que a captura esteja errada nem que o redimensionador a tenha causado.

## Causa e classificação

| Camada | Constatação |
|---|---|
| Segurança | `pdf-imposition-compatibility.js` bloqueia por presença de AcroForm, OCProperties, OutputIntents e ações não suportadas. No original, o impeditivo é `OCProperties`; no sintético com OutputIntents, o código desse recurso foi confirmado separadamente. |
| Preparador | `podePreparar` exige recursos removíveis e ausência de impeditivos, exceto ActualText sujeito à sua prova específica. Com OCG, permanece falso e o botão fica desabilitado por projeto. Não existe hoje uma preparação vetorial homologada de camadas ou perfis de saída. |
| Motor | `gerarMontagem()` diagnostica a compatibilidade antes do planejamento e da materialização. Ao recusar, não gera bytes nem resultado de montagem/livreto. Compartilhamento desse motor explica o bloqueio nas duas ferramentas. |
| Renderização | O original abriu no editor e renderizou em PDF.js nos intents display e print, além de renderizar no Poppler. Não foi reproduzida falha de renderização. Nenhuma exceção de execução foi capturada no Chrome. |
| Interface | Em erro, a montagem desenha a folha de destino vazia com `desenharFolhaVazia`; o livreto a desenha antes da geração. É um placeholder sem conteúdo validado, não uma página original vazia. Falta sinalização sobreposta mais clara e orientação específica ao recurso bloqueado. |
| Estado | `result` fica nulo; gerar/imprimir continuam bloqueados corretamente. Não foi observado botão preso por trabalho antigo ou race condition nessa reprodução. |

Não há evidência de que marcas de corte, espaçamento ou sangria provoquem esse bloqueio: a recusa acontece antes do cálculo geométrico.

## Por que não remover apenas a verificação

O materializador cria um PDF novo e incorpora páginas como Form XObjects. Recursos de página, inclusive ICCBased, são transportados; Group de transparência recebe tratamento explícito. **O catálogo de OCProperties e OutputIntents não é transferido pelo fluxo atual.** Permitir a entrada sem implementar essa responsabilidade pode deixar OCG sem configuração ou perder a intenção de saída.

`embedPage` no pdf-lib utiliza seu próprio `PDFObjectCopier`. Copiar depois o catálogo por outro copier não demonstra que as referências OCG dentro dos Forms sejam as mesmas registradas em OCProperties. Igualdade de nomes de camadas não comprova identidade. Essa alteração exige testes estruturais e de visibilidade, não uma exceção por nome de arquivo.

A [referência oficial PDF1.7 da Adobe](https://opensource.adobe.com/dc-acrobat-sdk-docs/pdfstandards/pdfreference1.7old.pdf), seções4.10 e10.10, descreve a configuração de conteúdo opcional e as intenções de saída. [PDOCConfig](https://opensource.adobe.com/dc-acrobat-sdk-docs/acrobatsdk/apireference/PD_Layer/PDOCConfig.html) distingue configuração padrão e estados de visibilidade. São a base conceitual para preservar esses recursos, sem apagá-los indiscriminadamente.

## Alternativas e proposta para a próxima etapa

### Preferência: preparação sem rasterização

Tratar OCG e OutputIntents separadamente. Para OCG, implementar transferência consistente de configurações e referências, ou uma preparação restrita que resolva a visibilidade de impressão e remova somente os operadores de camada comprovadamente dispensáveis por meio de parser de conteúdo, nunca regex genérica. O caso de uma camada sempre ligada é candidato, não autorização para apagar camadas arbitrárias. OFF, OCMD, Usage/Print, configurações alternativas e camadas com estados distintos precisam de provas próprias ou continuar bloqueados.

Para OutputIntents, estudar transporte integral das intenções e perfis ao PDF de destino, com comparação dos streams ICC e dos vínculos; não converter silenciosamente para RGB nem aplicar um perfil de um documento a conteúdos com intenções conflitantes. Caso um arquivo provenha de união, sua proveniência de cores/camadas também precisa ser considerada. Não se deve prometer manutenção da conformidade PDF/X apenas copiando essa entrada.

Manter preparação transacional fora do workspace, comparação de todas as páginas antes/depois nos modos de visualização/impressão pertinentes, dimensões, texto quando existente e validação da montagem final. Cancelar não publica uma cópia. Preview e saída só usam bytes finais aprovados. Não foi aplicada uma correção funcional pequena porque a simples liberação do diagnóstico não satisfaz essas garantias.

### Alternativa explícita: cópia raster de impressão

Foi exercitado **o renderizador/preparador já existente**, sem mudar seu código: `reconstruirCopiaImpressao`, `rasterizarPaginaImpressao` e `conferirCopiaImpressao`, modo documento inteiro, em300 e600DPI. Sem ignorar verificações da montagem, as cópias aprovadas passaram pelo diagnóstico normal e geraram PDFs independentes de Step & Repeat e livreto.

| Prova |300DPI |600DPI |
|---|---:|---:|
| Imagem da página |1034×591 |2067×1182 |
| Componentes RGB comparados após reabrir |1.833.282 |7.329.582 |
| Diferenças acima da tolerância de2 níveis |0 |0 |
| Dimensão física |87,5×50mm |87,5×50mm |
| Diagnóstico da cópia |sem impeditivos |sem impeditivos |

**Limite da prova:** o comparador mede a imagem proposta pelo PDF.js contra a cópia reaberta na mesma resolução. Não é prova de equivalência de todos os pixels do original em qualquer escala, nem de fidelidade colorimétrica no RIP. Há reamostragem; a inspeção em resolução de tela mostra suavização. A cópia usa imagem DeviceRGB opaca sobre branco; o vínculo ICC original não é mantido nessa reconstrução. No original examinado já não existem textos/vetores, mas em outros documentos eles seriam convertidos em imagem, com perda de texto pesquisável, edição e possível diferença de transparência/cor.

A eventual integração à montagem deve oferecer ação separada, resolução e avisos explícitos, prévia revisável e confirmação do operador, sempre em uma cópia da janela, sem substituir o documento do editor. Não tornar essa escolha automática e não apresentar rasterização como preservação do perfil ou recuperação de resolução. Não habilitar saídas antes da comparação/revisão final. **Essa alternativa foi demonstrada localmente, não adicionada à interface da .27.**

### Apresentação sugerida

Título: “Montagem bloqueada por compatibilidade”. Resumo específico: “Este PDF usa camadas OCG; a montagem atual ainda não preserva sua configuração com segurança.” Para OutputIntents, mensagem própria identificando gerenciamento de cores de saída. Manter detalhes completos e motivo de indisponibilidade do botão.

Sobre a folha de espera: “Prévia da montagem indisponível: conteúdo bloqueado por compatibilidade. O documento original não está vazio.” Diferenciar de falha de renderização e de excesso geométrico. Quando houver alternativa validada, mostrar opção e perdas previstas. Não reutilizar a autorização de corte nas bordas para liberar compatibilidade.

## Caixas de página — investigação separada

No original, MediaBox, CropBox, TrimBox, BleedBox e ArtBox coincidem em87,5×50mm, com diferenças de representação de zero apenas. **Não há sangria adicional declarada nessas caixas.** Não inferir sangria por aparência visual.

O motor usa TrimBox explícita como peça final; sem ela, usa CropBox, e sem esta, MediaBox. BleedBox cai para TrimBox quando ausente. Exige contenção das caixas e limita a sangria solicitada ao que está disponível, considerando rotação/UserUnit. O recorte incorporado usa o TrimBox mais o padding da sangria escolhido. ArtBox não é uma opção de área na interface.

O sintético com MediaBox300×200pt, CropBox280×180pt, TrimBox260×160pt e BleedBox270×170pt foi aceito com peça final91,7222×56,4444mm, correspondente à TrimBox, e texto preservado. Não foi confundido com incompatibilidade de catálogo.

Futuro seletor, a aprovar separadamente: Corte final/TrimBox; incluir sangria limitada à BleedBox; página inteira/MediaBox, incluindo marcas existentes; área personalizada validada. Deve deixar claras a área escolhida, origem dos valores, inclusão de marcas e relação com CropBox, inclusive herança de caixas. **Nenhuma dessas opções foi implementada nesta etapa.**

## Testes executados e limites

-76 testes Node pertinentes: compatibilidade de montagem/impressão, geometria, excesso, motor e paginação de livreto. Todos aprovados. O runner emitia avisos de `Math.sumPrecise` da biblioteca em Node; não reprovaram testes nem apareceram como exceções no Chrome.
-13 testes Node de redimensionamento aprovados após a informação sobre a captura intermediária. Total desta investigação:89 testes existentes aprovados; não representa nova execução da suíte completa399.
-Chrome headless real, perfil isolado: abertura do original no editor; bloqueio nas duas janelas; preparar/gerar/imprimir indisponíveis conforme aplicável; renderização display/print; cópias300/600 e geração subsequente sem bypass.
-Sete casos sintéticos: comum, ICCBased, OutputIntents, OCG ON, OCG OFF, caixas diferentes e formulário. Comum/ICCBased/caixas aceitos com texto pesquisável; stream do perfil ICC no caso aceito idêntico byte a byte. OutputIntents/OCG/formulário continuam bloqueados. A camada OFF produziu página branca, demonstrando por que não se pode simplesmente ligar/apagar camadas para contornar o diagnóstico.
-Inspeção visual de original e cópia300DPI; comparação automatizada de reabertura300/600DPI. Original também renderizado pelo Poppler. Não houve avaliação colorimétrica, física ou em múltiplos RIPs.
-Original preservado por SHA256 antes/depois. Arquivos derivados e capturas somente na pasta ignorada. Não executados novos testes Electron instalado, NSIS ou toda a suíte Chrome; pacote/código da .27 não foram alterados.

## Pendências independentes

1. Corrigir compatibilidade OCG/OutputIntents conforme proposta aprovada e melhorar o estado visual de bloqueio, com testes permanentes sintéticos.
2. Obter/reproduzir o estado intermediário exato da captura para explicar a presença de OutputIntents; os dois destinos de teste do redimensionador atual não reproduziram essa entrada.
3. **Redimensionamento e exportação de `8) MIDIAS_Canais.pdf`: problema relatado pelo operador, registrado separadamente, sem diagnóstico, correção ou declaração de aprovação nesta etapa.** Não usar os resultados do Carmecita para considerar esse caso resolvido.
4. Autoscroll das miniaturas continua com a pendência anterior. Não investigado nem alterado nesta etapa. Livretos, cor e cenários de produção mantêm suas homologações específicas pendentes.

Sem commit/push, instalação, alterações de impressoras ou dependências.

Encaminhamento posterior de06/11/2026: alternativa raster explícita implementada na experimental.28, sem remover os bloqueios vetoriais aqui diagnosticados. Ver [COPIA_IMAGEM_MONTAGEM_EXPERIMENTAL28.md](COPIA_IMAGEM_MONTAGEM_EXPERIMENTAL28.md). Este relatório permanece como histórico da investigação anterior; a origem do OutputIntents intermediário, o seletor de caixas e o caso MÍDIAS continuam pendentes separadamente.
