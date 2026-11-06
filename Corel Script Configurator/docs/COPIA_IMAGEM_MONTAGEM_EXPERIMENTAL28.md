# Cópia como imagem para montagem/livreto — experimental.28

**Aceitação manual registrada em06/11/2026:** o usuário aprovou a preparação como imagem da .28 com o PDF Carmecita e autorizou o checkpoint conjunto com a .27. Aprovação restrita ao fluxo e ao arquivo efetivamente testados; **impressão física das novas cópias, fidelidade ICC/RIP e livretos ainda sem homologação completa**. Permanecem separados o caso `8) MIDIAS_Canais.pdf`, aprimoramento do reconhecimento/escolha de caixas de corte e sangria, autoscroll inferior, limites de zoom próximos às bordas e preparação vetorial futura para OCG. Os resultados automatizados e a avaliação instalada abaixo complementam, sem ampliar, essa aprovação manual.

06/11/2026. Base `c9384f5`, branch `main`. Os17 caminhos pendentes da .27/diagnóstico foram copiados com hashes e patch para `.desktop-test/raster28/before/`, ignorada, antes de alterar o código. Nenhuma alteração anterior descartada, commit, push, instalação ou impressão executada nesta etapa.

## Uso e comportamento

Quando a montagem ou o livreto encontra incompatibilidade sem preparação vetorial disponível, aparece uma alternativa explícita **Preparar cópia como imagem**, com 300/600 DPI. O botão vetorial indisponível não concorre visualmente com a ação válida. Motivos técnicos e perdas da cópia ficam em seções recolhíveis; a confirmação repete as perdas antes de iniciar qualquer rasterização.

A elegibilidade usa as verificações do preparador de impressão existente, em modo documento inteiro, e os limites de memória em ambas as resoluções. XFA, ações/scripts, aparências dinâmicas, assinatura visual não comprovável e outros casos recusados por esse preparador não são liberados. Falha ou limite explica a indisponibilidade; não reduz DPI nem remove verificações para continuar.

Se a tentativa vetorial falhar na comparação, a mesma alternativa pode ser oferecida nas duas janelas, após o preflight e mediante **nova confirmação**. A cópia vetorial recusada não é publicada e a saída permanece bloqueada. Esse caminho também recebeu teste de falha injetada e recuperação.

A folha sem conteúdo recebe mensagem explícita de **prévia bloqueada por compatibilidade**, sem ser apresentada como montagem final. Preparar não habilita saída imediatamente: todas as páginas da cópia são rasterizadas sequencialmente, reabertas e comparadas; o diagnóstico normal de montagem também precisa aceitar a cópia. Depois é gerada/renderizada a montagem/livreto com a geometria atual. Os bloqueios de geometria, seleção e aceite de brancos/cortes continuam valendo.

Apenas os bytes da janela são substituídos. O snapshot vem do PDFWorkspace atual, inclusive rotações e outras alterações não exportadas. O original, o documento do editor, sua revisão e sua condição de exportação permanecem intactos. Seleção, grade, quantidade, papel, escalas e marcas não são redefinidos pelo preparo. **Voltar ao original** recupera o snapshot original da janela; Cancelar fecha e descarta o preparo sem substituir o workspace. Ao fechar durante a operação, as tarefas são canceladas e não publicam resultados antigos.

Depois do sucesso, a origem indica imagem RGB e DPI, não oferece preparação recursiva e permite gerar/exportar pelos fluxos já existentes. Step & Repeat e N-up usam a cópia da mesma janela; livreto mantém sua janela independente. A impressão existente só fica acessível quando o resultado final é válido e o operador a solicita. Nenhuma impressão foi usada nos ensaios desta etapa.

## Limites e cores

- Todas as páginas tornam-se imagens RGB opacas sobre branco. Texto pesquisável, vetores, interatividade, camadas, acessibilidade e assinaturas ativas não são transportados. Não se preserva validade criptográfica de assinatura na cópia.
- ICCBased e OutputIntents não são apagados do original. Na reconstrução raster, sua influência depende da renderização PDF.js; o PDF resultante usa DeviceRGB e não mantém o vínculo ICC/intenção de saída como no original. Não há promessa de fidelidade colorimétrica, PDF/X, cores especiais, sobreimpressão ou equivalência de RIP.
- 300/600DPI são escolhas explícitas. Mais DPI não recupera resolução ausente. No original real examinado, o conteúdo já era uma imagem de aproximadamente 200 ppi.
- Reutilizados limites de 36 milhões de pixels e12000 por lado por página. Preflight adicional soma RGB estimado e recusa acima de 256 MiB; pode ser conservador mesmo quando a compressão final seria pequena. O limite de 256 MiB comprimidos/finais também é verificado. Orientação/UserUnit/CropBox entram no cálculo. Não há redução automática de resolução.
- A área renderizada é a área visível/CropBox em sua interseção com MediaBox. O preparador existente conserva/transforma as caixas conforme a orientação visual e normaliza UserUnit/Rotate. As caixas declaradas continuam guiando a montagem. **Não foi criado seletor novo de TrimBox/BleedBox**, sangria ou área personalizada.
- A comparação verifica a cópia reaberta na resolução escolhida contra a imagem proposta pelo mesmo renderizador, com tolerância existente de 2 níveis RGB e limite de 0,1%. Isso não comprova igualdade visual em qualquer escala ou com outro RIP. Revisão humana continua obrigatória antes de produzir.

## Arquitetura e arquivos desta etapa

Novos módulos:

- `public/js/tools/pdf-imposition-raster.js`: preflight, consentimento obrigatório, orquestração do preparador existente, comparação e recursos isolados de PDF.js.
- `public/js/tools/pdf-imposition-raster-ui.js`: opção comum às duas janelas, resolução, avisos e descarte de respostas antigas de elegibilidade.

Integração em `pdf-imposition-ui.js`, `pdf-booklet-ui.js` e CSS restrito em `theme.css`. Cada modal mantém sua sequência de operações, AbortController e revisão. Tarefas/Workers próprios são encerrados uma vez; não se chama destroy no documento PDF.js nem em tarefas do editor. Falhas não publicam cópias e permitem nova tentativa; cancelar continua disponível.

Não foram alterados os verificadores de compatibilidade, o preparador/rasterizador anterior, PDFWorkspace, geometria/materialização de montagem e livreto, redimensionamento, PDFium, fontes C#, mecanismos Windows, associação de PDFs, atualizadores ou dependências. A carga do módulo raster é adiada até existir uma oferta de compatibilidade.

Testes novos: `tests/pdf-imposition-raster.test.mjs`, `tests/pdf-imposition-raster-browser.mjs`, `tests/helpers/pdf-imposition-raster-browser.mjs`, `tests/helpers/pdf-imposition-raster-modules.mjs`. O teste de prévia existente recebeu o adaptador da nova composição de UI no Node, preservando suas asserções. Builder Hold Test passa somente a versão para `.28`. Documentação de estado/distribuição e este guia atualizados; diagnóstico anterior preservado como histórico.

## Validação executada

| Verificação | Resultado |
|---|---|
| Suíte Node final |405/405 aprovados; seis novos testes sobre consentimento, elegibilidade, memória, cancelamento e propriedade de recursos |
| Chrome headless |20 preparações pela interface,300/600DPI, montagem e livreto; Step & Repeat seguido de N-up; exportação normal e bytes salvos conferidos |
| Electron isolado |20 preparações equivalentes, incluindo caixas/rotação/UserUnit; salvamento pelo IPC Desktop idêntico aos bytes validados |
| Regressão Chrome anterior de montagem/livreto |Aprovada: vetor, preparação anterior/ActualText, seleção, geometria, retry, cancelamentos e temas |
| Apresentação final1366×768 |Conferida Chrome/Electron: botão/resolução visíveis sem rolar o aviso; mensagem de prévia bloqueada; origem identifica RGB/DPI após sucesso |
| Original real |300/600DPI;87,5×50mm preservados; hash original idêntico; prévias e saídas Step & Repeat, N-up e livreto conferidas localmente |
| Casos sintéticos |OCG ON/OFF, OutputIntents com perfil sRGB do sistema, três páginas em ordem, origem girada/pendente, caixas distintas com270°/UserUnit2, PDF comum, XFA e documento acima do limite |
| Falhas/cancelamentos |Recusa da confirmação, erro de rasterização e retry na mesma janela, render pendente cancelado ao fechar, abort inicial, saída bloqueada até cópia válida |
| Pacote |ASAR458, código/recursos conferidos; conteúdo do próprio Setup extraído e auditado; identidade Hold Test preservada |

Uma rodada ampliada do Electron executada com outra bateria concorrente excedeu o prazo de180s do harness durante comparação de imagem, sem exceção capturada do aplicativo. A causa não foi isolada. A rodada final sem concorrência, com foco emulado/primeiro plano, passou nos20 casos, incluindo a exportação. O harness passou a explicitar esse foco. Não se declara que esse ensaio tenha identificado/corrigido um defeito do renderizador; avaliar demora em segundo plano se houver relato real. Depois do ajuste final de posicionamento dos controles, houve nova conferência visual Electron e nova rodada Chrome completa; lógica de preparo não mudou.

Aceitação final em `chrome-acceptance/` e `electron-acceptance/`:20 preparações em cada, mais dois casos em cada roteiro de oferta raster após falha de comparação vetorial (montagem e livreto), recuperados com nova confirmação. ExportsMatchValidatedBytes e preservação do original/workspace confirmados nos dois relatórios. Esses resultados cobrem o código final, inclusive a recuperação adicional; nenhum envio de impressão.

Resultados/cópias/capturas autorizadas somente em `.desktop-test/raster28/`, ignorada. Os testes permanentes geram sintéticos; o PDF real só é lido se indicado explicitamente por variável de ambiente. Nenhum arquivo de cliente ou captura faz parte do instalador/repositório. O runner Node mantém os avisos anteriores de Math.sumPrecise da biblioteca, sem falhas nos testes. A suíte geral Chrome de miniaturas não foi repetida; autoscroll continua com a pendência conhecida da .27.

### Reproduzir sem imprimir

```powershell
node --test tests/*-*.test.mjs
node tests/pdf-imposition-raster-browser.mjs
$env:RASTER_ELECTRON='1'
node tests/pdf-imposition-raster-browser.mjs
```

O roteiro browser usa Chrome ou Electron instalado nas dependências do projeto, perfil de teste isolado e perfil sRGB padrão do Windows somente para gerar o sintético. Pode indicar um original autorizado com `RASTER_REAL_PATH` e mudar a pasta de provas com `RASTER_RESULTS`. Não usar caminho de cliente fixo em testes versionados. Os ensaios geram/exportam PDFs localmente e nunca acionam a impressão.

## Distribuição

Setup: `dist-desktop/hold-native-test/4PrintShops-Hold-Test-1.0.0-experimental.28-Windows-x64-Setup.exe`.

Mesmos `br.com.4printshops.holdtest`, produto/executável, atalhos, chave NSIS e perfil `%APPDATA%/4PrintShops Hold Test`; preferências/dados não são apagados. **Não instalado nesta etapa: o registro Windows permanece em1.0.0-experimental.27.** Sem assinatura comercial; nenhuma verificação do atualizador foi desabilitada. Sem publicação.

Setup final:126.423.277 bytes. SHA256: `B83266368F994F07337CA5067D764738A54C121AEDAC5CE21CDE5A14D657C07C`. Não confundir o pacote intermediário gerado antes do ajuste de visibilidade dos controles com o Setup final entregue.

## Conferência manual após instalação autorizada

1. Abrir um PDF de teste incompatível e entrar em Montagem ou Livreto. Conferir o motivo técnico, mensagem de prévia bloqueada e saídas indisponíveis.
2. Selecionar300DPI, abrir perdas/limites e clicar Preparar cópia como imagem. Recusar uma vez: nada deve mudar.
3. Confirmar; aguardar a comparação e conferir conteúdo, dimensões, orientação e indicação RGB/DPI. As configurações de montagem devem permanecer iguais.
4. Gerar o PDF, reabrir e conferir as peças no tamanho esperado. Testar Step & Repeat e N-up; no livreto, revisar e aceitar brancos propostos conscientemente.
5. Voltar ao original ou cancelar; o documento e as alterações pendentes do editor devem permanecer intactos. Repetir a600DPI a partir do original, não de uma cópia raster anterior.
6. Em um arquivo simples, verificar que o fluxo vetorial anterior continua normal. Casos recusados por risco/dinâmica/memória devem continuar bloqueados com motivo.

Pendências independentes: redimensionamento/exportação de `8) MIDIAS_Canais.pdf`; seletor futuro de áreas TrimBox/BleedBox; origem do OutputIntents da captura intermediária anterior; autoscroll; homologações físicas/colorimétricas/livreto e dos atualizadores permanecem separadas. Nenhuma dessas pendências foi considerada resolvida por esta entrega.

Verificação final: sintaxe dos18 arquivos JS/MJS/CJS pendentes aprovada e `git diff --check` sem erros. Git com25 caminhos pendentes (17 anteriores preservados), staging vazio. Nenhuma alteração no código dos motores/desktop ou instalação desta versão.

## Instalação autorizada e avaliação na gráfica — 06/11/2026

Esta seção é posterior à geração descrita acima: **experimental.28 efetivamente instalada sobre a .27**, após autorização expressa, sem nova compilação ou alteração de código. Setup e SHA256 conferidos novamente; aplicação/auxiliares fechados e sem documento aberto. As preferências foram lidas na .27 e ela foi fechada normalmente antes da execução do instalador, que terminou com código0.

Mesmo diretório: `.desktop-test/hold-lab/installed-validation/installed/4PrintShops Hold Test` neste checkout. Mesmo perfil `%APPDATA%/4PrintShops Hold Test`, origem `app://4printshops` e chave NSIS `90ff2ff0-f013-571c-b109-7edb265dacc8`. Uma única instalação Hold Test registrada. FileVersion e registro confirmam1.0.0-experimental.28; aplicativo realmente iniciado desse executável após a atualização. Executável, ASAR, ícone PDF, auxiliares e PDFium instalados conferidos por hash contra a extração do Setup; auditoria ASAR458 aprovada.

Preferências reais comparadas integralmente antes da instalação, na primeira abertura .28 e após os ensaios:

| Preferência | Antes .27 | Após .28/testes |
|---|---|---|
| Painéis PDF |versão1, esquerda312, direita230 |Idêntica |
| Tema |light |Idêntica |
| Marcador de configuração |preservado |Idêntica |

Foram verificados os menus reais e a importação no editor. Step & Repeat, N-up e Livreto disponíveis. O código do próprio aplicativo instalado passou na matriz de20 preparações de montagem/livreto a300/600DPI, incluindo o original autorizado, sintéticos OCG ON/OFF, OutputIntents e caixas/270°/UserUnit2. Preservação de configuração/seleção/escala, avisos de perda de vetores/texto e diferenças de cor, recusa de confirmação, cancelamento de render pendente, falha/retry e alternativa após falha vetorial aprovados. Saídas permaneceram bloqueadas antes da cópia validada; não houve preparo recursivo após sucesso.

**Voltar ao original** foi exercitado em todos os20 casos e recuperou exatamente os bytes anteriores. O workspace de teste com alterações não exportadas permaneceu intacto e pendente. No original real,87,5×50mm preservados; comparação da cópia reaberta sem diferenças acima da tolerância em1.833.282 componentes RGB a300DPI e7.329.582 a600DPI. Essa prova continua limitada ao PDF.js, sem equivalência de RIP ou cor física.

Neste roteiro instalado, as funções de exportar foram omitidas somente na cópia ignorada do harness de avaliação: **não foi salvo/exportado nenhum PDF nem enviada impressão física/virtual**. A geração dos bytes validados e a prévia foram testadas; a exportação permanece com os resultados aprovados da etapa de desenvolvimento anterior. O original no disco manteve SHA256. Nenhuma exceção de execução capturada. Preferências preservadas ao final e aplicativo fechado normalmente, sem encerrar à força ou descartar trabalho do usuário.

Evidências locais ignoradas: `.desktop-test/raster28/installed-evaluation/before.json`, `after.json`, logs, inventário de hashes anterior e capturas locais. Nenhum dado de cliente foi adicionado ao Git.25 caminhos pendentes preservados, staging vazio, apenas três documentos atualizados nesta avaliação. Sem commit/push ou alterações de impressoras. Permanecem as limitações de memória/RGB/RIP, MÍDIAS, futuro seletor de caixas, autoscroll e demais homologações anteriores.
