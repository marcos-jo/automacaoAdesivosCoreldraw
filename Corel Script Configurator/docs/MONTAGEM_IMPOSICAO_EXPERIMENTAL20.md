# Preparar cópia para montagem — experimental.20
**Validação manual informada pelo usuário em 05/11/2026:** o usuário relatou que testou a .20 e que está tudo certo, solicitando auditoria e checkpoint antes de continuar. Aprovação para os cenários por ele exercitados, sem extrapolar para todos os PDFs, cores ou modos de impressão. Inventário do checkpoint em [AUDITORIA_EXPERIMENTAIS16_20.md](AUDITORIA_EXPERIMENTAIS16_20.md).

Data: 05/11/2026. Continuação das .16–.19, sem commit/push. Base Git `6b8f63a`, branch main. Os29 caminhos pendentes foram copiados antes da edição em `.desktop-test/imposition20/baseline19-20261105-133813/` com inventário e patch; nenhum trabalho anterior foi descartado.

## Diagnóstico do bloqueio

O motor vetorial recusava qualquer catálogo com **StructTreeRoot**, antes de verificar anotações. Essa é a verificação que produzia a mensagem relatada, não a simples presença de links. Após esse bloqueio, a janela apagava a prévia e deixava apenas o erro no rodapé, sem uma ação de preparação.

No arquivo real autorizado foram identificadas1 página960×540pt,168 entradas ActualText na árvore e **nenhuma anotação/link PDF**. O nome e a aparência de um arquivo não bastam para identificar hiperlinks ativos. O original foi apenas lido e seu SHA256 permaneceu igual ao final dos testes. Não há documento/conteúdo de cliente em fixtures ou no pacote.

A nova camada reutiliza `diagnosticarImpressao`, `validarActualTextImpressao`, `prepararParaImpressao` e o comparador `validarCompatibilidadeImpressao`, sem alterar esses módulos. Ela separa:

- Marcação/acessibilidade, marcadores e destinos que podem ser removidos de uma cópia de impressão mediante confirmação e validação.
- ActualText na árvore: requer comprovação de redundância em seus próprios MCIDs e comparação de texto/renderização. ActualText essencial/ambíguo continua bloqueado. ActualText nos streams de conteúdo não é removido.
- Links **sem flag Print, AP ou AA**, que podem perder interatividade somente na cópia. Mesmo esses links exigem comparação integral da aparência impressa e do texto após a remoção.
- Formulários, aparências de anotações, links imprimíveis, notas/tipos não suportados, ações complexas, camadas opcionais e perfis OutputIntents: continuam bloqueados. AcroForm é conservadoramente bloqueado inclusive quando vazio. Não se afirma que todo documento Word/PowerPoint/Acrobat seja compatível.

## Fluxo para o operador

1. Abrir Montagem / Imposição sobre o workspace atual. Quando houver incompatibilidade, aparece uma explicação visível na área de prévia, com **Motivos técnicos** recolhíveis.
2. Se houver caminho vetorial de preparação, usar **Preparar cópia para montagem**. Ler e confirmar a remoção de acessibilidade/navegação/interatividade na cópia. Recusar não altera nada.
3. O aplicativo clona os bytes atuais, inclusive edições não exportadas. Mantém streams, fontes, imagens, vetores, dimensões, caixas e giro; remove somente os recursos diagnosticados/autorizados. Não rasteriza páginas no PDF de saída.
4. A preparação é validada: nova análise estrutural sem incompatibilidades, texto e renderização de impressão antes/depois iguais, mesmas dimensões/rotação. Só então a cópia passa a ser a fonte da montagem.
5. A janela informa **Cópia independente preparada e validada para montagem**. Continuar com Step & Repeat/N-up e salvar pelo botão habitual. O documento do editor principal e seu estado pendente não são modificados nem marcados como exportados.
6. **Voltar ao original** restaura o snapshot original da janela e seu diagnóstico, sem recarregar/tocar o workspace. Cancelar fecha e descarta a cópia temporária. Reabrir captura novamente o workspace.

Se um recurso não puder ser preparado, os motivos ficam acessíveis e gerar/imprimir permanecem bloqueados. Não se oferece PDF vazio como resultado. Não foi acrescentada rasterização automática: casos que dependem de aparência de formulário/anotação ainda precisam de tratamento específico e aprovação em uma etapa própria.

A geometria, as marcas inteligentes, matrizes/escala e o motor Windows não foram reescritos. A prévia continua renderizando os mesmos bytes gerados. Cancelamento, versões, fila, worker próprio e Tentar novamente da .18 permanecem; comparação usa tarefas isoladas e limpeza limitada, sem destruir a prévia do editor.

## O que a validação comprova e seus limites

O comparador existente lê **todas as páginas**, texto/MCIDs e renderização de impressão integral a72DPI; exige pixels exatamente iguais na preparação, sem amostra ou redução silenciosa. Máximo8 milhões de pixels/8192px por lado;30s por etapa de leitura/render. Documentos acima disso ou com falhas de render ficam bloqueados. A prova pode demorar; Cancelar continua disponível. Parsing/save PDF-lib têm trechos síncronos e não são preemptados por timeout.

Essa prova não é certificação de ICC/CMYK/cores especiais, sobreimpressão, PDF/X ou comportamento de qualquer RIP. A cópia remove estruturas semânticas autorizadas; não é uma preservação de acessibilidade ou validade de assinaturas. Assinaturas/formulários com aparência continuam bloqueados neste fluxo.

A comparação da peça já escalada é um ensaio separado, não uma flexibilização do teste de preparação. Foram comparados pixels na área interna de uma peça, alinhada à mesma escala/posição a144DPI, excluindo2px de borda. A renderização depende do ambiente:

- Sintético com fonte incorporada, símbolos, imagens e transparências:572.792 componentes comparados, diferença máxima1/255 e nenhum componente diferindo mais de2 níveis.
- Real no **Electron instalado**:572.792 componentes, diferença máxima1/255, nenhum acima de2 níveis.
- Mesmo real no **Chrome**:970/572.792 componentes (0,169346%) acima de2 níveis, máximo38/255. A comparação foi inspecionada lado a lado sem identificar conteúdo ausente; a causa dessas diferenças não foi atribuída a um único fator. O diagnóstico registra a necessidade de conferência humana, não declara igualdade perfeita.

O limiar0,1% do relatório opcional serve somente como referência diagnóstica. O teste operacional de preparação continua exigindo igualdade exata antes/depois; não foi afrouxado. Conferir a montagem exportada no fluxo real da gráfica antes da produção.

## Testes executados

- **307 testes Node aprovados**:301 anteriores +6 novos de diagnóstico/segurança, confirmação/prova obrigatória, cópia independente, conservação de streams/fontes/caixas/giro, Step & Repeat/N-up, falha/cancelamento e bloqueios de ActualText essencial/formulários/AP/links imprimíveis/camadas/perfis.
- Chrome real com sintéticos semelhantes ao PowerPoint: estrutura marcada, ActualText redundante e essencial, links internos/externos sem Print, navegação, fonte TrueType incorporada, símbolos, imagens, vetores/transparências e páginas960×540/A4. A preparação de duas páginas conservou4.080.928 componentes RGBA e a sequência textual. ActualText essencial permaneceu bloqueado mesmo com o comparador real. Links sem marcação e marcação sem ActualText também passaram.
- Interface: motivo visível e detalhes recolhidos, confirmação recusada/aceita, preparação cancelada em voo, Voltar ao original, nova montagem e origem pendente preservada. Formulário com aparência no navegador mostrou motivo e manteve preparar/gerar bloqueados. Regressão das prévias (nove tentativas/oito cancelamentos, falha/retry, guias durante carregamento) preservada.
- .20 realmente instalada na mesma identidade/diretório/perfil Hold Test: todos esses fluxos; salvar a montagem preparada pela janela Windows produziu os mesmos bytes da prévia. Impressão **somente de sintéticos na Microsoft Print to PDF**, incluindo N-up preparado de uma folha. Nenhuma impressão física foi enviada.
- Arquivo real autorizado, tanto no Chrome quanto no Electron instalado:168 substituições redundantes comprovadas;2.073.600 componentes RGBA iguais na preparação; texto PDF.js preservado (1.855 caracteres nas strings dos itens; contagens de outros extratores não são diretamente equivalentes);960×540pt e rotação0 preservados. Step & Repeat/N-up gerados; fonte original/hash preservados. O real não foi enviado à fila de impressão.
- Electron smoke aprovado (abrir/miniaturas/giro, cópia seletiva, salvar/cancelar/falhar, abas, proteção, persistência e encerramento). ASAR instalado auditado:446 entradas. Sintaxe e `git diff --check` executados.
- Uma primeira tentativa instalada esgotou a espera do roteiro pela janela de impressão; a janela apareceu na captura. A causa não foi determinada. A repetição diagnóstica e a rodada final passaram, sem alterar o motor de impressão. Registrar se voltar a ocorrer; não há homologação universal dessa janela por este ensaio.

Na primeira rodada Node, foi corrigida uma chamada incorreta a lookupMaybe sem tipo no novo diagnóstico, antes da liberação. A proteção de anotações foi preservada e seu teste passou a conferir o código técnico, em vez do texto genérico da mensagem.

## Fixtures, evidências e privacidade

As fixtures permanentes são geradas por código, não PDFs de clientes. Os arquivos de teste gerados ficam ignorados em `.desktop-test/imposition20/browser/` e `installed/`:

- `PowerPoint sintetico para montagem.pdf`, `ActualText essencial bloqueado.pdf`, `Formulario bloqueado.pdf`.
- `Copia vetorial preparada.pdf`, `Montagem PowerPoint Step Repeat.pdf`, `Montagem PowerPoint N-up.pdf`, `Montagem preparada salva.pdf`, `Montagem preparada virtual.pdf`.
- `Origem e montagem comparadas.png`, capturas `estruturas-especiais.png` e `copia-montagem-preparada.png`; métricas em `compatibility.json`.

Os derivados reais e sua comparação ficam exclusivamente em subpastas locais ignoradas `browser/real/` e `installed/real/`, com nomes genéricos. Nenhum conteúdo, nome de cliente ou caminho privado é necessário no código. Para teste local opcional, definir `IMPOSITION_REAL_PATH` fora do repositório e executar `node tests/pdf-imposition-browser.mjs`; sem essa variável o teste real não é realizado. Não publicar essa pasta ou seus derivados.

## Arquivos alterados/novos

Novos de produção: `public/js/tools/pdf-imposition-compatibility.js` (diagnóstico/preparação transacional) e `pdf-imposition-compare.js` (adaptador do comparador existente).

Alterados: `pdf-imposition-engine.js` (erro estruturado antes da geometria), `pdf-imposition-ui.js` (aviso/ação/estado/retorno), `public/css/theme.css`, `electron-builder.hold-native-test.cjs` (.20).

Testes: novo `tests/pdf-imposition-compatibility.test.mjs`; novos helpers `pdf-imposition-modules.mjs`, `pdf-imposition-compatibility-browser.mjs`, `pdf-imposition-real-browser.mjs`, `pdf-imposition-pixel-compare.mjs`; adaptações de `pdf-imposition.test.mjs` e `pdf-imposition-browser.mjs`.

Documentação: este guia, `ESTADO_PROJETO.md` e `DISTRIBUICAO_DESKTOP.md`. Não foram alterados PDFWorkspace, redimensionamento, motor geométrico ou de marcas, validador ActualText existente, impressão nativa, drivers ou atualizador.

## Entrega e conferência manual

Instalador: `dist-desktop/hold-native-test/4PrintShops-Hold-Test-1.0.0-experimental.20-Windows-x64-Setup.exe`.

SHA256: `7B06663C265C80B20C3340A7D71E7C84024E41AB0AEEBF4BB89205FC5E1A0925`.

Mesma identidade/appId/NSIS/perfil Hold Test, atualizado localmente. Sem assinatura Authenticode comercial; instalação automática do updater ainda não homologada. Pacotes anteriores preservados, sem publicação.

1. Exporte trabalhos pendentes e feche Hold Test antes de instalar .20 no destino existente.
2. Abra um PDF marcado de teste e faça uma alteração sem exportar. Entre em Montagem / Imposição: deve aparecer o aviso e a ação apropriada, não somente um canvas branco.
3. Recuse a preparação: nada deve mudar. Aceite e aguarde a comparação; se compatível, a prévia deve aparecer e indicar cópia independente preparada.
4. Confira dimensões, fonte, símbolos, imagens/transparências, Step & Repeat/N-up e marcas; salve a montagem e reabra. Compare visualmente antes de produzir, especialmente no RIP.
5. Use Voltar ao original e Cancelar; o editor e suas alterações pendentes devem permanecer. Preparar/exportar a montagem não substitui Salvar cópia do PDF original.
6. Teste o sintético de ActualText essencial e o formulário: devem continuar bloqueados com motivo. Não tentar contornar removendo tags/anotações manualmente.
7. Para ensaio de impressão, comece pela Microsoft Print to PDF. O novo caminho com PDF real ainda requer aceitação pessoal do operador; não houve impressão física automática.

Pendentes: aceitação operacional da .20, conferência visual das diferenças entre renderizadores, fontes/recursos complexos e homologação de outros PDFs. Permanecem bloqueios descritos, limites de memória anteriores e pendências de cor/RIP, SMB/UNC e instalação automática do atualizador.

Conferência adicional do real: os18 streams originais tiveram seus hashes de conteúdo preservados na cópia preparada, incluindo os recursos de fontes/imagens/vetores existentes. Evidência numérica em `.desktop-test/imposition20/real-streams.json`; nenhum stream de cliente foi incluído no repositório.
