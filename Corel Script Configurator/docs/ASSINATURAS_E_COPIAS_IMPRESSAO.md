# Assinaturas e cópias para impressão — 29/10/2026

O Gerenciador identifica estruturas de assinatura ao importar/adicionar PDFs e oferece **Resolver** no resumo técnico. A proposta de impressão é preparada localmente, sem upload e sem alterar o arquivo original. Não é remoção de senha, quebra de criptografia ou verificação jurídica/criptográfica.

## Diagnóstico

A leitura ocorre na origem **antes do primeiro `save()`**, com pdf-lib 1.17.1. Examina campos `/FT /Sig`, herança `/Parent`/`Kids`, valores `/V`, `/Contents`, `/ByteRange`, widgets `/AP`, `/Rect`, flags de impressão e `/Perms`. Dicionários órfãos e outros indícios são tratados separadamente. `/SigFlags` sozinho não comprova assinatura; zero não cria alerta. Imagem/rubrica/texto desenhado não é considerado assinatura digital.

| Estado | Interpretação |
| --- | --- |
| Assinatura digital detectada | Campo/catálogo associado a valor com conteúdo não vazio e ByteRange estruturalmente plausível. **Não** comprova certificado, integridade, identidade ou validade jurídica. |
| Campo de assinatura vazio | Campo Sig sem valor; não é contado como assinatura aplicada. Pode ter uma aparência de campo a preservar. |
| Detecção inconclusiva | Valores incompletos, referências órfãs, estruturas malformadas/dinâmicas ou indícios sem associação suficiente. Não significa arquivo livre de assinaturas. |
| Nenhuma identificada | Nenhuma estrutura reconhecida nas origens analisadas; também não é garantia universal de ausência. |

Widgets com área visível são associados às páginas. AP utilizável e flag de impressão são registradas separadamente. Uma assinatura com retângulo visível e AP ausente não é chamada de invisível. Sem widgets/área visível, a assinatura pode ser estruturalmente invisível; não se inventa a página que precisa de intervenção.

A assinatura protege os bytes da **revisão assinada**, não somente a página onde está o widget. “Páginas necessárias” identifica aparências a rasterizar; não limita o alcance criptográfico. Por isso a cópia remove as estruturas de assinatura do documento inteiro, mesmo quando apenas uma página recebe imagem. Nas múltiplas assinaturas incrementais, cada CMS é verificado sobre sua própria revisão.

O histórico guarda os diagnósticos dos arquivos de origem e os relaciona às identidades das páginas. **Informações** mantém contagens, páginas originais e limites da detecção; o alerta mostra a posição atual após reorganização/exclusão/união. A versão editável já passa por reserialização, mesmo para visualização: o diagnóstico da origem não declara que a assinatura continua válida na versão atual. Após Resolver, o histórico permanece, mas a cópia é identificada separadamente como preparada e verificada estruturalmente sem assinatura ativa. Unir outra origem assinada reativa o alerta.

## Operação

1. Importe o documento; veja o aviso superior compacto. **Ver detalhes / Ocultar detalhes** expande/recolhe o diagnóstico completo, também disponível em **Informações**. A lista do banner mostra até três páginas e a quantidade restante; a lista completa fica nos detalhes. Não há modal automático.
2. Clique **Resolver**. Isso só abre opções, sem alterar o workspace ou baixar arquivo.
3. Mantenha **Rasterizar somente páginas necessárias** (padrão) ou escolha **Rasterizar documento inteiro**. Escolha **300 DPI** (padrão) ou **600 DPI**. **Mostrar detalhes técnicos** reúne explicações, limites, diagnóstico e, após preparar, páginas rasterizadas/preservadas e avisos. Progresso, exigência de páginas manuais e erros permanecem visíveis. Os controles de conferência e aplicação aparecem quando a prévia estiver pronta.
4. O seletivo inclui páginas com aparência visual de assinatura e outras anotações imprimíveis. Informe páginas adicionais/manuais quando necessário. Sem páginas determinadas, indique-as ou escolha inteiro; não há seleção automática fictícia para assinaturas invisíveis.
5. **Preparar e pré-visualizar** gera uma cópia temporária, informa progresso e verifica o resultado. O documento central e suas alterações acumuladas permanecem intactos.
6. Confira todas as páginas rasterizadas usando os controles de revisão. O contador significa páginas **exibidas**, não prova automática de conferência humana. A comparação “antes do preparo” mostra o workspace atual pelo PDF.js, **não os bytes do original assinado**. Compare também com o original em leitor confiável, considerando giros, cortes e edições intencionais anteriores. Role a janela para visualizar a página e a confirmação.
7. Se faltarem elementos ou aparecerem quadrados pretos, **cancele**. Não confirme para tentar forçar a correção. Se estiver de acordo, marque a conferência e clique **Aplicar cópia ao documento**.
8. A confirmação explicita: “Será criada uma cópia para impressão. A nova versão não preservará a validade da assinatura digital. O documento original permanecerá intacto.” Recusar mantém a proposta, sem aplicar.
9. Aplicar confirma a revisão do workspace, atualiza visualização/resumo/miniaturas e marca alterações não exportadas. Use **Salvar cópia** para baixar `nome_IMPRESSAO.pdf`. Um original já com esse sufixo gera nome distinto (`..._COPIA_IMPRESSAO.pdf`). Não existe download automático em Resolver, preparar ou aplicar.

Cancelar/fechar/ESC descarta a proposta e mantém as edições anteriores. Alterar a configuração invalida a prévia e a conferência. Aplicar é uma transição protegida: aguarde sua conclusão. Depois, giro, organização, extração, união e redimensionamento continuam disponíveis. Rasterizar novamente em maior DPI não recupera detalhes perdidos; reimporte o original quando precisar de outra resolução real.

Após aplicar, o banner mostra **Cópia para impressão preparada / Sem assinatura ativa**, com **Refazer preparo** no lugar de Resolver. O diagnóstico do original permanece identificado como histórico nos detalhes e em Informações. Refazer abre a mesma preparação, sem modificar o documento; é necessário informar páginas ou escolher inteiro quando não há páginas visuais pendentes. Cancelar preserva a cópia atual. Outra origem assinada adicionada ao documento reativa **Resolver**, contando somente as origens ainda pendentes no resumo. Os detalhes do modal iniciam recolhidos em cada abertura; o banner recolhe ao trocar a origem ou passar de pendente para preparado, preservando a escolha ao navegar/girar/reorganizar.

## Aparência e reconstrução

- PDF.js local **6.4.299**, `intent: "print"`, `AnnotationMode.ENABLE`; não se usa uma renderização que omita formulários. A prévia da proposta usa o mesmo modo. Widgets de assinatura são tratados pelo renderizador como anotações; o conteúdo da página sozinho não basta.
- Verifica AP, associação de anotação ao operator list, presença de operadores de pintura e pixels de anotação na região visual de assinatura, comparando com anotações desabilitadas. AP ausente, vazia ou não disponibilizada bloqueia. Aparência visual sem flag de impressão também bloqueia, em vez de desaparecer silenciosamente.
- São salvaguardas do **mesmo motor**; não comprovam que o PDF.js reproduz o Acrobat ou o RIP. Outros elementos sobrepostos, fontes/recursos não suportados e bugs do renderizador exigem conferência externa. Um retângulo preto poderia ser conteúdo legítimo ou erro; não há heurística que o substitua silenciosamente.
- Cria um **PDF novo**, em contexto novo, sem copiar catálogo/AcroForm/Perms/DSS de assinatura. Antes de copiar páginas não rasterizadas, remove referências interativas da cópia transacional. Todas as aparências imprimíveis entram no plano de rasterização. Referências não imprimíveis/invisíveis/sem área não são transportadas e são indicadas nos avisos. Não apaga os elementos de aparência impressa apenas para remover o campo.
- Páginas preservadas são copiadas juntas, compartilhando recursos quando possível. Permanecem com conteúdo, fontes, vetores, imagens, transparências e caixas, sem nova rasterização. Já poderiam ser imagens no original; a ferramenta não afirma que toda página preservada é vetorial. OutputIntents são transportados no seletivo como recurso de cor, sem certificar sua conformidade.
- Nas páginas rasterizadas, captura a orientação real de impressão, inclusive widgets NoRotate. Incorpora rotação e UserUnit à imagem/caixas e normaliza Rotate=0 e unidade=1. Mantém tamanho **físico**, sequência, Media/Crop e caixas gráficas existentes; valores numéricos brutos de largura/altura podem trocar por causa de uma rotação de 90°. Não cria sangria nem marcas.
- RGB opaco sobre branco, compressão Flate **sem JPEG**. Faixas de 32 linhas evitam um PNG/RGB completo intermediário. Isso não preserva separações CMYK, cores especiais, overprint ou certificações PDF/X/PDF/UA; transparências são achatadas. A saída pode diferir no RIP. Páginas rasterizadas perdem texto pesquisável/selecionável e vetores, conforme aviso antes de preparar.
- Verifica todos os objetos da cópia, incluindo referências órfãs/dicionários diretos, impedindo FT Sig, tipos Sig/DocTimeStamp, ByteRange, SigFlags, AcroForm ou Perms remanescentes. A proposta reabre no PDF.js; compara cada página não rasterizada a 72 DPI com pixels exatos. Para imagens, descomprime RGB por faixas e compara a renderização da saída no DPI escolhido: menos de 0,1% dos componentes com diferença maior que 2 níveis. A prévia e a exportação também foram comparadas nos testes.

## Limites e alternativas

- OCProperties/camadas dependentes do catálogo bloqueiam o seletivo e oferecem **inteiro** como alternativa explícita. Diferença na página preservada também bloqueia e sugere inteiro; não muda a estratégia sozinho.
- XFA, NeedAppearances, NeedsRendering dinâmico, cálculos de campos, scripts/ações adicionais e cadeias de ação não tratadas bloqueiam. Uma assinatura/AP não renderizável não é recuperada pela opção inteiro do mesmo motor; é necessário outro leitor/processo confiável.
- Widgets diretos sem identificador estável, recursos incomuns e PDFs malformados podem não passar na associação/validação. Isso gera limite explícito, sem dizer que o documento está livre de assinatura ou corrigido.
- Páginas rasterizadas: até **36 milhões de pixels** e **12000 por lado**. Imagens comprimidas acumuladas: **256 MiB**. Páginas preservadas: comparação integral limitada a 8 milhões de pixels em 72 DPI. Não reduz resolução nem valida apenas amostras sem avisar.
- Até 60 s por leitura de página/operator list/render/fluxo de compressão. Operações PDF.js e canvases são cancelados e liberados; destruição de loaders é memoizada pelo dono (`PDFPreview`). Parsing, cópia/serialização da pdf-lib e alguns trechos nativos não se interrompem imediatamente; ao terminar, a geração cancelada é descartada.
- Um documento/candidato inteiro ainda existe em memória. O limite de pixels considera margem para decodificação/canvas; **não é medição ou garantia do pico total de RAM do Chrome**. Reduza resolução ou divida o trabalho se necessário. O limite bloqueia A3/600 DPI, sem tentar alocar seu canvas.

| Formato / DPI | Pixels | Memória exata do canvas RGBA | Resultado no Chrome |
| --- | --- | --- | --- |
| A4 / 300 | 2481 × 3508 | 33,2 MiB | Render, preparação e exportação passaram |
| A3 / 300 | 3508 × 4961 | 66,4 MiB | Passaram |
| A4 / 600 | 4961 × 7016 | 132,8 MiB | Passaram |
| A3 / 600 | 7016 × 9922 | 265,6 MiB (cálculo; não alocado) | Bloqueado pelo limite de 36 milhões de pixels |

## Testes e reprodução

- Referência inicial: **94 testes passaram**. Suíte final: **112 passaram**, incluindo detector, valores vazios/inconclusivos, metadados antes do save, assinatura só na última página, histórico após edições, transação/revisão, reconstrução sem referências de assinatura, nomes distintos, limites, reentradas e os novos testes criptográficos.
- Chrome completo passou com regressões anteriores e o **PowerPoint real aprovado** (`PDF_ACTUALTEXT_REAL`), preservando as 168 verificações ActualText e os testes de redimensionamento. Os casos estruturais simulados continuam para injetar erros/estruturas inválidas; eles foram complementados pela bateria de CMS reais descrita abaixo.
- Casos novos no navegador: visível/invisível/vazia/duas/última página, Stamp e widgets, seletivo e inteiro, mistura de páginas preservadas/rasterizadas, transparências/PNG 1024²/TrueType/símbolos, NoRotate/270°/Crop deslocado/UserUnit/caixas, comparação com documento antes do preparo, confirmação recusada/aceita, cancelamento, AP ausente/vazia/não imprimível, erro de render injetado com retry, reabertura, nenhum download automático e **pixels idênticos entre prévia e exportação**.
- Giro, exclusão, movimentação, extração, redimensionamento e nova união assinada após Resolver passaram. Nova origem assinada reativou o alerta e cancelar o segundo preparo preservou as edições. Cancelar preparo com 350 páginas manteve documento e controles responsivos.
- Benchmark de regressão 20/100/350 páginas confirmou giro **0 save / 0 getDocument**, movimento **1 / 0** e exclusão **1 / 1**. Medianas em 350 páginas nesta execução: 16,9 / 214,8 / 217,2 ms. Não é comparação controlada nem garantia para produção; as novas rotinas de rasterização não participam dessas operações.
- Sintaxe dos JS/MJS pendentes e `git diff --check` passaram. Nenhuma dependência instalada. `npm test` continua sendo placeholder; a suíte é executada explicitamente.

```powershell
node --test tests/pdf-real-signatures.test.mjs tests/pdf-signatures.test.mjs tests/pdf-print-copy.test.mjs tests/pdf-actualtext.test.mjs tests/pdf-workspace.test.mjs tests/pdf-integration.test.mjs tests/pdf-thumbnails.test.mjs tests/pdf-protection.test.mjs tests/pdf-visual-document.test.mjs tests/pdf-formats.test.mjs tests/pdf-resize.test.mjs tests/pdf-print-compatibility.test.mjs
node tests/pdf-browser-smoke.mjs
node tests/pdf-performance.mjs assinaturas-regressao
```

Para executar apenas o bloco de assinaturas no Chrome, definir `$env:PDF_TESTE_ASSINATURAS_APENAS='1'`; nesse modo as outras regressões são explicitamente informadas como **não executadas**. Remover a variável antes de rodar tudo. `PDF_ACTUALTEXT_REAL` continua opcional para o teste real PowerPoint; não aponta para um exemplo assinado.

## Certificados de teste e CMS criptograficamente reais

O usuário confirmou não ter um PDF problemático salvo e autorizou a geração de exemplos assinados com certificados de teste. Python/pyHanko não estavam disponíveis. Foi usado o **OpenSSL já instalado com o Git**, sem instalar dependências e sem alterar configuração/credenciais do Git/GitHub. Os executáveis são localizados também por `OPENSSL_PATH` ou PATH, para reprodução em outra máquina.

`tests/helpers/pdf-real-signatures.mjs` cria um certificado autoassinado **TEST ONLY**, RSA-2048/SHA-256, validade de um dia, CA:FALSE e uso de assinatura. Não instala nada no repositório de certificados do Windows. A chave privada fica em diretório temporário exclusivo e é removida após a geração, inclusive em caminhos de falha; os testes conferem sua ausência. Nunca há chave/certificado de cliente ou certificado de produção.

As fixtures reservam Contents, calculam ByteRange real e incorporam CMS detached com `adbe.pkcs7.detached`. Múltiplas assinaturas usam atualizações **incrementais** de objeto/xref/trailer com Prev; não reescrevem os bytes já assinados. São utilitários de teste para PDFs sintéticos controlados, **não um assinador de produção**.

| Fixture | Páginas | Assinaturas CMS | Aparências / seleção no Resolver |
| --- | --- | --- | --- |
| `real-visivel.pdf` | 3 | 1 | Página 1; outras 2 preservadas |
| `real-invisivel.pdf` | 3 | 1 | Nenhuma página visual inventada; teste seleciona 2 manualmente |
| `real-ultima.pdf` | 8 | 1 | Somente 8 rasterizada; outras 7 preservadas |
| `real-multiplas.pdf` | 4 | 2 em revisões incrementais | Páginas 1 e 4; outras 2 preservadas |
| `real-vazio.pdf` | 3 | 0 | Campo vazio com AP na página 3; não rotulado como aplicado |
| `real-misto.pdf` | 6 | 2 em revisões incrementais | Páginas 2 e 5; outras 4 preservadas |

O OpenSSL verificou **sete CMS** nos cinco PDFs assinados, antes e depois dos fluxos no app. `cms -verify -binary -noverify` verifica assinatura/digest, sem validar confiança da cadeia do certificado autoassinado. Um teste altera um byte protegido, mantém o PDF legível e exige falha criptográfica. Outro demonstra que reserializar pelo workspace não preserva a assinatura do original, enquanto o arquivo original permanece verificável e o diagnóstico de origem continua disponível.

No Chrome, os seis cenários passaram por Resolver, prévia, conferência/confirmar, Salvar cópia e reabertura dos bytes baixados. As cópias tiveram zero estruturas de assinatura ativa; os originais conservaram hash e CMS verificáveis. As páginas não afetadas mantiveram fontes/conteúdo vetorial. Também passou a igualdade de pixels da prévia/exportação. Isso comprova esses documentos de teste, **não confiança em certificados, validade jurídica, política DocMDP, LTV/PAdES ou compatibilidade com todo assinador**. A aplicação continua fazendo apenas detecção estrutural; a verificação OpenSSL existe exclusivamente nos testes.

O gerador imprime o diretório `%TEMP%/4printshops-certificados-teste-.../`, com PDFs, certificado público e relatório `verificacoes.json`, sem a chave privada. Essas fixtures geradas não são versionadas. Para repetir a geração/verificação basta `node --test tests/pdf-real-signatures.test.mjs`. Sem OpenSSL, a bateria informa a dependência e falha explicitamente, sem ser contada como aprovada nem instalar ferramentas automaticamente.

## Aceitação manual pendente com PDFs assinados reais

O usuário confirmou que não possui no momento um exemplar da falha **Acrobat correto × PDF.js incorreto**. Não foi testada impressão física, Acrobat ou RIP; não há promessa de recuperar conteúdo que o PDF.js omite. Os certificados/CMS reais de teste ampliam a validação, mas não reproduzem automaticamente o defeito de atendimento. A confirmação automática dos botões simula o fluxo, não substitui a aprovação do operador. Quando chegar um arquivo problemático da gráfica, fazer a validação complementar antes de considerar esse caso resolvido.

1. Trabalhar com cópias de arquivos reais e preservar o original assinado em local separado. Conferir contagem/campos vazios, AP, páginas originais e atuais, invisíveis e estado inconclusivo em Informações. Uma imagem/rubrica não deve gerar o alerta digital.
2. Abrir Resolver, cancelar sem preparar, depois preparar seletivo em 300 DPI. Comparar cada página necessária com o original em Acrobat/leitor confiável, em modo de impressão, levando em conta edições intencionais. Conferir assinaturas, textos pequenos, símbolos, imagens, transparências e bordas. Se faltar algo, cancelar.
3. Recusar a confirmação e verificar que nada foi aplicado. Confirmar somente após conferência; Salvar cópia `_IMPRESSAO`, reabrir em outro leitor/programa de destino e conferir tamanho físico, orientação e sequência. Imprimir uma prova no RIP com escala controlada antes de produzir.
4. Testar inteiro/600 DPI em formatos dentro dos limites, assinaturas invisíveis com seleção manual, formulários/anotações e os bloqueios explícitos. A3/600 deve bloquear sem alterar nada. Não tentar “resolver” AP ausente repetindo a mesma renderização.
5. Repetir após giro/reorganização/exclusão/união/redimensionamento. Testar pendências/aviso ao sair, cancelamento durante preparo, repetição e reentrada. Nova união assinada exige nova conferência. Nunca versionar PDFs de cliente, certificados, chaves ou capturas contendo conteúdo privado.
