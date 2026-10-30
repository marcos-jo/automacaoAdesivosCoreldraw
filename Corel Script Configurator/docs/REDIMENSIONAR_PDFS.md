# Redimensionar páginas e preparação para impressão

A ferramenta fica no painel direito do Gerenciador de PDFs. As operações anteriores continuam disponíveis. O PDF original nunca é sobrescrito; redimensionar não inicia download.

## Uso

1. Importe um PDF ou prepare o documento usando união, giro, exclusão e reorganização.
2. Abra **Redimensionar**. Escolha A3, A4, A5, A6 ou Personalizado. Informe largura/altura em mm e a orientação. Aceita vírgula ou ponto decimal, sem separador de milhar. Limites: 0,1–5000 mm.
3. Escolha página atual, intervalo (`2, 4-8`) ou todas. Se a página atual estiver fora do intervalo, a prévia mostra a primeira página do intervalo e identifica seu número.
4. Escolha o modo e clique em **Pré-visualizar**. A área central mostra o conteúdo real transformado. O texto informa dimensões anteriores/novas e avisos. O resumo técnico e as miniaturas ainda descrevem o documento definitivo.
   Se houver apenas recursos descartáveis de navegação/acessibilidade, o diagnóstico oferece **Continuar em modo de impressão**. Leia a lista e confirme: o descarte alcança o **documento inteiro**, mesmo ao redimensionar uma página. A confirmação gera uma prévia; ainda é possível cancelar sem alterar o workspace.
5. **Cancelar** restaura o documento anterior, incluindo edições anteriores. Trocar de ferramenta também descarta a prévia. Alterar uma configuração cancela/invalida o resultado anterior; gere uma nova prévia.
6. **Aplicar alterações** confirma a transformação e atualiza resumo, alertas, miniaturas e estado de exportação. A página originalmente selecionada é preservada. Continue editando e use **Salvar cópia** quando desejar.

Enquanto uma prévia está aberta, importação, navegação, edição e exportação ficam bloqueadas para evitar misturar estados. As configurações de redimensionamento e Cancelar continuam disponíveis. Durante a confirmação de Aplicar, as configurações também ficam bloqueadas.

Uma prévia não gera pendência de exportação. Se as páginas já tiverem as dimensões solicitadas e não houver preparação para impressão, a ferramenta informa que não há alteração e não permite aplicar novamente. Uma preparação que remova estruturas é uma alteração mesmo sem mudar dimensões. Cancelar continua disponível.

## Modos

| Modo | Escala do conteúdo | Comportamento |
| --- | --- | --- |
| Ajustar proporcionalmente (padrão) | Menor razão entre largura/altura de destino e origem | Preserva a área originalmente visível inteira; centraliza e pode deixar margens. |
| Centralizar sem alterar a escala | 100% do tamanho físico atual | Centraliza; formatos menores podem cortar conteúdo. |
| Preencher com corte | Maior razão entre largura/altura de destino e origem | Preenche o destino mantendo proporção; pode cortar conteúdo nas bordas. |

As margens resultantes não são sangria. O sistema não cria marcas, imposição, perfis de cor ou conversões RGB/CMYK.

## Formatos personalizados

Expanda **Formatos personalizados**, preencha nome e medidas acima e use **Salvar formato personalizado**. Não há medidas predefinidas para Formato 4 ou Formato 8.

- Selecionar um formato preenche suas dimensões. Selecionar/cadastrar não transforma o PDF.
- Para editar ou renomear, selecione o formato, altere os campos e use **Atualizar / renomear formato**.
- Excluir pede confirmação. A3/A4/A5/A6 são protegidos contra edição/exclusão; nomes equivalentes, sem distinção de maiúsculas e espaços excedentes, são rejeitados.
- Nome: 1–80 caracteres. Biblioteca: até 200 formatos personalizados.
- Persistência por `localStorage`, chave `4printshops.pdf.formatos.v1`, no navegador/origem atual. Dados inválidos são ignorados com aviso, sem impedir o editor. Um salvamento explícito substitui os dados inválidos.
- Armazenamento bloqueado ou quota esgotada: aviso e retenção em memória durante a vida da página, inclusive ao sair/voltar ao módulo PDFs. Recarregar/fechar a página pode perder esses cadastros não persistidos.
- Apagar dados do navegador, mudar de navegador, computador, porta/origem ou perfil não transporta a biblioteca.

`PDFFormatos` faz validação e CRUD; `PDFArmazenamentoFormatos` implementa apenas `ler()`/`gravar(valor)`. O serviço expõe `pronta` e operações assíncronas; aceita um adaptador cujo retorno seja uma Promise. A interface aguarda essas operações e utiliza a lista em cache. Um adaptador de API pode substituir o armazenamento sem alterar o motor de PDF. Autenticação, escopo por usuário/empresa, conflitos, sincronização, migrations e backup pertencem à futura camada de servidor; não estão implementados.

## Transformação e caixas

- Clona o documento completo em memória e prepara todas as páginas selecionadas antes de publicar qualquer alteração. Verifica versão/revisão e cancelamento novamente ao aplicar.
- Envolve os streams originais em estado gráfico, matriz afim e clipping; não converte o documento em imagens. Textos, vetores, imagens, recursos e operadores de cor existentes permanecem no PDF.
- Usa a interseção entre CropBox e MediaBox como área originalmente visível. O clipping explícito conserva áreas antes ocultas: aumentar o formato não revela conteúdo que estava fora do CropBox.
- Considera origem deslocada/negativa, rotações 0/90/180/270 e UserUnit. Incorpora a orientação ao conteúdo, normaliza Rotate para 0 e UserUnit para 1 nas páginas transformadas; novos giros continuam funcionando normalmente.
- MediaBox/CropBox passam ao formato solicitado, com origem 0,0. TrimBox/BleedBox/ArtBox **existentes** acompanham a transformação e são limitadas à área visível de origem/destino. Não cria caixas gráficas ausentes. Uma caixa que ficaria vazia bloqueia a operação, com indicação do problema.
- A existência de BleedBox não comprova sangria suficiente; cores, sangria, DPI e outros aspectos de pré-impressão continuam não verificados.

Referências conferidas: documentação de [PDFPage da pdf-lib](https://pdf-lib.js.org/docs/api/classes/pdfpage), código local da pdf-lib 1.17.1 e [PDF Reference da Adobe](https://opensource.adobe.com/dc-acrobat-sdk-docs/pdfstandards/pdfreference1.5_v6.pdf). Alterar tamanho da página e transformar conteúdo/anotações são responsabilidades diferentes; a implementação trata essas partes separadamente.

## Compatibilidade e limites

- Links externos URI/GoToR e notas Text/Popup simples, sem aparência AP: reposiciona retângulos, QuadPoints e bordas aplicáveis. Áreas cortadas deixam de ficar ativas fora da página; notas usam ícones definidos pelo leitor PDF.
- **Opcionais, com confirmação:** árvore de acessibilidade/estrutura marcada (`StructTreeRoot`, `MarkInfo`, referências `StructParent/StructParents` em páginas, anotações e XObjects), marcadores `Outlines`, destinos internos (`Dests` e `Names/Dests`), abertura em destino interno e destinos/ações GoTo de links simples. Cada categoria é identificada separadamente. A confirmação é específica à tentativa e não vale para outro documento, outra configuração ou revisão.
- A preparação remove as referências ativas dessas categorias em uma cópia transacional. Remove também `Tabs/S` e `PageMode/UseOutlines` associados às estruturas descartadas. Links internos conservam a anotação, retângulos e bordas; apenas sua navegação é retirada. Não elimina streams marcados `BDC/EMC`, propriedades de conteúdo, fontes/ToUnicode, imagens ou vetores. Objetos órfãos podem continuar nos bytes: não é higienização de informações confidenciais.
- **Preservados:** links externos URI/GoToR, outros nomes do catálogo, idioma, metadados, OutputIntents, OCProperties e recursos de produção existentes. A presença desses recursos não comprova conformidade PDF/X, PDF/UA, cor ou sangria. A árvore lógica pode conter atributos de layout/coordenadas que o motor não transforma; esta versão oferece descarte explícito, não promete reconstruir uma árvore acessível válida após a mudança geométrica.
- **Continuam bloqueados nas páginas selecionadas:** Widget, tipos de anotação além de Link/Text/Popup, aparências AP e ações complexas. Não apaga nem achata esses elementos para contornar o risco de perda de conteúdo impresso.
- **Continuam bloqueados no documento:** campos AcroForm, XFA, sinalizadores de assinatura SigFlags, ações adicionais de página/documento, scripts no catálogo de nomes e ações de abertura não tratadas. `ActualText` na árvore requer a validação específica abaixo; quando essencial ou não verificável, continua bloqueado. `ActualText` nos streams de conteúdo permanece intacto. O diagnóstico não oferece o modo de impressão quando identifica riscos não resolvidos.
- Assinaturas não são preservadas como válidas por uma reescrita PDF. Essa versão não é garantia de fidelidade para todo recurso de qualquer PDF. Variantes complexas, metadados geométricos e recursos não cobertos pelos testes precisam de validação específica.
- Uma falha de render na aplicação restaura o workspace. Se também falhar a restauração visual, remove o canvas temporário, informa erro e desbloqueia os controles; não apresenta a prévia descartada como versão definitiva.
- Cancelamento verifica a geração da configuração, AbortSignal, revisão e documento; o carregamento PDF.js cancelado usa a destruição memoizada existente. Algumas fases de serialização/parsing não oferecem interrupção imediata e precisam terminar antes da limpeza.

O caminho rápido de giro/movimentação foi preservado. Redimensionar altera conteúdo e exige preparar uma nova fonte PDF.js; ao aplicar, reaproveita essa fonte já preparada. Dimensões e miniaturas das páginas afetadas são invalidadas, enquanto as demais podem ser reutilizadas. O cache continua limitado a 32 canvases e a janela virtual a 16 itens.

Referência para marcação e texto substituto: [modelo de acessibilidade da Adobe](https://opensource.adobe.com/dc-acrobat-sdk-docs/library/accessibility/Access_DOM.html) e [PDF Reference](https://opensource.adobe.com/dc-acrobat-sdk-docs/pdfstandards/pdfreference1.5_v6.pdf). `ActualText` na estrutura e no conteúdo têm locais diferentes; remover a estrutura não autoriza remover operadores/propriedades de texto.

## Validação inicial da compatibilidade — 29/10/2026

- Suíte Node: 85 testes passaram, incluindo seis novos testes de motor/estrutura e um de integração da confirmação. As fixtures incluem 960×540 pt, marcação isolada, navegação/destinos, links, cinco entradas de recurso de fonte, TrueType incorporada, acentos/símbolos, seis imagens, vetores, transparências e Form XObject. A fonte Arial/DejaVu instalada é incorporada apenas aos PDFs temporários de teste; seu binário não faz parte do repositório.
- Verificados preservação byte a byte dos streams originais e recursos, limpeza consistente das referências estruturais, manutenção de nomes/metadados não relacionados, modos/formatos, 90°, páginas mistas, união/exclusão/reorganização anteriores, aplicação transacional, pendência e invalidação da confirmação. Formulários/AP, scripts e ActualText estrutural continuam bloqueados sem alteração parcial.
- Chrome headless: três comparações antes/depois da **preparação sem transformação geométrica**, cada uma com **2.073.600 componentes RGBA exatamente iguais**, e textos extraídos iguais com os símbolos. Separadamente, seis cenários de redimensionamento (três modos × 0/90°), confirmação nativa aceita/recusada, Cancelar, Aplicar, Salvar cópia e bloqueio de Widget passaram. Prévia/aplicação idênticas; exportação no mesmo teste de tolerância existente (<0,1% dos componentes diferindo mais de 2 níveis).
- A regressão de geometria já existente também passou: três modos × quatro rotações, centroides, 100/350 páginas, miniaturas, navegação, reordenação, giro, exclusão, extração, união, exportação, proteção e formatos persistentes.
- Na etapa inicial o PDF real não havia sido localizado; a validação posterior está registrada abaixo. A versão PDF.js usada nos testes Node emite avisos sobre build legacy/Math.sumPrecise; texto extraível foi conferido também no Chrome, onde ocorreu a validação visual. Não se certifica todo PDF de produção nem conformidade PDF/X/PDF/UA.

## ActualText: comprovação antes de remover a estrutura — 29/10/2026

O bloqueio por mera presença de ActualText foi substituído por uma prova conservadora. A extração global do PDF.js, sozinha, não basta: ela pode ignorar o texto substituto da árvore. Uma figura pode depender desse texto mesmo quando outra região já contém a mesma palavra.

1. O diagnóstico associa cada ActualText à sua página/referência e aos MCIDs indicados em `K` (número, array, elementos filhos ou MCR simples). A posição acompanha reorganizações; páginas já excluídas não voltam ao documento.
2. Uma segunda cópia temporária é preparada sem as referências estruturais, usando a mesma limpeza do modo de impressão. O workspace permanece intacto.
3. O verificador PDF.js lê cada página antes/depois, com conteúdo marcado e normalização automática desabilitada. Exige a mesma sequência de texto/quebras de linha e compara cada substituto com **seus próprios** MCIDs, sem procurar coincidências em outra região. Só normaliza Unicode NFC e espaços/quebras equivalentes; não ignora letras, símbolos, caixa ou pontuação. MCIDs ausentes/repetidos/ambíguos e substituições essenciais bloqueiam.
4. Compara integralmente a renderização de todas as páginas, uma de cada vez, a 72 dpi com intenção de impressão. Exige dimensões iguais e todos os componentes RGBA idênticos. Os canvases são apenas instrumentos de comparação; não substituem o conteúdo do PDF.
5. Somente após a prova aparece **Continuar em modo de impressão**, com confirmação explícita sobre a remoção de acessibilidade/semântica do documento inteiro. A confirmação executa a prova novamente na cópia atual. A autorização é efêmera, vinculada à cópia e ao diagnóstico; não há preferência/booleano persistido para ignorar a proteção.

Limites adicionais: páginas acima de 8 milhões de pixels ou 8192 pixels por lado em 72 dpi ficam bloqueadas, sem reduzir silenciosamente a resolução nem validar só uma amostra. Há limite de 30 s por leitura de página/texto/renderização. ActualText em MCR de XObject (`Stm/StmOwn`), árvores cíclicas/compartilhadas ou referências não identificáveis continuam bloqueados. Cancelamento interrompe o render, libera canvases e usa `PDFPreview.descartar()` para destruir cada tarefa uma vez. Parsing/save ainda têm os limites de cancelamento existentes. A prova integral pode demorar em documentos extensos; não é executada no fluxo de giro/movimento, nem para PDFs sem ActualText estrutural.

**Resultado com o arquivo real:** localizado `(PROPOSTA) SÃO JOSÉ.pdf` em Downloads, sem adicioná-lo ao Git. As duas cópias locais encontradas tinham o mesmo SHA-256. No Chrome headless, as **168 entradas ActualText** foram comprovadas redundantes nos seus blocos. Preparação: **2.073.600 componentes RGBA idênticos**, incluindo comparação adicional com os bytes originais, e extração antes/depois igual. A4/Ajustar: importar → pré-visualizar → recusar/aceitar → cancelar → aceitar/aplicar → Salvar cópia → reabrir passou. Prévia/aplicação idênticas; exportação/reabertura conferidas com a escala efetiva do render principal, mantendo o limiar existente. Os hashes de todos os streams originais foram preservados no PDF exportado; o SHA-256 do arquivo de origem permaneceu intacto. Capturas de origem, prévia A4 e reabertura também foram inspecionadas visualmente.

Na solicitação seguinte, o usuário informou que a correção PowerPoint foi aprovada em testes reais. A etapa de assinaturas preservou esse caminho e repetiu a regressão com o mesmo arquivo. O novo Resolver é um fluxo separado e explicitamente rasterizado para cópias de impressão: [ASSINATURAS_E_COPIAS_IMPRESSAO.md](ASSINATURAS_E_COPIAS_IMPRESSAO.md). Ele não substitui nem enfraquece os bloqueios/validações de ActualText do redimensionamento vetorial.

Nesta extração PDF.js, foram 1855 caracteres nas strings dos itens e **1892 incluindo as quebras de linha**, conservados após A4. A contagem externa de 1921 caracteres era de outro extrator. PyMuPDF/Poppler não estavam disponíveis; não se afirma ter reproduzido sua contagem nem seus testes. A igualdade foi verificada com o mesmo extrator em cada comparação.

**Testes:** 94 testes Node passaram (85 anteriores + oito testes específicos + um de cancelamento pela interface). Chrome headless passou com ActualText redundante, essencial, texto em vetores, mistura e símbolos; os três casos essenciais permaneceram bloqueados mesmo com extração global igual. Regressões anteriores de 100/350 páginas e demais ferramentas passaram. A biblioteca de formatos e a proteção contra perda de alterações foram preservadas. A aprovação manual desta correção pelo operador, inclusive em outro leitor/RIP, permanece recomendada; não equivale a certificação de todos os PDFs PowerPoint.

Para repetir sem incluir clientes no repositório: definir `PDF_ACTUALTEXT_REAL` com o caminho local do PDF e executar `node tests/pdf-browser-smoke.mjs`. Sem essa variável a parte real é explicitamente informada como não executada. PDFs derivados e capturas ficam em diretório temporário, nunca em fixtures Git. O teste registra apenas métricas/hashes e usa um nome genérico na interface.

Exemplo no PowerShell:

```powershell
$env:PDF_ACTUALTEXT_REAL = 'C:\caminho\fora-do-repositorio\arquivo.pdf'
node tests/pdf-browser-smoke.mjs
```

## Aceitação com arquivos reais

Em 29/10/2026, o usuário confirmou que todos os testes manuais realizados com PDFs reais da gráfica foram aprovados, incluindo PDF/X-1a, PDF/X-3, PDF/X-4, Acrobat 9, textos, curvas e páginas com dimensões diferentes. Esse registro se refere aos arquivos/cenários relatados; os limites de compatibilidade acima permanecem. O roteiro abaixo pode ser reutilizado em futuras alterações.

1. Usar cópias de PDFs reais da gráfica, de preferência com elementos reconhecíveis perto de cada borda.
2. Conferir A4→A5, A5→A4, A3→A4, personalizado `123,45 × 234,56` e paisagem/retrato, nos três modos. Verificar os avisos de corte antes de aplicar.
3. Testar página atual, intervalos e todas; incluir documentos mistos e previamente girados. Conferir número da página em prévia e preservação da seleção ao cancelar/aplicar.
4. Cancelar uma prévia limpa: não deve aparecer pendência. Repetir após girar/excluir/reorganizar: edições anteriores e aviso de pendência devem permanecer.
5. Durante o preparo, alterar medidas ou cancelar; nenhum resultado antigo deve reaparecer. Trocar de ferramenta e voltar; controles devem responder e não duplicar ações.
6. Aplicar, conferir dimensões/miniaturas/alertas, girar, excluir, unir e mover novamente. Salvar cópia, reabrir em outro leitor e conferir conteúdo, orientação, caixas e recursos importantes. Extrair também uma página transformada.
7. Conferir links externos/notas simples; testar uma página com formulário/recursos bloqueados e verificar mensagem, ausência de mudança parcial e preservação das demais funções.
8. Cadastrar um formato, fechar/reabrir o navegador, selecionar, editar/renomear e excluir cancelando/confirmando. Não editar os padrões. Testar entradas inválidas e nomes duplicados.
9. Repetir navegação/arraste/rolagem e operação em 100/350 páginas. PDFs complexos podem consumir tempo/memória superiores aos sintéticos de teste. Não há salvamento automático do documento.
10. Abrir uma cópia do PowerPoint real, escolher A4/Ajustar e Pré-visualizar. Conferir o diagnóstico de acessibilidade; recusar o modo de impressão e verificar que nada mudou. Repetir, aceitar e cancelar a prévia; conferir conteúdo/dimensões originais. Aceitar novamente, aplicar, exportar e reabrir. Comparar textos selecionáveis/pesquisa, símbolos, fontes, imagens, vetores e transparências. O original deve continuar intacto.
11. Repetir com marcadores/destinos/links internos e externos: conferir a lista de descarte do documento inteiro; links externos e bordas devem permanecer. Testar modos com corte conscientemente, página girada, páginas mistas e edições anteriores. Uma anotação com AP/formulário deve informar bloqueio sem oferecer descarte visual.
