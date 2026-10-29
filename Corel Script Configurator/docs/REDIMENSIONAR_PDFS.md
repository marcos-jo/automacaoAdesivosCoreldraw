# Redimensionar páginas — primeira versão

A ferramenta fica no painel direito do Gerenciador de PDFs. As operações anteriores continuam disponíveis. O PDF original nunca é sobrescrito; redimensionar não inicia download.

## Uso

1. Importe um PDF ou prepare o documento usando união, giro, exclusão e reorganização.
2. Abra **Redimensionar**. Escolha A3, A4, A5, A6 ou Personalizado. Informe largura/altura em mm e a orientação. Aceita vírgula ou ponto decimal, sem separador de milhar. Limites: 0,1–5000 mm.
3. Escolha página atual, intervalo (`2, 4-8`) ou todas. Se a página atual estiver fora do intervalo, a prévia mostra a primeira página do intervalo e identifica seu número.
4. Escolha o modo e clique em **Pré-visualizar**. A área central mostra o conteúdo real transformado. O texto informa dimensões anteriores/novas e avisos. O resumo técnico e as miniaturas ainda descrevem o documento definitivo.
5. **Cancelar** restaura o documento anterior, incluindo edições anteriores. Trocar de ferramenta também descarta a prévia. Alterar uma configuração cancela/invalida o resultado anterior; gere uma nova prévia.
6. **Aplicar alterações** confirma a transformação e atualiza resumo, alertas, miniaturas e estado de exportação. A página originalmente selecionada é preservada. Continue editando e use **Salvar cópia** quando desejar.

Enquanto uma prévia está aberta, importação, navegação, edição e exportação ficam bloqueadas para evitar misturar estados. As configurações de redimensionamento e Cancelar continuam disponíveis. Durante a confirmação de Aplicar, as configurações também ficam bloqueadas.

Uma prévia não gera pendência de exportação. Se as páginas já tiverem as dimensões solicitadas, a ferramenta informa que não há alteração e não permite aplicar novamente. Cancelar continua disponível.

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
- **Bloqueados nas páginas selecionadas:** Widget/campos de formulário, outros tipos de anotação, aparências AP, ações AA e ações de link não suportadas. Não remove nem achata esses recursos silenciosamente.
- **Bloqueados no documento:** destinos internos/nomeados e links GoTo, marcadores Outlines, estrutura marcada StructTreeRoot, OpenAction, XFA e sinalizadores de assinatura SigFlags. São restrições conservadoras desta versão; não significam que esses PDFs não possam ser abertos ou usados nas outras ferramentas.
- Assinaturas não são preservadas como válidas por uma reescrita PDF. Essa versão não é garantia de fidelidade para todo recurso de qualquer PDF. Variantes complexas, metadados geométricos e recursos não cobertos pelos testes precisam de validação específica.
- Uma falha de render na aplicação restaura o workspace. Se também falhar a restauração visual, remove o canvas temporário, informa erro e desbloqueia os controles; não apresenta a prévia descartada como versão definitiva.
- Cancelamento verifica a geração da configuração, AbortSignal, revisão e documento; o carregamento PDF.js cancelado usa a destruição memoizada existente. Algumas fases de serialização/parsing não oferecem interrupção imediata e precisam terminar antes da limpeza.

O caminho rápido de giro/movimentação foi preservado. Redimensionar altera conteúdo e exige preparar uma nova fonte PDF.js; ao aplicar, reaproveita essa fonte já preparada. Dimensões e miniaturas das páginas afetadas são invalidadas, enquanto as demais podem ser reutilizadas. O cache continua limitado a 32 canvases e a janela virtual a 16 itens.

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
