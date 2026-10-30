# Automação de Adesivos / CorelDRAW preservada

Em 30/10/2026, Adesivos foi retirado da aplicação principal do 4PrintShops. As regras de montagem variam entre máquinas; a futura ferramenta configurável de montagem será planejada dentro do Gerenciador de PDFs.

## Inventário e dependências

Os arquivos foram conservados nos caminhos atuais para preservar imports relativos, scripts avulsos e histórico. A preservação não representa um módulo ativo nem uma interface alternativa já pronta.

| Componente | Arquivos preservados | Dependências / situação |
| --- | --- | --- |
| Script original da máquina | `modelos/Adesivo 33x48_Modelo.js` | Original CorelDRAW, não alterado. Não usar como regra universal de imposição. |
| Adaptador do configurador | `public/js/app/adhesives.js` | Usa toolbar, renderer, state e API de scripts. Não importado nem registrado por `main.js`. |
| Editor de parâmetros | `public/js/toolbar.js`, `renderer.js`, `state.js` | Código legado preservado, não carregado pela aplicação principal. `events.js` é um arquivo vazio, também preservado. |
| Navegação anterior | `public/js/sidebar.js` | Não é montada pelo shell atual. O botão Adesivos também foi retirado daqui para evitar reintrodução pela navegação antiga. |
| Leitura e representação de modelos | `src/repository.js`, `parser.js`, `model.js` | Repository lê `modelos/`; parser cria o modelo de controles. Carregados pelo servidor somente quando uma rota legada é solicitada. |
| Gerador CorelDRAW | `src/generator.js` | Preservado, usado pelo teste avulso; não faz parte do processamento de PDFs. |
| Verificação avulsa | `teste-generator.js`, `teste2.mjs` | Mantidos na raiz. O gerador produz texto em memória/console sem sobrescrever o modelo. |

Antes da retirada, o commit `e78fc91` registra o configurador integrado à aplicação. Para estudar ou recuperar essa interface no futuro, consultar esse histórico e o inventário acima; reativação exige decisão explícita e validação própria. Não há recuperação automática por cartões, menus ou eventos antigos.

## Servidor e serviços compartilhados

- `app.js` e `src/server.js` continuam fornecendo o HTTP local e os arquivos estáticos usados por PDFs e QR Codes.
- As rotas de leitura `/api/scripts` e `/api/scripts/:nome` foram preservadas por compatibilidade com o trabalho anterior. Não são solicitadas pela interface atual; `repository.js`/parser/model não integram mais a importação inicial do servidor. O `import()` sob demanda usa o cache de módulos do Node.
- Não foram removidas bibliotecas nem alterados os contratos de PDF.js/pdf-lib, o workspace, os menus genéricos, as abas ou o serviço de exportação.
- A aplicação usa somente o registro de módulos de `public/js/main.js`: PDFs e QR Codes. Incluir uma ferramenta real nesse registro gera seu cartão; não criar espaços clicáveis fictícios. Comandos com um ID não registrado são recusados, preservando a aba atual.

## Montagem de páginas / Imposição — somente planejamento

A futura ferramenta pertence ao módulo PDFs, com folha padrão/personalizada, repetição/distribuição, margens/espaçamentos, escala proporcional, marcas de corte, tratamento da sangria existente, prévia antes da aplicação e configurações reutilizáveis independentes da máquina.

Não existem controles nem motor novos de imposição nesta retirada. A implementação futura deverá definir a transformação vetorial, as caixas/sangrias e a fidelidade de impressão, acumulando o resultado no PDFWorkspace e exportando somente quando solicitado. O script CorelDRAW preservado serve como referência histórica, sem reutilização automática de suas regras.

## Conferência

Validar Início → PDFs → QR Codes, fechamento/reabertura, proteção de alterações e exportação. O Chrome headless verifica também a ausência de pedidos dos módulos legados/API na UI, a recusa de comandos antigos e, por solicitações explícitas separadas da interface, as respostas das rotas legadas. Os testes avulsos e a comparação SHA-256 confirmam a conservação do trabalho CorelDRAW; resultados da tarefa em `ESTADO_PROJETO.md`.
