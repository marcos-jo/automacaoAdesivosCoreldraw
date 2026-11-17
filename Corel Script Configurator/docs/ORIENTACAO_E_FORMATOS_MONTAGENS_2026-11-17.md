# Orientação e formatos nas montagens — 17/11/2026

## Revisão autorizada e checkpoints separados

- **A:** `d1ad3f0a93012503d1cba5c1d68881d9622f4b46`, Corrige orientacao de folhas personalizadas no Step Repeat. Cinco arquivos: `public/js/tools/pdf-imposition-ui.js` (somente orientação), `public/js/tools/pdf-sheet-orientation.js`, `tests/pdf-imposition-preview.test.mjs`, `tests/pdf-repeat-interface-browser.mjs` e `tests/pdf-sheet-orientation.test.mjs`.
- **B:** este documento integra o checkpoint filho **Compartilha formatos personalizados nas montagens PDF**. Sete arquivos de código/testes descritos na Parte B abaixo e cinco documentos: PLANEJAMENTO, CONTEXTO_ATUAL, ESTADO_PROJETO, PROPOSTA_FORMATOS_COMPARTILHADOS e este relatório. O hash de B é consultável pelo Git; não se incorpora um hash autorreferente ao commit.
- Nenhum código/teste foi alterado para separar os commits. A UI compartilhada foi preparada por trechos somente no index; o checkout A+B ficou intacto. O checkpoint A foi testado isoladamente sem os módulos/integração B: **14/14** (orientação/prévia), sintaxe dos cinco arquivos e roteiro de interface Edge aprovados.
- A+B revalidados: **37/37** nos cinco módulos direcionados de orientação, formatos, catálogo, Redimensionar e prévia; roteiros Edge de interface e catálogo aprovados. Sintaxe de 11 arquivos, comparação dos hashes, staging e diff checks aprovados. A suíte completa **567/567 é a execução anterior**, não repetida aqui. Electron não foi reexecutado nem homologado.
- Logs e snapshots em `C:\Dev\4PrintShops_Auditoria_Formatos_2026-11-17\checkpoints`, incluindo `A-unitarios-previa.log`, `A-interface-Edge.log`, `AB-direcionados.log`, `AB-interface-Edge.log` e `AB-catalogo-Edge.log`. Os documentos anteriores foram preservados; os resultados originais abaixo permanecem como histórico.

## Próxima validação Electron — plano, não executado

1. **Preparação isolada:** após revisão, usar Electron 45.1.0 e o código dos dois checkpoints em um harness que injete um `profile` temporário explícito em `startDesktop`, mecanismo já existente. Verificar o caminho efetivo de `userData` antes de qualquer cadastro. Manter a origem `app://4printshops`, o sandbox e as proteções; bloquear impressão e atualização no harness. Fechar o Hold Test antes do ensaio para evitar interferência do controle de instância única. Resolver primeiro a inicialização/endpoint de depuração que falhou no ensaio anterior; não aumentar timeout ou usar o perfil operacional como alternativa.
2. **Dados sintéticos:** em perfil novo, sem copiar sessões/credenciais/perfil completo, cadastrar registros equivalentes F4 328×478, F6 318×318 e F8 238×328 mm e um formato temporário de teste. Abrir Redimensionar, Step & Repeat e N-up e conferir catálogo, IDs e medidas. Testar cadastro/duplicação somente aqui. Com PDF sintético, alternar Retrato/Paisagem/retorno e manter Girar arte 90°; conferir quadrado F6, prévia e dimensões sintéticas. Não imprimir.
3. **Persistência Electron:** registrar a chave v1 e IDs no perfil isolado, fechar normalmente e aguardar o processo encerrar. Reabrir com a mesma origem e exatamente o mesmo `userData`, conferir a chave/lista nas três ferramentas e comparar os dados. Se houver falha, registrar antes de alterar qualquer configuração; não declarar homologação.
4. **Dados operacionais, somente após autorização própria:** a experimental.40 atualmente instalada ainda não contém A/B. Antes de futura atualização autorizada, com aplicativo fechado, atualizar/verificar backup privado e instalador anterior, registrar IDs/nomes/dimensões e hash do catálogo v1 por inspeção somente de leitura. Guardar evidência privada, sem inserir dados de perfil no Git. Cumprir também as validações INST-VAL exigidas para a futura versão, em procedimento separado.
5. **Homologação supervisionada da futura versão instalada:** abrir o atalho conferido, selecionar F4/F6/F8 em cada ferramenta usando só PDF sintético. Não usar Salvar formato, Atualizar, Excluir, importar catálogo ou reinicializar armazenamento. Alternar orientação/arte apenas na montagem em memória. Fechar e reabrir o aplicativo normalmente; conferir presença, IDs e medidas e comparar o catálogo com a referência anterior, sem mudanças nos registros reais. Qualquer divergência interrompe o teste para diagnóstico; não restaurar/mesclar automaticamente.
6. **Registro:** identificar versão/commit, runtime, origem/perfil, cenários realmente executados, capturas e resultado antes/depois, distinguindo ensaio isolado de homologação operacional. Guardar os resultados sintéticos na auditoria e somente resumo sem dados sensíveis na continuidade. Reabertura aprovada no Edge não substitui essa validação Electron.

Próxima ação depende da revisão do responsável; nenhum ensaio Electron, instalador ou acesso ao perfil real foi iniciado nesta etapa. Os registros seguintes descrevem a implementação anterior aos checkpoints.

Estado: implementado no checkout e aguardando revisão, sem commit, build ou instalação. Base `6c5a7f8024a5130eb4ee5178037473655a8b04f2`. A experimental.40 instalada continua inalterada e não contém estas mudanças.

## Parte A — orientação personalizada

**Causa:** o evento de orientação em `pdf-imposition-ui.js` alterava largura/altura somente quando existia um preset. A folha `custom` mantinha 297×420, e plano/prévia utilizavam essas medidas diretamente.

**Correção:** helper puro `pdf-sheet-orientation.js` normaliza a dupla numérica para Retrato=min×max ou Paisagem=max×min. A UI altera somente as medidas efetivas. O cadastro salvo não é invertido; Girar arte, escala, grade e demais parâmetros permanecem independentes. Digitação manual indica a orientação compatível com a dupla atual. Folhas quadradas mantêm as medidas.

**Arquivos para revisão A:** helper novo, `tests/pdf-sheet-orientation.test.mjs`, ajuste de orientação em `pdf-imposition-ui.js`, regressões em `tests/pdf-repeat-interface-browser.mjs` e resolução das importações reais em `tests/pdf-imposition-preview.test.mjs`. O roteiro de interface aceita `BROWSER_PATH`, sem instalar outro navegador.

**Verificado:** alternâncias repetidas 297×420 ↔ 420×297 mantendo arte 90°, quadrado 318×318, presets A3/A4/SRA3; campos, plano, prévia e páginas PDF sintéticas exportadas coerentes. O documento de origem sintético permanece separado do resultado.

## Parte B — formatos persistentes compartilhados

**Estrutura existente:** PDFFormatos e PDFArmazenamentoFormatos, chave `4printshops.pdf.formatos.v1`, documento `{versao:1,formatos:[{id,nome,largura,altura}]}`, medidas em mm. Redimensionar já oferecia cadastro/edição/exclusão, mas a montagem possuía somente presets locais/medidas livres.

**Implementado:** instância compartilhada por renderer em `pdf-format-catalog.js`; Redimensionar, Step & Repeat e N-up usam o mesmo catálogo. Seletor preserva presets e acrescenta registros salvos. Seção recolhível Formatos salvos permite cadastrar/reutilizar sem aplicar operação ao PDF. Edição/exclusão continuam na interface existente de Redimensionar.

**Preservação:** sem nova chave, versão ou migração. Carregar não grava dados; IDs, nomes e medidas existentes permanecem nos registros originais, inclusive campos adicionais. Medidas textuais legadas são projetadas numericamente para uso, sem reescrever os registros. Nomes canônicos impedem duplicações por espaços/caixa/Unicode. Na montagem, mesmo nome/medidas reutiliza ID; mesmo nome com medidas diferentes recusa. Outro nome com medidas iguais, inclusive invertidas, exige confirmação explícita nas duas interfaces.

**Concorrência e falhas:** gravações aguardadas e serializadas, Web Locks quando disponíveis; antes de escrever, comparar o conteúdo armazenado com a leitura anterior. Conteúdo divergente exige reabertura/recarregamento, sem sobrescrever a outra edição. JSON inválido, versão desconhecida ou falha de carga bloqueiam edição preservando o conteúdo. O fallback de armazenamento preexistente continua limitado à sessão e mostra aviso: não prometer persistência se o navegador recusar a gravação. Catálogo não sincroniza automaticamente computadores/perfis/origens. Sem Web Locks, a garantia principal é a fila no mesmo renderer e a verificação de divergência; não se propõe sincronização distribuída.

**Cadastro em uso:** editar/excluir em outra ferramenta atualiza a lista, mas mantém dimensões/prévia correntes como personalizadas e solicita nova seleção. Nenhum papel/escala/rotação é aplicado silenciosamente.

**Arquivos para revisão B:** `pdf-format-catalog.js`, `pdf-formats.js`, `pdf-format-storage.js`, `pdf-resize-ui.js`, integração do catálogo em `pdf-imposition-ui.js`, `tests/pdf-format-catalog.test.mjs` e `tests/pdf-sheet-catalog-browser.mjs`.

**Pendências de integração:** Livreto tem regras específicas de tamanho final/papel; nenhuma liberação arbitrária foi feita. Menu global de gerenciamento, catálogo entre máquinas, importação/exportação de bibliotecas e reformulação visual do N-up não foram implementados. PDF-010 não foi iniciado.

## Validações efetivamente executadas

Runtime Node 26.13.1 isolado já existente. Edge instalado, executado com perfil temporário; nenhuma instalação de dependências.

| Verificação | Resultado |
|---|---|
| Suíte PDF executada serialmente | 567 aprovados; 0 falhas, cancelados ou ignorados |
| `tests/pdf-imposition-preview.test.mjs` | 12/12 aprovados |
| `tests/pdf-repeat-interface-browser.mjs` em Edge | Aprovado; cenários existentes de interface e regressões A |
| `tests/pdf-sheet-catalog-browser.mjs` em Edge | Aprovado; 8 grupos A e 6 grupos B, sem erros inesperados de console |
| Persistência após encerramento/reabertura reais do Edge | Aprovada, mesma origem/perfil temporários, IDs e conteúdo preservados |
| F4/F6/F8 sintéticos equivalentes aos legados | IDs, nomes, dimensões e metadados preservados |
| Exportação sintética | Repeat 420×297/297×420 e N-up 328×478 mm conferidos com pdf-lib |
| Concorrência | Escrita válida aceita; instância desatualizada recusada sem perda |
| Armazenamento inválido/versão desconhecida/falha | Escrita bloqueada ou falha propagada conforme o adaptador; conteúdo anterior preservado |
| Sintaxe e higiene do diff | `node --check` em 11 arquivos aprovado; `git diff --check` aprovado; index vazio |

Comandos dos roteiros: usar o Node em `C:\Dev\4PrintShops-Tools\node-v26.13.1-win-x64\node.exe`, na pasta `Corel Script Configurator`; para o roteiro de interface definir `BROWSER_PATH=C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe`. Roteiro do catálogo utiliza Edge e dados sintéticos isolados. Não apontar os testes para perfis operacionais.

O ensaio adicional de persistência no Electron genérico não conseguiu obter o endpoint DevTools, mesmo antes de executar as verificações do catálogo; **não aprovado**, causa ambiental ainda não determinada. Rascunhos e logs ficaram fora do checkout para não incluir um roteiro incompleto. A reabertura do aplicativo Electron com este código continua pendente de ensaio isolado/homologação futura, sem substituir agora a instalação. Não atribuir essa limitação à persistência nem declará-la resolvida.

Durante a validação foram ajustados o adaptador de importações do teste de prévia e uma dependência de DOM não disponível no simulador de Redimensionar. As execuções finais acima passaram; logs intermediários foram mantidos.

## Evidências e revisão

Auditoria local: `C:\Dev\4PrintShops_Auditoria_Formatos_2026-11-17`.

- `pdf-suite-revisada.log`: execução serial de referência desta rodada, 567/567.
- `preview-regressao.log`: 12/12.
- `parte-A-browser-final.log`: roteiro de interface e regressão A.
- `browser-fonte-final.log` e `browser-validacao-completa.log`: ensaios A/B, sem erros inesperados.
- `PARTE_A.patch` e `PARTE_B.patch`: separação sequencial das duas implementações; aplicação A→B conferida em árvore isolada, reproduz os 11 arquivos de código/testes do checkout sem tocar no index. Os documentos são um conjunto adicional de revisão.
- `evidencias-browser-finais/`: resultado/parcial JSON e três capturas PNG copiados para fora do TEMP, sem copiar o perfil do navegador. `verificacao-final.log`: sintaxe, comparação dos patches e index vazio.
- Logs/rascunhos `electron-persistencia*` e `rascunho-pdf-format-*`: ensaio pendente, sem aprovação atribuída.

Não foram usados/movidos/renomeados PDFs reais. Não foi acessado ou alterado o localStorage operacional, nem os formatos reais F4/F6/F8. Instalador, atualizador, perfil, mecanismos de proteção PDF e bloqueios de impressão não foram alterados. As duas modificações documentais anteriores foram preservadas integralmente; novos registros foram acrescentados, sem apagar o histórico.

Próximo passo: revisão separada A/B e dos documentos. Não criar checkpoints, empacotar ou instalar sem autorização. PDF-001/PDF-008/PDF-004/PDF-009 permanecem Em testes; não considerar a experimental.40 homologada por esta rodada.
