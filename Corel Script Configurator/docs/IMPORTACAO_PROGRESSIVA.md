# Conferência antecipada de importação — desenvolvimento após a .33

**.34 instalada para avaliação — 08–09/11/2026:** atualização controlada .33→.34
concluída com perfil/preferências preservados. Binário instalado e ASAR real
validados com sintéticos; testes mistos do ASAR usaram seletor simulado, deixando
o seletor nativo e a experiência visual para o operador. Não houve mudanças de
funcionalidade nesta instalação. Relatório e roteiro atualizados em
[EXPERIMENTAL34_IMPORTACAO_PROGRESSIVA.md](EXPERIMENTAL34_IMPORTACAO_PROGRESSIVA.md).
Menções abaixo a não instalada/checkout descrevem etapas anteriores.

**Empacotamento .34 — 08/11/2026:** a implementação descrita abaixo agora está no
Setup experimental.34, gerado e auditado para avaliação; não instalado. A .33
instalada permanece preservada. Não houve funcionalidade nova no empacotamento;
somente versão, documentação e ajuste temporal no teste Word. Resultados, hash,
metodologia e roteiro: [EXPERIMENTAL34_IMPORTACAO_PROGRESSIVA.md](EXPERIMENTAL34_IMPORTACAO_PROGRESSIVA.md).
As referências abaixo a somente checkout/sem pacote descrevem a etapa anterior.

Base: main/1845fc1 (experimental.33). Alteração somente no checkout, para avaliação.
Não foi gerado/instalado novo pacote e a instalação .33 permanece a anterior.

## Diagnóstico e solução incremental

O IPC `files:import` anterior só devolvia a seleção depois de ler todos os arquivos
e converter todos os Word. Em seguida, a janela ainda validava PDFs/imagens e
preparava miniaturas. O operador via uma espera genérica até essa preparação acabar.

Mantidos o seletor nativo, o conversor privado, os limites de128 arquivos/512MiB,
o preflight de todo o lote, PDFCombinacao e a adoção transacional no PDFWorkspace.
Acrescentados eventos de seleção, etapa e arquivo preparado, limitados à janela
autorizada e identificados por um token por operação. Eventos levam somente
nome-base/tipo/índice, bytes necessários à conferência e relatório Word existente;
não enviam caminhos completos ao novo painel nem criam cópias permanentes.

A janela Conferir e combinar arquivos abre após a seleção nativa, antes da
conversão. Cada arquivo recebe nome/tipo, espaço reservado e indicador animado.
Estados: Aguardando, Convertendo Word, Lendo PDF/imagem, Preparando imagem,
Gerando miniaturas, Pronto e Falha. A contagem reflete arquivos efetivamente
validados com primeira miniatura concluída, sem porcentagens estimadas.
Respeitada a preferência do sistema de reduzir animações.

O rascunho anterior é copiado para uma área provisória: cada arquivo novo passa
pelas mesmas validações da combinação. A primeira miniatura de cada origem só
é publicada após validar o documento e renderizar essa página. As páginas
restantes usam a fila preguiçosa anterior depois que o lote está pronto; não é
necessário rasterizar350 páginas para liberar a conferência. PDFs mantêm o
conteúdo vetorial na combinação pelo motor anterior.

O protocolo espera a entrega de todos os índices esperados E a fila de validação/
miniaturas antes de liberar a organização. Não presume ordem entre a resposta
final do IPC e os eventos de progresso. Existe prazo de15s para a entrega interna
após a resposta final; não é um prazo para converter Word. Cancelar continua
disponível durante essa espera. Um teste Electron detectou a corrida na primeira
implementação; a barreira de entrega e uma regressão específica foram acrescentadas.

## Operação e proteção

- Durante a preparação, adicionar, editar opções, organizar/remover, limpar e
  combinar ficam bloqueados. Cancelar importação, × e Escape cancelam a seleção.
- Falha em qualquer arquivo invalida toda a nova seleção. As miniaturas
  provisórias não constituem importação; o documento aberto e o rascunho anterior
  permanecem intactos. Cancele e selecione novamente para corrigir a seleção.
- Cancelamento aborta o conversor e o carregamento PDF.js, invalida a geração e
  ignora eventos/miniaturas tardios. Nova seleção tem outro token e outra geração.
- A liberação tardia de uma operação antiga não desbloqueia uma combinação nova.
  Limpar a lista também bloqueia os controles enquanto libera a prévia anterior.
- Depois de pronto, combinar continua sendo ação explícita. Criar novo PDF com
  edições pendentes mantém a confirmação de descarte existente.
- No Desktop, inclusive um único PDF selecionado pelo botão Importar agora passa
  pela conferência, para que a preparação tenha o mesmo fluxo de aprovação.
  No navegador, a abertura direta de um PDF conserva o fluxo anterior; lotes e
  imagens na combinação usam a preparação progressiva. Word continua exclusivo
  do Desktop, sem conversão HTTP ou serviço externo.

## Verificação e uso para avaliação

Resultados executados em08/11/2026:

| Verificação | Resultado |
|---|---|
| Suíte Node completa |468 aprovados; zero falhas, cancelamentos ou skips. Dez regressões novas: oito da conferência/ponte e duas do backend. |
| Chrome real | Conferência antecipada, miniatura enquanto outra origem espera, animação/contagem,1/100/350 páginas, falha intermediária, cancelar/reabrir, ignorar resultados tardios e recusa/aceite do descarte aprovados. |
| Electron de desenvolvimento | Mesma matriz de conferência e teste adicional do IPC real com seleção sintética aprovados, perfil isolado. |
| Electron + LibreOffice privado | PDF/PNG/JPG/DOCX/DOC em sete páginas; primeira miniatura durante conversão Word, batimento de interface, controles bloqueados até o fim, erro no meio/cancelamento e cinco originais intactos aprovados. |
| Electron smoke existente | Edição, abas, giro, salvamento/exportação, cancelamento, falhas nativas, proteção/persistência e fechamento aprovados. |
| Montagem Chrome existente | Step & Repeat, N-up tradicional/compacto/escalas, livretos, prévias, retry, cancelamentos e salvamento aprovados; sem fila física. |
| PDFs convertidos | Dois PDFs Word × Step & Repeat/N-up/livreto: seis cenários aprovados, bytes de origem preservados. |
| Sintaxe/diff | JS/MJS/CJS alterados e novos verificados; git diff --check aprovado. |

Os testes Node emitem avisos do PDF.js sobre ambiente/build legacy e
`Math.sumPrecise`; não houve falhas. As miniaturas reais também foram verificadas
em Chrome/Electron, sem substituir essa conferência pelos canvases simulados do Node.
Uma primeira rodada da automação teve código de teste sem wrapper async e nome
incorreto de botão; corrigidos somente no roteiro. A corrida real do novo IPC
descrita acima foi corrigida na implementação, e a repetição Word completa passou.
Também foi reproduzida uma corrida ao limpar e imediatamente adicionar arquivos;
Limpar agora bloqueia controles até liberar a prévia, preservando o novo rascunho.

Inventário desta etapa:

- Backend/ponte: `desktop/import-files.mjs`, `desktop/app.mjs`, `desktop/preload.cjs`,
  `public/js/app/desktop-client.js`.
- Conferência: `public/js/tools/pdf-import-review.js` (novo), `pdf-merge.js`,
  `pdfs.js` e `public/css/theme.css`.
- Testes: `tests/pdf-import-progress.test.mjs`, `tests/pdf-import-progress-browser.mjs`
  (novos); `desktop-import-files.test.mjs`, `desktop-word-smoke.mjs`,
  `pdf-integration.test.mjs`, `helpers/pdf-combine-modules.mjs`, `helpers/pdf-dom.mjs`.
- Documentação: este guia, `ESTADO_PROJETO.md` e `IMPORTACAO_WORD_LOCAL.md`.

Executar na pasta da aplicação, com dependências e runtime privado já preparados:

```powershell
node --test tests/pdf-import-progress.test.mjs tests/desktop-import-files.test.mjs
npm test
node tests/pdf-import-progress-browser.mjs
$env:IMPORT_PROGRESS_ELECTRON='1'
node tests/pdf-import-progress-browser.mjs
Remove-Item Env:IMPORT_PROGRESS_ELECTRON
node tests/desktop-word-smoke.mjs
node tests/desktop-smoke.mjs
```

Os roteiros usam somente documentos sintéticos e perfis isolados. O teste Word
usa o motor privado real, testa os cinco formatos em sete páginas e verifica
miniatura pronta enquanto outro arquivo está Convertendo Word, controles
bloqueados, batimento de interface, falha/cancelamento e originais preservados.
Os testes gerais de montagem podem ser executados em Chrome com
`IMPOSITION_DESKTOP=0`, sem encaminhar à fila de impressão.

Evidências desta etapa ficam em `.desktop-test/import-progress/` e nos diretórios
temporários identificados pelos roteiros Electron, todos fora do Git.

Roteiro manual de avaliação: selecionar um PDF pequeno e depois um lote misto;
observar a janela antes da conversão terminar, nomes sem caminhos e miniaturas
progressivas; cancelar durante Word, selecionar novamente, testar um arquivo
inválido no meio e confirmar que não há importação parcial; combinar com um
documento editado e recusar a substituição; conferir a ordem e os controles ao fim.

Limites: o indicador não mede porcentagem interna do LibreOffice. O tempo de
conversão não foi reduzido. A validação de PDFs/imagens continua no renderizador
existente, podendo consumir CPU em documentos complexos; não foi criado um motor
paralelo nem prometida responsividade uniforme para qualquer PDF. Apenas a
primeira página de cada arquivo é antecipada; as demais miniaturas aparecem sob
demanda. Testes sintéticos não homologam fontes/paginação/ICC/RIP, Word em Windows
limpo, documentos reais em massa ou impressão/livretos físicos. Motores de
Step & Repeat, N-up, livreto, redimensionamento, exportação e impressão não foram
alterados. Sem commit ou push autorizado nesta etapa.
