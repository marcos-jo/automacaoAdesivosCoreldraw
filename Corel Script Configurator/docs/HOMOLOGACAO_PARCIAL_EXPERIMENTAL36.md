# Experimental.36 — homologação parcial na gráfica

Data do relato: **10/11/2026**. Resultado informado pelo operador, sem execução de novos testes neste checkpoint.

## Instalação e auditoria técnica anteriores

A experimental.36 foi instalada com sucesso na estação real da gráfica, mantendo identidade, diretório e perfil Hold Test. Foi uma **instalação limpa após a desinstalação manual da experimental.29**, e não uma migração transacional .29 → .36. O diário da instalação limpa terminou em committed.

A auditoria técnica já realizada confirmou 19.510 arquivos do pacote íntegros e 19.410 arquivos íntegros do LibreOffice privado, com inicialização e conversões sintéticas DOC/DOCX aprovadas. Perfil, preferências, associações próprias de PDF e leitor padrão foram preservados. Esses resultados não foram repetidos nesta etapa documental.

Detalhes, limites dos ensaios, hashes e referências das evidências: [análise da instalação](ANALISE_DESINSTALACAO29_INSTALACAO36.md). Os testes anteriores de migração .34 → .36 em VM são distintos desta instalação real .29 removida → .36 limpa.

## Cenário real aprovado pelo operador

O operador criou um documento Word real contendo somente a linha:

> teste importação word

Relatou que:

1. A importação demorou um pouco, mas concluiu.
2. O arquivo abriu corretamente no 4PrintShops.
3. Foi possível realizar a montagem.
4. A exportação foi concluída com sucesso.

**Classificação: homologação parcial da importação Word em cenário simples, aprovada pelo operador.**

A extensão do Word, o tipo de montagem e o formato da exportação não foram especificados no relato; não atribuir esse resultado a todas as variantes. A duração dessa importação real não foi cronometrada, portanto o desempenho não está homologado. O documento e seu conteúdo binário não foram adicionados ao repositório.

## Pendências

Não classificar este resultado como homologação completa de DOC/DOCX. Continuam pendentes documentos complexos, fidelidade e substituição de fontes, imagens, tabelas, paginação, desempenho, compatibilidade ICC/RIP e impressão física. Uma aprovação em cenário simples não estende os resultados às demais combinações.

## Próxima prioridade

Investigar e implementar, em tarefa futura, **importação segura de PDFs protegidos, impedindo travamentos em lotes mistos**. Este checkpoint apenas registra a prioridade; não inicia diagnóstico, implementação ou novos testes dessa funcionalidade.

## Preservação e continuidade

Evidências, documentos de teste, instaladores e backups locais/externos existentes foram preservados, permanecendo fora do Git. O repositório transporta os fontes e a documentação; os recursos privados de runtime e os instaladores não acompanham o clone. Para preparar outra estação, consultar [DESENVOLVIMENTO_WINDOWS.md](DESENVOLVIMENTO_WINDOWS.md) e [DISTRIBUICAO_DESKTOP.md](DISTRIBUICAO_DESKTOP.md), sem assumir que os caminhos locais dos relatórios existirão em casa.

Este encerramento altera somente documentação e verifica o diff/staging. Nenhuma instalação, desinstalação, impressão, alteração funcional, novo teste ou envio ao GitHub foi realizado nesta etapa. Commit local autorizado pelo operador; push exclusivamente manual.
