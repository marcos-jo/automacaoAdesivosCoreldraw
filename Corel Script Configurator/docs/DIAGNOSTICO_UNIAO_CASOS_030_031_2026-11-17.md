# Diagnóstico de união — CASO-030 e CASO-031

## Decisão do projeto — diagnóstico revisado e encerrado em 17/11/2026

O responsável confirmou os resultados: os dois casos abrem individualmente, ambos têm ICC sRGB e a proteção atual de OutputIntents bloqueia intencionalmente a união. Dez ensaios concluídos (oito recusas envolvendo os novos casos e duas uniões de controles nas duas ordens); originais/cópias e documento aberto preservados. Esses resultados são os ensaios anteriores revisados, não uma nova execução nesta decisão.

Registrado **PDF-011 — Agrupamento seguro com OutputIntents**, após conferir ausência de identificador específico já existente. Estado Planejado, etapa independente antes da primeira estável. Abrange preservação ICC/vínculos, compatibilidade de perfis, referências globais, integridade da saída e bloqueio contínuo de casos incompatíveis. Não remover a proteção sem solução segura nem implementar agora. Ver critérios no planejamento.

A experimental.42 será dedicada exclusivamente ao gerenciamento geral de tamanhos; não incorpora alteração de agrupamento, PDF-010 ou Livreto. Este diagnóstico fica encerrado como registro de regressão; eventual implementação requer nova etapa/autorização. PDFs, hashes detalhados e evidências privadas continuam apenas na biblioteca local ignorada, fora deste checkpoint documental.

Investigação autorizada, separada do catálogo de tamanhos, PDF-010 e Livreto. Base `44a69dabc53ca68363e30832dddae7604bd0ebd2`, após os três checkpoints do catálogo. **Nenhuma correção implementada.** Este documento não inclui PDFs, conteúdo de páginas, imagens de clientes ou perfis.

## Registro e reprodução

- Dois PDFs novos localizados na entrada local e registrados como CASO-030/CASO-031, categoria união, com nomes originais, metadados, tamanhos e SHA-256 nos inventários ignorados. Cópias classificadas verificadas por hash; recebidos mantidos em seus caminhos/nomes originais. Base histórica e registros anteriores preservados; observações novas, sem sobrescrita.
- CASO-030 tem cinco páginas; CASO-031 tem uma. Ambos renderizaram todas as páginas individualmente. Controles conhecidos CASO-014/015 abriram uma página cada.
- Dez tentativas efetivas na interface: par novo nas duas ordens (recusado), cada novo sozinho (recusado), cada novo com CASO-014 nas duas ordens (novo recusado, controle pronto), CASO-014/015 nas duas ordens (aprovado, duas páginas).
- Em todas as oito recusas, o workspace aberto permaneceu com o mesmo documento/revisão/nome/quantidade de páginas. Não houve exportação ou impressão. Zero console.error/exceções inesperadas/avisos na execução final; hashes de recebidos/cópias/controles preservados.
- Ambiente: Edge 157.0.4341.60, localhost e perfil exclusivo. Onze módulos PDF envolvidos conferidos contra a fonte auditada do build .41, iguais (normalizadas quebras de linha quando necessário). A .41 instalada não foi aberta/modificada para este ensaio; equivalência dos módulos não é nova homologação do executável Electron.

## Causa confirmada

Ambos contêm `/OutputIntents` com a mesma declaração e perfil ICC sRGB incorporado. `PDFCombinacao.add`, em `public/js/tools/pdf-combine-session.js:51`, recusa explicitamente esse recurso antes da cópia das páginas. `pdf-import-worker.js` executa esse preflight durante **Preparar arquivo para união**. Abertura individual e agrupamento têm políticas diferentes; a ordem não muda a recusa. O diagnóstico existente não identificou assinatura ou conflito XRef nesses dois casos.

Mensagem técnica: **PDF com intenção de saída/perfil de impressão. A combinação ainda não preserva esse perfil global com segurança.**

O erro não possui código específico; `pdf-feedback.js` apresenta **Operação não concluída. Consulte os detalhes ou peça ajuda ao suporte.** A explicação fica em Ver detalhes. Stack reproduzido na classe aponta à reemissão em `pdf-combine-session.js:75:37`, com guarda na linha 51. O status geral bloqueia Combinar e informa que nenhuma página foi adicionada.

Os controles não possuem OutputIntents e foram combinados com sucesso. O agrupador cria documento novo/copia páginas, sem transferência explícita dessa declaração global; portanto a recusa atual é uma proteção deliberada e uma limitação funcional, não motivo para classificar os recebidos como inválidos. Metadados declaram Canva; isso não prova origem externa nem certifica PDF/A/CMYK. Há também estruturas globais de marcação/acessibilidade e marcadores que exigem análise antes de liberar a união.

## Proposta para revisão

1. Código/mensagem específicos para esse bloqueio, preservando detalhes e atomicidade.
2. Investigar suporte limitado a intenções/perfis comprovadamente equivalentes, preservando ICC e referências globais. Considerar marcação/acessibilidade, marcadores e metadados. Não basta remover a guarda ou descartar perfil; não afirmar conformidade PDF/A sem validação.
3. Manter perfis diferentes, mistos/ausentes, múltiplos ou inválidos bloqueados até política/testes explícitos. Futuras regressões devem incluir ambos os sentidos deste par, controles, preservação de perfil/recursos, páginas/ordem/renderização, falhas atômicas e proteções existentes.

Nenhuma remoção de estruturas, conversão, reparo ou geração forçada foi executada. A implementação depende de nova autorização; não iniciar PDF-010 nem misturar com tamanhos/Livreto.

## Evidências privadas e limites

Relatório completo, nomes, hashes e capturas ficam apenas em `.desktop-test/casos-reais/evidencias/uniao-030-031-20261117/RELATORIO.md`; execução final `execucao-F4GAzU/reproducao.json`. Inventários incrementais locais incluem os dois casos, recibo e observações; nenhum PDF/captura/perfil no Git. O manifesto final confirma hashes antes/depois.

Duas execuções preparatórias do harness tiveram esperas incorretas dos controles, corrigidas somente no roteiro local: reconhecer a transição de preparação para lista pronta e selecionar destino Novo após o início da seleção. Falhas preservadas; a execução final completa é a evidência de aprovação dos controles. Nenhum timeout foi aumentado ou erro silenciado.

Só foram executados estes casos/matriz e leitura estrutural. Suíte PDF completa, regressões do catálogo e ensaios de instalador não foram repetidos. Código funcional, .41 instalada, perfil real e formatos operacionais intactos. Sem commit/push, build, instalação ou impressão nesta tarefa.
