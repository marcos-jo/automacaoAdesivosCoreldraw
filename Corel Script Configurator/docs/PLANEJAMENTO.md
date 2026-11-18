# 4PrintShops — Planejamento Geral de Implementações

## 18/11/2026 — Preparação e Setup .42 autorizados, instalação não autorizada

- Responsável aprovou a limpeza e autorizou checkpoints separados de filtro/teste e documentação, preparação de 1.0.0-experimental.42, regressões, Electron isolado, INST-VAL sintético .41→.42 e geração/auditoria em pasta inédita somente após aprovação integral de testes, recursos e armazenamento. Sem push, instalação/desinstalação, perfil operacional, impressão ou redefinição real.
- Escopo funcional permanece no catálogo único/gerenciador/nove sugestões administráveis/backup e recuperação controlada, conservando as proteções da .41. PDF-010/PDF-011/abas/servidor/Livreto fora desta geração. Nenhuma atualização deve repovoar automaticamente o catálogo existente ou vazio.
- Filtro/teste revisados e registrados em `4bec43c`; nova execução válida de filtro/backup/distribuição: 15/15. Invocação inicial inválida usou cwd acima da aplicação e caminho inexistente para dependências, causando ENOENT/MODULE_NOT_FOUND; corrigidos somente cwd e variável do processo, sem alteração de produto/testes. Demais validações e pacote ainda pendentes; interromper o empacotamento se houver falha na execução válida.
- Documentação histórica e recibos privados da limpeza preservados. Metadados/pacote terão commit-base exato e comprovação de árvore limpa. Revisão do Setup pelo responsável precede qualquer futura instalação supervisionada.

## 18/11/2026 — Oito alvos removidos; revisão antes de empacotar .42

- Diretório padrão confirmado por Get-Location sem parâmetros/workdir: C:\Dev\4PrintShops. Backups no servidor e 86.087 registros de origem revalidados. Removidos exclusivamente os sete alvos mais a auxiliar de hardlinks; restante da auditoria preservado. A primeira tentativa esvaziou o projeto antigo, mas a raiz foi bloqueada por três auxiliares ociosos do Codex com cwd antigo; somente eles foram encerrados, a verificação foi ampliada para cwd/direito de exclusão e a retomada concluiu os oito alvos. Histórico completo no preflight.
- Livre: 10,768 → 20,584 GiB durante a operação; aumento líquido observado 9,816 GiB, sujeito à atividade normal do volume. Pós-conferência: 20,597 GiB. Não somar a auxiliar aos dados originais. Resultados individuais e bytes nos recibos privados da auditoria de armazenamento.
- Pós-limpeza: 2.660 arquivos do checkout íntegros, 19.515 estáticos da .41 iguais, 76 arquivos do perfil idênticos ao snapshot pré-limpeza; Setup .41 e Git íntegros. HEAD bccce139/main, 12 commits locais; seis caminhos pendentes preservados. Apenas quatro documentos de continuidade atualizados depois da conferência. Sem consulta nova ao remoto, commit ou push.
- Armazenamento resolvido para revisão; filtro e 15/15 são resultados anteriores. Metadados .41, sem build/Setup .42/instalação/reset/impressão. Setups .29/.36 seguem indisponíveis. **Aguardar revisão antes de retomar metadados, testes/preflight funcional e empacotamento .42.** Lista exata, bloqueio/resolução e medições em PREFLIGHT_FILTRO_E_ARMAZENAMENTO_42_2026-11-17.md. Instalação/perfis/PDFs reais/Setups preservados/recursos futuros fora da limpeza.


## 18/11/2026 — Limpeza autorizada; execução suspensa por diretório padrão

- Responsável autorizou exclusivamente os sete alvos e a auxiliar listados no preflight. A recuperação independente não executou exclusão: o diretório padrão da sessão ainda aponta ao projeto antigo, confirmado sem parâmetro workdir. Prosseguir somente após a execução estar efetivamente vinculada a `C:\Dev\4PrintShops`, mantendo a lista aprovada.
- Backups/inventários revalidados, instalação .41 conferida e perfil acessível; alterações operacionais e geração .42 continuam proibidas nesta recuperação. Livre ~10,81 GiB. Detalhes/diagnósticos persistentes em `PREFLIGHT_FILTRO_E_ARMAZENAMENTO_42_2026-11-17.md`. Sem nova implementação ou mudança de prioridade.

## 18/11/2026 — Preparação final da limpeza, exclusão ainda condicionada

- Filtro aprovado com 15/15 anteriores; preservar a exceção exata sem alterar catálogo ou versão. Nenhuma suíte funcional nova exigida ou executada para arquivar arquivos.
- Autorizado arquivar somente conteúdos ainda não preservados e conferir os sete alvos do plano de aproximadamente 9,72 GiB. A exclusão continua dependendo da última autorização do responsável; não gerar .42 agora.
- Preservar também os documentos/alterações pendentes e os 12 commits locais. Setups .29/.36 não localizados, sem garantia de reconstrução exata. Não duplicar os 31 casos reais nem o Setup .41 já verificados.
- Procedimento, recibos e pasta auxiliar de vínculos do arquivamento identificados em `PREFLIGHT_FILTRO_E_ARMAZENAMENTO_42_2026-11-17.md`. Após limpeza autorizada, medir espaço efetivo e retomar metadados, testes e preflight .42, mantendo .41 e perfis operacionais intactos.

## 17/11/2026 - Filtro .42 corrigido e plano de armazenamento em revisao

- Regra de backup mantida com inclusao exata de pdf-format-backup.js; legado excluido. FileMatcher real e regressao do backup/distribuicao: 15/15 aprovados, sintaxe e diff-check aprovados; sem build ou novas suites completas.
- Politica: .29 producao, .36 beta, .41 desenvolvimento. Projeto antigo e intermediarios candidatos a remocao somente apos ultima conferencia e preservacao verificada dos itens locais ignorados. Plano de sete caminhos recupera aproximadamente 9,72 GiB; C: tem cerca de 11,28 GiB livres. Nenhuma exclusao executada.
- 31 casos reais verificados por hash no servidor; bundles dos repositorios (incluindo 12 commits locais), alteracoes atuais e Setup .41 preservados com verificacao. Setups .29/.36 nao localizados; responsavel informou provavel exclusao anterior, preservacao dessas versoes continua pendente.
- Detalhes, caminhos completos, itens locais e condicoes em PREFLIGHT_FILTRO_E_ARMAZENAMENTO_42_2026-11-17.md. Arquivamento dos dados ignorados/geracoes e decisao sobre instaladores ausentes precedem limpeza. Metadados ainda .41; .42 aguarda armazenamento, preflight e autorizacao de empacotamento. Perfis, formatos reais e instalacoes intactos. Sem commit/push. Historico abaixo preservado.

## 17/11/2026 — Diagnóstico encerrado, PDF-011 planejado; preflight .42 interrompido

- Responsável revisou CASO-030/031: abertura individual aprovada, perfil ICC sRGB em ambos, agrupamento intencionalmente bloqueado por OutputIntents. Dez ensaios anteriores concluídos (oito recusas e dois controles aprovados); originais preservados. Não remover a proteção sem solução segura. Diagnóstico encerrado sem implementação.
- Conferidos identificadores existentes: PDF-004 é compatibilidade geral/recorte marcado e PDF-007 é inspeção; não havia ID específico para união com OutputIntents nem PDF-011 ocupado. Registrado **PDF-011 — Agrupamento seguro com OutputIntents**, Planejado, antes da primeira estável, fora da .42 e separado de PDF-010/Livreto. Critérios abaixo; não autorizado implementar agora.
- Experimental.42 autorizada exclusivamente para catálogo geral de tamanhos, base funcional `44a69da` mais documentação revisada. Conservar guardas/funcionalidades da .41 e não redefinir o catálogo operacional na instalação.
- **Preflight interrompido antes de metadados, regressões funcionais novas e build:** `electron-builder.config.cjs` exclui `!public/js/tools/*backup*`, atingindo o módulo obrigatório `pdf-format-backup.js`. Confirmado com minimatch e o filtro real FileMatcher do builder isolado: pdf-formats.js incluído, pdf-format-backup.js excluído, legado pdfs backup.js excluído. É necessário revisar a exclusão para manter fora o legado e incluir o módulo legítimo; nenhuma alteração feita nesta configuração nesta etapa, conforme ordem de parar em falha.
- Livre observado ~11,34 GiB; geração .41 anterior ocupa ~7,91 GiB nos grupos medidos (dependências/build/dist/audit), sem contar temporários e margem. A reserva anterior de 12 GiB não está disponível; replanejar armazenamento antes de retomar. Sem limpeza/cópia de build/Setup/instalação/push ou alteração de perfil. Electron/INST-VAL/regressões desta .42 ainda não executados; resultados históricos não são novos.

## 17/11/2026 — Casos reais 030/031: união bloqueada por perfil, somente diagnóstico

- Após concluir checkpoints do catálogo, registrados dois PDFs novos de união, com originais/nomes/hashes preservados na biblioteca local ignorada. Abertura integral de ambos aprovada; recusa nas duas ordens e contra controle confirmada por guarda OutputIntents. Controles conhecidos unidos com sucesso nas duas ordens.
- Diagnóstico e proposta em `DIAGNOSTICO_UNIAO_CASOS_030_031_2026-11-17.md`. Nova investigação separada de catálogo/PDF-010/Livreto, sem alteração de prioridades ou implementação. Revisar antes de autorizar qualquer suporte a perfis; não remover guarda/perfil para contornar o bloqueio.
- Nenhuma mudança funcional, instalação, perfil real, impressão, commit/push ou suíte completa. Evidências/PDFs privados exclusivamente em `.desktop-test/`, sem inclusão no Git. PDF-001/008/004/009 continuam Em testes.

## 17/11/2026 — Lista oficial confirmada e checkpoints A/C/B

- Responsável confirmou os nove formatos e todas as medidas: A3/A4/A5/A6/A7/F4/F8/F6/Faixa. **A4 210×297 mm está confirmado, sem pendência**; todos editáveis/excluíveis. Referências provisórias abaixo são histórico superado por esta decisão.
- Divisão segura: A `ab483e6` (seletores das montagens), C `0e80a68` (serviço, nove iniciais, backup/recuperação e testes), depois B (menu/gerenciador, integração e estes registros). Ordem evita duplicar modelos ou referenciar módulo ainda não registrado; cópias do índice A/C validadas antes dos commits.
- Nova execução final: **108/108** dirigidas seriais (94+14), Edge/Electron A/B com três reinícios e catálogo vazio persistente, backup/reset/restore sintéticos aprovados, zero erros. Fontes funcionais mantidas byte a byte em relação ao trabalho revisado; nenhum novo escopo.
- Redefinição operacional **não é automática nem está na interface**. Depende de autorização e confirmação do operador em procedimento técnico assistido, backup privado salvo/relido/SHA-256 e chamada explícita da API; a API não abre diálogo de confirmação. Somente a chave do catálogo é escrita/restaurada. Nenhum reset real nesta etapa.
- Preparar .42 depende de revisão e autorização de empacotamento, metadados, INST-VAL e auditoria/espaço do pacote; nada disso foi executado agora. Sem Setup/instalação/perfil real/impressão/push/PDF-010. PDF-001/008/004/009 Em testes; histórico abaixo preservado.

## 17/11/2026 — Catálogo inicial de nove tamanhos e redefinição controlada

- Nova decisão: dispensar compatibilidade completa com catálogos experimentais antigos. Preservadas A/B; serviço simplificado para sugestões A3/A4/A5/A6/A7/F4/F8/F6/Faixa somente em primeiro uso (chave ausente). Todos administráveis; qualquer catálogo existente, inclusive vazio, não é completado nem repovoado.
- Mesmo PDFFormatos/chave v1 nas quatro interfaces. Preparados backup SHA-256, redefinição condicionada ao backup atual e recuperação exata somente da chave. Mecanismo de manutenção assistida, sem reset automático/botão no aplicativo instalado e sem conversão de perfil real.
- 94/94 regressões seriais novas + 14/14 adicionais; Edge/Electron ampliados com três reinícios, vazio persistente e reset/restore sintéticos aprovados. Relatório `SINCRONIZACAO_E_GERENCIAMENTO_TAMANHOS_2026-11-17.md` atualizado sem apagar ensaios anteriores.
- A4 permanece interpretação provisória; confirmar antes de redefinir dados operacionais. Aguardar revisão A/B/procedimento e autorização específica para qualquer reset real, checkpoint ou nova distribuição. .41/perfis/PDFs reais intactos; sem commit/push/build/impressão. Prioridades e PDF-001/008/004/009 Em testes preservados.

## 17/11/2026 — Exclusões sincronizadas e Gerenciar tamanhos em revisão

- Ocorrência da homologação .41: SRA3 excluído no Redimensionar permanecia nas montagens por opção estática independente. Correção A usa somente PDFFormatos; todos os tamanhos iniciais passam a ser administráveis, conforme esclarecimento do responsável. Correção B acrescenta Arquivo → Gerenciar tamanhos com o mesmo serviço/chave.
- Implementado e testado somente no checkout e em perfis sintéticos. 87 regressões dirigidas seriais, roteiros Edge e Electron isolado com dois reinícios aprovados. Sem instalação, commit, push, impressão, alteração do perfil real ou dos cadastros F4/F6/F8 reais.
- Revisar A/B e extensão do documento v1 para catálogo integral antes dos checkpoints/distribuição: primeira edição explícita persiste a lista completa, e retorno à .41 antiga pode exigir restauração revisada do backup anterior. Plano e resultados em `SINCRONIZACAO_E_GERENCIAMENTO_TAMANHOS_2026-11-17.md`. Não executar conversão no perfil real nesta etapa.
- Atualização real .40→.41 anterior concluída e conferida; homologação operacional da .41 permanece pendente pela ocorrência, e este código ainda não está instalado. PDF-001/008/004/009 continuam Em testes; PDF-010/Livreto/reformulação visual do N-up não iniciados. Registros anteriores abaixo preservados como histórico.

## 17/11/2026 — Experimental.41 gerada, atualização operacional aguardando autorização

- Checkpoints .41 e documentação concluídos separadamente (fd8cd5a/a3dfdf7). Novo Setup .41 gerado da base a3dfdf7 em pasta inédita, preflight e auditoria integral aprovados. Fonte atual, formatos/orientação/mensagens e correções do atualizador incluídos; nenhuma nova funcionalidade.
- Relatório: `C:\Dev\4PrintShops-Builds\experimental.41-FINAL-20261117-dE9Y3n\RELATORIO_EXPERIMENTAL_41.md`. Instalador não assinado, sem publicação/canal automático; Setups .39/.40 e perfil real preservados. Livre ~15,31 GiB, sem limpeza.
- **Não instalado.** Aguardar revisão/autorização antes da atualização .40→.41, renovar backup privado e conferir formatos/preferências/aplicativos fechados/transações. Homologação operacional posterior; PDF-001/008/004/009 continuam Em testes, PDF-010/Livreto/reformulação do N-up não iniciados. Novos registros ficam para revisão, sem terceiro commit/push.

## 17/11/2026 — Retomada .41: Electron e INST-VAL aprovados sem pacote

- Harness corrigido/instrumentado sem alterar aplicativo: três fechamentos normais e duas reaberturas isoladas aprovados, catálogo/orientação preservados. Causa original do timeout permanece hipótese, sem defeito funcional confirmado.
- Validações pendentes executadas: 100 regressões de atualização, quatro casos .40→.41 e duas instalações/reinstalações sintéticas, todos aprovados. Espaço ~23,92 GiB suficiente para reserva incremental de 6–8 GiB (ou 9–12 GiB em pasta integral nova). Nenhuma limpeza.
- Aguardar revisão antes de gerar Setup. Preflight da fonte/configuração e integridade integral dos recursos privados do preparo interrompido devem preceder o próximo build; auditoria e homologação operacional permanecem posteriores. Não instalar, ampliar escopo ou iniciar PDF-010. Metadados .41 e documentos continuam pendentes, sem commit/push; PDF-001/008/004/009 Em testes.

## 17/11/2026 — .41 autorizada para preparo; geração interrompida por teste

- Validação Electron documentada no checkpoint 3774468. Versão .41 preparada somente em metadados, sem mudar identidade/perfil ou ampliar funcionalidades.
- 81 regressões dirigidas aprovadas; repetição Electron falhou no primeiro encerramento. Conforme decisão de interromper em falha, **não gerar pacote** até diagnosticar o fechamento e concluir reabertura/regressões de atualização/INST-VAL/auditoria. Relatório na pasta experimental.41-INSTVAL-20261117-CqiKQD em 4PrintShops-Builds.
- .40 instalada e formatos reais preservados; nenhuma autorização de instalação .41 executada. PDF-001/008/004/009 continuam Em testes; PDF-010, Livreto e reformulação do N-up não iniciados. Aguardar revisão para retomada.

## 17/11/2026 — Validação Electron autorizada e concluída em perfil isolado

- Orientação e compartilhamento entre Redimensionar/Step & Repeat/N-up aprovados no Electron 45.1.0 com encerramento e reabertura reais, perfil sintético exclusivo e quatro cadastros/IDs preservados. Nenhuma alteração do perfil operacional ou da .40 instalada.
- Responsável informou boa avaliação da interface/textos da .40. As melhorias A/B ainda exigem futura versão para homologação operacional; isso não homologa a produção nem muda o estado Em testes de PDF-001/008/004/009.
- Próximo passo: revisão para autorizar a preparação da experimental.41, mantendo INST-VAL/auditoria e preservação dos formatos reais. Nenhum instalador/build/instalação foi autorizado ou executado nesta etapa; não iniciar PDF-010/Livreto/reformulação do N-up. Detalhes na continuidade e auditoria Electron de 17/11.

## 17/11/2026 — Revisão autorizada: checkpoints separados e próximo ensaio

- A registrada em `d1ad3f0` (orientação personalizada). B registrada separadamente no checkpoint **Compartilha formatos personalizados nas montagens PDF**, que contém este registro e a integração local mínima aprovada. Os registros de implementação em revisão abaixo descrevem a etapa anterior.
- Próximo passo **somente planejado**: validar no Electron isolado e, após autorização específica de nova versão, homologar no instalado por seleção/leitura de F4/F6/F8 nas três ferramentas, reabertura e independência orientação/arte. Persistência real no Electron permanece pendente. Procedimento no relatório A/B.
- Nenhuma autorização implícita para gerar instalador, alterar perfis, executar o ensaio agora ou iniciar PDF-010/Livreto/reformulação do N-up. PDF-001/PDF-008/PDF-004/PDF-009 permanecem Em testes.

## 17/11/2026 — Rodada autorizada: orientação e formatos nas montagens

- **A — Orientação personalizada:** correção implementada no checkout, em revisão. Retrato/Paisagem normalizam as medidas efetivas sem alterar o cadastro nem Girar arte. Inclui regressões de prévia/exportação sintética, alternâncias repetidas e folha quadrada.
- **B — Catálogo local compartilhado:** implementado o reaproveitamento mínimo de PDFFormatos v1 entre Redimensionar, Step & Repeat e N-up, com cadastro/reutilização na montagem. Mesma chave, IDs e medidas preservados; sem migração de armazenamento ou acesso ao perfil operacional.
- Esta autorização específica antecipa somente essa integração local segura. **Não inicia PDF-010**, catálogo entre computadores, gerenciador global, reformulação do N-up ou integração arbitrária de papéis no Livreto. Essas partes maiores continuam pendentes de autorização.
- Situação das duas partes: **implementadas e em revisão, sem checkpoint**. Não estão no aplicativo instalado e não constituem homologação da experimental.40. PDF-001/PDF-008/PDF-004/PDF-009 permanecem Em testes.
- Resultados e limites separados em [ORIENTACAO_E_FORMATOS_MONTAGENS_2026-11-17.md](ORIENTACAO_E_FORMATOS_MONTAGENS_2026-11-17.md). Nenhum build, instalador, impressão, commit ou push nesta rodada.

**Revisão:** 0.1 — Rascunho inicial
**Data:** 12/11/2026
**Responsável pelas prioridades:** usuário responsável pelo projeto
**Situação:** aguardando revisão e aprovação das prioridades

## 1. Objetivo

Este documento é a referência central para o planejamento do 4PrintShops, sistema de ferramentas para gráfica.

Seu objetivo é permitir que diferentes agentes de IA e computadores trabalhem no mesmo projeto, respeitando as decisões já tomadas e evitando retrabalho, implementações duplicadas ou mudanças não autorizadas de prioridade.

Documentos complementares:

- `AGENTS.md` — regras permanentes para os agentes.
- `docs/CONTEXTO_ATUAL.md` — situação recente, checkpoints e passagem de trabalho.
- `docs/ESTADO_PROJETO.md` — histórico técnico detalhado e evidências.

Este arquivo organiza o trabalho futuro; não substitui a inspeção do código.

## 2. Situação atual do projeto

**Atualização de referência em 16/11/2026:** checkout principal em `C:\Dev\4PrintShops`, branch main, checkpoint Edge local `8081175ec099b1c2513008632ea6f4e19d730322`; origin/main observado em `5c547d2`, sem push do checkpoint. Mensagens simplificadas e nova interface do Step & Repeat presentes no histórico. A versão instalada continua experimental.39. Os registros de sincronização abaixo descrevem o momento inicial de 12/11, não o HEAD atual. PDF-001, PDF-008, PDF-004 e PDF-009 permanecem **Em testes**, com homologações operacionais pendentes.

**Decisão do responsável em 16/11/2026 — instalação:** planejar a atualização da experimental.39 existente, mantendo somente a versão mais recente do Hold Test, em vez de instalação lado a lado. Preservar F4/F6/F8, preferências e dados importantes; guardar o instalador anterior fora do cache mutável, verificar atualizador/diretórios de dados e documentar atualização e reversão. Nova versão continua experimental, sem homologação de produção. Preparação final, geração de instalador, instalação, desinstalação e substituição ainda dependem de revisão/autorização específica. Procedimento proposto em `C:\Dev\4PrintShops_Auditoria_2026-11-16\PROCEDIMENTO_ATUALIZACAO_HOLD_TEST.md`.

**Decisão posterior em 16/11/2026 — reversão simplificada e dependências:** responsável confirmou retorno anterior bem-sucedido por desinstalação da versão nova pelo Windows e reinstalação normal da antiga. Adotar esse procedimento com backup dos dados e Setup .39 preservado fora do cache, SHA-256 conferido e verificação de F4/F6/F8/preferências após reinstalar. Bloqueio de downgrade direto não impede a atualização; não exigir ensaio complexo sem outro impedimento técnico concreto. Preparação isolada das dependências desktop e backups autorizada nesta continuação. Geração de instalador, instalação, desinstalação e substituição continuam aguardando autorização; sem impressão.

**Fila de usabilidade aprovada em 16/11/2026:** após validar a nova versão instalada, tratar PDF-010 (Ctrl+A); somente depois avançar às implementações maiores do catálogo compartilhado de papéis e do N-up. Esta decisão não define a ordem relativa dessas duas implementações maiores nem autoriza iniciá-las.

**Última versão registrada localmente:** 1.0.0-experimental.39
**Commit local:** `a8e25f8` — Corrige instalador e registra experimental 39
**Último commit remoto confirmado:** `7558bf5` — experimental.38
**Sincronização:** o commit da .39 ainda aguarda push.

A experimental.39 corrigiu problemas de atualização e reinstalação do aplicativo.

A auditoria específica do instalador realizada em 12/11/2026 aprovou 28 testes, sem falhas. O histórico registra 520 testes Node aprovados anteriormente pelo Work; essa suíte completa não foi repetida nesta auditoria.

O aplicativo .39 foi instalado e abriu com sucesso na gráfica, conforme os registros disponíveis. Isso não representa homologação geral de todas as funções.

O problema de recorte envolvendo o atributo PDF `/Tabs` continua pendente.

## 3. Regras de planejamento

**16/11/2026 — retomada autorizada do Step & Repeat:** etapa de interface recuperada sobre o checkpoint de mensagens 9877601 e finalizada para revisão, com rotação na área básica e detalhes técnicos recolhidos, sem alterar motores ou a lógica de N-up/Livreto. 560 testes PDF anteriores conferidos; Chrome específico e regressão ampliada aprovados. Catálogo compartilhado permanece somente proposta investigativa em PROPOSTA_FORMATOS_COMPARTILHADOS.md; nenhum cadastro ou migração implementado. Prioridades e estados das tarefas anteriores preservados, sem autorização de commit/push/distribuição. Detalhes atuais em CONTEXTO_ATUAL.md e ESTADO_PROJETO.md.

**13/11/2026 — ajuste de apresentação autorizado:** mensagens curtas de incompatibilidade para balcão, detalhes técnicos recolhidos e cópia de código nos status/conferência. Implementação local aguardando revisão; sem mudança de prioridades ou novo suporte PDF. PDF-001/PDF-008/PDF-004/PDF-009 permanecem Em testes. Validação atual: 560 testes PDF/41 arquivos e Chrome com os sete casos reais prioritários, assinaturas, XRef e recorte. Diagnóstico específico da abertura individual protegida e padronização dos demais erros sem código permanecem separados; não inferir restrições por texto. Detalhes e limites nos registros atuais de CONTEXTO_ATUAL.md e ESTADO_PROJETO.md. Sem autorização de commit/push ou distribuição nesta etapa.

1. Este documento será a lista oficial de implementações após sua aprovação.
2. Cada tarefa deve possuir um identificador permanente.
3. Nenhum agente pode excluir, alterar prioridades aprovadas ou declarar uma tarefa concluída sem justificativa, evidências e concordância do usuário.
4. Novas ideias devem entrar como **Proposta**, sem substituir automaticamente as prioridades existentes.
5. Toda tarefa deve possuir objetivo, dependências e critérios de aceitação.
6. A implementação e a homologação operacional são etapas diferentes.
7. Antes de iniciar uma tarefa, conferir o Git, o contexto atual e eventuais mudanças locais não publicadas.
8. Alterações realizadas em casa e na gráfica só serão consideradas sincronizadas depois do envio e recebimento dos commits correspondentes.
9. Não fazer push, apagar alterações, instalar aplicativos ou executar impressão sem autorização específica.
10. Ao terminar uma etapa, atualizar a situação da tarefa e registrar os resultados técnicos.

### Validação obrigatória de novas versões — decisão de 16/11/2026

Testes de instalação, atualização entre versões e recuperação após falhas passam a ser parte obrigatória da validação de cada nova versão do 4PrintShops. Aprovação de testes PDF ou geração bem-sucedida do Setup não substitui essa validação.

Priorizar condições representativas dos computadores da gráfica, reproduzidas inicialmente em instalações sintéticas e ambientes isolados:

- **INST-VAL-001 — Instalação:** instalação inicial e reinstalação, identidade e versão corretas, sem duplicação indevida do aplicativo.
- **INST-VAL-002 — Atualização entre versões:** origem e destino identificados, arquivos instalados íntegros e compatibilidade do caminho normal de atualização.
- **INST-VAL-003 — Recuperação:** falhas e interrupções controladas, rollback e verificação da integridade do estado anterior; preservar arquivos e registros em situações ambíguas.
- **INST-VAL-004 — Perfis e preferências:** preservação dos dados importantes e formatos F4/F6/F8, com backup e comparação antes/depois, sem expor perfis ou credenciais nos relatórios compartilhados.
- **INST-VAL-005 — Condições reais e resíduos:** arquivos residuais conhecidos e desconhecidos, arquivos em uso/processos ativos e espaço temporário necessário; validador rigoroso e quarentena preservada quando aplicável.

Manter resultados identificados pelos IDs acima, seguindo a rastreabilidade das tarefas PDF: versão/commit de origem e destino, identificação e SHA-256 do instalador, ambiente, cenário, resultado esperado e observado, data, evidências/logs e pendências. Distinguir testes técnicos isolados de instalação e homologação operacional efetivamente realizadas; cenário não executado permanece pendente. Registrar a síntese em CONTEXTO_ATUAL.md e o histórico técnico em ESTADO_PROJETO.md, preservando os resultados anteriores.

Esta decisão não amplia a implementação atual: concluir somente o necessário para gerar e validar a experimental.40, mediante as autorizações de geração e instalação correspondentes. Não iniciar novas funcionalidades nem ampliar a infraestrutura de testes nesta etapa.

### Estados permitidos

- **Proposta:** identificada, ainda não autorizada.
- **Planejado:** incluído na fila por decisão do responsável, sem autorização para implementar nesta etapa.
- **Aprovada:** autorizada, aguardando implementação.
- **Em desenvolvimento:** implementação iniciada.
- **Bloqueada:** depende de uma condição não atendida.
- **Em testes:** implementação existente, aguardando validação.
- **Concluída:** critérios de aceitação cumpridos e conclusão confirmada.

### Prioridades

- **P0:** bloqueio, segurança, integridade ou problema importante para continuar.
- **P1:** próxima implementação relevante.
- **P2:** melhoria planejada.
- **P3:** evolução futura.

As prioridades na primeira versão são sugestões, exceto quando houver uma decisão explícita registrada.

## 4. Lista principal de implementações

| ID | Prioridade sugerida | Situação | Tarefa |
|---|---|---|---|
| SYS-001 | P0 | Em preparação | Organizar planejamento e sincronização entre computadores |
| PDF-001 | P0 | Em testes | Corrigir bloqueio de recorte por atributo `/Tabs` |
| INST-001 | P1 | Proposta | Investigar e reduzir o pico de espaço em disco do instalador |
| QA-001 | P1 | Em testes | Homologar Step & Repeat por grade e frente e verso |
| PDF-002 | P1 | Proposta | Reconhecer e preservar corretamente caixas de corte e sangria |
| WORD-001 | P1 | Em testes | Ampliar homologação de importação DOC/DOCX |
| PDF-003 | P2 | Proposta | Permitir senha para abertura de PDFs protegidos, quando autorizado |
| QR-001 | P2 | Proposta | Importar planilhas XLSX e gerar QR Codes de contatos em lote |
| PDF-004 | P2 | Em testes | Ampliar compatibilidade de PDFs complexos — etapa limitada de recorte marcado autorizada |
| PDF-005 | P2 | Proposta | Implementar Desfazer/Refazer |
| PDF-006 | P2 | Proposta | Integrar extração de páginas ao fluxo do workspace |
| PDF-007 | P2 | Proposta | Inspeção técnica de cores, fontes, imagens e perfis |
| PDF-008 | Não definida | Em testes | Proteção consistente de PDFs assinados — implementação autorizada em 13/11/2026 |
| PDF-009 | Não definida | Em testes | Abertura segura e bloqueio de reescrita XRef incompatível — CASO-029 |
| PDF-010 | Não definida | Planejado | Atalho Ctrl+A para selecionar todas as páginas do documento ativo |
| PDF-011 | Não definida | Planejado | Agrupamento seguro com OutputIntents — antes da primeira estável, fora da experimental.42 |
| IMP-001 | P2 | Em testes | Homologar fisicamente a montagem de livretos |
| DESK-001 | P3 | Proposta | Evoluir distribuição e atualização assistida |
| IMP-002 | P3 | Proposta | Avaliar Cut & Stack, múltiplos cadernos e creep |
| LEG-001 | Sem prioridade | Preservado | Manter legado CorelDRAW separado da aplicação principal |

**Observação:** esta tabela não constitui autorização para iniciar as tarefas. A ordem definitiva será aprovada pelo usuário.

## 5. Detalhamento das próximas etapas

### SYS-001 — Continuidade e sincronização

**Objetivo:** centralizar as decisões e permitir alternar com segurança entre gráfica, computador dedicado e casa.

Critérios de aceitação:

- Criar `docs/PLANEJAMENTO.md`.
- Criar `docs/CONTEXTO_ATUAL.md`.
- Atualizar `AGENTS.md` com regras de leitura e sincronização.
- Preservar `docs/ESTADO_PROJETO.md`.
- Registrar os documentos em commit separado do código funcional.
- Enviar os commits autorizados ao GitHub.
- Clonar e conferir o repositório no computador dedicado.
- Confirmar que os agentes reconhecem as mesmas prioridades.

### PDF-001 — Correção do atributo /Tabs

**13/11/2026:** checkpoint de implementação exclusivamente local autorizado pelo usuário, mantendo **Em testes**. Auditoria real: sete origens recortadas, seis com Tabs /S preservado; operação adicional na interface sobre a mesma origem Word. Dois PowerPoint continuam bloqueados por estrutura marcada. Comparação visual independente e homologação completa pendentes. Testes anteriores aprovados não serão repetidos sem motivo. Nenhuma autorização de push, nova experimental ou correção de assinaturas. Prioridades preservadas; detalhes e limitações no checkpoint de ESTADO_PROJETO.md.

**12/11/2026:** implementação autorizada pelo usuário e preparada sobre 118032e no computador dedicado. Testes automatizados e Chrome sintético aprovados; aguardando revisão e amostras Word/PowerPoint/comuns. Não concluída nem homologada. Prioridade sugerida preservada; nenhuma autorização para gerar versão. Detalhes em ESTADO_PROJETO.md.

**Problema:** PDFs comuns, PDFs gerados pelo Word e PDFs exportados do PowerPoint apresentaram a mensagem:

`Recorte bloqueado: atributo Tabs ainda não suportado.`

PDF criado a partir de imagem funcionou no teste relatado.

**Objetivo:** aceitar os casos válidos do atributo `/Tabs` sem enfraquecer as restrições de segurança do recorte.

Critérios de aceitação:

- Inspecionar o comportamento real na versão atual.
- Validar os valores e a semântica do atributo.
- Preservar os bloqueios de documentos que não possam ser processados com segurança.
- Acrescentar testes de regressão.
- Confirmar que outras páginas e os arquivos originais permanecem intactos.
- Validar PDFs representativos do Word, PowerPoint e PDFs comuns.

Não remover indiscriminadamente todas as verificações de atributos PDF.

### PDF-008 — Proteção consistente de PDFs assinados

**13/11/2026:** ID disponível e execução autorizada explicitamente pelo usuário sobre main/f910315. Prioridades e ordem das outras tarefas preservadas. Implementação local em testes; sem autorização de commit, push, pacote ou mudança da experimental.39.

**Objetivo:** reutilizar o detector estrutural para bloquear operações de alteração/reescrita quando houver assinatura aplicada ou indícios inconclusivos, inclusive Perms/DocMDP. Abrir e visualizar permanece permitido, com bytes originais. Campos vazios continuam distintos de assinatura aplicada, sujeitos às restrições conservadoras de cada motor.

**Dependências:** detector existente de assinaturas, checkpoint PDF-001 preservado e dependências locais; nenhuma instalação nova.

**Critérios de aceitação:** proteção nos módulos internos e interface para recorte, união, rotação/exclusão/reordenação/duplicação, redimensionamento, imposição (Step & Repeat/N-up/livreto), exportação que reescreve e preparos que descartariam assinaturas; bloqueio antes de modificar/serializar; mensagens claras; regressões de documentos comuns e PDF-001 aprovadas; originais e biblioteca preservados. Não certificar validade criptográfica nem afirmar que ausência de detecção prova ausência histórica de assinatura.

**Resultado local:** política central e separação de bytes de visualização implementadas; testes Node e Chrome, incluindo worker de união e janelas de montagem/livreto, registrados em ESTADO_PROJETO.md. Homologação operacional e Electron pendentes. Não ampliar suporte a estrutura marcada, camadas ou caixas gráficas. Aguardar revisão antes de checkpoint.

### INST-001 — Consumo de espaço do instalador

**Problema:** uma tentativa de instalação da .39 encontrou falta de espaço em C: durante o processamento.

**Objetivo:** medir o espaço temporário necessário e avaliar alternativas que não comprometam a integridade ou a recuperação da instalação anterior.

Critérios de aceitação:

- Medir o consumo real das etapas de instalação.
- Identificar quais cópias temporárias são indispensáveis.
- Estimar o espaço mínimo exigido antes de iniciar.
- Impedir atualizações inseguras por falta de espaço.
- Preservar rollback, perfil e arquivos não gerenciados.
- Executar testes em ambiente isolado antes de qualquer instalação operacional.

### QA-001 — Step & Repeat e duplex

**Situação:** a montagem por linhas e colunas foi implementada na experimental.37.

Permanecem pendentes verificações operacionais e físicas.

Critérios de aceitação:

- Testar grade 3 × 3 e outras configurações.
- Conferir quantidades, marcas e posições.
- Verificar frente e verso, borda longa e borda curta.
- Comparar resultado com impressora/RIP da gráfica.
- Registrar eventuais diferenças de posicionamento.

Nenhuma impressão física sem autorização.

### PDF-002 — Corte e sangria

**Objetivo:** melhorar o tratamento de `MediaBox`, `CropBox`, `TrimBox` e `BleedBox`.

A primeira implementação de recorte modifica as caixas da página, mas não cria sangria nem apaga permanentemente conteúdo oculto.

Critérios de aceitação:

- Identificar caixas existentes com confiabilidade.
- Diferenciar corte visual de remoção definitiva de conteúdo.
- Preservar sangria válida quando tecnicamente possível.
- Informar claramente quando a sangria não estiver verificada.
- Validar exportação, montagem e comportamento no RIP.

### WORD-001 — Compatibilidade Word

**Situação:** importação de um documento Word simples foi aprovada em teste operacional. DOC e DOCX sintéticos também passaram em testes técnicos.

Pendente:

- Documentos com imagens e tabelas.
- Fontes diferentes e substituições.
- Documentos extensos.
- Paginação e layout.
- Desempenho em lotes.
- Compatibilidade de impressão.

Não considerar todos os documentos Word homologados com base no teste simples.

## 6. Outras implementações planejadas

### PDF-003 — PDFs protegidos

A conferência individual de PDFs protegidos e inválidos já foi implementada.

A próxima evolução é permitir informar senha, quando houver autorização, e preparar uma cópia adequada para processamento. Não contornar restrições nem comprometer os originais.

### QR-001 — QR Codes por planilhas

Importar planilhas XLSX de contatos e permitir:

- Mapear as colunas para campos de contato.
- Selecionar quais dados entram no QR Code.
- Gerar vCards compatíveis.
- Produzir vários QR Codes em lote.
- Exportar PNG, SVG e ZIP.
- Validar os resultados antes da exportação.

O módulo atual de QR Codes já existe, mas essa importação automatizada ainda é futura.

### PDF-004 — Compatibilidade avançada

**13/11/2026 — ajuste lexical autorizado, Em testes:** revisão identificou aceitação indevida de MCIDs reais cujo valor JavaScript parecia inteiro. Correção localizada no lexer: preservar lexema, exigir sintaxe de inteiro PDF antes de Number e manter limites/ParentTree. Sem ampliar perfil nem alterar prioridades. Novos 35 testes literais: antes, 27 passaram e oito reproduziram a falha; depois, 35 passaram. Regressão atual: **513 testes PDF em 39 arquivos aprovados**, incluindo os 478 anteriores. CASO-008/009 novamente recortados/exportados/reabertos em memória, com manifestos/streams preservados e saídas idênticas por SHA256 às anteriores. Conferência independente e acessibilidade continuam pendentes; sem checkpoint autorizado.

**13/11/2026 — etapa limitada autorizada, Em testes:** recorte de documento marcado com uma página, árvore verificável e MCIDs no Contents da página, sem remover a marcação. Prioridade P2 e ordem das demais tarefas preservadas; não representa conclusão da PDF-004. Base main/a9071ac, checkpoint local da PDF-008; PDF-001 e PDF-008 permanecem Em testes. Implementação não depende do nome ou produtor do arquivo.

**Aceitação desta etapa:** relações P/Pg/K, MCIDs e ParentTree conferidas antes/depois da serialização; streams e objetos preservados, exceto caixas da página e detalhes normais de serialização. Estruturas desconhecidas, múltiplas páginas marcadas, anotações, formulários, assinaturas, camadas e XObjects fora do perfil continuam bloqueados. Aviso explícito de recorte visual, conteúdo ainda extraível e ausência de certificação PDF/UA.

**Resultado local:** 478 testes PDF em 39 arquivos aprovados (422 anteriores + 56 novos), Chrome de recorte e dois casos reais aprovados. CASO-008: 99 elementos/47 MCIDs; CASO-009: 343/204, preservados após exportar/reabrir. Comparações da região visível no mesmo PDF.js sem diferenças RGBA. Leitor independente, tecnologia assistiva e aplicativo instalado ainda pendentes. Detalhes em ESTADO_PROJETO.md. Aguardar revisão antes de checkpoint; sem commit, push ou alteração da experimental.39 nesta etapa.

Investigar separadamente documentos com formulários, anotações, estruturas marcadas, destinos internos e outros recursos que hoje possuem restrições de processamento.

Não remover bloqueios preventivos sem testes de fidelidade.

### PDF-009 — Abertura segura do CASO-029 e gerações XRef concorrentes

**Complemento autorizado em 13/11/2026 — concorrência:** geração monotônica por tentativa de abertura, validada junto à identidade/revisão antes da publicação. Tentativa recente que falha também invalida as anteriores. Interface não restaura snapshots/mensagens antigos; prévia verifica vigência antes de publicar canvas. Correção localizada, sem ampliar compatibilidade XRef ou criar cancelamento global. Antes: 15 regressões, 14 falhas; depois: 19 novas regressões aprovadas, **547 testes PDF/40 arquivos aprovados**, Chrome XRef/CASO-029/comum, assinaturas e recorte comum/marcado aprovados. CASO-029 continua somente visualizável, sem reescrita segura. Estado Em testes; pendências de aplicativo instalado, leitor independente e homologação operacional mantidas. Prioridades preservadas e checkpoint ainda depende de revisão/autorização. Registros iniciais abaixo mantidos como histórico.

**13/11/2026:** ID disponível conferido; implementação limitada autorizada pelo usuário sobre main/9eb6a7f. Em testes, sem prioridade nova ou mudança de ordem das tarefas anteriores. PDF-001, PDF-008 e PDF-004 permanecem Em testes. Experimental.39 inalterada.

**Objetivo:** visualizar a origem sem reserialização quando o contexto carregado contém gerações concorrentes do mesmo número de objeto; recusar operações que o mecanismo atual não consegue reescrever com segurança. Sem reparar/renumerar objetos ou escolher geração por heurística.

**Resultado:** bytes de origem associados à identidade/revisão do workspace; guarda compartilhada antes de editar, serializar, extrair, unir ou preparar cópias. Mensagem explícita de visualização com reescrita bloqueada. CASO-029 abre no Chrome atual, com bytes exatos da cópia, mas edição/exportação PDF continuam incompatíveis e são recusadas antes de salvar/publicar. Geração não zero isolada continua aceita.

**Validação:** 528 testes PDF/40 arquivos aprovados (513 anteriores + 15 novos), Chrome XRef com sintético/real/comum, Chrome de assinaturas e recorte comum/marcado aprovados. Abertura instalada na .39, leitor independente e serialização efetivamente segura desse perfil permanecem pendentes. A alternativa sem object streams abriu com recuperação/avisos e não foi adotada como correção. Detalhes em ESTADO_PROJETO.md. Aguardar revisão antes de checkpoint; sem commit/push ou pacote.

### PDF-010 — Atalho Ctrl+A para selecionar todas as páginas do documento

**Situação:** Planejado, por decisão do responsável em 16/11/2026; não implementar agora.

**Problema:** Ctrl+A seleciona textos da interface quando o operador espera selecionar todas as páginas do PDF aberto.

**Objetivo:** na área de navegação e seleção de páginas, selecionar todas as páginas do documento ativo conforme o modelo de seleção múltipla existente, com indicação visual clara e sem selecionar textos da interface inteira.

**Dependências e posição na fila:** validar primeiro a nova versão instalada; executar esta melhoria de usabilidade antes das implementações maiores do catálogo compartilhado de papéis e do N-up. Conferir os contextos de foco, documento ativo e abas antes da implementação.

**Critérios de aceitação:**

- Ctrl+A na navegação/seleção de páginas seleciona todas as páginas apenas do documento ativo.
- Dentro de campos de texto editáveis, mantém o comportamento padrão de selecionar o texto do campo.
- Respeita a seleção múltipla existente e indica visivelmente quais páginas foram selecionadas.
- Não seleciona textos da interface inteira nem páginas de outros documentos ou abas.
- Selecionar não executa automaticamente operações destrutivas.
- Testar documentos de uma página e de várias páginas, troca de documento/aba e foco nos campos editáveis.

### PDF-011 — Agrupamento seguro com OutputIntents

**Decisão de 17/11/2026:** investigar e validar em etapa independente, antes da primeira versão estável. Não integrar à experimental.42, dedicada ao gerenciamento geral de tamanhos. Não iniciar implementação sem autorização específica; não muda o estado de PDF-001/004/008/009 nem inicia PDF-010/Livreto.

**Regressões de referência:** CASO-030 (cinco páginas) e CASO-031 (uma), preservados apenas na biblioteca privada, com ICC sRGB incorporado. Abrem individualmente; a guarda atual recusa a união. Dez ensaios revisados: oito recusas previstas envolvendo os novos casos e duas uniões de controles aprovadas, documento corrente e originais intactos.

**Escopo e aceitação:** preservar perfil ICC e vínculo OutputIntents, comparar compatibilidade de perfis/intenções (inclusive equivalentes, diferentes, ausentes/mistos, múltiplos e inválidos), conferir referências globais/mapeamento de páginas, marcação/acessibilidade, marcadores/metadados e integridade da saída. Verificar quantidade/ordem das páginas, streams/recursos/ICC, renderização local e atomicidade em falhas. Manter bloqueados os casos incompatíveis ou não comprovados e conservar as políticas de assinaturas, XRef, camadas e formulários. Não apagar perfil/estruturas, converter cores silenciosamente nem certificar PDF/A/PDF/X só pela presença de OutputIntents. Melhorar código/mensagem de diagnóstico sem enfraquecer guardas.

**Dependências:** revisão do desenho seguro da união e testes sintéticos/perfis controlados, além dos dois casos reais locais em ambas as ordens e controles conhecidos. Nada de PDFs reais/evidências privadas no Git. Diagnóstico: `DIAGNOSTICO_UNIAO_CASOS_030_031_2026-11-17.md`.

### PDF-005 — Desfazer e refazer

Planejar histórico de operações do PDFWorkspace, com recuperação segura de edições e controle de consumo de memória.

### PDF-006 — Extração integrada

Evoluir a extração de páginas, que ainda utiliza download direto no fluxo legado, para permitir abrir ou preparar outro documento de trabalho sem exportação automática.

### PDF-007 — Inspeção gráfica

Planejar inspeção confiável de:

- RGB, CMYK e cores especiais.
- Perfis ICC.
- Fontes incorporadas ou substituídas.
- Resolução efetiva das imagens.
- Caixas de corte e sangria.
- Alertas de pré-impressão.

Até existir análise confiável, apresentar **Não verificado**.

### IMP-001 e IMP-002 — Livretos e imposição avançada

Homologar primeiro a montagem de livreto grampeado e o duplex físico.

Somente depois avaliar funcionalidades adicionais, como Cut & Stack, múltiplos cadernos e compensação de lombada (creep).

### DESK-001 — Distribuição e atualização

Avaliar futuramente o fluxo de atualização assistida, assinatura e distribuição entre máquinas, sem reduzir proteções de integridade ou segurança.

### LEG-001 — CorelDRAW

A automação antiga de Adesivos/CorelDRAW foi retirada da aplicação principal em 30/10/2026.

Preservar código, modelos e documentação legados. Uma retomada ou integração futura dependerá de decisão separada.

## 7. Implementações importantes já existentes

Estas funcionalidades não devem ser recriadas por outro agente sem antes conferir o código:

- Gerenciador de PDFs e PDFWorkspace.
- Importação, união, exclusão e rotação de páginas.
- Miniaturas, reorganização e controles de visualização.
- Redimensionamento vetorial.
- Proteção contra perda de alterações.
- Exportação PDF e imagens.
- Montagem e imposição de PDFs.
- Step & Repeat por grade.
- Primeira etapa de Recortar página.
- Diagnóstico de PDFs protegidos em lotes mistos.
- Conversão Word com LibreOffice privado.
- Aplicativo Electron para Windows.
- Atualização transacional e rollback.
- Configurações e recursos de impressão Hold Test.
- Módulo inicial de QR Codes.

A existência no código não significa homologação universal.

## 8. Processo de execução

Antes de iniciar:

1. Consultar `AGENTS.md`, `PLANEJAMENTO.md`, `CONTEXTO_ATUAL.md` e `ESTADO_PROJETO.md`.
2. Conferir branch, HEAD, alterações locais e situação do remoto.
3. Confirmar o ID da tarefa e sua autorização.
4. Inspecionar a implementação existente para evitar refazer trabalho.
5. Definir escopo, arquivos afetados e testes de aceitação.

Durante a execução:

- Fazer alterações pequenas e isoladas.
- Preservar originais e recursos existentes.
- Executar testes proporcionais ao risco.
- Registrar impedimentos e descobertas relevantes.
- Não mudar a ordem das tarefas sem aprovação.

Ao encerrar:

1. Documentar o que mudou.
2. Registrar testes executados e não executados.
3. Atualizar o status da tarefa.
4. Atualizar o contexto de passagem de trabalho.
5. Preparar um checkpoint Git quando autorizado.
6. Informar se existem alterações não sincronizadas.

## 9. Registro de decisões

| Data | Decisão ou registro |
|---|---|
| 30/10/2026 | Legado CorelDRAW preservado e separado da aplicação principal |
| 10/11/2026 | Step & Repeat por grade registrado na experimental.37 |
| 10/11/2026 | Primeira etapa do recorte vetorial implementada |
| 11/11/2026 | Experimental.39 preparada com correções do instalador |
| 12/11/2026 | Auditoria específica: 28 testes do instalador aprovados |
| 12/11/2026 | Commit local `a8e25f8` criado para preservar a .39 |
| 12/11/2026 | Decidida a organização do trabalho entre chat de planejamento, Codex, Work e GitHub |

## 10. Mudanças futuras de prioridade

Toda alteração de prioridade aprovada deverá registrar:

- Data.
- ID da tarefa.
- Prioridade anterior e nova.
- Motivo da alteração.
- Aprovação do usuário.
- Commit que contém a mudança, quando disponível.

O histórico de decisões não deve ser apagado.

## 11. Próxima ação

Consolidar este planejamento com as listas que possam existir nas conversas do computador de casa.

Após aprovação, finalizar a documentação compartilhada, sincronizar os commits e preparar o computador dedicado.

**Nenhuma implementação nova é autorizada automaticamente por este documento.**
