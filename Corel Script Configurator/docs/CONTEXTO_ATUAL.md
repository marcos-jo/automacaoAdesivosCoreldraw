# 4PrintShops — Contexto Atual e Passagem de Trabalho

## 16/11/2026 — Revisão final e checkpoint da quarentena

- Criado checkpoint local autorizado `5899f75123a710caa0d7565ee8f157035fd62259`, Preserva residuos LibreOffice na atualizacao transacional, exclusivamente os três módulos desktop e o teste de transação. Revisão confirmou registro antes do movimento, recuperação nos estados anterior/posterior ao movimento, preservação em colisões, bloqueio de arquivos em uso e validador inalterado. A quarentena é irmã da instalação, fora de lo e de PLUGINSDIR; nenhuma limpeza de sucesso examinada a exclui. Sem risco de exclusão identificado nesse ciclo.
- Conferido log anterior de 72/72 testes aprovados, sem repetir a suíte; sintaxe dos quatro JS e diff staged --check novamente aprovados. Registros históricos abaixo mantidos integralmente, somente acréscimos documentais. Fonte pronto para incorporar em novo instalador após autorização: usar checkout deste checkpoint, verificar inclusão dos três módulos, pacote/versão/hashes e ausência de dados privados; renovar backup/perfil/F4/F6/F8, conferir Setup .39 para retorno, espaço, aplicativos fechados e ausência de transação pendente antes de instalar. Sem geração, Setup, limpeza operacional ou push; .39 preservada. Este registro e as tentativas anteriores seguem no checkpoint documental separado autorizado.

## 16/11/2026 — Quarentena transacional restrita do lock OpenCL (aguarda revisão)

- Base `2affffb2b99a5be54485f995f156b63e6c91dae8`; correção local sem commit/push. Relatos pendentes das duas tentativas abaixo preservados. Instalação operacional .39, perfil/F4/F6/F8, PDFs reais e ambos Setups não foram modificados nesta tarefa.
- `desktop/update-quarantine.mjs` reconhece exclusivamente o não gerenciado `lo/program/opencl/.~lock.cl-test.ods#`: 1–4096 bytes, UTF-8 válido, cinco campos CSV, primeiro vazio, proprietário não vazio, host local, data no formato do lock e URI file de `profile` ausente em `TEMP/4printshops-word-<alfanumérico>/profile`. TEMP é comparado por caminho real, inclusive aliases 8.3. Nome sem conteúdo reconhecido, perfil existente, links ou qualquer outro extra em lo causam recusa.
- `update-transaction.mjs` registra caminho original, tamanho e SHA-256 no diário antes de mover. Destino preservado: `<pasta da instalação>/../.4ps-upgrade-<id>/quarantine/libreoffice-opencl-lock`, fora de lo e do diretório temporário do Setup. Movimento nativo sem substituição, hash conferido com handle que recusa leitores/escritores concorrentes. Rollback reverte e confere o hash; interrupção após mover é recuperável pelo diário. Colisão/adulteração impede sobrescrita e pode exigir recuperação assistida (`recovery-required`).
- `update-registry.mjs` e o helper recusam processos ativos soffice/soffice.bin, opencltest e oosplash; fechar LibreOffice antes da atualização. Verificador/manifesto do runtime permanecem intactos e rigorosos. Retry EPERM de save() inalterado.
- 72 testes aprovados, zero falhas/ignorados: transação (incluindo 14 novos), EPERM (5), runtime LibreOffice, registro isolado, assinatura de atualização e atualizador. Sintaxe dos quatro arquivos JS e `git diff --check` aprovados. Node 26 isolado; dependências existentes externas resolvidas somente no processo de teste, sem instalação. Log: `C:\Dev\4PrintShops_Auditoria_Instalacao40_2026-11-16\quarentena-final.log`. Tentativa intermediária das duas suítes ESM não iniciou por dependência fora do checkout; execução final com resolução isolada aprovou ambas.
- Em sucesso, reter diário/quarentena e backups para revisão; não restaurar o lock em lo nem apagá-lo automaticamente. Liberação/descarte futuro depende de revisão. Nenhum Setup/build/atualização executado. Apto à revisão para checkpoint; geração e instalação posteriores exigem autorização. Homologação visual .40 pendente; PDF-001/008/004/009 continuam Em testes.

## 16/11/2026 — Segunda tentativa .40: recuperação e origem do lock conferidas

- Segunda transação `.4ps-upgrade-m8zN4A` rolled-back sem erro de recuperação; .39 instalada. Conferidos 19.515 arquivos anteriores e 76 arquivos do perfil/backup, zero divergências; F4/F6/F8 preservados, sem processo pendente.
- Lock OpenCL já estava na .39 antes da primeira tentativa, como arquivo não gerenciado; não está em nenhum dos dois Setups, recursos preparados ou win-unpacked. Conteúdo aponta a perfil temporário de conversão Word de 16/11 às 10:50; processo criador exato não identificado. Preservação do lock antigo no destino conflita com o verificador estrito de lo; filtro de build não resolveria.
- Sem alteração de código/política: aguardar revisão de tratamento transacional com preservação externa/quarentena antes de implementar. 15 testes existentes do runtime/EPERM aprovados; sem build/Setup/commit/push. Setup 78C10BE97F84DE19A64B9EF0AD76527483B2A9258994886E02CD27A0AD61F7FA preservado, reprovado para esta atualização. Relatório: `C:\Dev\4PrintShops_Auditoria_Instalacao40_2026-11-16\SEGUNDA_FALHA_LIBREOFFICE.md`.


## 16/11/2026 — Tentativa .40 abortada; .39 recuperada pelo próprio Setup

- Checkpoint dos três metadados/documentos aprovado e criado: `da04ca52e0eba49e889d801ac4ea69a15fbd6275`, Registra versao experimental.40 e seu empacotamento. Árvore limpa após o commit, main ahead 3 de origin/main, sem push. Registros desta tentativa ficam pendentes, sem outro commit.
- Todos os pré-requisitos passaram: Setup .40 no hash autorizado, .39 preservada, backup/perfil de 76 arquivos íntegros, F4/F6/F8 presentes, aplicativo fechado, sem instalador ativo/recuperação pendente e ~49,1 GiB livres. Tentativa única por usuário no mesmo destino, com ajuda do operador na UI por falhas de captura/entrada do Computer Use; sem pedido administrativo observado.
- Instalação abortada por EPERM ao renomear `journal.json.tmp` para `journal.json` em `%LOCALAPPDATA%\Programs\.4ps-upgrade-bQ4TGq`, fase copying, Exit=2. O Setup executou sua recuperação transacional, estado final **rolled-back**, sem recoveryError. Não houve desinstalação/reinstalação manual, nova tentativa, alteração de permissões/código ou novo pacote. Motivo do bloqueio do arquivo ainda indeterminado; não atribuir a antivírus/administração sem evidência.
- Conferência independente: 19.513 arquivos gerenciados e dois não gerenciados iguais aos hashes anteriores; 76 arquivos do perfil e backup iguais, atalhos e Setup/cache .39 íntegros. Apenas um registro instalado do Hold Test, versão **1.0.0-experimental.39**. Tela abortada encerrada, sem processo ativo desta tentativa. Não foi aberto o aplicativo após a falha; testes visuais/abertura de PDF na .40 permanecem **pendentes**.
- Relatório/evidências: `C:\Dev\4PrintShops_Auditoria_Instalacao40_2026-11-16\RELATORIO_TENTATIVA_EXPERIMENTAL_40.md`. Preservar diário/backups, aguardar revisão da ocorrência antes de nova tentativa/correção. Sem push, impressão, homologação ou implementação de PDF-010/catálogo/N-up. Config/fonte continuam .40 no checkpoint; instalação efetiva permanece .39.


## 16/11/2026 — Checkpoint criado e experimental.40 gerada; não instalada

- Checkpoint autorizado concluído: `be9e8e572d87abca5913f22990451b53f5aa9777`, mensagem Atualiza smoke desktop e planejamento da proxima experimental, exatamente os três documentos e três arquivos de teste/fixture aprovados. Sem push. main agora ahead 2 de origin/main; após o commit, somente metadado .40 e registros desta geração ficam pendentes, sem segundo commit.
- `electron-builder.hold-native-test.cjs`: única alteração de configuração é extraMetadata.version .39 → **1.0.0-experimental.40**. AppId/GUID/nome/perfil/atalhos/cache/políticas preservados; nenhum módulo funcional modificado. Build a partir do checkpoint + esse diff, em pasta exclusiva `C:\Dev\4PrintShops-Builds\experimental.40-2026-11-16`, usando dependências isoladas e publish never.
- Instalador: `source\dist-desktop\hold-native-test\4PrintShops-Hold-Test-1.0.0-experimental.40-Windows-x64-Setup.exe`, relativo à pasta de build; **515.067.551 bytes**, SHA-256 **47B96CE6D446EEF64B44285819F4456E29DDBA2E2017FD7473B5BBFA2FF16846**. Sidecar .sha256 junto do Setup. Config, PE e ASAR identificam .40; sem assinatura Authenticode, release ou canal publicado.
- Validação atual: 36 regressões de assinaturas/fixture, 79 testes de segurança/empacotamento/simulados e 1 transação sintética .39→.40 aprovados; desktop-smoke aprovado no novo staging. Executável empacotado --verify-distribution aprovou os 19.410 arquivos e inicialização do LibreOffice privado em perfil temporário. ASAR/recursos/fuses e carga do Setup conferidos; 111 arquivos próprios iguais ao staging, mudanças visuais presentes, sem biblioteca/perfis reais no pacote. Referência PDF anterior 560/560 serial, não repetida.
- .39 continua instalada; perfil e backup (76 arquivos), F4/F6/F8/preferências e Setup .39 reconferidos íntegros. Não houve instalação/desinstalação, impressão, limpeza manual ou alteração das pastas antigas/restauração. Atualização futura única com mesma identidade/perfil; retorno Windows desinstalar nova/reinstalar .39 preservada. Aguardar autorização para instalar e revisão de metadados/documentação antes de outro commit.
- Relatório completo e evidências: `C:\Dev\4PrintShops-Builds\experimental.40-2026-11-16\RELATORIO_EXPERIMENTAL_40.md`. PDF-001/008/004/009 Em testes; PDF-010 Planejado, sem catálogo/N-up ou novas funcionalidades. Avaliação visual/operacional e homologação continuam pendentes.


## 16/11/2026 — Desktop-smoke corrigido; aguardando revisão/checkpoint

- Base main/8081175, três documentos locais anteriores preservados. Corrigidos somente `tests/desktop-smoke.mjs`, nova fixture `tests/helpers/pdf-print-prepared-fixture.mjs` e seu contrato `tests/desktop-prepared-fixture.test.mjs`; nenhum módulo do aplicativo alterado. Roteiro antigo tentava reconstruir assinatura aplicada, operação corretamente bloqueada pela PDF-008, e clicava Prever sem abrir a sessão.
- Desktop agora reutiliza recusas/snapshots e console estrito do helper já aprovado no Edge; exportação legítima usa PDF sintético já preparado, sem assinatura e sem histórico interno forjado. Três recusas reconhecidas, bytes/objetos preservados; eventos adicionais/exceções continuam fatais. Único beforeunload esperado é validado na tentativa deliberada de fechar/cancelar; nenhum diálogo HTML pendente. Sem timeout aumentado.
- Desktop-smoke final aprovado no Electron 45.1.0; 36 regressões direcionadas aprovadas; controle negativo externo com erro inesperado reprovou com saída 1, como exigido. Sintaxe/diff --check aprovados. Suíte PDF anterior 560/560 serial preservada como referência, não repetida nem somada aos resultados atuais. Perfil original/backup (76 arquivos) iguais por SHA-256.
- Pronto para revisão/checkpoint e autorização de geração experimental; proposta inequívoca **Hold Test 1.0.0-experimental.40**, ainda sem mudar metadados .39 do config ou gerar pacote. Mesma identidade/perfil, atualização única da instalação .39; retorno por desinstalar nova no Windows/reinstalar .39 preservada. Instalação só após verificar pacote/autorização. Relatório: `C:\Dev\4PrintShops_Auditoria_DesktopSmoke_2026-11-16\RELATORIO.md`.
- GitHub main reconferido por consulta somente leitura em 5c547d2; main local 8081175 ahead 1. Quatro modificados (três docs anteriores e smoke), dois novos testes/fixtures, nada staged; sem commit/push/instalador/impressão/atualização. PDF-001/008/004/009 Em testes; PDF-010 Planejado, catálogo/N-up não iniciados. Experimental.39, formatos, biblioteca, checkout antigo e restauração intactos.


## 16/11/2026 — Dependências desktop isoladas e reversão simplificada

- Reversão adotada conforme experiência prática confirmada pelo responsável: desinstalar a nova versão pelo Windows e reinstalar normalmente a .39, com backup e conferência de F4/F6/F8/preferências. Bloqueio de downgrade direto não impede atualização; sem exigência de ensaio complexo na ausência de outro impedimento concreto.
- Setup .39 e perfil (76 arquivos) preservados fora do cache em `%LOCALAPPDATA%\4PrintShops-Backups\pre-atualizacao-2026-11-16-65bd9755-c29e-4678-8345-3dd88f1889be`, com hashes conferidos e dados originais inalterados. F4 328×478, F6 318×318 e F8 238×328 mm confirmados na cópia privada; nenhum perfil incluído no Git.
- Dependências preparadas em `C:\Dev\4PrintShops-Tools\desktop-prep-2026-11-16`: Electron 45.1.0, builder 26.18.0, updater 6.8.11, recursos privados LibreOffice/PDFium e ferramentas do builder verificados. Fonte de staging do commit 8081175, sem misturar os três documentos locais pendentes nem alterar node_modules do checkout. Canvas aprovado em Node 26 e Electron, sem rebuild; 75 testes direcionados aprovados.
- Pendência: desktop-smoke reprovou por esperar preparo seletivo de assinatura aplicada; botão atual mostra Ver restrições e não abre esse diálogo. Não modificar a proteção para atender o roteiro antigo. Relatório em `C:\Dev\4PrintShops_Auditoria_2026-11-16\RESULTADO_PREPARACAO_DESKTOP.md`; revisar teste de forma localizada antes de recomendar liberação da geração/instalação. Não houve correção de teste/código nesta preparação, instalador, atualização, desinstalação, impressão, commit ou push. .39, biblioteca e formatos preservados; tarefas PDF mantêm estados anteriores.


## 16/11/2026 — Planejamento da atualização única e Ctrl+A

- Checkpoint Edge concluído: `8081175ec099b1c2513008632ea6f4e19d730322`, main um commit à frente do origin/main observado (`5c547d2`), sem push. Árvore limpa antes desta atualização exclusivamente documental.
- Responsável substituiu a preferência lado a lado por atualizar a experimental.39 existente, preservando perfil/F4/F6/F8 e instalador anterior. Procedimento e diagnóstico em `C:\Dev\4PrintShops_Auditoria_2026-11-16\PROCEDIMENTO_ATUALIZACAO_HOLD_TEST.md`; preparação final e instalação aguardam revisão. Nenhuma dependência instalada, pacote gerado ou aplicativo/perfil modificado.
- PDF-010 (Ctrl+A no documento ativo) registrado **Planejado**: depois da validação da nova versão instalada e antes das implementações maiores do catálogo compartilhado e N-up. Sem implementação. PDF-001/PDF-008/PDF-004/PDF-009 continuam Em testes.


## 16/11/2026 — Validação Edge de assinaturas ajustada; aguardando revisão

- Base main/5c547d2, HEAD e origin/main iguais, árvore inicialmente limpa. Ambiente aprovado em C:\Dev\4PrintShops; Node 26.13.1 isolado já validado na etapa anterior com 560 testes PDF em execução serial. Essa suíte e o harness Electron 45 não foram repetidos neste ajuste exclusivo de testes.
- ConsoleAssinaturas, em tests/helpers/pdf-signature-console.mjs, exige exatamente um console.error por recusa de giro nos cenários aplicada, multipla e perms: janela temporal da ação, origem girarPagina em pdfs.js, frames da proteção/workspace, mensagem exata, codigo PDF_ASSINATURA_PROTEGIDA e assinaturaDetectada=true. Validação final exige os três cenários em ordem e nenhum registro adicional; Runtime.exceptionThrown nunca é aceito como recusa esperada. Console da aplicação não foi alterado ou silenciado.
- Integrados o roteiro pdf-browser-smoke e o chamador pdf-signature-policy-browser; removido desse último o filtro genérico por texto. Helper compartilhado compara SHA-256 dos bytes de entrada/visualização, revisões, páginas e manifesto dos objetos antes/depois de cada giro recusado. Espião de capturarEstado somente observa a instância e delega sem mudar retorno/estado; restaurado inclusive no controle negativo. Nenhuma proteção ou implementação de rotação modificada.
- Validação: cinco arquivos de teste com sintaxe aprovada; 13 testes Node do verificador aprovados. Edge 157.0.4341.60, perfil isolado, ambos os roteiros aprovados com três recusas esperadas/zero inesperados e documentos sintéticos preservados. Controle negativo PDF_TESTE_CONSOLE_INJETAR_ERRO=1 introduziu um console.error adicional dentro da primeira recusa: roteiro reprovou com saída 1 e dois registros, como exigido. Duplicatas idênticas, exceções não tratadas, origem/código/mensagem/momento divergentes e registros fora da ação também são recusados pelos testes do verificador.
- Evidências novas fora do checkout: C:\Dev\4PrintShops_Auditoria_Edge_2026-11-16. Sem biblioteca real nos cenários, dependência desktop, instalador, impressão, commit ou push. PDF-001/PDF-008/PDF-004/PDF-009 permanecem Em testes, prioridades e experimental.39 preservadas. Próximo passo: revisão do usuário antes de checkpoint; homologação instalada/operacional continua pendente.

## 16/11/2026 — Retomada do Step & Repeat; interface preparada, revisão pendente

- Recuperação após desligamento: main/9877601052aafe3415e55a1c5aa5814a2a8bb4a2, cinco commits locais à frente de origin/main. O checkpoint das mensagens já existe; não havia alterações pendentes daquela tarefa. Encontrados cinco arquivos modificados e um teste novo do Step & Repeat, todos reconhecidos na conversa. Estado recuperado preservado em .desktop-test/repeat-retomada-20261116-srmIhG/; nada descartado ou refeito.
- Parte A já implementada antes da interrupção: único seletor de rotação realocado para a área básica somente no Step & Repeat, com rótulo Girar arte, ao lado da orientação da folha; Colunas/Linhas e frente/verso mantidos. Resumo curto de folha/quantidade/escala; explicações e diagnósticos completos em Ver detalhes técnicos. Avisos geométricos, aceites e guardas permanecem ativos. CSS responsivo exclusivo de repeat; N-up restaura os mesmos controles e a apresentação anterior. Nenhum cálculo, algoritmo de imposição, motor PDF ou Livreto modificado.
- Evidências anteriores recuperadas: 560 testes PDF/41 arquivos aprovados; Chrome específico com antes/depois, orientação/rotação independentes, grades, duplex, teclado, navegação/guias e exportação; Chrome de assinaturas aprovado. Não repetida a suíte Node, pois os dois arquivos de produção permanecem idênticos ao estado recuperado validado. A regressão Chrome ampliada havia ficado sem resultado final; foi concluída agora, com aprovação. O teste específico foi complementado para conferir o plano antes/depois de abrir detalhes, rodapé na resolução menor, origem intacta e reabertura da exportação também no PDF.js.
- Resultados atuais: Chrome ampliado aprovou apresentação básica/avançada, escalas, geometria, marcas, erros/retry, cancelamento e regressões N-up/Livreto. Chrome específico aprovado em 1366×900, 1024×768 e 720×640, com exportação sintética de duas páginas e reabertura em pdf-lib/PDF.js. N-up teve estado completo idêntico ao baseline; Step & Repeat manteve plano/configuração iniciais idênticos. Nenhuma diferença de cálculo encontrada. Testes são do fonte atual no Chrome, não do aplicativo instalado.
- Capturas anteriores mantidas em .desktop-test/repeat-interface-byi02q/ (antes), .desktop-test/repeat-interface-iJZ2Dj/ (depois) e nova execução .desktop-test/repeat-interface-uTV7Ni/. Regressão retomada em .desktop-test/repeat-retomada-20261116-montagem/ e respectivo log. Não usar datas internas do host dos artefatos como data de homologação; a retomada foi em 16/11/2026.
- Parte B somente investigada: já existem PDFFormatos e adaptador localStorage na chave 4printshops.pdf.formatos.v1, com IDs/nome/largura/altura em mm; fallback em memória não persiste após encerrar. Proposta separada em docs/PROPOSTA_FORMATOS_COMPARTILHADOS.md. Nenhum catálogo, menu, migração, cadastro ou seletor compartilhado implementado/alterado. Perfis reais não foram abertos para inspecionar cadastros.
- Próximo passo: revisão do usuário antes do checkpoint. Homologação no aplicativo instalado, uso pelos operadores e papel/impressão permanecem pendentes. PDF-001/PDF-008/PDF-004/PDF-009 e as homologações anteriores continuam Em testes; prioridades, experimental.39, biblioteca e checkpoints preservados. Sem commit/push, dependência, instalador ou impressão nesta retomada. Registros abaixo são históricos.

## 13/11/2026 — Mensagens de balcão: ajustes autorizados, aguardando revisão

- Base main/23c1e39a4fcf7dfcf790c9ffa3c4b79122c52b3a, quatro commits locais à frente de origin/main, árvore inicialmente limpa. Esse é o checkpoint local já criado da PDF-009; os registros abaixo descrevem etapas anteriores. PDF-001/PDF-008/PDF-004/PDF-009 permanecem Em testes; prioridades e experimental.39 preservadas.
- Novo componente de apresentação pdf-feedback.js: título/frase simples e Ver detalhes recolhido, diagnóstico original, operação e código quando disponível. Copiar código envia somente o código; se a API não estiver disponível ou recusar, oferece seleção/cópia manual. Sem novos modais, classificação por texto ou alteração do documento ao consultar detalhes. Erros desconhecidos ficam genéricos no resumo, sem expor caminhos ou termos técnicos; não são classificados como PDF restrito.
- Reutilizados status do editor, conferência de união e painéis existentes de assinatura/montagem. Camadas e recorte marcado receberam apenas códigos nos pontos de recusa existentes; condições mantidas. Metadado de assinatura detectada acompanha os erros pelo worker para distinguir indícios inconclusivos. Nomes duplicados nos erros de união removidos; nome original segue no item e na origem. Ver restrições abre o diagnóstico de assinatura, sem iniciar preparo. Campo vazio continua distinto e mantém os caminhos permitidos.
- Recorte mantém aceite explícito para remover ajustes de corte/sangria; aviso curto explica que ocultar não apaga conteúdo. Limitações de acessibilidade permanecem em detalhes. A confirmação de cópia para impressão não afirma invalidar uma assinatura inexistente. Na interface de montagem, somente resumo de anotação e apresentação dos códigos foram alterados; nenhum motor, geometria ou fluxo de Step & Repeat/N-up/Livreto modificado.
- Validação atual: **560 testes PDF em 41 arquivos aprovados**, zero falhas/cancelados/ignorados (547 anteriores + 13 novos). Chrome específico: sete casos reais CASO-003/005/006/011/012/013/029 e dois sintéticos (comum/assinado), com detalhes, cópia do código, preservação do estado, aceite de sangria obrigatório e cancelamento sem aplicação. Área de transferência simulada no roteiro; nenhum conteúdo copiado para o Windows. Chrome de assinaturas, XRef/CASO-029/comum e recorte comum/marcado também aprovados. Não equivale a aplicativo instalado ou leitor independente.
- Evidências novas ignoradas: .desktop-test/mensagens/, .desktop-test/mensagens-chrome-EFjd3A/ e .desktop-test/xref-chrome-Wz9b0S/. Nenhuma origem, inventário ou evidência anterior da biblioteca alterada. Pendentes: diagnóstico específico na abertura individual de PDF protegido (sem inferir restrição pelo erro genérico da biblioteca), padronização dos demais erros sem código e homologação com operadores/no aplicativo instalado. Sem novo suporte de senha, camadas ou serialização XRef. Sem commit/push, dependência, impressão ou instalador. Aguardar revisão antes de checkpoint.

## PDF-009 — 13/11/2026, concorrência da abertura corrigida, Em testes

- Autorização específica após revisão: corrigir publicação obsoleta em abrir(), sem ampliar compatibilidade XRef. Base main/9eb6a7f, três commits locais à frente de origin/main; preservados os 15 modificados/quatro novos anteriores e índice vazio. Agora 16 modificados/quatro novos, pois pdf-preview.js também precisa impedir a publicação de canvas atrasado. Experimental.39 e prioridades inalteradas; PDF-001/PDF-008/PDF-004/PDF-009 permanecem Em testes.
- Cada tentativa incrementa geracaoAbertura antes da leitura, inclusive quando falha. Identidade/revisão e geração são conferidas após os awaits e imediatamente antes da única publicação; bytes, páginas e diagnósticos são preparados localmente. Erros antigos recebem PDF_ABERTURA_OBSOLETA. Geração não integra o snapshot restaurável. A tentativa recente inválida conserva o documento corrente e não revalida a abertura anterior.
- pdfs.js verifica geração e versão efetivamente publicada, descarta respostas/erros obsoletos sem rollback, mudança de mensagens ou navegação; ocupação continua pertencendo à tentativa da interface. pdf-preview.js recebe verificação opcional antes de publicar o canvas/adotar a fonte ou restaurar uma prévia; demais chamadores mantêm o padrão existente. A guarda de união continua independente e uma tentativa de abertura a invalida desde o início.
- Regressões antes: 15 casos, um aprovado/14 falhas. Guarda inicial: 15 aprovados; dois casos adicionais de prévia expuseram um canvas tardio (um aprovado/uma falha), corrigido antes da regressão final. Acrescentados também dois casos de load tardio. Resultado final: 19 novas regressões aprovadas; **547 testes PDF em 40 arquivos aprovados, sem falhas/cancelados/ignorados**. Contagens intermediárias são subconjuntos, não somar.
- Chrome atual: CASO-029/sintético XRef visualizados, origens idênticas, bloqueios sem downloads e worker recusado; comum girado/exportado/reaberto. Assinaturas PDF-008 e recorte PDF-001/PDF-004 sintético aprovados (12 viewports e 1.285.504 componentes RGBA iguais no marcado). CASO-029 também reaberto no PDF.js Node: uma página/116 operadores, zero saves; SHA256 c816a5d2208cc8860f86af0814d0d430ccd25b994847531bf0de4239b229f66b preservado. Biblioteca não modificada.
- Evidências incrementais ignoradas em .desktop-test/pdf009-abertura/ e .desktop-test/xref-chrome-RhL356/. Limite: operações obsoletas podem terminar em memória; não há cancelamento global ou isolamento/timeout rígido. Sem nova capacidade de reescrita XRef; aplicativo instalado, leitor independente e homologação operacional pendentes. Aguardar revisão antes de checkpoint; sem commit/push, instalação, impressão ou pacote. Registros abaixo são históricos.

## PDF-009 — 13/11/2026, CASO-029: abertura segura, reescrita bloqueada, Em testes

- Base main/9eb6a7f77161ff5c72ec93f3e11aa0ff959fe64b, três commits à frente de origin/main; Git limpo no início. Checkpoint PDF-004 já criado, conforme autorização anterior. ID PDF-009 disponível, sem prioridade definida e sem alterar prioridades/ordem anteriores. PDF-001/PDF-008/PDF-004 continuam Em testes.
- Diagnóstico: o arquivo recebido mantém revisões incrementais com Root 3 0 R e 3 1 R. pdf-lib carrega ambas no contexto. O aplicativo descartava os bytes de comuns e serializava antes da primeira prévia: a XRef gerada fazia a entrada do objeto 4 apontar a 3 1 obj, causando a recusa pelo PDF.js. Não era necessário reescrever só para visualizar a origem.
- Correção limitada: guarda de gerações concorrentes em pdf-xref-policy.js; cache seletivo dos bytes originais para assinatura protegida ou contexto incompatível, associado a PDF/revisão e incluído no snapshot/rollback. Prévia recebe cópia dos bytes enquanto identidade/revisão correspondem. Comuns sem esse risco mantêm o fluxo existente, sem cache extra; edições suportadas continuam vindo do workspace atualizado. Bloqueios de assinatura mantidos e anteriores ao novo diagnóstico.
- Gerações concorrentes são recusadas antes de editar/exportar, unir/extrair ou preparar cópias nos módulos internos pertinentes, incluindo worker via rascunho de união. Nada de reparo ou seleção de maior geração. Interface informa “PDF aberto para visualização” com reescrita bloqueada; ações não publicam sucesso ou download ao recusar. Step & Repeat, N-up e Livreto não alterados.
- Testes: antes da correção, fixture final com 12 testes, dois aprovados e dez falhas que demonstravam falta de cache/guardas. Suíte final: **528 testes PDF em 40 arquivos, todos aprovados**, incluindo 15 novos e os 513 anteriores. Chrome: sintético XRef, CASO-029 e comum; preservação de bytes, bloqueios sem downloads, worker e giro/exportação/reabertura do comum. Chrome de PDF-008 e de recorte PDF-001/PDF-004 aprovados. Registros em .desktop-test/xref-fix/ e evidências do novo browser em .desktop-test/xref-chrome-WDoszA/.
- CASO-029: visualização usa os mesmos 378.402 bytes e SHA256 c816a5d2208cc8860f86af0814d0d430ccd25b994847531bf0de4239b229f66b, zero saves; PDF.js Node lê uma página/116 operadores e Chrome renderiza. Giro/recorte/exportação recusados com PDF_XREF_INCOMPATIVEL sem mutação. Salvar sem object streams em ensaio isolado abriu com avisos/recuperação de XRef, sem comprovação de integridade; alternativa não adotada nem publicada.
- Limites: diagnóstico detecta a colisão de números/gerações conhecida, não certifica toda XRef ou conformidade PDF. Edição e exportação desse perfil continuam bloqueadas, mesmo sem alterações; não se adicionou “salvar original” como novo fluxo. .39 instalada existe, mas não foi executada; testes são do fonte atual em Chrome/Node. Sem instalador, impressão, dependência, versão nova, commit ou push. CASO-025–028 e seus seis PNGs não foram testados/corrigidos. Biblioteca e inventários anteriores preservados.
- Próximo passo: revisão do usuário antes de checkpoint; eventual suporte de reescrita exigirá proposta e regressões próprias. PDF-009 permanece Em testes.

## PDF-004 — 13/11/2026, validação lexical de MCIDs corrigida, Em testes

- Base main/a9071ac, dois commits locais à frente de origin/main. Preservados os nove arquivos modificados e quatro novos da implementação anterior; índice vazio. Autorização específica para corrigir a validação lexical, sem ampliar o perfil marcado ou mudar prioridades.
- Lexer mantém o lexema numérico, exige `[+-]?\d+` para MCID antes de converter e depois confere inteiro seguro, não negativo, menor que 10.000 e não duplicado. ParentTree e demais vínculos continuam obrigatórios. Inteiros com sinal/zeros à esquerda, inclusive -0 (valor zero), mantidos; reais com ponto, arredondamento ou underflow recusados. Não alterada a interpretação global de números pela pdf-lib.
- Novos 35 casos com tokens literais nos streams, conferidos após serialização para excluir normalização prévia: antes 27 aprovados e oito falhas reproduzidas; após correção 35 aprovados. Direcionados de recorte: 175 aprovados. Suíte PDF atual: **513 testes em 39 arquivos, todos aprovados, sem falhas/cancelados/ignorados**, incluindo as regressões dos 478 anteriores, PDF-001, PDF-008 e concorrência. Contagens direcionadas são subconjuntos, não somar à suíte.
- CASO-008 conserva 99 elementos/47 MCIDs; CASO-009, 343/204. Novo recorte com margens 10 mm/exportação/reabertura em memória aprovado; manifestos e bytes de streams conferidos. SHA256 das duas saídas idênticos aos resultados Node anteriores, sem criar outros PDFs nem sobrescrever evidências. Os 24 originais e 24 cópias continuam íntegros por SHA256. Comparação Chrome anterior permanece evidência do mesmo resultado binário, sem nova execução visual nesta correção.
- Alterados nesta correção somente pdf-crop-structure.js, tests/pdf-crop-structure.test.mjs e os três documentos de continuidade. Evidências novas ignoradas em .desktop-test/pdf004-mcid-lexical/{antes.tap,depois.tap,direcionados.tap,regressao-pdf.tap,reais.json}. Código das demais proteções preservado, experimental.39/dependências sem mudanças. Nenhum commit/push, instalação, impressão ou instalador.
- Limites continuam cooperativos, sem isolamento de memória ou prazo rígido. Mantidas as pendências de leitor independente, ferramenta assistiva e aplicativo instalado; não certificar PDF/UA nem homologação de produção. PDF-004, PDF-001 e PDF-008 permanecem Em testes. Aguardar revisão antes do checkpoint. Registros abaixo são históricos.

## PDF-004 — 13/11/2026, recorte marcado limitado, Em testes

- Base confirmada: main/a9071acf523b939ccfd5f807dcb427fe6b2bddde, árvore limpa no início, dois commits locais à frente de origin/main/118032e. a9071ac é o checkpoint local autorizado da PDF-008, incluindo concorrência da união; os registros abaixo descrevem etapas anteriores. Nenhum push efetuado nesta etapa.
- Autorização específica para compatibilidade de recorte com estrutura marcada verificável. Novo diagnóstico limitado a uma página, ParentTree Nums, elementos indiretos e MCIDs no Contents; valida referências, pais, página, índices, vínculos reversos, duplicações e ciclos, com limites de recursos. Não repara nem remove marcação. Suporte independente de nome/produtor.
- Cópia integral e recorte por caixas mantidos; manifesto de objetos/streams antes/depois da preparação e da exportação. RoleMap, MarkInfo, idioma, ActualText, Alt e Tabs preservados. Bloqueios da PDF-008 e validação de Tabs da PDF-001 mantidos. Aviso de conteúdo oculto ainda extraível/acessível e ausência de certificação PDF/UA ampliado; sangria continua exigindo aceite quando aplicável.
- Validação final: 478 testes PDF em 39 arquivos, todos aprovados, sem ignorados; inclui 56 novos casos e as regressões de assinaturas, Tabs e concorrência. Chrome: 12 viewports existentes, botão integrado e sintético marcado + CASO-008/009 com cancelar/aplicar/exportar/reabrir e aviso. CASO-008 conserva 99 elementos/47 MCIDs; CASO-009 conserva 343/204. Streams/manifestos conferidos; regiões renderizadas com zero diferenças. Hashes dos originais e cópias preservados.
- Evidências locais ignoradas: .desktop-test/pdf004-recorte/ (relatórios Node, PDFs derivados, imagens antes/região original/depois e JSON em chrome/). Não incorporar biblioteca ou dados reais ao Git. Nenhuma instalação, impressão, dependência ou mudança da experimental.39.
- Limitações: somente perfil estrutural restrito; múltiplas páginas marcadas, MCR/OBJR/XObjects, atributos de layout e demais estruturas não reconhecidas ficam bloqueados. O conteúdo oculto não é apagado nem filtrado da leitura lógica. Comparação no mesmo PDF.js e inspeção das imagens não homologam leitor independente, PDF/UA, tecnologia assistiva, cor/RIP ou aplicativo instalado.
- Próximo passo: revisão do usuário e homologação complementar; checkpoint requer nova autorização. PDF-004 nesta etapa, PDF-001 e PDF-008 permanecem Em testes, sem mudança de prioridades. Alterações desta etapa ainda sem staging/commit.

## PDF-008 — 13/11/2026, concorrência da união corrigida, Em testes

- Autorização adicional limitada à publicação de união obsoleta, testes e comentário do roteiro de assinaturas. Base permanece main/f910315, um commit local à frente de origin/main; alterações anteriores da PDF-008 preservadas, sem commit ou push.
- Reprodução permanente criada antes da correção: união suspensa na serialização, nova abertura comum/assinada ou edição, depois liberação. Antes: oito testes, um aprovado e sete falhas. Depois: geração monotônica por operação, identidade do PDF e revisão verificadas antes da publicação; erros tardios são classificados como PDF_UNIAO_OBSOLETA. A interface não restaura snapshot antigo nem publica sucesso quando a operação perdeu validade.
- Validação atual: 141 testes direcionados aprovados; depois, **422 testes PDF em 38 arquivos aprovados**, sem falhas/ignorados. Os 141 são subconjunto, não somar aos 422. Incluídos oito casos de concorrência e três de interface; regressões de assinaturas PDF-008 e recorte Tabs PDF-001 preservadas. Chrome de assinaturas e conferência progressiva aprovados. Cinco arquivos JS/MJS desta etapa com sintaxe aprovada e git diff --check aprovado.
- Relatórios locais em .desktop-test/pdf008-concorrencia/antes.tap, depois-direcionados.tap e regressao-pdf.tap. Biblioteca de PDFs reais não alterada; nenhum instalador, dependência, impressão ou mudança da experimental.39.
- Limites: geração da união não é um novo mecanismo global de cancelamento. Preparações obsoletas podem terminar em memória, mas não publicam resultados; temporários são locais e liberados pelo coletor, sem destruir documentos/tarefas alheios. Outras operações assíncronas fora da união não foram refatoradas. Electron/app instalado e homologação operacional continuam pendentes.
- Aguardar nova revisão antes do checkpoint. PDF-001 e PDF-008 permanecem Em testes; prioridades e planejamento preservados.

## PDF-008 — 13/11/2026, implementação local Em testes

- Base: main/f910315 (checkpoint PDF-001), um commit local à frente de origin/main/118032e. Checkout limpo no início. PDF-001 permanece Em testes; seu commit não foi alterado nem enviado.
- Usuário autorizou PDF-008: política conservadora central no detector existente, aplicada antes das operações de alteração/reescrita. Abrir/visualizar assinados usa cópia dos bytes originais sem save(); somente documentos protegidos mantêm esse cache extra. Campo vazio não é assinatura aplicada; restrições de formulários permanecem.
- Cobertura: recorte e adoção de cópia preparada; união no workspace/rascunho/worker; giro, exclusão, reordenação, movimento, duplicação; redimensionamento; Step & Repeat, N-up, livreto; exportação PDF/extração legada; preparos de impressão/vetorial/imagem. O antigo Resolver não remove mais assinatura aplicada ou inconclusiva. Interface de Step & Repeat/N-up não alterada.
- Regressões Node e Chrome executadas com dependências existentes. Relatórios desta tarefa em .desktop-test/pdf008; detalhes, contagens finais e limitações no registro PDF-008 de ESTADO_PROJETO.md. Biblioteca anterior: 24 originais e cópias/11 resultados conferidos por SHA256, sem alterações; leitura dos seis sintéticos confirma cinco bloqueados para reescrita e um campo vazio tratado separadamente.
- Limitações: detecção estrutural, sem verificação de confiança/certificados; não se afirma segurança absoluta ou que um PDF nunca foi assinado. Electron/instalado e documentos arbitrários ainda não homologados. Nada impresso, instalado ou empacotado; experimental.39 inalterada.
- Alterações ainda sem commit; não houve push, instalação de dependências nem mudança de prioridade de outra tarefa. Próximo passo: revisão do usuário e autorização específica para checkpoint. Registros abaixo são históricos.

## Checkpoint autorizado — 13/11/2026, PDF-001 Em testes

- Usuário autorizou commit exclusivamente local da implementação PDF-001 e seus testes/documentos, com mensagem `Corrige atributo Tabs no recorte de PDFs`. Base antes do commit: main/118032e, sincronizada com origin/main. Hash será informado após a criação; sem push autorizado.
- Código funcional inalterado desde a implementação revisada. A versão de distribuição continua experimental.39; não houve novo pacote, instalação ou impressão.
- Validações anteriores: 84 testes de recorte (74 novos), regressões relacionadas, 28 testes transacionais e Chrome sintético aprovados. Não repetidos para este checkpoint.
- Auditoria local de 24 PDFs autorizados: sete documentos únicos com recorte aplicado/exportado/reaberto pelos módulos, seis com Tabs /S preservado; uma operação adicional pela interface no mesmo CASO-002 (Word), totalizando oito PDFs de recorte. Dois documentos PowerPoint (CASO-008/009) continuam corretamente bloqueados por StructTreeRoot; não homologados para recorte.
- Revisão final conferiu dimensões esperadas dos sete recortes, Tabs, reabertura e hashes dos 11 resultados. Os 24 originais e as cópias preservadas permaneceram intactos. Comparação visual completa em leitor independente e homologação operacional continuam pendentes; PDF-001 permanece Em testes.
- Biblioteca, inventários, PDFs e evidências ficam exclusivamente em .desktop-test/casos-reais, ignorados pelo Git. O commit inclui somente os seis arquivos autorizados, sem dados ou nomes de clientes.
- Risco separado e preexistente: inspeção e estrutura sintética em memória mostraram divergência entre detector de assinaturas e diagnóstico de montagem para Perms/DocMDP sem formulário/anotações. A montagem desse caso não foi executada. Rotação e união não possuem bloqueio universal; preservação criptográfica não demonstrada. Nenhuma correção de assinaturas iniciada, nenhuma prioridade alterada.
- Próximo passo: aguardar nova autorização do usuário. Checkpoint de implementação não significa homologação nem autorização de push ou de nova experimental. Os registros abaixo são históricos.

## Checkpoint — 12/11/2026, computador dedicado, PDF-001

- Revisão complementar: diff de código somente lido; nenhuma suíte repetida. Validação de PDFName e bloqueios preservados. A/W pertencem ao PDF 2.0 e às extensões Adobe 1.7; não há certificação de versão/extensão. Nenhuma amostra real encontrada no projeto. Pastas locais ignoradas preparadas em .desktop-test/pdf001-real/{originais,resultados,evidencias}; aguardando caminhos e autorização de amostras Word/PowerPoint/outra ferramenta. Riscos e roteiro no novo registro de revisão de ESTADO_PROJETO.md. Tarefa continua Em testes.

- HEAD e origin/main: 118032e; a8e25f8 (.39) presente no histórico publicado. As seções antigas abaixo registram o estado anterior à sincronização.
- PDF-001 autorizada pelo usuário e implementada localmente; em testes, aguardando revisão. Somente prepararRecorte ganhou validação explícita de Tabs R/C/S/A/W, com preservação e bloqueios anteriores mantidos.
- 84 testes de recorte aprovados; regressões relacionadas e 28 testes transacionais aprovados. Chrome headless aprovou recorte/exportação/reabertura com Tabs /S. Sem instalação de dependências; usadas bibliotecas existentes e dependências transacionais isoladas.
- Alterações ainda não publicadas: public/js/tools/pdf-workspace.js, tests/pdf-crop.test.mjs, tests/pdf-crop-browser.mjs e estes três documentos. Sem commit/push.
- Pendências: revisão do usuário e PDFs representativos Word/PowerPoint/comuns; estrutura marcada e anotações reais continuam bloqueadas. Não preparar nova experimental antes da revisão. Detalhes no checkpoint PDF-001 de ESTADO_PROJETO.md.


**Atualizado em:** 12/11/2026
**Ambiente de referência:** gráfica
**Projeto:** 4PrintShops Hold Test
**Branch principal:** `main`

## 1. Finalidade

Este documento registra a situação mais recente do projeto para permitir a continuidade do desenvolvimento entre a gráfica, o computador dedicado de testes e o computador de casa.

Consultar também:

- `AGENTS.md` — regras permanentes.
- `docs/PLANEJAMENTO.md` — prioridades e tarefas planejadas.
- `docs/ESTADO_PROJETO.md` — histórico técnico detalhado.

Este documento não substitui a inspeção do Git ou dos arquivos reais.

## 2. Situação do Git

Repositório:

`https://github.com/marcos-jo/automacaoAdesivosCoreldraw`

Pasta principal:

`Corel Script Configurator/`

**Último commit local informado:**

`a8e25f8 — Corrige instalador e registra experimental 39`

**Último commit remoto confirmado:**

`7558bf5 — Gera e valida experimental 38 com recorte de pagina`

Na última conferência, a branch local estava um commit à frente de `origin/main`.

O commit da experimental.39 ainda não havia sido enviado ao GitHub.

Após esse checkpoint, foi iniciado o trabalho de documentação:

- `docs/PLANEJAMENTO.md` — criado manualmente, aguardando commit.
- `docs/CONTEXTO_ATUAL.md` — documento em preparação.
- `AGENTS.md` — atualização das regras de continuidade ainda pendente.

Antes de alterar arquivos, conferir novamente `git status -sb` e `git log -1 --oneline`.

## 3. Versão atual

**Versão experimental mais recente na gráfica:** 1.0.0-experimental.39.

Principais alterações:

- Correção das regras de atualização entre versões experimentais.
- Suporte à reinstalação da própria versão.
- Tratamento seguro de determinados resíduos de desinstalação.
- Preservação de arquivos não gerenciados.
- Melhoria da mensagem de erro e do diagnóstico de rollback.

A documentação do Work registra 520 testes Node aprovados e uma instalação bem-sucedida da .39.

Em 12/11/2026, foram executados manualmente os 28 testes específicos de `tests/update-transaction.test.mjs`:

- 28 aprovados.
- 0 falhas.
- 0 ignorados.

A verificação de sintaxe passou.

Esses resultados não representam homologação completa do sistema em produção.

## 4. Situação dos computadores

### Computador de trabalho da gráfica

- Contém o checkout local com a experimental.39 registrada no commit `a8e25f8`.
- A aplicação .39 foi instalada e aberta com sucesso, conforme os relatos e registros existentes.
- Os arquivos de planejamento estão sendo preparados neste ambiente.
- Não realizar intervenções de instalação ou limpeza sem autorização.

### Computador dedicado ao desenvolvimento

- Foi preparado pelo usuário para desenvolvimento e testes.
- Possui aproximadamente 100 GB livres, conforme relato inicial.
- VS Code com Codex e aplicativo ChatGPT instalados e conectados.
- Ainda precisa receber o repositório sincronizado.
- Será utilizado para implementar, testar e gerar novas versões antes da distribuição.

### Computador de casa

- Utilizado para desenvolvimento complementar.
- O usuário informou que não realizou novas alterações no projeto na noite de 11/11/2026.
- O estado exato do Git local deve ser conferido antes de qualquer sincronização.
- Não presumir que arquivos locais ou evidências não versionadas foram transferidos pelo GitHub.

## 5. Problemas conhecidos

### Recortar página — atributo /Tabs

A experimental.39 apresenta o bloqueio:

`Recorte bloqueado: atributo Tabs ainda não suportado.`

Relatos de ocorrência:

- PDFs comuns.
- PDFs originados no Word.
- PDFs exportados do PowerPoint.

PDFs criados a partir de imagens funcionaram nos testes relatados.

A causa provável está nas verificações conservadoras de atributos de página em `pdf-workspace.js`, identificadas na análise da versão publicada .38.

A implementação local da .39 deverá ser inspecionada antes de modificar o código.

**Tarefa:** PDF-001 no planejamento.

### Espaço temporário do instalador

Houve uma tentativa de instalação com esgotamento do espaço disponível em C:.

A correção da .39 resolveu problemas de compatibilidade e resíduos, mas não foi demonstrada uma redução do pico de consumo de espaço.

**Tarefa:** INST-001.

### Homologações pendentes

- Step & Repeat por grade em cenários reais de produção.
- Frente e verso com impressora e RIP.
- Documentos Word complexos.
- Tratamento completo de caixas de corte e sangria.
- Compatibilidade com PDFs complexos.
- Validação mais ampla antes da distribuição.

Consultar o planejamento para os respectivos IDs.

## 6. Próxima tarefa autorizada

**SYS-001 — Organização e sincronização do desenvolvimento.**

Sequência:

1. Finalizar `PLANEJAMENTO.md` e `CONTEXTO_ATUAL.md`.
2. Atualizar `AGENTS.md` com as regras de sincronização.
3. Revisar as alterações documentais.
4. Criar um commit local separado do checkpoint da experimental.39, após autorização.
5. Conferir os commits e executar push autorizado para o GitHub.
6. Clonar o repositório atualizado no computador dedicado.
7. Confirmar que o computador dedicado possui a mesma versão e os mesmos documentos.
8. Iniciar a próxima tarefa aprovada no planejamento.

A correção do atributo `/Tabs` será tratada em uma etapa separada.

## 7. Regras para retomar o trabalho

Antes de implementar:

- Conferir o Git local e remoto.
- Identificar alterações não publicadas.
- Ler o planejamento e confirmar a prioridade aprovada.
- Verificar se a tarefa já foi iniciada em outro ambiente.
- Preservar alterações anteriores, originais, perfis e backups.
- Não fazer reset destrutivo, instalação, impressão ou push sem autorização.

Ao concluir uma tarefa:

- Informar arquivos modificados.
- Registrar testes e limitações.
- Atualizar o planejamento, quando houver mudança de status autorizada.
- Atualizar este documento com a nova situação.
- Registrar o histórico técnico em `ESTADO_PROJETO.md` quando apropriado.
- Informar explicitamente se o trabalho foi enviado ao GitHub.

## 8. Pendência de sincronização

**Situação no momento da criação deste documento: não sincronizado.**

O checkpoint local da experimental.39 e os arquivos novos de planejamento ainda precisam ser publicados.

Não considerar o computador dedicado ou o computador de casa atualizados até que o Git seja conferido nesses ambientes.

---

**Próximo passo:** finalizar a documentação de continuidade, revisar e preparar o segundo commit.
