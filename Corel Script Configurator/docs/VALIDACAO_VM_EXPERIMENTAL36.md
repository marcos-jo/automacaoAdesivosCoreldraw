# Validação em VM da atualização .34 → .36

## Avaliação do operador e continuidade na gráfica — 10/11/2026

Após os ensaios automatizados, o operador abriu a .36 dentro da VM e relatou avaliação visual positiva: **"ficou ótimo"**. Esse é um resultado manual informado pelo operador, separado das verificações automatizadas abaixo. Não houve descrição individual de cada ação do roteiro; não afirmar testes específicos de edição, exportação ou imposição como integralmente aprovados só por esse relato. A falha de atualização foi corrigida no cenário reproduzido no laboratório; homologação no computador real da gráfica ainda pendente.

Autorizado commit local para continuidade, partindo de main/4efcc569ab283cb96638ffb82177f7ce9bf81cf9. Inventário desta etapa: somente `docs/ESTADO_PROJETO.md` e `docs/VALIDACAO_VM_EXPERIMENTAL36.md`. Código da correção .36 já consta do commit-base; nenhum código/motor novo nesta consolidação. Push será feito manualmente pelo operador. Nenhuma credencial, runtime, instalador, VM/ISO, perfil, sintético ou evidência binária incluído.

O Setup existente fica em `dist-desktop/hold-native-test/4PrintShops-Hold-Test-1.0.0-experimental.36-Windows-x64-Setup.exe`, com514.970.609 bytes e SHA256 `1A39DAD02CD34C79E26D7439F5DA076D7118B68469F7775CF2C396F231762B02`. **O Git não transporta esse instalador nem os recibos/VM ignorados.** Para levar o mesmo pacote, copiar separadamente por meio escolhido pelo operador e conferir o SHA256 no destino; uma reconstrução é um novo artefato e deve ser auditada novamente, sem presumir o mesmo hash.

Na gráfica, começar lendo AGENTS.md, ESTADO_PROJETO.md, este relatório e ATUALIZACAO_TRANSACIONAL_EXPERIMENTAL36.md. Conferir branch/HEAD/status e versão/diretório/perfil realmente instalados, preservando alterações locais. Planejar atualização controlada somente após fechar normalmente o aplicativo/documentos e fazer backup verificado de perfil/instalação/registro pertinentes. O caminho172 e a conta da VM são fixtures de laboratório, não destinos para aplicar automaticamente na máquina real. Continuam pendentes as políticas próprias da gráfica, preferências reais, fontes/paginação/RIP e homologação física. Não imprimir, limpar arquivos ou iniciar nova atualização automaticamente por causa deste checkpoint. **.36 ainda não homologada para produção.**

Interrupção abrupta/queda de energia e recuperação manual continuam como ensaios separados não realizados nesta VM. O rollback aprovado cobre falha tratada por bloqueio real; não declarar recuperação automática de uma transação interrompida à força.

## Resultado atual: migração e Word automatizados aprovados — 10/11/2026

Operador confirmou a elevação para executar **somente a continuação**, sem repetir o Setup. Recibos completos em `.desktop-test/vm36-environment-20261110-055213/real-post-migration36-20261110-123517/`, especialmente `result.json`, `word-pdf-inspection.json`, `owned-registry-differences.json` e `guest-evidence/`. Processo terminou com saída0 e success=true. Os registros de espera/cancelamento abaixo são históricos, anteriores a esta conclusão.

| Verificação automatizada | Resultado comprovado |
| --- | --- |
| Setup .34 doméstico original, destino172 e LongPathsEnabled0 | Reproduziu instalação incompleta:18.402 presentes íntegros e1.108 ausentes; saída0 inadequadamente favorável |
| Setup .36 com bloqueio real no app.asar da .34 | Saída2/EBUSY esperadas, diário rolled-back;18.404 arquivos anteriores, perfil e registro/default restaurados |
| Atualização sem bloqueio pelo Setup .36 original | Saída0; diário committed, origem1.0.0-experimental.34 |
| Versão/diretório/identidade após atualização | .36 registrada no mesmo destino172; appId/GUID Hold Test preservados |
| Payload instalado, conferido novamente na continuação |19.510 arquivos iguais por tamanho/SHA256; zero ausentes/diferentes |
| Verificação nativa antes e depois das conversões | Ambas saída0; Office privado26.8.2.2 e19.410 arquivos íntegros |
| DOCX e DOC sintéticos | Ambos convertidos pelo soffice.com privado instalado, saída0, originais preservados por SHA256 |
| PDFs gerados, inspecionados no host após cópia por PowerShell Direct | Dois PDFs, duas páginas A4 cada, texto/acentos preservados; hashes iguais aos recibos do convidado |
| Perfil sintético | Inventário integral antes da instalação e depois dos testes igual |
| Leitor PDF padrão | UserChoice e valor padrão de .pdf iguais antes/depois |
| Registro próprio | Somente DisplayName, DisplayVersion e EstimatedSize mudaram; demais valores/árvores auditados iguais |

O diário novo é `.4ps-upgrade-8QyCF3/journal.json`: committed,18.404 arquivos gerenciados anteriores,19.512 arquivos novos (payload mais auxiliares do instalador) e zero arquivos não gerenciados nessa fixture. O diário anterior rolled-back e seu diagnóstico EBUSY foram preservados. Um erro antigo registrado não deve ser confundido com falha da migração posteriormente confirmada.

**Escopo Word:** chamado explicitamente o `lo/program/soffice.com` da instalação da VM, com ambiente de processo isolado igual ao privado e perfil seguro separado para cada conversão. Nenhuma descoberta/fallback para Office externo utilizada nesse teste; nenhuma instalação global de Office feita no convidado. Verificação nativa posterior confirmou ausência de arquivos extras no runtime. Estes são testes CLI/integração do runtime real, **não testes da janela de importação**.

PDFs sintéticos preservados no host, relativos à pasta de evidências acima:

| Saída | Bytes | Páginas/dimensões | SHA256 |
| --- | --- | --- | --- |
| `guest-evidence/output-docx/synthetic-word.pdf` |48.495 |2×210,01×297mm |`2516618f98cba99ad84c7d7d0e6dcf696fbc1a0a0748bfe8172572e6bd847b9b` |
| `guest-evidence/output-doc/synthetic-word.pdf` |48.495 |2×210,01×297mm |`d23d9bfe10fc76158ab9103d9c932f1d7a138b4ce9a97e4b189a9ff744becb47` |

PDF.js encontrou as cinco frases sintéticas esperadas, inclusive acentos, e três fontes por saída. Emissão de avisos: uso do build não legacy no Node e `TT: undefined function:21`. Não impediram extração/conferências; não equivalem a aprovação visual de fontes nem foram corrigidos fora do escopo. Originais DOCX/DOC mantiveram os hashes do kit.

**Limites e pendências:** marcador/perfil sintético preservado não prova preferências reais de operador. Ainda falta uso visual da .36 instalada na VM (abrir PDF, editar/organizar/exportar, importar Word pela interface e conferir prévias), comparação visual de fontes/paginação e homologação de drivers/RIP na gráfica. A origem é o Setup .34 doméstico, não o rebuild de hash distinto da gráfica; quantitativos/cenário da falha foram reproduzidos, mas a causa exata da recusa histórica do Office da gráfica permanece indeterminada. Windows da VM tem LongPathsEnabled0, porém outras políticas da gráfica não foram reproduzidas integralmente. Os testes gerais489 Node pertencem à validação anterior, não foram reexecutados nesta etapa documental/VM.

Próximo roteiro manual na VM, somente com sintéticos: abrir a .36 e conferir a versão; importar PDF, girar/reordenar/exportar uma cópia; importar DOCX e DOC pela interface e conferir fontes/duas páginas; abrir prévias Step & Repeat/N-up/livreto e cancelar; fechar normalmente. **Não imprimir nem atualizar a instalação ativa do host nesta etapa.**

Nenhuma alteração de produção ou novo pacote. Dois documentos pendentes no Git, staging vazio; sintaxe dos auxiliares e git diff --check aprovados. Sem limpeza, impressão, commit ou push. **Não homologada para produção.**

## Ponto atual de retomada: .36 instalada na VM, verificações finais pendentes — 10/11/2026

**Não executar novamente o instalador.** O Setup .36 original já terminou com saída0 e os19.510 arquivos instalados conferiram por tamanho/SHA256, sem ausências/diferenças. A interrupção posterior foi do harness, na consulta ao endereço errado de InstallLocation, conforme seção abaixo. A sequência integral ainda não passou.

Preparado `post-migration36-controller.txt` exclusivamente no ambiente ignorado: consulta o valor na chave correta, exige versão .36, reaproveita os snapshots anteriores à instalação, confere novamente o payload e faz as verificações nativas/conversões DOC/DOCX. Não contém chamada de instalação; não substitui arquivos/registro para satisfazer o teste. `inspect-guest-word-pdfs.mjs` está preparado para inspecionar quantidade de páginas, dimensões A4 e texto dos PDFs sintéticos quando existirem. Sintaxe conferida, resultados não antecipados.

Pedido administrativo para essa continuação terminou com **cancelamento informado pelo Windows antes de iniciar o componente**. Nenhum ponteiro `latest-post-migration36-job.txt` nem resultado complementar criado. Recibo `post-migration36-launch-canceled-20261110-114413.json` na raiz do ambiente VM. Não repetir automaticamente, reinstalar, restaurar checkpoint ou contornar permissões. Operador deverá acompanhar e confirmar o próximo aviso normal de administrador.

| Verificação | Situação comprovada |
| --- | --- |
| Setup .34 original em raiz172/LongPathsEnabled0 | Falha original reproduzida:1.108 ausentes |
| Bloqueio real app.asar e rollback .36 | Aprovado:18.404 arquivos/perfil/registro/default restaurados |
| Setup .36 sobre essa .34 | Concluído:saída0 |
| Payload .36 instalado | Aprovado:19.510 arquivos, zero ausentes/diferentes |
| Diário/registro final e preservação final de perfil/default | Conferência complementar pendente |
| DOCX/DOC pelo Office privado no convidado | Pendente |
| Interface e uso visual da .36 instalada na VM | Pendente |
| Fontes/paginação/RIP e homologação física na gráfica | Pendentes |

Instalação ativa do host preservada; somente dois documentos pendentes, staging vazio, sem código de produção, impressão, limpeza, commit ou push. **Não homologada para produção.**

## Migração real: Setup e payload aprovados; harness interrompido após auditoria — 10/11/2026

Com elevação confirmada pelo operador, executado o Setup .36 original **somente na VM**, sem bloqueio artificial, no mesmo destino172 da instalação original .34 incompleta. Resultado em `real-migration36-20261110-112832/guest-evidence/installer-exit.json`: **saída0**, duração625.157,7896ms. `payload-audit.json`: **19.510 arquivos esperados/conferidos, zero ausentes e zero diferentes** por tamanho/SHA256.

A sequência automatizada terminou com `success=false` depois da auditoria, antes dos testes Word e da comparação final de perfil. Causa comprovada no harness: consulta a InstallLocation sob a chave Uninstall, onde o NSIS não grava esse valor. O snapshot anterior (`registry-before.json`) mostra InstallLocation em `Software\90ff2ff0-f013-571c-b109-7edb265dacc8`; Uninstall contém DisplayVersion e comandos de desinstalação. A chamada TrimEnd em valor nulo interrompeu o teste. **Não foi identificado defeito do aplicativo por essa exceção; não corrigir o Registro para satisfazer o teste.**

Preservados todos os recibos da interrupção. Preparada continuação `post-migration36-controller.txt`, corrigindo somente a consulta da automação e reaproveitando o baseline anterior à instalação. Essa continuação **não executa Setup novamente**; exige registro .36, audita o payload, confere diário e faz testes Word/verificação de perfil/default. Até obter seus resultados, não considerar a sequência integral aprovada. Sem código de produção, instalação ativa, impressão, limpeza, commit ou push alterados.

## Retomada após limite de uso: resultado recuperado, próxima etapa bloqueada — 10/11/2026

Conferidos o recibo final e as evidências do ensaio de bloqueio descrito abaixo: **success=true**, saída2 esperada,18.404 arquivos anteriores restaurados, perfil e registro/default iguais. O limite do chat não interrompeu aquele processo já iniciado. Nenhuma instalação foi repetida para recuperar esse resultado.

Preparados, exclusivamente no ambiente ignorado de laboratório, `migration36-controller.txt`, `migration36-body.txt` e `inspect-guest-word-pdfs.mjs`. Preveem atualização pelo Setup .36 original na mesma raiz172, auditoria dos19.510 arquivos, verificação nativa antes/depois de conversões DOC/DOCX sintéticas e comparação de perfil/leitor padrão. Conversão CLI não será apresentada como validação da interface. Sintaxe PowerShell/Node conferida; **nenhum desses novos ensaios foi aprovado por antecipação**.

Na tentativa de iniciar essa etapa, o Windows informou **“A operação foi cancelada pelo usuário”** antes de entregar um processo administrativo. Nenhum ponteiro `latest-migration36-job.txt` ou resultado da migração foi criado. Recibo `migration36-launch-canceled-20261110-105131.json` na raiz do ambiente. A saída do invólucro de lançamento não é evidência de sucesso; a mensagem de cancelamento e ausência de processo prevalecem. Sem nova tentativa automática ou contorno de permissões. Retomada exige operador acompanhando o aviso normal de administrador.

Situação: **Setup original .34 incompleto reproduzido e rollback real .36 aprovados; migração integral e Word no convidado ainda pendentes**. Instalação ativa do host não foi alvo de atualização. Nenhuma alteração de produção, impressão, limpeza, commit ou push. Git mantém somente ESTADO_PROJETO.md modificado e este relatório novo; staging vazio.

## Bloqueio real e recuperação da .34: aprovados — 10/11/2026

Estado do Setup .34 incompleto preservado no checkpoint `VM36-Setup34-original-incompleto-20261110-092721`, ID `09eb275a-4f25-4e52-a3f5-7bed9b6e67b1`. Uma sessão PowerShell Direct existente perdeu a conexão depois do checkpoint; o harness não entregou resultado (`real-lock36-20261110-092719/controller-error.json`). Nova sessão confirmou versão .34, nenhum Setup ativo e nenhuma pasta de evidência anterior no convidado antes de iniciar outro ensaio. Não foi necessário restaurar checkpoint, apagar arquivos ou repetir transação incompleta.

Ensaio concluído em `real-lock36-20261110-093025/result.json`, com evidências completas copiadas do convidado para `guest-evidence/`. Setup .36 original executado no mesmo destino172, offline e LongPathsEnabled0. Um FileStream de teste manteve bloqueio real sobre o app.asar anterior, sem modificar seu conteúdo. Instalador terminou com saída **2** após aproximadamente430 segundos; diário confirmou **EBUSY durante moving** e fase final **rolled-back**. O diagnóstico exato alcançou `.4ps-upgrade-last-error.ini`, comprovando a comunicação GUI/NSIS nesse caso.

Após liberar somente o bloqueio do próprio teste, conferidos por nome/tamanho/SHA256 os **18.404 arquivos anteriores**, incluindo18.402 do payload e dois auxiliares do instalador. Inventário completo do perfil sintético igual antes/depois. Árvores/valores próprios do registro e associação PDF padrão iguais antes/depois; versão registrada permanece .34. **Recuperação após bloqueio aprovada**. Não equivale à migração .36 concluída nem a teste visual dos módulos PDF. O processo terminou e seus recibos ficaram disponíveis mesmo após a interrupção do chat por limite de uso.

Próxima etapa: atualização sem bloqueio pelo Setup .36 original, auditoria dos19.510 arquivos e conversões DOC/DOCX sintéticas pelo Office privado no Windows convidado. Esses passos ainda não aprovados neste registro. Instalação ativa do host preservada, sem impressão, limpeza, commit ou push.

## Setup .34 original: instalação incompleta reproduzida — 10/11/2026

Após confirmação administrativa do operador, executado **somente na VM** o Setup doméstico .34 original, com hash previamente conferido, instalação currentuser e destino de172 caracteres. Laboratório offline/ativado, LongPathsEnabled0. Duração253.382ms; saída do instalador **0**.

A auditoria de todos os19.510 arquivos esperados identificou **18.402 presentes, todos iguais por tamanho/SHA256, e1.108 ausentes**. Corresponde aos quantitativos da falha registrada na gráfica. Portanto o cenário original foi reproduzido; não declarar a .34 íntegra nem homologada por sua saída0. Versão .34 e diretório conferidos no registro próprio. Leitor PDF padrão e marcador sintético de perfil preservados; marcador não equivale a preferências reais nem prova de interface utilizável.

Evidências no host: `.desktop-test/vm36-environment-20261110-055213/real-setup34-20261110-090600/result.json`. Evidências no convidado: `C:\4PS-TestKit\real-setup34-20261110-090621`. Resultado detalhado inclui os1.108 nomes ausentes; não contém documentos de clientes. Inventários .34/.36 copiados para o kit convidado. Sem modificar políticas ou preencher manualmente os arquivos ausentes.

Próxima etapa: preservar este estado por checkpoint e avaliar o Setup .36 real, começando por falha de bloqueio e recuperação, depois migração completa. **Nenhum Setup .36 foi executado até este registro.** Instalação ativa do host intocada; sem impressão, limpeza, commit ou push. Os registros históricos abaixo de cancelamento precedem e não invalidam esta execução concluída.

## Laboratório pronto, offline e com checkpoint atualizado — 10/11/2026

Com operador acompanhando RustDesk, conferência final concluída com saída0. Recibos em `offline-ready-20261110-084406/`: adaptador da VM **desconectado**, nenhuma rede física alterada; licença permanece **status1**, período129.590 minutos, Windows26300.9457 e LongPathsEnabled0. O recibo anterior de verificação de rede com erro de parâmetro foi superado por esta leitura válida, sem apagar o histórico.

Criado checkpoint **VM36-Windows-pronto-testes-20261110**, ID `83f4b8bb-2603-46ae-8b2e-fa72834a532c`; tipo de snapshot retornado Standard, configuração da VM Production. Recibo `windows-ready-checkpoint.json`. Preserva o laboratório ativado, a credencial definida pelo operador e o kit; usar este baseline nos ensaios novos, sem apagar o checkpoint anterior à credencial.

Inventários dos payloads **originais extraídos e anteriormente auditados** preparados por leitura: .34 e .36 com19.510 arquivos cada, manifesto privado19.410 arquivos conferido em ambas as fontes. Arquivos `guest-kit/expected-34.json` e `expected-36.json`. Não foi usado o win-unpacked com diagnóstico extra como fonte íntegra. A etapa seguinte executa somente na VM o Setup .34 doméstico original no caminho172, compara os arquivos, leitor PDF e marcador de perfil sintético; processo ainda depende da confirmação administrativa normal. Não antecipar resultado de instalação ou migração.

Pedido de elevação dessa etapa terminou com cancelamento informado pelo Windows antes de iniciar o processo elevado; recibo `real-setup34-20261110-085238/launch-error.json`. Portanto o Setup .34 **não foi executado**, e os inventários novos ainda não foram copiados ao convidado. Etapa preparada, aguardando o operador acompanhar o próximo aviso; não repetir automaticamente nem declarar teste reprovado de aplicativo por este cancelamento.

## Avaliação Windows ativa; conferência final de isolamento pendente — 10/11/2026

Operador autorizou internet temporária somente na VM para tentativa oficial de ativação. Conectado seu único adaptador ao **Default Switch existente**, sem criar rede/alterar adaptadores físicos. Consulta anterior: `LicenseStatus=5`, período0. Ao receber rede, Windows passou a **LicenseStatus=1** e **129.600 minutos (90 dias)**; tentativa explícita `slmgr.vbs /ato` retornou `0xC004E028` (outra ativação em andamento). O resultado favorável é a consulta de licença, não uma aprovação do retorno do comando concorrente. Não repetir ativação enquanto o estado for1. Sem troca de chave, rearm, alteração do relógio, políticas ou licença do host.

Recibos em `.desktop-test/vm36-environment-20261110-055213/official-activation-20261110-083247/`: `before.json`, `activation-result.json` e logs correspondentes dentro do convidado. A chamada de desconexão foi executada sem erro, mas a consulta subsequente usou parâmetro `VMId` não suportado em `Get-VMNetworkAdapter`, deixando o recibo conservador `networkDisconnected=false`. Não apresentar esse recibo como prova de desconexão nem como prova de que a desconexão falhou. Consulta separada corrigida usa `Get-VM` por ID e `Get-VMNetworkAdapter -VM`, desconecta qualquer vínculo remanescente e confere novamente a licença offline; na última conferência aguarda confirmação administrativa normal do operador.

Depois da confirmação serão gerados `network-offline-confirmed.json`, `license-offline-confirmed.json` e novo checkpoint `VM36-Windows-pronto-testes-20261110`, sem substituir o anterior. Ainda não considerar esse checkpoint criado. Nenhum Setup .34/.36 executado até este registro. Referência oficial para a concorrência: [Microsoft — erros de ativação](https://support.microsoft.com/en-gb/windows/activation/get-help-with-windows-activation-errors).

**Resultado do pedido final de elevação:** Windows devolveu “A operação foi cancelada pelo usuário” antes de executar essa conferência; recibo `offline-confirmation-launch-canceled.json`. A ativação permanece comprovada pelo estado1/período90dias da leitura anterior, mas rede offline e novo checkpoint ainda não foram confirmados/criados. Não repetir pedidos automaticamente; combinar a nova confirmação com o operador acompanhando a tela do RustDesk.

## Credencial validada; testes aguardam regularização da avaliação Windows — 10/11/2026

Operador definiu credencial somente na conta de laboratório e autorizou seu uso na automação; o valor não é registrado aqui. PowerShell Direct autenticou com sucesso. Uma primeira coleta retornou campos nulos por formatação incorreta de hashtable no harness, sem aprovar essa coleta; a repetição usou PSCustomObject e validação explícita de campos. Durante a sequência a VM desligou, impedindo a leitura seguinte. Relativa às tentativas: `guest-baseline-attempt-20261110-081607/` e `guest-baseline-checked-20261110-081822/`, preservadas no ambiente ignorado.

Religada somente a VM após confirmar estado desligado; DVD virtual desconectado e disco virtual definido como primeiro boot, preservando a ISO no host. Consulta concluída em `guest-restart-inspection-20261110-082226/result.json`: **cinco arquivos recebidos conferidos por tamanho/SHA256** (Setups .34/.36, PDF/DOCX/DOC); `LongPathsEnabled=0`, igual à política registrada na gráfica. Windows EnterpriseEval, DisplayVersion26H2, build26300.9457; o Registro mantém o rótulo legado ProductName “Windows 10 Enterprise Evaluation”. Não foram encontrados executáveis de Office nos caminhos globais usuais consultados, sem varredura exaustiva de outros caminhos. Nenhum Setup do 4PrintShops executado no convidado.

**Causa do desligamento confirmada no Windows convidado:** evento System1074, processo `wlms.exe`, comentário explícito de período de licença expirado. Licença `TIMEBASED_EVAL`, `LicenseStatus=5`, `GracePeriodRemaining=0`. Esse desligamento é do Windows, não falha de atualização .36. Não prosseguir com ensaios integrais sujeitos a interrupção periódica. A mídia continua sendo a ISO oficial já auditada; a causa do estado de licenciamento inicial ainda não foi corrigida nem atribuída a arquivo alterado.

Solicitada autorização para conectar **somente a VM** temporariamente ao Default Switch já existente, tentar ativação oficial Microsoft e desconectar ao terminar. Nenhuma conexão/ativação executada até esse registro. Não modificar relógio, políticas, licença do host ou reiniciar período de avaliação. Se a tentativa oficial falhar, preservar o diagnóstico e definir mídia/licença adequada antes de aprovar testes. O checkpoint Windows limpo existente foi criado antes da definição da credencial e não deve ser confundido com baseline autenticado/ativado; preservar e criar outro quando o laboratório estiver adequado.

## Windows limpo preservado e Setups transferidos — 10/11/2026

Área de trabalho da VM conferida visualmente; operador informou conta sintética sem senha. Criado checkpoint **VM36-Windows-limpo-20261110**, ID `196b57b2-a6bb-4efa-8c6d-fe4092c1178e`, tipo retornado **Standard**, em `2026-11-10T10:35:28.1848870Z`. Não apresentar como checkpoint Production; é ponto de restauração do laboratório, não backup externo. Recibo `windows-clean-checkpoint.json` no ambiente VM.

Habilitada somente a integração **Interface de Serviço de Convidado** da VM para copiar arquivos sem rede nem senha. Primeira tentativa imediatamente após habilitar retornou `0x800710DF` (dispositivo não pronto), antes de copiar arquivos. Estado consultado novamente e repetição única posterior aprovada. `guest-kit-transfer.json` comprova envio a `C:\4PS-TestKit` dos Setups originais domésticos .34/.36, `manifest.json` e `DESTINO-172-CARACTERES.txt`. Fontes reconferidas por SHA256, sem sobrescrever arquivos anteriores. Hashes dentro do convidado ainda por verificar. Nenhum Setup executado. Rede continua desconectada; senha/políticas não alteradas.

Preparados PDF sintético quatro páginas A6 e DOCX pelo helper de testes existente. Uma tentativa de gerar DOC com o Office privado do **build descompactado do host**, dentro do ambiente restrito de ferramenta, terminou sem produzir saída e deixou `dist-desktop/hold-native-test/win-unpacked/lo/program/0f5a35c9-a328-44ee-9d4c-4a040bf6af3c.dmp` (137.998 bytes; criado durante essa execução). Repetição fora da restrição foi recusada pela verificação de integridade por causa desse arquivo extra, antes de converter. **Não removido nem permitido no manifesto**, sem relaxar proteção ou alterar motores; causa específica do encerramento não isolada. O build descompactado não deve ser reutilizado como íntegro sem nova auditoria. Setups originais não foram alterados e seus hashes continuam aprovados.

Para não insistir no runtime com diagnóstico extra, reutilizado exclusivamente o DOC **sintético** do ensaio aprovado do Setup .36 anterior, identificado pelo recibo `hold36-setup-audit.log`. DOCX daquela fixture conferido byte a byte por SHA256 com o DOCX atual antes de copiar o DOC; fonte DOC preservada e cópia verificada por hash. Manifesto dos três sintéticos em `guest-kit/sinteticos/synthetic-manifest.json`. A preparação no host não comprova conversão Word no Windows convidado. Transferência dos sintéticos e sondagem read-only do convidado são etapas seguintes, com recibos próprios; sem mudança de senha ou políticas para forçar acesso.

**Resultado posterior:** `synthetic-transfer.json` confirma PDF/DOCX/DOC copiados para `C:\4PS-TestKit\sinteticos`. Tentativa única de consulta PowerShell Direct com a conta sintética informada e senha vazia retornou **“A credencial é inválida”** (`PSDirectException`); recibo `guest-readonly-probe-error.json`. Operador confirmou o nome composto da conta. Não tentar variantes de credenciais nem enfraquecer políticas para forçar entrada. A causa exata da recusa não foi isolada; hashes dos arquivos já recebidos e LongPathsEnabled no convidado ainda não foram conferidos. Solicitada escolha entre definir senha de laboratório pelo próprio operador (sem transmissão no chat) para automação, ou preservar conta sem senha e seguir manualmente. Nenhum Setup iniciado; checkpoint e arquivos prontos para a próxima etapa.

## Instalação Windows e configuração inicial pelo operador — 10/11/2026

Após a primeira tela UEFI sem boot, operador refez a inicialização e confirmou abertura do Windows Setup. Não foi necessário alterar Secure Boot/TPM/ISO. Conferida visualmente a seleção de **um único Disco0, espaço não alocado80GiB**, correspondente ao VHDX vazio do laboratório; instalação iniciada pelo operador. Posteriormente relatou configuração sem internet, usuário sintético e recusa dos serviços opcionais Microsoft que lhe foram apresentados. Essas escolhas foram feitas pelo operador, sem automação de segurança/privacidade ou credenciais pelo agente.

Na última observação há console da VM e diálogo **Conectar a 4PrintShops-Upgrade36-VM** com resolução1366×768. Impressoras já desmarcadas; solicitado ao operador desmarcar Área de Transferência e conectar. A área de trabalho ainda precisa ser confirmada antes do checkpoint do Windows limpo e da transferência controlada de Setups/sintéticos. Nenhuma senha coletada; nenhum teste .34→.36 aprovado por este relato. Não há alteração de fonte ou instalação ativa do 4PrintShops.

## UAC remoto aprovado e VM criada — 10/11/2026

Operador confirmou que a aprovação de administrador pelo RustDesk funcionou. Sondagem somente leitura retornou saída0, `administrator=True`, `success=True` e zero VMs antes da criação; recibo `.desktop-test/rustdesk-uac-probe-20261110-064606/result.json`.

Criação posterior concluída com saída0. VM **4PrintShops-Upgrade36-VM**, ID `70cfd57c-daa6-43c8-9340-e82afc447097`, geração2, memória estática6GiB, duas CPUs, VHDX dinâmico até80GiB (4MiB inicialmente), Secure Boot On, TPM virtual habilitado, rede desconectada e início automático Nothing. ISO validada conectada e DVD definido como primeiro dispositivo de boot. Configuração e disco estão exclusivamente em `.desktop-test/vm36-environment-20261110-055213/vm/`; recibo `vm-created.json` na raiz desse ambiente.

Console elevado aberto e conferido visualmente, com título identificando a VM e estado **Desligada**. Clique automatizado em Iniciar não alterou o estado; solicitado ao operador iniciar pelo RustDesk e pressionar uma tecla se o boot DVD pedir. Não repetir cliques sem observar mudança. Nenhum Windows instalado no convidado nem teste dos Setups .34/.36 executado até esse ponto. Instalação ativa preservada. Os registros abaixo de criação bloqueada são históricos, superados por esta aprovação.

## RustDesk instalado com operador presente — 10/11/2026

Operador retornou ao computador e autorizou a continuidade da instalação já solicitada. MSI oficial reconferido por SHA256/assinatura e executado com elevação normal, `/passive /norestart`; saída **0**, sem reinício necessário. Executável em `C:\Program Files\RustDesk\rustdesk.exe`, versão **1.5.1+68**, assinatura **Valid/PURSLANE**. Serviço `RustDesk` confirmado **Running**, inicialização automática do instalador. Janela RustDesk aberta para conferência pelo próprio operador.

Recibo `install-result.json` e log `install-20261110-063840.log` em `.desktop-test/rustdesk-preparation-20261110-061827/`, ignorados. Não foi configurada senha permanente nem registrada credencial/ID de acesso remoto em documentação ou Git. A conexão pelo outro dispositivo ainda depende de teste do operador. A instalação do RustDesk não aprova a atualização .36: a VM continua por criar e nenhum Setup do 4PrintShops foi executado no host. Sem commit/push. A seção anterior de instalador somente preparado, abaixo, permanece como histórico.

## Preparação de acesso remoto solicitada — 10/11/2026

Operador, sem acesso físico no momento, solicitou instalar RustDesk para aprovar a elevação remotamente. Não foi identificado cliente/serviço RustDesk instalado nos locais consultados. A documentação oficial informa que a versão portátil sem elevação não controla UAC; a instalação/elevação inicial também exige aprovação administrativa no computador. Essa dependência não foi contornada.

Deixado o instalador oficial RustDesk 1.5.1 x64 em `.desktop-test/rustdesk-preparation-20261110-061827/rustdesk-1.5.1-x86_64.msi`: 26.718.208 bytes, SHA256 `71069183CBC6124273138684AB4032D2211E83EFA035886F9C774FE1D3EBA3F0`, igual ao digest do asset oficial GitHub, assinatura Authenticode **Valid**, assinante `PURSLANE`. Recibo `installer-verified.json` na mesma pasta ignorada. **Não instalado nem executado**; nenhum serviço, senha, regra de firewall ou acesso remoto permanente configurado. Próximo passo depende de alguém aprovar a primeira instalação no computador ou de acesso administrativo remoto já provisionado pelo operador. A VM e os testes continuam pendentes. Referência: [RustDesk — elevação portátil](https://rustdesk.com/docs/en/client/windows/windows-portable-elevation/).

## Mídia pronta; criação da VM não iniciada — 10/11/2026

Download autorizado concluído com HTTP 200 e saída 0, em aproximadamente 721 segundos. Arquivo final preservado em:

```text
.desktop-test/vm36-environment-20261110-055213/media/Windows11EnterpriseEval-pt-BR-x64.iso
Bytes: 8170147840
SHA256: 29BD7592E832D32073D1637CE83689DE5F375700A963B2D753E437CC64D76C09
```

Esse SHA256 é o digest calculado localmente, **não uma comparação com hash publicado pela Microsoft**. Origem oficial HTTPS registrada. Estrutura UDF testada integralmente pelo 7-Zip 26.01 já instalado: 1.066 arquivos/102 pastas, `Everything is Ok`. Extraído somente `setup.exe` para inspeção, sem executá-lo; assinatura Authenticode **Valid**, emissor do arquivo `CN=Microsoft Corporation, O=Microsoft Corporation, L=Redmond, S=Washington, C=US`. Recibo final `iso-validated.json`, logs `iso-udf-test.log` e `iso-udf-extract.log` no ambiente acima. Não foram instaladas ferramentas novas.

A primeira validação foi interrompida por caminho de ferramenta inexistente. A leitura automática ISO9660 com a ferramenta antiga só expôs o volume simplificado, sem `setup.exe`; não foi considerada prova de integridade completa. A inspeção final usou explicitamente UDF e ferramenta existente compatível; todos os arquivos e diagnósticos anteriores foram preservados. Arquivo parcial foi renomeado para `.iso` somente após as verificações finais.

**Impedimento atual:** consulta administrativa de inventário passou, mas a elevação para criar a VM foi cancelada pelo Windows antes de iniciar o processo elevado. Nenhuma VM, VHDX ou rede criada. O plano `vm-plan.json` está pronto; será necessário acompanhar e confirmar o próximo pedido normal de administrador. Não repetir automaticamente, alterar grupos ou contornar o UAC. Não há instalação Windows no convidado nem teste real de upgrade aprovado nesta etapa. Instalação ativa .34/registro .26 permanece intocada.

Somente os mesmos dois documentos pendentes no Git; `git diff --check` conferido, sem build, limpeza, impressão, commit ou push. Os registros de download em andamento abaixo são históricos e foram superados por esta conclusão.

## Preparação autorizada do ambiente — 10/11/2026

Operador autorizou usar a avaliação oficial Microsoft. Ambiente isolado em `.desktop-test/vm36-environment-20261110-055213/`, ignorado pelo Git. Consulta administrativa concluída com sucesso: hipervisor ativo, **zero VMs existentes** e nome `4PrintShops-Upgrade36-VM` disponível. A rede `Default Switch` já existia; nenhuma rede nova criada.

Download iniciado pelo link oficial PT-BR x64 abaixo, exclusivamente HTTPS, para `media/Windows11EnterpriseEval-pt-BR-x64.iso.part`. Tamanho esperado 8.170.147.840 bytes; o arquivo parcial não é mídia aprovada. `environment.json` registra origem e `iso-download.log` registra a transferência. Download e integridade final ainda em validação no momento deste registro.

Plano concreto em `vm-plan.json`: geração 2, 6 GiB de memória estática, duas CPUs virtuais, disco VHDX expansível até 80 GiB, Secure Boot Microsoft Windows e TPM virtual. Rede desconectada, início automático desabilitado e sem redirecionar impressoras/pastas do host. Criação solicitada por elevação administrativa normal; só considerar concluída quando `vm-created.json` confirmar todos os parâmetros. Não iniciar VM com ISO parcial nem reutilizar/alterar VM preexistente. Nenhum Setup do 4PrintShops foi executado.

A consulta administrativa foi aprovada, mas o pedido seguinte para criar a VM retornou **“A operação foi cancelada pelo usuário”** antes de iniciar o componente elevado. Nenhuma VM criada; registro `vm-creation-not-started.json`. Essa parte foi interrompida, sem repetição automática ou alteração de grupos/permissões. Será necessário confirmar o pedido normal de administrador com o operador presente, ou criar a VM manualmente a partir do plano auditado. O download autorizado e sua validação são independentes dessa etapa.

## Retomada após reinício — 10/11/2026

Operador informou reinício concluído. Consulta somente leitura confirmou `HypervisorPresent=True`, `vmms=Running` e `vmcompute=Running`; cmdlets `Get-VM`, `Get-VMHost` e `New-VM` disponíveis. Habilitação efetiva do hipervisor confirmada. O volume C: tinha aproximadamente 153 GiB livres. Branch/HEAD preservados (`main/4efcc56`), com somente os mesmos dois documentos pendentes.

A consulta direta de VMs/configurações foi recusada pelo token sem elevação. Iniciada consulta **somente leitura** via pedido normal de administrador; na última conferência o processo `consent` estava presente e o relatório administrativo ainda não tinha sido entregue. Não alterar grupos/permissões para contornar esse requisito nem presumir que não existem VMs sem o resultado. Evidências desta retomada em `.desktop-test/vm36-postboot-20261110-053634/`; Setups .34 e .36 reconferidos com os hashes da tabela abaixo.

Não há ISO Windows nos arquivos diretamente consultados em Downloads (somente ISO Linux). Localizada opção pública oficial de avaliação: [Windows 11 Enterprise](https://www.microsoft.com/en-us/evalcenter/download-windows-11-enterprise), português Brasil x64. O link oficial `https://go.microsoft.com/fwlink/?LinkId=2382605&clcid=0x409&country=us&culture=en-us` redireciona para o domínio Microsoft `software-static.download.prss.microsoft.com`; consulta HTTP HEAD retornou 200 e tamanho **8.170.147.840 bytes**. É avaliação de 90 dias; cadastro/licença/conta que forem necessários ficam com o operador, sem dados inventados. Foi solicitada a escolha do operador entre avaliação oficial e mídia própria. Nenhum download foi iniciado.

**Situação:** Hyper-V ativo, aguardando confirmação administrativa e escolha da mídia antes de criar/iniciar VM. Nenhum teste de atualização em VM executado. Nenhuma VM, rede ou disco virtual criado nesta retomada; instalação/perfis reais preservados, sem build, limpeza, impressão, commit ou push. Esta seção supera o estado anterior de reinício pendente.

## Hyper-V preparado com autorização — 09/11/2026

Após a pré-verificação abaixo, o operador autorizou preparar Hyper-V neste computador, deixando o reinício sob seu controle. A tentativa de executar o arquivo local `enable-hyperv.ps1` foi bloqueada pela política de execução do Windows PowerShell elevado, antes de modificar recursos; o diagnóstico foi capturado em `launch-error.log`. Nenhuma política de scripts foi alterada e o arquivo não foi executado com bypass.

Usada a alternativa nativa documentada pela Microsoft: DISM elevado com `/Online /Enable-Feature /FeatureName:Microsoft-Hyper-V-All /All /NoRestart`. Resultado **3010: sucesso com reinicialização necessária**, registrado em `2026-11-10T00:57:54.8396626Z` (09/11 à noite no horário local). Nenhum reinício foi solicitado pelo agente. Evidências locais ignoradas: `.desktop-test/hyperv-prepare-20261109-215427/`, incluindo `dism-result.json`, `dism-enable.log` e o diagnóstico da tentativa PowerShell.

**Próximo passo:** operador salvar os trabalhos e reiniciar quando conveniente; depois confirmar disponibilidade do hipervisor, serviço/gerenciador Hyper-V e recursos antes de criar a VM. Não há VM criada nem teste .34→.36 iniciado. Ainda será necessário obter mídia Windows autorizada para o convidado; a inspeção não recursiva de Downloads encontrou somente ISO Linux. Não foi baixada imagem Windows, criada rede virtual ou executado Setup do 4PrintShops. A autorização concedida nesta etapa não inclui commit/push ou instalação do aplicativo ativo.

O impedimento inicial abaixo é histórico: a preparação do Hyper-V foi agora autorizada e concluída com reinício pendente; os demais limites de teste continuam válidos.

## Pré-verificação em casa — 09/11/2026

Base: `main/4efcc569ab283cb96638ffb82177f7ce9bf81cf9`, árvore inicialmente limpa e referência local `origin/main` coincidente. Nenhum fetch, instalação ou teste de atualização realizado nesta etapa.

Consultas somente leitura confirmaram Windows Professional, build 26200; aproximadamente 16 GiB de RAM; virtualização no firmware e SLAT disponíveis; `HypervisorPresent=False`. Não foram encontrados cmdlets/serviços usuais Hyper-V, executáveis padrão de VirtualBox/VMware ou Windows Sandbox. Isso não exclui ferramentas ou VMs em outros caminhos. A consulta `Get-WindowsOptionalFeature -Online` exigiu elevação administrativa e não foi concluída; não classificar os recursos como desabilitados com base somente nessa consulta.

O volume C: tinha aproximadamente 152 GiB livres. O host mantém `LongPathsEnabled=1`; não alterar essa política para executar o ensaio. A política 0 registrada na gráfica deverá ser aplicada **somente ao Windows convidado descartável**.

## Instaladores disponíveis e reconferidos

Ambos em `dist-desktop/hold-native-test/`, preservados sem executar:

| Versão | Arquivo | Bytes | SHA256 |
| --- | --- | ---: | --- |
| .34 | `4PrintShops-Hold-Test-1.0.0-experimental.34-Windows-x64-Setup.exe` | 514965183 | `6E9C7AD7CC1DD41939D044FB03BD891C1C8E68E6230BCEB48865F8EC8F9FD3C4` |
| .36 | `4PrintShops-Hold-Test-1.0.0-experimental.36-Windows-x64-Setup.exe` | 514970609 | `1A39DAD02CD34C79E26D7439F5DA076D7118B68469F7775CF2C396F231762B02` |

O Setup .34 disponível é o original doméstico, não o artefato recompilado na gráfica (`B029B493…129E`). Ele permite avaliar o instalador/desinstalador original doméstico; não comprova reprodução byte a byte da instalação da gráfica. O inventário real de 1.108 ausências também continua indisponível. Não fabricar essa lista.

## Roteiro preparado — ainda não executado

1. Disponibilizar uma VM Windows descartável, com instalação Windows autorizada, distinta do host e sem dados/contas de atendimento. Registrar versão, política de caminhos e identificação do convidado. Preparar checkpoint antes de instalar os Setups; nenhuma pasta do host deve ser compartilhada com escrita. Não usar a instalação ativa como laboratório.
2. Transferir somente os dois instaladores e documentos sintéticos; reconferir os hashes dentro da VM. Não transferir perfil real, repositório completo, credenciais ou documentos de clientes. Não disponibilizar impressoras físicas à VM.
3. No convidado, instalar a .34 pelo Setup original em caminho sintético longo (raiz de aplicativo com 172 caracteres, conforme laboratório anterior), mantendo a identidade Hold Test. Registrar modo por usuário/todos os usuários e o destino efetivamente utilizado. Inventariar arquivos, registro, associações, leitor PDF padrão/UserChoice e preferências sintéticas. Registrar faltas reais sem presumir que serão exatamente 1.108. Fechar o aplicativo e preservar checkpoint desse estado.
4. Executar o Setup .36 no mesmo destino e modo de instalação. Conferir versão, todos os hashes, perfil/preferências, atalhos, associações próprias, leitor padrão e ausência de instalação duplicada. Abrir o aplicativo e converter DOC/DOCX pelo motor privado sem Office externo; preservar os originais sintéticos.
5. Em estado restaurado do checkpoint .34, testar bloqueio real de arquivo e conferir recusa segura/restauração, registro e possibilidade de reabrir/editar PDF na versão anterior. Não aplicar variantes de falha do Upgrade Lab ao namespace Hold Test: falhas controladas de cópia/registro/integridade/Office já têm resultados de laboratório separados; sua reprodução na VM com identidade real exige um roteiro específico antes de executar.
6. Interrupção abrupta e recuperação manual são ensaios separados, depois dos casos anteriores e com checkpoint preservado. Não apresentar o mecanismo como recuperação automática após queda de energia. Exportar evidências sintéticas para pasta dedicada, sem apagar backups/diários.

## Situação e impedimento

**Nenhum teste em VM aprovado nesta etapa; ainda não há VM acessível identificada.** Próxima decisão: usar ambiente existente fornecido pelo operador ou preparar Hyper-V neste Windows Pro. A habilitação exige administrador e pode exigir reinício; não foi executada nem autorizada automaticamente pela pré-verificação. Não instalar hipervisor, baixar imagem Windows ou reiniciar sem combinar com o operador.

Instalação ativa, perfil, política do host, drivers e impressoras não foram modificados. Não houve build, limpeza, impressão, commit ou push. A divergência ativa arquivos .34/registro .26 continua intocada.

Referências: [relatório da retomada](RETOMADA_EXPERIMENTAL36_CASA.md), [instalação Hyper-V — Microsoft](https://learn.microsoft.com/en-us/windows-server/virtualization/hyper-v/get-started/install-hyper-v), [checkpoints — Microsoft](https://learn.microsoft.com/en-us/windows-server/virtualization/hyper-v/checkpoints).
