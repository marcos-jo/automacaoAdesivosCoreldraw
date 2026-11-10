# Retomada da experimental.36 em casa

09/11/2026. **Validação local concluída; Setup Hold Test .36 gerado e auditado. Não homologado para produção e não instalado no aplicativo ativo.** Sem limpeza, impressão, commit ou push.

## Ponto de partida e evidências

Retomada na `main`, HEAD `522e9f62ee64056752832cf3fbd6bdd0c131b91d` (`WIP experimental 36 - atualizacao transacional em desenvolvimento`), árvore e staging limpos, remoto original preservado. Os caminhos descritos como pendentes nos guias da gráfica já estão nesse checkpoint; não eram alterações locais a descartar ou consolidar novamente.

Lidos AGENTS.md e os cinco documentos solicitados, incluindo o histórico integral de ESTADO_PROJETO.md. Conferidos código, templates NSIS fixados, testes e configurações reais. Nenhum processo de instalação/conversão anterior da .36 foi localizado neste computador.

O ponto documentado da interrupção anterior era a validação integral do NSIS após corrigir a leitura física de `app.asar` com `original-fs` e reconstruir o Upgrade Lab: instalação limpa iniciada, sem resultado final confirmado. Não havia motivo para recomeçar a arquitetura. Sem os logs ignorados da gráfica, esse término anterior continua indeterminado; os resultados abaixo são novos ensaios locais completos, com seus próprios recibos.

Os logs/diários/builds ignorados da .35/.36 da gráfica **não vieram com o Git**: não existem aqui `upgrade36-node-final.log`, o pacote Upgrade Lab anterior, os laboratórios integrais interrompidos ou o inventário dos 1.108 arquivos ausentes da gráfica. Portanto não foi possível recuperar o resultado final do ensaio anterior nem atribuir sua recusa Office ao defeito físico de app.asar.

Disponíveis em casa: Setup .34 de 514.965.183 bytes, SHA256 `6E9C7AD7CC1DD41939D044FB03BD891C1C8E68E6230BCEB48865F8EC8F9FD3C4`, payload público e evidências domésticas anteriores em `.desktop-test/release34-1794168709616/`; runtime privado e quatro arquivos-fonte correspondentes preparados. A cópia Office do payload .34 passou na verificação de todos os 19.410 arquivos. Não foram usados documentos de clientes.

Evidências desta retomada: pasta indicada em `.desktop-test/resume36-home-current.txt` (inicialmente `.desktop-test/resume36-home-20261109-193338/`). Há inventário por SHA256 da instalação ativa e de seus perfis, snapshot do registro próprio e UserChoice, histórico Git inicial, logs de testes/builds e referências às pastas temporárias novas. Evidências/instaladores/backups existentes foram preservados.

A consulta inicial encontrou uma divergência pré-existente em casa: executável e ASAR ativos informam .34, enquanto o registro NSIS consultado informa .26. Não houve tentativa de corrigir ou normalizar esse estado; a comparação final usou o inventário inicial real. Os estados da estação da gráfica não são inferidos a partir desta máquina. O preparador recusa versões anteriores fora de .34/.35/.36 e divergências de metadados; não foi contornado para permitir uma atualização real desse estado doméstico.

## Defeitos e ajustes confirmados nesta retomada

1. **Perfil do Upgrade Lab:** o nome `4printshops-upgrade-lab` não era reconhecido por `startDesktop`; sua abertura normal escolheria `4PrintShops` e o perfil principal. Um teste do módulo real do ASAR, com app injetado que interrompe antes de acessar qualquer perfil, reproduziu a escolha incorreta. `desktop/app.mjs` e a configuração Upgrade Lab agora identificam explicitamente esse produto, usam AppUserModelID próprio e `lab-profile` ao lado do destino do laboratório. Não muda a identidade/perfil Hold Test.
2. **Cobertura insuficiente de rollback:** o ensaio anterior comparava marcadores inertes e um JSON de preferências sem iniciar os módulos antigos. O roteiro foi ampliado para abrir os módulos originais do ASAR .34 com perfil/identidade explicitamente isolados, renderizar e girar um PDF sintético, conferir a proteção de fechamento e sair normalmente antes/depois das falhas. Executável e desinstalador legados continuam inertes; não se executa o desinstalador original.
3. **Leitor padrão:** o snapshot passou a ler também o valor padrão de `HKCU/Software/Classes/.pdf`, além do UserChoice. Esse dado é somente lido e comparado; não entra na restauração/escrita do registro. Testes verificam ambos.
4. **Roteiro doméstico:** a migração recebe `--source=` explícito para o payload .34 disponível. `--missing=` é opcional e só pode apontar a evidência real da gráfica. Sem ela, a fixture usa o payload completo e registra que o padrão exato de 1.108 ausentes não foi reproduzido. Não se fabrica esse inventário. Backups novos da remoção seletiva também são mantidos.
5. **Detalhe vazio no Setup completo:** o primeiro ensaio de recusa/rollback retornou 2, mas o INI do NSIS tinha `Detail=` vazio. A saída de console do executável GUI não foi um canal confiável para `nsExec`. A restauração dos 19.512 arquivos e do registro foi conferida separadamente e os módulos .34 voltaram a funcionar; o cenário foi mantido como parcial, sem atribuir uma causa não recebida. `desktop/update-transaction.mjs` agora registra fase/mensagem/código da falha no diário e escreve resultado UTF-16 atômico na pasta validada do instalador. `build/nsis/upgrade-transaction.nsh` verifica ação, código e mensagem, com sequência única por chamada, recusando resultados ausentes/antigos/incoerentes mesmo se o processo sair 0. Nenhuma guarda foi removida.

Um primeiro ensaio da nova automação de abertura esperava a conferência de PDF único ao injetar um evento de teste; esse caminho já havia aberto diretamente e renderizado o PDF. A espera foi corrigida para aceitar os dois caminhos reais. Repetição aprovou renderização, edição, preferência sintética e encerramento normal; nenhum motor de importação foi alterado.

## Resultados intermediários comprovados

- Suíte inicial: **486 testes Node aprovados**, zero falhas/skips; o número real supera o registro histórico de 485.
- Regressão real Electron/original-fs: falha de leitura virtual de app.asar reproduzida, leitura física/migração/rollback aprovados.
- NSIS neutro de remoção seletiva: raiz 172, 19.004 entradas; sucesso no caso longo e recusa com restauração nos casos de lock real/inventário inválido.
- 20 testes focados de transação/registro/template/distribuição passaram após os ajustes, incluindo erro real de cópia por arquivo de staging removido e recusa de inicialização do componente após a cópia.
- Primeiro Setup completo reconstruído do checkpoint: instalação limpa em raiz 172, saída 0 em 324.056 ms; 19.510 arquivos esperados comparados, 19.513 instalados contando metadados extras; inicialização privada validada pelo instalador e novamente por CLI. DOCX/DOC sintéticos convertidos sem descoberta externa, originais preservados. Desinstalação seletiva saiu 0 em 131.532 ms, preservou arquivo/pasta vazia do usuário e perfil sintético; associação do laboratório removida e backup verificado mantido. Esse pacote inicial tem o defeito de perfil da abertura normal e **não é a entrega final**; está preservado como `lab-initial-Setup.exe` nas evidências.
- Os módulos .34 em perfil isolado renderizaram/editaram o sintético, mantiveram a preferência, exibiram a proteção de fechamento e encerraram normalmente. Isso não equivale a inicializar o executável legado original ou homologar Word no layout longo antigo.
- Nova regressão real de ASAR/original-fs passou após o ajuste de diagnóstico. Ensaio neutro real `tests/desktop-update-report-lab.mjs` compilou programa GUI sem console e usou o macro NSIS original: erro Unicode preservado, relatório ausente, ação errada e resultado antigo recusados; sucesso coerente aceito.
- A distribuição privada íntegra foi copiada para raiz de aplicativo 196/Office 199; a validação de hashes passou e a inicialização foi recusada. Sondagem nativa independente com `--version` encerrou em `3765269347` (`0xe06d7363`). O resultado demonstra uma falha nesse caminho sem depender de app.asar, mas não identifica a função interna nem confirma a causa exata da recusa antiga da gráfica, cujos logs estão indisponíveis.
- `scripts/build-upgrade-lab.mjs` reconstrói o laboratório e compila duas variantes de falha com o mesmo arquivo de pacote/desinstalador. Instrumenta somente o processo de build do electron-builder 26.18.0, conferindo a identidade exclusiva; não edita node_modules, não habilita injeções no Hold Test e mantém warnings-as-errors. Os três Setups anteriores ao ajuste de diagnóstico foram copiados/conferidos em `before-bridge-fix/` nas evidências antes de reconstruir.

## Ensaios integrais com as correções

- `full-startup-bridge.log` / `full-startup-receipt.json`: Setup completo em raiz 196/Office 199, recusa real de inicialização após copiar/verificar os arquivos, saída 2. Diário registra a fase `verifying` e a causa recebida. Todos os arquivos anteriores e o registro foram restaurados; os módulos .34 renderizaram/editaram o PDF e mantiveram a preferência sintética. A fixture tinha associações anteriores ausentes. Somente após essa prova foi restaurado o namespace inicial do laboratório; arquivos, perfis, diários e backups continuam preservados.
- `full-clean-bridge.log` / `full-clean-receipt.json`: Setup completo em raiz 172, saída 0 em 267.819 ms; abertura normal real do executável .36 aprovada. Hashes integrais, associação própria, leitor padrão/UserChoice, cache, DOCX/DOC privados e remoção seletiva aprovados, com dados/perfil e backup sintéticos preservados.
- `lab-package-audit-final.log`: pacote de laboratório corrigido auditado novamente; ASAR com 471 entradas e Office com 19.410 arquivos, identidade exclusiva e correspondência com o código atual aprovados.
- `full-migrate-bridge.log` / `full-migrate-receipt.json`: falhas após cópia (430.652 ms), após registro (432.506 ms) e com bloqueio real de `resources/app.asar` (154.541 ms), todas com saída 2 e causa esperada recebida. Em cada caso: registro/arquivos anteriores integralmente iguais, atalhos/cache antigos restaurados, dados/pastas de usuário preservados, diário `rolled-back`, abertura/edição PDF pelos módulos .34 e preferência sintética aprovados.
- No mesmo laboratório, migração sem falha injetada aprovada: Setup 0 em 323.502 ms, raiz/diretório e perfil preservados, abertura normal real da .36 aprovada; tema e preferência sintética herdados da .34. Todos os 19.510 arquivos esperados iguais (19.514 presentes, contando metadados/arquivo de usuário), DOCX/DOC privados aprovados, associações/atalhos/cache corretos no namespace exclusivo, nenhuma duplicata global, leitor padrão e UserChoice preservados.
- `node-final-bridge.log`: **489/489 testes Node aprovados**, zero falhas, cancelamentos ou skips, após todas as correções e a adulteração real do staging no teste de integridade. O caso de integridade e o erro real de cópia são testes da transação Node com fixture reduzida; os casos completos de rollback por cópia/registro/inicialização/lock usam NSIS e o payload público .34.
- `active-prepackage.log` / `active-preserved-proof-prepackage.json`: 19.512 arquivos ativos e 80 arquivos do perfil Hold Test iguais por nome/tamanho/SHA256; perfil principal vazio preservado. Registro original/UserChoice iguais; valor padrão de `.pdf` comparado separadamente com o snapshot ampliado obtido antes dos ensaios de Setup. A divergência pré-existente .26/.34 permaneceu intocada.

## Pacote final e conferência de encerramento

Build com `PRINT_GHOSTSCRIPT=none`, configuração existente `electron-builder.hold-native-test.cjs`, NSIS x64 e `--publish never`, saída 0. Mesma identidade `br.com.4printshops.holdtest`, GUID `90ff2ff0-f013-571c-b109-7edb265dacc8`, nome/perfil/diretório Hold Test preservados. Nenhuma execução do Setup final sobre o app ativo.

```text
C:\Users\Marcos\Documents\automacaoAdesivosCoreldraw\Corel Script Configurator\dist-desktop\hold-native-test\4PrintShops-Hold-Test-1.0.0-experimental.36-Windows-x64-Setup.exe
Tamanho: 514.970.609 bytes
SHA256: 1A39DAD02CD34C79E26D7439F5DA076D7118B68469F7775CF2C396F231762B02
Authenticode: NotSigned (experimental)
```

`hold36-setup-audit.log`, `hold36-setup-audit-exit.txt=0` e `%TEMP%/4ps-setup-audit-DOm1Wx/result.json`: 19.510 arquivos extraídos iguais por SHA256 a `win-unpacked`; ASAR com 471 entradas correspondente ao código testado, identidade Hold Test correta e Office íntegro (19.410 arquivos/bibliotecas/fontes/licenças). A entrega por Robocopy numa nova pasta raiz 172/Office 175 também conferiu todos os hashes. Somente os modos internos isolados foram executados: inicialização privada aprovada, DOCX e DOC com duas páginas A4 cada (210,01×297 mm), 105 caracteres/frases/acentos preservados, busca por Office externo desativada e originais intactos. Isso é auditoria/entrega de arquivos, não uma segunda instalação do Hold Test nem execução de seu fluxo normal/perfil real.

`active-final.log`, `active-final-exit.txt=0` e `active-preserved-proof-final.json`: novamente todos os 19.512 arquivos ativos e 80 arquivos do perfil real iguais por SHA256 ao início, perfil principal vazio preservado, registro próprio/UserChoice e padrão `.pdf` iguais. `pipeline-exit.txt=0`. O laboratório atualizado permanece registrado apenas em seu namespace exclusivo, para inspeção; seus diários/backups e os dos ensaios anteriores não foram limpos.

Sintaxe dos 11 JS/MJS/CJS novos/alterados e `git diff --check` aprovados. `git status -sb`: `main...origin/main`, 11 caminhos modificados e cinco novos (16 no total), staging vazio; HEAD permanece `522e9f62ee64056752832cf3fbd6bdd0c131b91d`. Nenhum commit/push. Inventário exato na seção seguinte e em `source-final-check.log` nas evidências.

As falhas iniciais de perfil/diagnóstico e a espera incorreta do roteiro estão identificadas acima e foram repetidas com sucesso após correção. Não há falha essencial aberta nos ensaios locais concluídos. Continuam pendentes os limites de VM/Windows/política, executável legado original, interrupção abrupta e avaliação gráfica listados ao fim; a causa exata da recusa antiga da gráfica não foi recuperada.

## Arquivos da alteração

- Execução/registro da atualização: `desktop/update-transaction.mjs`, `desktop/update-registry.mjs`, `build/nsis/upgrade-transaction.nsh`.
- Identidade/perfil exclusivos do laboratório: `desktop/app.mjs`, `electron-builder.upgrade-lab.cjs`.
- Build reproduzível do laboratório: `scripts/build-upgrade-lab.mjs` (novo).
- Ensaios: `tests/desktop-update-package-lab.mjs`, `tests/update-registry.test.mjs`, `tests/update-transaction.test.mjs`; novos `tests/desktop-update-report-lab.mjs`, `tests/desktop-upgrade-startup.mjs`, `tests/helpers/upgrade-profile-probe.mjs`.
- Documentação: este relatório (novo), `docs/ESTADO_PROJETO.md`, `docs/CONTINUIDADE_EXPERIMENTAL36.md`, `docs/ATUALIZACAO_TRANSACIONAL_EXPERIMENTAL36.md`.

São 16 caminhos de código/testes/documentação. Os motores aprovados de PDF, edição, imposição, N-up, Step & Repeat, livreto, impressão e importação permanecem fora do diff, assim como dependências e configurações do Hold Test. Artefatos, instaladores, runtime, laboratórios e evidências continuam ignorados pelo Git. Os auxiliares de auditoria desta execução estão no workspace ChatGPT, fora do repositório, e seus resultados estão nas evidências.

## Reprodução e avaliação posterior

Comandos abaixo pressupõem o diretório da aplicação, runtime privado já preparado e **namespace de laboratório livre**. Os roteiros recusam um registro prévio; não apagar registros/backups para repetir automaticamente. A migração bem-sucedida pode deixar a .36 do laboratório registrada para inspeção, sempre fora do Hold Test ativo.

```powershell
node scripts/build-upgrade-lab.mjs
node tests/desktop-update-report-lab.mjs
node scripts/test-update-asar.mjs dist-desktop/upgrade-transaction-lab/win-unpacked/resources/app.asar
node tests/desktop-update-package-lab.mjs startup --source=.desktop-test/release34-1794168709616/payload
node tests/desktop-update-package-lab.mjs clean --source=.desktop-test/release34-1794168709616/payload
node tests/desktop-update-package-lab.mjs migrate --source=.desktop-test/release34-1794168709616/payload
node --test --test-concurrency=2 tests/*-*.test.mjs
```

O inventário ausente da gráfica só deve entrar por `--missing=` se o arquivo original for recuperado. A regressão de integridade altera realmente um arquivo do staging **depois** da preparação e exige a recusa pelo hash calculado em `install`, com restauração da fixture; não substitui a guarda por um verificador artificial. Isso é um ensaio de transação Node, diferente dos Setups integrais de cópia/registro/Office/lock.

Próxima avaliação: VM descartável com o instalador .34 original e política de caminhos igual à gráfica, incluindo o inventário real de ausências se recuperado; atualização pelo Setup .36, associações e reinício/uso da aplicação anterior após falhas. Exercitar interrupção abrupta e recuperação manual separadamente. Depois, conferir visualmente documentos sintéticos representativos, fontes, paginação e o fluxo de produção/RIP. Qualquer futura intervenção na instalação ativa precisa de autorização própria e diagnóstico da divergência .26/.34 encontrada em casa. A aprovação dos ensaios deste relatório não autoriza instalar em produção.

## Limites do laboratório

O Upgrade Lab usa `br.com.4printshops.upgradelab.transaction36`, GUID `54308dbf-3408-5266-9771-b636750e1a20`, registro em `Software/4PrintShops/UpgradeLab`, atalhos/cache/perfil próprios. Não registra um leitor PDF real. Verificações extras conferem ausência de registro duplicado no namespace normal do Windows.

O sistema local informa build 10.0.26200 e `LongPathsEnabled=1` na consulta somente leitura. Esse valor não foi alterado; a estação da gráfica tinha 0 nos registros históricos. Os ensaios locais não homologam outra política/Windows. O pacote .34 anterior em `dist-desktop/hold-native-test/win-unpacked` foi identificado e preservado por movimentação integral para `preserved-experimental34-win-unpacked-20261109` no mesmo diretório de artefatos, antes de o builder gerar a .36; não houve exclusão de arquivos. Seu ASAR SHA256 está em `preserved-hold34.json`. O Setup .34 e o payload usado pelas fixtures permaneceram nos caminhos anteriores.

A migração é simulada com os arquivos públicos .34 e marcadores inertes para os dois executáveis legados. A prova funcional abre seus módulos imutáveis pelo Electron de teste com perfil e identidade forçados ao laboratório; isso não substitui uma VM com o Setup/desinstalador .34 original. O Office antigo em caminhos longos já tem limitações conhecidas; não se declara que seu Word está funcional antes/depois. Queda de energia/encerramento forçado, outros Windows/antivírus, fidelidade de fontes/paginação/ICC/RIP e aceitação na gráfica permanecem sem homologação. Não há recuperação automática após interrupção abrupta nem limpeza de backups.

## Consolidação local autorizada — 09/11/2026

Usuário autorizou registrar estas alterações em commit local, sem push. Auditados os mesmos 16 caminhos; nenhuma alteração nos motores aprovados ou no instalador existente. Sintaxe dos 11 JS/MJS/CJS aprovada e tamanho/SHA256 do Setup reconferidos iguais aos acima. A primeira repetição Node no ambiente restrito encontrou EPERM ao renomear arquivos temporários sintéticos e foi interrompida. Repetição fora dessa restrição, sem mudança de código: **489 testes aprovados, zero falhas, cancelamentos ou skips**, em `.desktop-test/commit36-node-unrestricted-20261109-214238.log`. Ensaios integrais anteriores aproveitados, sem reexecutar Setup ou alterar a instalação ativa nesta consolidação. Mensagem do commit: `Consolida validacao local da atualizacao experimental 36`. Instaladores, builds, runtime, perfis, backups e evidências não fazem parte do commit. Os registros de staging vazio/sem commit nas seções anteriores são o estado histórico anterior à autorização. Todas as limitações e pendências acima continuam válidas.
