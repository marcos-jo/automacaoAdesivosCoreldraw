# Abrir PDFs pelo Windows — experimental.4

Implementação em 03/11/2026, base `90cc0ce`. **O usuário concluiu pessoalmente os quatro testes manuais com o 4PrintShops PDF Test experimental.4, informou que todos passaram e aprovou a funcionalidade para checkpoint no Git.** O teste em compartilhamento SMB real permanece pendente. Não houve commit/push nem impressão nesta auditoria.

Histórico: a automação visual anterior por Computer Use foi interrompida com Esc e não foi retomada. Isso não se confunde com a validação manual posterior, realizada e aprovada pelo usuário. Os testes automatizados instalados descritos abaixo usaram o Shell real do Windows.

## Associação sem trocar o leitor padrão

`build/pdf-association.nsh` é incluído pelo NSIS. Registra por usuário o ProgID `<appId>.PDF`, ícone, comando de abertura, `Applications/<executável>/SupportedTypes`, capacidades e `RegisteredApplications`, além do valor próprio em `.pdf/OpenWithProgids`. O comando é `"<executável>" --open-pdf "%1"`, com caminhos entre aspas. Não grava o valor padrão de `.pdf`, `UserChoice`, verbos print ou chaves do Acrobat/outros leitores. A associação padrão é sempre escolha do operador no Windows.

Na desinstalação, só remove as chaves/valores próprios se o comando ainda aponta para aquela instalação; containers da extensão são removidos somente se estiverem vazios. Não escolhe automaticamente outro leitor ao remover o aplicativo. Se o usuário o tiver definido como padrão, deverá escolher seu substituto pela interface do Windows.

O macro padrão de `fileAssociations` do builder 26 não foi utilizado porque pode gravar o default da extensão. O include customizado permite registrar Abrir com sem essa alteração. Configs do laboratório de update desabilitam esse include, evitando registrar o Lab como leitor.

## Abertura e documentos em andamento

- Caminhos da primeira inicialização são enfileirados antes de carregar a interface. A ponte só entrega quando o renderer pede a lista; não se depende de um evento enviado cedo demais.
- Segunda abertura usa a instância única existente e `additionalData`, conservando os argumentos originais antes de possíveis modificações do Chromium. A janela é restaurada se minimizada e trazida à frente.
- Sem documento e com um único PDF, a aba PDFs é ativada e ele é importado automaticamente, com PDFWorkspace, PDF.js, miniaturas e operações existentes.
- A arquitetura tem uma aba por ferramenta e um workspace de PDF. Se há documento aberto ou vários arquivos, a janela contextual oferece **Substituir/Abrir primeiro**, **Adicionar ao documento** ou **Manter o atual/Dispensar abertura**. Não existe um segundo sistema de abas.
- Substituir reutiliza a confirmação de alterações não exportadas. Cancelar mantém o PDF, o giro e a pendência. Adicionar usa a união transacional existente e preserva edições; iniciar um documento por união marca trabalho pendente. Downloads continuam só pelo fluxo de exportação (extração legada permanece como antes).
- Quando solicitações chegam durante uma operação/decisão, as demais ficam na fila. Abrir primeiro conclui somente aquele pedido. IDs/caminhos pendentes normalizados são deduplicados; reabrir intencionalmente depois da conclusão é permitido. A ordem é a de chegada das invocações do Windows, não uma garantia de ordem alfabética da seleção.
- Enquanto a importação externa está ativa, a interface e os diálogos nativos concorrentes são protegidos. Atualização não pode instalar enquanto há PDFs externos pendentes/ativos. Erros permitem repetir ou dispensar, sem substituir parcialmente o workspace. A fila é de sessão, não sobrevive ao encerramento do aplicativo.

## Validação de entrada e limites

`desktop/external-pdfs.mjs` aceita apenas caminhos absolutos de filesystem Windows com extensão PDF, incluindo UNC e prefixos longos de disco/UNC. Rejeita URLs, relativos, ADS e dispositivos. Argumentos de opções não são tratados indiscriminadamente como documentos. O renderer recebe nomes/IDs e bytes; não pode fornecer caminhos arbitrários via IPC. Os canais exigem o frame/origem da janela existente.

A leitura verifica arquivo regular, tamanho, acesso, cabeçalho `%PDF-` nos primeiros 1024 bytes e alteração durante a leitura. O parsing completo continua nas bibliotecas atuais. Limite agregado de **512 MiB por importação** e **128 caminhos pendentes**. Exceder a fila não remove pedidos anteriores. A leitura tem deadline de 30 s; uma chamada filesystem de rede já bloqueada no sistema operacional pode demorar a terminar. Nesse caso novas leituras externas são recusadas até a anterior terminar, sem tornar possível acumular leituras ilimitadas. Originais são abertos somente para leitura.

Caminhos reais acima de 260 caracteres foram lidos em Node no Windows; componentes individuais continuam sujeitos aos limites do filesystem/Shell. UNC com nome simples de servidor foi validado em sintaxe, mas não havia share SMB de teste: permissões/latência/credenciais reais de rede precisam de teste do operador. **Limitação identificada na auditoria:** nomes de servidor com pontos (FQDN ou IPv4) são rejeitados pela validação atual, mesmo antes da leitura. Deve ser corrigida e testada na futura etapa SMB; o código aprovado não foi alterado nesta auditoria. Arquivos protegidos por senha seguem as limitações anteriores do módulo.

## Instalador de teste e testes executados

Gerado e realmente instalado:

`dist-desktop/pdf-open-test/4PrintShops-1.0.0-experimental.4-Windows-x64-Setup.exe`

- Produto **4PrintShops PDF Test**, appId `br.com.4printshops.opentest`, perfil `%APPDATA%\4PrintShops PDF Test`. Não substitui a instalação aprovada nem os pacotes de update .2/.3.
- Config `electron-builder.open-test.cjs`; versão de fonte/manifests `.4`, sem dependências novas. O build comum 4PrintShops usa o mesmo runtime/include e pode ser gerado com `npm run desktop:dist`, mas o instalador comum não foi gerado nesta rodada.
- SHA-256: `3158d2c2e6a4c51923a2f50448746948a357da63fa8105976f916ad9a1177ea0` (123.035.080 bytes). Experimental sem Authenticode; não foi desativada proteção do Windows.
- Auditoria ASAR: **412 entradas**, fontes de runtime idênticas, sem clientes/chaves/testes/legado empacotados.
- **184 testes Node aprovados**, zero falhas/skips. Incluem caminhos/argumentos, UNC/longos em sintaxe, leitura real longa/acentos/espaços, fila/deduplicação/inicialização anterior ao renderer, inexistente/inválido/limite, cancelamento e aceite de substituição no módulo real. Log `%TEMP%/4printshops-open-node-complete.log`.
- Chrome headless curto: PDFs/QR, abas 100/350, exportação e proteção aprovados. Artefatos `%TEMP%/4printshops-browser-rc6DdA/`, log `%TEMP%/4printshops-open-browser.log`.
- Electron em desenvolvimento: regressões existentes aprovadas, inclusive cópia seletiva, exportação/cancelamento/falha/proteção. Artefatos `%TEMP%/4printshops-desktop-vnOw2Q/`, log `%TEMP%/4printshops-open-desktop.log`.
- `tests/desktop-open-installed.mjs` aprovado, com NSIS instalado e sem mocks no pacote. Abertura fria com arquivo/renderer ainda não pronto; chamadas reais `ShellExecuteEx` pelo ProgID registrado (equivalentes à invocação por arquivo do Shell); acentos/espaços, janela minimizada restaurada, segunda instância, cancelamento de substituição, rajada de dois PDFs sem perda, adição após giro, exportação nativa de sete páginas com giro preservado, arquivos inválidos/inexistentes com controles responsivos e originais idênticos. Registro de default e `UserChoice` mantido antes/depois da instalação e desinstalação, associação própria removida. Log final `%TEMP%/4printshops-open-installed-restored.log`, artefatos `%TEMP%/4printshops-open-installed-6uVTbj/`.
- A reinstalação temporária para conferência visual foi removida após o Esc pelo desinstalador próprio, sem tocar no app principal. Preferências de teste ficam preservadas conforme a política existente.
- Sintaxe dos **14 JS/MJS/CJS** alterados/novos, parsing do helper PowerShell e `git diff --check` (incluindo whitespace dos novos não rastreados) aprovados.

A aprovação dos quatro testes manuais de abertura é um resultado informado pelo usuário, separado das execuções automatizadas acima. Não se afirma testado um share SMB real. O harness de update assinado não foi executado; a homologação da instalação/reinício automáticos pelo atualizador continua pendente.

## Escolher leitor padrão no Windows 10 (manual)

Depois de instalar o pacote desejado, clicar com botão direito em um PDF → **Abrir com → Escolher outro aplicativo**. Selecionar **4PrintShops** (ou **4PrintShops PDF Test**, neste instalador isolado). Para testar sem trocar o padrão, deixar desmarcada a opção de sempre usar o aplicativo. Se não aparecer na lista curta, expandir Mais aplicativos ou localizar o executável instalado.

Para tornar padrão por decisão própria, marcar **Sempre usar este aplicativo para abrir arquivos .pdf** e confirmar. Alternativamente, em **Configurações → Aplicativos → Aplicativos padrão → Escolher aplicativos padrão por tipo de arquivo**, localizar `.pdf` e escolher o programa. O aplicativo/instalador não faz essa escolha automaticamente. Orientação oficial: [Microsoft — alterar aplicativos padrão](https://support.microsoft.com/en-us/windows/apps/change-default-apps-in-windows).

## Roteiro de validação e repetição

Mantido como referência para novas validações; os quatro testes manuais informados pelo usuário já foram aprovados. A pendência de rede é o compartilhamento SMB real, incluindo a limitação de FQDN/IPv4 descrita acima.

1. Instalar o PDF Test em uma pasta de teste; conferir Abrir com e que seu leitor padrão anterior permanece selecionado.
2. Usar Abrir com em um PDF com espaços/acentos, primeiro com o app fechado; conferir nome, páginas, miniaturas e conteúdo.
3. Girar uma página, minimizar e abrir outro PDF. Cancelar a substituição e conferir edições; repetir adicionando ao documento e exportar para outro nome/pasta.
4. Selecionar dois ou mais PDFs no Explorer e abrir com o aplicativo escolhido; conferir pedidos/contagens, decisões de união e ausência de duplicação. Se o Windows não oferecer Abrir com na seleção múltipla, invocar arquivos individualmente com o aplicativo já aberto; não mudar silenciosamente o padrão para contornar o teste.
5. Se desejar, escolher voluntariamente como padrão, fechar e testar duplo clique; repetir com programa aberto e PDF pendente em aba oculta.
6. Testar arquivo inválido, removido após a seleção, nome/caminho longo e share UNC acessível. Conferir erros e que controles voltam a responder; não configurar credenciais/segurança de rede para forçar aprovação.
7. Desinstalar o PDF Test, conferir remoção de suas entradas e preservação dos outros leitores. Voltar ao seu padrão preferido pela interface do Windows, se o tiver alterado manualmente.

## Auditoria final do checkpoint

Base/último commit: `90cc0ce9e83619127669f6f935042998ea2d192c` (`Implementa atualizador Desktop e prepara laboratorio de assinatura`). **Staging vazio; nenhum commit ou push executado.** Inventário: **21 candidatos, 13 modificados e 8 novos, sem remoções**. `git diff --stat` não inclui os novos não rastreados; eles foram conferidos separadamente.

Todos os candidatos pertencem à associação/abertura, integração, testes ou documentação. `build/pdf-association.nsh` é fonte necessária do instalador, não um build gerado. Os ajustes em `electron-builder.update-lab.cjs` e `scripts/authenticode-lab/prepare-kit.mjs` são compatibilidade necessária: impedir associação do Lab e incluir o novo módulo no kit. Não foram identificados arquivos de outros trabalhos ou alterações acidentais; nenhuma alteração local foi descartada.

| Estado | Arquivo | Finalidade |
| --- | --- | --- |
| M | `desktop/app.mjs` | Inicialização, instância única, fila/IPC e proteção do atualizador durante pedidos externos. |
| M | `desktop/preload.cjs` | Ponte restrita de nomes/IDs, leitura e conclusão/liberação dos pedidos. |
| M | `desktop/security.mjs` | Allowlist do novo módulo frontend. |
| A | `desktop/external-pdfs.mjs` | Argumentos, caminhos, leitura e fila com limites/deduplicação. |
| M | `electron-builder.config.cjs` | Empacotar frontend e incluir fonte NSIS. |
| M | `electron-builder.update-lab.cjs` | Manter Lab fora das associações do Windows. |
| A | `electron-builder.open-test.cjs` | Build isolado experimental.4 para testar instalação/desinstalação. |
| A | `build/pdf-association.nsh` | Registro/remoção de capacidade e Abrir com sem alterar padrão. |
| M | `package.json` | Versão experimental.4, sem novas dependências. |
| M | `package-lock.json` | Versão coerente com o manifest. |
| M | `public/js/main.js` | Conectar ao shell de abas existente. |
| M | `public/js/tools/pdfs.js` | Importação compartilhada e getters, preservando workspace/proteção. |
| A | `public/js/app/desktop-open.js` | Recebimento e decisões sem substituição silenciosa. |
| M | `scripts/authenticode-lab/prepare-kit.mjs` | Incluir o novo módulo nos kits futuros. |
| M | `tests/pdf-integration.test.mjs` | Regressão de importação e cancelamento/aceite de substituição. |
| A | `tests/desktop-external-pdfs.test.mjs` | Caminhos, fila, duplicatas, limites e leitura real. |
| A | `tests/desktop-open-installed.mjs` | Fluxo instalado/Shell, registro, exportação e desinstalação. |
| A | `tests/helpers/windows-open-pdf.ps1` | Invocar ProgID de teste pelo Shell sem mudar o padrão. |
| M | `docs/ESTADO_PROJETO.md` | Aprovação manual, auditoria e pendências. |
| M | `docs/DISTRIBUICAO_DESKTOP.md` | Distribuição e estado atual da validação. |
| A | `docs/ABRIR_PDFS_WINDOWS.md` | Funcionamento, roteiro, limites e esta lista completa. |

Verificações desta auditoria: **184 testes Node aprovados, zero falhas/skips**, log `%TEMP%/4printshops-pdf-open-audit-node.log`; sintaxe dos **14 JS/MJS/CJS**, parsing do helper PowerShell e `git diff --check` aprovados. Chrome headless curto aprovado: navegação, abas 100/350, QR, exportação e proteção; log `%TEMP%/4printshops-pdf-open-audit-browser.log`, artefatos `%TEMP%/4printshops-browser-c8b5XZ/`. Electron real em perfil temporário aprovado: worker/miniaturas, edição/exportação, cancelamento/falha, proteção, persistência/sandbox e cópia seletiva; log `%TEMP%/4printshops-pdf-open-audit-desktop.log`, artefatos `%TEMP%/4printshops-desktop-OQ3Yfq/`.

PDFWorkspace, módulo de proteção e motores do atualizador não têm alterações diretas desde `90cc0ce`; a integração acrescenta o bloqueio por pedidos externos. Não se repetiu instalação/desinstalação, que poderia interferir no PDF Test usado pelo usuário, nem builds ou a suíte longa sem mudança funcional nesta auditoria. Os resultados anteriores e a aprovação pessoal são preservados sem inventar novas execuções.

O `.gitignore` continua protegendo instaladores/builds/kits, logs/temporários, PDFs, certificados/chaves e credenciais locais. Nenhum desses artefatos é candidato; os scripts de teste e NSIS são fontes necessárias. A base já versionava dependências em node_modules; não foram removidas ou modificadas nesta etapa.

Mensagem proposta: `Adiciona abertura de PDFs pelo Windows ao 4PrintShops`. **Aguardar autorização explícita após análise deste relatório**, adicionando somente os 21 arquivos acima. Aprovação funcional não equivale à autorização para criar o commit agora. Pendências mantidas: SMB real/FQDN/IPv4 e homologação de instalação automática assinada do atualizador.
