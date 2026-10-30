# 4PrintShops Desktop — distribuição experimental Windows

Versão `1.0.0-experimental.1`, Windows x64. O major 1 foi mantido a partir do `package.json` existente; o sufixo identifica o primeiro instalador de testes. Não é uma publicação comercial. Nada foi enviado a releases ou serviços de hospedagem.

## Executar em desenvolvimento

Na raiz do projeto, com Node.js/npm instalados (ambiente validado: Node 24):

```powershell
npm ci
npm run desktop
```

Se a política de scripts do npm impedir o download do runtime Electron, executar explicitamente `node node_modules/electron/install.js`. Não é necessário aprovar scripts de pacotes de outros alvos, como Squirrel. O launcher de desenvolvimento retira do processo filho `ELECTRON_RUN_AS_NODE`, variável que pode ser herdada de um IDE Electron.

A versão de navegador continua disponível com `node app.js` ou `npm start`, em `http://localhost:3000`. O servidor escuta somente `127.0.0.1`. O aplicativo instalado não exige Node.js, npm, VS Code ou servidor externo.

## Arquitetura e dados

| Execução | Entrada e origem | Salvamento |
| --- | --- | --- |
| Navegador | `app.js` → HTTP local, `http://localhost:3000` | Mesmo download Blob do fluxo existente. |
| Desktop | `desktop/main.mjs` → `app://4printshops/index.html` | Diálogo nativo seguido de gravação temporária e renomeação no destino escolhido. |

O Desktop serve os assets empacotados pelo protocolo próprio, sem iniciar servidor HTTP, reservar porta ou carregar o backend CorelDRAW. A origem permanece fixa entre reinícios e versões. PDF.js, worker, pdf-lib, HTML/CSS e os módulos PDF/QR são os mesmos da versão web.

- `desktop/app.mjs`: instância única, janela, protocolo, restrições e encerramento. Uma segunda abertura restaura/focaliza a janela existente.
- `desktop/security.mjs`: validação de origem/frame/IPC, caminhos permitidos e dados de exportação.
- `desktop/preload.cjs`: somente abrir PDFs escolhidos pelo operador, salvar PDF, informar estado do documento e solicitar a verificação opcional de atualizações. Não expõe `require`, caminhos arbitrários ou IPC genérico.
- `desktop/files.mjs`: recebe os caminhos exclusivamente dos diálogos do processo principal. A importação devolve nome/bytes, não o caminho; a validação do conteúdo continua nos parsers existentes. O salvamento valida PDF e nome simples; falhas não truncam o destino existente. Exportações ou seleções de importação acima de **512 MiB** são recusadas nesta ponte experimental, sem rasterização para reduzir o arquivo.
- `public/js/app/save-file.js`: escolhe a ponte nativa quando disponível; no navegador mantém Blob/download. Cancelamento/falha preserva pendências; somente a gravação nativa concluída marca a revisão capturada como exportada. `desktop-client.js` integra o seletor nativo aos mesmos inputs/eventos de importação e união, com listener removido ao desmontar PDFs.

Preferências, formatos personalizados, armazenamento Chromium e configurações Desktop ficam em **`%APPDATA%\4PrintShops`**, separados da instalação. Reinstalar/atualizar conserva essa pasta. A desinstalação também a conserva (`deleteAppDataOnUninstall: false`). Preferências do navegador e do Desktop têm origens/perfis diferentes e não são migradas automaticamente. Documentos abertos e alterações em memória não são recuperados após reiniciar.

A janela conserva título, Início, abas, menus HTML e painéis; não recebe uma segunda barra nativa de menus. Pode ser redimensionada, minimizada e maximizada. O ícone experimental usa `public/img/logo.png`, convertido pelo builder para o executável. Uma identidade compacta definitiva pode substituir `win.icon` e o ícone da janela; o logo original não foi alterado.

## Gerar o instalador

```powershell
npm run desktop:pack
npm run desktop:dist
node scripts/check-desktop-package.mjs
```

`desktop:pack` prepara a pasta sem gerar o instalador. `desktop:dist` gera NSIS x64 com `--publish never`. Saídas:

- **`dist-desktop/4PrintShops-1.0.0-experimental.1-Windows-x64-Setup.exe`** — instalador.
- `dist-desktop/4PrintShops-1.0.0-experimental.1-Windows-x64-Setup.exe.blockmap` — metadado do builder; sua existência não habilita atualizações.
- `dist-desktop/win-unpacked/4PrintShops.exe` — aplicativo empacotado antes da instalação.

A auditoria verifica assets essenciais, exclusões e igualdade dos arquivos de runtime no ASAR com o código atual. O pacote inclui apenas o frontend ativo, bibliotecas locais, processo Desktop e dependências de produção da atualização. Exclui testes, modelos CorelDRAW, backend legado, documentos PDF, chaves, logs, backups e dependências de desenvolvimento. As pastas geradas são ignoradas pelo Git. O repositório já continha `node_modules` versionado; esta tarefa não remove esse histórico. A fonte reproduzível de dependências é o `package-lock.json` da raiz.

## Instalar e desinstalar

1. Executar o `Setup.exe`, escolher a pasta e concluir. A instalação é **por usuário**, sem instalação de Node e sem solicitar elevação para instalar para todos os usuários.
2. Abrir pelo atalho da Área de Trabalho ou do menu Iniciar. O instalador não abre o programa automaticamente ao finalizar.
3. Para remover, usar Aplicativos do Windows ou `Uninstall 4PrintShops.exe` na pasta escolhida. A pasta de preferências é conservada.

O instalador desta fase **não possui assinatura Authenticode**. O Windows pode pedir confirmação de procedência/SmartScreen. Não foi desativada nenhuma proteção do Windows. A validação foi feita no Windows 10 x64, build 19045; instalação em outros computadores e Windows 11 precisa de aceitação complementar.

## Segurança e encerramento

`contextIsolation: true`, `nodeIntegration: false`, sandbox, `webSecurity`, CSP e validação do frame principal permanecem ativos. O renderer não navega para sites/arquivos externos, abre popups ou incorpora webviews; permissões adicionais são negadas. O protocolo só lê caminhos permitidos dentro do pacote. Os fuses desabilitam execução como Node, opções de Node por ambiente e inspeção do processo principal; exigem ASAR e sua integridade.

O estado de pendência continua vindo do PDFWorkspace. `beforeunload` aciona a confirmação nativa ao fechar, inclusive com PDF em aba oculta. Cancelar conserva o documento. Operações em curso/gravação devem terminar antes de fechar. Fechar o aplicativo termina seus processos Electron; não existe backend filho para deixar órfão. As portas de depuração usadas nos harnesses são somente opções explícitas dos testes, não da inicialização normal.

Cancelar o diálogo nativo de destino impede a gravação. Se o operador já confirmou o caminho e a escrita começou, cancelar depois a janela de exportação conserva a pendência, mas uma cópia que terminou de gravar pode permanecer no destino. Encerramentos forçados pelo sistema ou reinstalações externas não substituem o fechamento protegido; exportar/concluir o trabalho antes de executar manualmente outro instalador.

## Atualizações opcionais

Estão **desativadas por padrão**: sem provedor publicado, agenda, download ao iniciar ou instalação ao sair. No Desktop, Arquivo → Verificar atualizações explica essa condição; no navegador essa ação não aparece.

`desktop/updates.mjs` implementa consulta, nenhuma versão disponível, versão/novidades, download/progresso, pronto, erro/repetição e instalação explícita. `update-ui.mjs` usa texto simples para novidades, progresso na janela e as escolhas **Atualizar agora / Mais tarde**. `autoDownload` e `autoInstallOnAppQuit` são `false`. Instalar exige documento sem pendências/operação; o estado é conferido novamente após a confirmação. A atualização não exporta nem descarta documentos.

A ativação futura exige, após autorização de distribuição:

1. Build e instalador assinados para uma identidade de editor validada; metadado `app-update.yml` do pacote contendo `publisherName`.
2. Provedor HTTPS sem credenciais na URL. Não há token ou endereço privado embutido nesta versão.
3. Configuração local `%APPDATA%\4PrintShops\updates.json` com `enabled: true` e `url` do provedor autorizado. Sem esses requisitos o serviço permanece inativo; a verificação Authenticode do electron-updater não é desabilitada.
4. Teste controlado com duas versões, consentimento, bloqueio com documento pendente, integridade do download e preservação de preferências/arquivos.

Os estados e bloqueios foram testados com um updater simulado. **Não houve teste de atualização real entre duas versões, publicação, hospedagem ou distribuição comercial.** Isso continua pendente; nenhuma aprovação de distribuição é inferida desta estrutura.

## Testes e aceitação

```powershell
npm test
npm run test:browser
npm run test:desktop
```

`test:desktop` usa Electron real, perfil temporário e diálogos simulados no bootstrap exclusivo dos testes, que não é empacotado. Verifica ponte, gravação real, cancelamento/falha, proteção, sandbox, origem, persistência e encerramento.

`tests/desktop-installed.mjs` é o harness técnico da instalação temporária identificada em `%TEMP%/4printshops-instalacao-teste.txt`; usa o executável instalado, sem mocks no pacote e sem Node no PATH do aplicativo. A automação Windows restringe-se ao PID/título/controles da instância de teste. Seus resultados e artefatos estão registrados em `ESTADO_PROJETO.md`; não devem ser confundidos com a aprovação manual em uma máquina de atendimento.

Aceitação manual: abrir o programa e tentar uma segunda abertura; importar PDF real; girar/reordenar/excluir/unir/redimensionar; conferir a cópia preparada de impressão; salvar em pasta escolhida e cancelar outra tentativa; alternar Início/PDFs/QR; fechar e cancelar com trabalho pendente; repetir com janela pequena/maximizada; reiniciar e conferir preferências; desinstalar/reinstalar e conferir dados. Comparar PDFs complexos no Acrobat/RIP antes de liberar para produção.

Para distribuição comercial futura ainda faltam assinatura de código e identidade final, revisão de licenças/avisos das bibliotecas, homologação em outras máquinas, canal de atualizações autorizado e seu teste entre versões, política de suporte/recuperação e aceitação de desempenho/fidelidade em arquivos reais. Não há autenticação, serviço pago ou novas ferramentas de edição nesta implementação.

Referências técnicas: [segurança do Electron](https://www.electronjs.org/docs/latest/tutorial/security), [protocolos locais](https://www.electronjs.org/docs/latest/api/protocol), [instalador NSIS](https://www.electron.build/nsis/), [assinatura Windows](https://www.electron.build/docs/features/code-signing/code-signing-win/).
