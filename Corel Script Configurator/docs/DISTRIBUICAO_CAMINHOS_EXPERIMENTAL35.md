# Distribuição privada do LibreOffice — diagnóstico e experimental.35

## Diagnóstico da remoção antiga concluído em laboratório — 09/11/2026

O [relatório específico de atualização NSIS](DIAGNOSTICO_UPGRADE_NSIS35.md)
comprovou falha de un.atomicRMDir no caminho longo:18304 marcadores com nomes
públicos passaram em raiz curta, falharam em raiz172 num recurso de260 caracteres
e passaram com prefixo estendido, todos com conteúdo conferido. API Win32 comum
retorna erro3 para esse recurso, enquanto a forma estendida o acessa; teste de
lock real separado continua recusado com erro32 e rollback. Na instalação ativa,
somente leitura confirmou o mesmo problema de endereçamento e18402 arquivos
presentes iguais por hash ao pacote .34; os1108 faltantes são os originais e75
arquivos do perfil permanecem íntegros. Caixa do Setup12488 não foi tocada.

Novo teste opt-in `tests/desktop-nsis-upgrade-lab.mjs`: seis casos com executáveis
neutros sem registro/perfil real, argumentos do desinstalador, reprodução/rollback,
alias, caminhos estendidos, lock e ordenação de metadados.14 testes Node focados
passaram. Alias sozinho não foi aprovado como migração: initMultiUser relê a
pasta longa do registro. Correção proposta exige componente de compatibilidade
auditado e validação completa antes de novo pacote; não implementada em produção.
Não houve nova versão/instalação real. Procedimento de cancelamento e recuperação
das entradas próprias de Abrir com está apresentado, **não executado**, no guia.

## Mensagem da segunda tentativa identificada — 09/11/2026

Transcrição do operador: “Não é possível fechar o 4PrintShops Hold Test. Feche a
janela do 4PrintShops Hold Test e clique em repetir para continuar.” Ícone amarelo,
mesmo título do instalador; botões relatados como “repetir” e “continuar”, sem
confirmação visual dos rótulos. Nenhum botão acionado automaticamente; Setup
PID12488 continua ativo e o registro de versão continua .34.

Não há processo do aplicativo, do conversor ou desinstalador ativo na consulta:
somente o Setup. O texto corresponde a `appCannotBeClosed`, reutilizado por
vários pontos do template. Em `include/installUtil.nsh`, o mesmo aviso aparece
após mais de cinco tentativas sem sucesso do desinstalador anterior. O temporário
contém old-uninstaller.exe, idêntico por SHA256 ao desinstalador instalado .34.
Isso, somado ao estado abaixo, aponta para falha da remoção anterior; não tratar
o texto como prova de aplicativo aberto nem presumir um arquivo bloqueado.
A operação/arquivo exato que provoca o retorno do desinstalador ainda não foi
identificado. Nenhuma execução direta de desinstalador foi feita pela automação.

**Efeito no registro detectado:** em comparação à coleta antes da primeira
tentativa, desapareceram ProgID/Applications/Capabilities/RegisteredApplications
próprios do Hold Test. O UserChoice do leitor PDF padrão continua igual. O hook
customUnInstall remove essas entradas antes da fase de arquivos, que pode abortar;
restaurar os arquivos não implica restaurar essas entradas. Não houve snapshot
entre as duas tentativas que permita atribuir a remoção especificamente à segunda.
Não fizemos restauração manual ou modificação de registro. Isso impede afirmar
que toda a instalação .34 esteja intacta, apesar dos arquivos/perfil preservados.

Evidências: `warning-reported.json` e `associations-at-warning.json` na pasta
install35-visible-O4UnEP, mais o snapshot original preservado. Manter o aviso
aberto para registro pelo operador, sem Repetir/Continuar/Fechar automático.
Investigar o desinstalador anterior em laboratório antes de qualquer novo pacote
ou tentativa, sem enfraquecer a proteção de trabalho/documentos ou caminhos.

## Segunda tentativa autorizada, visível — 09/11/2026

Operador autorizou uma única reprodução com interface visível para registrar a
mensagem, sem fechar avisos. Antes de iniciar: .34 confirmada, processos do
aplicativo/conversor/Setup ausentes, SHA256 FC319905…32AB108 integralmente
reconferido. Ambos os backups anteriores e75 arquivos atuais do perfil conferidos
por hash. Nova cópia dos75 arquivos/8.451.685 bytes verificada.18 caminhos
pendentes preservados e inventariados antes do novo registro documental.

Setup iniciado uma vez às13:47:26 locais, PID12488, por Start-Process com
WindowStyle Normal e argumento somente `/D=` do diretório existente; **sem /S**.
Janela “Instalação do 4PrintShops Hold Test” confirmada. Nenhum botão foi acionado
automaticamente. Solicitado ao operador que avance somente nas telas normais e
transcreva qualquer aviso/erro, deixando-o aberto. Não há resultado final ou
aprovação da instalação nesta observação. O monitor somente registra títulos
e, quando o processo terminar por ação normal, código de saída; não fecha janelas.

Evidências próprias em `.desktop-test/install35-visible-O4UnEP/`: `preflight.json`,
`profile-proof.json`, `profile-backup`, `worktree.json`, `visible-started.json`,
`process.json` e `window-titles.log`; ponteiro `install35-visible-current.json`.
Primeira tentativa e seus scripts/logs intactos; nenhum código de produção,
registro, permissão, impressora, política de segurança, commit ou push alterado
manualmente. Não executar outra tentativa nem contornar verificação se falhar.

## Encerramento da tentativa real — 09/11/2026

O operador informou ter fechado a janela que aparentava erro. Nova leitura
confirmou encerramento do Setup/controlador e resultado código2 em
`installer-result.json`; duração6.157.013ms inclui o tempo aguardando a janela,
não representa tempo útil de instalação. Saída de console capturada vazia,
sem o texto exato da mensagem. Não houve segunda execução. Registro continua
.34;75 arquivos do perfil e cinco arquivos principais foram reconferidos iguais
aos anteriores. Evidência `after-error-state.json` na mesma pasta da tentativa.
A causa específica ainda não está comprovada. Não é possível restaurar a janela
encerrada; eventual reprodução requer diagnóstico e autorização antes de repetir
o instalador. Sem instalação .35 homologada, sem novas conversões/impressão.
O estado de processo ativo na seção abaixo é histórico e foi superado por este
encerramento com falha, mantendo-se todos os logs, scripts e backups.

## Retomada da tentativa real — 09/11/2026, situação ainda pendente

O Setup autorizado iniciou às11:57:34 locais (14:57:34Z), PID7908, pelo script
`.desktop-test/run-install35.mjs`. Na retomada às13:32, ele e seu controlador
Node PID5932 ainda estavam ativos, sem `installer-result.json` nem resultado
final. Existe a janela “Instalação do 4PrintShops Hold Test”; o texto interno não
foi obtido porque o controle visual falhou ao inicializar (`failed to write
kernel assets`). Foi solicitado ao operador o texto da janela, sem clicar em
Repetir ou abrir outro Setup. Não foi iniciada uma segunda tentativa nem
encerrado qualquer processo.

Estado confirmado por leitura nesta retomada:

- Registro e package.json do ASAR continuam1.0.0-experimental.34.
- Cinco arquivos principais instalados iguais aos anteriores; novo `lo` ausente.
- Runtime antigo continua com1108 ausentes; Word permanece bloqueado na .34.
-75 arquivos atuais do perfil e os75 do backup correspondem integralmente aos
  hashes anteriores; backup8.451.685 bytes íntegro.
- Os18 caminhos pendentes correspondem byte a byte ao inventário anterior à
  instalação; base main/c85efe2, staging vazio.
- Temporário `nsf81AC.tmp` contém `old-uninstaller.exe` e plugins; não há processo
  filho do Setup na leitura realizada. Isso localiza a tentativa antes da entrega
  do pacote novo, mas não comprova a causa da espera ou uma falha do desinstalador.
- Consulta de eventos de aplicativo não forneceu evidência adicional de falha.

Evidências preservadas em `.desktop-test/install35-ci7i9u/`: `installation.json`,
`installer-started.json`, `before/profile-proof.json`, `before/profile-backup`,
`before/worktree.json`, `before/associations.json` e `resume-state.json`.
Os dois scripts anteriores foram apenas lidos, sem alteração ou reexecução.
Não há homologação da instalação .35 nem conversões na instalação nova a relatar.
Próximo passo: identificar a mensagem pendente e decidir com o operador; eventual
falha não autoriza reinstalação automática. As provas de laboratório abaixo
continuam válidas, mas não substituem a atualização real ainda inconclusiva.

## Diagnóstico confirmado em 09/11/2026

Base `main/c85efe2903b70bcee11da272328e957cecd52ad5`. Três documentos locais da
.34 já estavam modificados; foram preservados e copiados para evidências antes
das novas entradas. Nenhuma nova instalação do aplicativo foi realizada.

O destino ativo continua o diretório Hold Test anterior, com raiz do aplicativo
de172 caracteres e raiz `resources/libreoffice` de194. Não foram alterados
perfil, associações, registro, permissões, antivírus ou políticas Windows.
`LongPathsEnabled` foi somente consultado:0, mantido assim durante os ensaios.

Inventário confrontado com o manifesto e o Setup .34:

| Item | Resultado |
|---|---|
| Arquivos Office esperados |19410, mais o manifesto |
| Conteúdo Office instalado |18302 presentes, todos legíveis e com SHA256/tamanho corretos |
| Ausentes |1108; nenhuma substituição ou reparação na instalação ativa |
| Payload .34 |Todos os19410 arquivos com os hashes esperados |
| Validador |Recusa correta por ENOENT; não é falha de interpretação do inventário |

### Padrão dos arquivos ausentes

| Grupo | Ausentes |
|---|---:|
| help/media |28 |
| program/python-core-3.13.17 |416 |
| share/Scripts |2 |
| share/config |464 |
| share/extensions |172 |
| share/template |1 |
| share/xpdfimport |25 |

Extensões:414 `.py`,303 `.xml`,162 `.ui`,138 `.properties`,31 `.png`,30 sem
extensão e30 `.gen/.cfs/.xhp/.txt/.ott/.default/.js` no total,
discriminados no inventário local. A lista completa, com os1108 nomes relativos,
grupo, extensão e comprimento, está em
`.desktop-test/office35-wIwpZu/inventory.json` e `missing-files.csv`.

Os caminhos completos ausentes variam255–294 caracteres;59 têm menos de260.
Há arquivos presentes com caminho de276 caracteres. Portanto uma regra única
“todo caminho maior que260 falha” é falsa. A profundidade interfere de formas
diferentes nas duas primitivas usadas pelo instalador, inclusive na criação dos
diretórios. A relação foi demonstrada executando o mesmo processo, não inferida
somente desses comprimentos.

### Causa: cópia final e fallback do NSIS

O template real de electron-builder26.18.0,
`templates/nsis/include/extractAppPackage.nsh`, faz:

1. Nsis7z extrai o pacote para `$PLUGINSDIR/7z-out`.
2. `CopyFiles` tenta copiar essa árvore para o destino final.
3. Após erros/repetições, o fallback extrai diretamente no destino longo e
   ignora erros dessa última tentativa. Não havia conferência integral final.

Ensaio com NSIS3.0.4.1 e os mesmos plugins, apenas marcadores sintéticos com os
19410 nomes públicos do Office, sem registro, atalhos ou instalação de aplicativo:

| Operação | Resultado |
|---|---|
| Extração temporária |19410 arquivos, contagem conferida antes da cópia |
| CopyFiles para destino curto |19410 copiados, sem erro |
| CopyFiles para destino com o mesmo comprimento do ativo |Erro, somente5311 copiados |
| Fallback Nsis7z para esse destino longo |18302 presentes; **exatamente os mesmos1108 nomes ausentes** da instalação ativa |
| Robocopy para outro destino igualmente longo |19410 presentes; retorno1, que indica cópia realizada |
| Adaptador definitivo |19412 marcadores incluindo guardas/manifesto, todos entregues no destino longo |

Isso confirma o defeito no processo de entrega ao diretório final. O pacote e a
extração curta estão completos. O diretório temporário da instalação original
não foi preservado pelo NSIS; sua integridade foi demonstrada na reprodução,
sem alegar inspeção retroativa. Consulta ao antivírus não retornou detecções
relacionadas; isso não equivale a prova universal de ausência de antivírus.
O padrão exato reapareceu com conteúdo sintético, sem mudar permissões/políticas,
e desapareceu ao mudar a primitiva de cópia, descartando essas hipóteses como
causa primária da reprodução. A proteção Office permanece indispensável.

Na máquina de casa os relatórios registram instalação em `%LOCALAPPDATA%/Programs`,
mais curta, com inventário integral aprovado. Não houve acesso remoto nem nova
execução naquela máquina nesta etapa.

Referências primárias: [NSIS, comandos e limites de caminhos](https://nsis.sourceforge.io/Docs/Chapter4.html),
[Microsoft, Robocopy e códigos de retorno](https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/robocopy),
[Microsoft, limites das APIs de caminhos](https://learn.microsoft.com/en-us/windows/win32/fileio/maximum-file-path-limitation).

## Correção isolada

### Segunda limitação: inicialização do próprio LibreOffice

A primeira correção entregou os19510 arquivos reais do aplicativo em uma raiz
de196 caracteres, todos iguais por SHA256 à origem, e a validação nativa retornou0.
Porém a conversão Word falhou. Prova separada capturou `0xC0000409` ao executar
`soffice.com --version` e conversão nesse destino, com runtime íntegro. No caminho
curto de win-unpacked, ambos retornaram0. Na cópia de laboratório com raiz172,
igual ao comprimento da instalação ativa, a conversão também falhou no layout
antigo. Não basta corrigir apenas a cópia dos arquivos.

Somente na cópia própria de laboratório, o runtime foi movido para `lo` na raiz
do aplicativo, mantendo a raiz172 e todos os nomes internos: a raiz Office caiu
de194 para175 caracteres. Manifesto/hashes conferidos, `--version` e conversão
DOCX passaram. Os dumps gerados pelos testes de falha foram preservados na pasta
de evidências, sem apagar arquivos ou relaxar a proteção de extras. A referência
antiga em196 fica como teste de estresse reprovado, não como instalação homologada.
Isso demonstra sensibilidade ao caminho na inicialização; não identifica a função
interna do LibreOffice que aborta nem estabelece um limite universal de caracteres.

### Layout curto e entrega validada

Não foram renomeados arquivos internos do LibreOffice, removidos componentes,
movido o perfil ou criado um cache externo que exija configuração nas estações.
O operador continua instalando somente o 4PrintShops. O runtime passa a ocupar
`lo` junto ao executável, via `extraFiles`; a resolução empacotada acompanha essa
mudança, sem fallback para o layout antigo ou para Office global. A raiz de
instalação, identidade e perfil permanecem iguais. Os outros recursos seguem
em `resources`. Só encurtar nomes não resolveria a cópia legada em todos os casos.

A extração permanece em pasta temporária curta. A entrega final passa a usar o
Robocopy nativo do Windows com suporte a caminhos longos. Não usa `/256`,
`/MIR`, `/PURGE`, mudanças de ACL ou ferramentas encontradas pelo PATH.
`/IS /IT` evita pular arquivos por data/tamanho. Retorno8 ou superior, erro ao
iniciar ou ausência de componentes interrompem a instalação com código2; não
há fallback silencioso para a cópia antiga.

Depois da cópia, o aplicativo executa `--verify-distribution` sem janela, sem
instância normal, sem acesso ao perfil do operador e sem impressão/conversão.
O verificador existente confere o manifesto confiável e todos os arquivos,
tamanhos, hashes, componentes obrigatórios, links e extras. Em seguida consulta
`--version` do próprio runtime com perfil temporário, ambiente privado e processo
isolado. Falha de inicialização também impede sucesso; integridade sozinha não
é suficiente. Falha/timeout interrompe o instalador com código2. O algoritmo de
integridade usado na importação continua igual.

O adaptador mantém o template padrão e o processo padrão de assinatura do
desinstalador. `node_modules` não foi editado. A troca de include ocorre apenas
no compilador NSIS; versão e hash normalizado do template original são fixados.
Uma futura atualização de electron-builder exige revalidar esse contrato.

Arquivos de produção/configuração:

- `build/desktop-installer.nsh`, `build/nsis/extractAppPackage.nsh`.
- `build/pdf-association.nsh`: somente chamada condicional da validação antes de
  registrar associações, preservando os comandos existentes.
- `scripts/check-nsis-distribution.mjs` e `electron-builder.config.cjs`.
- `desktop/distribution-validation.mjs` e `desktop/main.mjs`: entrada interna de
  validação do próprio pacote, preservando a inicialização normal síncrona.
- `desktop/libreoffice-runtime.mjs`: somente localização do bundle empacotado;
  verificação criptográfica/ambiente privado mantidos.
- `scripts/check-desktop-package.mjs`: auditoria do novo layout.
- `electron-builder.hold-native-test.cjs`: versão1.0.0-experimental.35, mesma identidade.

PDFWorkspace, motores PDF/imposição/Windows, importador/conversor Word, bytes do
runtime, algoritmo de integridade Office, preferências e atualizadores preservados.

## Validação e artefato

**Setup definitivo gerado e validado em laboratório, NÃO instalado.**

| Verificação final | Resultado |
|---|---|
| Node, concorrência2 |472 aprovados, nenhuma falha/skip |
| Electron smoke repetido após layout curto |Editor, miniaturas, giro, abas, exportação/cancelamento/falha, proteção de fechamento, preferências e cópia seletiva aprovados; nenhuma impressão |
| Auditoria do pacote |467 entradas ASAR, ícones, auxiliares/PDFium e19410 arquivos Office íntegros |
| Entrega pelo macro NSIS real |19510 arquivos, todos com nomes/hashes iguais à origem; nenhuma ausência, raiz172 e runtime175 |
| Validação instalada simulada |Executável empacotado `--verify-distribution` retornou0 após hashes e inicialização privada |
| DOCX e DOC sintéticos |Dois A4 cada;210,01×297mm, fontes presentes,105 caracteres e frases/acentos esperados preservados |
| Isolamento Word |Busca de Office global desativada; originais iguais, cancelamento e runtime ausente recusados; integridade reconferida após uso |
| Manifesto adulterado somente na fixture |Mesmo macro NSIS recusou com código2; restauração somente da fixture e validação posterior aprovadas |
| Runtime íntegro mas sem inicialização, raiz Office199 |Função de validação com app simulado, hashes reais19410 e processo nativo real: recusou com código2, sem relaxar integridade |
| Instalação ativa e perfil |75 arquivos do perfil iguais, registro NSIS igual e cinco arquivos principais ativos iguais; nenhuma instalação realizada |
| Sintaxe/whitespace |JS/MJS/CJS alterados e `git diff --check` conferidos ao encerrar |

As regressões Node incluem geometria de N-up/Step & Repeat/livreto e proteções
do editor. Não houve nova homologação física nem execução de impressão virtual
nesta investigação. A primeira suíte simultânea ao ensaio de cópia teve uma
expectativa antiga do include e um timeout do auxiliar simulado; expectativa
atualizada e repetição integral passou, sem mudar timeout do produto. Depois do
layout curto a suíte inteira foi repetida e passou com472 testes.

Artefato definitivo:

- `dist-desktop/hold-native-test/4PrintShops-Hold-Test-1.0.0-experimental.35-Windows-x64-Setup.exe`.
- 514.961.998 bytes; SHA256 `FC3199052123BB5A64E6CBD3EF058626024A65310BD97CC01113F74E432AB108`.
- Identidade `br.com.4printshops.holdtest`, produto/executável/perfil Hold Test;
  versão1.0.0-experimental.35; mesma chave NSIS e destino previsto da .34.
- Sem certificado comercial, sem publicação; as verificações dos atualizadores
  permanecem intactas. O primeiro Setup de diagnóstico foi preservado como
  `initial-35-setup.exe` nas evidências e não deve ser instalado.

Testes permanentes: `tests/distribution-validation.test.mjs`,
`tests/desktop-distribution-copy.mjs`, ampliação de
`tests/desktop-private-office-package.mjs` para texto/A4 e ajuste da expectativa
em `tests/assisted-publishing.test.mjs`.

Reprodução controlada, após gerar o Setup:

```powershell
node --test --test-concurrency=2 tests/*-*.test.mjs
node tests/desktop-distribution-copy.mjs dist-desktop/hold-native-test/4PrintShops-Hold-Test-1.0.0-experimental.35-Windows-x64-Setup.exe 172
```

O segundo comando entrega o payload real pelo macro NSIS em uma pasta nova de
laboratório, sem instalar/registrar o produto. Compara todos os arquivos por
SHA256, valida Office, converte DOCX/DOC, testa rejeição do manifesto adulterado e
compara o perfil/registro/arquivos principais da instalação ativa antes/depois.
Nunca o confundir com execução do Setup completo sobre a .34. Um ensaio inicial
de validação usou OUTDIR sem preencher INSTDIR; corrigido somente o roteiro. O
primeiro build também precisou preservar o include raiz do desinstalador após
a seleção do adaptador no compilador. Logs/falhas mantidos. O teste não envia
trabalhos para impressoras. Evidências em `.desktop-test/office35-wIwpZu/` e nas
pastas novas `distribution35-*`, fora do Git. Falhas e logs iniciais preservados.

Resultado final completo em `.desktop-test/distribution35-b44PBG/result.json`;
inventário antes do teste em `before.json`, conversões em `private-office.log`.
Logs em `.desktop-test/graph34/`: `build35-compact.log`, `package35-compact.log`,
`delivery35-compact.log`, `office35-node-compact.log`, `smoke35-compact.log` e
`startup-guard35.log`. O comando de reprodução cria uma pasta nova e preserva
os ensaios anteriores. Os caminhos/arquivos de laboratório não são versionados.

Limites: isso não comprova execução do LibreOffice em qualquer profundidade.
O layout curto foi aprovado no comprimento real da estação; o caminho de
estresse excessivo foi recusado. Não foi alterada a política de caminhos longos.
Uma futura configuração Authenticode deve reconferir os hashes pós-build:
electron-builder pode assinar executáveis de extraFiles, modificando os bytes
do fornecedor. A validação não autoriza aceitar esses bytes por exceção. Sem
certificado nesta etapa, os bytes Office permaneceram idênticos aos preparados.
Como na geração local .34, `PRINT_GHOSTSCRIPT=none`; diagnóstico experimental
PS/Hold permanece bloqueado sem Ghostscript, impressão Windows/PDFium inalterada.

## Inventário Git e próxima validação autorizada

Base mantida `main/c85efe2`, staging vazio.18 caminhos:11 modificados e7 novos.
Os três documentos já modificados no início foram preservados; não houve
alterações em node_modules, arquivos dos clientes ou políticas/configurações
do Windows. Pacotes, caches, marcadores, dumps e provas continuam ignorados.

Modificados:

- `build/pdf-association.nsh`.
- `desktop/libreoffice-runtime.mjs`, `desktop/main.mjs`.
- `electron-builder.config.cjs`, `electron-builder.hold-native-test.cjs`.
- `scripts/check-desktop-package.mjs`.
- `tests/assisted-publishing.test.mjs`, `tests/desktop-private-office-package.mjs`.
- `docs/ESTADO_PROJETO.md`, `docs/DISTRIBUICAO_DESKTOP.md`,
  `docs/EXPERIMENTAL34_IMPORTACAO_PROGRESSIVA.md`.

Novos:

- `build/desktop-installer.nsh`, `build/nsis/extractAppPackage.nsh`.
- `desktop/distribution-validation.mjs`.
- `scripts/check-nsis-distribution.mjs`.
- `tests/distribution-validation.test.mjs`, `tests/desktop-distribution-copy.mjs`.
- Este relatório `docs/DISTRIBUICAO_CAMINHOS_EXPERIMENTAL35.md`.

Somente após nova autorização para atualizar a estação:

1. Salvar/exportar trabalhos e fechar normalmente o aplicativo; reconferir
   processos, hash do Setup, perfil/preferências e associações, com backup.
2. Usar o Setup definitivo sobre o diretório Hold Test existente. Não remover
   manualmente a .34 nem mover o perfil. Erro do instalador exige investigação;
   não implica rollback automático nem autoriza reparo manual de arquivos.
3. Confirmar versão .35, único registro NSIS, mesmo diretório/perfil e runtime
   `lo` completo, depois executar a auditoria e o teste privado nesse destino.
4. Pela interface, importar DOCX/DOC sintéticos e conferir as duas páginas A4,
   texto/tabela/acentos, preferências e proteção de edições. Fazer regressão curta
   de N-up/Step & Repeat/livreto sem imprimir. Comparação visual/manual de Word e
   validação em outros Windows continuam necessárias.

A .34 ativa permanece incompleta e Word continua bloqueado. Não instalar a
correção automaticamente. Requer autorização específica, aplicativo fechado,
backup de perfil e nova conferência de preferências/associações após o upgrade.
Homologação real .34→.35, testes em outros Windows e fidelidade universal
fontes/paginação/ICC/RIP continuam pendentes. Sem commit ou push.
