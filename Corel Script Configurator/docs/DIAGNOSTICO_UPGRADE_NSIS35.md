# Atualização .34 → .35 — diagnóstico NSIS e recuperação proposta

09/11/2026. Base main/c85efe2. Investigação autorizada, sem interagir com o Setup
ativo PID12488, sem executar instalador/desinstalador real novamente e sem mudar
políticas, permissões, registro, impressoras, documentos ou perfil da estação.
Nenhum pacote experimental novo foi gerado. O Setup .35 mantém seu SHA256
FC3199052123BB5A64E6CBD3EF058626024A65310BD97CC01113F74E432AB108.

## Conclusão e alcance da prova

A rotina de remoção antiga `un.atomicRMDir` do NSIS/electron-builder26.18.0 não
consegue renomear determinados caminhos longos do LibreOffice. Isso foi
reproduzido com o código original e os mesmos nomes públicos de18302 arquivos
Office presentes, mais dois marcadores de aplicação. O arquivo de falha do
ensaio possui260 caracteres e existe/está legível por APIs com caminho estendido.
Não havia aplicativo aberto ou lock nesse cenário. O desinstalador aborta e o
instalador apresenta uma mensagem genérica de aplicativo que não pode ser fechado.

A prova é mais específica que o código2: o retorno aponta o caminho da falha;
`MoveFileW` com caminho comum retorna falso/erro3 (ERROR_PATH_NOT_FOUND), enquanto
o mesmo movimento com `\\?\` tem sucesso na fixture. Um lock real separado
retorna erro32 (ERROR_SHARING_VIOLATION) e continua impedindo a operação corrigida.
Não é uma justificativa para ignorar bloqueios ou tratar todos os arquivos
maiores que260 da mesma maneira.

**Limite da afirmação:** não instrumentamos nem executamos o desinstalador real
da .34 nesta investigação. Portanto não há log com o primeiro arquivo exato da
execução real. A etapa e o defeito da rotina foram comprovados por reprodução e
leitura do mesmo recurso real; o fluxo completo do produto ainda não foi
homologado. Não afirmar que houve uma atualização real bem-sucedida em laboratório.

## Fluxo real analisado

1. `installSection.nsh` verifica processos e chama `uninstallOldVersion`.
2. `include/installUtil.nsh` obtém o UninstallString/InstallLocation, copia o
   desinstalador para `$PLUGINSDIR/old-uninstaller.exe` e usa:

   ```text
   "<old-uninstaller.exe>" /S /KEEP_APP_DATA /currentuser --updated _?=<diretório existente>
   ```

3. A inicialização `initMultiUser/setInstallModePerUser` relê InstallLocation
   do registro e redefine INSTDIR; o caminho de `_?=` sozinho não é autoridade
   final. Com `--updated`, o desinstalador .34 executa `customUnInstall`, depois
   tenta mover a árvore existente para `$PLUGINSDIR/old-install`, por `Rename`.
4. `un.atomicRMDir` retorna o caminho que não pôde mover. A rotina tenta devolver
   os arquivos já movidos com `un.restoreFiles` e faz Abort. Esse rollback não
   contempla o registro de associação.
5. O instalador tenta novamente internamente; após cinco retornos sem sucesso
   reaproveita `appCannotBeClosed`. Essa mensagem também serve para outros casos:
   o texto sozinho não comprova que exista processo aberto.
6. A entrega da nova árvore e a validação do LibreOffice da .35 ainda não foram
   alcançadas. Continuar apesar desse resultado não é uma correção.

O binário temporário copiado confere com o desinstalador .34 registrado, SHA256:
`D45AD680FF936BCD0BDF40B9B9E96F79D01938B8FA5092ECA3AF5FE9C9B8837B`.
O teste neutro validou os argumentos com espaços e confirmou que `_?=` escolhe
a pasta correta e conserva `/KEEP_APP_DATA /currentuser --updated`.

### Efeito no Abrir com

`build/pdf-association.nsh/customUnInstall` remove ProgID, Applications,
Capabilities, RegisteredApplications e o próprio valor OpenWithProgids antes
de saber se os arquivos serão removidos com sucesso. Essa ordem explica por
que os arquivos voltaram, mas as entradas próprias do Hold Test desapareceram.
O UserChoice do leitor PDF padrão permaneceu igual. Não existe snapshot entre
as duas tentativas reais que permita atribuir a remoção só à segunda.

### Relação com as alterações .35

O novo layout `lo`, Robocopy e verificação integral/inicialização privada são
posteriores à etapa de remoção antiga. Eles não corrigem retroativamente o
desinstalador .34 já compilado. A troca para extraFiles não é a origem da falha
reproduzida: ela aconteceu em fixtures contendo somente a árvore antiga.
Não foram encontradas evidências de processo do aplicativo ou conversor aberto;
a reprodução independente não depende de antivírus, política de execução ou
ACL especiais. Não se alterou nenhuma dessas configurações para obter sucesso.

## Evidências de caminhos e integridade da estação

Leitura Win32, sem Rename/Delete na instalação ativa, do recurso público:

```text
resources/libreoffice/help/media/icon-themes/chart2/res/valueaxisdirectpoints_52x60.svg
```

| Acesso | Resultado |
|---|---|
| Caminho comum completo260 |GetFileAttributesW inválido, erro3 |
| Mesmo arquivo com prefixo estendido |Arquivo encontrado, atributo32 |
| Mesmo arquivo pelo alias curto existente |Arquivo encontrado, atributo32 |
| Canonicalização do alias |Mesma pasta física, sem mover ou renomear a instalação |

No diretório ativo, o alias existente tem98 caracteres, versus172 do nome longo.
Com esse alias, até os caminhos potenciais do manifesto Office completo ficam
em220 caracteres. A consulta foi somente leitura; nenhuma criação de nome8.3
ou mudança de política do volume foi feita. GetLastError só foi interpretado
quando a API falhou; valores residuais após sucesso não são erros adicionais.

Auditoria completa por SHA256 contra o payload .34 anterior:

- 19510 arquivos esperados;18402 presentes, todos iguais por hash.
- 1108 ausentes, exatamente o conjunto original; zero arquivos alterados.
- 75 arquivos do perfil iguais ao backup anterior,8.451.685 bytes.
- Desinstalador registrado igual à cópia temporária; versão continua .34.
- O registro de associação está parcialmente removido; não declarar a instalação
  integralmente restaurada. Word segue bloqueado e o novo `lo` está ausente.

Hashes de referência para uma recuperação autorizada:

| Arquivo | SHA256 |
|---|---|
| 4PrintShops Hold Test.exe |48860071613d06b1d15d5afa84508ce8b68d0c4a92d16f197d3a91d71ef0fe64 |
| resources/app.asar |4d6392090d6723df94365c575bb07bb67c6a8ff6bfce12b390dd0813c022e165 |
| Desinstalador .34 |d45ad680ff936bcd0bdf40b9b9e96f79d01938b8fa5092eca3af5fe9c9b8837b |

## Reprodução isolada e testes

Não foi executado o Setup comercial/experimental real em outro diretório:
mesmo com `/D` diferente, ele compartilha GUID, registro, atalhos e mutex do
produto, podendo atingir a instalação ativa ou sua janela. Não há VM descartável
configurada nesta estação. Em vez disso usamos executáveis NSIS neutros, sem
qualquer escrita de registro, atalho, perfil ou chamada de impressão.

Ensaio amplo com18304 marcadores de nomes públicos, conteúdo conferido:

| Caso | Resultado |
|---|---|
| Rotina antiga, raiz67 |Sucesso;18304 arquivos transferidos |
| Rotina antiga, raiz172 |Falha no recurso de260 caracteres;18304 marcadores restaurados |
| Mesma rotina com caminhos estendidos, raiz172 |Sucesso;18304 transferidos, conteúdos iguais |

Teste permanente opt-in `tests/desktop-nsis-upgrade-lab.mjs`, seis casos:

| Caso | Resultado |
|---|---|
| Desinstalador neutro original, caminho curto |Código0, arquivos preservados na quarentena |
| Original, caminho longo |Código2, Rename falha, MoveFileW comum erro3; versão estendida move e restaura a mesma fixture |
| Original chamado pelo alias curto da mesma pasta longa |Código0; mesmos argumentos/modo/pasta física |
| Algoritmo com prefixo estendido e limpeza de metadados adiada |Código0 |
| Arquivo realmente aberto com FileShare.None |Código2 e erro32, arquivos restaurados e metadados-simulados preservados |
| Alias seguido da reescolha do caminho longo |Código2 novamente; demonstração do efeito de redefinir INSTDIR como faz a inicialização multiusuário |

O último caso simula explicitamente a reescolha com StrCpy, baseado na leitura
do template multiUser; não escreve ou lê chaves reais. Ele evita concluir que
alterar apenas `_?=` resolveria o fluxo completo. O teste de metadados usa um marcador de arquivo no laboratório, não o registro
real. Ele demonstra a ordem transacional, não homologa restauração de chaves
Windows em produção. Os testes de prefixo são para caminhos locais de unidade;
UNC precisaria de normalização própria e testes adicionais. A variante com alias
requer um alias existente: não permite ativar nomes8.3 globalmente.

Os14 testes Node focados de distribuição/runtime também passaram. Sintaxe do
novo teste e git diff --check conferidos; não repetida a suíte472 anterior,
Chrome/Electron ou impressão porque nenhum motor de produção foi modificado.
Pequenos erros iniciais da instrumentação foram corrigidos somente nos ensaios
(registradores System e interpolação do template); logs preservados.

Reprodução, sem instalar o produto:

```powershell
node tests/desktop-nsis-upgrade-lab.mjs
node --test tests/distribution-validation.test.mjs tests/desktop-libreoffice-runtime.test.mjs
```

O teste fixa versão26.18.0 e hash do template de remoção. Em outra versão do
builder, revalidar o componente antes de usar o protótipo.

## Correção recomendada, ainda não integrada à distribuição

1. Para migrar a .34 já instalada, preparar e auditar um **desinstalador de
   compatibilidade temporário, embutido no pacote de migração**, com a rotina de
   caminhos corrigida e os mesmos GUID, produto, currentuser e flags de
   preservação. Não substituir manualmente o desinstalador registrado. Seleção
   estrita da versão/identidade/hash antigos e da mesma pasta; parent mantém
   verificação de processo e recusa de retorno não zero. Validar essa integração
   completa em VM antes de liberar outra tentativa na gráfica. Esse componente
   de migração ainda não foi implementado ou homologado.
2. No componente de compatibilidade e nos desinstaladores futuros, tratar os caminhos longos em atomicRMDir E
   restoreFiles, com normalização adequada e rollback que continue recusando
   arquivos realmente bloqueados. Não basta corrigir só a cópia de instalação.
3. Mover a retirada das associações próprias para o sucesso da desinstalação;
   guardar previamente os valores próprios e prever restauração controlada
   em caso de falha da atualização. Não tocar no UserChoice,
   default de .pdf ou entradas de outros leitores. O marcador do laboratório
   não substitui testes reais dessa integração com registro isolado/VM.
4. Preservar integralmente a verificação do LibreOffice e seu layout curto .35;
   recusar retorno não zero, inconsistência de caminho, arquivos extras ou falha
   do runtime. Não substituir o desinstalador instalado manualmente nem ignorar
   a falha para sobrepor arquivos.

**Alternativa descartada como solução pronta:** passar só um alias curto em
`_?=` passou no componente neutro, mas não no caso que reescolhe o caminho longo.
O desinstalador real relê InstallLocation; alterar apenas esse argumento pode não
ter efeito e ainda muda a interpretação das verificações baseadas em caminho.
Não alterar registro temporariamente, usar argumentos malformados ou remover
checagens para forçar essa alternativa. O alias real foi consultado somente como
evidência de acesso à mesma pasta, não aplicado como correção na estação.

Ainda não foi feita uma modificação de produção desses hooks. É necessária
uma alteração de empacotamento para mudar a chamada que hoje está compilada no
Setup; executar o mesmo Setup .35 de novo não aplica este protótipo. Nenhum novo
instalador ou número experimental foi gerado para esta investigação.

## Procedimento de recuperação proposto — NÃO EXECUTADO

Antes de qualquer nova instalação, revisar e autorizar separadamente:

1. Registrar a caixa atual e cancelar a tentativa pela ação Cancelar do aviso,
   sem Repetir. O template usa MB_RETRYCANCEL. Se o rótulo exibido for realmente
   “Continuar”, não adivinhar sua ação: usar a saída de cancelamento/X/Escape
   somente após confirmar a janela; nunca confirmar continuação da instalação.
   Se surgir outro aviso, registrar antes de aceitar apenas o encerramento.
   Se o cancelamento normal não estiver disponível, parar e não forçar processos.
2. Aguardar o término normal do Setup/controlador e guardar visible-result.json.
   Não executar aplicativos ou outro instalador enquanto houver operação pendente.
3. Fazer nova cópia do perfil e validar os75 hashes; conferir versão .34,
   executável/ASAR/desinstalador pelos hashes acima, UninstallString/currentuser,
   mesmo diretório e inventário completo. O estado esperado atual é18402 arquivos
   corretos e os1108 já ausentes, não um LibreOffice funcional.
4. Comparar todas as entradas próprias de Abrir com ao snapshot anterior à
   primeira tentativa. Após autorização específica, restaurar somente os valores
   do ProgID Hold Test, Applications do seu executável, Capabilities,
   RegisteredApplications e o próprio OpenWithProgids. Validar os caminhos dos
   binários antes. Não importar indiscriminadamente a árvore .pdf ou o snapshot
   completo; jamais restaurar/escrever UserChoice ou seu Hash.
5. Reconferir associações e leitor padrão, preferências, arquivos e única
   identidade instalada. Só então decidir a instalação de um pacote corrigido
   e auditado, com novo backup e teste de integridade/conversão após instalar.

Não autorizada recuperação automática nesta tarefa. A caixa continua aberta,
nenhum arquivo real foi movido e não se tentou reinstalação. A associação ainda
não está recuperada, e Word continua indisponível na .34.

## Evidências e estado do Git

- `.desktop-test/nsis-upgrade-proof.log`; evidência ampla em
  `%TEMP%/4ps-upgrade-proof-8tFrO2/results.json`.
- `.desktop-test/nsis-upgrade-lab-complete.log`; seis casos permanentes em
  `%TEMP%/4ps-nsis-upgrade-7KdeIC/results.json`.
- `.desktop-test/active-path-read-result.json`, prova somente de leitura do
  recurso ativo e canonicalização do alias.
- `.desktop-test/active-upgrade-integrity.json`, hashes/conjunto exato de ausentes.
- `.desktop-test/nsis-upgrade-focused.log`,14 testes Node.
- Primeiras tentativas, backups e logs originais preservados.

Os14 arquivos anteriores fora de docs foram comparados ao inventário pré-instalação
e permanecem byte a byte iguais. Adicionados somente este relatório e o teste
permanente; atualizados os registros de contexto/distribuição.20 caminhos pendentes,
staging vazio; sem commit, push, limpeza ou nova versão.

Referências primárias: [Microsoft, caminhos e namespaces](https://learn.microsoft.com/en-us/windows/win32/fileio/naming-a-file),
[GetShortPathNameW e ausência possível de aliases](https://learn.microsoft.com/en-us/windows/win32/api/fileapi/nf-fileapi-getshortpathnamew),
[NSIS, instruções/retornos e parâmetros](https://nsis.sourceforge.io/Docs/Chapter4.html).
Aliases não são garantidos pelo Windows; sua disponibilidade deve ser verificada,
sem modificar o volume. A prova nesta estação não homologa todos os Windows.
