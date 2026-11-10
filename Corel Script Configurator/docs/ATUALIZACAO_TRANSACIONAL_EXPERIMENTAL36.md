# Atualização transacional — experimental.36

**Retomada em casa concluída — 09/11/2026:** base main/522e9f62ee64056752832cf3fbd6bdd0c131b91d, inicialmente limpa. Evidências ignoradas .35/.36 da gráfica indisponíveis; novos ensaios locais concluídos: 489 Node, ASAR/original-fs, NSIS limpo e migração simulada .34→.36 em raiz 172, rollbacks integrais após cópia/registro/Office/lock com todos os arquivos restaurados e módulos PDF .34 utilizáveis. Corrigidos perfil exclusivo do Upgrade Lab e diagnóstico perdido pelo processo GUI/NSIS, com resultado UTF-16 validado e causa no diário. Teste de integridade usa adulteração real do staging. Setup Hold Test .36 gerado/auditado: 514.970.609 bytes, SHA256 `1A39DAD02CD34C79E26D7439F5DA076D7118B68469F7775CF2C396F231762B02`; 19.510 arquivos extraídos/entregues iguais, Office privado DOCX/DOC aprovado. Instalação ativa 19.512 arquivos/perfil 80/registro/leitor padrão preservados; divergência doméstica pré-existente arquivos .34/registro .26 intocada. Sintaxe/diff aprovados, 16 caminhos pendentes e staging vazio, sem impressão, limpeza, commit ou push. **Não instalado no app ativo, não homologado para produção**: VM/desinstalador original, política da gráfica, fontes/paginação/RIP e interrupção abrupta pendentes. Relatório, arquivos e roteiro: [RETOMADA_EXPERIMENTAL36_CASA.md](RETOMADA_EXPERIMENTAL36_CASA.md). Esta entrada substitui as pendências locais de ensaios/pacote dos registros históricos abaixo.


Data: 09/11/2026. **Em validação; não instalar na estação da gráfica ainda.**

## Recuperação da tentativa anterior

O operador cancelou normalmente o Setup .35. O monitor registrou saída 2 às
17:34:55 UTC; o tempo inclui a espera pela janela. Não foi repetido o Setup nem
executado o desinstalador ativo. Auditoria posterior: 19.510 arquivos esperados,
18.402 presentes iguais por SHA256, 1.108 ausentes exatamente como na primeira
auditoria, zero arquivos alterados e os 75 arquivos do perfil íntegros. A .34
continua registrada. As entradas próprias Hold Test de Abrir com permanecem
ausentes desde as falhas anteriores; UserChoice/leitor padrão não mudou. Não
houve restauração manual. Backups e logs anteriores foram preservados.

## Correção em desenvolvimento

A causa documentada em DIAGNOSTICO_UPGRADE_NSIS35.md é a remoção antiga
un.atomicRMDir, não um aplicativo aberto. O novo instalador não executa essa
rotina da .34. Antes de alterar a instalação, extrai o pacote para uma pasta
temporária curta e executa o próprio aplicativo novo em modo interno isolado:

1. Confere identidade, versão anterior, diretório, processos em execução e
   integridade/inicialização do LibreOffice privado do pacote.
2. Inventaria os arquivos e guarda registro próprio, atalhos e cache anterior.
3. Move apenas arquivos gerenciados para backup persistente ao lado da
   instalação, usando operações Windows de caminhos longos do Node/libuv.
4. Copia e verifica os novos arquivos, incluindo o Office privado em `lo`.
5. Atualiza metadados/atalhos/associação própria e confirma a transação.
6. Em falhas tratadas, restaura arquivos e metadados anteriores. Arquivos não
   inventariados, pastas vazias do usuário e perfil não são descartados.

A migração aceita atualmente instalações por usuário das experimentais .34,
.35 e .36. Não é uma autorização para atualizar instalações de outras
identidades ou versões. O manifesto Office legado é reconhecido por hashes
fixos; recursos desconhecidos são preservados. Diretórios vinculados são
recusados. Diário pendente bloqueia nova tentativa automática.

O novo desinstalador usa inventário seletivo e caminhos estendidos, restaura
movimentos se encontrar bloqueio real e só retira a associação própria depois
de concluir a fase de arquivos. Não escreve UserChoice nem o leitor padrão.
O rollback restaura o estado anterior à tentativa, inclusive se a associação
já estava ausente antes dela; não inventa uma associação histórica.

Os modos internos do instalador usam perfil temporário e processo separado.
Não abrem PDFs, não imprimem e não iniciam o aplicativo normal. As verificações
de integridade e inicialização do Office não foram removidas. Não há mudança
de política de caminhos, segurança, antivírus, permissões ou dependência de
Office/LibreOffice global.

## Código e reprodução

- `desktop/update-transaction.mjs`: inventário, diário, aplicação e rollback.
- `desktop/update-registry.mjs`: leitura/restauração restrita ao registro próprio.
- `desktop/main.mjs`: entrada interna do instalador, separada da abertura normal.
- `scripts/prepare-nsis-upgrade.mjs`: gera adaptadores com hashes dos templates
  electron-builder 26.18.0; não edita node_modules.
- `build/desktop-installer.nsh`, `build/nsis/upgrade-transaction.nsh`,
  `build/nsis/uninstall-owned.nsh`, `build/pdf-association.nsh`: integração NSIS.
- `electron-builder.upgrade-lab.cjs`, `build/nsis/upgrade-lab.nsh`: identidade
  exclusiva, registro em Software/4PrintShops/UpgradeLab, atalhos/cache locais
  ao laboratório. Não registrar como leitor de PDFs real.
- Testes Node: `npm.cmd test`.
- Remoção neutra: `node tests/desktop-owned-uninstall-lab.mjs`.
- Pacote completo de laboratório: gerar com electron-builder e a configuração
  upgrade-lab; depois `node tests/desktop-update-package-lab.mjs clean` ou
  `migrate`, exclusivamente conforme as pré-condições do script. Uma falha
  interrompe o cenário e preserva evidências; não repetir indiscriminadamente.

O cenário migrate usa cópia do payload público .34 e a lista local dos recursos
ausentes, nunca a instalação ativa. Executável e desinstalador anteriores são
marcadores inertes para provar que não são executados. Não equivale a executar
o Setup original em uma VM; essa diferença deve permanecer explícita.

## Evidências já obtidas

- 485 testes Node aprovados, sem falhas/skips, em upgrade36-node-final.log.
- Testes de transação: caminhos longos, arquivos do usuário, pastas vazias,
  instalação limpa, falhas de verificação/metadados, lock Windows real,
  resultado antigo alterado, inventário inválido, junction e diário pendente.
- Registro real em namespace aleatório de laboratório: tipos String,
  ExpandString, Binary, DWord, QWord e MultiString preservados; leitor padrão
  igual. Contêineres vazios do laboratório podem permanecer, sem registro ativo.
- NSIS neutro: raiz 172, inventário 19.004 entradas, remoção 0 em 46.188 ms;
  lock real recusado com saída 2 e restauração; caminho inválido recusado com
  saída 2 antes de mover. Primeira tentativa do teste revelou necessidade de
  normalizar `..` antes do prefixo estendido; corrigido e repetido com sucesso.
- Compilação NSIS completa do laboratório passou mantendo warnings-as-errors.
  Ajustes de hooks/includes foram necessários; builds intermediários não são
  instaladores aprovados. Artefatos e logs ficam ignorados em .desktop-test e
  dist-desktop.

## Pendente antes de entregar a .36

O primeiro Setup completo de laboratório foi recusado com saída2 em821.521ms,
antes de preparar o diário ou mover arquivos antigos. A verificação recusou a
inicialização do Office; essa ocorrência ainda não tem causa específica
confirmada. Sondagens independentes iniciaram o mesmo runtime nas raízes122,
149 e em pasta temporária curta. O bloqueio foi preservado, sem fallback.

Um segundo defeito de integração foi reproduzido independentemente: o `fs`
virtualizado do Electron trata app.asar como diretório, levando o leitor físico
a EISDIR. O módulo transacional agora usa original-fs para inventariar/copiar
bytes, mantendo fs virtualizado apenas para ler os metadados embutidos. Não
desabilita ASAR, fuses ou verificação de integridade. A regressão real Electron
`node scripts/test-update-asar.mjs` passou em leitura física, migração e
rollback. A entrada empacotada corrigida preparou e cancelou uma transação
real de19mil arquivos em pasta temporária, sem alterações de instalação.
Pacote de laboratório está sendo reconstruído com essa correção; o primeiro
Setup e seus logs continuam preservados e não são aprovados para uso.

Instalação limpa e migração integral pelo NSIS real; falha deliberada depois da
cópia e depois do registro; integridade instalada de todos os arquivos Office;
inicialização e conversões DOC/DOCX; pacote Hold Test final e SHA256; auditoria
final da estação ativa e autorização específica para instalá-lo.

Falha de energia/encerramento forçado não é homologada como recuperação
automática: o diário/backup persistem e uma recuperação incompleta exige
suporte, sem apagar backup nem repetir automaticamente. Necessário espaço
para pacote temporário, versão nova e backup; não há limpeza automática de
backups antigos de produção nesta etapa. Outros Windows/antivírus/fontes e
paginação permanecem sujeitos a homologação. Nenhuma impressão foi executada.
