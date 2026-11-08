# Manutenção de armazenamento — somente simulação

## Conferência durante o checkpoint .33 — 08/11/2026

Nova execução de `npm run cleanup` (sem argumentos): saída0,5,170GB potenciais,
fotografia da árvore antes/depois sem diferenças. Evidência em
`.desktop-test/checkpoint33-1794162139516/cleanup-proof.json` e `cleanup.log`.
Backup real do perfil da instalação .32 preservado em `before/profile-backup`
nessa mesma pasta, fora da allowlist. Nenhuma operação destrutiva acrescentada.
Nova suíte completa458/458 aprovada, incluindo17 testes da ferramenta.
ESTADO_PROJETO recebe agora o registro autorizado da instalação/avaliação;
a preservação byte a byte dos13 caminhos descrita abaixo refere-se à etapa
anterior, antes dessas adições documentais autorizadas.

## Escopo e continuidade — 08/11/2026

Ferramenta de desenvolvimento separada do aplicativo e do instalador. Não altera
motores, builder, versão, preferências ou instalação. Os 13 caminhos pendentes da
experimental.33 devem permanecer byte a byte iguais. Por isso este registro é
separado de ESTADO_PROJETO.md, que já pertence àqueles 13 caminhos.

Não há exclusão, movimento, limpeza Git ou modo de aplicação implementado.
Nenhum manifesto é executável nem pode ser fornecido como instrução de limpeza.
Uma etapa destrutiva futura exigiria nova autorização, projeto e revisão.

## Uso

No diretório da aplicação (Corel Script Configurator):

```powershell
npm run cleanup
npm run --silent cleanup -- --json
node --test tests/storage-cleanup.test.mjs
```

O primeiro comando mostra total lógico do repositório, potencial por grupo,
caminhos, razões, alertas e proteções. O segundo emite o manifesto completo em
stdout: grupos, candidatos individuais com tamanho, SHA256 e assinatura de stat,
arquivos protegidos com motivos, Setups preservados com SHA256, raiz, HEAD e
estado Git. Não cria relatório nem cache automaticamente. Para guardar stdout,
escolher um arquivo NOVO ignorado em .desktop-test/storage-cleanup; nunca
redirecionar sobre arquivo existente. A geração de um relatório é uma gravação
intencional do operador, separada da inspeção somente leitura.

Parâmetros desconhecidos, --apply, --delete, --force e caminhos livres falham
antes da varredura. Não há dependências novas. package-lock não precisa mudar:
somente um script npm foi acrescentado.

## Lista explícita de candidatos

Extrações de payload: home24/installer-payload, nup25/installer-payload,
nup26/installer-payload, word30/setup-payload, import31/setup-payload e
release33/payload, todas abaixo de .desktop-test. Cada uma está vinculada ao
Setup da versão correspondente (24/25/26/30/31/33) em dist-desktop/hold-native-test.
Exigem executável com cabeçalho MZ, resources/app.asar e Setup original regular,
legível, com cabeçalho MZ e hash estável. O hash é registrado; isso não equivale
a provar sozinho que a extração corresponde ao Setup. A auditoria do pacote
continua necessária antes de qualquer autorização futura de exclusão.

Extrações NSIS: word30/setup-extracted, import31/setup-extracted e
release33/setup-extracted. Exigem também $PLUGINSDIR/app-64.7z e o Setup associado.

Perfis: somente os pais de execução enumerados em profileRuns no utilitário,
revisados a partir da auditoria (autoscroll, ActualText, checkpoints e testes das
versões conhecidas). Um filho chamado profile ou profile-NÚMERO não basta:
exige estrutura Chrome (Local State e Default) ou Electron (DevToolsActivePort,
Cache e Local Storage). Pais desconhecidos e profile-backup não são selecionados.

Mesmo nesses perfis, SOMENTE os caches enumerados em profileCaches entram na
análise: Cache, Code Cache, GPUCache, GPUPersistentCache, GrShaderCache,
ShaderCache, DawnGraphiteCache, DawnWebGPUCache, optimization_guide_model_store,
component_crx_cache e extensions_crx_cache. Isso inclui caches sob Default.
Preferências, histórico, cookies e armazenamento local ficam protegidos.

Novos caminhos/versões precisam de revisão explícita da finalidade e testes;
não há descoberta por idade nem varredura genérica de qualquer pasta profile.

## Proteções e limites

- Tudo fora da lista permanece protegido, incluindo código, testes, docs,
  configurações, AGENTS, node_modules, .git, runtime/cache do LibreOffice,
  Setups .32/.33 e win-unpacked atual.
- Arquivos rastreados são protegidos, com nova conferência ao final. Git é
  consultado com --no-optional-locks para não atualizar o índice na inspeção.
- Dentro dos grupos, documentos/imagens, JSON/XML/configurações textuais,
  backups, patches, manifestos, hashes, logs, relatórios, provas, avisos de
  licença e arquivos ligados por hardlink permanecem protegidos.
- Um grupo NÃO autoriza selecionar seu diretório inteiro: somente os arquivos
  individuais listados em candidates foram classificados. Muitos grupos contêm
  arquivos protegidos que devem permanecer no lugar.
- Os caminhos são limitados ao repositório, verificados por lstat e realpath,
  inclusive os ancestrais. Links/junctions não são seguidos. Se houver um no
  grupo candidato, o grupo inteiro é recusado. Tamanho total não conta o alvo
  de links; falhas de leitura são registradas e tornam o total parcial.
- Arquivo ilegível/bloqueado, mudança de conteúdo/metadados ou nova entrada
  durante a inspeção recusa o grupo inteiro. Setups são reconferidos ao final.
- O manifesto é uma fotografia, não uma garantia futura contra corridas. Não
  detecta universalmente todos os processos ativos nem prova que um perfil está
  livre. Um futuro executor precisaria validar novamente tudo imediatamente
  antes de agir; não pode confiar num manifesto antigo.
- Extensões e nomes não identificam universalmente dados de clientes. A lista
  fica limitada às extrações e caches sintéticos de finalidade conhecida.
  Qualquer dúvida requer preservar o arquivo e rever a classificação.

## Resultado da simulação real inicial

Total lógico antes de guardar o manifesto: 12.455.494.433 bytes. Foram ignoradas
duas junctions criadas nas fixtures de segurança, sem percorrer seus alvos.
Três perfis vazios/não reconhecidos foram excluídos. Comparação da árvore antes
e depois da execução: nenhuma criação, remoção ou alteração pela inspeção.

| Grupo | Arquivos candidatos | Bytes potenciais |
|---|---:|---:|
| Payloads conhecidos | 12.703 | 3.502.450.420 |
| Extrações NSIS | 33 | 769.111.263 |
| Caches de perfis sintéticos | 3.367 | 898.375.357 |
| Total | 16.103 | 5.169.937.040 |

27.910 arquivos dentro dos grupos ficaram protegidos. A estimativa anterior de
5,669 GB incluía árvores completas; o utilitário reduz para 5,170 GB porque
preserva documentos, manifestos, evidências e dados de perfil. Nenhum espaço foi
liberado. Relatórios e fixtures acrescentados durante esta tarefa explicam
pequenas diferenças de total em inspeções posteriores.

## Testes e retenção

Validação desta implementação: 17 testes específicos aprovados; suíte completa
`npm test` com 458 aprovados, zero falhas/cancelamentos/skips. Sintaxe dos dois
MJS e `git diff --check` conferidos. Não houve teste de interface, novo build,
instalação ou impressão: esta ferramenta não participa desses fluxos.

Os testes usam repositórios sintéticos dentro de .desktop-test/storage-cleanup-tests.
Não criam commits nem alteram o índice do repositório principal. Exercitam
allowlist, proteções, rastreados, backups, junctions externas/ancestrais,
hardlinks, inexistentes, EACCES simulado, bloqueio real FileShare.None no Windows,
novos arquivos e alterações durante inspeção, mudança de Setup e inclusão no
índice durante a varredura. Fotografias com SHA256 conferem preservação das fixtures.

Nesta etapa sem exclusões, as pequenas fixtures são deliberadamente retidas;
o relatório detalhado também é histórico protegido. Não incluí-los em candidatos
automaticamente. Após autorização futura, definir retenção dessas evidências;
o comando padrão não grava nem multiplica relatórios por conta própria.

Manter no desenvolvimento o runtime/cache do LibreOffice, a versão instalada
para retorno e o candidato atual. Priorizar extrações/caches revisados antes
de estudar redução do runtime. Não alterar licenças, fontes, componentes ou
empacotamento com base apenas no tamanho.
