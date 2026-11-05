# Auditoria do checkpoint — experimentais .16 a .20

Data: 05/11/2026. Base `6b8f63a` (`Refina interface compacta e adiciona temas claro e escuro`), branch `main`. Auditoria anterior ao commit; **nenhum commit ou push criado nesta etapa**.

O usuário informou que testou a .20 e que está tudo certo. Aprovação manual registrada para os cenários por ele exercitados, sem ampliar para homologação universal de impressão, cores/RIP ou todos os PDFs.

## Conclusão e plano

Os37 caminhos inicialmente pendentes correspondem ao desenvolvimento .16–.20. Somado este relatório, são **38 candidatos:13 modificados e25 novos**, sem remoções e sem staging preparado. Não foram identificadas alterações de outro trabalho entre os candidatos.

Recomendação: **um checkpoint conjunto**, preservando o estado final aprovado. `main.js`, `pdfs.js`, temas, harness de integração, builder e documentos foram alterados por várias versões. Dividir agora exigiria reconstruir estados intermediários e repetir validações; não é necessário para este checkpoint de segurança.

Mensagem sugerida:

`Consolida exportacao de imagens e montagem de PDFs com compatibilidade de impressao`

Aguardar autorização explícita para registrar somente os38 caminhos abaixo. Antes de criar o commit, conferir novamente HEAD, status, diff e conteúdo do staging. Não usar `git add .`/`git commit -a` nem incluir artefatos ignorados. Manifesto local de caminhos/hashes em `.desktop-test/audit16-20/candidates.json` permite conferir divergências antes do registro.

## Inventário completo

### Exportação de imagens e atalhos — .16 (12)

| Estado | Caminho | Responsabilidade |
|---|---|---|
| M | `desktop/app.mjs` | Canal de imagens validado e limpeza do trabalho |
| M | `desktop/preload.cjs` | Ponte restrita de exportação |
| M | `desktop/security.mjs` | Allowlist do módulo de atalhos |
| M | `electron-builder.config.cjs` | Inclusão do módulo de atalhos no pacote |
| M | `public/js/app/export-dialog.js` | PDF/PNG/JPEG, seleção, DPI e qualidade |
| M | `public/js/app/menus.js` | Rótulos dos atalhos |
| N | `desktop/image-exports.mjs` | Gravação ordenada/exclusiva dos lotes |
| N | `public/js/app/shortcuts.js` | Ctrl+O/S/P/E |
| N | `public/js/tools/pdf-image-export.js` | Renderização sequencial, DPI e cancelamento |
| N | `tests/image-export.test.mjs` | Unidade/segurança de arquivos/atalhos |
| N | `tests/image-export-browser.mjs` | Chrome/instalado e gravação Windows |
| N | `docs/EXPORTAR_IMAGENS_EXPERIMENTAL16.md` | Operação, testes e limites |

### Montagem, estabilidade e compatibilidade — .17 a .20 (18)

| Estado | Caminho | Responsabilidade |
|---|---|---|
| N | `public/js/tools/pdf-imposition-layout.js` | Geometria Step & Repeat/N-up e quatro modos de marcas |
| N | `public/js/tools/pdf-imposition-engine.js` | Geração vetorial e diagnóstico antes da geometria |
| N | `public/js/tools/pdf-imposition-ui.js` | Janela, prévia, preparação, exportação e retorno |
| N | `public/js/tools/pdf-imposition-preview.js` | Cancelamento, isolamento, limites e limpeza |
| N | `public/js/tools/pdf-imposition-compatibility.js` | Diagnóstico/preparação independente e validação |
| N | `public/js/tools/pdf-imposition-compare.js` | Adaptador isolado da comparação existente |
| N | `tests/pdf-imposition.test.mjs` | Geometria, marcas, recursos e workspace pendente |
| N | `tests/pdf-imposition-preview.test.mjs` | Cancelamentos, falhas, retry e resultados antigos |
| N | `tests/pdf-imposition-compatibility.test.mjs` | Tagged/ActualText, links, aparência e bloqueios |
| N | `tests/pdf-imposition-browser.mjs` | Fluxo real Chrome/Electron/instalado |
| N | `tests/helpers/pdf-imposition-modules.mjs` | Adaptação de imports para Node |
| N | `tests/helpers/pdf-imposition-compatibility-browser.mjs` | Fixtures sintéticas/fluxo preparado e virtual |
| N | `tests/helpers/pdf-imposition-pixel-compare.mjs` | Comparação visual alinhada da peça |
| N | `tests/helpers/pdf-imposition-real-browser.mjs` | Teste local opcional por variável, sem dados incorporados |
| N | `docs/MONTAGEM_IMPOSICAO_EXPERIMENTAL17.md` | Montagem vetorial e aprovação geométrica limitada |
| N | `docs/MONTAGEM_IMPOSICAO_EXPERIMENTAL18.md` | Diagnóstico/correção da prévia e cantos do grupo |
| N | `docs/MONTAGEM_IMPOSICAO_EXPERIMENTAL19.md` | Marcas de divisão somente nas bordas |
| N | `docs/MONTAGEM_IMPOSICAO_EXPERIMENTAL20.md` | Preparação independente e testes/limites de fidelidade |

### Compartilhados e registro do checkpoint (8)

| Estado | Caminho | Responsabilidade |
|---|---|---|
| M | `public/js/main.js` | Atalhos e contexto independente da montagem para imprimir |
| M | `public/js/tools/pdfs.js` | Integração com snapshot do workspace, menus/exportação |
| M | `public/css/theme.css` | Exportação, montagem, temas e aviso de compatibilidade |
| M | `electron-builder.hold-native-test.cjs` | Versão .20 na mesma identidade/perfil Hold Test |
| M | `tests/pdf-integration.test.mjs` | Importação dos módulos novos no harness |
| M | `docs/ESTADO_PROJETO.md` | Histórico .16–.20 e aprovação/auditoria atuais |
| M | `docs/DISTRIBUICAO_DESKTOP.md` | Instaladores, identidade e pendências |
| N | `docs/AUDITORIA_EXPERIMENTAIS16_20.md` | Este inventário/escopo exato |

## Segurança e privacidade

- Os candidatos são exclusivamente JS/MJS/CJS, CSS e Markdown. Sem binários, byte nulo, PDF, captura, log, instalador, build, arquivo temporário, certificado, credencial ou dado DPAPI entre eles.
- Inspeção das diferenças e dos novos arquivos, incluindo o helper de teste real. Ele recebe `IMPOSITION_REAL_PATH` em execução; não contém caminho privado, documento de cliente ou base64 literal do real. O conteúdo é lido somente para teste local, com derivados em `.desktop-test/` e conferência do hash original. Fixtures permanentes são sintéticas.
- Verificação dirigida de marcadores de chave privada, formatos comuns de token e blobs base64 literais extensos não apontou candidatos suspeitos. Isso não substitui revisão humana nem é garantia universal contra qualquer segredo; não foram encontrados indícios na revisão deste conjunto.
- `.gitignore` exclui `.desktop-test/`, `dist-desktop/`, builds nativos, PDFs, logs, temporários, chaves, certificados, `.env`/credenciais e DPAPI. Conferido com exemplos reais dos derivados e caminhos representativos desses tipos.
- Atualizador, verificações Authenticode, fuses, sandbox/context isolation, credenciais e configuração GitHub não foram enfraquecidos. Ampliação da allowlist inclui somente o módulo de atalhos; IPC de imagens valida origem, plano e tokens. Instalação principal/associação não receberam identidade nova.
- Originais dos clientes e workspace foram preservados. Sem nova impressão física, alteração de driver, novo instalador ou publicação durante esta auditoria.

### Pendência histórica encontrada: node_modules já rastreado

O Git contém **2.212 caminhos de node_modules já rastreados**, todos presentes no HEAD atual e anteriores a este desenvolvimento. A entrada `.gitignore` não retira arquivos já rastreados. O histórico consultado inclui o commit `3f6b150`.

Nenhum desses caminhos está entre os38 candidatos, nem tem alteração no diff desta tarefa. Não foram removidos do disco, do índice ou do histórico. O checkpoint proposto não acrescentará mudanças de node_modules, mas seu snapshot continuará herdando o conteúdo histórico do commit-base. Recomenda-se uma limpeza separada, autorizada e testada; não misturar essa remoção com a preservação das funcionalidades aprovadas.

## Validações desta auditoria

- `npm test`: **307 aprovados**,0 falhas, com fixtures sintéticas; sem execução de PDF real ou envio físico nesta rodada.
- `node --check`: **29 scripts JS/CJS/MJS** modificados/novos aprovados.
- Pacote .20 instalado auditado novamente: **446 entradas**, fontes/assets correspondentes, exclusões/allowlist aprovadas.
- `git diff --check` aprovado. Staging vazio; sem remoções e sem operação de commit em curso.
- Registros anteriores Chrome/Electron/instalado consultados: atalhos/imagens, Step & Repeat/N-up, quatro modos de marcas, prévia/cancelamentos/retry, origem pendente, diagnóstico/preparação, salvar Windows e impressão exclusivamente virtual. Não foram repetidos ensaios demorados nem instalações na auditoria.
- Aprovação manual da .20 informada pelo usuário registrada. Aprovação física A3→A6→Step & Repeat permanece limitada ao cenário relatado.

## Limitações que permanecem documentadas

- Campos/formulários, aparências de anotações, ActualText essencial/ambíguo, camadas/perfis e outras estruturas não comprovadas continuam bloqueados; sem rasterização silenciosa.
- Preparação do real conservou texto/pixels/dimensões/giro e18 streams. Na comparação da peça escalada, Electron teve máximo1/255; Chrome teve0,169% de componentes acima de2 níveis. Conferência visual/RIP continua necessária; aprovação manual não é certificação universal.
- Tetos de memória/resolução, possíveis resíduos após encerramento forçado, trechos síncronos de PDF-lib e limites de comparação permanecem. N-up misto/bleed e semântica das marcas seguem os respectivos guias.
- Instalação automática/reinício do atualizador, piloto assistido no servidor real/dois PCs, SMB/UNC e verificações funcionais específicas do Icon Test conservam suas pendências anteriores.
- Limpeza do node_modules histórico é trabalho futuro separado.

Próximo passo: usuário analisar este inventário e autorizar o registro local do checkpoint proposto. Não prosseguir com novas funcionalidades antes de concluir esse registro. Push sempre manual pelo usuário.
