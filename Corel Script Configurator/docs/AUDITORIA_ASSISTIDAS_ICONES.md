# Auditoria conjunta — atualizações assistidas e ícones

Data: 03/11/2026. Base: `c450fed` — `Adiciona abertura de PDFs pelo Windows ao 4PrintShops`. Commits anteriores: `90cc0ce` (atualizador integrado/laboratório de assinatura) e `e7b0ae3` (primeiro Desktop).

**Recomendação: dois commits íntegros, na ordem solicitada. Nenhum commit criado.** Staging vazio; não houve push, reset, checkout destrutivo, descarte ou remoção de arquivos existentes. Esta auditoria modificou somente documentação versionável e criou instrumentação/saídas ignoradas. Não implementou novas funcionalidades.

## Inventário final

48 candidatos: **16 modificados e 32 novos**, nenhuma remoção. Classificação: **25 exclusivos de atualizações assistidas**, **16 exclusivos de identidade visual/ícones** e **7 compartilhados**. Não foi encontrada alteração alheia ou acidental entre os candidatos. O documento desta auditoria é compartilhado por assunto, mas será incluído somente no segundo commit.

Na tabela, M = modificado desde HEAD; N = novo. `1 → 2` significa conteúdo intermediário completo no primeiro commit e complemento no segundo, não um arquivo de runtime incompleto.

| Arquivo | Estado | Classificação | Commit |
| --- | --- | --- | --- |
| `.gitignore` | M | Atualizações assistidas | 1 |
| `build/assisted-trust.cjs` | N | Atualizações assistidas | 1 |
| `build/icons/app.ico` | N | Identidade visual/ícones | 2 |
| `build/icons/app.png` | N | Identidade visual/ícones | 2 |
| `build/icons/pdf.ico` | N | Identidade visual/ícones | 2 |
| `build/icons/pdf.png` | N | Identidade visual/ícones | 2 |
| `build/icons/preview.png` | N | Identidade visual/ícones | 2 |
| `build/icons/sources/icone4ps.png` | N | Identidade visual/ícones | 2 |
| `build/icons/sources/iconepdf.png` | N | Identidade visual/ícones | 2 |
| `build/pdf-association.nsh` | M | Identidade visual/ícones | 2 |
| `desktop/app.mjs` | M | Compartilhado | 1 → 2 |
| `desktop/assisted-http.mjs` | N | Atualizações assistidas | 1 |
| `desktop/assisted-manifest.mjs` | N | Atualizações assistidas | 1 |
| `desktop/assisted-ui.mjs` | N | Atualizações assistidas | 1 |
| `desktop/assisted-updates.mjs` | N | Atualizações assistidas | 1 |
| `desktop/preload.cjs` | M | Atualizações assistidas | 1 |
| `desktop/security.mjs` | M | Atualizações assistidas | 1 |
| `docs/ATUALIZACOES_ASSISTIDAS.md` | N | Compartilhado | 1 → 2 |
| `docs/AUDITORIA_ASSISTIDAS_ICONES.md` | N | Compartilhado | 2 |
| `docs/DISTRIBUICAO_DESKTOP.md` | M | Compartilhado | 1 → 2 |
| `docs/ESTADO_PROJETO.md` | M | Compartilhado | 1 → 2 |
| `docs/ICONES_DESKTOP.md` | N | Identidade visual/ícones | 2 |
| `electron-builder.assisted-lab.cjs` | N | Atualizações assistidas | 1 |
| `electron-builder.assisted.cjs` | N | Atualizações assistidas | 1 |
| `electron-builder.config.cjs` | M | Compartilhado | 1 → 2 |
| `electron-builder.icon-test.cjs` | N | Identidade visual/ícones | 2 |
| `package-lock.json` | M | Atualizações assistidas | 1 |
| `package.json` | M | Atualizações assistidas | 1 |
| `public/js/app/desktop-assisted.js` | N | Atualizações assistidas | 1 |
| `public/js/app/menus.js` | M | Atualizações assistidas | 1 |
| `public/js/main.js` | M | Atualizações assistidas | 1 |
| `scripts/assisted/hosting.mjs` | N | Atualizações assistidas | 1 |
| `scripts/assisted/keys.mjs` | N | Atualizações assistidas | 1 |
| `scripts/assisted/prepare-lab.mjs` | N | Atualizações assistidas | 1 |
| `scripts/assisted/publish.mjs` | N | Atualizações assistidas | 1 |
| `scripts/authenticode-lab/prepare-kit.mjs` | M | Compartilhado | 1 → 2 |
| `scripts/check-desktop-package.mjs` | M | Identidade visual/ícones | 2 |
| `scripts/generate-icons.ps1` | N | Identidade visual/ícones | 2 |
| `scripts/icons/IconAssets.cs` | N | Identidade visual/ícones | 2 |
| `tests/assisted-key-windows.mjs` | N | Atualizações assistidas | 1 |
| `tests/assisted-publishing.test.mjs` | N | Atualizações assistidas | 1 |
| `tests/desktop-assisted-installed.mjs` | N | Atualizações assistidas | 1 |
| `tests/desktop-assisted.test.mjs` | N | Atualizações assistidas | 1 |
| `tests/desktop-icons-installed.mjs` | N | Identidade visual/ícones | 2 |
| `tests/desktop-icons.test.mjs` | N | Identidade visual/ícones | 2 |
| `tests/helpers/desktop-entry.mjs` | M | Atualizações assistidas | 1 |
| `tests/helpers/windows-icon-resources.ps1` | N | Identidade visual/ícones | 2 |
| `tests/helpers/windows-save-dialog.ps1` | M | Atualizações assistidas | 1 |

## Plano exato de commits

1. **`Implementa atualizacoes assistidas pela rede local`** — 31 arquivos: 14 modificados e 17 novos. Módulos, scripts, configuração de confiança, testes e documentação assistida completos. Conserva o logo/ícone anterior e o NSIS anterior de associação. Não inclui `build/icons/`, Icon Test, conversor nem teste de ícones.
2. **`Integra icones personalizados do 4PrintShops e PDFs`** — 23 caminhos: 15 arquivos adicionados e oito alterados em relação ao primeiro commit. Inclui assets/fontes, conversor, config de Icon Test, testes/validação de pacote, ícone do ProgID PDF e documentação/aprovação. Atualizador assistido do primeiro commit permanece completo.

Há seis arquivos efetivamente presentes nos dois commits:

| Arquivo | Conteúdo do commit 1 | Complemento do commit 2 |
| --- | --- | --- |
| `desktop/app.mjs` | Serviço assistido, IPC, isolamento do Assisted Lab, proteção e travas por proprietário | Identidade isolada Icon Test e PNG da janela |
| `electron-builder.config.cjs` | Somente acréscimo do frontend assistido à allowlist | PNG no ASAR, ICO do executável/NSIS e recurso externo do PDF |
| `scripts/authenticode-lab/prepare-kit.mjs` | Allowlist do frontend assistido | Três assets necessários para futuros builds do kit |
| `docs/ESTADO_PROJETO.md` | Estado/registro do fluxo assistido | Ícones aprovados, auditoria e pendências conjuntas |
| `docs/DISTRIBUICAO_DESKTOP.md` | Distribuição assistida e limites | Referências aos ícones e relatório conjunto |
| `docs/ATUALIZACOES_ASSISTIDAS.md` | Guia completo com resultados históricos de 199 testes | Adendo de compatibilidade com o novo pacote contendo ícones |

**Prova da separação:** foi criada uma cópia de auditoria de fontes em `.desktop-test/auditoria-conjunta-20261103/primeiro-commit/`, sem alterar o working tree ou staging. Ela restaura o ícone anterior, mantém completos os módulos assistidos e passou nos **199 testes Node**, sem depender dos assets novos. O conjunto atual passou nos **202 testes**. O plano intermediário de 31 arquivos e seus SHA256 está em `plano-primeiro-commit.json`, na mesma pasta ignorada.

Após autorização, preparar o índice com os arquivos inteiros exclusivos e os conteúdos intermediários verificados dos seis compartilhados; criar o primeiro commit; depois preparar os 23 caminhos do segundo. Preservar os arquivos finais do working tree durante o processo. Não usar `git add .` para o primeiro commit nem restaurar versões sobre arquivos finais. Antes de cada commit, conferir status, diff do índice e a correspondência com o plano. Se algum candidato mudar depois desta auditoria, rever os hashes e os testes afetados antes de usar o snapshot.

## Verificações e resultados

| Verificação | Resultado / alcance |
| --- | --- |
| Node da projeção do commit 1 | **199 aprovados**, zero falhas/skips, sem ícones |
| Node do conjunto final | **202 aprovados**, zero falhas/skips; inclui manifestos, replay/downgrade, hash, rede, proteção e ícones |
| Chrome headless curto | Aprovado; web via servidor Node, PDFs/QR, abas 100/350, exportação/proteção e comandos legados |
| Electron real com perfil temporário | Aprovado; sandbox/isolamento, PDF.js/miniaturas, giro, exportação/cancelar/falhar, fechamento, persistência e cópia seletiva |
| DPAPI Windows | Aprovado com chave descartável nova; ida/volta, correspondência pública e rejeição após adulteração; nenhum certificado instalado |
| Icon Test já instalado | Releitura aprovada, sem reinstalação: executável/Setup/desinstalador/PDF em 16/32 px iguais aos ICOs; atalhos e ProgID/comando de abertura corretos |
| Pacote assistido com ícones | Novo NSIS real e ASAR de **420 entradas** auditados; PNG e ICO externo iguais às fontes; ausência de documentos/chaves/testes no pacote |
| Publicação/download desse novo NSIS | Ed25519/DPAPI, produto/canal/raiz incorporada, detecção como `.5` → `.6`, adiar sem baixar, SHA512 e preparo aprovados; manifesto/Setup adulterados recusados, offline tratado |
| Sintaxe / espaços / segurança Git | **28 JS/MJS/CJS**, parsing de **três PowerShell** e `git diff --check` aprovados; novos arquivos incluídos na inspeção sem fazer staging |

A primeira tentativa Electron atingiu timeout na prévia seletiva enquanto o build estava em andamento. A repetição isolada passou sem mudança no código; causa não estabelecida conclusivamente. Log inicial preservado. O erro de gravação em pasta inexistente mostrado no log pertence ao cenário de falha intencional do próprio teste.

A conferência de ícones requer **Windows PowerShell 5.1/System.Drawing**; a tentativa com host PowerShell 7 não encontrou a referência .NET necessária. Usar o host indicado, sem mudar políticas de execução ou instalar dependências. Nenhum desses dois eventos motivou alteração nos motores PDF.

O pacote adicional está em `dist-desktop/auditoria-conjunta/assisted-com-icones/4PrintShops-Assisted-Lab-1.0.0-experimental.6-Windows-x64-Setup.exe`, SHA256 `279b627a146f06fee24bb7d3e2309679edf3e37d286d672f3c5bdb20b68c059e`. É somente material de auditoria, **não uma nova release**. Instaladores históricos e feed anterior foram preservados. Cliente/cache/repositório novos foram usados no teste; perfis/estado antirreplay instalados não foram alterados.

**Limite dessa prova adicional:** diálogos e localização no Explorador foram simulados, usando o cliente/publicador reais e o instalador real. O pacote novo não foi instalado nem executado. Não se repetiu a homologação de UI instalada entre duas versões; a validação instalada da etapa assistida anterior permanece registrada como histórica. Os ícones têm aprovação visual pessoal do usuário, distinta dessa prova criptográfica.

Logs novos em `%TEMP%/4printshops-auditoria-*.log`; artefatos Chrome `%TEMP%/4printshops-browser-r3eJYr`, Electron `%TEMP%/4printshops-desktop-7126Hr`; recursos relidos, plano e relatório criptográfico em `.desktop-test/auditoria-conjunta-20261103/`. Esses caminhos não pertencem ao commit.

## Segurança do que será versionado

- Lista revisada arquivo a arquivo; não contém instaladores, builds, PDFs, logs, arquivos temporários, certificados, chaves privadas ou dados DPAPI. PNGs/ICOs são os sete assets definitivos, incluindo cópias preservadas das artes fornecidas e prancha de qualidade.
- `.gitignore` cobre instaladores/ASAR/blockmaps, `dist-desktop/`, `.desktop-test/`, PDFs, logs, temporários, `.part`, `.dpapi`, certificados/chaves, `.env`/senhas e repositório de versões gerado. Também foram conferidos arquivos reais gerados; a proteção não depende somente de extensões hipotéticas.
- Chave de publicação de laboratório permanece fora da raiz Git e do repositório de distribuição. O builder lê somente `public-trust.json`; esse arquivo local não é candidato. Código que gera, protege e utiliza a chave deve ser versionado; o material privado não.
- Não houve mudança direta em PDFWorkspace, proteção/edição PDF, QR Codes, `desktop/updates.mjs` ou `desktop/update-signature.mjs` desde HEAD. Authenticode do atualizador integrado permanece obrigatório, sem fallback de falha para execução assistida.
- Associação continua apenas no ProgID/Open With próprios: `DefaultIcon` usa o recurso instalado e `--open-pdf` permanece igual. Não grava UserChoice nem altera o leitor padrão; a conferência funcional completa de Abrir com no Icon Test continua pendente.

## Riscos, pendências e decisão

1. **Arquivos compartilhados:** incluir seu conteúdo final inteiro no primeiro commit anteciparia referências aos ícones sem os assets. Usar a projeção testada e fazer a atualização completa no segundo evita esse problema. Não é necessário um checkpoint único.
2. **Versão `.6` reempacotada:** os bytes com ícones diferem dos históricos. Não sobrescrever uma release existente nem reutilizar esse pacote de auditoria na distribuição principal. Publicação real posterior deve avançar versão/seq, com a identidade e a chave principais aprovadas.
3. **Homologações distintas:** `.6` preparada, não executada pelo mecanismo assistido. Instalação/reinício automáticos continuam **não homologados**, pendentes de VM/certificado confiável. Aprovação visual não homologa instalação ou fluxo Abrir com.
4. **Rede interna:** ainda faltam aprovação de pasta/contas/ACL/serviço/porta no servidor real, base principal confiável, guarda/recuperação da chave, piloto em dois PCs, execução manual do Setup oferecido e preservação de preferências entre versões. Nenhum serviço, porta, permissão ou certificado do servidor foi alterado.
5. **Limitações existentes:** SMB/UNC com FQDN/IPv4 e teste real de compartilhamento pendentes; sem rotação/revogação remota, rollback/downgrade automático, retomada parcial ou limpeza automática do cache. Preparar não fecha documentos nem executa Setup; execução externa exige exportar/fechar conscientemente.

**Aguardar autorização explícita do usuário após esta auditoria.** Mensagens propostas são exatamente as duas solicitadas; nenhum commit ou git push foi executado.
