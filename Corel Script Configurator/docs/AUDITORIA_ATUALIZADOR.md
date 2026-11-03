# Auditoria do checkpoint do atualizador — 03/11/2026

Base/HEAD: `e7b0ae35fe93fe6e7e9e7b607ea411630d5c1980` (`e7b0ae3`). Auditados `git status`, diff da base, `git diff --stat` e staging. **Staging vazio. Nenhum commit ou push realizado.** Alterações anteriores preservadas.

**Instalação/reinício automáticos NÃO homologados.** O updater tem implementação, testes unitários e ensaio instalado de detecção/download/rejeição de pacotes sem assinatura; a instalação `.3` anterior foi complementar e manual. O teste completo assinado e seus negativos com certificados reais permanecem para uma etapa futura. O usuário decidiu não configurar VM agora.

## Conteúdo proposto para o commit

**34 arquivos: 18 modificados (M), 16 novos (A).** Os novos ainda não rastreados não aparecem em `git diff --stat`; foram inventariados separadamente. Não há remoções nem alterações identificadas fora do escopo. Os módulos PDF/QR não foram reescritos: `pdfs.js` apenas expõe a leitura de estado pendente/ocupação para a consulta de todas as abas. O manifest e lock apenas passam de `.1` para `.3`, sem novas dependências neste checkpoint.

### Configuração, versão e exclusões — 7

| Estado | Arquivo | Finalidade |
| --- | --- | --- |
| M | `.gitignore` | Excluir builds, certificados, credenciais locais, PDFs e temporários. |
| M | `package.json` | Versão experimental.3. |
| M | `package-lock.json` | Mesma versão na raiz do lock; dependências preservadas. |
| M | `electron-builder.config.cjs` | Incluir o adaptador de avisos no frontend empacotado. |
| A | `electron-builder.updates.cjs` | Build com provedor HTTPS/editor e assinatura obrigatória. |
| A | `electron-builder.update-lab.cjs` | Lab unsigned isolado .2/.3, sem distribuição comercial. |
| A | `electron-builder.authenticode-lab.cjs` | Build assinado exclusivo do guest, com barreiras e verificação ativa. |

### Runtime Desktop — 7

| Estado | Arquivo | Finalidade |
| --- | --- | --- |
| M | `desktop/app.mjs` | Consulta inicial, estado/IPC e bloqueio ao instalar; identidade/perfil separados do Lab. |
| M | `desktop/preload.cjs` | Ponte restrita de atualização, consulta de trabalho e remoção de listeners. |
| M | `desktop/security.mjs` | Permitir o novo módulo de interface na allowlist local. |
| M | `desktop/updates.mjs` | Estados, consentimento, validações e configuração do updater. |
| M | `desktop/update-ui.mjs` | Fluxo discreto/manual, confirmação e tratamento de falhas. |
| A | `desktop/work-state.mjs` | Consulta nova por token/frame, timeout e falha fechada. |
| A | `desktop/update-signature.mjs` | Prova adicional Authenticode que bloqueia resultados inconclusivos. |

### Integração com a interface — 4

| Estado | Arquivo | Finalidade |
| --- | --- | --- |
| M | `public/js/main.js` | Conectar avisos e estado de todas as abas. |
| M | `public/css/style.css` | Aviso compacto e progresso, só no Desktop. |
| M | `public/js/tools/pdfs.js` | Getter de estado do workspace/modal, sem mudar edição/exportação. |
| A | `public/js/app/desktop-updates.js` | Avisos, adiamento, progresso e agregação das abas. |

### Scripts reproduzíveis — 5

| Estado | Arquivo | Finalidade |
| --- | --- | --- |
| A | `scripts/serve-update-lab.mjs` | Servidor em loopback, allowlist, Range e banda limitada nos testes. |
| A | `scripts/authenticode-lab/guest-guard.cjs` | Bloquear host e exigir guest/manifesto/consentimento. |
| A | `scripts/authenticode-lab/prepare-kit.mjs` | Criar kit selecionado e hashes fora do Git. |
| A | `scripts/authenticode-lab/create-identity.mjs` | Planejar criação/trust de identidade não exportável só no guest. |
| A | `scripts/authenticode-lab/prepare-feed.mjs` | Preparar feeds válido, editor divergente e adulterações sem mudar origens. |

Esses scripts são fontes necessárias para reproduzir o laboratório; **o kit gerado não entra no commit**. A geração/trust de certificados e os builds assinados continuam não executados.

### Testes e instrumentação — 7

| Estado | Arquivo | Finalidade |
| --- | --- | --- |
| M | `tests/desktop-updates.test.mjs` | Estados, consentimento, pendências, falhas, origem e desbloqueio. |
| M | `tests/desktop-smoke.mjs` | Regressões Electron e PDF pendente em aba oculta. |
| M | `tests/helpers/desktop-entry.mjs` | Instrumentação exclusiva dos testes, não empacotada. |
| M | `tests/helpers/windows-save-dialog.ps1` | Escolhas dos diálogos reais restritas ao processo de teste. |
| A | `tests/desktop-update-signature.test.mjs` | Validação de editor, assinatura, PowerShell e falhas. |
| A | `tests/desktop-authenticode-lab.test.mjs` | Barreiras do guest/host sem criar certificados. |
| A | `tests/desktop-update-real.mjs` | Ensaio instalado unsigned/negativos e instalação manual complementar; não é harness positivo assinado. |

### Documentação — 4

| Estado | Arquivo | Finalidade |
| --- | --- | --- |
| M | `docs/ESTADO_PROJETO.md` | Histórico, checkpoint, pendências e próxima funcionalidade. |
| M | `docs/DISTRIBUICAO_DESKTOP.md` | Distribuição/configuração e limite de homologação explícito. |
| A | `docs/LABORATORIO_AUTHENTICODE.md` | Procedimento de VM, assinatura, testes e descarte futuro. |
| A | `docs/AUDITORIA_ATUALIZADOR.md` | Esta relação exata e evidências da auditoria. |

## Exclusões conferidas

- `dist-desktop/` inteiro: instaladores, win-unpacked/ASAR, blockmaps, metadados, hashes e pacotes/feeds/kits de laboratório.
- `node_modules/`, `.desktop-test/` e logs; nenhuma dependência gerada deve ser adicionada neste checkpoint.
- Instaladores/pacotes gerados fora dessas pastas: EXE, MSI, MSIX, APPX, blockmap, ASAR e NSIS 7z.
- PFX/P12/PEM/KEY, certificados públicos e requisições, pastas locais de certificados/segredos, `.env` (com exceção apenas de exemplo), arquivos locais de senha e temporários.
- PDFs, inclusive em pasta de clientes. Fixtures de PDF desta implementação são geradas pelos testes, não documentos reais versionados.

`git check-ignore` confirmou artefatos reais de distribuição/kit e nomes representativos de certificado, chave, `.env`, senha local, temporário e PDF. Não foram criados segredos para fazer essa verificação. Não há esses artefatos nos candidatos do checkpoint. Referências a variáveis de ambiente e valores fictícios de testes não são credenciais implantadas. O ignore não remove arquivos já rastreados; o histórico prévio de node_modules permanece preservado.

## Verificações executadas neste checkpoint

| Verificação | Resultado |
| --- | --- |
| Node (`npm test`) | 176 aprovados, zero falhas/skips; `%TEMP%/4printshops-checkpoint-node.log`. |
| Sintaxe | `node --check` nos 25 JS/MJS/CJS candidatos; parsing PowerShell do helper aprovado. |
| Chrome headless curto | Abas PDFs 100/350, QR/Início, exportação, proteção e fechamento/reabertura; aprovado. Log `%TEMP%/4printshops-checkpoint-browser.log`, artefatos `%TEMP%/4printshops-browser-u8it03/`. |
| Electron real em desenvolvimento | Worker/miniaturas, giro/exportação, cancelamento/falha, PDF pendente oculto, persistência/origem, sandbox e cópia seletiva sem AcroForm; aprovado. Log `%TEMP%/4printshops-checkpoint-desktop.log`, artefatos `%TEMP%/4printshops-desktop-y3cktv/`. |
| Whitespace | `git diff --check` sem erros. |
| Git | Staging vazio; inventário inclui novos não rastreados. Sem commit/push. |

Não se repetiram builds/reinstalações ou a suíte longa de PDFs sem alteração nos motores; seus resultados anteriores estão no histórico. **Não foi configurada VM, criada confiança, assinado pacote nem homologada instalação/reinício automáticos nesta auditoria.**

## Mensagem proposta e próxima etapa

`Implementa atualizador Desktop e prepara laboratorio de assinatura`

Criar o commit somente após autorização explícita do usuário, adicionando apenas os 34 arquivos acima e conferindo novamente status/diff/staging. A próxima funcionalidade será a associação de PDFs no Windows para abertura por duplo clique; **não foi iniciada**. O ensaio de atualização assinada fica pendente para uma etapa futura.
