# LibreOffice privado — experimental.32

**Checkpoint local autorizado em 08/11/2026:** os21 arquivos desta integração foram auditados, com nova execução dos437 testes Node, sintaxe dos13 scripts e diff aprovados. Consolidação local com a mensagem `Inclui motor privado LibreOffice na experimental 32`, sem runtime/cache, binários gerados, perfis, evidências ou clientes no Git. Nenhum novo build, instalação, impressão ou push. Inventário e contexto em [ESTADO_PROJETO.md](ESTADO_PROJETO.md); menções anteriores de “sem commit” descrevem o desenvolvimento anterior ao checkpoint. Windows limpo e testes em massa na gráfica continuam pendentes.

Etapa de 08/11/2026, branch `main`, base `e350f126acda0ea2dfc1213c11c2f744fba48a21`. O Desktop passa a levar uma cópia privada do LibreOffice para converter DOCX/DOC. A entrada continua sendo **Importar** para PDF, imagens e Word. Não há instalação separada de Office, registro de fontes/DLLs no Windows ou novos atalhos/associações de Word por esse componente.

## Organização e execução

- `build/libreoffice-lock.json`: versão 26.8.2.2 x64, build oficial, URLs e SHA256 do MSI e quatro arquivos de código-fonte.
- `scripts/prepare-libreoffice.mjs`: baixa somente com `--download`, verifica hashes, extrai imagem administrativa do MSI e prepara `build/libreoffice/runtime`. Não instala o MSI no Windows.
- `desktop/libreoffice-runtime.mjs`: valida manifesto e todos os arquivos, recusa ausências, alterações, arquivos extras, links e caminhos fora da árvore; resolve o executável privado e prepara o ambiente do processo filho.
- `electron-builder.config.cjs`: exige a preparação íntegra antes do build; incorpora a confiança no ASAR e copia o motor para `resources/libreoffice`. A variante Hold Test usa experimental.32, mantendo sua identidade/perfil.
- `desktop/app.mjs`, `import-files.mjs`, `word-import.mjs`: encaminham o resolvedor privado aos dois canais Word/unificado, preservando a política de conversão e os motores anteriores.

No aplicativo empacotado, motor privado ausente ou inválido bloqueia Word com orientação de reinstalar. Não procura uma instalação externa nem aceita redirecionamento pela variável `PRINTSHOPS_LIBREOFFICE`. Em desenvolvimento, a instalação externa continua sendo alternativa somente quando a pasta privada não existe; uma pasta privada inválida também bloqueia.

A verificação integral acontece uma vez por processo do aplicativo, antes de usar o motor, com oito leitores concorrentes. O ambiente filho remove caminhos UNO/Python herdados e restringe PATH ao System32. O cache Python fica no temporário da conversão, evitando gravação na árvore distribuída. Perfil, saída e cache são exclusivos da operação. Isso não constitui isolamento completo do Windows.

Os binários originais foram mantidos sem alterações de código. As fontes da pasta administrativa `Fonts` são copiadas para `share/fonts/truetype`, e as DLLs x64 de `System64` ficam junto ao programa. A primeira versão conserva a distribuição completa: 19.410 arquivos inventariados, 1.581.539.508 bytes, além do manifesto. A redução do pacote fica para uma etapa posterior com testes próprios.

## Preparar em outra máquina

Com Windows x64 e as dependências npm do projeto disponíveis, executar na subpasta `Corel Script Configurator`:

```powershell
node scripts/prepare-libreoffice.mjs --download
npx electron-builder --config electron-builder.hold-native-test.cjs --win nsis --x64 --publish never
```

Para aproveitar um MSI já disponível, definir `PRINTSHOPS_LIBREOFFICE_MSI` com seu caminho absoluto. O hash deve coincidir com o lock. Os quatro arquivos-fonte também precisam estar no cache; `--download` permite obtê-los. Sem essa opção, arquivos ausentes causam erro, sem downloads automáticos. A preparação não substitui silenciosamente um motor existente divergente.

`build/libreoffice/`, `build/libreoffice-cache/`, instaladores, perfis e evidências são gerados/ignorados no Git. É preciso preparar os recursos em cada máquina de build; eles não acompanham o clone do repositório. Reservar vários GB livres para cache, extração, runtime, pacote e compressão.

## Licenças e fontes

Foram preservados LICENSE.html, license.txt, NOTICE, CREDITS.fodt e readmes originais, inclusive avisos de componentes/fontes. `third-party/libreoffice/NOTICE.txt` acrescenta identificação da redistribuição, sem atribuir autoria ou endosso da The Document Foundation ao 4PrintShops. `SOURCE-CODE.json` acompanha o motor com versão, build, URLs exatas e hashes; os quatro arquivos-fonte correspondentes foram baixados e verificados no cache local e são exigidos no build.

Os arquivos-fonte ficam separados do instalador binário. Ao distribuir a terceiros, manter sua disponibilidade e atender também às condições dos componentes descritas nos avisos originais. Esta etapa local não publica o produto nem certifica toda a distribuição comercial. Referências: [licenças do LibreOffice](https://www.libreoffice.org/licenses/) e [fontes oficiais da versão](https://download.documentfoundation.org/libreoffice/src/26.8.2/).

## Testes executados

- 437 testes Node aprovados, zero falhas/skips. Os dez novos testes cobrem integridade, corrupção/ausência/extras/traversal, resolução sem fallback no pacote, desenvolvimento e propagação/isolamento do ambiente.
- Electron com motor privado: PDF + PNG + JPG + DOCX + DOC resultaram em sete páginas, com originais preservados; revisão, adicionar Word, falha/cancelamento, proteção de descarte e exportação reaberta passaram. Descoberta externa desativada no processo de teste.
- Teste dos módulos reais dentro do ASAR: DOCX e DOC produziram duas páginas cada, com caminhos externos inválidos; preservação dos originais, cancelamento, recusa do motor ausente e integridade após conversão passaram.
- Pacote de diretório e conteúdo extraído do Setup auditados: 465 entradas ASAR, igualdade com o código atual e todos os arquivos privados íntegros. Nenhuma impressão física ou virtual foi realizada.

Uma execução inicial identificou cache `__pycache__` criado pelo próprio LibreOffice dentro do runtime; o ambiente foi corrigido, o cache gerado foi removido e os testes finais passaram com a árvore íntegra. O primeiro build foi interrompido e substituído pelo build final. Logs locais em `.desktop-test/office32/`, não versionados.

Reprodução sem imprimir:

```powershell
npm test
node tests/desktop-word-smoke.mjs
node scripts/test-private-office.mjs dist-desktop/hold-native-test/win-unpacked
node scripts/check-desktop-package.mjs dist-desktop/hold-native-test/win-unpacked/resources/app.asar
```

O teste privado também aceita o diretório da instalação no lugar de `win-unpacked`. Ele executa conversões sintéticas usando os módulos e o motor daquele pacote, sem imprimir ou alterar documentos do cliente.

## Atualização local

A .32 foi instalada sobre a .31 em `%LOCALAPPDATA%/Programs/4PrintShops Hold Test` com retorno0 e mesma identidade/perfil. Backup anterior em `.desktop-test/office32/before/`. Abertura real confirmou Importar/ponte disponíveis, preferências idênticas e fechamento normal. Executável, ASAR, dois auxiliares, PDFium e ícone conferidos com o pacote. O teste privado foi repetido contra os módulos e o motor realmente instalados: DOCX/DOC de duas páginas, originais preservados, descoberta externa desativada e19.410 arquivos íntegros. Auditoria da instalação465 entradas aprovada. O LibreOffice global foi mantido. Sem impressão.

## Artefato e limites

Setup: `dist-desktop/hold-native-test/4PrintShops-Hold-Test-1.0.0-experimental.32-Windows-x64-Setup.exe`, 514.962.605 bytes. SHA256: `EBE8AFE2C8128BA646326056470D9523380748F1C64C818342AF54D998CCCFE7`. Manifesto do motor: `ae42c73b898efeb80813b024634d91cfd0e64503702b3c67e1588ea4b31d5781`.

Permanece pendente testar em um Windows limpo, sem Office: esta máquina mantém o LibreOffice global previamente autorizado. Invalidar a descoberta externa comprova o caminho usado, mas não substitui a validação das dependências em outra instalação do Windows. Também faltam os testes em massa com Word reais na gráfica, fontes/paginação e cor/RIP. Os limites e a revisão descritos em [IMPORTACAO_WORD_LOCAL.md](IMPORTACAO_WORD_LOCAL.md) continuam válidos. Motores aprovados não foram modificados. Nenhum commit ou push nesta etapa.
