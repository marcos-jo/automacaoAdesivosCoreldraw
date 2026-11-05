# Livreto grampeado — experimental.24

Data: 05/11/2026. Base limpa `2a6b49c` (checkpoint .21–.23 concluído). Uma janela independente recebe uma cópia dos bytes atuais do PDFWorkspace, incluindo giro, exclusão, união, reorganização e demais alterações ainda não exportadas. Não substitui o original, não marca a origem como exportada e não inicia download ao configurar.

## Escopo e paginação

Primeiro modo: **um único caderno com dobra central e grampo canoa**, leitura da esquerda para a direita. Origem deve estar na ordem normal de leitura. Não há corte/empilhamento, lombada colada, múltiplos cadernos ou compensação automática de espessura.

O total é completado até múltiplo de quatro, com brancos **no final**. A prévia mostra a proposta; salvar exige autorizar esses brancos. Não recebe tinta um rótulo de página branca. Faces consecutivas da externa à interna: frente da folha 1, verso da folha 1, frente da folha 2, verso da folha 2 etc.

| Originais | Brancos propostos nas posições | Total lógico | Folhas físicas | Faces do PDF |
|---:|---|---:|---:|---:|
| 4 | nenhum | 4 | 1 | 2 |
| 8 | nenhum | 8 | 2 | 4 |
| 12 | nenhum | 12 | 3 | 6 |
| 16 | nenhum | 16 | 4 | 8 |
| 5 | 6, 7, 8 | 8 | 2 | 4 |
| 7 | 8 | 8 | 2 | 4 |
| 10 | 11, 12 | 12 | 3 | 6 |

Exemplo 8 páginas, folha paisagem e verso normal: externa frente `[8 | 1]`, externa verso `[2 | 7]`, interna frente `[6 | 3]`, interna verso `[4 | 5]`. A prévia de leitura percorre capa, aberturas `[2 | 3]`, `[4 | 5]`, `[6 | 7]` e contracapa; deriva da simulação de dobra/aninhamento dessas mesmas faces. Não há uma segunda tabela de paginação independente.

Limite inicial: 2.000 páginas de origem, um caderno; limite técnico, **não recomendação de produção** para dobrar/grampear esse volume. O limite PDF de 256 MiB e os prazos/cancelamentos da prévia continuam ativos. Documentos grandes geram também uma cópia vetorial de leitura e podem consumir memória; ainda não há geração incremental das faces do livreto.

## Medidas, escala e frente/verso

- Final nominal A5: **148 × 210 mm**. A4 paisagem: **297 × 210 mm**. Dobrar A4 produz **148,5 × 210 mm antes do refile**, não A5 nominal exato. As duas molduras de 148 mm ficam unidas no centro e deixam 0,5 mm em cada borda externa. Não há redução escondida.
- Final A4: **210 × 297 mm** em A3 paisagem **420 × 297 mm**, exatamente duas páginas. A5 em A3 também é permitido para reservar espaço de marcas/refile. A4 em A4 recebe diagnóstico e permanece bloqueado.
- Papel retrato gira **toda a montagem em 90°**, incluindo conteúdos, molduras e marcas. Não confundir esse giro com a orientação do documento de origem nem com o duplex do driver.
- Conteúdo pode ficar a 100%, ajustar proporcionalmente ou usar escala explícita com vírgula/ponto. Ajustar pode reduzir **ou ampliar** e informa o percentual efetivo. O formato final não muda. Conteúdo que excede a moldura ou suas margens bloqueia a saída até ajuste; não há autorização para perder texto/arte.
- Margens são relativas a cada página na leitura, não espelhadas. Registro X/Y desloca somente conteúdo do verso nos eixos físicos da folha; não move a dobra e não altera configurações permanentes. Deslocamento que corte conteúdo fica bloqueado.
- Verso normal ou giro explícito 180°: a escolha é visível antes de salvar. O modelo normal corresponde à virada sobre o eixo da dobra; o giro 180° compensa o outro eixo. Testes geométricos cobrem os dois eixos nas duas orientações. **Isso não homologa o duplex de um driver ou o caminho físico de uma máquina.**

Duplex pode ser automático ou manual; o próprio [guia de livretos do Acrobat](https://helpx.adobe.com/in/acrobat/kb/print-booklets-acrobat-reader.html) distingue esses caminhos. A Microsoft define `DMDUP_HORIZONTAL` como encadernação na borda curta na documentação de [DEVMODEW](https://learn.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-devmodew), mas este módulo **não escreve duplex no DEVMODE** nem supõe que nomes/rotação/alimentação de todas as impressoras sejam equivalentes. A correspondência efetiva deve ser conferida com quatro páginas.

## Marcas, sangria e fidelidade

Marcas de corte nos cantos externos do bloco e dois indicadores externos de dobra. A guia de dobra/molduras na tela não vira uma linha atravessando as artes. Sangria exige BleedBox real e quantidade disponível após escala; não é inventada. Clip vetorial no encontro central impede a sangria de uma página de cobrir a vizinha. A prévia de leitura limita cada página ao corte final; não é um PDF de produção alternativo.

Marcas/sangria externas que saem da folha aparecem em área ampliada com hachura. É possível aumentar o papel, reduzir/desligar marcas ou autorizar explicitamente **somente esse recorte externo**. A autorização é invalidada quando muda a configuração. Nunca autoriza conteúdo fora da página final nem incompatibilidade técnica. A área imprimível do driver não é consultada nesta janela: é informada como não verificada.

O gerador compartilha os mesmos Forms vetoriais, fontes, grupos de transparência e transformações do motor já existente, acrescentando clip apenas nos posicionamentos de livreto. Step & Repeat/N-up conservam seus planejadores. Sem rasterização de páginas inteiras. ICC, cores especiais, overprint e fidelidade de RIP não recebem nova garantia universal. Arquivos com estrutura especial usam o diagnóstico/preparação já existente: consentimento, comparação de texto/aparência e bloqueio de ActualText essencial, formulários e aparências não comprovadas. Preparar ou voltar ao original afeta somente a cópia da janela.

## Operação

1. Abrir o PDF normal e concluir as edições desejadas. Pode permanecer com alterações não exportadas.
2. **Editar → Livreto grampeado**, ou selecionar Livreto no tipo da janela Montagem / Imposição.
3. Escolher final, papel e orientação. Conferir contagem, proposta de brancos e escala efetiva. Abrir Configurações avançadas somente quando necessário.
4. Conferir frente/verso de cada folha e mudar para Ordem de leitura. Anterior/Próxima e zoom são somente da prévia. Cancelar deixa o workspace intacto.
5. Autorizar brancos e eventual recorte externo de marcas/sangria. **Gerar PDF de livreto** pede confirmação e salva um arquivo independente, sem substituir o original.
6. Não há botão de envio físico novo. Para testar o PDF gerado, abri-lo no editor pelo fluxo normal, protegendo o descarte de qualquer trabalho anterior. No driver: papel/orientação corretos, escala 100%, sem livreto/N-up adicional. Primeiro conferir virtualmente; depois uma prova física deliberada pelo operador.

## Exemplos e roteiro de dobra/grampeamento

Gerar com `node scripts/generate-booklet-examples.mjs`. Os 28 arquivos sintéticos (origem, imposto, conferência de leitura e mapa JSON para cada quantidade) ficam em `.desktop-test/booklet24/examples/`, ignorados pelo Git. Não usar o PDF de conferência de leitura para produção: imprimir somente o arquivo **Livreto … A4 paisagem.pdf**.

1. Começar por quatro páginas numeradas A5. Conferir visualmente frente `[4 | 1]`, verso `[2 | 3]`, topo e base identificáveis.
2. No equipamento escolhido, fazer somente uma prova autorizada: A4 paisagem, escala 100%, uma cópia, sem segunda imposição. Selecionar duplex conforme o eixo de dobra e as instruções do driver. Se necessário, testar alimentação manual de frente/verso. Nenhum ajuste global é feito pelo aplicativo.
3. Dobrar ao meio e conferir capa 1, abertura 2/3 e contracapa 4, todos em pé. Verso invertido exige corrigir virada do driver **ou** giro do verso no PDF; não aplicar ambas as compensações sem conferir.
4. Conferir refile nominal de 0,5 mm para A5. A marca de dobra externa pode não caber em A4; use papel maior para marcas ou reconheça que elas serão cortadas. Não inferir área imprimível suficiente a partir do encaixe geométrico.
5. Repetir com 8/12/16 páginas, aninhar as folhas da externa à interna, dobrar e só então grampear no vinco. Conferir todas as páginas e o registro frente/verso. Repetir com 5/7/10 para conferir os brancos finais.
6. Homologar separadamente A4 em A3, bandeja/duplex/alimentação/gramatura/driver de cada impressora e tiragem. Esta etapa não compensa creep, espessura, grampo, compensação de dobra ou limitações de acabamento.

## Validação e distribuição

**Concluído:** 362 testes Node (328 anteriores + 34 novos); roteiros integrados completos Chrome headless e Electron **realmente instalado**; Electron smoke; sintaxe dos18 arquivos JS/CJS/MJS candidatos; `git diff --check`; auditoria ASAR454 contra os fontes. Os roteiros cobrem paginação, matrizes do PDF, duas prévias, cancelamento, cinco renderizações/quatro cancelamentos, retry, brancos, reentrada, temas, cancelamento de salvar e cópia compatível. Na prova sintética de 7 páginas, 18 pares de regiões/10.953.600 componentes RGBA foram **idênticos** entre origem, imposição, cópia de inspeção e leitura; sete números de página extraíveis. Não equivale à homologação de todos os PDFs ou RIPs.

O primeiro Electron smoke excedeu o prazo na preparação seletiva. O ensaio utilizava uma espera genérica que não assegurava o diagnóstico do arquivo recém-importado; foi acrescentada espera explícita pelo nome/diagnóstico. A repetição passou sem alterar detecção/preparação do aplicativo. O roteiro integrado também expôs uma limpeza incorreta do teste compacto: destruía a janela mantida pelo controlador para reabrir. O teste agora apenas fecha a instância do editor e destrói somente as instâncias próprias do ensaio; não foi alterado o ciclo de vida do aplicativo para contornar isso.

Houve ainda uma espera excedida, **não reproduzida nas duas rodadas seguintes**, na prévia de impressão da cópia preparada pelo fluxo antigo (antes de abrir o livreto). A janela permanecia ocupada na inicialização da prévia; não foi confirmada a causa nem aplicada correção ao motor. Os roteiros completos finais passaram, mas não demonstram resolução de uma condição intermitente. Logs preservados, sem confundir essa observação com a falha de prévia da .23, cuja causa também é desconhecida.

O roteiro instalado final salvou o livreto e produziu **quatro faces exclusivamente Microsoft Print to PDF**, além das três regressões virtuais anteriores (montagem, preparação compatível e compacto). O PDF gerado tem A4 exata; a saída do driver virtual mede aproximadamente210,01567 ×297,01068 mm. Na prova visual separada da saída virtual, a fração de componentes com diferença maior que16/255 ficou entre0,4255% e0,8856% por face (limiar do ensaio2%); não é igualdade de pixels nem homologação universal. Foram encontrados0 recursos de imagem e0 fontes no PDF virtual: o caminho Windows/PDFium existente converteu texto em contornos. O **PDF imposto salvo conserva as fontes e o texto extraível**, como comprovado separadamente. Não prometer pesquisabilidade da saída de impressão.

### Instalador e preservação do perfil

`dist-desktop/hold-native-test/4PrintShops-Hold-Test-1.0.0-experimental.24-Windows-x64-Setup.exe`

SHA256: `AF1C4B0761EC6F427C0A1C2DDDBE2E22ADAA788F727E45B0E8E332EE4DCCE254`.

Atualizou .23 → .24 no mesmo diretório Hold Test, appId `br.com.4printshops.holdtest`, chave NSIS e perfil `%APPDATA%/4PrintShops Hold Test`, origem `app://4printshops`. Comparados/preservados os valores anteriores de três preferências (painéis PDF, marcador da configuração de impressão e tema); um marcador específico adicional atravessou a instalação e foi removido ao final. Não havia formatos personalizados salvos nesse perfil para comparação valor a valor. Nenhum mecanismo de migração/exclusão de dados foi adicionado. A instalação principal permanece separada. Consulta somente de leitura confirmou a impressora padrão Konica e sua porta inalteradas.

Instalador experimental sem certificado Authenticode comercial; proteções do atualizador intactas. Não foi publicado, oferecido pelo mecanismo assistido ou tratado como atualização automática homologada.

### Reproduzir as verificações

Na raiz do checkout:

```powershell
node --test tests/*-*.test.mjs
node scripts/generate-booklet-examples.mjs
node tests/desktop-smoke.mjs
$env:IMPOSITION_DESKTOP='0'
$env:IMPOSITION_RESULTS='.desktop-test/booklet24/chrome-integrated-final'
node tests/pdf-imposition-browser.mjs
# Fechar a Hold Test depois de exportar trabalhos antes do roteiro instalado.
$env:IMPOSITION_DESKTOP='1'
$env:IMPOSITION_INSTALL='0' # usar 1 somente para instalar o pacote .24 já gerado
$env:IMPOSITION_RESULTS='.desktop-test/booklet24/installed-final'
node tests/pdf-imposition-browser.mjs
node scripts/check-desktop-package.mjs '.desktop-test/hold-lab/installed-validation/installed/4PrintShops Hold Test/resources/app.asar'
git diff --check
```

`BOOKLET_ONLY=1` executa só os cenários novos; o roteiro padrão inclui as regressões. Os testes instalados enviam **somente para Microsoft Print to PDF**, verificam o nome da fila antes de clicar e usam destinos sintéticos. Não são um comando de homologação física. `tests/desktop-profile-proof.mjs before` requer .23 fechada/instalada, captura o perfil e grava prova local; `after` requer .24 fechada/instalada e compara os valores. Não repetir `before` sobre .24 nem usar esse ensaio com trabalho aberto.

Evidências ignoradas em `.desktop-test/`: `booklet24-node-final.log`, `booklet24-electron-recheck.log`, `booklet24-chrome-integrated-final.log`, `booklet24-installed-final.log`; planos/PDFs/capturas e provas de pixels em `booklet24/chrome-integrated-final/` e `booklet24/installed-final/`; comparação das preferências em `booklet24/preferences-after.json`; análise da saída virtual em `booklet24/virtual-analysis.json` e `booklet24/virtual-appearance.json`. Preservados também os logs das tentativas inconclusivas. Todos são sintéticos/locais, fora do Git.

### Arquivos da implementação

Novos módulos `public/js/tools/pdf-booklet-layout.js` (paginação/geometria) e `pdf-booklet-ui.js` (janela/cópia/prévias). Alterados `pdf-imposition-engine.js` (materializador comum e clip adicional do livreto), `pdf-imposition-ui.js` (entrada para a janela), `pdfs.js` (captura compartilhada/menus/encerramento), `public/css/theme.css` (estilo restrito ao livreto) e `electron-builder.hold-native-test.cjs` (versão24). Planejadores Step & Repeat/N-up, PDFWorkspace, prévia principal e todos os motores nativos de impressão permanecem inalterados.

Novos testes `tests/pdf-booklet-layout.test.mjs`, `pdf-booklet-engine.test.mjs`, `desktop-profile-proof.mjs`; auxiliares `tests/helpers/pdf-booklet-{browser,fixtures,pixel-compare}.mjs`; gerador `scripts/generate-booklet-examples.mjs`. Adaptados `tests/pdf-imposition-browser.mjs`, `tests/desktop-smoke.mjs`, `tests/helpers/pdf-imposition-{modules,proof,compact-browser}.mjs`. Documentação neste guia, `ESTADO_PROJETO.md` e `DISTRIBUICAO_DESKTOP.md`.

Estado final:22 caminhos pendentes (12 modificados/10 novos), branch `main`, base `2a6b49c`, staging vazio. Nenhum commit/push. Os2.212 arquivos node_modules historicamente versionados não foram alterados. Instalador, exemplos e evidências continuam ignorados pelo Git.

**Pendência histórica preservada:** desaparecimento ocasional da prévia Step & Repeat/N-up em uma combinação ainda não identificada, relatado na .23. Não foi declarado corrigido por esta janela independente. Duplex físico/livretos grampeados ainda não homologados. Cut & Stack exige outra paginação/empilhamento; múltiplos cadernos exigem limites por caderno; creep necessita espessura medida, quantidade de folhas e estratégia validada de compensação. Apenas possibilidades futuras, sem controles fictícios.

Nenhum trabalho físico enviado pelo desenvolvimento, alteração global de impressora, commit ou push nesta etapa. Assinatura comercial/updater automático, piloto assistido em dois PCs e SMB real mantêm as limitações anteriores.

## Pesquisa para etapas futuras — não implementadas

Os [tipos de livreto documentados pela Adobe](https://helpx.adobe.com/ie/indesign/desktop/print/print-booklets/booklet-types.html) separam grampo canoa de Perfect Bound, que organiza páginas em cadernos de tamanho definido. Para uma futura lombada colada, o projeto precisará representar cadernos/limites, sequência de reunião, capa/lombada e preparação física; não basta trocar o nome do modo atual. Cut & Stack, solicitado para outra etapa, precisará de um plano próprio que modele posição na folha, ordem de corte, pilhas e reunião, com simulação inversa e provas numeradas antes de impressão.

A [compensação de creep da Adobe](https://helpx.adobe.com/indesign/desktop/print/print-booklets/adjust-creep-settings-in-booklets.html) desloca páginas progressivamente em relação à lombada para compensar espessura/dobra; os [ajustes de livreto](https://helpx.adobe.com/indesign/desktop/print/print-booklets/booklet-printing-settings.html) distinguem sangria, margem, espaçamento, creep e caderno. Para o 4PrintShops, recomenda-se exigir papel medido, folhas por caderno, regra de deslocamento validada e diagnóstico de perda de conteúdo. Não deduzir creep apenas pela gramatura e não ativá-lo automaticamente. Essas recomendações são planejamento próprio baseado nos conceitos, sem copiar algoritmos de terceiros.
