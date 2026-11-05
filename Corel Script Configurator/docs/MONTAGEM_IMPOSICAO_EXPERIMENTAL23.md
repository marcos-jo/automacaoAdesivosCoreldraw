# N-up em células e compacto — experimental.23

## Aprovação manual e checkpoint .21–.23 — 05/11/2026

O usuário testou fisicamente a experimental.23 na gráfica e confirmou que o **N-up compacto posicionou as peças conforme esperado** e que **Step & Repeat continua funcionando**. Aprovação restrita aos cenários efetivamente exercitados; o relato não discrimina todas as grades, orientações, escalas, materiais ou opções de driver. Não se extrapola para homologação universal de cores, RIP, sangria, acabamentos ou tamanhos. Os testes físicos foram executados pelo usuário; não houve envio físico automatizado pelo desenvolvimento.

O usuário autorizou o commit local das experimentais .21/.22/.23 com a mensagem `Consolida melhorias de imposicao e N-up compacto`, mantendo a pendência abaixo. Auditoria:32 candidatos relacionados (15 modificados/17 novos), staging inicialmente vazio, base `529a144`. Instaladores, builds, PDFs, capturas, logs e segredos locais permanecem ignorados. Não houve alteração funcional nem tentativa de corrigir a prévia neste checkpoint.

Verificações rápidas do checkpoint:328 testes Node novamente aprovados; sintaxe dos25 scripts JS candidatos e do auxiliar PowerShell, integridade do pacote instalado (452 entradas) e `git diff --check` conferidos. A correspondência entre inventário, conteúdo dos arquivos e staging deve ser verificada antes do commit. Esses resultados não encerram o problema ocasional da prévia relatado abaixo. A bateria de navegador/instalado anterior foi aproveitada, sem repetição de impressão ou geração de instalador.

## Problema conhecido — desaparecimento ocasional da prévia

**Estado: aberto, não resolvido.** O usuário relatou uma combinação ocasional de controles que faz a prévia desaparecer. Segundo o relato, isso não impede o fluxo normal da gráfica no momento. A sequência exata, os parâmetros envolvidos e a causa ainda não foram determinados. Não se presume que seja o mesmo problema de cancelamento corrigido em etapas anteriores, nem se declara resolvido pelos testes automatizados já aprovados.

Registrar para investigação posterior: coletar a sequência de controles, configuração de montagem e eventuais erros do console, preferindo um PDF sintético reproduzível e sem incluir documentos de clientes no Git. Nenhuma investigação corretiva ou implementação foi realizada nesta etapa, conforme solicitado. O checkpoint é aceito **com essa pendência conhecida**.

Data: 05/11/2026. Base `529a144`, branch main. Antes das mudanças, os 27 caminhos pendentes da .21/.22 foram preservados em `.desktop-test/imposition23/before-1793906792135/` com inventário, hashes, cópia integral e patch. Sem commit/push. Todos os PDFs desta etapa são sintéticos.

## Causa reproduzida

O N-up tradicional usa células iguais, desconta reservas de marcas/sangria e centraliza cada página na sua célula. Intervalo zero elimina o espaço **entre células/reservas**, mas não a folga interna de uma arte menor que a célula. O giro automático existente era opcional e inicialmente desligado.

Quatro páginas de 105 × 148 mm, grade 2 × 2, margens/intervalos zero e Contorno externo + divisões (afastamento3, comprimento3 e traço0,1 mm) reproduziram:

| Em células | Escala | Peça final | Folga entre colunas |
|---|---|---|---|
| A4 retrato, sem giro | 94,238095% | 98,95 × 139,472381 mm | 0 mm; ainda há folga entre linhas |
| A4 paisagem, sem giro | 66,858108% | 70,201014 × 98,95 mm | 72,248986 mm |
| A4 paisagem, giro permitido | 94,238095% | 139,472381 × 98,95 mm | 2,977619 mm |

Essas reduções são consequência do encaixe em células e da reserva das marcas, não uma alteração do documento original. Mesmo com giro, as marcas reservadas impedem 100%. O que faltava era outro modelo de distribuição; o tradicional continua disponível e inalterado.

## Comportamentos explícitos

**Em células (tradicional)** mantém centralização e ajuste por posição. A interface explica que pode haver espaço interno mesmo com intervalo zero.

**Compacto (peças encostadas)** coloca peças por suas dimensões finais, sem células uniformes. Preenchimento horizontal cria faixas da esquerda para a direita; a faixa seguinte começa abaixo da maior altura da anterior. Vertical cria colunas de cima para baixo e avança pela maior largura da anterior. Preserva a ordem/identidade das páginas e não sobrepõe retângulos.

Na direção de preenchimento, os intervalos entre cortes são exatamente os configurados. Para peças iguais, intervalo zero encosta nas duas direções. Peças mistas podem deixar vazios inevitáveis no outro eixo da faixa. Não é um encaixotador arbitrário de mosaicos, não reorganiza páginas para diminuir resíduos e não promete contato de todas as bordas de retângulos incompatíveis. Não utiliza uma grade uniforme para deformar ou redimensionar peças diferentes.

A grade limita peças/faixas por folha; 10 páginas em 4-up geram 3 folhas, sem duplicar ou eliminar páginas. Folhas parciais usam o bloco realmente ocupado. Vagas explícitas `0` permanecem suportadas em células; compacto pede páginas reais, evitando inventar dimensões de uma vaga.

## Escala e giro

- Em tamanho real, usa 100%. Em personalizada, usa o percentual informado. Não reduz por causa de marcas externas.
- Em ajuste proporcional, aplica **um fator comum** ao documento para que o maior bloco caiba nas margens, preservando as relações físicas entre páginas diferentes. Esse encaixe não aumenta acima de 100% e não reserva espaço para marcas externas. Os percentuais calculados continuam explícitos no resumo.
- Giro automático de 0°/90° só ocorre com **Permitir giro automático** marcado. O mesmo controle existente fica visível na área básica do compacto, sem duplicação de campos/listeners. Mostra a orientação final e o giro de cada página no resumo/detalhes/confirmacão antes da saída.
- Até 8 peças por folha, avalia todas as combinações de giro. Acima disso, usa busca determinística limitada a 32 candidatos por etapa, sem prometer ótimo global. Não faz busca exponencial em documentos grandes. Nas 350 posições sintéticas, a busca permaneceu limitada, sem alterar ordem ou sobrepor peças.
- Trocar a distribuição remove o fator adicional anterior e restaura a grade pré-ajuste quando aplicável. Métodos/percentuais escolhidos continuam explícitos. As configurações próprias de Step & Repeat/N-up da .22 permanecem, incluindo distribuição N-up ao sair/voltar. Papel/orientação compartilhados continuam visíveis. Não há fator oculto.

### Referências geométricas confirmadas

Sem margens, sangria ou marcas:

| Compacto 2 × 2 | Peça | Bloco completo | Escala |
|---|---|---|---|
| A4 retrato 210 × 297 mm | 105 × 148 mm, giro0° | 210 × 296 mm | 100% |
| A4 paisagem 297 × 210 mm, giro autorizado | 148 × 105 mm, giro90° | 296 × 210 mm | 100% |

Sobra 1 mm no eixo correspondente, distribuído em 0,5 mm de cada lado quando centralizado. Se o giro estiver desligado em paisagem, o ajuste proporcional pode reduzir para aproximadamente70,945946%; a opção real mantém100% e avisa que o bloco excede o papel. Não se gira silenciosamente.

## Marcas e sangria

Contorno externo + divisões projeta os cortes nas bordas do retângulo envolvente real das peças. O caso 2 × 2 uniforme produz12 segmentos, fora do conteúdo. Apenas contorno externo produz8. Sem marcas não produz segmentos. A prévia e a saída usam o mesmo plano/linhas vetoriais.

Quatro A6 a100% em A4 não deixam espaço para marcas externas. O compacto **mantém as peças em100%**, mostra o excesso e exige a decisão de corte da .21. Marcas fora da folha não serão impressas. Alternativas: folha maior, aceitar conscientemente o recorte ou usar **Ajustar montagem para caber**, que então reserva o espaço das marcas e mostra o novo percentual antes de aplicar. No exemplo retrato, esse ajuste explícito resultou em94,238095%; remover restaura100%.

Os intervalos do compacto são **entre cortes finais**. Com sangria positiva, exige pelo menos duas vezes essa sangria entre peças adjacentes; intervalo zero só encosta sem sobrepor sangria quando ela é zero. Não cria nem compartilha artificialmente sangrias. As validações de sangria real após escala permanecem.

Marcas individuais exigem espaço para não atravessar a peça vizinha. Em vez de adicionar espaço escondido ou desenhar sobre conteúdo, a combinação insegura é bloqueada e oferece **Usar marcas externas + divisões**. Com intervalos suficientes, as individuais continuam disponíveis. Marcas externas precisam ficar fora da reserva de sangria, incluindo metade do traço. A verificação individual usa consulta espacial e orçamento limitado de comparações; se excedido, pede reduzir peças ou usar externas. Não permite saída não verificada.

Com tamanhos mistos, os cortes projetados representam coordenadas de peças diferentes; nem toda divisão permite passar a guilhotina pela folha inteira. O resumo/confirmacão avisam que podem ser necessários refiles separados. A grade compacta não é homologação de um plano de cortes. Guias azuis são apenas de inspeção; não são impressas no PDF.

## Preservação da base

- `pdf-imposition-compact.js` isola o novo posicionamento/rotação e o limite de escala por faixas.
- `pdf-imposition-crop-marks.js` compartilha as **mesmas fórmulas** de marcas existentes. Foram comparados **768 planos tradicionais/Step & Repeat com o backup .22, todos idênticos**, incluindo giros, ordens, percentuais, fatores e quatro modos de marcas. Não houve alteração de tamanho/escala/posição do Step & Repeat.
- `pdf-imposition-layout.js` encaminha apenas a opção compacta ao novo módulo; `pdf-imposition-bounds.js` calcula o ajuste explícito usando a extensão real das faixas.
- `pdf-imposition-ui.js`, `pdf-imposition-settings.js`, `pdf-imposition-summary.js` e `public/css/theme.css` acrescentam seleção, giro acessível, explicações, resumo de orientação e recuperação da escolha de marcas. Opção em células continua padrão.
- `pdf-imposition-engine.js`, gerador vetorial, PDFWorkspace, preparação ActualText, worker/cancelamento, exportação, associação e motores Windows **não foram alterados nesta etapa**. O motor recebe o plano compacto pelo mesmo caminho que já gerava/imprimia o PDF.
- Testes novos em `tests/pdf-imposition-compact.test.mjs` e `tests/helpers/pdf-imposition-compact-browser.mjs`; roteiro geral/browser básico adaptado e build Hold Test incrementado para .23.

## Validação e instalador

- **328 testes Node aprovados**, incluindo8 novos testes compactos: A6 retrato/paisagem, comparação de células, marcas externas, ajuste explícito repetido, mistos, intervalos, sangria, individuais, vagas,350 páginas e350 posições. O último teste confere64 posicionamentos pelas Matrix/BBox realmente gravadas nos PDFs, com originais byte a byte intactos.
- **Chrome headless** completo aprovado: comportamento básico/avançado, alternância/restauração, escala, folhas externas, aceite/revogação de corte, retry/cancelamento, preparação sintética ActualText e compacto. O roteiro compacto captura11 estados/56 posicionamentos serializados. Inclui o bloqueio de marcas individuais sobre artes e recuperação por ação explícita.
- Electron smoke aprovado: segurança, PDF.js/miniaturas, edição, abas, salvamento/cancelamento/falha, proteção de fechamento e persistência.
- Sintaxe de29 scripts JS e `git diff --check` aprovados. Logs e PDFs sintéticos em `.desktop-test/imposition23/` e `.desktop-test/imposition23-*.log`, ignorados.
- **Validação instalada concluída:** o log final, de 05/11/2026 às16:44, registra aprovação do roteiro completo. A instalação encontrada nesta retomada já é a .23 no mesmo diretório da .22. Foram exercitados 12 estados compactos e60 posicionamentos, incluindo A6 retrato/paisagem a100%, giro autorizado, divisões externas, aviso/aceite de excesso, recuperação das marcas individuais e restauração de opções. Original e pendências preservados.
- **Impressão exclusivamente virtual concluída:** existem os três PDFs do roteiro instalado — montagem tradicional, cópia preparada e compacto A6. O compacto gerou1 página A4 de aproximadamente210,01567 ×297,01068 mm, dentro da tolerância de0,1 mm do teste da fila Windows. A fonte vetorial da montagem também foi salva e suas matrizes conferidas. Nenhum trabalho foi enviado a fila física; o ensaio não homologa a área imprimível da Konica.
- **Reconferência nesta retomada:** sem repetir Node/Chrome/Electron nem reinstalar, foram lidos os logs aprovados, conferidos novamente os116 posicionamentos dos artefatos Chrome/instalado (56+60), os arquivos virtuais e o hash do instalador. A auditoria do ASAR instalado passou com452 entradas e igualdade dos arquivos de produção com o código atual. `git diff --check` passou novamente. Relatório local em `.desktop-test/imposition23/final-recovery-check.json`.
- **Perfil e preferências:** identidade `br.com.4printshops.holdtest`, origem `app://4printshops` e diretório `%APPDATA%\4PrintShops Hold Test` permanecem iguais aos da .22; o perfil existente foi criado em04/11/2026 e não foi substituído por um perfil de versão nova. O instalador mantém `deleteAppDataOnUninstall:false`; os testes anteriores de persistência/origem do Electron passaram. Não há inventário anterior valor a valor dos formatos/preferências do operador, portanto essa comparação individual não é afirmada; conferir as personalizações usuais na aceitação manual.

Instalador gerado: `dist-desktop/hold-native-test/4PrintShops-Hold-Test-1.0.0-experimental.23-Windows-x64-Setup.exe`.

SHA256: `77B9836E1DDA54398D3143A5EE2157BD60CC6F57553BE4FCAD7764C9D94ED653`.

Mesma identidade `br.com.4printshops.holdtest`, diretório e perfil Hold Test da .22. Continua experimental, sem publicação nem certificado instalado. Atualizador automático/assinatura, piloto assistido e SMB conservam as pendências anteriores.

Instalação efetivamente conferida: `.desktop-test/hold-lab/installed-validation/installed/4PrintShops Hold Test/4PrintShops Hold Test.exe`. Registro do Windows e metadados do ASAR indicam `1.0.0-experimental.23`. Não foi necessário gerar ou aplicar outro instalador nesta retomada. Não se encontrou defeito novo que exigisse alterar código; foram concluídos o registro dos resultados e a documentação.

## Conferência manual na gráfica

1. Exportar trabalhos pendentes, fechar a Hold Test e aplicar a .23. Conferir a versão e a ausência de instalação duplicada.
2. Abrir quatro páginas A6 de105 ×148 mm. Em Montagem, escolher N-up, origem1–4,4-up, A4 retrato e margens/intervalos/sangria zero. Selecionar Sem marcas e Compacto. Conferir peças105 ×148, escala100%, bloco210 ×296 e contato das bordas. Se vier de Individuais e houver conflito, usar a ação de marcas externas ou escolher Sem marcas.
3. Mudar para A4 paisagem. Marcar Permitir giro automático; conferir peças148 ×105, giro90°, escala100% e bloco296 ×210. Desmarcar e observar explicitamente a nova escala de encaixe ou excesso em tamanho real. Voltar a retrato e conferir a restauração de100%, sem apagar os parâmetros escolhidos.
4. Selecionar Contorno externo + divisões. Conferir12 marcas externas, nenhuma sobre as artes. Desativar guias azuis se necessário. Verificar o aviso de marcas fora do papel. Não considerar essas marcas imprimíveis em A4; escolher uma folha maior, ajuste explícito ou corte autorizado.
5. Aplicar Ajustar montagem para caber e ler o percentual; remover deve voltar ao tamanho-base. Trocar entre células e compacto e conferir as diferenças anunciadas, mantendo Step & Repeat independente.
6. Usar10 páginas de tamanhos diferentes, intervalos0 e5 mm e preenchimentos horizontal/vertical. Conferir sequência,3 folhas em4-up e as folgas locais de faixas mistas. Avaliar os refiles antes de produzir.
7. Gerar o PDF e conferir dimensões/rotações no arquivo reaberto. O original e suas alterações pendentes devem continuar intactos.
8. Testar primeiro em Microsoft Print to PDF, com A4 e100%, sem N-up adicional no driver. A área imprimível da Konica pode impedir bordas até a extremidade; esses encaixes são geométricos, não homologação física do driver. Qualquer envio real depende do operador, não foi executado nesta etapa.

Pendentes após a aprovação manual: investigar o desaparecimento ocasional da prévia; validar cenários adicionais de medidas, materiais/cores/RIP e plano de refiles mistos. A aprovação física anterior A3 → A6 → Step & Repeat permanece restrita àquele cenário, assim como a nova aprovação acima se restringe aos testes relatados. Parsing/save PDF-lib ainda têm trechos síncronos; os limites de lote, palco e renderização da .21 permanecem.

Histórico da retomada da validação:32 caminhos pendentes (15 modificados/17 novos), somando .21/.22/.23, base `529a144`, staging vazio. Naquela retomada, somente este guia, `docs/ESTADO_PROJETO.md` e `docs/DISTRIBUICAO_DESKTOP.md` receberam atualizações documentais. Os 2.212 arquivos históricos de node_modules rastreados permanecem intocados. Instaladores/builds, PDFs sintéticos, perfis, capturas e logs continuam ignorados. O registro local foi autorizado posteriormente, conforme o checkpoint no início deste guia; nunca executar push automaticamente.

## Inventário auditado do checkpoint

M = modificado em relação ao commit-base; A = novo. Todos os caminhos são relativos à raiz deste aplicativo. A lista abaixo contém exatamente32 arquivos; nenhum artefato binário, documento de cliente ou configuração secreta.

| Estado | Arquivo | Finalidade |
|---|---|---|
| M | `docs/DISTRIBUICAO_DESKTOP.md` | Pacotes e validação das .21–.23 |
| M | `docs/ESTADO_PROJETO.md` | Estado, aprovação manual e pendência da prévia |
| A | `docs/MONTAGEM_IMPOSICAO_EXPERIMENTAL21.md` | Excesso geométrico e interface básica/avançada |
| A | `docs/MONTAGEM_IMPOSICAO_EXPERIMENTAL22.md` | Escalas e opções independentes |
| A | `docs/MONTAGEM_IMPOSICAO_EXPERIMENTAL23.md` | Compacto, limitações e este checkpoint |
| M | `electron-builder.hold-native-test.cjs` | Versão experimental.23, mesma identidade |
| M | `public/css/theme.css` | Interface de montagem e avisos |
| M | `public/js/main.js` | Última geometria pública consultada do driver |
| M | `public/js/tools/pdfs.js` | Repasse da área imprimível à montagem |
| M | `public/js/tools/pdf-imposition-engine.js` | Prévia ampliada e metadados de medidas |
| M | `public/js/tools/pdf-imposition-layout.js` | Excesso optativo e encaminhamento compacto |
| M | `public/js/tools/pdf-imposition-preview.js` | Enquadramento/zoom com limites |
| M | `public/js/tools/pdf-imposition-ui.js` | Controles, escalas e confirmações |
| A | `public/js/tools/pdf-imposition-bounds.js` | Limites, aproveitamento e ajuste |
| A | `public/js/tools/pdf-imposition-compact.js` | Faixas e giro do compacto |
| A | `public/js/tools/pdf-imposition-crop-marks.js` | Fórmulas compartilhadas de marcas |
| A | `public/js/tools/pdf-imposition-settings.js` | Configurações por modo |
| A | `public/js/tools/pdf-imposition-summary.js` | Resumo físico/escala do resultado |
| A | `public/js/tools/pdf-imposition-view.js` | Sobreposições somente da prévia |
| M | `tests/helpers/pdf-imposition-compatibility-browser.mjs` | Confirmação no fluxo preparado |
| M | `tests/helpers/pdf-imposition-modules.mjs` | Carregador dos módulos no Node |
| M | `tests/helpers/windows-save-dialog.ps1` | Espera segura de controles nativos |
| M | `tests/pdf-imposition-browser.mjs` | Roteiro Chrome/instalado .23 |
| M | `tests/pdf-imposition-preview.test.mjs` | Cancelamento, recuperação e aceite |
| A | `tests/helpers/pdf-imposition-basic-browser.mjs` | Interface básica/avançada |
| A | `tests/helpers/pdf-imposition-compact-browser.mjs` | Compacto e impressão virtual |
| A | `tests/helpers/pdf-imposition-overflow-browser.mjs` | Excesso e ajuste explícito |
| A | `tests/helpers/pdf-imposition-proof.mjs` | Matrix/BBox do PDF efetivamente gerado |
| A | `tests/helpers/pdf-imposition-scale-browser.mjs` | Alternâncias/escala no navegador |
| A | `tests/pdf-imposition-compact.test.mjs` | Geometria, mistos, marcas e matrizes |
| A | `tests/pdf-imposition-overflow.test.mjs` | Excesso, limites e regressões |
| A | `tests/pdf-imposition-scale.test.mjs` | Estado independente e dimensões |
