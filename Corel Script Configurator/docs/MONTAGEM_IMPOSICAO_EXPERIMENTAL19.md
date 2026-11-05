# Contorno externo + divisões — experimental.19

05/11/2026. Complemento da experimental.18, sem commit/push. Os28 caminhos pendentes foram preservados antes da edição em `.desktop-test/imposition19/baseline18-20261105-114501/`, com inventário e patch. Mantidos os três modos anteriores e a correção da prévia.

## Comportamento

Novo modo no seletor **Marcas de corte → Contorno externo + divisões**:

- Mantém as marcas dos quatro cantos do bloco.
- Projeta cada coordenada de corte vertical nas bordas superior e inferior do bloco, com um segmento em cada borda.
- Projeta cada coordenada de corte horizontal nas bordas esquerda e direita, com um segmento em cada borda.
- Deduplica cortes coincidentes com tolerância numérica de1e-7mm. Duas peças encostadas compartilham uma única posição de marca por borda.
- Nenhum segmento entra no retângulo que envolve as artes. Comprimento/espessura/afastamento e reservas de sangria seguem os parâmetros existentes em mm.

**Exemplo2×1, intervalo/sangria0:** oito segmentos dos cantos + um acima e outro abaixo da divisão =10 segmentos. Uma grade regular C×L encostada tem2×(C+1)+2×(L+1) segmentos. Não são geradas marcas individuais completas.

Quando há intervalo positivo, os dois lados do intervalo são cortes distintos e recebem marcas distintas; não se inventa uma linha de corte no meio do espaço. N-up de formatos mistos projeta as bordas reais de cada peça, que podem produzir mais posições de marcas do que uma grade de peças iguais. Vagas não geram peças/marcas fictícias. Folhas parciais usam somente peças presentes e uma folha totalmente vazia não recebe marcas.

O novo modo reutiliza exatamente as posições e reservas do modo grupo; não altera escala, rotação, transformação vetorial ou papel. Com sangria positiva, o espaço necessário para a sangria continua preservado: intervalo0 não remove essa reserva.

Continuam disponíveis:

- **Individuais:** marcas completas por peça.
- **Apenas no grupo / contorno externo:** somente oito segmentos dos cantos.
- **Sem marcas:** sem segmentos de corte.

Guias azuis continuam apenas na prévia, não no PDF. Para avaliar visualmente somente as marcas pretas, desative as guias.

## Referência conceitual

A documentação oficial de [Smart crop marks do Quite Imposing](https://www.quite.com/docs/qi6/en/qi6_manual/b6_0012.html) descreve marcas que evitam sobrepor páginas e contempla tratamento de marcas entre peças. Ela foi consultada como referência conceitual. Este modo segue o exemplo solicitado pelo usuário: todas as marcas ficam exclusivamente nas bordas externas, sem reproduzir a política de marcas internas ou de dobra do Quite Imposing. Não se declara equivalência completa com aquele produto.

## Alterações

- `public/js/tools/pdf-imposition-layout.js`: modo `bordas`, projeção e deduplicação. Os ramos existentes mantêm seus cálculos; reservas externas são compartilhadas com `grupo`.
- `public/js/tools/pdf-imposition-ui.js`: quarto modo e explicação compacta.
- `tests/pdf-imposition.test.mjs`: cinco testes novos, incluindo coordenadas dos operadores vetoriais do PDF.
- `tests/pdf-imposition-browser.mjs`: seleção/visualização/exportação do novo modo e exemplo sintético2×1; regressões de prévia preservadas.
- `electron-builder.hold-native-test.cjs`: versão .19 na mesma identidade Hold Test.
- Documentação: este guia, `ESTADO_PROJETO.md`, `DISTRIBUICAO_DESKTOP.md`.

O gerador vetorial já desenha a lista de marcas externas do plano: não foi necessário alterar `pdf-imposition-engine.js`, a prévia assíncrona, os motores de impressão ou o PDFWorkspace.

## Validações executadas

- **301 testes Node aprovados** (296 anteriores +5).
- Exemplo2×1: exatamente10 segmentos, com divisão central no topo e embaixo. Grades1×3,2×2,3×2 e4×3, espaçamento positivo, N-up uniforme/misto, vagas e folhas parciais/vazias cobertos. Coordenadas em pontos no PDF conferidas contra os valores em mm, tolerância1e-6pt. Segmentos fora das artes e dentro da folha.
- Comparação de216 configurações dos três modos anteriores com a cópia da .18: plano completo idêntico, incluindo escala, posições, capacidade e marcas.
- Chrome real: seletor, prévia2×1, temas, geração após mudanças, cancelamentos e erro/Tentar novamente aprovados. PDF e PNG2×1 foram gerados com fontes/vetores e inspecionados visualmente; não são capturas da interface.
- .19 realmente instalada sobre a Hold Test existente, no mesmo diretório/perfil/identidade. Prévia, salvamento nativo, N-up e proteção da origem aprovados; impressão **exclusivamente Microsoft Print to PDF**. Nenhum trabalho físico enviado.
- Regressões de prévia: nove renderizações/oito cancelamentos, incluindo alternar guias durante carregamento, recuperação após falha injetada e ações liberadas.
- Comparação sintética de texto/transparência preservada:82.944 componentes RGBA idênticos. Pacote instalado auditado com444 entradas; sintaxe e `git diff --check` verificados.

Logs, comparação e exemplos ignorados em `.desktop-test/imposition19/`: `node-final.log`, `chrome.log`, `installed.log`, `package.log`, `legacy-geometry.json`; exemplo `browser/Contorno divisoes 2x1.pdf` e PNG correspondente; captura `installed/contorno-divisoes-2x1.png`.

## Instalador e teste manual

`dist-desktop/hold-native-test/4PrintShops-Hold-Test-1.0.0-experimental.19-Windows-x64-Setup.exe`

SHA256: `A8DA4CCA0F857771D34156FF3583A78773AF76B5AB9DAD1964BE22EA525FDB8B`.

Mesmo appId/NSIS/executável/perfil Hold Test; .18 preservada no diretório de artefatos. Sem assinatura Authenticode comercial; instalação automática do atualizador segue não homologada.

1. Exporte trabalho pendente, feche Hold Test e instale .19 no destino existente.
2. Em Montagem / Imposição, escolha grade2 colunas×1 linha, duas peças, espaços0/sangria0 e **Contorno externo + divisões**. Confira uma marca central acima e outra abaixo, sem traço entre as artes.
3. Teste2×2 e3×2, conferindo marcas das colunas no topo/embaixo e das linhas à esquerda/direita. Teste espaço positivo e observe duas linhas de corte em cada intervalo.
4. Salve o PDF de montagem e confira novamente em leitor. Alterne os outros três modos para comparar; a origem deve continuar intacta e pendente quando aplicável.
5. Teste N-up com formatos iguais e diferentes e folha parcial. A impressão física do novo modo ainda depende de validação do operador.

Limitações da .18 continuam: retângulo envolvente, sem faca irregular; N-up misto pode ter espaços internos; sem equivalência universal de cor/PDF-X/RIP. Aprovação física anterior limita-se ao cenário A3 → A6 → Step & Repeat relatado pelo usuário. Sem novas funcionalidades de livreto/frente-verso, mudanças de driver, commit ou push.

## Conferência na retomada de 05/11/2026

A implementação e o instalador .19 já estavam concluídos antes da interrupção. Foram relidos AGENTS, estado e guia; comparados os28 caminhos da cópia anterior com os atuais. Sete caminhos compartilhados contêm as mudanças esperadas da .19 (layout, interface, dois testes, builder, estado e distribuição), além deste novo guia. Não foi encontrada operação Git incompleta (`index.lock` ausente), staging preparado ou alteração inesperada. Os29 caminhos pendentes incluem as etapas anteriores; nada foi descartado.

Nesta retomada não houve nova implementação, renomeação, reconstrução de pacote ou impressão. O nome disponível **Contorno externo + divisões** identifica o modo solicitado como Grupo com divisões. O hash do instalador confere com o registrado acima. Auditoria do ASAR realmente instalado refeita:444 entradas e fontes/configurações correspondentes ao checkout.

Reexecutados somente os30 testes rápidos de geometria/estabilidade:30 aprovados (`.desktop-test/imposition19/resume-tests.log`), além de sintaxe e `git diff --check`. Os resultados anteriores de301 testes Node, Chrome e aplicativo instalado foram conferidos nos logs, sem repetir instalação e impressão virtual. Pendência: conferência manual das marcas na gráfica. Exemplo2×1:10 segmentos;2×2:12;3×2:14, para grades regulares encostadas sem sangria. Desativar guias azuis para avaliar apenas as marcas pretas exportadas.
