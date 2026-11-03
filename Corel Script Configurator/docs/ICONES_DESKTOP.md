# Identidade visual e ícones Windows

## Estado e aprovação

Integração concluída no código e **aprovada visualmente pelo usuário**, conforme informado na auditoria conjunta de 03/11/2026. Essa aprovação é dos ícones; não equivale a concluir a verificação funcional de **Abrir com** no Icon Test. Não houve commit ou push.

O Computer Use da etapa anterior foi interrompido por Esc. O harness instalou o pacote e iniciou um processo, mas não confirmou a conclusão da abertura do PDF. A inspeção visual posterior do usuário resolve a aprovação da arte; startup/Abrir com/segunda abertura no **Icon Test** permanecem pendentes. A abertura Windows experimental.4 já tinha aprovação manual própria, preservada como histórico separado.

## Recursos definitivos

| Arquivo em `build/icons/` | Formato/conteúdo |
| --- | --- |
| `sources/icone4ps.png` | Original copiado sem alterar bytes, PNG RGBA sRGB, 1077 × 1082 px |
| `sources/iconepdf.png` | Original copiado sem alterar bytes, PNG RGBA sRGB, 1091 × 1239 px |
| `app.png` | Ícone principal, PNG RGBA 1024 × 1024 px |
| `pdf.png` | Ícone de documento, PNG RGBA 1024 × 1024 px |
| `app.ico` | Ícone principal multirresolução |
| `pdf.ico` | Ícone PDF multirresolução |
| `preview.png` | Prancha de conferência em fundos claro/escuro, tamanhos reais de 16 a 128 px |

Ambos os ICOs têm imagens de **16, 24, 32, 48, 64, 128 e 256 px**, com alpha de 32 bits. Quadros até 128 px usam DIB/BGRA com máscara AND; 256 px usa PNG. A prancha tem fundos somente para conferência, não adicionados aos ícones.

A conversão mantém o enquadramento completo e a proporção, centralizando a arte numa tela quadrada transparente. Não recorta, redesenha, acrescenta contorno ou preenche o fundo. Usa reamostragem bicúbica de alta qualidade; medidas fracionárias evitam deformação por arredondamento. Os detalhes naturalmente ficam menores em 16 px, sem desenho alternativo.

Hashes SHA256 das cópias originais, conferidos nos testes:

- Aplicativo: `2001def4aae94153a077f9548be1676185390f8c83142c243212dbb8cfd6eeca`.
- PDF: `5fd6f58cc1a567faa0d5fa5e2ca7c35ac04b74ad66b77e7a5ad0c4e5168c07cc`.

## Geração reproduzível

`scripts/generate-icons.ps1` utiliza `scripts/icons/IconAssets.cs` e System.Drawing do **Windows PowerShell 5.1/.NET Framework**, sem dependência npm nova. Com as cópias originais versionadas, executar na raiz:

```powershell
& "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -NonInteractive -Command "& ([scriptblock]::Create((Get-Content -Raw scripts/generate-icons.ps1)))"
```

Não alterar a política de execução do Windows. Para importar as artes pela primeira vez, o script aceita `-AppSource` e `-PdfSource`; recusa substituir uma cópia original já existente por bytes diferentes. Regerar modifica somente os derivados conhecidos em `build/icons/`. Os arquivos externos do Desktop não são sobrescritos. Em execução fora da raiz, informar `-ProjectRoot`.

## Integração

- `electron-builder.config.cjs`: `win.icon`, `nsis.installerIcon` e `nsis.uninstallerIcon` apontam para `build/icons/app.ico`. O ícone do executável é usado pelos atalhos, menu Iniciar, barra de tarefas e registro do aplicativo.
- `desktop/app.mjs`: a janela usa `build/icons/app.png`, incluído no ASAR com o mesmo caminho relativo. O logo horizontal da interface web permanece preservado.
- `extraResources`: copia somente `build/icons/pdf.ico` para `resources/icons/pdf.ico`, acessível ao Shell fora do ASAR.
- `build/pdf-association.nsh`: o `DefaultIcon` do ProgID próprio usa `"$INSTDIR\resources\icons\pdf.ico",0`; `ApplicationIcon` da capacidade usa o executável. Comandos `--open-pdf`, OpenWithProgids, proteção de desinstalação e escolha de leitor padrão permanecem intactos.
- O Windows mostra o ícone do leitor padrão vigente. Os PDFs só adotarão o ícone do 4PrintShops quando o usuário escolher esse aplicativo como padrão; a instalação não grava UserChoice nem substitui o default de `.pdf`.
- `scripts/check-desktop-package.mjs` confere também o PNG no ASAR e o ICO PDF externo. O kit de laboratório Authenticode inclui os três recursos necessários ao build; certificados, originais e ferramentas de geração não entram no pacote do aplicativo.

## Laboratório e testes

Configuração isolada `electron-builder.icon-test.cjs`: appId `br.com.4printshops.icontest`, produto/perfil **4PrintShops Icon Test**, versão `1.0.0-experimental.7`, saída `dist-desktop/icon-test/`. Não substitui o produto principal, PDF Test ou Assisted Lab.

Instalador gerado na etapa anterior:

`dist-desktop/icon-test/4PrintShops-Icon-Test-1.0.0-experimental.7-Windows-x64-Setup.exe`

- **202 testes Node aprovados** na integração e novamente na auditoria: 199 anteriores + três testes de originais/PNGs, estrutura/alpha dos ICOs e configuração/associação.
- Instalação NSIS isolada conferiu preservação do leitor padrão/UserChoice, registro do ícone PDF e atalhos. Recursos extraídos pelo Windows de executável, Setup, desinstalador e PDF em **16/32 px** coincidiram pixel a pixel com os ICOs. Releitura na auditoria aprovou os mesmos recursos/atalhos, sem reinstalar nem mudar associação.
- O helper trata caminhos longos e equivalentes 8.3; uma divergência inicial do verificador de atalhos foi corrigida no teste, sem defeito no atalho.
- Chrome headless curto e Electron aprovados na auditoria. O Electron teve uma tentativa inicial com timeout na prévia seletiva e passou na repetição isolada, sem mudança de código. Causa do timeout não foi conclusivamente estabelecida; logs preservados.
- O pacote Icon Test teve ASAR auditado; o novo Assisted Lab com ícones também passou com **420 entradas**. Testes do mecanismo assistido com o novo instalador estão registrados em `ATUALIZACOES_ASSISTIDAS.md`.

## Conferência manual ainda recomendada

1. Abrir `build/icons/preview.png`, `app.png` e `pdf.png`. Conferir 16/32 px e transparência; a aprovação visual do usuário já foi recebida.
2. No Icon Test, conferir ícone do aplicativo/atalhos e usar **Abrir com** para um PDF de teste, sem marcar como padrão. Confirmar página, miniaturas e navegação; repetir com a janela aberta e documento modificado, cancelando uma substituição para verificar preservação.
3. Somente se o usuário desejar mudar o leitor padrão, selecioná-lo nas configurações do Windows e conferir o ícone dos PDFs no Explorador. Não limpar caches nem mudar UserChoice por script. O cache do Shell pode demorar a atualizar.
4. O teste completo de Abrir com/segunda abertura nessa identidade continua pendente até execução explícita. Não confundir a aprovação visual com homologação funcional ou com atualização automática.
