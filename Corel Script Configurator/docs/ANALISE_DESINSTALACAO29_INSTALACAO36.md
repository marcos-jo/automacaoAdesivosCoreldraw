# Análise da rota .29 → desinstalação manual → instalação limpa .36

Data: 10/11/2026. Base: main/37c4a11, código 4efcc56. **Desinstalação manual .29 e instalação limpa .36 concluídas; verificações técnicas aprovadas na estação.** Operador aprovou parcialmente a importação Word em cenário real simples, incluindo montagem e exportação; documentos complexos e homologação física continuam pendentes. Os registros de preparação abaixo são históricos.

## Encerramento documental — relato do operador em 10/11/2026

Word real com somente “teste importação word”: importação um pouco demorada, mas concluída, abertura correta, montagem e exportação aprovadas pelo operador. **Homologação parcial em cenário simples**, sem estender a aprovação a DOC/DOCX completos, documentos complexos, fontes, imagens, tabelas, desempenho, ICC/RIP ou impressão física. Extensão, modo de montagem e formato da exportação não foram informados. Nenhum novo teste foi executado para registrar este relato. A instalação foi limpa depois da retirada da .29, não migração transacional .29→.36. Próxima prioridade registrada: importação segura de PDFs protegidos, impedindo travamentos em lotes mistos, sem iniciar essa implementação. Ver [relatório curto](HOMOLOGACAO_PARCIAL_EXPERIMENTAL36.md). Evidências e backups existentes preservados fora do Git; commit apenas documental autorizado, sem push automático.

## Instalação limpa .36 na gráfica — concluída em 10/11/2026

Após a conferência da remoção, o operador autorizou explicitamente uma execução visível do Setup local de SHA256 `FE6E93BC632784FF08D340D55CCB3339A2763AB633A797426C986BFF021B7056` (514.968.613 bytes). Hash, tamanho, ausência do alvo/registro anterior, perfil e disponibilidade do backup externo reconferidos antes de iniciar. **Uma única execução, sem instalação silenciosa ou repetição.** Não houve migração direta .29→.36: a .29 já havia sido removida manualmente e o destino estava ausente.

Leitura dos controles do Windows confirmou a versão e o destino correto; captura de imagem e clique automatizado falharam na ferramenta, sem alteração do instalador. Operador clicou em Instalar e depois em Concluir. Tela de sucesso registrada por acessibilidade; **código de saída0 capturado nesta instalação**, diferente do código indisponível da remoção .29. Abertura do Setup15:58:05 e encerramento16:11:55, horário de Brasília; aproximadamente13min50, **incluindo espera e interação nas telas**, não uma medição isolada de cópia.

Resultados independentes:

- Diário `%LOCALAPPDATA%/Programs/.4ps-upgrade-RB0wY7/journal.json` em **committed**, `previous:null`, `oldExists:false`, sem falha. Cópia do diário preservada nas evidências.
- Versão do ASAR, FileVersion do executável e registro de desinstalação **1.0.0-experimental.36**. Mesmo appId `br.com.4printshops.holdtest`, GUID `90ff2ff0-f013-571c-b109-7edb265dacc8`, destino `%LOCALAPPDATA%/Programs/4PrintShops Hold Test` e perfil `%APPDATA%/4PrintShops Hold Test`.
- **19.510 arquivos do pacote conferidos por tamanho/SHA256, zero ausentes ou diferentes**. Total instalado19.513: payload mais desinstalador, uninstallerIcon.ico e inventário seletivo 4ps-installed-files.ini. Os dois auxiliares conferidos contra o diário; inventário próprio coerente com a identidade e os19.513 arquivos, sem extras desconhecidos.
- Executável instalado em modo `--verify-distribution`: saída0 em19.692ms; **LibreOffice privado26.8.2.2,19.410 arquivos íntegros e inicialização real aprovada**, usando perfil temporário, sem abrir o perfil do operador.
- `scripts/test-private-office.mjs` contra a instalação real: saída0, passed=true, aproximadamente108,2s para **o teste completo**, não só a primeira conversão. DOCX e DOC sintéticos convertidos, duas páginas A4 em cada,210,01×297mm,105 caracteres conferidos por resultado, acentos/frases preservados e fontes detectadas. Descoberta de Office global desabilitada; executável efetivo `lo/program/soffice.com`. Originais intactos por bytes e integridade do runtime reconferida após uso; cancelamento antecipado e recusa de runtime ausente também passaram no teste existente.
- Antes da primeira abertura normal, **76 arquivos do perfil idênticos por SHA256 ao baseline**. Leitura adicional em cópia descartável confirmou exatamente as três preferências anteriores: tema light, painéis320/230 e print-config-check=preservado. Perfil real não foi aberto/modificado pelo leitor auxiliar.
- Aplicativo instalado aberto normalmente e janela `app://4printshops/index.html` conferida por acessibilidade. PDF sintético resultante da conversão aberto pelo comando `--open-pdf` do executável instalado, na janela existente: status PDF carregado, duas páginas A4, duas miniaturas e controles do editor presentes; divisores320/230 confirmados na interface. Nenhum cliente/PDF real utilizado. Isso testa o comando usado pela associação, **não um novo teste manual de Abrir com no Explorador**.
- Uma única entrada Hold Test nas áreas usuais HKCU/HKLM32/64, dois atalhos apontando ao executável correto. Associações próprias presentes. UserChoice e padrão .pdf iguais ao baseline inclusive após abrir o aplicativo; nenhuma alteração manual do leitor padrão.

Avisos/limitações: PDF.js emitiu avisos de build Node e `TT: undefined function:32`; não impediram texto/dimensões, mas não homologam aparência de todas as fontes. O log redirecionado pelo Windows PowerShell era UTF-16 e foi decodificado corretamente para ler o resultado, sem repetir conversões ou Setup; NativeCommandError no texto do log corresponde aos avisos em stderr, com saída final0. A verificação da janela foi textual porque a captura visual falhou nesta sessão. Documentos Word reais, substituição de fontes/paginação, ICC/RIP, impressão física e teste de desinstalador bloqueado/VM continuam pendentes. Nenhuma proteção foi desabilitada.

Operador observou LED de atividade do SSD quase continuamente aceso. Confirmados avanços nos contadores de leitura/gravação do processo e no diário. Isso indica atividade durante extração/cópia/hashes, **não uma medição de100% de utilização ou diagnóstico da saúde do SSD**. Não foram alteradas políticas, antivírus, drivers, impressoras ou configuração de desempenho.

Evidências privadas ignoradas em `.desktop-test/supervised36-install-20261110-155805/`, especialmente `setup-result.json`, `committed-journal.json`, `installed-package-audit.json`, `installed-verify-distribution-result.json`, `installed-private-word-result.json`, `preferences-after-install.json`, `registration-final.json`, `installed-startup-ui.txt`, `installed-synthetic-pdf-ui.txt` e **final-result.json**. Setup, backups e documentos sintéticos permanecem fora do Git. Aplicativo deixado aberto com o PDF sintético, sem edições. Somente documentação versionada atualizada, sem código de produção, restauração, limpeza manual, impressão, commit ou push.

## Resultado real da desinstalação manual .29 — 10/11/2026

Operador informou que o desinstalador concluiu, sem perguntar sobre manter dados/configurações. Essa ausência de pergunta é compatível com `deleteAppDataOnUninstall:false` da configuração .29; a preservação foi comprovada separadamente pelos hashes do perfil.

Conferência posterior somente de leitura:

- Diretório `%LOCALAPPDATA%/Programs/4PrintShops Hold Test` **ausente**, sem arquivos ou diretórios residuais no alvo. Nenhuma limpeza manual necessária ou executada.
- Chaves próprias de instalação/desinstalação, ProgId, Applications, Capabilities e valores próprios OpenWithProgids/RegisteredApplications ausentes, como esperado após desinstalação. Busca adicional nas áreas usuais HKCU/HKLM,32/64bits: zero entradas de desinstalação Hold Test.
- Os dois atalhos próprios foram removidos. Leitor PDF padrão/UserChoice e padrão .pdf permaneceram iguais ao baseline, sem escrita nesses valores.
- **76 arquivos do perfil idênticos ao baseline por nome, tamanho e SHA256**, sem faltas ou arquivos adicionais. As três preferências previamente lidas permanecem preservadas pelos mesmos bytes: tema light, painéis320/230 e print-config-check=preservado. Não foi aberto outro aplicativo para regravar o perfil.
- Backup local da instalação com1.220 arquivos, perfil e cache conferidos novamente por SHA256; atalhos de backup íntegros. Backup externo previamente validado com1.502 arquivos: recibo/manifesto reconferidos iguais e tamanhos dos arquivos reconferidos; **não repetida nesta etapa a leitura integral dos hashes externos**, já aprovada ao concluir a cópia.
- Setups .29 de recuperação e .36 local disponíveis e hashes novamente conferidos. Nenhum processo de aplicativo/conversor/Setup/desinstalador encontrado nas consultas finais. Nenhum diário transacional pendente para o destino.
- **Código de saída indisponível:** o monitor somente de leitura não observou o processo usado na remoção e foi encerrado pelo próprio marcador de parada. Não inferir saída0. Não foram capturadas mensagens de erro; o relato do operador é de conclusão, não de uma transcrição integral das telas. Não repetir desinstalação para obter um código.

Evidências privadas identificadas pelo ponteiro `.desktop-test/supervised36-post-uninstall-current.json`, que informa a pasta efetiva, com result.json, registry-after.json e profile-after-inventory.json. Monitor em `.desktop-test/supervised36-uninstall-20261110-152046/`. Esses materiais permanecem ignorados e não incluem autorização de nova instalação.

**Condições técnicas para uma tentativa supervisionada de instalação limpa .36 atendidas.** Não há resíduos a autorizar ou registros antigos conflitantes. Ainda aguarda autorização explícita para uma única execução do Setup local auditado FE6E93BC…21B7056. Preservar destino/perfil/identidade, não restaurar associações manualmente. Isso comprova o caso real de remoção manual na estação; não homologa comportamento com arquivo bloqueado nem equivale a ensaio do desinstalador em VM. .36 ainda não instalada nem homologada para produção nesta estação. Nenhum código funcional alterado, restauração, impressão, commit ou push.

## Mudança de estratégia pelo operador — 10/11/2026

O operador decidiu **não criar nem executar VM nesta estação de produção**, por recursos limitados, e autorizou preparar a rota manual supervisionada .29 → instalação limpa .36. Essa decisão substitui a recomendação operacional de aguardar VM nas seções históricas abaixo; **não significa que o desinstalador original foi homologado em VM**.

Condições atuais: nenhuma alteração de código, nenhum novo ensaio de laboratório e nenhuma execução automática de desinstalação/instalação. Operador acompanha e executa a desinstalação pela interface normal do Windows. Antes disso, backup externo íntegro e Setup original .29 disponível para recuperação. Depois da remoção, conferir arquivos/diretórios, Registro, associações próprias, leitor padrão, perfil/preferências e processos; resíduos exigem classificação e autorização específica, sem limpeza genérica. Setup .36 depende de outra autorização após apresentar a conferência.

Backup local e instalação/perfil foram novamente conferidos iguais aos inventários, fora do alvo de desinstalação; Setup .29 original de126.424.116 bytes reconfirmado por SHA256. O operador autorizou a unidade removível D: para a segunda cópia. **Cópia concluída e conferida**, em pasta nova `D:/4PrintShops-Backups/pre36-20261110-134755-copy-2026-11-10T18-07-16-207Z/`:1.502 arquivos/1.415.012.505 bytes, incluindo os1.501 arquivos do baseline completo e o Setup .29 separado para recuperação. Cada origem/destino conferidos por tamanho/SHA256, origem reconferida após cópia, zero divergências. Manifesto externo `external-copy-inventory.json`, SHA256 `F8B177F867854020F34F5DF54E0047E6D2C19570DF486F0F338D90C5BF0CF077`, e receipt.json gravados ao final; ponteiro local `.desktop-test/supervised36-external-backup-current.json` com status verified. Mídia FAT32 preserva nesta cópia os bytes conferidos; não afirmar preservação de ACLs/recursos NTFS. Nenhum arquivo existente da mídia foi substituído ou excluído.

Após a cópia, última conferência novamente aprovada: .29 instalada,1.220 arquivos e76 do perfil iguais aos backups, sem arquivos novos/links, Registro/atalhos/default iguais, Setup .36 íntegro e nenhum aplicativo/conversor/Setup/desinstalador detectado. Monitor somente leitura preparado para observar processos do desinstalador original e códigos de saída quando capturados; não lança nem fecha programas. Operador será orientado a usar appwiz.cpl e selecionar exclusivamente a entrada Hold Test .29. **Desinstalação ainda não executada neste registro; aguardando ação manual.** Se aparecer aviso/erro, preservar a tela e informar sem Repetir/Continuar. Conferência posterior e classificação de resíduos obrigatórias; não executar Setup .36 automaticamente.

## Preparação do ensaio com o binário original — 10/11/2026

O usuário autorizou o teste real em Windows descartável. **Ele ainda não foi executado: falta um ambiente Windows isolado acessível nesta estação.** A autorização do ensaio não foi usada para executar os binários no host de produção.

Consulta local: Windows 10 Pro/build19045, aproximadamente8GiB de RAM, virtualização de firmware/SLAT disponíveis e `HypervisorPresent=False`. Não foram encontrados comandos/instalações usuais de Hyper-V, Windows Sandbox, VirtualBox, VMware ou QEMU nos locais consultados. Não há ambiente equivalente provisionado identificado; isso não exclui ferramentas em locais não consultados. O relatório VM anterior refere-se ao computador de casa, não a uma VM presente nesta estação. Não foram habilitados recursos, alteradas políticas, instalado virtualizador ou solicitado reinício. Consultado o operador sobre usar a VM de casa, disponibilizar outra ou avaliar a preparação local.

Materiais privados preparados e conferidos em `.desktop-test/vm29-original-kit-9SeBFE/`, **não um resultado de homologação**:

- Setups .29 original e .36 local auditado, com hashes indicados neste relatório.
- Os1.119 arquivos adicionais `lo`, copiados do runtime preparado conhecido, conferidos com o manifesto confiável e o inventário anterior. Nenhuma cópia do perfil ou Registro de produção foi incluída.
- Inventários de referência dos arquivos .29 e dos19.510 arquivos .36; plano de destino sintético com raiz63 caracteres e identidade Hold Test.
- DOC e DOCX sintéticos provenientes do teste privado anterior; o DOC tem assinatura binária OLE válida. Não são arquivos reais de clientes.
- README com sequência de instalação .29 original, preferências reais sintéticas, checkpoint, desinstalação original normal, ramo com bloqueio real, inspeção de resíduos e instalação limpa .36 somente se houver remoção comprovada.
- Verificador PowerShell somente de leitura. Sua chamada local foi **recusada pela política de execução**; não foi contornada nem alterada. Análise sintática passou. Auditoria independente com Node conferiu todos os arquivos por tamanho/SHA256, sem executar PowerShell, instalador ou desinstalador.

Conferidos1.129 arquivos de materiais, total878.572.023 bytes, mais o manifesto `kit-files.json`. Zero arquivos adicionais fora do inventário. SHA256 do manifesto, a conferir separadamente no destino: `3D99F1E458CA368B07D3523DD30343279F8078CB6403EC8AF84D589EE7C2C5DB`. Esse digest vincula o transporte ao recibo local; não substitui assinatura comercial ou autenticação externa. Kit, binários, inventários e evidências continuam ignorados pelo Git. Ponteiro/recibos em `.desktop-test/vm29-original-kit-current.json`, `vm29-original-material-audit.json` e `vm29-original-environment.json`.

Preservação do host novamente conferida: versão .29;1.220 arquivos instalados e76 arquivos de perfil iguais aos backups, sem arquivos extras ou links; Registro/atalhos/leitor padrão iguais ao baseline. SHA256 do Setup .36 reconfirmado. Apenas C: disponível na consulta: **a segunda cópia externa dos backups ainda não foi realizada**.

**Situação para decisão:** não declarar o desinstalador original .29 homologado; ainda faltam os dois ramos reais na VM, a inspeção de arquivos/diretórios/Registro e a instalação limpa .36 com conversões/aplicativo. Não avançar para a intervenção na gráfica com base somente no kit ou nos ensaios NSIS neutros. A opção de usar a VM já existente em casa evita preparar um hipervisor no Windows de produção.

## Conclusão

A rota é tecnicamente viável, com a mesma identidade Hold Test e o perfil preservado fora da pasta do programa. A instalação atual tem caminhos curtos, sem evidência de que o defeito de caminhos longos vá ocorrer neste destino. Entretanto, **a desinstalação manual antiga não é transacional nem deve ser considerada incondicionalmente segura**: pode remover parcialmente os arquivos e as associações próprias, mesmo retornando sucesso. A alternativa exige autorização específica, backup conferido e uma barreira de inspeção entre desinstalar e instalar.

O laboratório desta etapa reproduziu a lógica de remoção NSIS em executáveis neutros. **Não foi executado o desinstalador original completo em VM**. Para homologar integralmente esse binário, falta um ensaio em Windows descartável, sem compartilhar o Registro, o perfil e a instalação da estação. Os testes anteriores da VM .34→.36 não substituem esse teste .29.

## Estado real e backups

- Instalação: `%LOCALAPPDATA%/Programs/4PrintShops Hold Test`, raiz com 63 caracteres; versão .29 confirmada por Registro, executável e ASAR. Nenhum aplicativo/conversor/Setup/desinstalador ativo nas consultas finais.
- 1.220 arquivos, 624.748.184 bytes. Todos os 99 arquivos do payload original .29 estão presentes e iguais por SHA256; desinstalador original também conferido. Dois arquivos são auxiliares NSIS e 1.119 estão em `lo`.
- Maior caminho atual: 127 caracteres, em `lo/share/extensions/dict-ru/pythonpath/lightproof_opts_ru_RU.py`; zero caminhos com 260 caracteres ou mais. Não há fundamento para atribuir uma futura falha neste destino atual a MAX_PATH apenas pelo histórico da .34.
- Backups: `.desktop-test/pre36-20261110-134755/`, fora do diretório de instalação e ignorados pelo Git. Instalação completa, perfil com **76 arquivos/8.458.106 bytes**, cache do Setup .29, dois atalhos e snapshot tipado das configurações próprias/associações foram reconferidos. Arquivos atuais e cópias permanecem iguais; inventários finais sem arquivos adicionais nem links/junções.
- O backup está no mesmo volume C:. É adequado para recuperação da desinstalação, mas **não protege contra falha do disco**. Uma cópia externa verificada é recomendável antes da intervenção em produção; não foi feita automaticamente.
- Preferências preservadas: tema `light`, painéis esquerdo/direito 320/230 e indicador de conferência de impressão `preservado`. Perfil não foi aberto pelo aplicativo real nesta análise.
- Leitor PDF padrão continua sendo o ProgId `br.com.4printshops.icontest.PDF`; entradas próprias Hold Test em Abrir com estão presentes. Registro, atalhos e perfil permaneceram iguais ao baseline. Nenhum valor UserChoice foi escrito.

## Os 1.119 arquivos adicionais de `lo`

Todos correspondem, por caminho e SHA256, ao runtime privado conhecido do LibreOffice 26.8.2.2. Zero arquivos desconhecidos ou diferentes nesse conjunto. Manifesto conferido contra a referência confiável do projeto: `01f0286083b702da3c085ec7219c22b98db2d86aa8ae3f6c3633f6707537247e`.

Esses arquivos **não pertencem ao payload original .29**. A igualdade dos bytes identifica o runtime, mas não comprova qual tentativa anterior ou cópia os introduziu. Não atribuir sua origem a uma instalação específica sem evidência adicional. O conjunto é parcial; `lo/program/version.ini` está ausente e não foi considerado um LibreOffice utilizável.

A remoção manual antiga apaga recursivamente a pasta do aplicativo, inclusive esse `lo` e qualquer dado desconhecido que alguém tenha colocado ali. O ensaio com um marcador sintético comprovou essa característica. Nenhum arquivo atual foi excluído; todo o conjunto está no backup. Antes de uma autorização futura, repetir o inventário e interromper se surgirem arquivos novos, alterados ou desconhecidos.

## Desinstalador .29 e caminhos longos

O Setup .29 original, SHA256 `FCAA3F9DBC9909F16D44EBC761A4404A3DED59237B1D6E125F6AD8FB79EF5E92`, foi somente extraído. O desinstalador instalado corresponde ao original extraído, SHA256 `0D2E006D1AAF52DC00E3C2006168E084D362D29788020958446A8BBBDE50A94B`.

A configuração histórica do commit e6f605b e o lockfile confirmam electron-builder 26.18.0, `deleteAppDataOnUninstall:false` e ausência de uma rotina própria de remoção. O template correspondente foi conferido por digest antes do laboratório.

- Em modo de atualização (`--updated`), o template usa `un.atomicRMDir`, baseado em renomeações, e tenta restaurar os arquivos se falhar. Compartilha a classe de defeito identificada na .34.
- Na desinstalação manual normal, não usa essa rotina: executa `RMDir /r $INSTDIR`. Também apresenta limitações com caminhos longos e arquivos bloqueados, e não interrompe explicitamente o fluxo imediatamente após um erro dessa remoção.
- O macro histórico .29 remove as associações próprias **antes** da remoção dos arquivos. Não oferece restauração transacional dessas entradas.
- Não usar `--delete-app-data`. O perfil deve permanecer fora da pasta removida. Não transformar a operação manual em atualização com `--updated`.

O binário original não foi redirecionado para uma pasta de teste neste Windows: a inicialização multiusuário pode consultar a identidade real no Registro e resolver o destino ativo. Usar apenas `_?=` não constitui isolamento suficiente para autorizar esse risco.

## Laboratório executado

Executáveis NSIS neutros, sem acesso ao Registro real, sem atalhos reais e sem APIs de impressão. Copiaram os arquivos do backup verificado para raízes temporárias exclusivas. A rotina `un.atomicRMDir`/restauração foi extraída do template correspondente; a remoção manual e a ordem dos efeitos foram reproduzidas. Associações e metadados foram representados por arquivos marcadores externos. Portanto, os resultados desses marcadores **não são um teste real de registro/desregistro do aplicativo**.

| Caso | Maior caminho | Resultado observado |
| --- | ---: | --- |
| Manual, cópia equivalente ao destino atual | 127 | Saída 0; todos os 1.221 arquivos removidos, incluindo marcador sintético interno. |
| Manual, `app.asar` realmente bloqueado | 127 | Saída 0, mas um arquivo permaneceu; marcadores de associações/metadados removidos. Confirma necessidade de inspecionar resíduos. |
| Manual, estresse sintético | 260 | Saída 0, mas um arquivo de caminho longo permaneceu; mesmos efeitos parciais nos marcadores. |
| Atualização, estresse sintético | 260 | Falha com saída 2; 1.222 arquivos restaurados iguais; marcador de associação já removido. |

Marcadores externos de perfil e leitor padrão permaneceram intactos nos quatro ensaios. Cada alvo foi validado como descendente da pasta temporária exclusiva antes da remoção. Nenhum binário original do produto foi executado nesses ensaios. Resíduos e logs do laboratório foram preservados.

Também executados: **14/14 testes Node de `tests/update-transaction.test.mjs`**, incluindo instalação limpa, rollback, preservação de arquivos e bloqueio real Windows. São testes de componente com Registro simulado, não uma nova execução do Setup instalado. Sintaxe dos três auxiliares locais e `git diff --check` conferidos.

Evidências privadas ignoradas: `.desktop-test/uninstall29-analysis-20261110-143618/`, com `verification.json`, `lo-classification.json`, `neutral-results.json`, `transaction-tests.log`, `final-verification.log` e `final-readonly.json`. A localização das cópias neutras está em `neutral-lab-location.json`. Não publicar backups, capturas de configurações ou perfis.

## Compatibilidade da instalação limpa .36

Configuração existente: appId `br.com.4printshops.holdtest`, GUID `90ff2ff0-f013-571c-b109-7edb265dacc8`, produto/executável `4PrintShops Hold Test` e mesmo perfil do usuário. Não há necessidade de outra identidade ou outra instalação paralela.

O migrador continua aceitando apenas versões registradas .34/.35/.36. O código atual permite instalação limpa quando não existe versão anterior registrada e o diretório está ausente ou vazio. **Se sobrar arquivo no alvo após desinstalar, a instalação limpa será recusada**; não remover guardas, normalizar o Registro ou apagar resíduos automaticamente.

Artefato local já auditado e novamente conferido nesta etapa:

```text
dist-desktop/hold-native-test/g36-20261110-130256/
4PrintShops-Hold-Test-1.0.0-experimental.36-Windows-x64-Setup.exe
514.968.613 bytes
SHA256 FE6E93BC632784FF08D340D55CCB3339A2763AB633A797426C986BFF021B7056
```

Payload e entrega em diretório longo com 19.510 arquivos, Office privado com 19.410 arquivos e conversões DOC/DOCX foram auditados na etapa anterior, sem execução desse Setup na estação. O hash doméstico identifica outro artefato. Não confundir instalação limpa com migração direta suportada da .29. A .36 continua sem homologação de produção nesta estação.

## Procedimento proposto — depende de nova autorização

1. Imediatamente antes da intervenção, fechar o aplicativo normalmente e reconferir processos, hashes e inventários atuais/backup, preferências e associações. Preservar este backup; guardar uma segunda cópia externa verificada, se disponível. Interromper se houver divergências.
2. Com autorização explícita para desinstalar, executar **uma única vez** o desinstalador registrado da .29, na interface normal do Windows e para o usuário atual. Não selecionar remoção de dados; não passar flags de exclusão de perfil ou de atualização. Registrar horário, telas e saída. Não tratar código 0 como prova suficiente.
3. **Parar para inspeção antes de instalar**: conferir processos, destino vazio/ausente, remoção das chaves próprias de instalação/desinstalação e atalhos, perfil e leitor padrão intactos. A ausência temporária de Hold Test em Abrir com é esperada após uma desinstalação concluída. Registro inconsistente ou resíduos exigem análise; não executar Setup nem apagar dados para contornar a recusa.
4. Após essa conferência e autorização de instalação, usar exclusivamente o .36 auditado, escolhendo o mesmo diretório/perfil Hold Test. Não criar outro destino ou identidade. Registrar logs/diário transacional.
5. Conferir versão .36, diário committed, todos os arquivos/hashes do inventário aplicável, LibreOffice privado em `lo` e inicialização real. Conferir conversões DOC/DOCX sintéticas, preferências, atalhos, Abrir com e leitor padrão. Só depois iniciar avaliação manual; sem impressão física automática.

A recomendação anterior de exigir o ensaio original .29 na VM foi substituída pela decisão explícita do operador de seguir com remoção manual supervisionada, conforme a seção mais recente; o binário original continua sem homologação em VM. A segunda cópia deve usar outro disco/dispositivo, em pasta nova dedicada, sem substituir backups anteriores e sem opções de espelhamento/exclusão. Conferir por SHA256 o conteúdo integral copiado, incluindo inventários, instalação, perfil, atalhos, cache e snapshots próprios; registrar destino e contagens. Não considerar duas pastas no mesmo C: como duas cópias independentes. A autorização da mídia D: e o resultado da transferência devem ser registrados na seção mais recente.

Se o ensaio real encontrar resíduos, emitir primeiro uma lista por caminho absoluto, tamanho e SHA256. Somente arquivos idênticos às referências .29/`lo`, contidos no alvo correto, sem links e sem bloqueio, podem ser candidatos a uma remoção seletiva **após autorização específica**. Arquivo modificado, desconhecido ou diretório não vazio fora dessa lista deve ser preservado. Não há limpeza genérica nem remoção seletiva executada ou homologada por antecipação.

## Recuperação proposta — não executada

- Se a desinstalação falhar: interromper e preservar destino, registro, perfil, logs e backups. Pode haver remoção parcial do programa e das associações. Não repetir, forçar exclusões ou lançar a .36 sobre resíduos. Não abrir um aplicativo parcialmente removido para continuar produção.
- Se a instalação limpa .36 falhar: preservar diário e registros e conferir rollback. O rollback de uma instalação limpa **não recupera a .29 já desinstalada**; o estado anterior dessa transação era sem aplicativo.
- Recuperar a .29 exigirá autorização separada e um plano baseado no estado observado: preservar por quarentena os resíduos da tentativa, conferir o backup integral e repor a instalação no destino original vazio, com restauração estritamente das configurações próprias/atalhos necessários. Alternativamente, avaliar o Setup .29 original conferido. Nenhuma dessas ações está automatizada ou autorizada por esta análise.
- Nunca importar cegamente todo o snapshot: UserChoice e padrão .pdf são somente referências de comparação; não restaurá-los automaticamente nem tocar em registros de outros leitores. Não sobrescrever um perfil mais recente com o backup sem verificar diferenças e obter autorização específica.
- Validar hashes do programa recuperado, Registro, atalhos e preferências antes de reabrir. Guardar todas as evidências, sem excluir a tentativa malsucedida para obter um estado aparentemente limpo.

Não houve alteração de código funcional, políticas, drivers, impressoras ou preferências. Nenhuma impressão, commit ou push. A decisão operacional permanece com o usuário após conhecer as limitações; homologação integral do desinstalador original .29 em VM continua sendo a opção de maior segurança antes da intervenção real.
