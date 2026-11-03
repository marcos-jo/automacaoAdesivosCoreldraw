# Laboratório descartável de Authenticode — 4PrintShops

Análise e preparação em 03/11/2026. **Nenhuma atualização automática assinada foi executada nesta etapa.** O ensaio unsigned anterior é distinto: download real aprovado, instalação pelo updater corretamente bloqueada, seguida de instalação manual complementar da `.3`.

## Viabilidade e ambiente encontrado

É tecnicamente viável confiar em uma identidade autoassinada **somente dentro de um guest Windows descartável**, assinar `.2` e `.3` com a mesma chave e manter todas as validações do aplicativo. Um certificado apenas criado/embutido, mas não confiável para o Windows do guest, não basta: `Get-AuthenticodeSignature` precisa retornar `Valid`, assinatura incorporada e editor esperado. Não há promessa de sucesso antes do ensaio real.

Versões verificadas nos manifests e fontes instalados:

| Componente | Versão | Requisito observado |
| --- | --- | --- |
| Electron | 45.1.0 | Manter runtime e isolamento existentes. |
| electron-builder | 26.18.0 | `win.signtoolOptions`, `certificateSha1`/store ou PFX, `publisherName`, `forceCodeSigning`. Não usar schema 27. |
| electron-updater | 6.8.11 | NSIS, metadados reais/SHA-512 e editor de `app-update.yml`; `verifyUpdateCodeSignature` é callback, não booleano. |

O updater usa também a validação adicional de `desktop/update-signature.mjs`: falha de PowerShell/JSON/timeout, assinatura ausente/não confiável, catálogo ou editor/caminho divergente bloqueiam. Nenhum desses controles é alterado por este laboratório. O servidor HTTP em `127.0.0.1` é a exceção **já existente e exclusiva do pacote Lab**, não uma configuração da produção.

Inventário somente leitura do host: **Windows 10 Pro 19045**, virtualização no firmware e SLAT reportados como ativos, `HypervisorPresent=False`. Não encontrados cmdlets/serviços usuais do Hyper-V, executáveis do VirtualBox/VMware nos caminhos padrão ou `WindowsSandbox.exe`. Isso não exclui instalações em outros caminhos, mas **não há uma VM acessível identificada**. Não foram habilitados recursos, reiniciado Windows, criado/importado certificado, alterado store, política, firewall ou confiança do host.

## Preparar o guest (intervenção manual)

Preferir uma VM Windows 10/11 limpa, descartável, com snapshot anterior à criação dos certificados, 6–8 GiB de RAM se disponíveis e espaço para Node, dependências, caches e três builds. Usar uma instalação Windows de laboratório autorizada, sem contas/credenciais de atendimento. Não clonar o Windows principal. Windows Sandbox é uma alternativa se for disponibilizado separadamente; sua ativação exige decisão do operador e possivelmente reinício, não foi feita aqui.

1. Provisionar VM e Node.js 24; instalar dependências somente no guest. Não habilitar pastas do usuário, discos de clientes, área de transferência, impressoras ou drag-and-drop compartilhados.
2. Transferir **apenas o kit** descrito abaixo por ISO ou compartilhamento somente leitura. Copiar para disco interno do guest, por exemplo `C:\4PrintShopsLab`. Remover o compartilhamento depois. Nunca mapear a raiz do projeto/usuário com escrita.
3. Durante o provisionamento pode ser necessário acesso controlado à internet para dependências/runtime/cache público do builder; usar NAT, sem ponte para a rede da gráfica. Não copiar `.npmrc`, `.env`, certificados comerciais ou tokens. Depois de preparar caches, desconectar a NIC antes de gerar chaves e testar. Servidor e updater comunicam-se por loopback do próprio guest.
4. Registrar snapshot, versão/build do Windows, espaço/RAM, versão de Node e hashes do kit. Se qualquer barreira ou controle corporativo impedir a execução, interromper; não mudar ExecutionPolicy, AppLocker, SmartScreen, Defender ou política de confiança para contornar a falha.

O kit é preparado no host por `node scripts/authenticode-lab/prepare-kit.mjs`, em `dist-desktop/authenticode-vm-kit/`. Contém fontes selecionadas e `lab-kit.json` com SHA-256 de cada arquivo e fingerprint irreversível do host. **Não contém Node, node_modules, cache, Git, PFX/chaves, PDFs ou documentos de clientes.** O preparador recusa sobrescrever um kit anterior. Arquivos gerados permanecem ignorados pelo Git.

Antes de executar os scripts no guest, conferir cada SHA-256 de `source/<file>` contra `lab-kit.json`. Um hash comprova igualdade, não a autenticidade de uma origem desconhecida. O kit deve vir desta sessão confiável. As barreiras host/guest evitam acidentes; não são uma fronteira contra um operador que edite deliberadamente scripts/manifestos.

### Provisionamento dentro da VM

Abrir um PowerShell exclusivo do guest e entrar em `C:\4PrintShopsLab\source` (ajustar ao local copiado). Com rede temporária de provisionamento:

```powershell
npm ci
$env:PRINTSHOPS_LAB_VERSION='1.0.0-experimental.2'
& .\node_modules\.bin\electron-builder.cmd --config electron-builder.update-lab.cjs --win nsis --x64 --publish never
if ($LASTEXITCODE -ne 0) { throw 'Falha no preparo dos caches; interromper.' }
node -e "require('app-builder-lib/out/toolsets/windows').getSignToolPath(undefined,true).then(x=>console.log(x))"
if ($LASTEXITCODE -ne 0) { throw 'Ferramenta de assinatura indisponível; interromper.' }
```

Esse build unsigned é apenas aquecimento de dependências/cache; não comprova assinatura nem deve ser usado no ensaio positivo. A chamada interna de cache foi conferida na versão 26.18.0; se outra versão estiver instalada, parar e revisar. O builder pode precisar de Electron, NSIS e ferramentas de assinatura públicos no cache, mesmo com `ELECTRON_BUILDER_OFFLINE=true`: nessa versão, a variável suprime timestamp/TSA, **não é garantia de bloqueio de rede nem substitui caches**. Desconectar a NIC efetivamente e registrar que a porta 18083 será somente loopback.

## Criar as identidades somente no guest

```powershell
$env:PRINTSHOPS_VM_MANIFEST='C:\4PrintShopsLab\lab-kit.json'
node scripts/authenticode-lab/create-identity.mjs --confirm-disposable-guest
```

O script compara fingerprint com o host e exige modelo de VM reconhecido. Sem confirmação, sem identificação confiável ou na máquina original, aborta antes de criar certificados. Modelo não reconhecido exige revisão do laboratório; não editar o manifesto para contornar.

O comando autoriza no guest duas identidades de 30 dias, RSA 3072/SHA-256, EKU Code Signing e chaves **não exportáveis**:

- A: `4PrintShops Update Lab (TEST ONLY)`, usada nos builds positivos `.2` e `.3`.
- B: `4PrintShops Other Publisher (TEST ONLY)`, usada no negativo de editor diferente.

Chaves ficam em `Cert:\CurrentUser\My` do guest; só certificados públicos `.cer` e thumbprints são gravados em `%LOCALAPPDATA%\4PrintShops Authenticode Lab`. A confiança é inserida em **CurrentUser\Root e CurrentUser\TrustedPublisher somente do guest**. Confirmar eventuais avisos de inclusão de certificado nessa VM. Não exporta chave, PFX ou senha. Se parar parcialmente, descartar/reverter o guest em vez de reaproveitar um estado desconhecido. Não repetir a geração entre `.2` e `.3`: reutilizar `identities.json` e a chave A.

## Build assinado (ainda não executado)

```powershell
$env:PRINTSHOPS_DISPOSABLE_GUEST='YES'
$env:PRINTSHOPS_LAB_IDENTITIES=Join-Path $env:LOCALAPPDATA '4PrintShops Authenticode Lab\identities.json'
$env:ELECTRON_BUILDER_OFFLINE='true'
$env:PRINTSHOPS_LAB_SIGNER='primary'
$env:PRINTSHOPS_LAB_VERSION='1.0.0-experimental.2'
& .\node_modules\.bin\electron-builder.cmd --config electron-builder.authenticode-lab.cjs --win nsis --x64 --publish never
if ($LASTEXITCODE -ne 0) { throw 'Build .2 A falhou; interromper.' }
$env:PRINTSHOPS_LAB_VERSION='1.0.0-experimental.3'
& .\node_modules\.bin\electron-builder.cmd --config electron-builder.authenticode-lab.cjs --win nsis --x64 --publish never
if ($LASTEXITCODE -ne 0) { throw 'Build .3 A falhou; interromper.' }
$env:PRINTSHOPS_LAB_SIGNER='foreign'
& .\node_modules\.bin\electron-builder.cmd --config electron-builder.authenticode-lab.cjs --win nsis --x64 --publish never
if ($LASTEXITCODE -ne 0) { throw 'Build .3 B falhou; interromper.' }
```

Interromper após qualquer código de saída diferente de zero; não prosseguir com artefato anterior. A config nova herda apenas o Lab, exige guest, manifesto, identidade e `forceCodeSigning: true`; mantém `verifyUpdateCodeSignature: true`. Seleciona thumbprint exato e SHA-256. Conferir no log seleção de `CurrentUser\My`, não outro certificado/store. Saídas ficam em `dist-desktop/authenticode-lab/{primary|foreign}/<version>/`, separadas das anteriores e de produção. `.2` e `.3` positivos devem usar **o mesmo thumbprint A**; `.3` negativo usa B. Não assinar o Setup depois de gerados os metadados: deixar o builder assinar executáveis/instalador/desinstalador e gerar os hashes/blockmaps finais.

Por ser offline e de curta duração, este laboratório não usa TSA/timestamp; após expirar, reverter o snapshot e repetir com novas identidades. Isso **não é a política comercial**, que exige cadeia pública apropriada, timestamp e demais requisitos de distribuição. Não reaproveitar variáveis, certificado ou chaves de teste numa sessão de release.

Antes de servir qualquer build, executar `Get-AuthenticodeSignature` no Setup e em `win-unpacked\4PrintShops Update Lab.exe`: exigir `Status=Valid`, `SignatureType=Authenticode`, `SignerCertificate.Thumbprint` correto e EKU de Code Signing. Conferir `publisherName` de `resources\app-update.yml`, `version` de `experimental.yml` e SHA-512/tamanho calculados **depois** da assinatura. Se não retornar `Valid`, parar; não ignorar a validação. Guardar logs públicos, hashes e fingerprints, nunca chaves.

## Feeds de ensaio e critérios de aprovação

Depois das verificações de assinatura dos três builds:

```powershell
node scripts/authenticode-lab/prepare-feed.mjs --confirm-disposable-guest
node scripts/serve-update-lab.mjs dist-desktop/authenticode-lab/feeds/valid
```

O preparador não altera os originais: gera quatro pastas separadas. Reiniciar o servidor com a pasta desejada; só um servidor em `127.0.0.1:18083`. Não publicar na internet. Começar cada cenário com snapshot limpo e instalação `.2` A; o cache não deve confundir casos. **Não usar o harness unsigned `tests/desktop-update-real.mjs` como homologação positiva**, pois ele espera rejeição e instala `.3` manualmente.

| Caso | Procedimento e resultado exigido |
| --- | --- |
| Atualização válida | Instalar `.2` A e servir `valid` (`.3` A); detectar, mostrar novidades, adiar e usar PDF/QR. Retomar download, verificar integridade/assinatura e chegar a pronto sem instalar sozinho. |
| Mais tarde/cancelamento | Com download pronto, escolher Mais tarde na confirmação. App deve permanecer aberto, sem instalação agendada ao sair. |
| PDF não salvo | Importar PDF sintético, girar e ocultar sua aba. Tentar instalar: deve bloquear, manter página/rotação/pendência e liberar controles. Cancelar fechamento também conserva trabalho. Exportar ou fechar a aba conscientemente antes de retomar. |
| Instalar/reiniciar | Sem pendências, autorizar explicitamente instalação. Concluir NSIS, reiniciar, confirmar `.3` **no processo em execução**, não só no arquivo de Setup. Registrar PID, caminho, versão, hashes e retorno. Recarregar preferências e formato personalizado; comparar PDF original por hash. |
| Editor diferente | Servir `foreign`: B precisa ser válido/confiável no guest, mas `.2` espera A. Exigir `ERR_UPDATER_INVALID_SIGNATURE`, nenhuma instalação e PDF preservado. Isso isola o teste de identidade do teste de cadeia não confiável. |
| Hash adulterado | Servir `hash-mismatch`: binário alterado com metadado original. Exigir falha de checksum antes da instalação. |
| Assinatura adulterada | Servir `signature-tampered`: cópia alterada com SHA-512 do feed recalculado deliberadamente. Exigir falha Authenticode, mesmo com hash do transporte compatível. Nunca executar o binário alterado manualmente. |

O gate atual compara identidade de editor/DN, **não fixa thumbprint da chave**. Uma nova chave confiável com o mesmo CN não deve ser descrita como automaticamente rejeitada; trocas legítimas de chave são um tema separado. O negativo acima usa outro CN e ambas as cadeias confiáveis. Não mudar o mecanismo comercial nem acrescentar pinning improvisado para fazer o teste passar.

Para confirmar a versão ativa, abrir o processo instalado e consultar a ponte já existente `window.desktop4PrintShops.estadoAtualizacao()` na instrumentação de laboratório, conferindo `currentVersion`, ou combinar caminho do processo iniciado após o update com ProductVersion do seu executável e evidência visual da execução. Portas de depuração são somente instrumentação temporária da VM; não habilitá-las na produção. Se o reinício perder os argumentos de depuração, reconectar pela nova sessão; isso não é prova de falha/sucesso de instalação por si só.

## Evidências, descarte e pendências comerciais

Registrar uma linha por caso: Windows/Node/versões, thumbprints/subjects públicos, `Get-AuthenticodeSignature`, hashes, escolha do operador, versão antes/depois, logs/capturas e resultado (aprovado/falhou/não executado). **Só aprovar instalação automática após o caso positivo passar de verdade**, incluindo reinício. Não confundir um resultado de mocks ou instalação manual com esse fluxo. Exportar apenas esses relatórios públicos por um canal controlado; não exportar perfil da VM, certificados com chave, registro ou cache contendo segredos.

Ao terminar, encerrar servidor/app e **descartar toda a VM ou restaurar o snapshot anterior aos certificados**. Não exportar a chave para facilitar novo ensaio, não promovê-la a release e não confiar esses certificados no Windows principal. Se precisar preservar evidências, usar diretório de saída dedicado, nunca compartilhar os documentos da gráfica.

Pendências nesta sessão: disponibilizar VM/Windows Sandbox, provisionar dependências/cache, criar confiança só no guest, assinar os três builds, executar todos os casos da tabela e revisar evidências. Nenhum desses passos de assinatura/instalação foi executado no host. O teste de laboratório não confere reputação SmartScreen, cadeia pública, revogação, timestamp de longo prazo, HTTPS/provedor real, rotação de certificados ou homologação em máquinas de atendimento. Distribuição comercial continua exigindo essas decisões e aceite separado.

Validações efetivamente realizadas nesta preparação: inventário do host em leitura; 176 testes Node aprovados (174 anteriores mais dois de barreiras do guest); sintaxe dos seis JS/MJS/CJS novos; teste real da barreira no próprio host sem invocar ações de certificado; kit com 66 arquivos e todos os SHA-256 conferidos; auditoria do ASAR `.3` anterior ainda idêntico ao runtime fonte (410 entradas). Nenhum teste de certificado, build assinado, confiança no guest, rejeição com certificados reais ou instalação/reinício automático foi executado nesta etapa. Não se repetiram testes de navegador por não haver mudança em runtime/frontend/configuração de produção.

## Fontes primárias

- [Microsoft: certificados autoassinados para testes](https://learn.microsoft.com/en-us/powershell/module/pki/new-selfsignedcertificate).
- [Microsoft: confiança de certificados autoassinados](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_signing).
- [Microsoft: isolamento/configuração do Windows Sandbox](https://learn.microsoft.com/en-us/windows/security/application-security/application-isolation/windows-sandbox/windows-sandbox-configure-using-wsb-file).
- [electron-builder v26: opções Windows](https://www.electron.build/v26/docs/win/).
- As fontes locais `windowsSignToolManager.js`, `NsisUpdater.js` e `desktop/update-signature.mjs` foram conferidas; não presumir comportamento de builder 27/updater 7.
