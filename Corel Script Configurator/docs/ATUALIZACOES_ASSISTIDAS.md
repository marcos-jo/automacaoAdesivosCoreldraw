# Atualizações assistidas na rede interna — laboratório

Registro de 03/11/2026. Implementação baseada no checkpoint `c450fed`. Electron 45.1.0, electron-updater 6.8.11 e electron-builder 26.18.0, sem novas dependências npm. **Somente laboratório local; servidor da gráfica e computadores de atendimento não foram configurados.**

## O que está implementado

- Consulta uma vez após abrir a interface e manualmente em Arquivo → **Atualização assistida pela rede** (menu de uma aba PDF/QR). Não aguarda a rede para mostrar o aplicativo.
- Aviso compacto com versão, novidades em texto, **Mais tarde**, download com progresso e **Preparar instalação**. Adiar não inicia download; o menu permite retomar.
- Preparo confere novamente validade/integridade e consulta o trabalho de **todas as abas**, inclusive ocultas, antes e depois da confirmação. PDFs pendentes, operações, seletor/gravação/fila externa de PDFs e instalação integrada impedem a preparação. Não há opção de descartar PDFs nesse fluxo: exportar primeiro.
- Com autorização, apenas `shell.showItemInFolder()` localiza o instalador no Explorador. **Não executa, não reinicia e não instala silenciosamente.** O operador exporta os PDFs, fecha o aplicativo e executa conscientemente o arquivo.
- Falha de validação bloqueia a oferta/preparação. Erro de rede não impede trabalhar. Consulta não envia documentos nem preferências ao servidor.
- Fluxo separado do atualizador integrado. `desktop/updates.mjs` e `desktop/update-signature.mjs` não foram alterados; não existe fallback de uma rejeição Authenticode para execução assistida. Instalação/reinício automáticos continuam **NÃO homologados**.

## Confiança e limites

O servidor entrega um envelope JSON: `keyId`, `payload` em base64url e `signature` Ed25519. A assinatura cobre os **bytes exatos** do payload. O aplicativo incorpora somente a chave pública em seu pacote ASAR; nunca aceita uma chave recebida do servidor ou da configuração de origem.

Depois de autenticar a assinatura, valida schema, produto, canal, Windows x64, SemVer, sequência crescente, emissão/expiração, notas e caminho `releases/<versao>/<nome-seguro>.exe`. O manifesto autentica tamanho e SHA512 do instalador. Download em streaming, limite 512 MiB, sem redirecionamentos/compressão, prazo limitado e arquivo `.part` exclusivo; só renomeia após conferir tamanho/hash. Confere novamente antes de localizar.

Maior sequência/digest vistos ficam no perfil local, com escrita atômica. Recusa downgrade, replay inferior e reutilização de sequência com bytes diferentes. Estado local corrompido falha fechado. O relógio deve estar correto: validade padrão 14 dias, máximo 31, tolerância de emissão 5 minutos. Republicar o mesmo instalador inalterado renova a validade com **nova sequência**; não exige nova instalação.

HTTP é permitido somente quando explicitamente autorizado para **este fluxo assinado**. Protege autenticidade/integridade do artefato, mas não confidencialidade, disponibilidade ou ocultação de versões/endereço. Um servidor comprometido pode negar/atrasar atualizações; uma publicação válida ainda dentro da validade pode ser repetida em um computador que nunca viu sequência superior. Não é um sistema completo de transparência/TUF. HTTPS continua preferível quando houver confiança TLS administrada, sem ignorar erros de certificado. A confiança TLS do Node usado pelo cliente precisa de validação específica; não presumir equivalência com o repositório de certificados Windows.

**Assinatura de publicação não é Authenticode.** Os dois instaladores de laboratório não têm assinatura Authenticode; o Windows pode mostrar editor desconhecido/SmartScreen. O fluxo integrado continuará recusando instaladores incompatíveis. Depois de revelar um arquivo no Explorador, o aplicativo não controla sua execução externa nem alterações posteriores no disco: ele não é um broker de execução. Salvar/fechar é obrigação explícita antes da execução manual. Malware com acesso ao usuário publicador ou à instalação local está fora dessa proteção.

## Chave privada do publicador

`scripts/assisted/keys.mjs` gera Ed25519 e protege PKCS8 usando **DPAPI CurrentUser**. A chave em claro só existe em memória/pipe local durante geração/assinatura; não é passada em argumento ou escrita em PEM. A nova pasta local tem ACL explícita somente do publicador e SYSTEM. Disco de rede/UNC e pasta dentro do Git são recusados. Não se cria nem instala certificado Windows.

- `publication.dpapi`: segredo criptografado, **fora do Git, do webroot e do servidor**, no computador publicador. Não compartilhar a conta publicadora.
- `public-trust.json`: somente produto, canal e chave pública/fingerprint. Pode ser copiado para o computador de build; a chave privada não entra no builder nem no instalador.
- DPAPI depende da identidade/ambiente Windows. Copiar só o arquivo não constitui backup portátil garantido. Planejar recuperação segura do perfil e da chave antes de produção. Perda/comprometimento requer nova raiz e reinstalação de base confiável; rotação/revogação remota não está implementada. Há suporte a até quatro chaves previamente incorporadas, sem protocolo automático de rotação.

## Reproduzir o laboratório local

Não reutilizar identidades, instaladores nem chaves dos laboratórios `.2`/`.3`, PDF Test `.4` ou Assisted Lab na distribuição principal.

```powershell
node scripts/assisted/prepare-lab.mjs
$labInfo = Get-Content dist-desktop/assisted-lab/lab-local.json | ConvertFrom-Json
$env:FOURPRINT_PUBLIC_TRUST = $labInfo.publicTrust
$env:FOURPRINT_ASSISTED_LAB = '1'
$env:FOURPRINT_LAB_VERSION = '1.0.0-experimental.5'
.\node_modules\.bin\electron-builder.cmd --config electron-builder.assisted-lab.cjs --win nsis --x64 --publish never
$env:FOURPRINT_LAB_VERSION = '1.0.0-experimental.6'
.\node_modules\.bin\electron-builder.cmd --config electron-builder.assisted-lab.cjs --win nsis --x64 --publish never

# Criar notas.txt com novidades curtas, uma por linha, nesta pasta de laboratório.
node scripts/assisted/publish.mjs $labInfo.repository dist-desktop/assisted-lab/1.0.0-experimental.6 $labInfo.publicTrust $labInfo.key dist-desktop/assisted-lab/notas.txt
node scripts/assisted/hosting.mjs $labInfo.repository dist-desktop/assisted-lab/Caddyfile
```

`prepare-lab` recusa um laboratório já preparado: reutilizar seu `lab-local.json` em vez de gerar outra chave sobre os pacotes anteriores. A geração exige Windows e pode exigir permissão do ambiente de desenvolvimento para iniciar PowerShell/icacls. Não mudar políticas de execução.

Obter Caddy portátil Windows x64 pelo projeto oficial e conferir o checksum da mesma versão, observando o algoritmo publicado. O laboratório usou **2.11.7**, ZIP SHA512 `c308154504e53755958ffc62d8a2657ab17eefd6dd25c45dbb2c3d93d29519971f77ca652f7d0f2bcb7f980b9ba3822478865e29969d0272d7613550a43152b1`. Esse checksum de download da ferramenta não é o mecanismo de confiança dos instaladores 4PrintShops.

```powershell
# Executável portátil, sem serviço, firewall ou certificados; somente loopback.
& 'C:\caminho-local\caddy.exe' validate --config dist-desktop/assisted-lab/Caddyfile --adapter caddyfile
& 'C:\caminho-local\caddy.exe' run --config dist-desktop/assisted-lab/Caddyfile --adapter caddyfile
```

Configuração gerada: HTTP explícito em `127.0.0.1:18084`, administração/HTTPS automático/persistência de configuração desativados; somente GET/HEAD dos manifestos e `.exe` autorizados, sem listagem/upload. Não usa `tls internal` nem instala confiança local. Caddy é um executável independente; Node fica apenas no computador de desenvolvimento/publicação.

## Repositório e publicação

```text
repository/
  channels/internal-lab.json       # ou internal.json na identidade principal
  releases/1.0.0-experimental.6/<Setup>.exe
```

O publicador lê o `app.asar` do build e exige identidade, versão e raiz incorporada correspondentes. Recusa produtos de laboratórios antigos. Usa o nome de instalador esperado do builder, calcula SHA512, copia e confere os bytes copiados; publica o manifesto assinado **por último**, com renomeação. Não sobrescreve uma versão com bytes diferentes. Há trava exclusiva `.publish.lock`; após interrupção, investigar a execução e o conteúdo antes de remover uma trava abandonada. Só um publicador de cada vez. O processo de build/publicador é parte da origem confiável: não é uma análise antimalware do executável.

Não colocar código-fonte, backups, PDFs, arquivos de configuração privada ou chaves nesse repositório. A auditoria rejeita links/junctions e arquivos sensíveis antes de publicar/gerar configuração; não conceder gravação aos clientes. Caddy não oferece isolamento contra links adicionados posteriormente: permissões do diretório continuam essenciais.

Para publicar uma vez no servidor futuro: usar uma pasta de versões dedicada, acessível por SMB **apenas ao publicador para gravação**, ou preparar localmente e copiar a árvore completa de releases antes de substituir o manifesto. Não copiar só o manifesto. Renomeação/ACL/permissões/latência precisam ser testadas em SMB real. O cliente consulta por HTTP(S), sem montar unidade SMB nem guardar credenciais no aplicativo.

## Preparar a primeira base principal — ainda não realizado

1. Criar **nova chave principal**, fora do Git/servidor, sem `--lab`:

   ```powershell
   node scripts/assisted/keys.mjs 'C:\pasta-local-protegida-NOVA'
   $env:FOURPRINT_PUBLIC_TRUST = 'C:\pasta-local-protegida-NOVA\public-trust.json'
   .\node_modules\.bin\electron-builder.cmd --config electron-builder.assisted.cjs --win nsis --x64 --publish never
   ```

2. Essa configuração herda appId/produto/executável/NSIS/associação/perfil **4PrintShops** e exige produto `br.com.4printshops.desktop`, canal `internal`. Recusa a raiz `internal-lab`, mesmo com variável de ambiente do laboratório ativa. O build comum `npm run desktop:dist` continua sem raiz assistida e não a ativa sozinho.
3. Distribuir a primeira base por canal administrado confiável (por exemplo, mídia controlada), conferindo hash/fingerprint obtidos diretamente do publicador por canal separado. Não aprender a chave pública pela rede que se deseja autenticar. A instalação inicial dessa base é necessária uma vez; as versões atuais não têm essa raiz.
4. Após exportar todos os PDFs/fechar o app, instalar sobre a identidade principal existente. Não usar um Setup Assisted Lab como atualização principal. Validar preservação do perfil e associação antes de ampliar.
5. Configurar no perfil `%APPDATA%\4PrintShops\assisted-updates.json`:

   ```json
   {"enabled":true,"url":"http://NOME-DO-SERVIDOR:8084/","allowHttp":true}
   ```

   Para HTTPS confiável, usar URL HTTPS e `allowHttp:false`. Sem credenciais/query/fragmento; terminar em `/`. Só a origem é configurável: a chave de confiança permanece no aplicativo. Sem arquivo, desativado ou sem raiz incorporada, não consulta a rede.

## Servidor Windows — proposta, não executar ainda

Recomendado Caddy portátil com repositório em disco local do servidor, publicado por SMB. IIS é alternativa caso já esteja administrado; habilitar componentes IIS apenas para isso exige mais configuração. SMB direto não é implementado no cliente desta etapa.

Antes de qualquer ação real, levantar nome/IP, versão do Windows, sub-rede, pasta local e compartilhamento, contas de publicação/leitura, política de backup e regra de firewall existente. **Apresentar configuração/comandos finais e solicitar autorização** para serviços/tarefas agendadas, portas, ACLs ou certificados. Nada disso foi feito no servidor.

- Copiar somente Caddy, configuração pública e repositório. Chave DPAPI fica no publicador.
- Adaptar o site `http://127.0.0.1:18084`/`bind` para o IP interno e a porta aprovados. Manter rotas restritas, sem browse/upload/admin API. Executar inicialmente em primeiro plano com `caddy run --config "C:\4PrintShopsUpdates\Caddyfile" --adapter caddyfile` **somente depois de aprovada a escuta de rede**.
- Definir conta dedicada com leitura/execução no Caddy/configuração e somente leitura no repositório; publicador/admins são os únicos escritores. Avaliar tarefa agendada/serviço e firewall limitado aos computadores/sub-rede da gráfica em etapa autorizada. Não executar como administrador por conveniência.
- Publicação autenticada permite começar com HTTP interno aprovado sem certificado, aceitando os limites descritos. HTTPS exige configurar e homologar confiança TLS separadamente, nunca ignorar erros. Authenticode interno é outra etapa, dependente da VM/certificado confiável.

## Resultados efetivamente obtidos

- **199 testes Node**, zero falhas/skips: 184 existentes + 15 de manifesto, publicação, download, progresso, replay/expiração/versão, separação de identidades, falhas e confirmação/PDFs pendentes.
- Chrome headless curto aprovado: versão web com Node/HTTP, PDFs/QR, abas 100/350, exportação, proteção e comandos legados recusados. Artefatos `%TEMP%\4printshops-browser-v2JMs8`.
- Electron real aprovado: sandbox/isolamento/worker/miniaturas, giro, exportação/cancelamento/falha, proteção, preferências e cópia seletiva. Foi corrigida somente uma corrida no **harness de teste**, que lia JSON de status durante escrita; agora o helper publica esse arquivo atomicamente. Artefatos `%TEMP%\4printshops-desktop-8IfYE8`.
- **Assisted Lab .5 realmente instalada por NSIS em pasta temporária isolada**. Caddy real serviu a `.6`, com a mesma chave incorporada. Startup detectou versão/notas; Mais tarde e uso de QR funcionaram. Aplicativo instalado recusou manifesto adulterado, downgrade assinado e instalador alterado. Download correto gerou arquivo verificado; PDF girado em aba oculta bloqueou preparo, exportação real manteve giro/original, cancelar preservou estado e confirmar localizou o instalador sem executar/fechar. Offline e retorno do servidor passaram; preferências/formatos persistiram após reiniciar a `.5`. Artefatos `%TEMP%\4printshops-assisted-installed-hEEn2u`, relatório `result.json`.
- Conferência separada: Caddy serviu manifesto real com PATH sem Node, somente loopback. ACL da chave: apenas usuário e SYSTEM, sem herança. Nenhum certificado foi instalado.
- Consulta somente de leitura ao Shell do Windows confirmou o Explorador na pasta correta, com o instalador verificado selecionado. Não se executou o arquivo.
- `node tests/assisted-key-windows.mjs` passou: geração descartável em disco local, proteção/desproteção DPAPI real, correspondência da chave pública e recusa após adulterar o arquivo protegido. Só os arquivos desse teste foram removidos; a chave do laboratório foi preservada fora do Git.
- Dois ASAR auditados, **417 entradas** cada: runtime igual às fontes, bibliotecas presentes e ausência de testes/docs/PDFs/chaves. Sintaxe JS/MJS/CJS, parsing do helper PowerShell e `git diff --check` aprovados. Artefatos ignorados; nenhum commit/push.

| Instalador em `dist-desktop/assisted-lab/` | Bytes | SHA256 |
| --- | ---: | --- |
| `1.0.0-experimental.5/4PrintShops-Assisted-Lab-1.0.0-experimental.5-Windows-x64-Setup.exe` | 123039514 | `607f2e4c527dfe44ebd001c4589dcf31d5cdc9b070e4158a1f3c295dabc46ac2` |
| `1.0.0-experimental.6/4PrintShops-Assisted-Lab-1.0.0-experimental.6-Windows-x64-Setup.exe` | 123039577 | `49ca3dd9a86e98b7ef5284e971b45d85fcef75839051647efe4e20cd8a91ca54` |

Para repetir o ensaio instalado: `node tests/desktop-assisted-installed.mjs`, com Caddy em `%TEMP%\4printshops-caddy-2.11.7\caddy.exe` ou variável `FOURPRINT_CADDY`. O harness provisiona somente a base .5 isolada e usa diálogos nativos; **não executa a .6**. Não apontar para perfil/servidor reais. Reinícios de laboratório não substituem uma validação de atualização principal em duas máquinas.

## Pendências e piloto de dois computadores

1. Aprovar configurações reais de rede/ACL/operação e política de guarda/recuperação da chave principal. Servidor real, SMB, HTTPS interno e Windows de outras máquinas **não testados** nesta etapa.
2. Produzir a base principal confiável, fazer backup dos formatos/preferências com aplicativo fechado e instalar no PC A. Confirmar identidade/perfil, PDFs/QR, Abrir com/duplo clique e importação/exportação.
3. Instalar a mesma base no PC B. Publicar uma versão principal superior uma vez. A deve detectar/adiar e continuar trabalhando; B deve detectar e baixar. Manter um PDF alterado em aba oculta: preparar deve bloquear. Exportar, cancelar/reabrir o preparo e localizar.
4. No B, fechar conscientemente o app e **executar manualmente** o Setup verificado; confirmar versão, formatos/preferências, associação e PDFs originais. Depois repetir no A em outro horário. Esta execução da versão oferecida ainda **não foi realizada** no laboratório atual.
5. Testar servidor desligado, perda de conexão, manifestação/Setup adulterados, versão velha, permissões e clock. Não apagar o estado antirreplay para contornar rejeições. Cache é reaproveitado só após nova publicação válida; não há retomada parcial/gerenciamento automático de versões antigas do cache.
6. Em falha, adiar no outro computador e interromper a oferta. Publicar correção em versão/seq superior. Guardar instalador anterior confiável e backup do perfil; retorno manual supervisionado somente com app fechado e após conferir compatibilidade dos dados. Não há downgrade/rollback automático nem garantia de restaurar documentos em memória.

**Não homologado:** instalação/reinício automáticos pelo electron-updater, certificados Authenticode internos, atualização principal real em duas máquinas e execução manual do Setup .6 oferecido. Pendências anteriores de PDFs via SMB/UNC com FQDN/IPv4 permanecem documentadas em `ABRIR_PDFS_WINDOWS.md`.

Referências oficiais: [Caddy estático](https://caddyserver.com/docs/quick-starts/static-files), [opções globais](https://caddyserver.com/docs/caddyfile/options), [file_server e limites com symlinks](https://caddyserver.com/docs/caddyfile/directives/file_server), [releases oficiais](https://github.com/caddyserver/caddy/releases/tag/v2.11.7).
