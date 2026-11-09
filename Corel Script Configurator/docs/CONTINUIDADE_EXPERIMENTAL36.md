# 4PrintShops — Continuidade da experimental.36

**Estado:** Desenvolvimento em andamento — NÃO HOMOLOGADO
**Objetivo:** Continuar no computador de casa o desenvolvimento e os testes da atualização transacional do 4PrintShops.

## 1. Contexto do projeto

- Repositório: https://github.com/marcos-jo/automacaoAdesivosCoreldraw
- Branch principal: `main`
- Aplicativo: 4PrintShops Desktop, Electron/Node.js, Windows.
- Última versão instalada e preservada na gráfica: experimental.34.
- Experimental.35: corrigiu o empacotamento do LibreOffice privado, mas falhou ao atualizar a instalação .34.
- Experimental.36: mecanismo de atualização transacional em desenvolvimento.

Leia obrigatoriamente `AGENTS.md`, `docs/ESTADO_PROJETO.md`, `docs/DISTRIBUICAO_CAMINHOS_EXPERIMENTAL35.md`, `docs/DIAGNOSTICO_UPGRADE_NSIS35.md` e `docs/ATUALIZACAO_TRANSACIONAL_EXPERIMENTAL36.md`, quando existentes.

## 2. Problema original

A experimental.34 introduziu LibreOffice privado para converter DOC/DOCX sem precisar instalar outro aplicativo.

Na estação da gráfica foram detectados 1.108 arquivos ausentes após a instalação, apesar de estarem presentes no pacote.

A experimental.35 corrigiu a distribuição, encurtando o caminho do LibreOffice para `lo` e introduzindo verificações de integridade.

Entretanto, a atualização .34 → .35 falhou no desinstalador NSIS antigo.

## 3. Diagnóstico confirmado

- `un.atomicRMDir` falha com determinados caminhos longos.
- Reprodução em laboratório com caminho de 260 caracteres.
- Operações usando caminhos estendidos do Windows funcionaram nos cenários testados.
- Arquivos realmente bloqueados continuaram protegidos.
- A remoção antecipada de associações em "Abrir com" deixou alterações parciais no Registro após a falha.
- Corrigir somente o instalador novo não resolve necessariamente o desinstalador antigo já distribuído.

## 4. Desenvolvimento iniciado na experimental.36

Foi iniciada uma arquitetura de atualização transacional com:

- Inventário de arquivos.
- Preservação da instalação anterior.
- Registro das operações.
- Validação de integridade.
- Tratamento de rollback.
- Migração compatível com instalações anteriores.
- Preservação das configurações e associações do aplicativo.

Um problema relacionado ao tratamento de `app.asar` pelo Electron foi identificado e corrigido nos testes correspondentes.

O último relatório informou 485 testes Node aprovados.

Esses resultados não representam homologação integral da atualização .36.

## 5. Estado no momento da interrupção

O Codex foi interrompido pelo limite de uso durante a execução dos ensaios integrais.

A instalação limpa em laboratório estava nas etapas finais de registro e cache, mas seu resultado definitivo não foi confirmado no histórico disponível.

Antes de executar qualquer novo teste:

1. Inspecionar o Git e os arquivos locais.
2. Recuperar os logs e resultados persistidos.
3. Determinar quais ensaios realmente terminaram.
4. Verificar se existem processos de instalação ainda ativos no ambiente atual.
5. Não presumir que o laboratório da gráfica foi transferido pelo GitHub.

Na última conferência relatada, havia aproximadamente 38 caminhos pendentes de commit.

## 6. Testes que faltam concluir

**Instalação limpa**
- Integridade do pacote.
- Instalação e inicialização do LibreOffice privado.
- Conversão DOC/DOCX sem Office global.
- Registro, associações e limpeza do instalador.

**Atualização .34 → .36**
- Migração de instalação com caminhos longos.
- Preservação do executável, perfil e preferências.
- Preservação do leitor PDF padrão e `UserChoice`.
- Funcionamento das associações próprias do Hold Test.
- Ausência de instalações duplicadas.

**Rollback**
- Falha na cópia.
- Falha nas operações de registro.
- Arquivo bloqueado.
- Manifesto inválido.
- Falha na validação de integridade.
- Falha de inicialização de componentes.

O rollback somente será considerado aprovado quando a versão anterior continuar íntegra e utilizável.

## 7. Segurança

- Não executar instalação, desinstalação ou testes destrutivos na estação da gráfica a partir de casa.
- Utilizar laboratórios isolados.
- Não realizar impressões físicas.
- Não alterar drivers, impressoras ou políticas globais do Windows.
- Não remover verificações de integridade para aprovar testes.
- Não apagar trabalhos existentes.
- Não incluir arquivos reais de clientes, credenciais, instaladores ou backups pessoais no GitHub.
- Não criar commits adicionais nem executar push sem autorização.

## 8. Estado da estação da gráfica

A experimental.34 continuava registrada após a falha da .35.

Os 75 arquivos do perfil foram conferidos e estavam íntegros.

As entradas próprias do Hold Test em "Abrir com" desapareceram durante a tentativa de atualização. O leitor PDF padrão não foi alterado.

O estado final da janela do Setup e de eventuais operações de recuperação deve ser conferido na própria estação; não deve ser presumido a partir deste documento.

## 9. Objetivo do próximo Codex

Retomar a implementação existente, não recomeçá-la.

Concluir a auditoria da instalação limpa, a migração .34 → .36 e os testes de rollback.

Corrigir apenas os defeitos confirmados.

Depois de aprovar integralmente os testes, finalizar e auditar a experimental.36, documentando os riscos restantes.

Não instalar automaticamente na gráfica.

## 10. Funcionalidades futuras — fora deste trabalho

- Importação de PDFs criptografados em lote, com estados individuais e sem travamentos.
- Solicitação de senha e cópia desbloqueada quando autorizada.
- Reconhecimento de TrimBox, BleedBox e marcas de corte.
- Homologação física dos livretos.
- Melhorias restantes de zoom e compatibilidade PDF.

**Prioridade absoluta: concluir e validar o mecanismo de atualização da experimental.36.**