# Formatos personalizados compartilhados — investigação de 16/11/2026

## Revisão de 17/11/2026 — checkpoint local da integração mínima

A orientação foi registrada separadamente em `d1ad3f0`. A integração local descrita a seguir compõe o checkpoint **Compartilha formatos personalizados nas montagens PDF**, junto à continuidade. Registros anteriores mantidos; compartilhamento no Electron ainda depende do ensaio isolado/reabertura e de futura homologação autorizada, conforme o plano no relatório A/B. Sem migração, instalação ou mudança de escopo.

## Atualização de 17/11/2026 — integração local autorizada

O texto de 16/11 abaixo permanece como histórico da proposta. Nesta rodada foi autorizado e implementado somente o compartilhamento local mínimo entre Redimensionar, Step & Repeat e N-up, incluindo salvar/reutilizar na montagem, serviço por renderer, notificações, serialização e preservação do armazenamento v1. Não houve migração nem acesso ao perfil operacional. Gerenciador global em menu, catálogo entre computadores e papéis arbitrários no Livreto continuam propostas pendentes; não foram antecipados. Resultado em revisão, ainda fora do aplicativo instalado: [relatório A/B](ORIENTACAO_E_FORMATOS_MONTAGENS_2026-11-17.md).

Proposta para revisão. Nenhuma implementação do catálogo compartilhado foi realizada nesta etapa. A interface do Step & Repeat foi tratada separadamente; prioridades e demais ferramentas permanecem preservadas.

## Armazenamento e uso atuais

- `public/js/tools/pdf-formats.js` já fornece o catálogo `PDFFormatos`: carregar, listar, validar, salvar/atualizar e excluir. A propriedade `pronta` permite aguardar a carga, inclusive com adaptador assíncrono. Cada instância mantém uma lista própria; ainda não há um serviço único com notificação de alterações entre ferramentas.
- `public/js/tools/pdf-format-storage.js` implementa `ler/gravar` em `localStorage`, chave `4printshops.pdf.formatos.v1`. O documento persistido tem a forma `{ "versao": 1, "formatos": [{ "id": "custom-…", "nome": "Convite", "largura": 140, "altura": 180 }] }`.
- Todas as medidas atuais são em **mm**. Não existe campo de unidade no esquema v1; não se pode recuperar uma unidade de entrada que nunca foi armazenada. Não converter nem arredondar dimensões ao compartilhar.
- Nomes são normalizados (espaços, Unicode) e comparados sem distinção de maiúsculas para impedir homônimos. IDs personalizados são estáveis, normalmente UUID. Há limite de 200 cadastros, nomes de 1–80 caracteres e medidas de 0,1–5000 mm. Nomes diferentes com medidas iguais ainda são permitidos.
- O adaptador conserva dados em memória e avisa quando leitura/escrita persistente falha. Essa alternativa **não sobrevive ao encerramento do processo**. A persistência normal depende da mesma origem/perfil, de gravação bem-sucedida e da preservação dos dados do aplicativo.
- No Desktop, `desktop/security.mjs` define a origem `app://4printshops`; `desktop/app.mjs` escolhe o diretório `userData` conforme o produto/perfil. Navegador HTTP, outra origem/porta ou outro perfil têm armazenamento separado. Não há arquivo JSON de catálogo mantido no repositório nem API de formatos no servidor atual.
- Esta investigação leu código e testes; não abriu nem alterou o banco de dados do perfil instalado e não inventaria os cadastros reais do operador.

## Criação, edição, exclusão e seleção

`public/js/tools/pdf-resize-ui.js` cria uma instância de `PDFFormatos`, aguarda `pronta` e popula o seletor. Salvar cria um ID; atualizar reutiliza o ID selecionado; excluir requer confirmação. Formatos padrão não podem ser alterados/excluídos. O catálogo apenas armazena dimensões: a transformação do PDF pertence ao fluxo separado de prévia/aplicação de Redimensionar.

Seletores atualmente independentes:

| Ferramenta | Origem das opções | Cuidados |
|---|---|---|
| Redimensionar | `PDFFormatos` + opção de medidas livres | Padrões A3/A4/A5/A6 e personalizados persistentes |
| Step & Repeat / N-up | `presets` e marcação em `pdf-imposition-ui.js` | Folhas A3/A4/SRA3 e medidas livres; não consulta o catálogo de Redimensionar |
| Livreto | `pdf-booklet-ui.js` e `pdf-booklet-layout.js` | Formato final A5/A4 e papel A4/A3 vinculados a regras próprias; não liberar formatos arbitrários apenas trocando o seletor |

## Menor evolução proposta

1. Reutilizar `PDFFormatos`, seu validador e `PDFArmazenamentoFormatos`. Criar uma instância compartilhada no renderer, com carga única e assinatura de eventos de alteração, mantendo o contrato existente. Serializar gravações para evitar perda de atualizações entre painéis; tratar a divergência entre abas/processos antes de prometer compartilhamento entre eles.
2. Acrescentar uma ação global **Formatos personalizados…** ao menu superior em `public/js/app/menus.js`, conectada pelo controlador do aplicativo. Um diálogo reutilizável fará o CRUD no serviço, sem depender de PDF aberto. Redimensionar passa a consumir o mesmo serviço, preservando IDs e dados.
3. Inicialmente manter os padrões fixos de cada ferramenta e acrescentar os mesmos formatos personalizados aos seletores compatíveis. Selecionar um cadastro no Step & Repeat/N-up deve preencher **somente as dimensões da folha**, por ação explícita do operador. Preservar rotação da arte, dimensões da peça, escala e grade. Orientação da folha é uma escolha independente; eventual troca largura/altura não deve regravar o cadastro.
4. Propagar alterações de lista sem aplicar outro tamanho automaticamente ao documento ou à montagem. Se o formato selecionado for renomeado/excluído, preservar as dimensões em uso e informar a situação; não regenerar com medidas diferentes silenciosamente.
5. Evitar novas cópias do cadastro por ferramenta. Reutilizar IDs; manter rejeição de nomes duplicados. Para dimensões coincidentes sob nomes diferentes, propor aviso/opção de reutilização, sem mesclar, excluir ou renomear automaticamente. Orientações invertidas e nomes diferentes podem ter significado operacional legítimo.

Nenhuma mudança em motores de PDF ou impressão é necessária para essa primeira integração. Livreto e formatos fora dos limites suportados exigem avaliação própria, sem ampliação automática de compatibilidade.

## Preservação de versões anteriores

- Manter inicialmente a chave e o esquema v1, sem migração destrutiva. Copiar o valor bruto antes da primeira alteração que eventualmente exija migração; não apagar o original nem trocar os IDs.
- Não tratar falha de leitura como catálogo vazio a ser sobrescrito. O comportamento atual avisa e permite substituir dados inválidos ao salvar; uma evolução compartilhada deve oferecer recuperação/backup e decisão explícita antes dessa substituição.
- Se houver esquema novo, tornar a migração validada e repetível, preservar nomes, IDs, dimensões e informações desconhecidas para recuperação. Só persistir o resultado completo após validação; não registrar uma migração parcial.
- Perfis/origens antigos diferentes podem precisar de exportação/importação autorizada. Não supor que manter o nome da chave transfere dados entre `http://localhost`, `app://4printshops` ou perfis de produtos diferentes.
- Não oferecer sincronização entre computadores como consequência de um serviço no renderer; isso exige outro escopo.

## Testes de uma futura implementação

Reutilizar `tests/pdf-formats.test.mjs`: CRUD, IDs, validação de nomes/dimensões, dados inválidos, quota e adaptadores assíncronos. Acrescentar testes de atualização simultânea de listas, conflitos de gravação, cadastros v1 antigos, reinício real no mesmo perfil isolado, indisponibilidade de armazenamento, alteração/exclusão do cadastro em uso e cópias de segurança. Conferir que selecionar uma folha não altera escala/arte e que abrir/fechar o gerenciador não modifica o PDF.

Complexidade estimada: **média**. Recomenda-se uma etapa posterior separada e autorizada, começando pelo serviço compartilhado e sua preservação de dados, seguida do menu e dos consumidores compatíveis. Não implementado nesta tarefa.
