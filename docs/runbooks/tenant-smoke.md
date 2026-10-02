# Smoke test em tenant autorizado

Este runbook leva o plugin de "verificado offline" a "verificado no real". Ele é o único
documento que autoriza alguém a afirmar que uma permissão foi concedida ou que uma operação
funciona contra o Microsoft Graph — testes offline, `plugins validate` e `plugins doctor` não
substituem esta evidência.

Execute no mesmo ambiente Hermes que fornece `agent.secret_scope`; o plugin resolve
`MICROSOFT365_CLIENT_SECRET` dentro do escopo de segredo. O segredo nunca entra no repositório,
no chat, nem neste documento.

## Regra de evidência

Registre **apenas**: data/hora, versão do plugin (`git rev-parse HEAD`), tenant anonimizado
(ex.: `tenant-A`), operação, `status`, categoria de erro e contagem de itens.

Nunca registre: token, segredo, endereço de e-mail completo, assunto, corpo, nome de arquivo
real, `chat_id`, `tenant_id` real ou `upload_url`. Os envelopes de resultado já redigem
identidades e URLs sensíveis por construção — se algum aparecer numa evidência, isso é um
defeito a reportar, não um dado a colar.

## Pré-requisitos

- tenant Microsoft 365 com conta de teste dedicada, sem dados pessoais;
- consentimento administrativo para as permissões app-only de
  [`graph-permissions.md`](../references/graph-permissions.md);
- `tenant_id`, `client_id`, `user_id` configurados no namespace do plugin;
- `MICROSOFT365_CLIENT_SECRET` disponível somente no secret scope;
- para as etapas delegadas: identidade delegada configurada e consentimento concedido;
- para a etapa 6: registro de bot no Bot Framework e o webhook apontando para o host.

## Ordem obrigatória

As etapas são cumulativas e cada uma é pré-requisito da seguinte. **Pare na primeira falha** e
corrija a causa antes de seguir: uma falha de autenticação na etapa 1 torna todo o resto
ininterpretável.

### Etapa 0 — configuração, sem rede

Confirme que o plugin está configurado antes de tocar o Graph:

```text
hermes plugins validate microsoft365
hermes plugins doctor microsoft365 --ci
```

Esperado: as capabilities que você pretende testar aparecem habilitadas e as credenciais são
reconhecidas como presentes. Uma capability desabilitada não deve aparecer na superfície do
modelo.

### Etapa 1 — leitura mínima

A menor leitura possível, para provar autenticação e consentimento isoladamente:

```text
outlook.search(user_id=<configured-user>, query="", top=1)
```

Esperado: `status=succeeded` com no máximo 1 item.

**Triage:**

| Resultado | Significado | Ação |
| --- | --- | --- |
| `authentication_required` | credencial ausente ou recusada | confira escopo de segredo e `client_id`/`tenant_id` |
| `consent_required` | consentimento administrativo não concedido | conceda no portal e repita |
| `permission_denied` | consentimento existe, app role não atribuída | atribua a app role e repita |
| `configuration_error` | `user_id` ausente para operação user-scoped | configure `user_id`; o plugin é fail-closed aqui |
| `throttled` | limite do tenant atingido | aguarde; não aumente `top` |
| `transport_error` | rede/proxy | confirme a saída para `graph.microsoft.com` |

### Etapa 2 — leituras dos quatro serviços

Uma leitura por serviço, para confirmar que as permissões são por recurso e não globais:

```text
sharepoint.search(query="*", top=1)
onedrive.search(query="*", drive_id=<drive-id>, top=1)
teams.list_teams(user_id=<configured-user>)
calendar.search(user_id=<configured-user>, top=1)
planner.list_plans()
todo.list_task_lists(user_id=<configured-user>, top=1)
```

Esperado: `status=succeeded` em cada uma. Uma falha isolada indica permissão faltando **naquele**
recurso; não é evidência de que os outros estão errados.

### Etapa 3 — transferência grande, nos dois sentidos

Esta etapa é a que mais vale exercitar, porque é onde um `2xx` pode não significar o que parece.

1. Faça upload de um arquivo pequeno com o comportamento definitivo de conflito:

   ```text
   sharepoint.upload_files(
     drive_id=<drive-id>, item_path="smoke/small.bin",
     content_base64=<base64>, content_type="application/octet-stream",
     conflict_behavior="replace"
   )
   ```

   Esperado: `status=succeeded`, `result.transfer="simple"` e
   `reconciliation.status="confirmed"`.

2. Faça upload de um arquivo **acima de 10 MiB**:

   Esperado: `result.transfer="session"` e `reconciliation.status="confirmed"`. O caminho de
   sessão usa fragmentos de 320 KiB e retoma pelo `nextExpectedRanges`.

3. Baixe o mesmo arquivo e compare o tamanho:

   ```text
   sharepoint.download_files(drive_id=<drive-id>, item_path="smoke/big.bin")
   ```

   Esperado: `result.size` igual ao tamanho enviado. Acima de 10 MiB a resposta vem por leituras
   em range, e cada `Content-Range` é conferido contra a janela pedida.

4. Leia o veredicto de reconciliação com atenção. Ele é a parte que distingue sucesso de
   verdade:

   | `reconciliation.status` | Significado | Ação |
   | --- | --- | --- |
   | `confirmed` | o item relido bate com o que foi pedido | prossiga |
   | `mismatched` | o Graph reportou sucesso e o item contradiz | **incidente**; pare e investigue |
   | `unverified` | a escrita pode estar boa; o plugin não pôde confirmar | investigue a releitura |

**Triage dos status de transferência:**

| `status` | Significado |
| --- | --- |
| `upload_session_incomplete` | a sessão parou sem completar; nada foi dado como sucesso |
| `download_size_mismatch` | o total montado difere do tamanho declarado nos metadados |
| `range_window_mismatch` | o servidor devolveu uma janela diferente da pedida — transferência recusada |
| `range_not_content_partial` | o servidor ignorou o `Range`; recusado em vez de aceito como arquivo inteiro |
| `download_too_large` | acima de 250 GiB; use transferência em seções |
| `invalid_conflict_behavior` | `conflict_behavior` fora de `replace`/`rename`/`fail` |
| `invalid_chunk_size` / `invalid_range_size` | janela não é múltiplo de 320 KiB |
| `not_a_file` / `metadata_without_size` | o alvo é pasta ou não declara tamanho |

5. Limpeza: remova `smoke/small.bin` e `smoke/big.bin` pela operação aprovada e confirme por
   leitura que não estão mais visíveis.

### Etapa 4 — escrita reversível

Só execute com o Hermes contendo o patch de aprovação em argumentos finais
([`hermes-core-approval-fix.md`](../hermes-core-approval-fix.md)). Sem ele, mantenha
`outlook.create_draft` desabilitada.

1. Habilite **somente** `outlook.create_draft`.
2. Crie um rascunho identificável por um marcador não sensível.
3. Confirme `reconciliation.status="confirmed"` — o rascunho foi relido e o `subject` confere.
4. Remova o rascunho pela operação de limpeza aprovada.
5. Confirme por leitura que não está mais disponível.
6. Desabilite a capability novamente.

**Não use `outlook.send` nesta etapa.** Envio é irreversível e exige campanha própria, com
destinatário interno controlado, revisão adversarial e aprovação explícita.

### Etapa 5 — autenticação delegada

Somente se a identidade delegada estiver configurada:

1. habilite uma única leitura delegada (ex.: `teams.list_chats`);
2. complete o fluxo device-code;
3. confirme que o cache de token é o protegido pelo sistema operacional — um cache não
   criptografado é recusado pelo plugin, e essa recusa é o comportamento correto;
4. confirme `status=succeeded`;
5. registre se houve prompt de consentimento e para qual conta.

Esperado: a operação usa escopos delegados derivados do conjunto habilitado, não os escopos
app-only. Se o modo delegado usar credencial de aplicação, isso é um defeito a reportar.

### Etapa 6 — Teams: entrada autenticada e resposta roteada

Esta é a única etapa que não pode ser provada offline, porque depende de um listener.

1. Registre a identidade externa contra uma sessão Hermes e vincule a rota de resposta, no host:

   ```text
   registry.bind(Identity(tenant_id, sender_id, conversation_id), session_id)
   router.bind(Identity(...), ReplyRoute(chat_id=<graph-chat-id>))
   ```

   Atenção: `chat_id` é um identificador do **Graph**, não o `conversation.id` do Bot Framework.
   São identificadores de sistemas diferentes.

2. Envie uma mensagem do Teams para o bot.
3. Confirme, na decisão de ingestão: `session_bound=true`, `duplicate=false`, e que o texto
   chegou à sessão vinculada.
4. Confirme que a resposta saiu **no chat vinculado** — e em nenhum outro.
5. Reenvie a mesma Activity (ou force um retry) e confirme que a decisão traz `duplicate=true`
   e que **nenhuma segunda resposta** foi enviada. Reentrega **não** é erro: a fronteira devolve
   uma decisão, não um status de falha, e é assim de propósito -- o Bot Framework reenvia até
   receber sucesso.

**Triage:**

| `status` | Significado | Ação |
| --- | --- | --- |
| `invalid_token` | header ausente, malformado ou assinatura inválida | confira o registro do bot e a URL do webhook |
| `token_algorithm_not_allowed` | token não é RS256 | investigue: pode ser tentativa de forjar |
| `token_audience_mismatch` | token emitido para outro bot | confirme o App ID |
| `token_issuer_mismatch` | token não veio do Bot Framework | idem |
| `token_expired` | token fora da janela | o Bot Framework reenvia; se persistir, confira o relógio do host |
| `unknown_identity` | remetente não vinculado | vincule antes de esperar entrega; o plugin não adivinha sessão |
| `no_reply_route` | sem rota vinculada | vincule a rota; o destino nunca vem do corpo |
| `reply_not_deliverable` | Activity de ciclo de vida, ou duplicada | não responda: não é um turno |

## Critério de aprovação

O tenant só pode ser marcado como verificado quando **todas** as etapas aplicáveis concluírem,
incluindo leitura pós-escrita e limpeza, com a evidência sanitizada registrada.

Ao final, atualize [`graph-permissions.md`](../references/graph-permissions.md) marcando cada
permissão testada como verificada, e registre no `CHANGELOG.md` a data e a versão testada.

## O que este runbook não prova

Diga isto junto com o resultado, para que ninguém leia mais do que a evidência sustenta:

- **Não prova durabilidade.** O estado de correlação e de rota do Teams é local ao processo;
  reinício perde vinculações e múltiplas réplicas precisariam de store compartilhado.
- **Não prova escala.** As transferências são bufferizadas em memória; o teto de 250 GiB é do
  Graph, não do que uma chamada de tool consegue carregar.
- **Não prova idempotência sob concorrência.** Não há verificação de ETag em leitura pós-escrita
  para todas as operações, nem coordenação entre instâncias.
- **Não cobre as escritas withheld.** `planner.*`, `todo.*`, `calendar.update_events` e
  `teams.send_messages` continuam não executáveis; o smoke não as exercita de propósito.
- **Não substitui revisão de segurança.** O secret scan e a suíte offline não são um threat model.
