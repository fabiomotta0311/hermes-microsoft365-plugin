# Tenant smoke test

Este runbook valida a instalação contra um tenant Microsoft 365 autorizado sem colocar
segredos no repositório. Ele deve ser executado no mesmo ambiente Hermes que fornece
`agent.secret_scope`; o plugin continua sendo responsável por resolver
`MICROSOFT365_CLIENT_SECRET` dentro do escopo de segredo.

## Pré-requisitos

- consentimento administrativo para as permissões app-only descritas em
  [`graph-permissions.md`](../references/graph-permissions.md);
- `tenant_id`, `client_id` e `user_id` configurados no namespace do plugin;
- `MICROSOFT365_CLIENT_SECRET` disponível somente no secret scope do Hermes;
- uma conta de teste dedicada, sem dados pessoais;
- `outlook.create_draft` só deve ser usado com o Hermes contendo o patch de aprovação final-argumento; sem ele, mantenha a capability desabilitada.

## Sequência obrigatória

1. Execute primeiro somente uma leitura pequena:

   ```text
   outlook.search(user_id=<configured-user>, query="", top=1)
   ```

2. Registre apenas metadados sanitizados: data/hora, versão do plugin, tenant anonimizado,
   operação, status, categoria de erro e correlation ID. Nunca registre token, segredo,
   endereço completo de e-mail, assunto ou corpo.
3. Confirme que a resposta contém `status=succeeded` e que a quantidade de itens não excede
   `top`.
4. Em qualquer falha de autenticação, consentimento, boundary de usuário ou transporte,
   interrompa o teste e corrija a causa antes de testar outra operação.

## Validação de escrita

A escrita real não deve ser liberada pela superfície do modelo enquanto o host não oferecer
aprovação baseada nos argumentos finais. Quando essa versão suportada estiver instalada:

1. habilite somente `outlook.create_draft`;
2. use uma conta de teste e destinatário interno controlado;
3. crie um rascunho identificável por um marcador não sensível;
4. leia o rascunho pelo ID retornado e confirme remetente, destinatário e assunto;
5. remova o rascunho pela operação de limpeza aprovada;
6. confirme por leitura que o recurso não está mais disponível;
7. registre a evidência sanitizada e desabilite a capability novamente.

Não use `outlook.send` neste smoke test. Envio é irreversível e requer uma campanha de
validação separada, com aprovação explícita e revisão adversarial.

## Critério de aprovação

O tenant somente pode ser marcado como verificado quando a sequência inteira for concluída,
incluindo leitura pós-escrita e limpeza. Testes offline, `plugins validate` e `plugins doctor`
não substituem essa evidência remota.
