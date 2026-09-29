# Hermes core approval fix

As escritas deste plugin dependem de uma garantia do host: a aprovação precisa ser
reavaliada depois que todo middleware terminar de modificar os argumentos. Caso contrário,
um hook pode aprovar um payload e o executor receber outro.

O patch [`hermes-core-approval-final-args.patch`](hermes-core-approval-final-args.patch)
implementa essa garantia no Hermes core:

- revalida a decisão quando argumentos são modificados;
- não reaplica modifiers durante a revalidação;
- bloqueia se os argumentos finais não receberem uma aprovação válida;
- revalida novamente quando `tool execution middleware` altera o payload;
- registra teste adversarial para alteração pós-aprovação.

## Aplicação

O patch deve ser aplicado no checkout do Hermes Agent, não neste repositório:

```bash
git am docs/hermes-core-approval-final-args.patch
```

Depois, execute a suíte do Hermes e instale esse plugin no mesmo ambiente. O commit de
referência no checkout usado para produzir o patch é:

```text
97104da814a fix(approval): bind write approval to final tool arguments
```

## Limite atual

Este patch corrige o bloqueio estrutural do host, mas não é uma evidência de tenant. A
superfície de escritas deste plugin continua deliberadamente não executável até haver:

1. uma versão do Hermes contendo este patch;
2. consentimento administrativo no tenant;
3. execução remota autorizada conforme o runbook
   [`tenant-smoke.md`](runbooks/tenant-smoke.md);
4. promoção explícita de uma operação reversível, começando por `outlook.create_draft`.

Não aplique o patch e habilite `outlook.send` como parte do mesmo rollout.
