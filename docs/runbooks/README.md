# Operational runbooks

These runbooks apply to the `pro/microsoft365-plugin` line. They describe safe local triage, the
tenant acceptance procedure, and release operations. Running the tenant smoke runbook against an
authorised tenant is the only thing that turns a permission claim into evidence; the other
runbooks describe how to operate and recover, and none of them confers production support on a
deployment that has not run it.

## Safety rules

- Never collect or transmit access tokens, client secrets, tenant/user IDs, message/file/calendar/chat/task content, raw URLs with identifiers, headers, or request/response bodies.
- Use operation names, sanitized status codes, package version, commit SHA, Python version, and timestamps only when needed.
- Do not add telemetry, debug logging, or diagnostic uploads to investigate an incident.
- Stop and escalate privately when a security issue or suspected credential exposure is found.

## Runbooks

- [Consent and tenant readiness](consent.md)
- [Tenant smoke test](tenant-smoke.md): the acceptance procedure -- configuration, minimal reads, large transfers in both directions, reversible writes, delegated auth, and the Teams inbound loop, with per-status triage
- [Graph response status triage](graph-errors.md): 401, 403, 404, 409, 412, 429, and 5xx
- [Rollback](rollback.md)
- [Incident response](incident-response.md)

CORE-1, CORE-2, and tenant verification are release blockers. A runbook cannot waive them or turn an offline result into remote evidence.
