# Graph response status triage

This runbook handles sanitized HTTP status metadata only. Do not log or attach tokens, headers, identifiers, URLs with IDs, or request/response bodies. The plugin is Alpha-only; these steps do not imply Graph behavior has been verified in a tenant.

## Common first steps

1. Record plugin version/commit, operation name, authentication mode, UTC timestamp, and status code.
2. Retry only when the status-specific section permits it; use a bounded attempt count and backoff.
3. Preserve the original sanitized status if a retry changes the outcome.
4. Stop for unexpected writes, data exposure, or repeated failures and use [incident response](incident-response.md).

## 401 Unauthorized

**Likely causes:** expired/missing token, wrong audience, invalid client credentials, or authentication mode mismatch.

**Action:** stop retries; refresh credentials through Hermes secret handling and re-run a minimal synthetic read after confirming consent. Never request or paste a token. If the failure persists, treat tenant authentication as unverified and escalate with sanitized metadata.

## 403 Forbidden

**Likely causes:** missing admin consent, insufficient application role, tenant policy, or delegated/application mismatch.

**Action:** do not retry repeatedly and do not broaden permissions automatically. Compare the operation with endpoint-specific permission evidence, obtain tenant-owner approval, and follow [consent](consent.md). Keep the operation blocked until permissions are verified.

## 404 Not Found

**Likely causes:** wrong resource scope, stale/invalid synthetic identifier, unsupported endpoint, or tenant visibility rules.

**Action:** verify only placeholder identifiers and required container relationships. Do not disclose or paste the missing resource identifier. Retry only after correcting the sanitized input; otherwise report not found without inferring that the tenant or resource exists.

## 409 Conflict

**Likely causes:** resource state conflict, duplicate creation, or concurrent update.

**Action:** do not blindly retry writes. Re-read a synthetic resource when authorized, resolve the conflict according to the operation contract, and require host approval. If the operation is a non-executable write, stop: CORE-1/CORE-2 remain blockers.

## 412 Precondition Failed

**Likely causes:** missing or stale `If-Match` ETag/precondition.

**Action:** do not remove the precondition to force a write. Re-read the synthetic resource, obtain a fresh ETag through the approved path, and retry once only if the contract requires it. Keep writes disabled until CORE-1/CORE-2 are closed.

## 429 Too Many Requests

**Likely causes:** Graph throttling or tenant rate policy.

**Action:** honor `Retry-After` when available without recording its surrounding response; otherwise use bounded exponential backoff with jitter. Cap attempts, stop on continued throttling, and avoid parallel retries. Do not claim throttling support as tenant-verified based on this runbook.

## 5xx (500, 502, 503, 504)

**Likely causes:** transient service or gateway failure, tenant service health, or an unverified endpoint path.

**Action:** retry idempotent reads only with bounded backoff. Do not automatically retry writes. If failures persist, stop and check the approved Microsoft 365 service-health channel without copying tenant data here; escalate with sanitized metadata.

## Closure

Record the outcome as resolved, blocked, or unverified. Never close an issue by deleting evidence or by asserting tenant support from an offline test.
