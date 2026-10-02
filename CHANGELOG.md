# Changelog

All notable changes to this project are documented here.

## [Unreleased]

### Added

- **Delegated authentication.** Device-code OAuth as a first-class auth mode alongside
  application credentials, with a token cache protected by the operating system; an
  unencrypted cache is refused. Delegated scopes are derived from the enabled capability set,
  and application credentials are never read on the delegated path.
- **Resumable upload sessions.** Transfers above the 10 MiB simple ceiling run through Graph
  upload sessions with aligned 320 KiB fragments, explicit byte ranges, `nextExpectedRanges`
  resume and a bounded attempt budget. A session that stops making progress is reported as
  `upload_session_incomplete` rather than assumed complete.
- **Verified ranged downloads.** Transfers above the simple ceiling run as aligned ranged reads
  that require `206`, verify every `Content-Range` against the window requested, and reject a
  mismatch instead of accepting whole-file bytes as the requested range.
- **Post-write reconciliation.** Every confirmed write is re-read and reported as `confirmed`,
  `mismatched` or `unverified`; `mismatched` is never repaired or silently retried, and
  `unverified` is never promoted to `confirmed`. Callers may opt out explicitly with
  `verify: false`.
- **Outlook draft creation and send**, **SharePoint and OneDrive upload**, and
  **calendar event creation with `transactionId`** promoted to executable operations.
- **Teams 1:1 chats.** `teams.list_chats`, `teams.read_chat_messages` and
  `teams.send_chat_message`, with `Chat.ReadBasic.All` and `Chat.Read.All` in application mode.
- **Authenticated inbound Teams activities.** A boundary that verifies the bearer token before
  reading the body, pins RS256 by name against algorithm confusion, checks audience, issuer and
  lifetime against an explicit clock and tolerance, projects the Activity onto a closed field
  set, and correlates identity by `(tenant_id, sender_id, conversation_id)` with bounded
  deduplication. Admitting an Activity can never send anything.
- **A Bot Framework webhook.** `microsoft365/teams_webhook.py` is a WSGI application (no new
  dependency) that serves one path, `POST` only, as JSON, with a required and bounded
  `Content-Length`, and renders the boundary's answer as a status code. Every distinguishable
  token failure leaves as the same anonymous `401`; a failure of ours is a retryable `500`; a
  failed reply is a `200` carrying the failure, because a `5xx` would provoke a redelivery that
  the duplicate guard would suppress anyway. Verified over a real TCP socket, not only through
  WSGI.
- **Operator-bound reply loopback.** A non-duplicate inbound decision becomes a reply whose
  destination comes from the operator's route and never from the Activity body, delivered only
  through the separately authorised send operation.
- **Explicit conflict policy** for drive writes: `conflict_behavior` of `replace`, `rename` or
  `fail`, with the boolean `overwrite` retained as a deprecated alias in both directions.
- **Fail-closed user boundary.** User-scoped operations without a configured `user_id` fail with
  `configuration_error` instead of trusting a `user_id` from the payload.
- **A sanitized tenant smoke runbook** covering configuration, minimal reads, large transfers,
  reversible writes, delegated auth and the Teams inbound loop, with per-status triage and an
  explicit statement of what the evidence does not prove.
- **Cross-layer drift guards.** Tests that read the handler modules as source and fail when a
  handler consults an argument the contract does not declare, when a write honours `verify`
  without publishing it, when a status carries a documented action but is never emitted, when
  the runbook triages a failure the code cannot produce, or when the docs name an operation
  that is not in the registry.

### Changed

- Write approval is bound to the final tool arguments after modifier hooks, not to the
  pre-modifier arguments. This requires a Hermes host carrying
  `docs/hermes-core-approval-final-args.patch`; without it, keep the write capabilities
  disabled.

### Fixed

- An unreachable JWKS endpoint was reported as an authentication refusal, telling Bot Framework to
  stop retrying and silently dropping a legitimate message. Failures of ours are now
  `configuration_error`; failures about the token stay sanitized `401`s.
- The endpoint read its body by asking the stream for the maximum body size, which under WSGI
  blocks until the peer closes the connection. The reader now asks for exactly `Content-Length`
  bytes, the only delimiter a request body has.

- `verify` on `outlook.create_draft` was honoured by the handler but never declared by the
  contract, so the documented opt-out was refused.
- The deprecated `overwrite` alias was honoured by the upload handler but rejected by the
  contract, so the documented backwards compatibility never worked. Passing `overwrite` and
  `conflict_behavior` together is now refused rather than letting one silently outrank the
  other.

## [0.1.0a4] - 2026-09-17

- Added sanitized GraphError envelopes at registered dispatch.
- Added bounded retry for idempotent reads only; writes remain single-attempt.
- Exposed strict discriminated operation schemas with runtime-matching limits.

## [0.1.0a3] - 2026-09-17

- Enforced single-user resource boundaries and redacted preflight identifiers.
- Added endpoint-derived preflight requirements and Planner-only diagnostics.
- Added CI quality gates, coverage baseline, support policy, and operational runbooks.

## [0.1.0a2] - 2026-09-17

- Refreshed endpoint-level Microsoft Graph permission evidence without promoting any
  claim to remote verification.
- Expanded adversarial approval and read-to-write mutation coverage.
- Added bounded streaming for simple drive downloads.

## [0.1.0a1] - 2026-09-17

- Established the offline-verified standalone Microsoft 365 Graph plugin foundation.
- Added strict capability validation, operation registry reporting, bounded file transfer,
  paging, and typed SDK request contracts.
- No Microsoft 365 tenant was contacted; the catalog entry is proposed in upstream PR
  #114530, and this pre-release makes no signature, tenant, or production support claim.

[Unreleased]: https://github.com/fabiomotta0311/hermes-microsoft365-plugin/compare/main...HEAD
[0.1.0a4]: https://github.com/fabiomotta0311/hermes-microsoft365-plugin/releases/tag/v0.1.0a4
[0.1.0a3]: https://github.com/fabiomotta0311/hermes-microsoft365-plugin/releases/tag/v0.1.0a3
[0.1.0a2]: https://github.com/fabiomotta0311/hermes-microsoft365-plugin/releases/tag/v0.1.0a2
[0.1.0a1]: https://github.com/fabiomotta0311/hermes-microsoft365-plugin/releases/tag/v0.1.0a1
