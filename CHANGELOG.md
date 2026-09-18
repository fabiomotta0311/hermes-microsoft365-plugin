# Changelog

All notable changes to this project are documented here.

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

## [Unreleased]

- Added release engineering for reproducible wheel builds, isolated install smoke tests,
  SHA-256 checksums, and a lockfile-derived CycloneDX SBOM artifact.
- Documented the release procedure and the boundary between build evidence and catalog
  or tenant verification.

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
