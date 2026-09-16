# Known limitations and remaining blockers

This document states the boundary of the first standalone milestone. It is not release evidence.

## Runtime and authentication

- Only application/client-credentials client construction is implemented. Delegated OAuth, secure delegated token caching, and account-selection UX remain open.
- Authentication, tenant consent, endpoint permissions, throttling, pagination, and write cleanup have not been tested against an authorized test tenant.
- Operations lacking a strict executable contract remain in the administrative registry and preflight status; they are not deleted and are not registered in active action enums.
- Planner ownership/container semantics, Planner write serialization/ETag behavior, To Do application-write support, Microsoft Search chat-message support, and large-file transfer sessions require further endpoint-specific work.

## Host dependency

Hermes currently has a generic approval modification-order defect described in the governing plan: an earlier hook can modify arguments after another hook evaluated approval against different arguments. This repository does not add a Microsoft-specific workaround. Write execution must remain disabled until the separate generic Hermes core fix binds approval to the final arguments that execute.

## Files and pagination

- The offline foundation verifies generated drive collection request information and binary normalization, not complete path-based `/content` upload/download behavior.
- Multi-page Graph traversal is implemented as an offline helper with authenticated next-link validation, cycle detection and a global page cap (`microsoft365/paging.py`, WP5). No executable operation consumes it yet, so multi-page behavior against Graph remains unverified.
- The 10 MiB simple-transfer product contract is retained in metadata, but large-file upload sessions are not implemented.

## Delivery

- The repository is not yet in the Hermes plugin catalog and has no SHA-pinned catalog entry.
- Plugin validation/doctor are CI gates only where the pinned Hermes release exposes the commands compatibly.
- No production support, release readiness, or remote tenant behavior is claimed.
