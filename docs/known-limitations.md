# Known limitations and remaining blockers

This document states the boundary of the first standalone milestone. It is not release evidence.

## Runtime and authentication

- Only application/client-credentials client construction is implemented. Delegated OAuth, secure delegated token caching, and account-selection UX remain open.
- Authentication, tenant consent, endpoint permissions, throttling, pagination, and write cleanup have not been tested against an authorized test tenant.
- Operations lacking a strict executable contract remain in the administrative registry and preflight status; they are not deleted and are not registered in active action enums.
- To Do application-write support, Microsoft Search chat-message support, and large-file transfer sessions require further endpoint-specific work.

## Planner

- Planner containers, request bodies and `If-Match` handling are pinned by strict offline request-contract tests (`tests/test_planner_contracts.py`): the tenant-wide and group-owned plan listings, the plan-scoped bucket and task listings, the task-scoped read and update, `planId`/`bucketId` in the creation body, and a mandatory ETag on `planner.update_tasks`. A blank required identifier, and a blank optional `group_id`, are refused before any client, approval or token work, with the endpoint the identifier belongs to named in the error.
- No Planner operation is executable yet. All six Planner operations stay in the registry with `implementation_status = contract_verified` and `executable = False`, and no handler exists for them, so a Planner call cannot reach Microsoft Graph through this plugin.
- Planner permission claims are endpoint-specific and labelled `documented_not_verified`: each claim cites the endpoint it belongs to, no endpoint page was re-read offline, and no tenant has been contacted (R10, WP16). Assignments on task creation are not modelled, and any additional role that endpoint may require is unverified.
- The shared partial-model helper cannot express a PATCH field whose value is a *list* of generated models (for example `attendees` on `calendar.update_events`). It refuses such a field instead of silently dropping it, so affected updates stay unimplemented until the helper is extended deliberately.

## Host dependency

Hermes currently has a generic approval modification-order defect described in the governing plan: an earlier hook can modify arguments after another hook evaluated approval against different arguments. This repository does not add a Microsoft-specific workaround. Write execution must remain disabled until the separate generic Hermes core fix binds approval to the final arguments that execute.

## Files and pagination

- The offline foundation verifies generated drive collection request information and binary normalization, not complete path-based `/content` upload/download behavior.
- Multi-page Graph traversal, authenticated next-link validation, cycle detection, and global result caps remain open.
- The 10 MiB simple-transfer product contract is retained in metadata, but large-file upload sessions are not implemented.

## Delivery

- The repository is not yet in the Hermes plugin catalog and has no SHA-pinned catalog entry.
- Plugin validation/doctor are CI gates only where the pinned Hermes release exposes the commands compatibly.
- No production support, release readiness, or remote tenant behavior is claimed.
