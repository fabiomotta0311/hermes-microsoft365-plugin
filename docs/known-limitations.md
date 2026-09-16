# Known limitations and remaining blockers

This document states the boundary of the first standalone milestone. It is not release evidence.

## Runtime and authentication

- Only application/client-credentials client construction is implemented. Delegated OAuth, secure delegated token caching, and account-selection UX remain open.
- Authentication, tenant consent, endpoint permissions, throttling, pagination, and write cleanup have not been tested against an authorized test tenant.
- Operations lacking a strict executable contract remain in the administrative registry and preflight status; they are not deleted and are not registered in active action enums.
- Microsoft Search chat-message support and large-file transfer sessions require further endpoint-specific work.

## Planner

- Planner containers, request bodies and `If-Match` handling are pinned by strict offline request-contract tests (`tests/test_planner_contracts.py`): the tenant-wide and group-owned plan listings, the plan-scoped bucket and task listings, the task-scoped read and update, `planId`/`bucketId` in the creation body, and a mandatory ETag on `planner.update_tasks`. A blank required identifier, and a blank optional `group_id`, are refused before any client, approval or token work, with the endpoint the identifier belongs to named in the error.
- No Planner operation is executable yet. All six Planner operations stay in the registry with `implementation_status = contract_verified` and `executable = False`, and no handler exists for them, so a Planner call cannot reach Microsoft Graph through this plugin.
- Planner permission claims are endpoint-specific and labelled `documented_not_verified`: each claim cites the endpoint it belongs to, no endpoint page was re-read offline, and no tenant has been contacted (R10, WP16). Assignments on task creation are not modelled, and any additional role that endpoint may require is unverified.
- The shared partial-model helper cannot express a PATCH field whose value is a *list* of generated models (for example `attendees` on `calendar.update_events`). It refuses such a field instead of silently dropping it, so affected updates stay unimplemented until the helper is extended deliberately.

## Microsoft To Do

- To Do request semantics are pinned by strict offline request-contract tests (`tests/test_todo_contracts.py`): one endpoint per operation with its container and identifier set, the `contains(title,'…')` task filter -- never the Outlook `subject` field, which no To Do model has -- and a mandatory non-blank `user_id`/`todo_list_id`/`todo_task_id` before any request is built.
- A To Do write may carry only `title`, `body`, `due_date_time` and `status`. Every other property of the generated `TodoTask` model, and any field the model does not declare, is refused with the SDK case name instead of being sent and silently dropped by Graph. `status` is mapped to the generated `TaskStatus` member when accepted, because Kiota's JSON writer omits an enum field that is not a generated member.
- Filtering by `status` is not sent as an OData filter: the endpoint's acceptance of one is not recorded offline, so matching stays client-side and limited to the tasks a caller has already retrieved (`todo_task_status_matches()`).
- No Microsoft To Do operation is executable yet. All five To Do operations stay in the registry with `implementation_status = contract_verified` and `executable = False`, and no handler exists for them, so a To Do call cannot reach Microsoft Graph through this plugin.
- To Do permission claims are endpoint-specific and labelled `documented_not_verified`: each claim cites the endpoint it belongs to, only the task-listing endpoint carries a recorded page reference, no endpoint page was re-read offline, and no tenant has been contacted (R10, WP16). Application-mode write support for `todo.create_tasks` and `todo.update_tasks` is `not_verified` and is derived from those endpoint claims, so it cannot be promoted without recorded endpoint evidence -- promoting it fails a test. Whether the task update endpoint requires an `If-Match` ETag is not recorded, so the plugin sends no required header and adds `If-Match` only when the caller supplies one.

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
