# Known limitations and remaining blockers

This document states the boundary of the first standalone milestone. It is not release evidence.

## Runtime and authentication

- Only application/client-credentials client construction is implemented. Delegated OAuth, secure delegated token caching, and account-selection UX remain open.
- Authentication, tenant consent, endpoint permissions, throttling, pagination, and write cleanup have not been tested against an authorized test tenant.
- Operations lacking a strict executable contract remain in the administrative registry and preflight status; they are not deleted and are not registered in active action enums. Exactly five operations are executable today -- `outlook.search`, `outlook.read`, `calendar.search`, `teams.list_teams` and `teams.list_channels` -- and every write stays non-executable until the host approval fix under "Host dependency" lands.
- Microsoft Search chat-message support and large-file transfer sessions require further endpoint-specific work.

## Planner

- Planner containers, request bodies and `If-Match` handling are pinned by strict offline request-contract tests (`tests/test_planner_contracts.py`): the tenant-wide and group-owned plan listings, the plan-scoped bucket and task listings, the task-scoped read and update, `planId`/`bucketId` in the creation body, and a mandatory ETag on `planner.update_tasks`. A blank required identifier, and a blank optional `group_id`, are refused before any client, approval or token work, with the endpoint the identifier belongs to named in the error.
- No Planner operation is executable yet. All six Planner operations stay in the registry with `implementation_status = contract_verified` and `executable = False`, and no handler exists for them, so a Planner call cannot reach Microsoft Graph through this plugin.
- Planner permission claims are endpoint-specific and labelled `documented_not_verified`: each claim cites the endpoint it belongs to, no endpoint page was re-read offline, and no tenant has been contacted (R10, WP16). Assignments on task creation are not modelled, and any additional role that endpoint may require is unverified.
- The shared partial-model helper (`microsoft365/sdk_contract.py`) cannot express a PATCH field whose value is a *list* of generated models (for example `attendees` on `calendar.update_events`). It refuses such a field instead of silently dropping it. This applies to the Planner and To Do contract cases that use it: `calendar.update_events` does **not** -- its handler builds its own partial `Event` from real generated models, so a list-valued field is accepted and a test proves every accepted field changes the serialized request.

## Outlook and Calendar

- Outlook and Calendar request semantics are pinned by strict offline request-contract tests (`tests/test_handlers_outlook_calendar.py`): each operation is rebuilt through the generated builders and the handler's executed request is compared with the `sdk_contract` request of the same operation byte for byte -- method, URL, typed `$top`/`$filter`/`$select`, serialized body and `If-Match`.
- The endpoint rows of both services live in `microsoft365/contract.py` (`OUTLOOK_ENDPOINTS`, `CALENDAR_ENDPOINTS`), each carrying the container, identifier set, application role, required header and contract case it belongs to. A test rebuilds every declared row and compares the rendered method and path with the declaration, so a declared endpoint cannot be invented. No documentation page is recorded for any row: no endpoint reference was re-read offline and no tenant was contacted (R10, WP16), so every claim is `documented_not_verified`.
- Exactly three operations are executable: `outlook.search`, `outlook.read` and `calendar.search`. They are the only operations that appear in a model-facing action enum, and their registry status is `implemented`.
- The four writes -- `outlook.create_draft`, `outlook.send`, `calendar.create_events`, `calendar.update_events` -- are implemented, contract-pinned and exercised offline through the execution seam, and deliberately **non-executable**: `active_actions()` never returns them, they appear in no action enum, and both the hook and the dispatch path refuse them. They are not dead code -- the `HANDLER_TABLE` entries for them exist and the tests drive each one to a real request through the strict double -- and the flip is gated on the host approval fix below (R5, D1).
- `microsoft365.registration.HANDLER_TABLE` is filled from the self-populating handler registry (`microsoft365/handlers/`, `load_handlers()`): adding `handlers/<service>.py` registers a service without editing `registration.py`. The table holds **every** implemented operation, writes included, so membership in it is never a permission to run.
- The dispatch path re-evaluates executability on the **final arguments** (`operation_is_executable`, invariant 15) with the same `active_actions()`/`operation_status()` gate the `pre_tool_call` hook uses. A payload whose `action` is changed from a read into a write after a hook approved the read is refused before any handler, secret, credential, client or request work, with the same message the hook returns. This is the plugin-side mitigation of the host defect below, not a replacement for the core fix.
- `outlook.send` has two endpoints -- `/users/{user_id}/sendMail` for a composed message and `/users/{user_id}/messages/{message_id}/send` for a draft that already exists -- and supplying both forms, or neither, is refused instead of silently sending one.
- An offsetless date and time must be accompanied by `time_zone`, and `calendar.search` refuses `start_date_time`/`end_date_time`/`time_zone` because the endpoint's acceptance of a date filter is not recorded offline; `calendar.update_events` requires the ETag it cannot be conditional without.
- `calendar.search` and `outlook.search` apply the caller's item budget through the paginator, so the first multi-page paths in the plugin are the two executable reads.

## Teams Graph

- Teams Graph handlers are registered for `list_teams`, `list_channels`, `search_messages` and `send_messages`. The first two are implemented and executable in application mode through `users/{user_id}/joinedTeams` and `teams/{team_id}/channels`.
- `joinedTeams` does not accept OData query parameters, and the channels endpoint does not accept `$top`; the handlers process returned collections within the shared item budget. Channels may use typed `$select`.
- `search_messages` uses the generated `search/query` POST body (`QueryPostRequestBody` containing `SearchRequest`/`SearchQuery` for `chatMessage`), and `send_messages` uses a generated `ChatMessage` on the channel `messages` collection. Both remain delegated-only: application mode has no permission claim and dispatch refuses them before client or credential work; delegated scopes are recorded as `Chat.Read` + `ChannelMessage.Read.All` and `ChannelMessage.Send`, respectively.
- Delegated authentication is not implemented until WP14, so delegated Teams operations are described in the matrix but are not executable. Bot Framework files (`plugins/platforms/teams` and `teams_pipeline`) are intentionally untouched.

## Microsoft To Do

- To Do request semantics are pinned by strict offline request-contract tests (`tests/test_todo_contracts.py`): one endpoint per operation with its container and identifier set, the `contains(title,'…')` task filter -- never the Outlook `subject` field, which no To Do model has -- and a mandatory non-blank `user_id`/`todo_list_id`/`todo_task_id` before any request is built.
- A To Do write may carry only `title`, `body`, `due_date_time` and `status`. Every other property of the generated `TodoTask` model, and any field the model does not declare, is refused with the SDK case name instead of being sent and silently dropped by Graph. `status` is mapped to the generated `TaskStatus` member when accepted, because Kiota's JSON writer omits an enum field that is not a generated member.
- Filtering by `status` is not sent as an OData filter: the endpoint's acceptance of one is not recorded offline, so matching stays client-side and limited to the tasks a caller has already retrieved (`todo_task_status_matches()`).
- No Microsoft To Do operation is executable yet. All five To Do operations stay in the registry with `implementation_status = contract_verified` and `executable = False`, and no handler exists for them, so a To Do call cannot reach Microsoft Graph through this plugin.
- To Do permission claims are endpoint-specific and labelled `documented_not_verified`: each claim cites the endpoint it belongs to, only the task-listing endpoint carries a recorded page reference, no endpoint page was re-read offline, and no tenant has been contacted (R10, WP16). Application-mode write support for `todo.create_tasks` and `todo.update_tasks` is `not_verified` and is derived from those endpoint claims, so it cannot be promoted without recorded endpoint evidence -- promoting it fails a test. Whether the task update endpoint requires an `If-Match` ETag is not recorded, so the plugin sends no required header and adds `If-Match` only when the caller supplies one.

## Host dependency

Hermes currently has a generic approval modification-order defect described in the governing plan: an earlier hook can modify arguments after another hook evaluated approval against different arguments. This repository does not add a Microsoft-specific workaround. Write execution must remain disabled until the separate generic Hermes core fix binds approval to the final arguments that execute.

## Files and pagination

- Path-based drive `/content` transfer is implemented with validated relative-path addressing, the 10 MiB simple-transfer bound enforced before allocation in both directions, and byte-exact round trips (`microsoft365/files.py`, WP4). Large-file upload sessions and ranged download are not implemented and are reported explicitly (`upload_session_not_implemented`, `download_range_not_implemented`). No executable operation consumes it yet.
- Multi-page Graph traversal is implemented as an offline helper with authenticated next-link validation, cycle detection and a global page cap (`microsoft365/paging.py`, WP5). The two executable reads consume it (`outlook.search`, `calendar.search`), so a traversal follows a next link as soon as those calls can run; multi-page behavior against a real tenant remains unverified.
- The 10 MiB simple-transfer contract is enforced in code; the large-file upload-session request shape is contract-verified but deliberately not executed.

## Delivery

- The repository is not yet in the Hermes plugin catalog and has no SHA-pinned catalog entry.
- Plugin validation/doctor are CI gates only where the pinned Hermes release exposes the commands compatibly.
- No production support, release readiness, or remote tenant behavior is claimed.
