# Known limitations and remaining blockers

This document states the boundary of the first standalone milestone. It is not release evidence.

## Runtime and authentication

- Only application/client-credentials client construction is implemented. Delegated OAuth, secure delegated token caching, and account-selection UX remain open.
- Exactly eighteen reads are executable in the current application-mode matrix, plus the explicitly gated `outlook.create_draft`, `outlook.send` and `sharepoint.upload_files` writes; all remote verification remains pending.
- Authentication, tenant consent, endpoint permissions, throttling, pagination, and write cleanup have not been tested against an authorized test tenant.
- Operations lacking a strict executable contract remain in the administrative registry and preflight status; they are not deleted and are not registered in active action enums. The eighteen verified reads plus the gated `outlook.create_draft` pilot are executable in application mode. Other writes stay non-executable until separately promoted.
- Microsoft Search chat-message support and large-file transfer sessions require further endpoint-specific work.

## Planner

- Planner containers, request bodies and `If-Match` handling are pinned by strict offline request-contract tests (`tests/test_planner_contracts.py`): the tenant-wide and group-owned plan listings, the plan-scoped bucket and task listings, the task-scoped read and update, `planId`/`bucketId` in the creation body, and a mandatory ETag on `planner.update_tasks`. A blank required identifier, and a blank optional `group_id`, are refused before any client, approval or token work, with the endpoint the identifier belongs to named in the error.
- The four Planner reads are executable through synchronous handlers and the shared paging/execution seams; `planner.create_tasks` and `planner.update_tasks` have handlers and strict real-model contracts but remain non-executable until CORE-1/CORE-2.
- Planner permission claims are endpoint-specific and labelled `documented_not_verified`: each claim cites the endpoint it belongs to, no endpoint page was re-read offline, and no tenant has been contacted (R10, WP16). Assignments on task creation are not modelled, and any additional role that endpoint may require is unverified.
- The shared partial-model helper (`microsoft365/sdk_contract.py`) cannot express a PATCH field whose value is a *list* of generated models (for example `attendees` on `calendar.update_events`). It refuses such a field instead of silently dropping it. This applies to the Planner and To Do contract cases that use it: `calendar.update_events` does **not** -- its handler builds its own partial `Event` from real generated models, so a list-valued field is accepted and a test proves every accepted field changes the serialized request.

## Outlook and Calendar

- Outlook and Calendar request semantics are pinned by strict offline request-contract tests (`tests/test_handlers_outlook_calendar.py`): each operation is rebuilt through the generated builders and the handler's executed request is compared with the `sdk_contract` request of the same operation byte for byte -- method, URL, typed `$top`/`$filter`/`$select`, serialized body and `If-Match`.
- The endpoint rows of both services live in `microsoft365/contract.py` (`OUTLOOK_ENDPOINTS`, `CALENDAR_ENDPOINTS`), each carrying the container, identifier set, application role, required header and contract case it belongs to. A test rebuilds every declared row and compares the rendered method and path with the declaration, so a declared endpoint cannot be invented. No documentation page is recorded for any row: no endpoint reference was re-read offline and no tenant was contacted (R10, WP16), so every claim is `documented_not_verified`.
- Three of the nine executable reads belong to Outlook and Calendar: `outlook.search`, `outlook.read` and `calendar.search`. Their registry status is `implemented`.
- `outlook.create_draft` is the first executable write pilot. It is contract-pinned, uses `Mail.ReadWrite`, requires the configured single-user boundary, and remains subject to the host final-argument approval patch and remote tenant smoke test. `outlook.send` is now executable behind the same approval gate; it supports composed-message and draft-send forms and is irreversible, so tenant smoke tests must use an explicitly approved recipient. Calendar writes and all other writes remain non-executable.
- `microsoft365.registration.HANDLER_TABLE` is filled from the self-populating handler registry (`microsoft365/handlers/`, `load_handlers()`): adding `handlers/<service>.py` registers a service without editing `registration.py`. The table holds **every** implemented operation, writes included, so membership in it is never a permission to run.
- The dispatch path re-evaluates executability on the **final arguments** (`operation_is_executable`, invariant 15) with the same `active_actions()`/`operation_status()` gate the `pre_tool_call` hook uses. A payload whose `action` is changed from a read into a write after a hook approved the read is refused before any handler, secret, credential, client or request work, with the same message the hook returns. This is the plugin-side mitigation of the host defect below, not a replacement for the core fix.
- `outlook.send` has two endpoints -- `/users/{user_id}/sendMail` for a composed message and `/users/{user_id}/messages/{message_id}/send` for a draft that already exists -- and supplying both forms, or neither, is refused instead of silently sending one.
- An offsetless date and time must be accompanied by `time_zone`, and `calendar.search` refuses `start_date_time`/`end_date_time`/`time_zone` because the endpoint's acceptance of a date filter is not recorded offline; `calendar.update_events` requires the ETag it cannot be conditional without.
- `calendar.create_events` accepts the optional Graph `transaction_id` and serializes it as
  `transactionId`. Clients should provide a stable value when their workflow may retry a create;
  this improves duplicate protection but does not itself enable writes or replace host approval.
- `calendar.search` and `outlook.search` apply the caller's item budget through the paginator.

## SharePoint and OneDrive

- SharePoint and OneDrive request semantics are pinned by strict offline request-contract tests (`tests/test_handlers_files.py`): each operation is rebuilt through the generated builders and the handler's executed request is compared with the `sdk_contract` request of the same operation -- method, URL, typed query, the `search(q='…')` path, the site→drive resolution path, the item read path and the `/content` transfer path.
- The endpoint rows of both services live in `microsoft365/contract.py` (`DRIVE_ENDPOINTS`), each carrying the container, identifier set, application role and contract case it belongs to. A test rebuilds every declared row and compares the rendered method and path with the declaration, so a declared endpoint cannot be invented. No documentation page is recorded for any row: no endpoint reference was re-read offline and no tenant was contacted (R10, WP16), so every claim is `documented_not_verified`.
- `sharepoint.search` scopes a search by one container at most: with no container it searches **sites** (`GET /sites?$search=…`), with `drive_id` it searches that drive (`GET /drives/{drive_id}/search(q='…')`), and with `site_id` it resolves the site to its default document library (`GET /sites/{site_id}/drive`) before searching the resolved drive. `onedrive.search` is drive-scoped only: a `site_id` is refused and a search without `drive_id` is refused, because OneDrive is not site-scoped and there is no tenant-wide drive search in application mode.
- `read` addresses one drive item by an opaque `item_id` or a validated relative `item_path` through `microsoft365.files.drive_item_address`; the generated builder renders the path and percent-encodes reserved and Unicode characters, never this plugin.
- Six of the nine executable reads belong to these two services: `sharepoint.search`, `sharepoint.read`, `sharepoint.download_files`, `onedrive.search`, `onedrive.read` and `onedrive.download_files`.
- `sharepoint.upload_files` and `onedrive.upload_files` are executable through the final-argument approval gate for **any** file size. Small `replace` uploads use one simple `PUT /content`; anything larger -- or `conflict_behavior` `fail`/`rename`, which a simple PUT cannot express -- is routed through the resumable upload session. `conflict_behavior` is the definitive argument name; the boolean `overwrite` is kept as a compatibility alias (`true` = `replace`, `false` = `fail`). `chunk_size` is optional and must be a multiple of 320 KiB.

- `active_actions()` never returns them for non-executable operations; promoted writes are the explicit exception.

## Teams Graph

- Teams Graph handlers are registered for `list_teams`, `list_channels`, `search_messages`, `send_messages`, `list_chats`, `read_chat_messages` and `send_chat_message`. The first two are implemented and executable in application mode through `users/{user_id}/joinedTeams` and `teams/{team_id}/channels`.
- The individual-chat handlers use `GET /users/{user_id}/chats`, `GET /chats/{chat_id}/messages` and `POST /chats/{chat_id}/messages`. They preserve chat context via the caller-supplied `chat_id`, use shared pagination, and serialize real `ChatMessage`/`ItemBody` models. Chat listing and message reading are executable in application mode with `Chat.ReadBasic.All` and `Chat.Read.All`; sending remains delegated-only with `ChatMessage.Send`.
- `joinedTeams` does not accept OData query parameters, and the channels endpoint does not accept `$top`; the handlers process returned collections within the shared item budget. Channels may use typed `$select`.
- `search_messages` uses the generated `search/query` POST body (`QueryPostRequestBody` containing `SearchRequest`/`SearchQuery` for `chatMessage`), and `send_messages` uses a generated `ChatMessage` on the channel `messages` collection. Both remain delegated-only: application mode has no permission claim and dispatch refuses them before client or credential work; delegated scopes are recorded as `Chat.Read` + `ChannelMessage.Read.All` and `ChannelMessage.Send`, respectively.
- Delegated authentication uses Azure Identity device-code OAuth with an OS-protected persistent token cache. Delegated Teams operations are now executable according to their contract; Bot Framework files (`plugins/platforms/teams` and `teams_pipeline`) remain intentionally untouched because transport is a separate Hermes integration boundary.

## Microsoft To Do

- To Do request semantics are pinned by strict offline request-contract tests (`tests/test_todo_contracts.py`): one endpoint per operation with its container and identifier set, the `contains(title,'…')` task filter -- never the Outlook `subject` field, which no To Do model has -- and a mandatory non-blank `user_id`/`todo_list_id`/`todo_task_id` before any request is built.
- A To Do write may carry only `title`, `body`, `due_date_time` and `status`. Every other property of the generated `TodoTask` model, and any field the model does not declare, is refused with the SDK case name instead of being sent and silently dropped by Graph. `status` is mapped to the generated `TaskStatus` member when accepted, because Kiota's JSON writer omits an enum field that is not a generated member.
- Filtering by `status` is not sent as an OData filter: the endpoint's acceptance of one is not recorded offline, so matching stays client-side and limited to the tasks a caller has already retrieved (`todo_task_status_matches()`).
- The three To Do reads are executable through synchronous handlers and the shared paging/execution seams; `todo.create_tasks` and `todo.update_tasks` have handlers and strict real-model contracts but remain non-executable until CORE-1/CORE-2.
- To Do permission claims are endpoint-specific and labelled `documented_not_verified`: each claim cites the endpoint it belongs to, only the task-listing endpoint carries a recorded page reference, no endpoint page was re-read offline, and no tenant has been contacted (R10, WP16). Application-mode write support for `todo.create_tasks` and `todo.update_tasks` is `not_verified` and is derived from those endpoint claims, so it cannot be promoted without recorded endpoint evidence -- promoting it fails a test. Whether the task update endpoint requires an `If-Match` ETag is not recorded, so the plugin sends no required header and adds `If-Match` only when the caller supplies one.

- Executable operation inventory: `outlook.search`, `outlook.read`, `outlook.create_draft`, `outlook.send`, `calendar.search`, `sharepoint.search`, `sharepoint.read`, `sharepoint.download_files`, `sharepoint.upload_files`, `onedrive.search`, `onedrive.read`, `onedrive.download_files`, `onedrive.upload_files`, `teams.list_teams`, `teams.list_channels`, `teams.list_chats`, `teams.read_chat_messages`, `todo.list_task_lists`, `todo.search`, `todo.read`, `planner.list_buckets`, `planner.list_plans`, `planner.list_tasks`, `planner.read`.

Hermes currently has a generic approval modification-order defect described in the governing plan: an earlier hook can modify arguments after another hook evaluated approval against different arguments. This repository does not add a Microsoft-specific workaround. Write execution must remain disabled until the separate generic Hermes core fix binds approval to the final arguments that execute.
A tested core patch for that dependency is included in [`docs/hermes-core-approval-fix.md`](hermes-core-approval-fix.md), but it is not applied automatically to the user's Hermes installation.

## Files and pagination

- Path-based drive `/content` transfer is implemented with validated relative-path addressing, the 10 MiB simple-transfer bound enforced before allocation in both directions, and byte-exact round trips (`microsoft365/files.py`, WP4). Resumable upload sessions lift the 10 MiB ceiling for **uploads** (`microsoft365/upload_session.py`: `createUploadSession`, 320 KiB-aligned fragments, resume from Graph's `nextExpectedRanges`, bounded retries, explicit `@odata.conflictBehavior`). **Download** of a file above the bound is still not implemented and is reported explicitly (`download_range_not_implemented`): range reads remain the missing half.
- Multi-page Graph traversal is implemented as an offline helper with authenticated next-link validation, cycle detection and a global page cap (`microsoft365/paging.py`, WP5). The four executable search reads consume it (`outlook.search`, `calendar.search`, `sharepoint.search`, `onedrive.search`), so a traversal follows a next link as soon as those calls can run; multi-page behavior against a real tenant remains unverified.
- The 10 MiB simple-transfer contract is enforced in code for the one-shot path. Above it, uploads execute through the resumable session up to Graph's own 250 GiB per-session ceiling; a transfer that never completes is reported as `upload_session_incomplete`, never as a success.
- Remote verification is `not_tested` for every operation: no tenant has been contacted (R10; WP16 is the only work package allowed to record evidence).

## Delivery

- The repository has a SHA-pinned catalog proposal in upstream PR #114530, but it is not official until that PR is merged.
- Plugin validation/doctor are CI gates only where the pinned Hermes release exposes the commands compatibly.
- No production support or remote tenant behavior is claimed; the current release remains pre-release.

## Microsoft 365 Files: upload sessions (PRO)

- `microsoft365/upload_session.py` owns the resumable transfer: `POST .../createUploadSession`
  with a real `CreateUploadSessionPostRequestBody`, then one `PUT` per 320 KiB-aligned fragment to
  the session's pre-authorized URL, resuming from the `nextExpectedRanges` Graph reports.
- The `uploadUrl` is bearer-equivalent. It never appears in a result, a log line or an exception;
  every failure is a sanitized taxonomy error with an explicit `status`.
- No tenant credential is attached to fragment PUTs: the URL is already authorized, and widening
  the blast radius to a second origin is not done.
- A stuck session cannot spin forever: consecutive replays and total fragments are both bounded.
- Not verified against a tenant: fragment sizing, `nextExpectedRanges` handling and the committed
  `DriveItem` are proven offline only. No remote evidence exists yet.
