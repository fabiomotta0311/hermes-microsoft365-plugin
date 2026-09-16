# Microsoft Graph permission matrix

This is a conservative pre-release matrix for local administrative reporting. It does not assert tenant consent or remote success. Runtime metadata in `microsoft365/contract.py` is authoritative and tested against the complete operation list.

| Service / operation | Application role(s) | Application status | Delegated status |
|---|---|---|---|
| Outlook search/read | `Mail.Read` | supported | not implemented |
| Outlook draft | `Mail.ReadWrite` | supported | not implemented |
| Outlook send | `Mail.Send` | supported | not implemented |
| SharePoint search/read | `Sites.Read.All` | supported | not implemented |
| SharePoint download/upload | `Files.Read.All` / `Files.ReadWrite.All` | supported | not implemented |
| OneDrive search/read/download/upload | `Files.Read.All` / `Files.ReadWrite.All` | supported | not implemented |
| Calendar search/create/update | `Calendars.Read` / `Calendars.ReadWrite` | supported | not implemented |
| Teams list teams/channels | `Team.ReadBasic.All` / `Channel.ReadBasic.All` | supported | not implemented |
| Teams message search | none claimed | unsupported auth mode | requires delegated implementation |
| Teams send message | none claimed | unsupported auth mode | requires delegated implementation |
| To Do list/search/read | `Tasks.Read.All` | supported | not implemented |
| To Do create/update | `Tasks.ReadWrite.All` candidate | not verified | requires endpoint verification |
| Planner (6 operations) | endpoint-specific, see "Planner endpoint claims" | contract verified offline, not executable | not implemented |

## Planner endpoint claims

Planner is reported one endpoint at a time. The blanket `Planner list/read` and `Planner create/update`
rows this matrix used to carry are gone: `Tasks.Read.All` and `Tasks.ReadWrite.All` are not a single
claim, each endpoint carries its own. `microsoft365/contract.py` declares one `PlannerEndpoint` per row
below and derives each operation's permission tuple, container, required identifiers and required
headers from those endpoints, so a claim cannot drift from the endpoint it belongs to, and
`tests/test_planner_contracts.py` fails if this table and the registry disagree.

| Operation | Endpoint | Application role | Permission claim | Implementation status | Recorded endpoint reference |
|---|---|---|---|---|---|
| `planner.list_plans` | `GET /planner/plans` | `Tasks.Read.All` | `documented_not_verified` | `contract_verified` | not recorded |
| `planner.list_plans` | `GET /groups/{group_id}/planner/plans` | `Tasks.Read.All` | `documented_not_verified` | `contract_verified` | not recorded |
| `planner.list_buckets` | `GET /planner/plans/{plan_id}/buckets` | `Tasks.Read.All` | `documented_not_verified` | `contract_verified` | not recorded |
| `planner.list_tasks` | `GET /planner/plans/{plan_id}/tasks` | `Tasks.Read.All` | `documented_not_verified` | `contract_verified` | https://learn.microsoft.com/en-us/graph/api/plannerplan-list-tasks |
| `planner.read` | `GET /planner/tasks/{planner_task_id}` | `Tasks.Read.All` | `documented_not_verified` | `contract_verified` | not recorded |
| `planner.create_tasks` | `POST /planner/tasks` | `Tasks.ReadWrite.All` | `documented_not_verified` | `contract_verified` | https://learn.microsoft.com/en-us/graph/api/planner-post-tasks |
| `planner.update_tasks` | `PATCH /planner/tasks/{planner_task_id}` | `Tasks.ReadWrite.All` | `documented_not_verified` | `contract_verified` | https://learn.microsoft.com/en-us/graph/api/plannertask-update |

### Container semantics (decision recorded by WP2)

A Planner request is only valid against one container, and the identifiers that identify that container
must be non-blank before anything is sent:

- `planner.list_plans` accepts an **optional** `group_id`. Omitted (or `None`), it reads the tenant-wide
  route `GET /planner/plans`; supplied, it reads the group-owned route
  `GET /groups/{group_id}/planner/plans`. A blank `group_id` is refused -- it never falls back silently
  to the tenant-wide route.
- `planner.list_buckets` and `planner.list_tasks` index by **plan**: `plan_id` is a required path
  identifier, and the path has no calling form without it.
- `planner.read` and `planner.update_tasks` index by **task**: `planner_task_id` is a required path
  identifier.
- `planner.create_tasks` carries both container identifiers in the **request body** (`planId`,
  `bucketId`) together with the task title: creation is bucket-scoped, and the path `POST /planner/tasks`
  carries no identifier at all.
- `planner.update_tasks` additionally requires an `If-Match` ETag header. The other endpoints here
  (`planner.list_plans`, `planner.list_buckets`, `planner.list_tasks`, `planner.read`,
  `planner.create_tasks`) declare no required header and send no `If-Match`.

`validate_planner_identifiers()` in `microsoft365/contract.py` refuses a blank or absent required
identifier -- and a blank optional one -- naming the endpoint the identifier belongs to, before any
client, approval or token work happens.

### Claim legend

- `documented_not_verified` -- the application role recorded for that endpoint. No endpoint page was
  re-read in this offline milestone and no tenant has been contacted (rule R10; remote verification is
  WP16), so the claim is **not** verified. Roles on rows without a recorded reference are inherited from
  this matrix's earlier revision and carry no page citation yet; every claim must be recomputed from its
  endpoint's permission table before the catalog release.
- `verified` -- reserved. Only for a claim recomputed against the endpoint reference **and** confirmed on
  a test tenant. No Planner claim may use it yet.

### Implementation status legend

- `contract_verified` -- every endpoint the operation uses is pinned by a strict offline request-contract
  test (`tests/test_planner_contracts.py`); no handler exists, so the operation is **not executable** and
  cannot reach Graph through this plugin.
- `implemented` -- reserved. Requires a registered handler **and** a strict contract test; no Planner
  operation qualifies yet.

Official endpoint references:

- https://learn.microsoft.com/en-us/graph/api/search-query
- https://learn.microsoft.com/en-us/graph/api/user-list-joinedteams
- https://learn.microsoft.com/en-us/graph/api/channel-list
- https://learn.microsoft.com/en-us/graph/api/todotasklist-list-tasks
- https://learn.microsoft.com/en-us/graph/api/planner-post-tasks
- https://learn.microsoft.com/en-us/graph/api/plannertask-update
- https://learn.microsoft.com/en-us/graph/api/plannerplan-list-tasks
