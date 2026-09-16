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
| Planner list/read | `Tasks.Read.All` endpoint-specific | supported, implementation incomplete | not implemented |
| Planner create/update | `Tasks.ReadWrite.All` | supported, implementation incomplete | not implemented |

Official endpoint references:

- https://learn.microsoft.com/en-us/graph/api/search-query
- https://learn.microsoft.com/en-us/graph/api/user-list-joinedteams
- https://learn.microsoft.com/en-us/graph/api/channel-list
- https://learn.microsoft.com/en-us/graph/api/todotasklist-list-tasks
- https://learn.microsoft.com/en-us/graph/api/planner-post-tasks
- https://learn.microsoft.com/en-us/graph/api/plannertask-update
- https://learn.microsoft.com/en-us/graph/api/plannerplan-list-tasks
