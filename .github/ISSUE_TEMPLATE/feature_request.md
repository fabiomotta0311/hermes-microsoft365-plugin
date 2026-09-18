---
name: Feature request
about: Propose a repository or documentation improvement within the Alpha boundary
title: ""
labels: "enhancement"
assignees: ""
---

## Support boundary

- [ ] I understand this project is Alpha-only and not production or tenant support.
- [ ] This proposal does not assume CORE-1, CORE-2, or tenant verification is complete.

## Problem and proposal

What problem should be solved, and what is the smallest safe change?

## Safety and scope

- Does this touch Microsoft Graph permissions, credentials, logging, or telemetry?
- How will payloads, tenant identifiers, and PII remain out of diagnostics and fixtures?
- Does it require tenant verification or a Hermes core change?

## Acceptance criteria

- [ ] Offline behavior and documentation boundaries are testable.
- [ ] No secrets, tenant data, payloads, or PII are included.
