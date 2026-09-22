# Review and acceptance policy

This document defines who may merge into `main`, what must pass before a merge,
and the approval flow for external contributions. It is the policy behind the
`main` branch protection rule; the rule enforces it mechanically.

## Roles

- **Maintainer** — `fabiomotta0311`. Sole owner, final approval level, and the
  only account listed in `.github/CODEOWNERS`. A maintainer is the last reviewer
  for every external pull request and is the only role that merges into `main`.
- **Contributor** — anyone opening a pull request. Contributors never merge
  directly; their changes enter `main` only through an approved pull request.

## Acceptance criteria (definition of done)

A pull request is mergeable only when **all** of the following hold:

1. **Scope** — the change is focused, conventional (`feat:`, `fix:`, `test:`,
   `docs:`, `chore:`), and preserves the operation catalog in
   `microsoft365.contract.OPERATIONS` unless product authorization says otherwise.
2. **Boundary honesty** — it does not claim tenant verification or production
   support, and CORE-1/CORE-2 and tenant-verification blockers remain accurately
   documented.
3. **No secrets or PII** — no payload, credentials, tenant/user identifiers, or
   secret-bearing fixtures were added; no telemetry or logging that could expose
   them.
4. **CI green** — the `ci-gate` check passes, which requires the full test matrix,
   strict SDK contract tests, `--cov-fail-under=90`, plugin validation, doctor,
   compileall, wheel build, isolated smoke, pip-audit, secret scan, and diff
   hygiene to all pass (see `.github/workflows/ci.yml`).
5. **Tests** — runtime behavior is covered by tests written test-first; SDK
   contract tests use real pinned `msgraph-sdk` builders, not permissive
   `__getattr__` fakes.
6. **Review approval** — for a contributor pull request, the maintainer
   (`fabiomotta0311`) approves it as Code Owner. For a maintainer's own pull
   request, no external approval is required (see below).
7. **Conversation resolved** — open review threads are resolved before merge.

## Approval flow

### External (contributor) pull requests

1. Contributor opens a pull request against `main` using the PR template.
2. CI runs `ci-gate`; it must be green before merge.
3. `CODEOWNERS` requests review from the maintainer, who is the final approval
   level. A stale approval (after new commits) is dismissed and must be re-given.
4. The maintainer resolves all threads, approves, and merges. Contributors cannot
   merge directly, force-push, or delete `main`.

### Maintainer's own pull requests

The maintainer may merge their own pull requests without an external approval
(`enforce_admins` is disabled), but the CI gate still applies. This keeps the
maintainer's velocity while keeping external changes gated behind review.

## Branch protection summary (`main`)

- Pull request required before merge; direct push to `main` is blocked.
- Required status check: `ci-gate`.
- One required approving review from Code Owners; stale reviews dismissed.
- Force pushes and branch deletion are blocked.
- `enforce_admins` is **off** (maintainer retains self-merge ability).

## Changing this policy

Changes to `.github/CODEOWNERS`, the PR template, `ci.yml`, or this file must go
through a pull request and the same CI gate. The branch protection rule itself is
GitHub repository state managed by the maintainer; it is not versioned in the
repository and must be kept consistent with this document.
