# Consent and tenant readiness

## Scope

This is a readiness checklist for a maintainer working with an authorized test tenant. It is not evidence that this repository has contacted or verified a tenant. The current release remains Alpha-only and has no production support boundary.

## Preconditions

1. Confirm the exact plugin tag/commit and Python version.
2. Confirm the test tenant, application registration, operator, and change window are authorized by the tenant owner.
3. Keep credentials in Hermes secret handling. Never place secrets in settings files, issues, logs, shell history, or fixtures.
4. Record only sanitized metadata: tenant alias (not an ID), authentication mode, operation name, status code, and UTC timestamp.
5. Confirm CORE-1 and CORE-2 are resolved before considering any write path; until then, writes remain non-executable.

## Consent procedure

1. Obtain tenant-owner approval for the exact Microsoft Graph application permissions needed by the operation under test.
2. Review the endpoint-specific permission evidence in `microsoft365/references/graph-permissions.md`; treat `documented_not_verified` as unverified.
3. Grant admin consent through the tenant's approved Microsoft Entra process, outside this repository.
4. Verify the application registration and permission state using the tenant's approved administrative process. Do not paste screenshots, IDs, tokens, or Graph responses into this repository.
5. Run one minimal read operation with synthetic/non-sensitive test data. Capture only sanitized success/failure metadata.
6. Test paging, throttling, and error handling only within the approved limits. Do not use real user or business content.
7. Revoke test consent and credentials when the approved test window ends, if required by the tenant owner.

## Stop conditions

Stop immediately on missing consent, unexpected permissions, a token/secret exposure, a write request before CORE-1/CORE-2, or any response that includes data outside the approved synthetic scope. Follow [incident response](incident-response.md) for exposure and [graph errors](graph-errors.md) for sanitized status triage.

## Evidence boundary

Offline tests and plugin validation prove repository behavior only. Tenant verification requires separately approved, repeatable evidence and must not be claimed by a release until the tenant-verification blocker is closed.
