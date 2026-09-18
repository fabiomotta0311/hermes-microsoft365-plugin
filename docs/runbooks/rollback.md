# Rollback

Use this procedure for a bad package, tag, configuration release, or operational change. The current project is Alpha-only and has no production deployment guarantee.

## Decide

1. Declare the release version/commit affected using public identifiers only.
2. Confirm whether the issue is repository behavior, host/core behavior, tenant configuration, or suspected security exposure.
3. For a security exposure, stop normal rollback and follow [incident response](incident-response.md) and `SECURITY.md`.
4. Do not roll back by deleting evidence, changing a published tag, or editing history.

## Package/repository rollback

1. Stop enabling the affected pre-release in new environments.
2. Disable the plugin through the host's normal plugin controls; do not delete tenant data.
3. Pin the last reviewed package/tag or commit in the consuming environment, if one exists and is approved.
4. If a release artifact is withdrawn, document the public version and reason without including payloads, PII, credentials, tenant IDs, or raw logs.
5. Open a corrective change from a clean branch. Keep the published tag immutable and use a new version/tag for any replacement.
6. Re-run the complete local gates and release smoke checks before proposing a replacement.

## Tenant configuration rollback

Tenant owners must revoke or remove consent, app assignments, or test configuration through their approved Entra/change-management process. Maintainers must not request or handle tenant credentials. Confirm only sanitized state (for example, `consent revoked: yes`) through the authorized owner.

## Verify

- Plugin is disabled or pinned as intended.
- No new write operation is executable.
- CORE-1/CORE-2 and tenant-verification blockers remain accurately stated.
- The working tree and release evidence identify the exact immutable commit.
- No diagnostic output contains payloads or PII.

Rollback restores a known repository state; it does not prove that Graph or a tenant behaved correctly.
