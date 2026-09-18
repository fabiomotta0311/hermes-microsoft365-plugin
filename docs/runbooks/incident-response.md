# Incident response

This is the Alpha-only response procedure for repository, release, host-integration, or suspected data/security incidents. It is not a promise of production on-call coverage.

## 1. Triage and contain

1. Assign an incident owner and record the UTC start time.
2. Classify as availability, correctness, release, security, or suspected data exposure.
3. Capture only sanitized metadata: version/commit, operation name, auth mode, status code, and affected environment alias. Do not collect payloads, PII, tenant/user IDs, tokens, headers, or raw traces.
4. Disable the affected plugin/release or stop the test window. Do not retry writes or ask users to resend sensitive content.
5. Preserve immutable repository/release references; do not rewrite tags or delete logs that contain no sensitive data.

## 2. Security or exposure path

If a credential, token, tenant identifier, or payload may have been exposed:

- Stop handling it in issues, chat, or commits.
- Ask the tenant owner to revoke/rotate the affected credential through the approved channel; maintainers never need the secret.
- Use the private GitHub security advisory flow described in `SECURITY.md`.
- Remove accidental secret material from the working tree and follow the repository's credential-remediation process, without pasting the value into any report.
- Preserve a sanitized timeline and affected version only.

## 3. Investigate

Reproduce with synthetic data or offline tests. Inspect code and configuration, not customer content. Separate facts from hypotheses and mark tenant behavior as unverified unless approved tenant evidence exists. CORE-1, CORE-2, and tenant verification remain blockers and cannot be waived during an incident.

## 4. Communicate

Use the issue or private advisory appropriate to the classification. State impact, affected version, containment, and next update without identifying tenants/users or quoting payloads. Do not promise a fix date or production support.

## 5. Recover and close

1. Apply the smallest reviewed fix or rollback.
2. Run the documented local gates and secret scan.
3. Verify the disabled/pinned state and the exact immutable commit.
4. Close with sanitized root cause, timeline, validation, and follow-up owners.
5. Add a regression test for code behavior when code is changed; documentation/config-only fixes may use direct validation.

## Required closure checklist

- [ ] Containment verified.
- [ ] No payload/PII telemetry or diagnostic collection added.
- [ ] Credentials rotated/revoked when applicable.
- [ ] Public/private communication used appropriately.
- [ ] Offline and release gates pass.
- [ ] Alpha-only boundary and remaining blockers are still accurate.
