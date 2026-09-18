# Support

## Support boundary

`hermes-microsoft365-plugin` is **Alpha-only** (`v0.1.0a2`). It is an offline-verified, opt-in pre-release and is not production-supported. No tenant authentication, consent flow, endpoint permissions, throttling, pagination, or mutation behavior has been verified against a real Microsoft 365 tenant.

Support is limited to reproducible repository behavior, documented configuration validation, and sanitized local test failures. We do not promise tenant-specific troubleshooting, uptime, data recovery, or Microsoft Graph support.

The following remain release blockers:

- **CORE-1**: the generic Hermes approval/final-argument binding fix required before write execution can be enabled.
- **CORE-2**: the complementary host/core dependency required by the release plan before write execution can be enabled.
- **Tenant verification**: authorized-tenant evidence for authentication, consent, permissions, endpoint behavior, throttling, pagination, and mutations.

Do not interpret a passing offline test, package build, or plugin doctor run as tenant verification.

## Before opening an issue

1. Reproduce on the published tag or identify the exact commit and Python version.
2. Run the relevant checks in [CONTRIBUTING.md](CONTRIBUTING.md).
3. Search existing issues and the [known limitations](docs/known-limitations.md).
4. Remove secrets and all tenant, message, file, calendar, task, and chat content from the report.

## What to include

Include only sanitized metadata: command name, local OS/Python version, plugin version, authentication mode (`application` or `delegated`), operation name, sanitized HTTP status/error code, and a short reproduction using placeholders. Include a traceback only after removing arguments, URLs containing identifiers, headers, tokens, request/response bodies, and file paths that reveal private data.

Never attach `.env`, tokens, client secrets, tenant IDs, user IDs, message/file content, raw Graph requests or responses, or telemetry dumps. This project does not add payload or PII telemetry. Do not enable ad-hoc logging that captures them.

## Channels

- General bugs and documentation: GitHub issue using the supplied templates.
- Security vulnerabilities: follow [SECURITY.md](SECURITY.md) and use GitHub's private security advisory flow; do not use a public issue.
- Operational handling: maintainers should use the runbooks in [`docs/runbooks/`](docs/runbooks/).

Issues may be closed as unsupported when they require tenant verification, production support, CORE-1/CORE-2, or an unimplemented operation.