# Security Policy

## Supported versions

No version is production-supported yet. The project is pre-release.

## Reporting

Do not open a public issue containing credentials, tenant identifiers, tokens, message/file content, or exploit details. Contact the repository maintainer privately through the GitHub security advisory flow.

## Security boundary

The plugin runs in-process and is not a sandbox. Microsoft Graph permissions and tenant policy remain authoritative. Secrets must be supplied through Hermes secret handling. Write actions require the Hermes host approval gate, and this plugin intentionally does not work around host approval-ordering defects.
