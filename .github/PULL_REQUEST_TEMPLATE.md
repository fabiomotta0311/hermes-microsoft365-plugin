## Summary

<!-- Describe the change and link an issue if applicable. -->

## Release and support boundary

- [ ] This change does not claim tenant verification or production support.
- [ ] CORE-1/CORE-2 and tenant-verification blockers remain accurately documented.
- [ ] No payload, PII, credentials, or secret-bearing fixtures were added.
- [ ] No telemetry or logging that could expose payloads/PII was added.

## Validation

- [ ] `python -m pytest -q`
- [ ] `python -m compileall -q microsoft365`
- [ ] `python -m hermes_cli.plugin_validate microsoft365`
- [ ] `hermes plugins doctor microsoft365 --ci`
- [ ] `git diff --check`
- [ ] Secret scan completed with zero findings.

## Notes

<!-- Include known limitations, migration/release impact, or explain skipped checks. -->
