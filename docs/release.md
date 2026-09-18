# Release procedure

This repository prepares reproducible release evidence. The current `0.1.0a1`
pre-release is eligible for a maintainer-reviewed tag and SHA-pinned catalog proposal;
that proposal remains non-official until upstream admission. This workflow does not
sign artifacts or generate a GitHub artifact attestation.

## Before a release candidate

1. Confirm the capability registry and manifest describe only the verified surface.
2. Run the full local gates from `CONTRIBUTING.md`.
3. Review `CHANGELOG.md` and set the intended version in both `pyproject.toml` and
   `microsoft365/plugin.yaml`.
4. Regenerate `uv.lock` only when dependency changes are intentional and review the
   resulting lockfile diff.
5. Confirm the working tree is clean except for the reviewed release commit.

## Build and verify locally

From the repository root:

```bash
rm -rf /tmp/hermes-m365-build-a /tmp/hermes-m365-build-b dist
python -m pip wheel . --no-deps --no-cache-dir --wheel-dir /tmp/hermes-m365-build-a
python -m pip wheel . --no-deps --no-cache-dir --wheel-dir /tmp/hermes-m365-build-b
sha256sum /tmp/hermes-m365-build-a/*.whl /tmp/hermes-m365-build-b/*.whl
python scripts/release_smoke.py /tmp/hermes-m365-build-a/*.whl
python scripts/generate_sbom.py --output /tmp/hermes-m365-sbom.json
```

The two wheel hashes must match. The smoke test installs the wheel into a fresh
virtual environment, then checks import, distribution
metadata, and manifest/package version agreement. The SBOM is deterministic for a
given `uv.lock` and contains the locked Python components and their recorded SHA-256
artifact hashes.

## CI evidence workflow

Run **Actions → Release artifacts → Run workflow** on the reviewed commit. The
workflow builds the wheel twice, compares hashes, runs the isolated install smoke
test, emits the lockfile-derived CycloneDX JSON, writes `SHA256SUMS`, and uploads
these files as a short-lived workflow artifact. Download and retain that evidence
with the release review if policy requires it.

The workflow's artifact upload is not a publication step. No credentials, tenant
configuration, Microsoft Graph calls, signing keys, or provenance attestations are
used. A future signing/attestation change must add an explicit trust policy and
verification instructions rather than implying provenance from this workflow.

## Catalog hand-off

Only after capability truth, host compatibility, tenant verification requirements,
and catalog review are complete may a maintainer create a version tag or submit a
SHA-pinned catalog entry. Those actions are intentionally outside this change.
