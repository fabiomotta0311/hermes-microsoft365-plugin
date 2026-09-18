# Changelog

All notable changes to this project are documented here.

## [Unreleased]

- Added release engineering for reproducible wheel builds, isolated install smoke tests,
  SHA-256 checksums, and a lockfile-derived CycloneDX SBOM artifact.
- Documented the release procedure and the boundary between build evidence and catalog
  or tenant verification.

## [0.1.0a1] - 2026-09-17

- Established the offline-verified standalone Microsoft 365 Graph plugin foundation.
- Added strict capability validation, operation registry reporting, bounded file transfer,
  paging, and typed SDK request contracts.
- No Microsoft 365 tenant was contacted; no catalog entry, tag, signature, or production
  support claim is made by this pre-release.

[Unreleased]: https://github.com/fabiomotta0311/hermes-microsoft365-plugin/compare/main...HEAD
[0.1.0a1]: https://github.com/fabiomotta0311/hermes-microsoft365-plugin/releases/tag/v0.1.0a1
