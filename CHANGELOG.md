# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

This project is licensed under the [Apache License 2.0](LICENSE).

## [Unreleased]

### Added

- Pluggable SQL connector registry (`sliger.connectors`) so warehouses share one `sql()` contract
- BigQuery extra (`sliger[bigquery]`) with `:named` parameters rewritten to `@name`
- Snowflake extra (`sliger[snowflake]`) with password or programmatic access token
- Databricks SQL extra (`sliger[databricks]`) using native `:named` parameters

## [0.2.0] - 2026-09-06

### Added

- Sliger Python client and `python -m sliger`
- uv/hatchling packaging, Ruff, pytest, GitHub Actions CI
- `--dry-run` for jinjify/imagify (preview without writing)
- User OAuth installed-app login as alternative to service accounts
- Temporary Drive uploads for imagify (file deleted after insert)
- PyPI release workflow on version tags

### Changed

- Duplicate presentation is not world-writable by default (`--anyone-can-edit` / `--anyone-can-view`)
- Slide text replace is per-shape (deleteText + insertText)
- Function map entries must be `module.function` dotted paths

### Removed

- Poetry; the unused `google` PyPI package; global CLI state in `__init__.py`

[0.2.0]: https://github.com/slidoapp/sliger/releases/tag/v0.2.0
