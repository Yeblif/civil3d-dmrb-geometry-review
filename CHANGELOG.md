# Changelog

All notable changes to this project are recorded here, following [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

The workbook and the exporter have separate version numbers. Tags are `workbook-vX.Y.Z` and `exporter-vX.Y.Z`. History from before this repository is in [docs/history.md](docs/history.md).

## [Unreleased]

### Added
- Workbook v6.10.11 as the starting point, released blank with all project data, document author details and saved file locations removed.
- VBA module exported to `src/vba/` for readable diffs.
- Superelevation exporter v1.2.4 source and no-SDK build script in `exporter/`.
- Synthetic test LandXML and superelevation CSV with a generator script.
- Documentation, issue templates and pull request template.

### Changed
- Exporter build script and README now direct build errors to GitHub issues.

### Fixed
