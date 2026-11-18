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
- Sheets are no longer protected. The buttons used to re-protect the raw import sheets after every run, which caused problems.
- Development workflow: the `.bas` is now the source of truth, code pull requests are opened as drafts and tested in Excel before merging, and the workbook is only updated at release. See `docs/developing.md`.
- Exporter build script and README now direct build errors to GitHub issues.

### Fixed
- Setup Status and Live Review Summary on Review Setup no longer have leftover dropdowns. A value picked there was overwritten by Update Reviews; Road Type and Include? are set in columns C and D.
- Running setup or repair repeatedly no longer adds duplicate **Clear / Start New Review** buttons.
- Update Reviews and other buttons failed with Error 91 ("Object variable or With block variable not set"), first seen while building clause 2.13 junction zones. `ResizeTableRows` now makes sure an emptied table has a data body after resizing, so the Control Zones, Review, superelevation, Review Setup and design-speed tables can be written again.
