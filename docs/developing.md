# Developing

`src/vba/C3D_Geometry_Reviewer_Importer.bas` is the source of truth for the code. Pull requests change the `.bas` only. The workbook (`Civil3D_Geometry_Review.xlsm`) is updated once per release, so between releases the module inside it can be behind `main`.

## Changing the code

1. Raise an issue for the bug, feature or task.
2. The change is made on a `claude/…` branch and opened as a **draft** pull request that links the issue (`Fixes #n`). Never commit directly to `main`.
3. **Test the change before it's merged.** Open a local copy of the workbook, then in the VBA editor (Alt+F11) remove the `C3D_Geometry_Reviewer_Importer` module and use **File → Import File…** to import the `.bas` from the pull request branch. Run the affected buttons against the files in `test-data/`.
4. If it fails, report the error on the pull request and the fix goes on the same branch. If it passes, tick *Tested in Excel*, mark the pull request ready for review and merge it.

Don't commit the workbook as part of a code change.

## Keep project data out of the repository

This repository is public. Before committing the workbook:

1. Click **Clear / Start New Review** on Review Setup.
2. Clear any table filters (Data → Clear).
3. **File → Info → Check for Issues → Inspect Document** → *Document Properties and Personal Information* → **Remove All**. This removes your name and the saved file location.
4. Save.

Never commit real LandXML or CSV exports, client or project names, or Autodesk DLLs. Keep real scheme files outside the repo folder, or in a `_local/` folder, which is ignored.

## Line endings

`.bas`, `.cls` and `.cmd` files are stored exactly as written, with Windows (CRLF) line endings, because the VBA editor's Import needs CRLF. `.gitattributes` handles this; don't change it.

## Test data

`test-data/generate_test_data.py` regenerates the synthetic files. It needs Python 3 and nothing else.

## Releases

When the merged changes make a version worth releasing:

1. Open the workbook, remove the `C3D_Geometry_Reviewer_Importer` module and import the `.bas` from `main`.
2. Run every button against `test-data/`.
3. Make the workbook blank (see *Keep project data out of the repository*).
4. Commit the `.xlsm` on a `claude/…` branch, move the *Unreleased* entries in `CHANGELOG.md` to a new version heading, and merge.
5. Tag the merge commit:

- workbook: `workbook-v6.10.12`, `workbook-v6.11.0`, …
- exporter: `exporter-v1.2.5`, …
