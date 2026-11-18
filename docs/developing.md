# Developing

The workbook is the working copy of the code. `src/vba/C3D_Geometry_Reviewer_Importer.bas` is an exported text copy, kept so every change can be read and diffed on GitHub. Commit both together.

## Changing the code

1. Open `Civil3D_Geometry_Review.xlsm` and make your changes in the VBA editor (Alt+F11).
2. Test with the files in `test-data/`.
3. In the VBA editor, right-click `C3D_Geometry_Reviewer_Importer` → **Export File…**, and save over `src/vba/C3D_Geometry_Reviewer_Importer.bas`.
4. **Make the workbook blank before committing** (see below).
5. Commit the `.xlsm` and the `.bas` together, add a line to `CHANGELOG.md` under *Unreleased*, and open a pull request.

To go the other way (for example after reviewing a change on GitHub), use **File → Import File…** in the VBA editor. Remove the existing module first.

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

When a version is finished, move the *Unreleased* entries in `CHANGELOG.md` to a new version heading and tag it:

- workbook: `workbook-v6.10.12`, `workbook-v6.11.0`, …
- exporter: `exporter-v1.2.5`, …
