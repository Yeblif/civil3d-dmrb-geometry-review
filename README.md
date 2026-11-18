# Civil 3D DMRB Geometry Review

An Excel/VBA workbook that imports highway geometry from Civil 3D and reviews it against DMRB CD 109, flagging potential departures, relaxations and coincident relaxations for an engineer to check.

## What it does

- **Imports Civil 3D data** from a LandXML export: alignments, horizontal elements (lines, clothoids, arcs), design and surface profiles, superelevation and design-speed stations.
- **Imports superelevation** from a CSV written by the companion Civil 3D command `DMRBEXPORTSUPER` (in [`exporter/`](exporter/)), because LandXML superelevation data is often incomplete.
- **Reviews geometry against CD 109**:
  - horizontal radius and superelevation;
  - crest and sag K values and gradients, using tangent grades (with a curve-grade fallback where there is no tangent);
  - stopping sight distance, from values you enter manually.
- **Applies design-step relaxations** by design speed and band (A/B), road type and context rules, capped at the clause 2.11 limits. Junction zones (clause 2.13) are generated from the junctions you enter.
- **Produces three outputs**: a full Geometry Review, a Departure Summary, and a Coincidence Review that pairs overlapping relaxations along each alignment.

> **This is a review aid, not a certified compliance checker.** Results depend on the quality of the Civil 3D export and on the reference values in the workbook. Check findings against the current edition of DMRB, and use your own engineering judgement, before relying on them for design.

## How to use

**Requirements:** desktop Excel on Windows with macros enabled. The workbook uses MSXML 6.0 and the Windows Scripting Runtime, so it won't run in Excel for Mac or Excel on the web. The superelevation exporter needs Civil 3D 2026.

1. **Download** [`Civil3D_Geometry_Review.xlsm`](Civil3D_Geometry_Review.xlsm). If Windows blocks the macros, right-click the file → Properties → tick **Unblock**.
2. **Export from Civil 3D:**
   - LandXML: *Output → Export to LandXML*, including the alignments and profiles you want to review.
   - Superelevation CSV: run `DMRBEXPORTSUPER`. See [`exporter/README.md`](exporter/README.md) for the one-off build and `NETLOAD`.
3. **Review Setup sheet:**
   1. Click **Refresh Civil 3D LandXML** and pick the XML file.
   2. Click **Import Superelevation CSV** and pick the CSV.
   3. For each alignment, choose the review profile, road type and whether to include it.
4. Complete the input sheets that apply (yellow cells are inputs):
   - **Design Speeds**: band A/B for each speed zone;
   - **Superelevation Setup**: rural/urban environment and assessment basis;
   - **SSD Review**: manual SSD records, junctions and relaxation contexts.
5. Click **Update Reviews**, then work through **Geometry Review**, **Departure Summary** and **Coincidence Review**.
6. To start a new scheme, click **Clear / Start New Review**. It clears all project data but keeps the rules, formulas and formatting.

**To try it without a real scheme**, use the synthetic files in [`test-data/`](test-data/).

## Files

| Path | What it is |
|---|---|
| `Civil3D_Geometry_Review.xlsm` | The workbook, blank and ready to use. Code, buttons, validation and CD 109 reference values included. |
| `src/vba/C3D_Geometry_Reviewer_Importer.bas` | The workbook's VBA module as text, so code changes can be read and diffed on GitHub. Must match the copy inside the workbook. |
| `exporter/` | Source and no-SDK build script for the `DMRBEXPORTSUPER` Civil 3D superelevation exporter (.NET 8). |
| `test-data/` | A synthetic LandXML and superelevation CSV, plus the script that generates them. |
| `docs/history.md` | Version history and known engineering issues from development before this repo. |
| `docs/developing.md` | How to change the workbook and commit it, including keeping project data out of the repo. |

## Changelog

See [CHANGELOG.md](CHANGELOG.md).
