# History before this repository

The workbook and exporter were developed through many iterations in September 2026, before version control. Most intermediate files were not kept. These notes were reconstructed afterwards from development notes, so changes marked *unknown* are not recorded. The only complete copies are the ones now in this repository: workbook v6.10.11 and exporter v1.2.4.

## Workbook

### Original compliance log (superseded)
- **Prototype / FULL**: a broad DMRB requirements log across GG 101, CD 109, CD 116, CD 123 and CD 195 (995 requirement rows) with coincidence and dashboard features. `#NAME?` errors were reported.
- **FULL_FIXED, C3D_Import, FIXED, FIXED3**: fixes for the `#NAME?` errors and import revisions; individual changes unknown.
- **LandXML importer v1–v3**: automated LandXML import, confirmed working.
- **LandXML importer v4**: vertical-curve grade checks using the maximum incoming/outgoing grade. This was later changed to avoid duplicate checks on curves.

### Geometry reviewer
| Version | Notes |
|---|---|
| v1 | Narrowed scope to Civil 3D geometry import and review. |
| v2 | Changes unknown. |
| v3 | Consolidated Review Setup, live formula review, chosen-profile filtering, junction zones, raw import sheets and coincidence review. |
| v4, v4.1, v5, v6 | Changes unknown. |
| v6.1 | LandXML refresh with settings preserved; raw table refresh; geometry/SSD findings; departure and coincidence outputs; helper visibility, table filters and refresh buttons. |
| v6.2 | Table protection and resizing repairs. |
| v6.3 | Formula and import repairs. |
| v6.4 | Coincidence overlaps must have positive length; same vertical element's K and grade no longer paired. |
| v6.4.1 | Tangent-only gradient checks with curve-grade fallback; shared point finding for adjoining curves; point findings excluded from coincidence. |
| v6.4.2, v6.4.3 | Repair releases; changes unknown. |
| v6.4.4 | Stable baseline: corrected output tables, coincidence/departure rebuild, tangent/fallback grades and clause 2.13 treatment. |
| v6.4.5, v6.5.x | Changes unknown (v6.5 split the VBA package). |
| v6.6.0 | Rebuilt from v6.4.4; restored populated reviews and tangent/curve grade logic. |
| v6.6.1 | Superelevation assessed only on circular curves with V²/R > 7; capped at 7% rural / 5% urban; missing full superelevation marked CHECK; Proposed/Existing basis added. |
| v6.6.2 – v6.9.0 | Changes unknown. |
| v6.10.0 | Integration of the superelevation CSV workflow. |
| v6.10.1 | Legacy macro clean-up; single `RepairGeometryReviewer` entry point (fixes duplicate-macro warnings). |
| v6.10.2 | Superelevation bounded by the review profile's extent; end-of-alignment fixes. |
| v6.10.3, v6.10.4 | Changes unknown. |
| v6.10.5 | Blank workbook release keeping VBA, buttons, formatting, validation, formulas and reference values. |
| v6.10.6 | **Clear / Start New Review** button. |
| v6.10.7 – v6.10.9 | Fixes for Error 91 on the clear button. |
| v6.10.10 | Import fix. |
| v6.10.11 | LandXML fix. **Starting point of this repository.** |

## Superelevation exporter
| Version | Notes |
|---|---|
| v1.0 | .NET plugin; alignment selection; uses `SuperelevationCriticalStations` / `GetSlope`; timestamped output. |
| v1.1 | `DMRBEXPORTSUPER` command for `NETLOAD`; combined CSV with alignment names and export timestamp. Needed the .NET 8 SDK. |
| v1.2 – v1.2.2 | Builds without the SDK, using the Windows `csc.exe` and the .NET 8 runtime shipped with Civil 3D; `/noconfig` fix. |
| v1.2.3 | Adds the `AecBaseMgd.dll` reference; fixes the `Label` name clash; prefers WindowsDesktop assemblies over facades. |
| v1.2.4 | Hides alignments without superelevation data; shows counts; sorts CSV by alignment, then numeric station. **Starting point of this repository.** |

## Known engineering issues to keep in mind
- LandXML profile PVI station/elevation and curve lengths can't always give a reliable K for non-tangential connections. Geometry shouldn't be inferred where the source is inadequate.
- LandXML superelevation records don't always contain full lane crossfall or critical-station data, which is why the CSV exporter exists.
- Relaxations depend on context and on design-speed band (70A/70B). Achieved values and limits should stay visible in the coincidence review. Clauses used include CD 109 2.11, 2.13, 3.5, 4.5, 5.1, 5.7 and 5.9; check them against the current edition.
- Avoid overlap false positives, duplicate vertical grade checks, and zero-length findings being treated as intervals.
- Areas that have broken before and are worth retesting after changes: table resizing and protection, formula compatibility, duplicate macros, Error 91 on clear, blank/reset behaviour, profile limits and LandXML refresh.
