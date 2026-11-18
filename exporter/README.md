# DMRB Superelevation Exporter (v1.2.4)

A Civil 3D 2026 command, `DMRBEXPORTSUPER`, that writes the superelevation critical stations of selected alignments to one CSV for the geometry review workbook. It exists because LandXML exports don't reliably include full lane crossfall data.

## Build (one-off, no SDK or admin rights needed)

The build uses only what's already on a Civil 3D 2026 PC:

- the Windows C# compiler (`csc.exe`, part of .NET Framework);
- the .NET 8 runtime installed with Civil 3D;
- the AutoCAD / Civil 3D managed API DLLs.

Nothing is installed and no system settings are changed. Autodesk DLLs are referenced from your Civil 3D installation and are not included here.

1. Download or clone this folder somewhere writable.
2. Double-click `BuildForNETLOAD_NoSDK.cmd`.
3. If it succeeds, it creates `NETLOAD\DMRBSuperelevationExporter.dll`.

If the build fails, please [open an issue](https://github.com/Yeblif/civil3d-dmrb-geometry-review/issues/new/choose) with the compiler output.

`API_TARGET.txt` lists the Civil 3D API versions the build was developed against.

## Run in Civil 3D

1. Run `NETLOAD` and select `NETLOAD\DMRBSuperelevationExporter.dll`.
2. Run `DMRBEXPORTSUPER`.
3. Tick the alignments to export. Only alignments that have superelevation data are listed; the window shows how many were hidden.
4. Choose the output folder and click **Export**.

## Output

One CSV, named `<drawing>_Superelevation_<yyyy-MM-dd_HHmmss>.csv`. Rows are sorted by alignment name, then station. Columns:

- `Exported At`, `Drawing`, `Alignment`, `Superelevation Region`, `Station`, `Description`, `Smoothing Curve Length`;
- the eight lane and shoulder crossfalls, in the same order as Civil 3D's Superelevation Tabular Editor.

Import it in the workbook with **Import Superelevation CSV**. There's an example in [`../test-data/`](../test-data/).

## Build notes

- `csc.exe` is run directly with `/noconfig`, so .NET Framework default references aren't mixed with .NET 8 references.
- `AecBaseMgd.dll` is referenced because Civil 3D alignment objects inherit from the AEC API.
- `System.Windows.Forms.Label` is fully qualified to avoid a clash with `Autodesk.Civil.DatabaseServices.Label`.
- WindowsDesktop assemblies take priority over same-named .NET Core facade assemblies such as `WindowsBase.dll`.
