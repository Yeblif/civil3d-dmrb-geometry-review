using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.Versioning;
using System.Text;
using System.Windows.Forms;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;

using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: TargetFramework(".NETCoreApp,Version=v8.0", FrameworkDisplayName = ".NET 8.0")]

namespace DMRB.SuperelevationExporter
{
    public sealed class Commands
    {
        private static readonly SuperelevationCrossSegmentType[] NativeColumnOrder =
        {
            SuperelevationCrossSegmentType.LeftOutShoulderCrossSlope,
            SuperelevationCrossSegmentType.LeftOutLaneCrossSlope,
            SuperelevationCrossSegmentType.LeftInLaneCrossSlope,
            SuperelevationCrossSegmentType.LeftInShoulderCrossSlope,
            SuperelevationCrossSegmentType.RightOutShoulderCrossSlope,
            SuperelevationCrossSegmentType.RightOutLaneCrossSlope,
            SuperelevationCrossSegmentType.RightInLaneCrossSlope,
            SuperelevationCrossSegmentType.RightInShoulderCrossSlope
        };

        [CommandMethod("DMRBEXPORTSUPER", CommandFlags.Modal)]
        public void ExportSuperelevation()
        {
            Document doc = AcApp.DocumentManager.MdiActiveDocument;
            if (doc == null)
                return;

            Editor ed = doc.Editor;
            Database db = doc.Database;

            try
            {
                int totalAlignmentCount;
                int withSuperelevationCount;
                int withoutSuperelevationCount;

                List<AlignmentChoice> alignments = ReadAlignmentChoices(
                    db,
                    out totalAlignmentCount,
                    out withSuperelevationCount,
                    out withoutSuperelevationCount);

                if (totalAlignmentCount == 0)
                {
                    ed.WriteMessage("\nNo Civil 3D alignments were found in the active drawing.");
                    return;
                }

                if (withSuperelevationCount == 0)
                {
                    string message = string.Format(
                        CultureInfo.InvariantCulture,
                        "{0} alignment{1} found. None contain superelevation data.",
                        totalAlignmentCount,
                        totalAlignmentCount == 1 ? "" : "s");

                    ed.WriteMessage("\n" + message);
                    MessageBox.Show(
                        message,
                        "DMRB Superelevation Exporter",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }

                string drawingBaseName = GetDrawingBaseName(db);
                string defaultFolder = GetDefaultOutputFolder(db);

                using (ExportForm form = new ExportForm(
                    alignments,
                    drawingBaseName,
                    defaultFolder,
                    withSuperelevationCount,
                    withoutSuperelevationCount))
                {
                    DialogResult result = AcApp.ShowModalDialog(form);
                    if (result != DialogResult.OK)
                        return;

                    IList<AlignmentChoice> selected = form.SelectedAlignments;
                    if (selected.Count == 0)
                    {
                        MessageBox.Show(
                            "Select at least one alignment.",
                            "DMRB Superelevation Exporter",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                        return;
                    }

                    DateTimeOffset exportedAt = DateTimeOffset.Now;
                    Directory.CreateDirectory(form.OutputFolder);

                    string filename = string.Format(
                        CultureInfo.InvariantCulture,
                        "{0}_Superelevation_{1}.csv",
                        SanitizeFileName(drawingBaseName),
                        exportedAt.ToString("yyyy-MM-dd_HHmmss", CultureInfo.InvariantCulture));

                    string outputPath = Path.Combine(form.OutputFolder, filename);

                    int rowCount = ExportCsv(db, selected, outputPath, drawingBaseName, exportedAt);

                    ed.WriteMessage(
                        string.Format(
                            CultureInfo.InvariantCulture,
                            "\nDMRB superelevation export complete: {0} rows -> {1}",
                            rowCount,
                            outputPath));

                    MessageBox.Show(
                        string.Format(
                            CultureInfo.InvariantCulture,
                            "Export complete.\n\nAlignments: {0}\nRows: {1}\n\n{2}",
                            selected.Count,
                            rowCount,
                            outputPath),
                        "DMRB Superelevation Exporter",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\nDMRBEXPORTSUPER failed: " + ex.Message);
                MessageBox.Show(
                    ex.ToString(),
                    "DMRB Superelevation Exporter - Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private static List<AlignmentChoice> ReadAlignmentChoices(
            Database db,
            out int totalAlignmentCount,
            out int withSuperelevationCount,
            out int withoutSuperelevationCount)
        {
            List<AlignmentChoice> output = new List<AlignmentChoice>();
            CivilDocument civilDoc = CivilApplication.ActiveDocument;

            totalAlignmentCount = 0;
            withSuperelevationCount = 0;
            withoutSuperelevationCount = 0;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                foreach (ObjectId id in civilDoc.GetAlignmentIds())
                {
                    Alignment alignment = tr.GetObject(id, OpenMode.ForRead) as Alignment;
                    if (alignment == null)
                        continue;

                    totalAlignmentCount++;

                    int criticalStationCount = 0;
                    int regionCount = 0;

                    try
                    {
                        criticalStationCount = alignment.SuperelevationCriticalStations.Count;
                    }
                    catch
                    {
                    }

                    try
                    {
                        regionCount = alignment.SuperelevationCurves.Count;
                    }
                    catch
                    {
                    }

                    bool hasSuperelevationData = HasSuperelevationData(alignment, regionCount);

                    if (hasSuperelevationData)
                    {
                        withSuperelevationCount++;
                        output.Add(new AlignmentChoice(
                            id,
                            alignment.Name,
                            criticalStationCount,
                            regionCount));
                    }
                    else
                    {
                        withoutSuperelevationCount++;
                    }
                }

                tr.Commit();
            }

            output.Sort(
                delegate(AlignmentChoice a, AlignmentChoice b)
                {
                    return StringComparer.OrdinalIgnoreCase.Compare(a.Name, b.Name);
                });

            return output;
        }

        private static bool HasSuperelevationData(Alignment alignment, int regionCount)
        {
            // SuperelevationCurves is the normal indicator that Civil 3D has
            // superelevation calculated/set for the alignment.
            if (regionCount > 0)
                return true;

            // Defensive fallback: keep an alignment if it contains a real
            // critical station other than Civil 3D's automatic Begin/End
            // Alignment rows. This avoids hiding manually edited data.
            try
            {
                SuperelevationCriticalStationCollection stations =
                    alignment.SuperelevationCriticalStations;

                for (int i = 0; i < stations.Count; i++)
                {
                    string stationType = stations[i].StationType.ToString();

                    if (!string.Equals(
                            stationType,
                            "BeginAlignment",
                            StringComparison.OrdinalIgnoreCase) &&
                        !string.Equals(
                            stationType,
                            "EndAlignment",
                            StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
            }
            catch
            {
            }

            return false;
        }

        private static int ExportCsv(
            Database db,
            IList<AlignmentChoice> selected,
            string outputPath,
            string drawingBaseName,
            DateTimeOffset exportedAt)
        {
            List<ExportRow> rows = new List<ExportRow>();
            List<AlignmentChoice> ordered = new List<AlignmentChoice>(selected);

            ordered.Sort(
                delegate(AlignmentChoice a, AlignmentChoice b)
                {
                    return StringComparer.OrdinalIgnoreCase.Compare(a.Name, b.Name);
                });

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                foreach (AlignmentChoice choice in ordered)
                {
                    Alignment alignment = tr.GetObject(choice.Id, OpenMode.ForRead) as Alignment;
                    if (alignment == null)
                        continue;

                    SuperelevationCriticalStationCollection stations =
                        alignment.SuperelevationCriticalStations;

                    for (int i = 0; i < stations.Count; i++)
                    {
                        SuperelevationCriticalStation station = stations[i];

                        string region = SafeString(
                            delegate { return station.SuperelevationCurveName; });

                        string description = SafeString(
                            delegate { return station.TransitionDescription; });

                        if (string.IsNullOrWhiteSpace(description))
                            description = PrettyEnumName(station.StationType.ToString());

                        string smoothingLength = GetRepresentativeSmoothingLength(station);

                        List<string> row = new List<string>();
                        row.Add(exportedAt.ToString("yyyy-MM-dd HH:mm:ss zzz", CultureInfo.InvariantCulture));
                        row.Add(drawingBaseName);
                        row.Add(alignment.Name);
                        row.Add(region);
                        row.Add(FormatStation(station.Station));
                        row.Add(description);
                        row.Add(smoothingLength);

                        foreach (SuperelevationCrossSegmentType segmentType in NativeColumnOrder)
                            row.Add(FormatSlope(TryGetSlope(station, segmentType)));

                        rows.Add(new ExportRow(
                            alignment.Name,
                            station.Station,
                            i,
                            row.ToArray()));
                    }
                }

                tr.Commit();
            }

            rows.Sort(
                delegate(ExportRow a, ExportRow b)
                {
                    int nameCompare = StringComparer.OrdinalIgnoreCase.Compare(
                        a.AlignmentName,
                        b.AlignmentName);

                    if (nameCompare != 0)
                        return nameCompare;

                    int stationCompare = a.Station.CompareTo(b.Station);

                    if (stationCompare != 0)
                        return stationCompare;

                    return a.OriginalOrder.CompareTo(b.OriginalOrder);
                });

            string[] header =
            {
                "Exported At",
                "Drawing",
                "Alignment",
                "Superelevation Region",
                "Station",
                "Description",
                "Smoothing Curve Length",
                "Left Outside Shoulder",
                "Left Outside Lane",
                "Left Inside Lane",
                "Left Inside Shoulder",
                "Right Outside Shoulder",
                "Right Outside Lane",
                "Right Inside Lane",
                "Right Inside Shoulder"
            };

            using (StreamWriter writer = new StreamWriter(
                outputPath,
                false,
                new UTF8Encoding(true)))
            {
                writer.WriteLine(JoinCsv(header));

                foreach (ExportRow row in rows)
                    writer.WriteLine(JoinCsv(row.Fields));
            }

            return rows.Count;
        }

        private static double? TryGetSlope(
            SuperelevationCriticalStation station,
            SuperelevationCrossSegmentType segmentType)
        {
            try
            {
                return station.GetSlope(segmentType);
            }
            catch (InvalidOperationException)
            {
                return null;
            }
        }

        private static string GetRepresentativeSmoothingLength(
            SuperelevationCriticalStation station)
        {
            SuperelevationCrossSegmentType[] preference =
            {
                SuperelevationCrossSegmentType.LeftOutLaneCrossSlope,
                SuperelevationCrossSegmentType.RightOutLaneCrossSlope,
                SuperelevationCrossSegmentType.LeftInLaneCrossSlope,
                SuperelevationCrossSegmentType.RightInLaneCrossSlope,
                SuperelevationCrossSegmentType.LeftOutShoulderCrossSlope,
                SuperelevationCrossSegmentType.RightOutShoulderCrossSlope,
                SuperelevationCrossSegmentType.LeftInShoulderCrossSlope,
                SuperelevationCrossSegmentType.RightInShoulderCrossSlope
            };

            foreach (SuperelevationCrossSegmentType type in preference)
            {
                try
                {
                    double length = station.GetSmoothingLength(type);
                    return length.ToString("0.000", CultureInfo.InvariantCulture);
                }
                catch (InvalidOperationException)
                {
                }
            }

            return string.Empty;
        }

        private static string FormatStation(double value)
        {
            return value.ToString("0.000", CultureInfo.InvariantCulture) + "m";
        }

        private static string FormatSlope(double? slope)
        {
            if (!slope.HasValue)
                return string.Empty;

            return (slope.Value * 100.0).ToString("0.000", CultureInfo.InvariantCulture) + "%";
        }

        private static string JoinCsv(string[] fields)
        {
            StringBuilder builder = new StringBuilder();

            for (int i = 0; i < fields.Length; i++)
            {
                if (i > 0)
                    builder.Append(',');

                builder.Append(CsvEscape(fields[i]));
            }

            return builder.ToString();
        }

        private static string CsvEscape(string value)
        {
            if (value == null)
                value = string.Empty;

            if (value.IndexOf('"') >= 0)
                value = value.Replace("\"", "\"\"");

            if (value.IndexOf(',') >= 0 ||
                value.IndexOf('"') >= 0 ||
                value.IndexOf('\r') >= 0 ||
                value.IndexOf('\n') >= 0)
            {
                return "\"" + value + "\"";
            }

            return value;
        }

        private static string SafeString(Func<string> getter)
        {
            try
            {
                string value = getter();
                return value ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string PrettyEnumName(string value)
        {
            if (string.IsNullOrEmpty(value))
                return string.Empty;

            StringBuilder output = new StringBuilder();
            output.Append(value[0]);

            for (int i = 1; i < value.Length; i++)
            {
                char c = value[i];
                char previous = value[i - 1];

                if (char.IsUpper(c) && !char.IsUpper(previous))
                    output.Append(' ');

                output.Append(c);
            }

            return output.ToString();
        }

        private static string GetDrawingBaseName(Database db)
        {
            string filename = db.Filename;

            if (string.IsNullOrWhiteSpace(filename))
                return "Untitled";

            return Path.GetFileNameWithoutExtension(filename);
        }

        private static string GetDefaultOutputFolder(Database db)
        {
            string filename = db.Filename;

            if (!string.IsNullOrWhiteSpace(filename))
            {
                string folder = Path.GetDirectoryName(filename);

                if (!string.IsNullOrWhiteSpace(folder) && Directory.Exists(folder))
                    return folder;
            }

            return Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        }

        private static string SanitizeFileName(string value)
        {
            char[] invalid = Path.GetInvalidFileNameChars();

            foreach (char c in invalid)
                value = value.Replace(c, '_');

            return value;
        }
    }

    internal sealed class ExportRow
    {
        public string AlignmentName { get; private set; }
        public double Station { get; private set; }
        public int OriginalOrder { get; private set; }
        public string[] Fields { get; private set; }

        public ExportRow(
            string alignmentName,
            double station,
            int originalOrder,
            string[] fields)
        {
            AlignmentName = alignmentName;
            Station = station;
            OriginalOrder = originalOrder;
            Fields = fields;
        }
    }

    internal sealed class AlignmentChoice
    {
        public ObjectId Id { get; private set; }
        public string Name { get; private set; }
        public int CriticalStationCount { get; private set; }
        public int RegionCount { get; private set; }

        public AlignmentChoice(
            ObjectId id,
            string name,
            int criticalStationCount,
            int regionCount)
        {
            Id = id;
            Name = name;
            CriticalStationCount = criticalStationCount;
            RegionCount = regionCount;
        }

        public override string ToString()
        {
            return string.Format(
                CultureInfo.InvariantCulture,
                "{0}    [{1} critical stations, {2} superelevation regions]",
                Name,
                CriticalStationCount,
                RegionCount);
        }
    }

    internal sealed class ExportForm : Form
    {
        private readonly CheckedListBox _alignmentList;
        private readonly TextBox _folderBox;

        public IList<AlignmentChoice> SelectedAlignments
        {
            get
            {
                List<AlignmentChoice> selected = new List<AlignmentChoice>();

                foreach (object item in _alignmentList.CheckedItems)
                {
                    AlignmentChoice choice = item as AlignmentChoice;
                    if (choice != null)
                        selected.Add(choice);
                }

                return selected;
            }
        }

        public string OutputFolder
        {
            get { return _folderBox.Text.Trim(); }
        }

        public ExportForm(
            IList<AlignmentChoice> alignments,
            string drawingBaseName,
            string defaultFolder,
            int withSuperelevationCount,
            int withoutSuperelevationCount)
        {
            Text = "DMRB Superelevation Exporter";
            StartPosition = FormStartPosition.CenterScreen;
            Width = 820;
            Height = 600;
            MinimizeBox = false;
            MaximizeBox = false;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            ShowIcon = false;

            System.Windows.Forms.Label title = new System.Windows.Forms.Label();
            title.Left = 16;
            title.Top = 15;
            title.Width = 770;
            title.Height = 38;
            title.Text = string.Format(
                CultureInfo.InvariantCulture,
                "Only alignments containing superelevation data are shown. {0} alignment{1} ha{2} superelevation data; {3} alignment{4} do{5} not and {6} hidden.",
                withSuperelevationCount,
                withSuperelevationCount == 1 ? "" : "s",
                withSuperelevationCount == 1 ? "s" : "ve",
                withoutSuperelevationCount,
                withoutSuperelevationCount == 1 ? "" : "s",
                withoutSuperelevationCount == 1 ? "es" : "",
                withoutSuperelevationCount == 1 ? "is" : "are");
            title.AutoSize = false;
            Controls.Add(title);

            _alignmentList = new CheckedListBox();
            _alignmentList.Left = 16;
            _alignmentList.Top = 58;
            _alignmentList.Width = 770;
            _alignmentList.Height = 350;
            _alignmentList.CheckOnClick = true;
            _alignmentList.HorizontalScrollbar = true;

            foreach (AlignmentChoice item in alignments)
                _alignmentList.Items.Add(item, true);

            Controls.Add(_alignmentList);

            Button selectAllButton = new Button();
            selectAllButton.Left = 16;
            selectAllButton.Top = 418;
            selectAllButton.Width = 100;
            selectAllButton.Height = 30;
            selectAllButton.Text = "Select All";
            selectAllButton.Click += delegate { SetAllChecked(true); };
            Controls.Add(selectAllButton);

            Button clearButton = new Button();
            clearButton.Left = 124;
            clearButton.Top = 418;
            clearButton.Width = 100;
            clearButton.Height = 30;
            clearButton.Text = "Clear All";
            clearButton.Click += delegate { SetAllChecked(false); };
            Controls.Add(clearButton);

            System.Windows.Forms.Label folderLabel = new System.Windows.Forms.Label();
            folderLabel.Left = 16;
            folderLabel.Top = 465;
            folderLabel.Width = 110;
            folderLabel.Height = 22;
            folderLabel.Text = "Output folder:";
            Controls.Add(folderLabel);

            _folderBox = new TextBox();
            _folderBox.Left = 126;
            _folderBox.Top = 461;
            _folderBox.Width = 575;
            _folderBox.Height = 26;
            _folderBox.Text = defaultFolder;
            Controls.Add(_folderBox);

            Button browseButton = new Button();
            browseButton.Left = 709;
            browseButton.Top = 459;
            browseButton.Width = 77;
            browseButton.Height = 30;
            browseButton.Text = "Browse...";
            browseButton.Click += BrowseButton_Click;
            Controls.Add(browseButton);

            System.Windows.Forms.Label namingLabel = new System.Windows.Forms.Label();
            namingLabel.Left = 16;
            namingLabel.Top = 495;
            namingLabel.Width = 770;
            namingLabel.Height = 30;
            namingLabel.Text = string.Format(
                CultureInfo.InvariantCulture,
                "Filename is generated at export time, e.g. {0}_Superelevation_{1}.csv",
                drawingBaseName,
                DateTime.Now.ToString("yyyy-MM-dd_HHmmss", CultureInfo.InvariantCulture));
            Controls.Add(namingLabel);

            Button exportButton = new Button();
            exportButton.Left = 590;
            exportButton.Top = 525;
            exportButton.Width = 95;
            exportButton.Height = 32;
            exportButton.Text = "Export";
            exportButton.DialogResult = DialogResult.OK;
            exportButton.Click += ExportButton_Click;
            Controls.Add(exportButton);
            AcceptButton = exportButton;

            Button cancelButton = new Button();
            cancelButton.Left = 691;
            cancelButton.Top = 525;
            cancelButton.Width = 95;
            cancelButton.Height = 32;
            cancelButton.Text = "Cancel";
            cancelButton.DialogResult = DialogResult.Cancel;
            Controls.Add(cancelButton);
            CancelButton = cancelButton;
        }

        private void SetAllChecked(bool state)
        {
            for (int i = 0; i < _alignmentList.Items.Count; i++)
                _alignmentList.SetItemChecked(i, state);
        }

        private void BrowseButton_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Choose the folder for the combined superelevation CSV";
                dialog.ShowNewFolderButton = true;

                if (Directory.Exists(OutputFolder))
                    dialog.SelectedPath = OutputFolder;

                if (dialog.ShowDialog(this) == DialogResult.OK)
                    _folderBox.Text = dialog.SelectedPath;
            }
        }

        private void ExportButton_Click(object sender, EventArgs e)
        {
            if (_alignmentList.CheckedItems.Count == 0)
            {
                MessageBox.Show(
                    this,
                    "Select at least one alignment.",
                    Text,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                DialogResult = DialogResult.None;
                return;
            }

            if (string.IsNullOrWhiteSpace(OutputFolder))
            {
                MessageBox.Show(
                    this,
                    "Choose an output folder.",
                    Text,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                DialogResult = DialogResult.None;
                return;
            }

            try
            {
                Directory.CreateDirectory(OutputFolder);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(
                    this,
                    "The output folder cannot be used:\n\n" + ex.Message,
                    Text,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                DialogResult = DialogResult.None;
            }
        }
    }
}
