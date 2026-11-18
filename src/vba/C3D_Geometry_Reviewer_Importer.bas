Attribute VB_Name = "C3D_Geometry_Reviewer_Importer"
Option Explicit

Private Const SH_SETUP As String = "Review Setup"
Private Const SH_SSD_REVIEW As String = "SSD Review"
Private Const SH_DESIGN_SPEEDS As String = "Design Speeds"
Private Const SH_SUPER_SETUP As String = "Superelevation Setup"
Private Const SH_REVIEW As String = "Geometry Review"
Private Const SH_DEPART As String = "Departure Summary"
Private Const SH_COINC As String = "Coincidence Review"
Private Const SH_ALIGN As String = "C3D Alignments"
Private Const SH_HORIZ As String = "C3D Horizontal"
Private Const SH_PROFILES As String = "C3D Profiles"
Private Const SH_VERT As String = "C3D Vertical"
Private Const SH_SUPER As String = "C3D Superelevation"
Private Const SH_SPEEDS As String = "C3D Speed Stations"
Private Const SH_CONTROL As String = "Control Zones"
Private Const SH_REF As String = "CD109 Reference Values"
Private Const SH_LISTS As String = "Lists"

Private Const TBL_SETUP As String = "tblAlignmentSetup"
Private Const TBL_SETUP_CHECKLIST As String = "tblSetupChecklist"
Private Const TBL_SSD_SUMMARY As String = "tblSSDSummary"
Private Const TBL_SSD As String = "tblManualSSD"
Private Const TBL_JUNCTIONS As String = "tblJunctions"
Private Const TBL_CONTEXTS As String = "tblContexts"
Private Const TBL_REVIEW As String = "tblGeometryReview"
Private Const TBL_DEPART As String = "tblDepartureSummary"
Private Const TBL_COINC As String = "tblCoincidenceReview"
Private Const TBL_ALIGN As String = "tblC3DAlignments"
Private Const TBL_HORIZ As String = "tblC3DHorizontal"
Private Const TBL_PROFILES As String = "tblC3DProfiles"
Private Const TBL_VERT As String = "tblC3DVertical"
Private Const TBL_SUPER As String = "tblC3DSuperelevation"
Private Const TBL_SPEEDS As String = "tblC3DSpeeds"
Private Const TBL_SPEED_ZONES As String = "tblSpeedZones"
Private Const TBL_SUPER_SETTINGS As String = "tblSuperSettings"
Private Const TBL_CONTROL As String = "tblControlZones"
Private Const TBL_CONTEXT_RULES As String = "tblContextRules"
Private Const TBL_LISTS As String = "tblLists"

Private Const LANDXML_NS As String = "xmlns:lx='http://www.landxml.org/schema/LandXML-1.2'"

Private Type FindingRec
    FindingRef As String
    Alignment As String
    Profile As String
    Source As String
    Parameter As String
    startCh As Double
    endCh As Double
    HasRange As Boolean
    achieved As Double
    HasAchieved As Boolean
    unitText As String
    speedCode As String
    desirable As Double
    HasDesirable As Boolean
    baseSteps As Long
    contextAdj As Long
    permittedSteps As Long
    permittedLimit As Double
    HasPermittedLimit As Boolean
    achievedSteps As Long       ' -1 = below CD 109 2.11 floor
    hasStepAssessment As Boolean
    ReviewStatus As String
    ClauseBasis As String
    candidateID As String
End Type

Private Type ZoneRec
    ZoneID As String
    junctionID As String
    Alignment As String
    Parameter As String
    startCh As Double
    endCh As Double
    ssd15 As Double
    refType As String
    ClauseBasis As String
End Type

Private Findings() As FindingRec
Private FindingCount As Long
Private Zones() As ZoneRec
Private zoneCount As Long
Private ReviewBuildSucceeded As Boolean

' ============================================================
' PUBLIC MACROS
' ============================================================



Public Sub RecoverReviewerTables()
    Dim oldCalc As Long
    Dim oldScreen As Boolean
    Dim oldEvents As Boolean
    Dim stage As String
    Dim errNo As Long
    Dim errDesc As String

    On Error GoTo Failed

    oldCalc = Application.Calculation
    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationAutomatic

    stage = "unprotecting workbook sheets"
    UnprotectRepairSheets

    stage = "preparing current workbook layout"
    PrepareCurrentLayout

    stage = "restoring Excel table structures and headers"
    EnsureWorkbookTableStructures
    EnsureAllTables
    VerifyRequiredTables

    stage = "restoring lists, formulas and input controls"
    UpdateLists
    RefreshSpeedZoneDerivedValues
    RefreshSuperSettingStatus
    InstallSetupStatusFormulas
    InstallJunctionFormulas
    InstallInputControls
    UpdateSetupChecklist

    stage = "restoring buttons"
    InstallButtons

    ApplyHelperVisibility
    ThisWorkbook.Worksheets(SH_SETUP).Activate

    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen

    MsgBox "Geometry Reviewer table headers and controls have been restored." & vbCrLf & _
           "You can now use Refresh Civil 3D LandXML again.", _
           vbInformation, "Civil 3D Geometry Reviewer"
    Exit Sub

Failed:
    errNo = Err.Number
    errDesc = Err.description

    On Error Resume Next
    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    On Error GoTo 0

    MsgBox "Recovery failed while " & stage & "." & vbCrLf & _
           "Error " & CStr(errNo) & ": " & errDesc, _
           vbCritical, "Civil 3D Geometry Reviewer"
End Sub

Public Sub ClearReviewData()
    Dim answer As VbMsgBoxResult
    Dim oldCalc As Long
    Dim oldScreen As Boolean
    Dim oldEvents As Boolean
    Dim stage As String
    Dim lo As ListObject
    Dim errNo As Long
    Dim errDesc As String

    answer = MsgBox( _
        "Clear all imported and project-specific review data?" & vbCrLf & vbCrLf & _
        "This clears Civil 3D imports, review setup inputs, design-speed and " & _
        "superelevation setup, SSD / junction / context inputs and generated outputs." & vbCrLf & vbCrLf & _
        "The workbook structure, tables, headers, formulas, formatting, validation " & _
        "and DMRB reference data are retained." & vbCrLf & vbCrLf & _
        "This action cannot be undone.", _
        vbYesNo + vbExclamation + vbDefaultButton2, _
        "Start New Geometry Review")

    If answer <> vbYes Then Exit Sub

    On Error GoTo Failed

    oldCalc = Application.Calculation
    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    stage = "unprotecting workbook sheets"
    UnprotectRepairSheets
    UnprotectRawSheets

    stage = "restoring table structures before clearing"
    EnsureWorkbookTableStructures

    stage = "clearing imported Civil 3D data"
    Set lo = GetTable(SH_ALIGN, TBL_ALIGN): ClearTableBody lo
    Set lo = GetTable(SH_SPEEDS, TBL_SPEEDS): ClearTableBody lo
    Set lo = GetTable(SH_HORIZ, TBL_HORIZ): ClearTableBody lo
    Set lo = GetTable(SH_PROFILES, TBL_PROFILES): ClearTableBody lo
    Set lo = GetTable(SH_VERT, TBL_VERT): ClearTableBody lo
    Set lo = GetTable(SH_SUPER, TBL_SUPER): ClearTableBody lo

    stage = "clearing alignment review setup"
    Set lo = GetTable(SH_SETUP, TBL_SETUP): ClearTableBody lo
    With ThisWorkbook.Worksheets(SH_SETUP)
        .Range("G5:G9").ClearContents
        .Range("G9").Value = 0
    End With

    stage = "clearing SSD, junction and context inputs"
    Set lo = GetTable(SH_SSD_REVIEW, TBL_SSD): ClearTableBody lo
    Set lo = GetTable(SH_SSD_REVIEW, TBL_JUNCTIONS): ClearTableBody lo
    Set lo = GetTable(SH_SSD_REVIEW, TBL_CONTEXTS): ClearTableBody lo

    stage = "clearing design-speed and superelevation setup"
    Set lo = GetTable(SH_DESIGN_SPEEDS, TBL_SPEED_ZONES): ClearTableBody lo
    Set lo = GetTable(SH_SUPER_SETUP, TBL_SUPER_SETTINGS): ClearTableBody lo

    stage = "clearing generated review outputs"
    Set lo = GetTable(SH_REVIEW, TBL_REVIEW): ClearTableBody lo
    Set lo = GetTable(SH_DEPART, TBL_DEPART): ClearTableBody lo
    Set lo = GetTable(SH_COINC, TBL_COINC): ClearTableBody lo
    Set lo = GetTable(SH_CONTROL, TBL_CONTROL): ClearTableBody lo

    stage = "restoring blank-workbook controls and formulas"
    UpdateLists
    RefreshSpeedZoneDerivedValues
    RefreshSuperSettingStatus
    InstallSetupStatusFormulas
    InstallJunctionFormulas
    InstallInputControls
    UpdateSetupChecklist
    Application.CalculateFull

    stage = "resetting review-engine state"
    FindingCount = 0
    zoneCount = 0
    Erase Findings
    Erase Zones
    ReviewBuildSucceeded = False

    ApplyHelperVisibility
    ThisWorkbook.Worksheets(SH_SETUP).Activate

    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen

    MsgBox "Review data cleared. The workbook is ready for a new LandXML import.", _
           vbInformation, "Civil 3D Geometry Reviewer"
    Exit Sub

Failed:
    errNo = Err.Number
    errDesc = Err.description

    On Error Resume Next
    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    On Error GoTo 0

    MsgBox "Clear Review Data failed while " & stage & "." & vbCrLf & _
           "Error " & CStr(errNo) & ": " & errDesc, _
           vbCritical, "Civil 3D Geometry Reviewer"
End Sub

Public Sub RepairGeometryReviewer()
    Dim stage As String
    Dim oldCalc As Long
    Dim oldScreen As Boolean
    Dim oldEvents As Boolean
    Dim errNo As Long
    Dim errDesc As String

    On Error GoTo Failed

    oldCalc = Application.Calculation
    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationAutomatic

    stage = "unprotecting workbook sheets"
    UnprotectRepairSheets

    stage = "preparing current workbook layout"
    PrepareCurrentLayout

    stage = "checking Excel table structures"
    EnsureWorkbookTableStructures

    stage = "arranging user sheets"
    ArrangeUserSheets

    stage = "enabling Excel table filters"
    EnsureAllTables
    VerifyRequiredTables

    stage = "refreshing speed-zone display values"
    RefreshSpeedZoneDerivedValues

    stage = "refreshing superelevation setting status"
    RefreshSuperSettingStatus

    stage = "installing setup formulas"
    InstallSetupStatusFormulas
    InstallJunctionFormulas

    stage = "installing input controls"
    UpdateLists
    InstallInputControls

    stage = "installing buttons"
    InstallButtons

    stage = "calculating workbook"
    Application.CalculateFull

    stage = "rebuilding review outputs"
    RebuildReviewsCore True, False
    If Not ReviewBuildSucceeded Then GoTo ReviewFailed

    stage = "restoring helper-sheet visibility"
    ApplyHelperVisibility


    ThisWorkbook.Worksheets(SH_SETUP).Activate

    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen

    MsgBox "Geometry Reviewer setup is ready.", _
           vbInformation, "Civil 3D Geometry Reviewer"
    Exit Sub

ReviewFailed:
    On Error Resume Next
    ApplyHelperVisibility
    ThisWorkbook.Worksheets(SH_SETUP).Activate
    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    On Error GoTo 0
    Exit Sub

Failed:
    errNo = Err.Number
    errDesc = Err.description

    On Error Resume Next
    ThisWorkbook.Worksheets(SH_SETUP).Activate
    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    On Error GoTo 0

    MsgBox "Geometry Reviewer setup failed while " & stage & "." & vbCrLf & _
           "Error " & CStr(errNo) & ": " & errDesc, _
           vbCritical, "Civil 3D Geometry Reviewer"
End Sub

Public Sub SetupLandXMLImporter()
    Dim stage As String
    Dim errNo As Long
    Dim errDesc As String

    On Error GoTo Failed
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationAutomatic

    stage = "preparing current workbook layout"
    PrepareCurrentLayout

    stage = "ensuring Excel table structures"
    EnsureWorkbookTableStructures
    UnprotectRawSheets

    stage = "arranging user sheets"
    ArrangeUserSheets

    stage = "enabling table filters"
    EnsureAllTables

    stage = "refreshing speed-zone display values"
    RefreshSpeedZoneDerivedValues

    stage = "refreshing superelevation setting status"
    RefreshSuperSettingStatus

    stage = "installing setup formulas"
    InstallSetupStatusFormulas
    InstallJunctionFormulas

    stage = "installing input controls"
    UpdateLists
    InstallInputControls
    Application.CalculateFull

    stage = "installing buttons"
    InstallButtons

    stage = "building current review outputs"
    RebuildReviewsCore False, False

    stage = "restoring helper-sheet visibility"
    ApplyHelperVisibility

    ThisWorkbook.Worksheets(SH_SETUP).Activate
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Civil 3D Geometry Reviewer setup is complete.", _
           vbInformation, "Civil 3D Geometry Reviewer"
    Exit Sub

Failed:
    errNo = Err.Number
    errDesc = Err.description

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error Resume Next
    ThisWorkbook.Worksheets(SH_SETUP).Activate
    On Error GoTo 0

    MsgBox "Setup failed while " & stage & "." & vbCrLf & _
           "Error " & CStr(errNo) & ": " & errDesc, _
           vbCritical, "Civil 3D Geometry Reviewer"
End Sub

Public Sub UpdateReviews()
    RebuildReviewsCore True, True
End Sub

Private Sub RebuildReviewsCore(ByVal preserveReviewerNotes As Boolean, ByVal showCompletion As Boolean)
    Dim stage As String
    Dim notes As Object
    Dim depCount As Long
    Dim pairCount As Long
    Dim errNo As Long
    Dim errDesc As String

    On Error GoTo Failed
    ReviewBuildSucceeded = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic

    If preserveReviewerNotes Then
        Set notes = SaveReviewerNotes()
    Else
        Set notes = CreateObject("Scripting.Dictionary")
    End If

    stage = "refreshing speed-zone display values"
    RefreshSpeedZoneDerivedValues

    stage = "refreshing superelevation setting status"
    RefreshSuperSettingStatus

    stage = "updating Setup Status"
    InstallSetupStatusFormulas

    stage = "checking input controls"
    InstallInputControls

    stage = "updating junction formulas"
    InstallJunctionFormulas
    Application.CalculateFull

    stage = "updating relaxation contexts"
    UpdateContextTable

    stage = "updating desirable SSD summary"
    UpdateSSDSummary

    stage = "building clause 2.13 junction zones"
    UpdateJunctionsAndBuildZones

    stage = "updating manual SSD assessments"
    UpdateManualSSDTable

    stage = "building geometry findings"
    BuildFindings

    stage = "writing Geometry Review"
    WriteGeometryReview notes

    stage = "building Coincidence Review"
    WriteCoincidenceReview

    stage = "building Departure Summary"
    WriteDepartureSummary

    stage = "updating setup summaries"
    UpdateAlignmentSummaries
    UpdateSetupChecklist

    stage = "enabling table filters"
    EnsureAllTables

    Application.CalculateFull
    ReviewBuildSucceeded = True
    ThisWorkbook.Worksheets(SH_SETUP).Activate
    Application.ScreenUpdating = True

    If showCompletion Then
        depCount = CountOutputRows(GetTable(SH_DEPART, TBL_DEPART), "DEP-")
        pairCount = CountOutputRows(GetTable(SH_COINC, TBL_COINC), "PAIR-")

        MsgBox "Reviews updated." & vbCrLf & vbCrLf & _
               CStr(FindingCount) & " geometry / SSD / superelevation findings" & vbCrLf & _
               CStr(depCount) & " departures" & vbCrLf & _
               CStr(pairCount) & " coincidence pairs", _
               vbInformation, "Civil 3D Geometry Reviewer"
    End If
    Exit Sub

Failed:
    errNo = Err.Number
    errDesc = Err.description

    Application.ScreenUpdating = True
    On Error Resume Next
    ThisWorkbook.Worksheets(SH_SETUP).Activate
    On Error GoTo 0

    MsgBox "Update Reviews failed while " & stage & "." & vbCrLf & _
           "Error " & CStr(errNo) & ": " & errDesc, _
           vbCritical, "Civil 3D Geometry Reviewer"
End Sub

Public Sub ImportRefreshLandXML()
    Dim filePath As String
    Dim xmlDoc As Object
    Dim savedSettings As Object
    Dim savedBands As Object
    Dim oldCalc As Long
    Dim oldScreen As Boolean
    Dim oldEvents As Boolean
    Dim stage As String

    On Error GoTo Failed

    oldCalc = Application.Calculation
    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    stage = "selecting LandXML file"
    filePath = DefaultXmlPath()
    If Len(filePath) = 0 Then
        filePath = PickLandXMLFile()
    ElseIf Not LocalFileExists(filePath) Then
        filePath = PickLandXMLFile()
    End If
    If Len(filePath) = 0 Then GoTo CleanExit

    PrepareCurrentLayout
    EnsureWorkbookTableStructures

    Set savedSettings = CreateObject("Scripting.Dictionary")
    SaveReviewSettings savedSettings

    Set savedBands = CreateObject("Scripting.Dictionary")
    SaveSpeedZoneBands savedBands

    stage = "opening LandXML"
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.async = False
    xmlDoc.validateOnParse = False
    xmlDoc.SetProperty "SelectionLanguage", "XPath"
    xmlDoc.SetProperty "SelectionNamespaces", LANDXML_NS

    If xmlDoc.Load(filePath) = False Then
        Err.Raise vbObjectError + 501, , "LandXML could not be read: " & xmlDoc.parseError.reason
    End If

    stage = "preparing raw import sheets"
    UnprotectRawSheets
    ResetRawTables

    stage = "importing alignments"
    ImportAlignments xmlDoc

    stage = "importing all design speed stations"
    ImportSpeedStations xmlDoc

    stage = "importing horizontal geometry"
    ImportHorizontal xmlDoc

    stage = "importing profiles and vertical geometry"
    ImportProfilesAndVertical xmlDoc

    stage = "resizing raw Excel Tables"
    ResizeRawTables

    stage = "synchronising Review Setup"
    SyncReviewSetup savedSettings
    SyncSpeedZones savedBands
    SyncSuperelevationSettings
    UpdateLists
    RecordRefresh filePath
    RefreshSpeedZoneDerivedValues
    RefreshSuperSettingStatus
    InstallSetupStatusFormulas
    InstallJunctionFormulas
    UpdateSetupChecklist

    stage = "rebuilding review outputs"
    Application.Calculation = xlCalculationAutomatic
    RebuildReviewsCore False, False

    ApplyHelperVisibility
    EnsureAllTables

    ThisWorkbook.Worksheets(SH_SETUP).Activate

    MsgBox "Civil 3D LandXML refresh complete. Superelevation crossfall is retained from the separately imported CSV.", _
           vbInformation, "Civil 3D Geometry Reviewer"

CleanExit:
    Application.StatusBar = False
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Exit Sub

Failed:
    MsgBox "LandXML refresh failed while " & stage & "." & vbCrLf & _
           "Error " & CStr(Err.Number) & ": " & Err.description, _
           vbCritical, "Civil 3D Geometry Reviewer"
    Resume CleanExit
End Sub

Public Sub ImportSuperelevationCSV()
    Dim filePath As String
    Dim csvWb As Workbook
    Dim csvWs As Worksheet
    Dim lo As ListObject
    Dim lastRow As Long
    Dim dataRows As Long
    Dim r As Long
    Dim c As Long
    Dim stage As String
    Dim expectedHeaders As Variant
    Dim oldCalc As Long
    Dim oldScreen As Boolean
    Dim oldEvents As Boolean

    On Error GoTo Failed

    oldCalc = Application.Calculation
    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    PrepareCurrentLayout
    EnsureWorkbookTableStructures

    filePath = PickSuperelevationCSVFile()
    If Len(filePath) = 0 Then GoTo CleanExit

    stage = "opening superelevation CSV"
    Set csvWb = Workbooks.Open(Filename:=filePath, Format:=2, ReadOnly:=True, Local:=True)
    Set csvWs = csvWb.Worksheets(1)

    expectedHeaders = Array( _
        "Exported At", "Drawing", "Alignment", "Superelevation Region", "Station", _
        "Description", "Smoothing Curve Length", "Left Outside Shoulder", _
        "Left Outside Lane", "Left Inside Lane", "Left Inside Shoulder", _
        "Right Outside Shoulder", "Right Outside Lane", "Right Inside Lane", _
        "Right Inside Shoulder")

    For c = 1 To 15
        If Trim$(CStr(csvWs.Cells(1, c).Value2)) <> CStr(expectedHeaders(c - 1)) Then
            Err.Raise vbObjectError + 701, "ImportSuperelevationCSV", _
                      "Unexpected CSV header in column " & CStr(c) & ". Expected '" & _
                      CStr(expectedHeaders(c - 1)) & "' but found '" & _
                      Trim$(CStr(csvWs.Cells(1, c).Value2)) & "'."
        End If
    Next c

    lastRow = csvWs.Cells(csvWs.Rows.count, 1).End(xlUp).Row
    dataRows = lastRow - 1
    If dataRows < 1 Then dataRows = 1

    stage = "writing Civil 3D superelevation critical stations"
    UnprotectRawSheets
    Set lo = GetTable(SH_SUPER, TBL_SUPER)
    ResizeTableRows lo, dataRows
    lo.DataBodyRange.ClearContents

    If lastRow >= 2 Then
        For r = 2 To lastRow
            lo.DataBodyRange.Cells(r - 1, 1).Value = CStr(csvWs.Cells(r, 1).Value2)
            lo.DataBodyRange.Cells(r - 1, 2).Value = CStr(csvWs.Cells(r, 2).Value2)
            lo.DataBodyRange.Cells(r - 1, 3).Value = CStr(csvWs.Cells(r, 3).Value2)
            lo.DataBodyRange.Cells(r - 1, 4).Value = CStr(csvWs.Cells(r, 4).Value2)
            lo.DataBodyRange.Cells(r - 1, 5).Value = CsvStationValue(csvWs.Cells(r, 5))
            lo.DataBodyRange.Cells(r - 1, 6).Value = CStr(csvWs.Cells(r, 6).Value2)
            lo.DataBodyRange.Cells(r - 1, 7).Value = CsvPlainNumber(csvWs.Cells(r, 7))
            lo.DataBodyRange.Cells(r - 1, 8).Value = CsvSlopePercent(csvWs.Cells(r, 8))
            lo.DataBodyRange.Cells(r - 1, 9).Value = CsvSlopePercent(csvWs.Cells(r, 9))
            lo.DataBodyRange.Cells(r - 1, 10).Value = CsvSlopePercent(csvWs.Cells(r, 10))
            lo.DataBodyRange.Cells(r - 1, 11).Value = CsvSlopePercent(csvWs.Cells(r, 11))
            lo.DataBodyRange.Cells(r - 1, 12).Value = CsvSlopePercent(csvWs.Cells(r, 12))
            lo.DataBodyRange.Cells(r - 1, 13).Value = CsvSlopePercent(csvWs.Cells(r, 13))
            lo.DataBodyRange.Cells(r - 1, 14).Value = CsvSlopePercent(csvWs.Cells(r, 14))
            lo.DataBodyRange.Cells(r - 1, 15).Value = CsvSlopePercent(csvWs.Cells(r, 15))
        Next r
    End If

    lo.ListColumns("Station").DataBodyRange.NumberFormat = "0.000"
    lo.ListColumns("Smoothing Curve Length").DataBodyRange.NumberFormat = "0.000"
    For c = 8 To 15
        lo.ListColumns(c).DataBodyRange.NumberFormat = "0.000"
    Next c

    csvWb.Close SaveChanges:=False
    Set csvWb = Nothing

    stage = "recording superelevation CSV refresh"
    RecordSuperelevationRefresh filePath
    UpdateSetupChecklist

    stage = "rebuilding review outputs"
    Application.Calculation = xlCalculationAutomatic
    RebuildReviewsCore False, False

    ApplyHelperVisibility
    EnsureAllTables
    ThisWorkbook.Worksheets(SH_SETUP).Activate

    MsgBox "Civil 3D superelevation CSV refresh complete." & vbCrLf & vbCrLf & _
           CStr(IIf(lastRow >= 2, lastRow - 1, 0)) & " critical-station row(s) imported.", _
           vbInformation, "Civil 3D Geometry Reviewer"

CleanExit:
    On Error Resume Next
    If Not csvWb Is Nothing Then csvWb.Close SaveChanges:=False
    On Error GoTo 0
    Application.StatusBar = False
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Exit Sub

Failed:
    MsgBox "Superelevation CSV refresh failed while " & stage & "." & vbCrLf & _
           "Error " & CStr(Err.Number) & ": " & Err.description, _
           vbCritical, "Civil 3D Geometry Reviewer"
    Resume CleanExit
End Sub

' ============================================================
' EXCEL TABLE / BUTTON SETUP
' ============================================================

Private Sub PrepareCurrentLayout()
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim headers As Variant
    Dim i As Long
    Dim isCurrent As Boolean

    Set ws = ThisWorkbook.Worksheets(SH_SUPER)
    ws.Unprotect Password:=""

    On Error Resume Next
    Set lo = ws.ListObjects(TBL_SUPER)
    On Error GoTo 0

    If Not lo Is Nothing Then
        If lo.ListColumns.count = 15 Then
            If Trim$(CStr(lo.HeaderRowRange.Cells(1, 1).Value2)) = "Exported At" Then isCurrent = True
        End If
    End If

    ' Title/note merged-cell layout is part of the workbook template. Runtime VBA
    ' writes values/formatting only and does not create, remove or repair merges.
    ws.Range("A1").Value = "Civil 3D Superelevation Critical Stations (CSV)"
    ws.Range("A2").Value = "Authoritative achieved-crossfall source imported from DMRBEXPORTSUPER. Exporter columns are preserved; station and slope text are normalised to numeric values on import for calculation. LandXML superelevation values are not used for achieved crossfall."

    If Not isCurrent Then
        For i = ws.ListObjects.count To 1 Step -1
            ws.ListObjects(i).Delete
        Next i

        ws.Range("A3:Z200").Clear

        headers = Array( _
            "Exported At", "Drawing", "Alignment", "Superelevation Region", "Station", _
            "Description", "Smoothing Curve Length", "Left Outside Shoulder", _
            "Left Outside Lane", "Left Inside Lane", "Left Inside Shoulder", _
            "Right Outside Shoulder", "Right Outside Lane", "Right Inside Lane", _
            "Right Inside Shoulder")

        For i = 0 To 14
            ws.Cells(4, i + 1).Value = headers(i)
        Next i

        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A4:O5"), , xlYes)
        lo.Name = TBL_SUPER
        lo.TableStyle = "TableStyleMedium2"
    End If

    ' Reapply the complete title/note formatting so no legacy partial formatting
    ' remains after an upgrade or repair.
    ws.Range("A1:O1").Interior.Color = RGB(31, 78, 120)
    ws.Range("A1:O1").Font.Color = RGB(255, 255, 255)
    ws.Range("A1:O1").Font.Bold = True
    ws.Range("A1:O1").Font.Size = 12

    ws.Range("A2:O2").Interior.Color = RGB(242, 242, 242)
    ws.Range("A2:O2").Font.Color = RGB(64, 64, 64)
    ws.Range("A2:O2").WrapText = True
    ws.Range("A2:O2").Font.Italic = True

    ws.Columns("A:O").ColumnWidth = 18
    ws.Columns("B:B").ColumnWidth = 30
    ws.Columns("C:C").ColumnWidth = 34
    ws.Columns("F:F").ColumnWidth = 24

    With ThisWorkbook.Worksheets(SH_SETUP)
        .Range("A2").Value = "Set alignment/profile review here. Civil 3D provides numeric design speed; select Band A/B on Design Speeds. Import superelevation separately from the DMRBEXPORTSUPER CSV."
        .Range("A3").Value = "LandXML = alignment/profile/horizontal/vertical/design-speed geometry. Superelevation CSV = authoritative achieved crossfall. Reviews remain clipped to the selected review profile extents."
        .Range("A7").Value = "Workflow"
        .Range("A8").Value = "1"
        .Range("B8").Value = "Refresh Civil 3D LandXML"
        .Range("A9").Value = "2"
        .Range("B9").Value = "Import the latest DMRBEXPORTSUPER superelevation CSV and complete required inputs"
        .Range("A10").Value = "3"
        .Range("B10").Value = "Click Update Reviews"
        .Range("F4").Value = "Current Review Status"
        .Range("F5").Value = "XML file"
        .Range("F6").Value = "XML refreshed"
        .Range("F7").Value = "SE CSV file"
        .Range("F8").Value = "SE CSV refreshed"
        .Range("F9").Value = "Imported alignments"
        .Range("F10").Value = "Alignment setup OK"
        .Rows(2).RowHeight = 34
        .Rows(3).RowHeight = 42
        .Rows(7).RowHeight = 42
        .Rows(8).RowHeight = 24
        .Rows(9).RowHeight = 38
        .Rows(10).RowHeight = 24
        .Range("A2:A3").WrapText = True
        .Range("B8:B10").WrapText = True
        .Range("G7").WrapText = True
        .Range("G7").Font.Size = 8
        .Range("B51").Value = "Where assessed; CSV required"
    End With

    ThisWorkbook.Worksheets(SH_SUPER_SETUP).Range("A2").Value = _
        "Set Environment for each active assessment range and choose Proposed or Existing retained. Achieved crossfall comes from the DMRBEXPORTSUPER Civil 3D critical-station CSV; horizontal geometry and selected-profile limits come from LandXML. CD 109 4.8.3 minimum longitudinal-gradient checks use the same CSV transition stations."
End Sub

Private Sub EnsureWorkbookTableStructures()
    EnsureTableExists SH_SETUP, TBL_SETUP, "A12:F42"
    EnsureTableExists SH_SETUP, TBL_SETUP_CHECKLIST, "A46:D51"

    EnsureTableExists SH_SSD_REVIEW, TBL_SSD, "A17:R47"
    EnsureTableExists SH_SSD_REVIEW, TBL_JUNCTIONS, "A52:O82"
    EnsureTableExists SH_SSD_REVIEW, TBL_CONTEXTS, "A87:K117"

    EnsureTableExists SH_DESIGN_SPEEDS, TBL_SPEED_ZONES, "A7:H67"
    EnsureTableExists SH_SUPER_SETUP, TBL_SUPER_SETTINGS, "A8:I38"

    EnsureTableExists SH_REVIEW, TBL_REVIEW, "A4:T5"
    EnsureTableExists SH_DEPART, TBL_DEPART, "A4:P5"
    EnsureTableExists SH_COINC, TBL_COINC, "A4:U5"

    EnsureCurrentRegionTable SH_ALIGN, TBL_ALIGN
    EnsureCurrentRegionTable SH_SPEEDS, TBL_SPEEDS
    EnsureCurrentRegionTable SH_HORIZ, TBL_HORIZ
    EnsureCurrentRegionTable SH_PROFILES, TBL_PROFILES
    EnsureCurrentRegionTable SH_VERT, TBL_VERT
    EnsureTableExists SH_SUPER, TBL_SUPER, "A4:O5"

    EnsureTableExists SH_CONTROL, TBL_CONTROL, "A4:J5"
    EnsureCurrentRegionTable SH_LISTS, TBL_LISTS

    ' Restore canonical column headings. This also repairs workbooks damaged by
    ' v6.10.8 ClearOutputArea clearing through vertically stacked tables.
    RestoreKnownTableHeaders
End Sub

Private Sub EnsureTableExists(ByVal sheetName As String, ByVal tableName As String, ByVal rangeAddress As String)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim existing As ListObject
    Dim target As Range

    Set ws = ThisWorkbook.Worksheets(sheetName)
    Set target = ws.Range(rangeAddress)

    On Error Resume Next
    Set lo = ws.ListObjects(tableName)
    On Error GoTo 0

    If Not lo Is Nothing Then Exit Sub

    Set existing = FindOverlappingTable(ws, target)

    If Not existing Is Nothing Then
        On Error GoTo RenameFailed
        existing.Name = tableName
        Exit Sub
    End If

    Set lo = ws.ListObjects.Add(xlSrcRange, target, , xlYes)
    lo.Name = tableName
    Exit Sub

RenameFailed:
    Err.Raise vbObjectError + 620, "EnsureTableExists", _
              "An Excel table already overlaps " & sheetName & "!" & rangeAddress & _
              " but could not be renamed to " & tableName & "."
End Sub


Private Sub EnsureCurrentRegionTable(ByVal sheetName As String, ByVal tableName As String)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim existing As ListObject
    Dim rng As Range

    Set ws = ThisWorkbook.Worksheets(sheetName)

    On Error Resume Next
    Set lo = ws.ListObjects(tableName)
    On Error GoTo 0

    If Not lo Is Nothing Then Exit Sub

    Set existing = FindTableContainingCell(ws, ws.Range("A4"))
    If Not existing Is Nothing Then
        existing.Name = tableName
        Exit Sub
    End If

    Set rng = ws.Range("A4").CurrentRegion
    If rng.Columns.count < 1 Then Set rng = ws.Range("A4:A5")
    If rng.Rows.count < 2 Then Set rng = rng.Resize(2, rng.Columns.count)

    Set lo = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
    lo.Name = tableName
End Sub

Private Function FindOverlappingTable(ByVal ws As Worksheet, ByVal target As Range) As ListObject
    Dim lo As ListObject

    For Each lo In ws.ListObjects
        If Not Application.Intersect(lo.Range, target) Is Nothing Then
            Set FindOverlappingTable = lo
            Exit Function
        End If
    Next lo
End Function

Private Function FindTableContainingCell(ByVal ws As Worksheet, ByVal targetCell As Range) As ListObject
    Dim lo As ListObject

    For Each lo In ws.ListObjects
        If Not Application.Intersect(lo.Range, targetCell) Is Nothing Then
            Set FindTableContainingCell = lo
            Exit Function
        End If
    Next lo
End Function


Private Sub VerifyRequiredTables()
    Dim names As Variant
    Dim sheets As Variant
    Dim i As Long
    Dim lo As ListObject

    names = Array(TBL_SETUP, TBL_SETUP_CHECKLIST, _
                  TBL_SSD, TBL_JUNCTIONS, TBL_CONTEXTS, _
                  TBL_SPEED_ZONES, TBL_SUPER_SETTINGS, _
                  TBL_REVIEW, TBL_DEPART, TBL_COINC, TBL_ALIGN, TBL_SPEEDS, TBL_HORIZ, _
                  TBL_PROFILES, TBL_VERT, TBL_SUPER)

    sheets = Array(SH_SETUP, SH_SETUP, _
                   SH_SSD_REVIEW, SH_SSD_REVIEW, SH_SSD_REVIEW, _
                   SH_DESIGN_SPEEDS, SH_SUPER_SETUP, _
                   SH_REVIEW, SH_DEPART, SH_COINC, SH_ALIGN, SH_SPEEDS, SH_HORIZ, _
                   SH_PROFILES, SH_VERT, SH_SUPER)

    For i = LBound(names) To UBound(names)
        Set lo = Nothing
        On Error Resume Next
        Set lo = ThisWorkbook.Worksheets(CStr(sheets(i))).ListObjects(CStr(names(i)))
        On Error GoTo 0

        If lo Is Nothing Then
            Err.Raise vbObjectError + 631, "VerifyRequiredTables", _
                      "Missing Excel table " & CStr(names(i)) & " on sheet " & CStr(sheets(i)) & "."
        End If

        lo.ShowHeaders = True
        lo.ShowAutoFilter = True
    Next i
End Sub

Private Sub ArrangeUserSheets()
    On Error Resume Next

    ThisWorkbook.Worksheets(SH_SSD_REVIEW).Move After:=ThisWorkbook.Worksheets(SH_SETUP)
    ThisWorkbook.Worksheets(SH_DESIGN_SPEEDS).Move After:=ThisWorkbook.Worksheets(SH_SSD_REVIEW)
    ThisWorkbook.Worksheets(SH_SUPER_SETUP).Move After:=ThisWorkbook.Worksheets(SH_DESIGN_SPEEDS)

    ThisWorkbook.Worksheets(SH_REVIEW).Move After:=ThisWorkbook.Worksheets(SH_SUPER_SETUP)
    ThisWorkbook.Worksheets(SH_DEPART).Move After:=ThisWorkbook.Worksheets(SH_REVIEW)
    ThisWorkbook.Worksheets(SH_COINC).Move After:=ThisWorkbook.Worksheets(SH_DEPART)

    On Error GoTo 0
End Sub

Private Sub EnsureAllTables()
    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            On Error Resume Next
            lo.ShowHeaders = True
            lo.ShowAutoFilter = True
            lo.TableStyle = "TableStyleMedium2"
            lo.ShowTableStyleRowStripes = True
            lo.ShowTableStyleFirstColumn = False
            lo.ShowTableStyleLastColumn = False
            On Error GoTo 0
        Next lo
    Next ws
End Sub


Private Sub InstallInputControls()
    Dim lo As ListObject
    Dim inputColor As Long
    Dim derivedColor As Long

    inputColor = RGB(255, 242, 204)
    derivedColor = RGB(221, 235, 247)

    Set lo = GetTable(SH_SETUP, TBL_SETUP)
    lo.DataBodyRange.Interior.Color = derivedColor
    SetColumnInputFill lo, "Selected Review Profile", inputColor
    SetColumnInputFill lo, "Road Type", inputColor
    SetColumnInputFill lo, "Include?", inputColor

    SetListValidation lo.ListColumns("Selected Review Profile").DataBodyRange, "Lists!$C$5:$C$34"
    SetListValidation lo.ListColumns("Road Type").DataBodyRange, "Lists!$D$5:$D$7"
    SetListValidation lo.ListColumns("Include?").DataBodyRange, "Lists!$A$5:$A$6"

    Set lo = GetTable(SH_SSD_REVIEW, TBL_SSD)
    lo.DataBodyRange.Interior.Color = derivedColor
    SetColumnInputFill lo, "SSD ID", inputColor
    SetColumnInputFill lo, "Alignment", inputColor
    SetColumnInputFill lo, "Direction", inputColor
    SetColumnInputFill lo, "Start Chainage", inputColor
    SetColumnInputFill lo, "End Chainage", inputColor
    SetColumnInputFill lo, "Achieved SSD (m)", inputColor
    SetColumnInputFill lo, "Active?", inputColor
    SetColumnInputFill lo, "Notes", inputColor

    SetListValidation lo.ListColumns("Alignment").DataBodyRange, "Lists!$B$5:$B$34"
    SetListValidation lo.ListColumns("Direction").DataBodyRange, "Lists!$G$5:$G$6"
    SetListValidation lo.ListColumns("Active?").DataBodyRange, "Lists!$A$5:$A$6"

    Set lo = GetTable(SH_SSD_REVIEW, TBL_JUNCTIONS)
    lo.DataBodyRange.Interior.Color = derivedColor
    SetColumnInputFill lo, "Junction ID", inputColor
    SetColumnInputFill lo, "Alignment", inputColor
    SetColumnInputFill lo, "Reference Type", inputColor
    SetColumnInputFill lo, "Reference Chainage", inputColor
    SetColumnInputFill lo, "Approach Direction", inputColor
    SetColumnInputFill lo, "Active?", inputColor
    SetColumnInputFill lo, "Override Start", inputColor
    SetColumnInputFill lo, "Override End", inputColor
    SetColumnInputFill lo, "Notes", inputColor

    SetListValidation lo.ListColumns("Alignment").DataBodyRange, "Lists!$B$5:$B$34"
    SetListValidation lo.ListColumns("Reference Type").DataBodyRange, "Lists!$E$5:$E$12"
    SetListValidation lo.ListColumns("Approach Direction").DataBodyRange, "Lists!$F$5:$F$7"
    SetListValidation lo.ListColumns("Active?").DataBodyRange, "Lists!$A$5:$A$6"

    Set lo = GetTable(SH_SSD_REVIEW, TBL_CONTEXTS)
    lo.DataBodyRange.Interior.Color = derivedColor
    SetColumnInputFill lo, "Context ID", inputColor
    SetColumnInputFill lo, "Alignment", inputColor
    SetColumnInputFill lo, "Start Chainage", inputColor
    SetColumnInputFill lo, "End Chainage", inputColor
    SetColumnInputFill lo, "Context Type", inputColor
    SetColumnInputFill lo, "Active?", inputColor
    SetColumnInputFill lo, "Notes", inputColor

    SetListValidation lo.ListColumns("Alignment").DataBodyRange, "Lists!$B$5:$B$34"
    SetListValidation lo.ListColumns("Context Type").DataBodyRange, "Lists!$L$5:$L$17"
    SetListValidation lo.ListColumns("Active?").DataBodyRange, "Lists!$A$5:$A$6"

    Set lo = GetTable(SH_DESIGN_SPEEDS, TBL_SPEED_ZONES)
    lo.DataBodyRange.Interior.Color = derivedColor
    SetColumnInputFill lo, "Band", inputColor
    SetListValidation lo.ListColumns("Band").DataBodyRange, "Lists!$I$5:$I$6"

    Set lo = GetTable(SH_SUPER_SETUP, TBL_SUPER_SETTINGS)
    lo.DataBodyRange.Interior.Color = inputColor
    lo.ListColumns("Status").DataBodyRange.Interior.Color = derivedColor

    SetListValidation lo.ListColumns("Alignment").DataBodyRange, "Lists!$B$5:$B$34"
    SetListValidation lo.ListColumns("Environment").DataBodyRange, "Lists!$J$5:$J$6"
    SetListValidation lo.ListColumns("Active?").DataBodyRange, "Lists!$A$5:$A$6"
    SetListValidation lo.ListColumns("Assessment Basis").DataBodyRange, "Lists!$K$5:$K$6"
End Sub

Private Sub SetColumnInputFill(ByVal lo As ListObject, ByVal columnName As String, ByVal fillColor As Long)
    If Not lo.ListColumns(columnName).DataBodyRange Is Nothing Then
        lo.ListColumns(columnName).DataBodyRange.Interior.Color = fillColor
    End If
End Sub

Private Sub SetListValidation(ByVal target As Range, ByVal sourceAddress As String)
    If target Is Nothing Then Exit Sub

    With target.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
             Formula1:="=INDIRECT(""" & sourceAddress & """)"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
        .ErrorTitle = "Select a listed value"
        .ErrorMessage = "Choose a value from the drop-down list."
    End With
End Sub

Private Sub InstallButtons()
    Dim ws As Worksheet
    Dim shp As Shape

    Set ws = ThisWorkbook.Worksheets(SH_SETUP)

    On Error Resume Next
    ws.Shapes("btnLandXMLRefresh").Delete
    ws.Shapes("btnSuperelevationCSV").Delete
    ws.Shapes("btnReviewsRefresh").Delete
    ws.Shapes("btnClearReviewData").Delete
    On Error GoTo 0

    Set shp = ws.Shapes.AddShape(5, ws.Range("A4").Left, ws.Range("A4").Top, 165, 34)
    With shp
        .Name = "btnLandXMLRefresh"
        .TextFrame.Characters.Text = "Refresh Civil 3D LandXML"
        .OnAction = "ImportRefreshLandXML"
        .Fill.ForeColor.RGB = RGB(31, 78, 121)
        .Line.ForeColor.RGB = RGB(31, 78, 121)
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Size = 10
    End With

    Set shp = ws.Shapes.AddShape(5, ws.Range("C4").Left, ws.Range("C4").Top, 165, 34)
    With shp
        .Name = "btnSuperelevationCSV"
        .TextFrame.Characters.Text = "Import Superelevation CSV"
        .OnAction = "ImportSuperelevationCSV"
        .Fill.ForeColor.RGB = RGB(47, 117, 101)
        .Line.ForeColor.RGB = RGB(47, 117, 101)
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Size = 10
    End With

    Set shp = ws.Shapes.AddShape(5, ws.Range("E4").Left, ws.Range("E4").Top, 150, 34)
    With shp
        .Name = "btnReviewsRefresh"
        .TextFrame.Characters.Text = "Update Reviews"
        .OnAction = "UpdateReviews"
        .Fill.ForeColor.RGB = RGB(91, 155, 213)
        .Line.ForeColor.RGB = RGB(91, 155, 213)
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Size = 10
    End With

    Set shp = ws.Shapes.AddShape(5, ws.Range("D8").Left, ws.Range("D8").Top, 170, 30)
    With shp
        .Name = "btnClearReviewData"
        .TextFrame.Characters.Text = "Clear / Start New Review"
        .OnAction = "ClearReviewData"
        .Fill.ForeColor.RGB = RGB(192, 80, 77)
        .Line.ForeColor.RGB = RGB(149, 55, 52)
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Size = 9
    End With
End Sub

Private Sub ClearTableBody(ByVal lo As ListObject)
    ' Deliberately clear only the table body. Do not scan below the table and
    ' do not resize it. Some sheets contain multiple vertically stacked tables.
    If lo Is Nothing Then Exit Sub
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.ClearContents
End Sub

Private Sub RestoreKnownTableHeaders()
    SetTableHeaders GetTable(SH_SETUP, TBL_SETUP), Array( _
        "Alignment", "Selected Review Profile", "Road Type", "Include?", _
        "Setup Status", "Live Review Summary")

    SetTableHeaders GetTable(SH_SETUP, TBL_SETUP_CHECKLIST), Array( _
        "Area", "When needed", "Current status", "Sheet")

    SetTableHeaders GetTable(SH_SSD_REVIEW, TBL_SSD), Array( _
        "SSD ID", "Alignment", "Direction", "Start Chainage", "End Chainage", _
        "Achieved SSD (m)", "Active?", "Review Design Speed", "Desirable SSD", _
        "Base Permitted Steps", "Context Adjustment", "Permitted Steps", _
        "Permitted Minimum SSD", "Achieved Steps", "Review Result", _
        "CD 109 2.13 Zone Overlap", "Clause Basis", "Notes")

    SetTableHeaders GetTable(SH_SSD_REVIEW, TBL_JUNCTIONS), Array( _
        "Junction ID", "Alignment", "Reference Type", "Reference Chainage", _
        "Approach Direction", "Active?", "Review Design Speed", "Desirable SSD", _
        "1.5 x SSD", "Override Start", "Override End", "Zone Start", "Zone End", _
        "CD 109 Clause 2.13 Basis", "Notes")

    SetTableHeaders GetTable(SH_SSD_REVIEW, TBL_CONTEXTS), Array( _
        "Context ID", "Alignment", "Start Chainage", "End Chainage", _
        "Context Type", "Active?", "Parameter", "Step Adjustment", _
        "Clause", "Eligibility", "Notes")

    SetTableHeaders GetTable(SH_DESIGN_SPEEDS, TBL_SPEED_ZONES), Array( _
        "Alignment", "Zone Start", "Zone End", "C3D Speed", "Band", _
        "Review Design Speed", "Source Station", "Status")

    SetTableHeaders GetTable(SH_SUPER_SETUP, TBL_SUPER_SETTINGS), Array( _
        "SE ID", "Alignment", "Start Chainage", "End Chainage", "Environment", _
        "Active?", "Status", "Notes", "Assessment Basis")

    SetTableHeaders GetTable(SH_REVIEW, TBL_REVIEW), Array( _
        "Finding Ref", "Alignment", "Profile", "Source", "Parameter", _
        "Start Chainage", "End Chainage", "Achieved", "Unit", _
        "Review Design Speed", "Desirable", "Base Permitted Steps", _
        "Context Adjustment", "Permitted Steps", "Permitted Limit", _
        "Achieved Steps", "Review Status", "Clause Basis", "Candidate ID", _
        "Reviewer Note")

    SetTableHeaders GetTable(SH_DEPART, TBL_DEPART), Array( _
        "Departure Ref", "Finding Ref", "Related Finding Ref", "Alignment", _
        "Issue Type", "Parameter", "Start Chainage", "End Chainage", "Achieved", _
        "Desirable", "Permitted / Requirement", "Related Achieved", _
        "Departure Reason", "Clause", "Control / Overlap Chainage", "Notes")

    SetTableHeaders GetTable(SH_COINC, TBL_COINC), Array( _
        "Pair Ref", "Alignment", "Finding Ref A", "Parameter A", "Achieved A", _
        "Desirable A", "Permitted A", "Steps A", "Clause A", "Finding Ref B", _
        "Parameter B", "Achieved B", "Desirable B", "Permitted B", "Steps B", _
        "Clause B", "Overlap Start", "Overlap End", "Result", "Rule", "Notes")

    SetTableHeaders GetTable(SH_ALIGN, TBL_ALIGN), Array( _
        "Alignment", "Start Chainage", "End Chainage", "Length", _
        "Imported Design Speed", "Speed Station")

    SetTableHeaders GetTable(SH_SPEEDS, TBL_SPEEDS), Array( _
        "Raw ID", "Alignment", "Station", "Speed")

    SetTableHeaders GetTable(SH_HORIZ, TBL_HORIZ), Array( _
        "Raw ID", "Alignment", "Type", "Start Chainage", "End Chainage", _
        "Length", "Radius", "Direction", "Radius Start", "Radius End", "Clothoid A")

    SetTableHeaders GetTable(SH_PROFILES, TBL_PROFILES), Array( _
        "Alignment", "Profile", "Profile Type", "State")

    SetTableHeaders GetTable(SH_VERT, TBL_VERT), Array( _
        "Raw ID", "Alignment", "Profile", "Type", "Start Chainage", _
        "End Chainage", "Length", "Grade In (%)", "Grade Out (%)", "K", _
        "PVI Chainage", "PVI Elevation")

    SetTableHeaders GetTable(SH_SUPER, TBL_SUPER), Array( _
        "Exported At", "Drawing", "Alignment", "Superelevation Region", "Station", _
        "Description", "Smoothing Curve Length", "Left Outside Shoulder", _
        "Left Outside Lane", "Left Inside Lane", "Left Inside Shoulder", _
        "Right Outside Shoulder", "Right Outside Lane", "Right Inside Lane", _
        "Right Inside Shoulder")

    SetTableHeaders GetTable(SH_CONTROL, TBL_CONTROL), Array( _
        "Zone ID", "Junction ID", "Alignment", "Parameter", "Start Chainage", _
        "End Chainage", "1.5 x SSD", "Active?", "Reference Type", "Clause Basis")

    SetTableHeaders GetTable(SH_LISTS, TBL_LISTS), Array( _
        "Yes / No", "Alignment Names", "Layout Profiles", "Road Types", _
        "Junction Reference Types", "Approach Directions", "SSD Directions", _
        "Review Design Speeds", "Band", "Environment", "Assessment Basis", _
        "Relaxation Context Types")
End Sub

Private Sub SetTableHeaders(ByVal lo As ListObject, ByVal headers As Variant)
    Dim i As Long
    Dim requiredCount As Long

    If lo Is Nothing Then Exit Sub

    requiredCount = UBound(headers) - LBound(headers) + 1
    If lo.ListColumns.count <> requiredCount Then
        Err.Raise vbObjectError + 646, "SetTableHeaders", _
                  "Table " & lo.Name & " has " & CStr(lo.ListColumns.count) & _
                  " columns; expected " & CStr(requiredCount) & "."
    End If

    For i = LBound(headers) To UBound(headers)
        lo.HeaderRowRange.Cells(1, i - LBound(headers) + 1).Value = CStr(headers(i))
    Next i
End Sub

Private Function GetTable(ByVal sheetName As String, ByVal tableName As String) As ListObject
    Set GetTable = ThisWorkbook.Worksheets(sheetName).ListObjects(tableName)
End Function

Private Sub ResizeTableRows(ByVal lo As ListObject, ByVal dataRows As Long)
    Dim firstRow As Long
    Dim firstCol As Long
    Dim colCount As Long
    Dim rng As Range

    If dataRows < 1 Then dataRows = 1

    firstRow = lo.HeaderRowRange.Row
    firstCol = lo.HeaderRowRange.Column
    colCount = lo.ListColumns.count

    Set rng = lo.Parent.Range( _
        lo.Parent.Cells(firstRow, firstCol), _
        lo.Parent.Cells(firstRow + dataRows, firstCol + colCount - 1))

    lo.Resize rng

    ' Excel can leave DataBodyRange as Nothing after Resize when the table was
    ' empty (e.g. after Clear / Start New Review). Callers write straight into
    ' DataBodyRange, which then fails with Error 91. Materialise the first row
    ' with ListRows.Add (it reuses the blank insert row, so nothing below the
    ' table is shifted) and re-apply the requested size.
    If lo.DataBodyRange Is Nothing Then
        lo.ListRows.Add
        If dataRows > 1 Then lo.Resize rng
    End If

    If lo.DataBodyRange Is Nothing Then
        Err.Raise vbObjectError + 630, "ResizeTableRows", _
                  "Table " & lo.Name & " on sheet " & lo.Parent.Name & _
                  " has no data rows after resizing to " & dataRows & " row(s)."
    End If
End Sub



Private Sub EnsureTableRowCount(ByVal lo As ListObject, ByVal requiredRows As Long)
    Dim firstRow As Long
    Dim firstCol As Long
    Dim colCount As Long
    Dim rng As Range

    If requiredRows < 1 Then requiredRows = 1

    ' Grow using ListRows.Add. This avoids the transient Nothing DataBodyRange
    ' condition Excel can produce immediately after ListObject.Resize.
    Do While lo.ListRows.count < requiredRows
        lo.ListRows.Add
    Loop

    ' Shrink only when necessary.
    If lo.ListRows.count > requiredRows Then
        firstRow = lo.HeaderRowRange.Row
        firstCol = lo.HeaderRowRange.Column
        colCount = lo.ListColumns.count

        Set rng = lo.Parent.Range( _
            lo.Parent.Cells(firstRow, firstCol), _
            lo.Parent.Cells(firstRow + requiredRows, firstCol + colCount - 1))

        lo.Resize rng
    End If

    ' Defensive fallback: Excel should now expose at least one ListRow.
    If lo.ListRows.count = 0 Then lo.ListRows.Add
End Sub

Private Sub ClearOutputArea(ByVal lo As ListObject)
    Dim ws As Worksheet
    Dim firstDataRow As Long
    Dim lastDataRow As Long
    Dim firstCol As Long
    Dim lastCol As Long
    Dim c As Long
    Dim candidateLastRow As Long

    Set ws = lo.Parent
    firstDataRow = lo.HeaderRowRange.Row + 1
    firstCol = lo.Range.Column
    lastCol = firstCol + lo.ListColumns.count - 1

    lastDataRow = lo.Range.Row + lo.Range.Rows.count - 1

    ' Values can remain below a ListObject after the table is resized smaller.
    ' Scan every output column so those stale rows are also cleared.
    For c = firstCol To lastCol
        candidateLastRow = ws.Cells(ws.Rows.count, c).End(xlUp).Row
        If candidateLastRow > lastDataRow Then lastDataRow = candidateLastRow
    Next c

    If lastDataRow >= firstDataRow Then
        ws.Range(ws.Cells(firstDataRow, firstCol), _
                 ws.Cells(lastDataRow, lastCol)).ClearContents
    End If
End Sub

Private Function TableNonBlankCount(ByVal lo As ListObject, ByVal keyCol As Long) As Long
    Dim r As Long
    If lo.DataBodyRange Is Nothing Then Exit Function

    For r = 1 To lo.ListRows.count
        If Len(Trim$(CStr(lo.DataBodyRange.Cells(r, keyCol).Value2))) > 0 Then
            TableNonBlankCount = TableNonBlankCount + 1
        End If
    Next r
End Function



Private Function CountOutputRows(ByVal lo As ListObject, ByVal prefixText As String) As Long
    Dim r As Long
    Dim v As String

    If lo.DataBodyRange Is Nothing Then Exit Function

    For r = 1 To lo.ListRows.count
        v = CStr(lo.DataBodyRange.Cells(r, 1).Value2)
        If Left$(v, Len(prefixText)) = prefixText Then
            CountOutputRows = CountOutputRows + 1
        End If
    Next r
End Function

' ============================================================
' SETUP STATUS / USER INPUT DERIVATIONS
' ============================================================

Private Sub InstallSetupStatusFormulas()
    Dim lo As ListObject
    Dim f As String

    Set lo = GetTable(SH_SETUP, TBL_SETUP)

    f = "=IF([@Alignment]="""",""""," & _
        "IF([@[Include?]]=""No"",""EXCLUDED""," & _
        "IF([@[Selected Review Profile]]="""",""PROFILE REQUIRED""," & _
        "IF(COUNTIFS(tblC3DProfiles[Alignment],[@Alignment]," & _
        "tblC3DProfiles[Profile],[@[Selected Review Profile]]," & _
        "tblC3DProfiles[Profile Type],""Layout"")=0,""PROFILE INVALID""," & _
        "IF([@[Road Type]]="""",""ROAD TYPE REQUIRED""," & _
        "IF(COUNTIF(tblSpeedZones[Alignment],[@Alignment])=0,""C3D SPEED REQUIRED""," & _
        "IF(COUNTIFS(tblSpeedZones[Alignment],[@Alignment],tblSpeedZones[Status],""<>OK"")>0," & _
        """DESIGN SPEED SETUP REQUIRED"",""OK"")))))))"

    If Not lo.DataBodyRange Is Nothing Then
        lo.ListColumns("Setup Status").DataBodyRange.Formula = f
    End If

    With ThisWorkbook.Worksheets(SH_SETUP)
        .Range("G10").Formula = "=COUNTIF(tblAlignmentSetup[Setup Status],""OK"")"
        .Range("H9").Formula = "=""Departures: ""&COUNTIF(tblDepartureSummary[Departure Ref],""DEP-*"")"
        .Range("H10").Formula = "=""Coincidence pairs: ""&COUNTIF(tblCoincidenceReview[Pair Ref],""PAIR-*"")"
    End With
End Sub

Private Sub InstallJunctionFormulas()
    Dim lo As ListObject
    Dim ws As Worksheet
    Dim firstRow As Long
    Dim lastRow As Long
    Dim r As Long
    Dim stage As String
    Dim f As String
    Dim simpleType As String
    Dim alnStart As String
    Dim alnEnd As String
    Dim errNo As Long
    Dim errDesc As String
    Dim alignmentName As String
    Dim fallbackCode As String
    Dim speedCode As String
    Dim refCh As Variant
    Dim ssd As Variant

    On Error GoTo Failed

    Set ws = ThisWorkbook.Worksheets(SH_SSD_REVIEW)
    Set lo = GetTable(SH_SSD_REVIEW, TBL_JUNCTIONS)
    If lo.DataBodyRange Is Nothing Then Exit Sub

    firstRow = lo.DataBodyRange.Row
    lastRow = firstRow + lo.ListRows.count - 1

    For r = firstRow To lastRow
        stage = "Review Design Speed row " & CStr(r)

        alignmentName = Trim$(CStr(ws.Cells(r, "B").Value2))
        refCh = ws.Cells(r, "D").Value2
        fallbackCode = AlignmentFallbackSpeedCode(alignmentName)
        speedCode = fallbackCode

        If Len(alignmentName) > 0 And IsNumeric(refCh) Then
            speedCode = ResolveReviewSpeed(alignmentName, CDbl(refCh), CDbl(refCh), fallbackCode)
        End If

        ws.Cells(r, "G").Value = speedCode

        stage = "Desirable SSD row " & CStr(r)
        If Len(speedCode) > 0 Then
            ssd = GetDesignValue(GetNumericSpeed(speedCode), "Stopping sight distance")
            ws.Cells(r, "H").Value = ssd
            ws.Cells(r, "I").Value = 1.5 * CDbl(ssd)
        Else
            ws.Cells(r, "H").ClearContents
            ws.Cells(r, "I").ClearContents
        End If

        simpleType = _
            "OR(COUNTIF($C" & r & ",""*Note 1*"")>0," & _
               "COUNTIF($C" & r & ",""*Note 2*"")>0," & _
               "COUNTIF($C" & r & ",""*Note 6*"")>0)"

        alnStart = _
            "IFERROR(INDEX('C3D Alignments'!$B$5:$B$5000," & _
            "MATCH($B" & r & ",'C3D Alignments'!$A$5:$A$5000,0)),0)"

        alnEnd = _
            "IFERROR(INDEX('C3D Alignments'!$C$5:$C$5000," & _
            "MATCH($B" & r & ",'C3D Alignments'!$A$5:$A$5000,0)),1E+99)"

        stage = "Zone Start row " & CStr(r)
        f = "=IF(OR($A" & r & "="""",$B" & r & "="""",$F" & r & "<>""Yes""),""""," & _
            "IF(AND(ISNUMBER($J" & r & "),ISNUMBER($K" & r & ")),MIN($J" & r & ",$K" & r & ")," & _
            "IF(" & simpleType & "," & _
            "IF(AND(ISNUMBER($D" & r & "),ISNUMBER($I" & r & "),$E" & r & "<>"""")," & _
            "MAX(" & alnStart & ",IF($E" & r & "=""Increasing chainage"",$D" & r & "-$I" & r & "," & _
            "IF($E" & r & "=""Decreasing chainage"",$D" & r & ",$D" & r & "-$I" & r & "))),""""),"""")))"
        ws.Cells(r, "L").Formula = f

        stage = "Zone End row " & CStr(r)
        f = "=IF(OR($A" & r & "="""",$B" & r & "="""",$F" & r & "<>""Yes""),""""," & _
            "IF(AND(ISNUMBER($J" & r & "),ISNUMBER($K" & r & ")),MAX($J" & r & ",$K" & r & ")," & _
            "IF(" & simpleType & "," & _
            "IF(AND(ISNUMBER($D" & r & "),ISNUMBER($I" & r & "),$E" & r & "<>"""")," & _
            "MIN(" & alnEnd & ",IF($E" & r & "=""Increasing chainage"",$D" & r & "," & _
            "IF($E" & r & "=""Decreasing chainage"",$D" & r & "+$I" & r & ",$D" & r & "+$I" & r & "))),""""),"""")))"
        ws.Cells(r, "M").Formula = f

        stage = "Clause basis row " & CStr(r)
        f = "=IF($C" & r & "="""",""""," & _
            "IF(COUNTIF($C" & r & ",""*Note 1*"")>0,""CD 109 2.13 NOTE 1 - minor road priority approach""," & _
            "IF(COUNTIF($C" & r & ",""*Note 2*"")>0,""CD 109 2.13 NOTE 2 - major road priority approach""," & _
            "IF(COUNTIF($C" & r & ",""*Note 3*"")>0,""CD 109 2.13 NOTE 3 - at-grade diverge taper""," & _
            "IF(COUNTIF($C" & r & ",""*Note 4*"")>0,""CD 109 2.13 NOTE 4 - at-grade diverge auxiliary lane""," & _
            "IF(COUNTIF($C" & r & ",""*Note 5*"")>0,""CD 109 2.13 NOTE 5 - at-grade merge taper""," & _
            "IF(COUNTIF($C" & r & ",""*Note 6*"")>0,""CD 109 2.13 NOTE 6 - roundabout approach""," & _
            "IF(COUNTIF($C" & r & ",""*Note 7*"")>0,""CD 109 2.13 NOTE 7 - grade-separated diverge""," & _
            "IF(COUNTIF($C" & r & ",""*Note 8*"")>0,""CD 109 2.13 NOTE 8 - grade-separated merge"","""")))))))))"
        ws.Cells(r, "N").Formula = f
    Next r

    Exit Sub

Failed:
    errNo = Err.Number
    errDesc = Err.description
    Err.Raise errNo, "InstallJunctionFormulas", stage & " - " & errDesc
End Sub

Private Sub UpdateSSDSummary()
    Dim ws As Worksheet
    Dim loZones As ListObject
    Dim speedOrder As Variant
    Dim r As Long
    Dim i As Long
    Dim speedValue As Variant
    Dim band As String
    Dim targetSpeed As Long
    Dim hasA As Boolean
    Dim hasB As Boolean
    Dim labelText As String
    Dim ssd As Variant
    Dim outRow As Long

    Set ws = ThisWorkbook.Worksheets(SH_SSD_REVIEW)
    Set loZones = GetTable(SH_DESIGN_SPEEDS, TBL_SPEED_ZONES)

    ws.Range("A8:C13").ClearContents
    speedOrder = Array(120, 100, 85, 70, 60, 50)
    outRow = 8

    For i = LBound(speedOrder) To UBound(speedOrder)
        targetSpeed = CLng(speedOrder(i))
        hasA = False
        hasB = False

        For r = 1 To loZones.ListRows.count
            If Trim$(CStr(loZones.DataBodyRange.Cells(r, 8).Value2)) = "OK" Then
                speedValue = loZones.DataBodyRange.Cells(r, 4).Value2
                band = UCase$(Trim$(CStr(loZones.DataBodyRange.Cells(r, 5).Value2)))

                If IsNumeric(speedValue) Then
                    If Abs(CDbl(speedValue) - CDbl(targetSpeed)) < 0.000001 Then
                        If band = "A" Then hasA = True
                        If band = "B" Then hasB = True
                    End If
                End If
            End If
        Next r

        If hasA Or hasB Then
            If hasA And hasB Then
                labelText = CStr(targetSpeed) & "A / " & CStr(targetSpeed) & "B"
            ElseIf hasA Then
                labelText = CStr(targetSpeed) & "A"
            Else
                labelText = CStr(targetSpeed) & "B"
            End If

            ssd = GetDesignValue(targetSpeed, "Stopping sight distance")
            ws.Cells(outRow, "A").Value = labelText
            ws.Cells(outRow, "B").Value = ssd
            ws.Cells(outRow, "C").Value = 1.5 * CDbl(ssd)
            outRow = outRow + 1
        End If
    Next i
End Sub

Private Sub UpdateContextTable()
    Dim lo As ListObject
    Dim loRules As ListObject
    Dim r As Long
    Dim i As Long
    Dim contextType As String
    Dim alignmentName As String
    Dim speedCode As String
    Dim roadType As String
    Dim param As String
    Dim adj As Variant
    Dim clauseText As String
    Dim eligible As String
    Dim dummyProfile As String
    Dim dummyInclude As String
    Dim dummyStatus As String
    Dim startCh As Variant
    Dim endCh As Variant

    Set lo = GetTable(SH_SSD_REVIEW, TBL_CONTEXTS)
    Set loRules = GetTable(SH_REF, TBL_CONTEXT_RULES)

    For r = 1 To lo.ListRows.count
        contextType = Trim$(CStr(lo.DataBodyRange.Cells(r, 5).Value2))
        alignmentName = Trim$(CStr(lo.DataBodyRange.Cells(r, 2).Value2))
        startCh = lo.DataBodyRange.Cells(r, 3).Value2
        endCh = lo.DataBodyRange.Cells(r, 4).Value2

        param = vbNullString
        adj = vbNullString
        clauseText = vbNullString
        eligible = vbNullString
        dummyProfile = vbNullString
        dummyInclude = vbNullString
        dummyStatus = vbNullString

        lo.DataBodyRange.Cells(r, 7).ClearContents
        lo.DataBodyRange.Cells(r, 8).ClearContents
        lo.DataBodyRange.Cells(r, 9).ClearContents
        lo.DataBodyRange.Cells(r, 10).ClearContents

        If Len(contextType) > 0 Then
            For i = 1 To loRules.ListRows.count
                If CStr(loRules.DataBodyRange.Cells(i, 1).Value2) = contextType Then
                    param = CStr(loRules.DataBodyRange.Cells(i, 2).Value2)
                    adj = loRules.DataBodyRange.Cells(i, 3).Value2
                    clauseText = CStr(loRules.DataBodyRange.Cells(i, 4).Value2)
                    Exit For
                End If
            Next i

            lo.DataBodyRange.Cells(r, 7).Value = param
            lo.DataBodyRange.Cells(r, 8).Value = adj
            lo.DataBodyRange.Cells(r, 9).Value = clauseText

            If Len(alignmentName) = 0 Then
                eligible = "ALIGNMENT REQUIRED"
            ElseIf Not FindAlignmentSetting(alignmentName, speedCode, roadType, dummyProfile, dummyInclude, dummyStatus) Then
                eligible = "ALIGNMENT NOT FOUND"
            Else
                If IsNumeric(startCh) And IsNumeric(endCh) Then
                    speedCode = ResolveReviewSpeed(alignmentName, CDbl(startCh), CDbl(endCh), speedCode)
                End If
                eligible = ContextEligibility(contextType, speedCode, roadType)
            End If

            lo.DataBodyRange.Cells(r, 10).Value = eligible
        End If
    Next r
End Sub

Private Function ContextEligibility(ByVal contextType As String, ByVal speedCode As String, ByVal roadType As String) As String
    Dim band As String
    Dim speedValue As Long

    band = GetBand(speedCode)
    speedValue = GetNumericSpeed(speedCode)

    If Len(speedCode) = 0 Then
        ContextEligibility = "DESIGN SPEED REQUIRED"
        Exit Function
    End If

    Select Case True
        Case Left$(contextType, 6) = "3.6(1)", Left$(contextType, 6) = "3.6(2)"
            If band = "A" Then ContextEligibility = "OK" Else ContextEligibility = "NOT ELIGIBLE"

        Case Left$(contextType, 6) = "3.6(3)"
            If roadType = "All-purpose single carriageway" Then ContextEligibility = "OK" Else ContextEligibility = "NOT ELIGIBLE"

        Case Left$(contextType, 6) = "3.7(1)"
            If roadType = "Motorway" Or roadType = "All-purpose dual carriageway" Then ContextEligibility = "OK" Else ContextEligibility = "NOT ELIGIBLE"

        Case Left$(contextType, 6) = "3.7(2)"
            If roadType = "All-purpose single carriageway" Then ContextEligibility = "OK" Else ContextEligibility = "NOT ELIGIBLE"

        Case Left$(contextType, 3) = "4.6"
            If band = "B" And roadType = "All-purpose single carriageway" Then ContextEligibility = "OK" Else ContextEligibility = "NOT ELIGIBLE"

        Case Left$(contextType, 6) = "4.7(1)"
            If band = "B" And (roadType = "Motorway" Or roadType = "All-purpose dual carriageway") Then ContextEligibility = "OK" Else ContextEligibility = "NOT ELIGIBLE"

        Case Left$(contextType, 6) = "4.7(2)"
            If band = "B" And roadType = "All-purpose single carriageway" Then ContextEligibility = "OK" Else ContextEligibility = "NOT ELIGIBLE"

        Case Left$(contextType, 6) = "5.7(1)"
            If roadType = "All-purpose single carriageway" Then ContextEligibility = "OK" Else ContextEligibility = "NOT ELIGIBLE"

        Case Left$(contextType, 6) = "5.7(2)"
            If band = "A" Then ContextEligibility = "OK" Else ContextEligibility = "NOT ELIGIBLE"

        Case Left$(contextType, 3) = "5.8"
            If roadType = "All-purpose single carriageway" Then ContextEligibility = "OK" Else ContextEligibility = "NOT ELIGIBLE"

        Case Left$(contextType, 4) = "5.10"
            If speedValue > 0 And speedValue <= 70 Then ContextEligibility = "OK" Else ContextEligibility = "NOT ELIGIBLE"

        Case Left$(contextType, 4) = "5.11"
            If (speedCode = "50B" Or speedCode = "60B" Or speedCode = "70B") And roadType = "All-purpose single carriageway" Then
                ContextEligibility = "OK"
            Else
                ContextEligibility = "NOT ELIGIBLE"
            End If

        Case Else
            ContextEligibility = "CHECK CONTEXT"
    End Select
End Function


Private Sub UpdateJunctionsAndBuildZones()
    Dim lo As ListObject
    Dim r As Long
    Dim junctionID As String
    Dim alignmentName As String
    Dim refType As String
    Dim activeText As String
    Dim zoneStart As Variant
    Dim zoneEnd As Variant
    Dim ssd15 As Variant
    Dim basis As String

    zoneCount = 0
    Erase Zones

    Application.CalculateFull

    Set lo = GetTable(SH_SSD_REVIEW, TBL_JUNCTIONS)

    For r = 1 To lo.ListRows.count
        junctionID = Trim$(CStr(lo.DataBodyRange.Cells(r, 1).Value2))
        alignmentName = Trim$(CStr(lo.DataBodyRange.Cells(r, 2).Value2))
        refType = Trim$(CStr(lo.DataBodyRange.Cells(r, 3).Value2))
        activeText = Trim$(CStr(lo.DataBodyRange.Cells(r, 6).Value2))
        ssd15 = lo.DataBodyRange.Cells(r, 9).Value2
        zoneStart = lo.DataBodyRange.Cells(r, 12).Value2
        zoneEnd = lo.DataBodyRange.Cells(r, 13).Value2
        basis = CStr(lo.DataBodyRange.Cells(r, 14).Value2)

        If activeText = "Yes" And Len(junctionID) > 0 And Len(alignmentName) > 0 Then
            If IsNumeric(ssd15) And IsNumeric(zoneStart) And IsNumeric(zoneEnd) Then
                AddZone junctionID, alignmentName, "Stopping sight distance", CDbl(zoneStart), CDbl(zoneEnd), CDbl(ssd15), refType, basis
                AddZone junctionID, alignmentName, "Crest curvature", CDbl(zoneStart), CDbl(zoneEnd), CDbl(ssd15), refType, basis
                AddZone junctionID, alignmentName, "Sag curvature", CDbl(zoneStart), CDbl(zoneEnd), CDbl(ssd15), refType, basis
            End If
        End If
    Next r

    WriteControlZones
End Sub

Private Function JunctionClauseBasis(ByVal refType As String) As String
    Select Case True
        Case InStr(1, refType, "Note 1", vbTextCompare) > 0
            JunctionClauseBasis = "CD 109 2.13 NOTE 1 - minor road priority approach"
        Case InStr(1, refType, "Note 2", vbTextCompare) > 0
            JunctionClauseBasis = "CD 109 2.13 NOTE 2 - major road priority approach"
        Case InStr(1, refType, "Note 3", vbTextCompare) > 0
            JunctionClauseBasis = "CD 109 2.13 NOTE 3 - at-grade diverge taper"
        Case InStr(1, refType, "Note 4", vbTextCompare) > 0
            JunctionClauseBasis = "CD 109 2.13 NOTE 4 - at-grade diverge auxiliary lane"
        Case InStr(1, refType, "Note 5", vbTextCompare) > 0
            JunctionClauseBasis = "CD 109 2.13 NOTE 5 - at-grade merge taper"
        Case InStr(1, refType, "Note 6", vbTextCompare) > 0
            JunctionClauseBasis = "CD 109 2.13 NOTE 6 - roundabout approach"
        Case InStr(1, refType, "Note 7", vbTextCompare) > 0
            JunctionClauseBasis = "CD 109 2.13 NOTE 7 - grade-separated diverge"
        Case InStr(1, refType, "Note 8", vbTextCompare) > 0
            JunctionClauseBasis = "CD 109 2.13 NOTE 8 - grade-separated merge"
    End Select
End Function

Private Function IsSimpleJunctionType(ByVal refType As String) As Boolean
    IsSimpleJunctionType = _
        (InStr(1, refType, "Note 1", vbTextCompare) > 0 Or _
         InStr(1, refType, "Note 2", vbTextCompare) > 0 Or _
         InStr(1, refType, "Note 6", vbTextCompare) > 0)
End Function

Private Sub CalculateSimpleZone(ByVal alignmentName As String, ByVal refCh As Double, ByVal direction As String, _
                                ByVal ssd15 As Double, ByRef zoneStart As Variant, ByRef zoneEnd As Variant)
    Dim alnStart As Double
    Dim alnEnd As Double

    If Not GetAlignmentLimits(alignmentName, alnStart, alnEnd) Then Exit Sub

    Select Case direction
        Case "Increasing chainage"
            zoneStart = Application.WorksheetFunction.Max(alnStart, refCh - ssd15)
            zoneEnd = Application.WorksheetFunction.Min(alnEnd, refCh)
        Case "Decreasing chainage"
            zoneStart = Application.WorksheetFunction.Max(alnStart, refCh)
            zoneEnd = Application.WorksheetFunction.Min(alnEnd, refCh + ssd15)
        Case "Both"
            zoneStart = Application.WorksheetFunction.Max(alnStart, refCh - ssd15)
            zoneEnd = Application.WorksheetFunction.Min(alnEnd, refCh + ssd15)
    End Select
End Sub

Private Sub SwapVariants(ByRef a As Variant, ByRef b As Variant)
    Dim t As Variant
    t = a
    a = b
    b = t
End Sub

Private Sub AddZone(ByVal junctionID As String, ByVal alignmentName As String, ByVal param As String, _
                    ByVal startCh As Double, ByVal endCh As Double, ByVal ssd15 As Double, _
                    ByVal refType As String, ByVal basis As String)
    zoneCount = zoneCount + 1
    ReDim Preserve Zones(1 To zoneCount)

    With Zones(zoneCount)
        .ZoneID = "ZONE-" & Format$(zoneCount, "000")
        .junctionID = junctionID
        .Alignment = alignmentName
        .Parameter = param
        .startCh = startCh
        .endCh = endCh
        .ssd15 = ssd15
        .refType = refType
        .ClauseBasis = basis
    End With
End Sub

Private Sub WriteControlZones()
    Dim lo As ListObject
    Dim r As Long

    Set lo = GetTable(SH_CONTROL, TBL_CONTROL)
    ResizeTableRows lo, IIf(zoneCount > 0, zoneCount, 1)
    lo.DataBodyRange.ClearContents

    For r = 1 To zoneCount
        With lo.DataBodyRange.Rows(r)
            .Cells(1, 1).Value = Zones(r).ZoneID
            .Cells(1, 2).Value = Zones(r).junctionID
            .Cells(1, 3).Value = Zones(r).Alignment
            .Cells(1, 4).Value = Zones(r).Parameter
            .Cells(1, 5).Value = Zones(r).startCh
            .Cells(1, 6).Value = Zones(r).endCh
            .Cells(1, 7).Value = Zones(r).ssd15
            .Cells(1, 8).Value = "Yes"
            .Cells(1, 9).Value = Zones(r).refType
            .Cells(1, 10).Value = Zones(r).ClauseBasis
        End With
    Next r
End Sub

Private Sub UpdateManualSSDTable()
    Dim lo As ListObject
    Dim r As Long
    Dim alignmentName As String
    Dim speedCode As String
    Dim roadType As String
    Dim profileName As String
    Dim includeText As String
    Dim setupStatus As String
    Dim achieved As Variant
    Dim startCh As Variant
    Dim endCh As Variant
    Dim activeText As String

    Dim desirable As Double
    Dim baseSteps As Long
    Dim contextAdj As Long
    Dim permittedSteps As Long
    Dim permittedLimit As Double
    Dim achievedSteps As Long
    Dim statusText As String
    Dim clauseText As String
    Dim overlapCount As Long
    Dim c As Long

    Set lo = GetTable(SH_SSD_REVIEW, TBL_SSD)

    For r = 1 To lo.ListRows.count
        alignmentName = Trim$(CStr(lo.DataBodyRange.Cells(r, 2).Value2))
        startCh = lo.DataBodyRange.Cells(r, 4).Value2
        endCh = lo.DataBodyRange.Cells(r, 5).Value2
        achieved = lo.DataBodyRange.Cells(r, 6).Value2
        activeText = Trim$(CStr(lo.DataBodyRange.Cells(r, 7).Value2))

        For c = 8 To 17
            lo.DataBodyRange.Cells(r, c).ClearContents
        Next c

        If Len(alignmentName) > 0 Then
            If FindAlignmentSetting(alignmentName, speedCode, roadType, profileName, includeText, setupStatus) Then
                If IsNumeric(startCh) And IsNumeric(endCh) Then
                    speedCode = ResolveReviewSpeed(alignmentName, CDbl(startCh), CDbl(endCh), speedCode)
                End If
                lo.DataBodyRange.Cells(r, 8).Value = speedCode

                If activeText = "Yes" And setupStatus = "OK" And IsNumeric(startCh) And IsNumeric(endCh) And IsNumeric(achieved) Then
                    EvaluateStepParameter "Stopping sight distance", CDbl(achieved), alignmentName, CDbl(startCh), CDbl(endCh), _
                                          speedCode, roadType, desirable, baseSteps, contextAdj, permittedSteps, permittedLimit, _
                                          achievedSteps, statusText, clauseText

                    clauseText = AppendSpeedInterfaceNote(clauseText, alignmentName, CDbl(startCh), CDbl(endCh))
                    overlapCount = CountZoneOverlaps(alignmentName, "Stopping sight distance", CDbl(startCh), CDbl(endCh))

                    lo.DataBodyRange.Cells(r, 9).Value = desirable
                    lo.DataBodyRange.Cells(r, 10).Value = baseSteps
                    lo.DataBodyRange.Cells(r, 11).Value = contextAdj
                    lo.DataBodyRange.Cells(r, 12).Value = permittedSteps
                    lo.DataBodyRange.Cells(r, 13).Value = permittedLimit
                    If achievedSteps = -1 Then
                        lo.DataBodyRange.Cells(r, 14).Value = "BELOW FLOOR"
                    Else
                        lo.DataBodyRange.Cells(r, 14).Value = achievedSteps
                    End If
                    lo.DataBodyRange.Cells(r, 15).Value = statusText
                    lo.DataBodyRange.Cells(r, 16).Value = overlapCount
                    lo.DataBodyRange.Cells(r, 17).Value = clauseText
                End If
            End If
        End If
    Next r
End Sub

Private Sub UpdateSetupChecklist()
    Dim loOut As ListObject
    Dim loSetup As ListObject
    Dim loZones As ListObject
    Dim loSSD As ListObject
    Dim loJ As ListObject
    Dim loC As ListObject
    Dim loSE As ListObject
    Dim r As Long
    Dim includedCount As Long
    Dim setupIssues As Long
    Dim zoneCount As Long
    Dim zoneIssues As Long
    Dim ssdCount As Long
    Dim junctionCount As Long
    Dim contextCount As Long
    Dim seCount As Long
    Dim seRequired As Long
    Dim seCsvAlignmentCount As Long
    Dim seCsv As Object
    Dim seAln As String
    Dim statusText As String

    Set loOut = GetTable(SH_SETUP, TBL_SETUP_CHECKLIST)
    Set loSetup = GetTable(SH_SETUP, TBL_SETUP)
    Set loZones = GetTable(SH_DESIGN_SPEEDS, TBL_SPEED_ZONES)
    Set loSSD = GetTable(SH_SSD_REVIEW, TBL_SSD)
    Set loJ = GetTable(SH_SSD_REVIEW, TBL_JUNCTIONS)
    Set loC = GetTable(SH_SSD_REVIEW, TBL_CONTEXTS)
    Set loSE = GetTable(SH_SUPER_SETUP, TBL_SUPER_SETTINGS)
    Set seCsv = CreateObject("Scripting.Dictionary")

    With GetTable(SH_SUPER, TBL_SUPER)
        For r = 1 To .ListRows.count
            seAln = Trim$(CStr(.DataBodyRange.Cells(r, 3).Value2))
            If Len(seAln) > 0 Then seCsv.Item(seAln) = True
        Next r
    End With
    seCsvAlignmentCount = seCsv.count

    If loOut.ListRows.count < 5 Then ResizeTableRows loOut, 5

    For r = 1 To loSetup.ListRows.count
        If Trim$(CStr(loSetup.DataBodyRange.Cells(r, 1).Value2)) <> "" And _
           Trim$(CStr(loSetup.DataBodyRange.Cells(r, 4).Value2)) = "Yes" Then
            includedCount = includedCount + 1
            If Trim$(CStr(loSetup.DataBodyRange.Cells(r, 5).Value2)) <> "OK" Then setupIssues = setupIssues + 1
        End If
    Next r

    If includedCount = 0 Then
        statusText = "No included alignments"
    ElseIf setupIssues = 0 Then
        statusText = "OK - " & CStr(includedCount) & " included"
    Else
        statusText = "ACTION - " & CStr(setupIssues) & " need setup"
    End If
    loOut.DataBodyRange.Cells(1, 3).Value = statusText

    For r = 1 To loZones.ListRows.count
        If Trim$(CStr(loZones.DataBodyRange.Cells(r, 1).Value2)) <> "" Then
            zoneCount = zoneCount + 1
            If Trim$(CStr(loZones.DataBodyRange.Cells(r, 8).Value2)) <> "OK" Then zoneIssues = zoneIssues + 1
        End If
    Next r

    If zoneCount = 0 Then
        statusText = "ACTION - no C3D speed zones"
    ElseIf zoneIssues > 0 Then
        statusText = "ACTION - " & CStr(zoneIssues) & " speed zone(s) need setup"
    Else
        statusText = "OK - " & CStr(zoneCount) & " speed zone(s)"
    End If
    loOut.DataBodyRange.Cells(2, 3).Value = statusText

    For r = 1 To loSSD.ListRows.count
        If Trim$(CStr(loSSD.DataBodyRange.Cells(r, 2).Value2)) <> "" And _
           Trim$(CStr(loSSD.DataBodyRange.Cells(r, 7).Value2)) = "Yes" Then ssdCount = ssdCount + 1
    Next r
    If ssdCount = 0 Then
        loOut.DataBodyRange.Cells(3, 3).Value = "No active SSD records"
    Else
        loOut.DataBodyRange.Cells(3, 3).Value = CStr(ssdCount) & " active SSD"
    End If

    For r = 1 To loJ.ListRows.count
        If Trim$(CStr(loJ.DataBodyRange.Cells(r, 2).Value2)) <> "" And _
           Trim$(CStr(loJ.DataBodyRange.Cells(r, 6).Value2)) = "Yes" Then junctionCount = junctionCount + 1
    Next r
    For r = 1 To loC.ListRows.count
        If Trim$(CStr(loC.DataBodyRange.Cells(r, 2).Value2)) <> "" And _
           Trim$(CStr(loC.DataBodyRange.Cells(r, 6).Value2)) = "Yes" Then contextCount = contextCount + 1
    Next r
    loOut.DataBodyRange.Cells(4, 3).Value = CStr(junctionCount) & " junction / " & CStr(contextCount) & " context record(s)"

    For r = 1 To loSE.ListRows.count
        If Trim$(CStr(loSE.DataBodyRange.Cells(r, 2).Value2)) <> "" And _
           Trim$(CStr(loSE.DataBodyRange.Cells(r, 6).Value2)) = "Yes" Then
            seCount = seCount + 1
            Select Case Trim$(CStr(loSE.DataBodyRange.Cells(r, 7).Value2))
                Case "ENVIRONMENT REQUIRED", "ASSESSMENT BASIS REQUIRED"
                    seRequired = seRequired + 1
            End Select
        End If
    Next r

    If seCsvAlignmentCount = 0 Then
        statusText = "ACTION - superelevation CSV not imported"
    ElseIf seCount = 0 Then
        statusText = "CSV: " & CStr(seCsvAlignmentCount) & " alignment(s); no active ranges"
    ElseIf seRequired > 0 Then
        statusText = "CSV: " & CStr(seCsvAlignmentCount) & " alignment(s); ACTION - " & CStr(seRequired) & " need input"
    Else
        statusText = "OK - " & CStr(seCsvAlignmentCount) & " CSV alignment(s); " & CStr(seCount) & " active range(s)"
    End If
    loOut.DataBodyRange.Cells(5, 3).Value = statusText
End Sub

' ============================================================
' REVIEW ENGINE
' ============================================================

Private Sub BuildFindings()
    FindingCount = 0
    Erase Findings

    BuildHorizontalFindings
    BuildVerticalFindings
    BuildSuperelevationFindings
    BuildManualSSDFindings
End Sub

Private Sub BuildHorizontalFindings()
    Dim lo As ListObject
    Dim r As Long
    Dim alignmentName As String
    Dim geomType As String
    Dim speedCode As String
    Dim roadType As String
    Dim profileName As String
    Dim includeText As String
    Dim setupStatus As String
    Dim startCh As Variant
    Dim endCh As Variant
    Dim achieved As Variant
    Dim rawID As String

    Dim desirable As Double
    Dim baseSteps As Long
    Dim contextAdj As Long
    Dim permittedSteps As Long
    Dim permittedLimit As Double
    Dim achievedSteps As Long
    Dim statusText As String
    Dim clauseText As String

    Set lo = GetTable(SH_HORIZ, TBL_HORIZ)

    For r = 1 To lo.ListRows.count
        geomType = CStr(lo.DataBodyRange.Cells(r, 3).Value2)

        If geomType = "Curve" Then
            alignmentName = CStr(lo.DataBodyRange.Cells(r, 2).Value2)

            If FindAlignmentSetting(alignmentName, speedCode, roadType, profileName, includeText, setupStatus) Then
                If includeText = "Yes" And setupStatus = "OK" Then
                    startCh = lo.DataBodyRange.Cells(r, 4).Value2
                    endCh = lo.DataBodyRange.Cells(r, 5).Value2
                    achieved = lo.DataBodyRange.Cells(r, 7).Value2
                    rawID = CStr(lo.DataBodyRange.Cells(r, 1).Value2)

                    If IsNumeric(startCh) And IsNumeric(endCh) And IsNumeric(achieved) Then
                        speedCode = ResolveReviewSpeed(alignmentName, CDbl(startCh), CDbl(endCh), speedCode)

                        EvaluateStepParameter "Horizontal curvature", CDbl(achieved), alignmentName, CDbl(startCh), CDbl(endCh), _
                                              speedCode, roadType, desirable, baseSteps, contextAdj, permittedSteps, permittedLimit, _
                                              achievedSteps, statusText, clauseText

                        clauseText = AppendSpeedInterfaceNote(clauseText, alignmentName, CDbl(startCh), CDbl(endCh))

                        AddFinding alignmentName, vbNullString, "Horizontal", "Horizontal curvature", _
                                   CDbl(startCh), CDbl(endCh), CDbl(achieved), "m", speedCode, _
                                   desirable, baseSteps, contextAdj, permittedSteps, permittedLimit, _
                                   achievedSteps, True, statusText, clauseText, "H:" & rawID
                    End If
                End If
            End If
        End If
    Next r
End Sub

Private Sub BuildVerticalFindings()
    Dim lo As ListObject
    Dim r As Long
    Dim alignmentName As String
    Dim profileName As String
    Dim selectedProfile As String
    Dim vertType As String
    Dim speedCode As String
    Dim roadType As String
    Dim includeText As String
    Dim setupStatus As String
    Dim startCh As Variant
    Dim endCh As Variant
    Dim kValue As Variant
    Dim gradeIn As Variant
    Dim rawID As String
    Dim parameterName As String

    Dim desirable As Double
    Dim baseSteps As Long
    Dim contextAdj As Long
    Dim permittedSteps As Long
    Dim permittedLimit As Double
    Dim achievedSteps As Long
    Dim statusText As String
    Dim clauseText As String

    Set lo = GetTable(SH_VERT, TBL_VERT)

    For r = 1 To lo.ListRows.count
        alignmentName = CStr(lo.DataBodyRange.Cells(r, 2).Value2)
        profileName = CStr(lo.DataBodyRange.Cells(r, 3).Value2)

        If FindAlignmentSetting(alignmentName, speedCode, roadType, selectedProfile, includeText, setupStatus) Then
            If includeText = "Yes" And setupStatus = "OK" And profileName = selectedProfile Then
                vertType = CStr(lo.DataBodyRange.Cells(r, 4).Value2)
                startCh = lo.DataBodyRange.Cells(r, 5).Value2
                endCh = lo.DataBodyRange.Cells(r, 6).Value2
                rawID = CStr(lo.DataBodyRange.Cells(r, 1).Value2)

                If IsNumeric(startCh) And IsNumeric(endCh) Then
                    speedCode = ResolveReviewSpeed(alignmentName, CDbl(startCh), CDbl(endCh), speedCode)

                    If vertType = "Crest" Or vertType = "Sag" Then
                        kValue = lo.DataBodyRange.Cells(r, 10).Value2

                        If IsNumeric(kValue) Then
                            If vertType = "Crest" Then
                                parameterName = "Crest curvature"
                            Else
                                parameterName = "Sag curvature"
                            End If

                            EvaluateStepParameter parameterName, CDbl(kValue), alignmentName, CDbl(startCh), CDbl(endCh), _
                                                  speedCode, roadType, desirable, baseSteps, contextAdj, permittedSteps, permittedLimit, _
                                                  achievedSteps, statusText, clauseText

                            clauseText = AppendSpeedInterfaceNote(clauseText, alignmentName, CDbl(startCh), CDbl(endCh))

                            AddFinding alignmentName, profileName, "Vertical", parameterName, _
                                       CDbl(startCh), CDbl(endCh), CDbl(kValue), "K", speedCode, _
                                       desirable, baseSteps, contextAdj, permittedSteps, permittedLimit, _
                                       achievedSteps, True, statusText, clauseText, "K:" & rawID
                        End If

                    ElseIf vertType = "Tangent" Then
                        gradeIn = lo.DataBodyRange.Cells(r, 8).Value2

                        If IsNumeric(gradeIn) Then
                            AddLongitudinalGradientFinding alignmentName, profileName, _
                                                           CDbl(startCh), CDbl(endCh), Abs(CDbl(gradeIn)), _
                                                           speedCode, roadType, "G:" & rawID, "Vertical tangent"
                        End If
                    End If
                End If
            End If
        End If
    Next r

    BuildFallbackCurveGradientFindings lo
End Sub

Private Sub BuildFallbackCurveGradientFindings(ByVal lo As ListObject)
    Dim r As Long
    Dim alignmentName As String
    Dim profileName As String
    Dim selectedProfile As String
    Dim vertType As String
    Dim speedCode As String
    Dim roadType As String
    Dim includeText As String
    Dim setupStatus As String
    Dim startCh As Variant
    Dim endCh As Variant
    Dim gradeIn As Variant
    Dim gradeOut As Variant
    Dim rawID As String

    For r = 1 To lo.ListRows.count
        alignmentName = CStr(lo.DataBodyRange.Cells(r, 2).Value2)
        profileName = CStr(lo.DataBodyRange.Cells(r, 3).Value2)
        vertType = CStr(lo.DataBodyRange.Cells(r, 4).Value2)

        If vertType = "Crest" Or vertType = "Sag" Then
            If FindAlignmentSetting(alignmentName, speedCode, roadType, selectedProfile, includeText, setupStatus) Then
                If includeText = "Yes" And setupStatus = "OK" And profileName = selectedProfile Then
                    startCh = lo.DataBodyRange.Cells(r, 5).Value2
                    endCh = lo.DataBodyRange.Cells(r, 6).Value2
                    gradeIn = lo.DataBodyRange.Cells(r, 8).Value2
                    gradeOut = lo.DataBodyRange.Cells(r, 9).Value2
                    rawID = CStr(lo.DataBodyRange.Cells(r, 1).Value2)

                    If IsNumeric(startCh) And IsNumeric(gradeIn) Then
                        If Not HasTangentEndingAt(lo, alignmentName, profileName, CDbl(startCh)) And _
                           Not HasCurveEndingAt(lo, alignmentName, profileName, CDbl(startCh), r) Then

                            speedCode = ResolveReviewSpeed(alignmentName, CDbl(startCh), CDbl(startCh), speedCode)

                            AddLongitudinalGradientFinding alignmentName, profileName, _
                                                           CDbl(startCh), CDbl(startCh), Abs(CDbl(gradeIn)), _
                                                           speedCode, roadType, "GP:" & rawID & ":IN", _
                                                           "Unrepresented curve entry grade"
                        End If
                    End If

                    If IsNumeric(endCh) And IsNumeric(gradeOut) Then
                        If Not HasTangentStartingAt(lo, alignmentName, profileName, CDbl(endCh)) Then
                            speedCode = AlignmentFallbackSpeedCode(alignmentName)
                            speedCode = ResolveReviewSpeed(alignmentName, CDbl(endCh), CDbl(endCh), speedCode)

                            AddLongitudinalGradientFinding alignmentName, profileName, _
                                                           CDbl(endCh), CDbl(endCh), Abs(CDbl(gradeOut)), _
                                                           speedCode, roadType, "GP:" & rawID & ":OUT", _
                                                           "Unrepresented curve exit grade"
                        End If
                    End If
                End If
            End If
        End If
    Next r
End Sub

Private Sub AddLongitudinalGradientFinding(ByVal alignmentName As String, ByVal profileName As String, _
                                            ByVal startCh As Double, ByVal endCh As Double, _
                                            ByVal achievedGrade As Double, ByVal speedCode As String, _
                                            ByVal roadType As String, ByVal candidateID As String, _
                                            ByVal sourceText As String)

    Dim desiredGrad As Double
    Dim permittedGrad As Double
    Dim statusText As String

    GetGradientLimits roadType, desiredGrad, permittedGrad

    If LEQ3(achievedGrade, desiredGrad) Then
        statusText = "COMPLIANT"
    ElseIf LEQ3(achievedGrade, permittedGrad) Then
        statusText = "RELAXATION - CD 109 5.1"
    Else
        statusText = "DEPARTURE - exceeds CD 109 5.1 permitted gradient"
    End If

    AddFinding alignmentName, profileName, sourceText, "Longitudinal gradient", _
               startCh, endCh, achievedGrade, "%", speedCode, _
               desiredGrad, 0, 0, 0, permittedGrad, _
               0, False, statusText, "CD 109 5.1 / Table 5.1", candidateID
End Sub

Private Function HasTangentEndingAt(ByVal lo As ListObject, ByVal alignmentName As String, _
                                    ByVal profileName As String, ByVal chainage As Double) As Boolean
    Dim r As Long
    Dim ch As Variant

    For r = 1 To lo.ListRows.count
        If CStr(lo.DataBodyRange.Cells(r, 2).Value2) = alignmentName And _
           CStr(lo.DataBodyRange.Cells(r, 3).Value2) = profileName And _
           CStr(lo.DataBodyRange.Cells(r, 4).Value2) = "Tangent" Then

            ch = lo.DataBodyRange.Cells(r, 6).Value2
            If IsNumeric(ch) Then
                If Abs(CDbl(ch) - chainage) <= 0.001 Then
                    HasTangentEndingAt = True
                    Exit Function
                End If
            End If
        End If
    Next r
End Function

Private Function HasTangentStartingAt(ByVal lo As ListObject, ByVal alignmentName As String, _
                                      ByVal profileName As String, ByVal chainage As Double) As Boolean
    Dim r As Long
    Dim ch As Variant

    For r = 1 To lo.ListRows.count
        If CStr(lo.DataBodyRange.Cells(r, 2).Value2) = alignmentName And _
           CStr(lo.DataBodyRange.Cells(r, 3).Value2) = profileName And _
           CStr(lo.DataBodyRange.Cells(r, 4).Value2) = "Tangent" Then

            ch = lo.DataBodyRange.Cells(r, 5).Value2
            If IsNumeric(ch) Then
                If Abs(CDbl(ch) - chainage) <= 0.001 Then
                    HasTangentStartingAt = True
                    Exit Function
                End If
            End If
        End If
    Next r
End Function

Private Function HasCurveEndingAt(ByVal lo As ListObject, ByVal alignmentName As String, _
                                  ByVal profileName As String, ByVal chainage As Double, _
                                  ByVal excludeRow As Long) As Boolean
    Dim r As Long
    Dim ch As Variant
    Dim vertType As String

    For r = 1 To lo.ListRows.count
        If r <> excludeRow Then
            If CStr(lo.DataBodyRange.Cells(r, 2).Value2) = alignmentName And _
               CStr(lo.DataBodyRange.Cells(r, 3).Value2) = profileName Then

                vertType = CStr(lo.DataBodyRange.Cells(r, 4).Value2)
                If vertType = "Crest" Or vertType = "Sag" Then
                    ch = lo.DataBodyRange.Cells(r, 6).Value2
                    If IsNumeric(ch) Then
                        If Abs(CDbl(ch) - chainage) <= 0.001 Then
                            HasCurveEndingAt = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next r
End Function

Private Function GetSelectedProfileExtent(ByVal alignmentName As String, ByVal profileName As String, _
                                          ByRef profileStart As Double, ByRef profileEnd As Double) As Boolean
    Dim lo As ListObject
    Dim r As Long
    Dim s As Variant
    Dim e As Variant
    Dim thisStart As Double
    Dim thisEnd As Double
    Dim found As Boolean

    Set lo = GetTable(SH_VERT, TBL_VERT)

    For r = 1 To lo.ListRows.count
        If Trim$(CStr(lo.DataBodyRange.Cells(r, 2).Value2)) = alignmentName And _
           Trim$(CStr(lo.DataBodyRange.Cells(r, 3).Value2)) = profileName Then

            s = lo.DataBodyRange.Cells(r, 5).Value2
            e = lo.DataBodyRange.Cells(r, 6).Value2

            If IsNumeric(s) And IsNumeric(e) Then
                thisStart = LesserDouble(CDbl(s), CDbl(e))
                thisEnd = GreaterDouble(CDbl(s), CDbl(e))

                If Not found Then
                    profileStart = thisStart
                    profileEnd = thisEnd
                    found = True
                Else
                    If thisStart < profileStart Then profileStart = thisStart
                    If thisEnd > profileEnd Then profileEnd = thisEnd
                End If
            End If
        End If
    Next r

    GetSelectedProfileExtent = found
End Function

Private Sub BuildSuperelevationFindings()
    Dim loH As ListObject
    Dim r As Long
    Dim i As Long
    Dim alignmentName As String
    Dim geomType As String
    Dim rawID As String
    Dim baseSpeedCode As String
    Dim roadType As String
    Dim selectedProfile As String
    Dim includeText As String
    Dim setupStatus As String
    Dim curveStart As Double
    Dim curveEnd As Double
    Dim reviewStart As Double
    Dim reviewEnd As Double
    Dim profileStart As Double
    Dim profileEnd As Double
    Dim radius As Double
    Dim points() As Double
    Dim pointCount As Long

    Set loH = GetTable(SH_HORIZ, TBL_HORIZ)

    For r = 1 To loH.ListRows.count
        geomType = Trim$(CStr(loH.DataBodyRange.Cells(r, 3).Value2))

        If geomType = "Curve" Then
            alignmentName = Trim$(CStr(loH.DataBodyRange.Cells(r, 2).Value2))

            If FindAlignmentSetting(alignmentName, baseSpeedCode, roadType, selectedProfile, includeText, setupStatus) Then
                If includeText = "Yes" And setupStatus = "OK" Then
                    If GetSelectedProfileExtent(alignmentName, selectedProfile, profileStart, profileEnd) Then
                        If IsNumeric(loH.DataBodyRange.Cells(r, 4).Value2) And _
                           IsNumeric(loH.DataBodyRange.Cells(r, 5).Value2) And _
                           IsNumeric(loH.DataBodyRange.Cells(r, 7).Value2) Then

                            curveStart = LesserDouble(CDbl(loH.DataBodyRange.Cells(r, 4).Value2), _
                                                      CDbl(loH.DataBodyRange.Cells(r, 5).Value2))
                            curveEnd = GreaterDouble(CDbl(loH.DataBodyRange.Cells(r, 4).Value2), _
                                                     CDbl(loH.DataBodyRange.Cells(r, 5).Value2))

                            reviewStart = GreaterDouble(curveStart, profileStart)
                            reviewEnd = LesserDouble(curveEnd, profileEnd)

                            If reviewEnd - reviewStart > 0.001 Then
                                radius = Abs(CDbl(loH.DataBodyRange.Cells(r, 7).Value2))
                                rawID = CStr(loH.DataBodyRange.Cells(r, 1).Value2)

                                CollectXfallBreakpoints alignmentName, reviewStart, reviewEnd, points, pointCount

                                For i = 1 To pointCount - 1
                                    If points(i + 1) - points(i) > 0.001 Then
                                        AddInterpolatedSuperelevationFinding alignmentName, selectedProfile, roadType, _
                                                                            baseSpeedCode, radius, points(i), points(i + 1), _
                                                                            profileStart, profileEnd, _
                                                                            rawID & ":" & CStr(i)
                                    End If
                                Next i
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next r

    BuildSuperelevationTransitionFindingsFromCSV
End Sub

Private Sub AddInterpolatedSuperelevationFinding(ByVal alignmentName As String, ByVal profileName As String, _
                                                 ByVal roadType As String, ByVal baseSpeedCode As String, _
                                                 ByVal radius As Double, ByVal startCh As Double, ByVal endCh As Double, _
                                                 ByVal profileStart As Double, ByVal profileEnd As Double, _
                                                 ByVal candidateID As String)
    Dim speedCode As String
    Dim speedValue As Long
    Dim v2r As Double
    Dim environment As String
    Dim assessmentBasis As String
    Dim hasSetting As Boolean
    Dim requiredSuper As Double
    Dim rawRequired As Double
    Dim maxSuper As Double
    Dim leftStart As Double
    Dim leftEnd As Double
    Dim rightStart As Double
    Dim rightEnd As Double
    Dim hasLeftStart As Boolean
    Dim hasLeftEnd As Boolean
    Dim hasRightStart As Boolean
    Dim hasRightEnd As Boolean
    Dim haveUsableSide As Boolean
    Dim achievedMin As Double
    Dim achievedMax As Double
    Dim sideMin As Double
    Dim statusText As String
    Dim clauseText As String

    speedCode = ResolveReviewSpeed(alignmentName, startCh, endCh, baseSpeedCode)
    speedValue = GetNumericSpeed(speedCode)

    If radius <= 0# Or speedValue <= 0 Then Exit Sub

    v2r = (CDbl(speedValue) * CDbl(speedValue)) / radius
    If v2r <= 5# Then Exit Sub

    hasSetting = GetSuperelevationSettingAt(alignmentName, (startCh + endCh) / 2#, environment, assessmentBasis)

    If Not hasSetting Or (environment <> "Rural" And environment <> "Urban") Then
        statusText = "SETUP REQUIRED - Environment in Superelevation Settings"
        clauseText = "CD 109 4.1-4.4; Table 2.10"
        AddSuperelevationFinding alignmentName, profileName, startCh, endCh, False, 0#, _
                                 speedCode, 0#, 0#, statusText, clauseText, candidateID
        Exit Sub
    End If

    maxSuper = GetMaximumSuperelevation(environment)

    If v2r > 7# Then
        rawRequired = (CDbl(speedValue) * CDbl(speedValue)) / (2.828 * radius)
        requiredSuper = LesserDouble(rawRequired, maxSuper)
        clauseText = "CD 109 4.2 Equation 4.2; requirement capped by 4.3/4.4 at " & _
                     Format$(maxSuper, "0.###") & "%; Table 2.10"
    Else
        requiredSuper = 2.5
        clauseText = "CD 109 4.1; Table 2.10"
    End If

    clauseText = AppendSpeedInterfaceNote(clauseText, alignmentName, startCh, endCh)
    clauseText = clauseText & "; achieved crossfall from Civil 3D superelevation CSV critical-station interpolation within selected-profile extent"

    GetCrossfallsAtChainageWithinProfile alignmentName, startCh, profileStart, profileEnd, _
                                         leftStart, hasLeftStart, rightStart, hasRightStart
    GetCrossfallsAtChainageWithinProfile alignmentName, endCh, profileStart, profileEnd, _
                                         leftEnd, hasLeftEnd, rightEnd, hasRightEnd

    achievedMin = 1E+30
    achievedMax = 0#

    If hasLeftStart And hasLeftEnd Then
        sideMin = MinimumAbsoluteLinearXfall(leftStart, leftEnd)
        achievedMin = LesserDouble(achievedMin, sideMin)
        achievedMax = GreaterDouble(achievedMax, GreaterDouble(Abs(leftStart), Abs(leftEnd)))
        haveUsableSide = True
    End If

    If hasRightStart And hasRightEnd Then
        sideMin = MinimumAbsoluteLinearXfall(rightStart, rightEnd)
        achievedMin = LesserDouble(achievedMin, sideMin)
        achievedMax = GreaterDouble(achievedMax, GreaterDouble(Abs(rightStart), Abs(rightEnd)))
        haveUsableSide = True
    End If

    If Not haveUsableSide Then
        If assessmentBasis = "Existing retained" Then
            statusText = "CHECK - existing retained; no usable outside-lane xfall in superelevation CSV"
            clauseText = clauseText & "; assessment basis: Existing retained"
        Else
            statusText = "CHECK - no usable outside-lane xfall in superelevation CSV"
        End If

        AddSuperelevationFinding alignmentName, profileName, startCh, endCh, False, 0#, _
                                 speedCode, requiredSuper, maxSuper, statusText, clauseText, candidateID
        Exit Sub
    End If

    If assessmentBasis = "Existing retained" Then
        clauseText = clauseText & "; assessment basis: Existing retained"

        If environment = "Rural" And achievedMax > maxSuper + 0.001 Then
            statusText = "CHECK - existing retained; CD 109 4.3 existing-road exception may apply"
        ElseIf achievedMin + 0.001 < requiredSuper Or _
               (environment = "Urban" And achievedMax > maxSuper + 0.001) Then
            statusText = "CHECK - existing retained xfall differs from new-design requirement"
        Else
            statusText = "COMPLIANT - existing retained xfall also meets new-design requirement"
        End If
    Else
        If achievedMax > maxSuper + 0.001 Then
            statusText = "DEPARTURE - exceeds CD 109 maximum superelevation"
        ElseIf achievedMin + 0.001 < requiredSuper Then
            statusText = "DEPARTURE - below required superelevation"
        Else
            statusText = "COMPLIANT"
        End If
    End If

    AddSuperelevationFinding alignmentName, profileName, startCh, endCh, True, achievedMin, _
                             speedCode, requiredSuper, maxSuper, statusText, clauseText, candidateID
End Sub

Private Sub AddSuperelevationFinding(ByVal alignmentName As String, ByVal profileName As String, _
                                     ByVal startCh As Double, ByVal endCh As Double, _
                                     ByVal hasXfall As Boolean, ByVal achievedXfall As Double, _
                                     ByVal speedCode As String, ByVal requiredSuper As Double, _
                                     ByVal maxSuper As Double, ByVal statusText As String, _
                                     ByVal clauseText As String, ByVal candidateID As String)
    Dim idx As Long

    If hasXfall Then
        AddFinding alignmentName, profileName, "Superelevation", "Superelevation", _
                   startCh, endCh, Abs(achievedXfall), "%", speedCode, requiredSuper, _
                   0, 0, 0, maxSuper, 0, False, statusText, clauseText, candidateID
    Else
        AddFinding alignmentName, profileName, "Superelevation", "Superelevation", _
                   startCh, endCh, 0#, "%", speedCode, requiredSuper, _
                   0, 0, 0, maxSuper, 0, False, statusText, clauseText, candidateID

        idx = FindingCount
        Findings(idx).HasAchieved = False
    End If
End Sub

Private Sub BuildSuperelevationTransitionFindingsFromCSV()
    Dim lo As ListObject
    Dim r As Long
    Dim j As Long
    Dim nextRow As Long
    Dim alignmentName As String
    Dim nextAlignment As String
    Dim sta As Variant
    Dim nextStaValue As Variant
    Dim staA As Double
    Dim staB As Double
    Dim bestNext As Double
    Dim leftA As Double
    Dim rightA As Double
    Dim leftB As Double
    Dim rightB As Double
    Dim hasLeftA As Boolean
    Dim hasRightA As Boolean
    Dim hasLeftB As Boolean
    Dim hasRightB As Boolean
    Dim baseSpeedCode As String
    Dim roadType As String
    Dim selectedProfile As String
    Dim includeText As String
    Dim setupStatus As String
    Dim profileStart As Double
    Dim profileEnd As Double
    Dim reviewStart As Double
    Dim reviewEnd As Double
    Dim environment As String
    Dim assessmentBasis As String
    Dim speedCode As String

    Set lo = GetTable(SH_SUPER, TBL_SUPER)

    For r = 1 To lo.ListRows.count
        If IsPreferredSuperelevationRow(lo, r) Then
            alignmentName = Trim$(CStr(lo.DataBodyRange.Cells(r, 3).Value2))
            sta = lo.DataBodyRange.Cells(r, 5).Value2

            If Len(alignmentName) > 0 And IsNumeric(sta) Then
                staA = CDbl(sta)
                nextRow = 0
                bestNext = 1E+30

                For j = 1 To lo.ListRows.count
                    If j <> r And IsPreferredSuperelevationRow(lo, j) Then
                        nextAlignment = Trim$(CStr(lo.DataBodyRange.Cells(j, 3).Value2))
                        nextStaValue = lo.DataBodyRange.Cells(j, 5).Value2

                        If nextAlignment = alignmentName And IsNumeric(nextStaValue) Then
                            If CDbl(nextStaValue) > staA + 0.001 And CDbl(nextStaValue) < bestNext Then
                                bestNext = CDbl(nextStaValue)
                                nextRow = j
                            End If
                        End If
                    End If
                Next j

                If nextRow > 0 Then
                    staB = bestNext
                    GetRowCrossfalls lo, r, leftA, hasLeftA, rightA, hasRightA
                    GetRowCrossfalls lo, nextRow, leftB, hasLeftB, rightB, hasRightB

                    If CrossfallRowsDiffer(leftA, hasLeftA, rightA, hasRightA, _
                                           leftB, hasLeftB, rightB, hasRightB) Then

                        If FindAlignmentSetting(alignmentName, baseSpeedCode, roadType, selectedProfile, includeText, setupStatus) Then
                            If includeText = "Yes" And setupStatus = "OK" Then
                                If GetSelectedProfileExtent(alignmentName, selectedProfile, profileStart, profileEnd) Then
                                    ' Only a critical station that actually falls inside the selected
                                    ' profile can complete a crossfall transition there. If the next
                                    ' point is beyond the profile end, the previous xfall is carried
                                    ' forward and no in-profile transition is created.
                                    If staB <= profileEnd + 0.000001 Then
                                        reviewStart = GreaterDouble(staA, profileStart)
                                        reviewEnd = LesserDouble(staB, profileEnd)

                                        If reviewEnd - reviewStart > 0.001 Then
                                            If GetSuperelevationSettingAt(alignmentName, (reviewStart + reviewEnd) / 2#, environment, assessmentBasis) Then
                                                speedCode = ResolveReviewSpeed(alignmentName, reviewStart, reviewEnd, baseSpeedCode)
                                                AddMinimumGradientFinding alignmentName, selectedProfile, speedCode, _
                                                                          reviewStart, reviewEnd, _
                                                                          "SE483CSV:" & CStr(r) & ":" & CStr(nextRow)
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next r
End Sub

Private Sub AddEdgeProfileFinding(ByVal alignmentName As String, ByVal profileName As String, _
                                  ByVal roadType As String, ByVal speedCode As String, _
                                  ByVal startCh As Double, ByVal endCh As Double, _
                                  ByVal fullSuper As Double, ByVal pivotToEdge As Double, _
                                  ByVal normalCrossfall As Double, ByVal candidateID As String)
    Dim transitionLength As Double
    Dim deltaCrossfall As Double
    Dim edgeVariation As Double
    Dim desirable As Double
    Dim permitted As Double
    Dim statusText As String
    Dim clauseText As String

    transitionLength = endCh - startCh
    If transitionLength <= 0# Then Exit Sub

    deltaCrossfall = Abs(normalCrossfall) + Abs(fullSuper)
    edgeVariation = pivotToEdge * deltaCrossfall / transitionLength

    permitted = 1#

    If roadType = "Motorway" Then
        desirable = 0.5
        clauseText = "CD 109 4.8.1 and 4.8.2"

        If edgeVariation <= 0.5 + 0.001 Then
            statusText = "COMPLIANT"
        ElseIf edgeVariation <= 1# + 0.001 Then
            statusText = "ADVISORY - above 0.5% motorway smoother-profile target"
        Else
            statusText = "ADVISORY - exceeds 1% edge-profile variation"
        End If
    Else
        desirable = 1#
        clauseText = "CD 109 4.8.1"

        If edgeVariation <= 1# + 0.001 Then
            statusText = "COMPLIANT"
        Else
            statusText = "ADVISORY - exceeds 1% edge-profile variation"
        End If
    End If

    clauseText = AppendSpeedInterfaceNote(clauseText, alignmentName, startCh, endCh)

    AddFinding alignmentName, profileName, "Superelevation", "SE edge grade variation", _
               startCh, endCh, edgeVariation, "%", speedCode, desirable, _
               0, 0, 0, permitted, 0, False, statusText, clauseText, candidateID
End Sub

Private Sub AddMinimumGradientFinding(ByVal alignmentName As String, ByVal profileName As String, _
                                      ByVal speedCode As String, ByVal startCh As Double, _
                                      ByVal endCh As Double, ByVal candidateID As String)
    Dim minGrade As Double
    Dim statusText As String
    Dim clauseText As String

    If Not MinimumAbsoluteProfileGrade(alignmentName, profileName, startCh, endCh, minGrade) Then Exit Sub

    If minGrade + 0.001 >= 0.5 Then
        statusText = "COMPLIANT"
    Else
        statusText = "ADVISORY - below 0.5% longitudinal gradient during superelevation"
    End If

    clauseText = AppendSpeedInterfaceNote("CD 109 4.8.3", alignmentName, startCh, endCh)

    AddFinding alignmentName, profileName, "Superelevation", "SE minimum longitudinal gradient", _
               startCh, endCh, minGrade, "%", speedCode, 0.5, _
               0, 0, 0, 0.5, 0, False, statusText, clauseText, candidateID
End Sub

Private Function FindSuperelevationEventForCurve(ByVal alignmentName As String, ByVal curveStart As Double, _
                                                 ByVal curveEnd As Double, ByRef hasFullSuper As Boolean, _
                                                 ByRef fullSuper As Double, ByRef rowIndex As Long) As Boolean
    Dim lo As ListObject
    Dim r As Long
    Dim eventStart As Variant
    Dim eventEnd As Variant
    Dim fullSta As Variant
    Dim runoffSta As Variant
    Dim fullValue As Variant
    Dim eventOverlap As Double
    Dim fullStart As Double
    Dim fullEnd As Double
    Dim fullOverlap As Double
    Dim bestFullOverlap As Double

    Set lo = GetTable(SH_SUPER, TBL_SUPER)
    hasFullSuper = False
    fullSuper = 0#
    bestFullOverlap = -1#

    For r = 1 To lo.ListRows.count
        If CStr(lo.DataBodyRange.Cells(r, 1).Value2) = alignmentName Then
            eventStart = lo.DataBodyRange.Cells(r, 2).Value2
            eventEnd = lo.DataBodyRange.Cells(r, 3).Value2
            fullSta = lo.DataBodyRange.Cells(r, 6).Value2
            fullValue = lo.DataBodyRange.Cells(r, 7).Value2
            runoffSta = lo.DataBodyRange.Cells(r, 8).Value2

            If IsNumeric(eventStart) And IsNumeric(eventEnd) Then
                eventOverlap = LesserDouble(curveEnd, CDbl(eventEnd)) - GreaterDouble(curveStart, CDbl(eventStart))

                If eventOverlap > 0.001 Then
                    FindSuperelevationEventForCurve = True

                    If IsNumeric(fullValue) And IsNumeric(fullSta) Then
                        fullStart = CDbl(fullSta)

                        If IsNumeric(runoffSta) Then
                            fullEnd = CDbl(runoffSta)
                        Else
                            fullEnd = CDbl(eventEnd)
                        End If

                        If fullEnd < fullStart Then SwapDoubleValues fullStart, fullEnd

                        fullOverlap = LesserDouble(curveEnd, fullEnd) - GreaterDouble(curveStart, fullStart)

                        If fullOverlap > 0.001 And fullOverlap > bestFullOverlap Then
                            bestFullOverlap = fullOverlap
                            rowIndex = r
                            fullSuper = CDbl(fullValue)
                            hasFullSuper = True
                        End If
                    End If
                End If
            End If
        End If
    Next r
End Function

Private Function GetMaximumSuperelevation(ByVal environment As String) As Double
    If environment = "Urban" Then
        GetMaximumSuperelevation = 5#
    ElseIf environment = "Rural" Then
        GetMaximumSuperelevation = 7#
    End If
End Function

Private Function GetSuperelevationSettingAt(ByVal alignmentName As String, ByVal chainage As Double, _
                                            ByRef environment As String, ByRef assessmentBasis As String) As Boolean
    Dim lo As ListObject
    Dim r As Long
    Dim s As Variant
    Dim e As Variant
    Dim activeText As String

    Set lo = GetTable(SH_SUPER_SETUP, TBL_SUPER_SETTINGS)
    environment = vbNullString
    assessmentBasis = vbNullString

    For r = 1 To lo.ListRows.count
        If Trim$(CStr(lo.DataBodyRange.Cells(r, 2).Value2)) = alignmentName Then
            activeText = Trim$(CStr(lo.DataBodyRange.Cells(r, 6).Value2))
            s = lo.DataBodyRange.Cells(r, 3).Value2
            e = lo.DataBodyRange.Cells(r, 4).Value2

            If activeText = "Yes" And IsNumeric(s) And IsNumeric(e) Then
                If chainage >= LesserDouble(CDbl(s), CDbl(e)) - 0.001 And _
                   chainage <= GreaterDouble(CDbl(s), CDbl(e)) + 0.001 Then
                    environment = Trim$(CStr(lo.DataBodyRange.Cells(r, 5).Value2))
                    assessmentBasis = Trim$(CStr(lo.DataBodyRange.Cells(r, 9).Value2))
                    GetSuperelevationSettingAt = True
                    Exit Function
                End If
            End If
        End If
    Next r
End Function

Private Function GetEntrySuperelevationInterval(ByVal alignmentName As String, ByVal fullSta As Variant, _
                                                ByVal beginRunout As Variant, ByVal beginRunoff As Variant, _
                                                ByRef startCh As Double, ByRef endCh As Double) As Boolean
    Dim candidate As Double
    Dim hasCandidate As Boolean
    Dim spiralStart As Double

    If Not IsNumeric(fullSta) Then Exit Function
    endCh = CDbl(fullSta)

    If IsNumeric(beginRunout) Then
        If CDbl(beginRunout) < endCh - 0.001 Then
            candidate = CDbl(beginRunout)
            hasCandidate = True
        End If
    End If

    If IsNumeric(beginRunoff) Then
        If CDbl(beginRunoff) < endCh - 0.001 Then
            If Not hasCandidate Or CDbl(beginRunoff) < candidate Then candidate = CDbl(beginRunoff)
            hasCandidate = True
        End If
    End If

    If PreviousSpiralStart(alignmentName, endCh, spiralStart) Then
        If spiralStart < endCh - 0.001 Then
            If Not hasCandidate Or spiralStart < candidate Then candidate = spiralStart
            hasCandidate = True
        End If
    End If

    If hasCandidate Then
        startCh = candidate
        GetEntrySuperelevationInterval = True
    End If
End Function

Private Function GetExitSuperelevationInterval(ByVal alignmentName As String, ByVal runoffSta As Variant, _
                                               ByVal startRunout As Variant, ByVal endRunout As Variant, _
                                               ByRef startCh As Double, ByRef endCh As Double) As Boolean
    Dim candidate As Double
    Dim hasCandidate As Boolean
    Dim spiralEnd As Double

    If Not IsNumeric(runoffSta) Then Exit Function
    startCh = CDbl(runoffSta)

    If IsNumeric(startRunout) Then
        If CDbl(startRunout) > startCh + 0.001 Then
            candidate = CDbl(startRunout)
            hasCandidate = True
        End If
    End If

    If IsNumeric(endRunout) Then
        If CDbl(endRunout) > startCh + 0.001 Then
            If Not hasCandidate Or CDbl(endRunout) > candidate Then candidate = CDbl(endRunout)
            hasCandidate = True
        End If
    End If

    If FollowingSpiralEnd(alignmentName, startCh, spiralEnd) Then
        If spiralEnd > startCh + 0.001 Then
            If Not hasCandidate Or spiralEnd > candidate Then candidate = spiralEnd
            hasCandidate = True
        End If
    End If

    If hasCandidate Then
        endCh = candidate
        GetExitSuperelevationInterval = True
    End If
End Function

Private Function PreviousSpiralStart(ByVal alignmentName As String, ByVal endAt As Double, ByRef spiralStart As Double) As Boolean
    Dim lo As ListObject
    Dim r As Long

    Set lo = GetTable(SH_HORIZ, TBL_HORIZ)

    For r = 1 To lo.ListRows.count
        If CStr(lo.DataBodyRange.Cells(r, 2).Value2) = alignmentName And _
           CStr(lo.DataBodyRange.Cells(r, 3).Value2) = "Spiral" And _
           IsNumeric(lo.DataBodyRange.Cells(r, 5).Value2) Then

            If Abs(CDbl(lo.DataBodyRange.Cells(r, 5).Value2) - endAt) <= 0.002 Then
                If IsNumeric(lo.DataBodyRange.Cells(r, 4).Value2) Then
                    spiralStart = CDbl(lo.DataBodyRange.Cells(r, 4).Value2)
                    PreviousSpiralStart = True
                    Exit Function
                End If
            End If
        End If
    Next r
End Function

Private Function FollowingSpiralEnd(ByVal alignmentName As String, ByVal startAt As Double, ByRef spiralEnd As Double) As Boolean
    Dim lo As ListObject
    Dim r As Long

    Set lo = GetTable(SH_HORIZ, TBL_HORIZ)

    For r = 1 To lo.ListRows.count
        If CStr(lo.DataBodyRange.Cells(r, 2).Value2) = alignmentName And _
           CStr(lo.DataBodyRange.Cells(r, 3).Value2) = "Spiral" And _
           IsNumeric(lo.DataBodyRange.Cells(r, 4).Value2) Then

            If Abs(CDbl(lo.DataBodyRange.Cells(r, 4).Value2) - startAt) <= 0.002 Then
                If IsNumeric(lo.DataBodyRange.Cells(r, 5).Value2) Then
                    spiralEnd = CDbl(lo.DataBodyRange.Cells(r, 5).Value2)
                    FollowingSpiralEnd = True
                    Exit Function
                End If
            End If
        End If
    Next r
End Function

Private Function MinimumAbsoluteProfileGrade(ByVal alignmentName As String, ByVal profileName As String, _
                                             ByVal startCh As Double, ByVal endCh As Double, _
                                             ByRef minGrade As Double) As Boolean
    Dim lo As ListObject
    Dim r As Long
    Dim vStart As Variant
    Dim vEnd As Variant
    Dim gIn As Variant
    Dim gOut As Variant
    Dim ovStart As Double
    Dim ovEnd As Double
    Dim g1 As Double
    Dim g2 As Double
    Dim thisMin As Double
    Dim found As Boolean
    Dim vertType As String

    Set lo = GetTable(SH_VERT, TBL_VERT)
    minGrade = 1E+99

    For r = 1 To lo.ListRows.count
        If CStr(lo.DataBodyRange.Cells(r, 2).Value2) = alignmentName And _
           CStr(lo.DataBodyRange.Cells(r, 3).Value2) = profileName Then

            vStart = lo.DataBodyRange.Cells(r, 5).Value2
            vEnd = lo.DataBodyRange.Cells(r, 6).Value2
            gIn = lo.DataBodyRange.Cells(r, 8).Value2
            gOut = lo.DataBodyRange.Cells(r, 9).Value2
            vertType = CStr(lo.DataBodyRange.Cells(r, 4).Value2)

            If IsNumeric(vStart) And IsNumeric(vEnd) And IsNumeric(gIn) Then
                ovStart = GreaterDouble(startCh, CDbl(vStart))
                ovEnd = LesserDouble(endCh, CDbl(vEnd))

                If ovEnd - ovStart > 0.000001 Then
                    If vertType = "Tangent" Or Not IsNumeric(gOut) Or Abs(CDbl(vEnd) - CDbl(vStart)) < 0.000001 Then
                        thisMin = Abs(CDbl(gIn))
                    Else
                        g1 = CDbl(gIn) + (CDbl(gOut) - CDbl(gIn)) * ((ovStart - CDbl(vStart)) / (CDbl(vEnd) - CDbl(vStart)))
                        g2 = CDbl(gIn) + (CDbl(gOut) - CDbl(gIn)) * ((ovEnd - CDbl(vStart)) / (CDbl(vEnd) - CDbl(vStart)))

                        If g1 = 0# Or g2 = 0# Or Sgn(g1) <> Sgn(g2) Then
                            thisMin = 0#
                        Else
                            thisMin = LesserDouble(Abs(g1), Abs(g2))
                        End If
                    End If

                    If thisMin < minGrade Then minGrade = thisMin
                    found = True
                End If
            End If
        End If
    Next r

    MinimumAbsoluteProfileGrade = found
End Function

Private Function GreaterDouble(ByVal a As Double, ByVal b As Double) As Double
    If a >= b Then
        GreaterDouble = a
    Else
        GreaterDouble = b
    End If
End Function

Private Function LesserDouble(ByVal a As Double, ByVal b As Double) As Double
    If a <= b Then
        LesserDouble = a
    Else
        LesserDouble = b
    End If
End Function

Private Sub BuildManualSSDFindings()
    Dim lo As ListObject
    Dim r As Long
    Dim ssdID As String
    Dim alignmentName As String
    Dim speedCode As String
    Dim roadType As String
    Dim profileName As String
    Dim includeText As String
    Dim setupStatus As String
    Dim startCh As Variant
    Dim endCh As Variant
    Dim achieved As Variant
    Dim activeText As String

    Dim desirable As Double
    Dim baseSteps As Long
    Dim contextAdj As Long
    Dim permittedSteps As Long
    Dim permittedLimit As Double
    Dim achievedSteps As Long
    Dim statusText As String
    Dim clauseText As String

    Set lo = GetTable(SH_SSD_REVIEW, TBL_SSD)

    For r = 1 To lo.ListRows.count
        ssdID = Trim$(CStr(lo.DataBodyRange.Cells(r, 1).Value2))
        alignmentName = Trim$(CStr(lo.DataBodyRange.Cells(r, 2).Value2))
        startCh = lo.DataBodyRange.Cells(r, 4).Value2
        endCh = lo.DataBodyRange.Cells(r, 5).Value2
        achieved = lo.DataBodyRange.Cells(r, 6).Value2
        activeText = Trim$(CStr(lo.DataBodyRange.Cells(r, 7).Value2))

        If Len(ssdID) > 0 And activeText = "Yes" Then
            If FindAlignmentSetting(alignmentName, speedCode, roadType, profileName, includeText, setupStatus) Then
                If includeText = "Yes" And setupStatus = "OK" And IsNumeric(startCh) And IsNumeric(endCh) And IsNumeric(achieved) Then
                    speedCode = ResolveReviewSpeed(alignmentName, CDbl(startCh), CDbl(endCh), speedCode)

                    EvaluateStepParameter "Stopping sight distance", CDbl(achieved), alignmentName, CDbl(startCh), CDbl(endCh), _
                                          speedCode, roadType, desirable, baseSteps, contextAdj, permittedSteps, permittedLimit, _
                                          achievedSteps, statusText, clauseText

                    clauseText = AppendSpeedInterfaceNote(clauseText, alignmentName, CDbl(startCh), CDbl(endCh))

                    AddFinding alignmentName, vbNullString, "Manual SSD", "Stopping sight distance", _
                               CDbl(startCh), CDbl(endCh), CDbl(achieved), "m", speedCode, _
                               desirable, baseSteps, contextAdj, permittedSteps, permittedLimit, _
                               achievedSteps, True, statusText, clauseText, "SSD:" & ssdID
                End If
            End If
        End If
    Next r
End Sub

Private Sub EvaluateStepParameter(ByVal parameterName As String, ByVal achieved As Double, _
                                  ByVal alignmentName As String, ByVal startCh As Double, ByVal endCh As Double, _
                                  ByVal speedCode As String, ByVal roadType As String, _
                                  ByRef desirable As Double, ByRef baseSteps As Long, ByRef contextAdj As Long, _
                                  ByRef permittedSteps As Long, ByRef permittedLimit As Double, _
                                  ByRef achievedSteps As Long, ByRef statusText As String, ByRef clauseText As String)

    Dim speedValue As Long
    Dim requestedSteps As Long

    speedValue = GetNumericSpeed(speedCode)
    desirable = GetDesignValue(speedValue, parameterName)
    baseSteps = GetBaseSteps(parameterName, roadType, speedCode)
    contextAdj = GetContextAdjustment(alignmentName, parameterName, startCh, endCh)

    requestedSteps = baseSteps + contextAdj
    If requestedSteps < 0 Then requestedSteps = 0

    permittedSteps = requestedSteps
    If permittedSteps > StepsTo50(speedValue) Then permittedSteps = StepsTo50(speedValue)

    permittedLimit = GetDesignValue(SpeedAfterSteps(speedValue, permittedSteps), parameterName)
    achievedSteps = GetAchievedSteps(achieved, speedValue, parameterName)

    clauseText = BaseClauseForParameter(parameterName)
    If contextAdj <> 0 Then clauseText = clauseText & "; context " & GetContextClauses(alignmentName, parameterName, startCh, endCh)

    If achievedSteps = -1 Then
        statusText = "DEPARTURE - CD 109 2.11 floor"
    ElseIf achievedSteps = 0 Then
        statusText = "COMPLIANT"
    ElseIf achievedSteps <= permittedSteps Then
        statusText = "RELAXATION - permitted"
    Else
        statusText = "DEPARTURE - exceeds permitted relaxation"
    End If
End Sub

Private Sub AddFinding(ByVal alignmentName As String, ByVal profileName As String, ByVal sourceText As String, _
                       ByVal parameterName As String, ByVal startCh As Double, ByVal endCh As Double, _
                       ByVal achieved As Double, ByVal unitText As String, ByVal speedCode As String, _
                       ByVal desirable As Double, ByVal baseSteps As Long, ByVal contextAdj As Long, _
                       ByVal permittedSteps As Long, ByVal permittedLimit As Double, _
                       ByVal achievedSteps As Long, ByVal hasStepAssessment As Boolean, _
                       ByVal statusText As String, ByVal clauseText As String, ByVal candidateID As String)

    FindingCount = FindingCount + 1
    ReDim Preserve Findings(1 To FindingCount)

    With Findings(FindingCount)
        .FindingRef = "AUTO-" & Format$(FindingCount, "0000")
        .Alignment = alignmentName
        .Profile = profileName
        .Source = sourceText
        .Parameter = parameterName
        .startCh = startCh
        .endCh = endCh
        .HasRange = True
        .achieved = achieved
        .HasAchieved = True
        .unitText = unitText
        .speedCode = speedCode
        .desirable = desirable
        .HasDesirable = True
        .baseSteps = baseSteps
        .contextAdj = contextAdj
        .permittedSteps = permittedSteps
        .permittedLimit = permittedLimit
        .HasPermittedLimit = True
        .achievedSteps = achievedSteps
        .hasStepAssessment = hasStepAssessment
        .ReviewStatus = statusText
        .ClauseBasis = clauseText
        .candidateID = candidateID
    End With
End Sub

' ============================================================
' OUTPUT TABLES
' ============================================================

Private Function SaveReviewerNotes() As Object
    Dim d As Object
    Dim lo As ListObject
    Dim r As Long
    Dim candidateID As String
    Dim noteText As String

    Set d = CreateObject("Scripting.Dictionary")
    Set lo = GetTable(SH_REVIEW, TBL_REVIEW)

    If lo.DataBodyRange Is Nothing Then
        Set SaveReviewerNotes = d
        Exit Function
    End If

    For r = 1 To lo.ListRows.count
        candidateID = CStr(lo.DataBodyRange.Cells(r, 19).Value2)
        noteText = CStr(lo.DataBodyRange.Cells(r, 20).Value2)

        If Len(candidateID) > 0 And Len(noteText) > 0 Then d.Item(candidateID) = noteText
    Next r

    Set SaveReviewerNotes = d
End Function

Private Sub WriteGeometryReview(ByVal notes As Object)
    Dim lo As ListObject
    Dim r As Long

    Set lo = GetTable(SH_REVIEW, TBL_REVIEW)
    ClearOutputArea lo
    ResizeTableRows lo, IIf(FindingCount > 0, FindingCount, 1)

    If FindingCount = 0 Then
        lo.DataBodyRange.Cells(1, 1).Value = "No findings - complete Review Setup and click Update Reviews"
        Exit Sub
    End If

    For r = 1 To FindingCount
        With lo.DataBodyRange.Rows(r)
            .Cells(1, 1).Value = Findings(r).FindingRef
            .Cells(1, 2).Value = Findings(r).Alignment
            .Cells(1, 3).Value = Findings(r).Profile
            .Cells(1, 4).Value = Findings(r).Source
            .Cells(1, 5).Value = Findings(r).Parameter
            .Cells(1, 6).Value = Findings(r).startCh
            .Cells(1, 7).Value = Findings(r).endCh
            If Findings(r).HasAchieved Then
                .Cells(1, 8).Value = Findings(r).achieved
            Else
                .Cells(1, 8).ClearContents
            End If
            .Cells(1, 9).Value = Findings(r).unitText
            .Cells(1, 10).Value = Findings(r).speedCode
            .Cells(1, 11).Value = Findings(r).desirable

            If Findings(r).hasStepAssessment Then
                .Cells(1, 12).Value = Findings(r).baseSteps
                .Cells(1, 13).Value = Findings(r).contextAdj
                .Cells(1, 14).Value = Findings(r).permittedSteps
                .Cells(1, 15).Value = Findings(r).permittedLimit

                If Findings(r).achievedSteps = -1 Then
                    .Cells(1, 16).Value = "BELOW FLOOR"
                Else
                    .Cells(1, 16).Value = Findings(r).achievedSteps
                End If
            Else
                .Cells(1, 15).Value = Findings(r).permittedLimit
            End If

            .Cells(1, 17).Value = Findings(r).ReviewStatus
            .Cells(1, 18).Value = Findings(r).ClauseBasis
            .Cells(1, 19).Value = Findings(r).candidateID

            If notes.Exists(Findings(r).candidateID) Then .Cells(1, 20).Value = notes.Item(Findings(r).candidateID)
        End With
    Next r
End Sub

Private Sub WriteCoincidenceReview()
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim i As Long
    Dim j As Long
    Dim pairCount As Long
    Dim outRow As Long
    Dim sheetRow As Long
    Dim dataRows As Long
    Dim clearLastRow As Long
    Dim overlapStart As Double
    Dim overlapEnd As Double
    Dim resultText As String
    Dim localStage As String
    Dim errNo As Long
    Dim errDesc As String

    On Error GoTo Failed

    localStage = "getting Coincidence Review worksheet"
    Set ws = ThisWorkbook.Worksheets(SH_COINC)

    localStage = "getting Coincidence Review table"
    Set lo = ws.ListObjects(TBL_COINC)

    localStage = "counting coincidence pairs"
    pairCount = 0
    For i = 1 To FindingCount - 1
        If IsRelaxedFinding(Findings(i)) Then
            For j = i + 1 To FindingCount
                If IsRelaxedFinding(Findings(j)) Then
                    If Findings(i).Alignment = Findings(j).Alignment And _
                       Findings(i).Parameter <> Findings(j).Parameter Then

                        overlapStart = GreaterOf(Findings(i).startCh, Findings(j).startCh)
                        overlapEnd = LesserOf(Findings(i).endCh, Findings(j).endCh)

                        ' Positive-length overlap only. Boundary/point contact is not coincidence.
                        If overlapEnd - overlapStart > 0.001 Then
                            pairCount = pairCount + 1
                        End If
                    End If
                End If
            Next j
        End If
    Next i

    localStage = "clearing old Coincidence Review values"
    clearLastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    If clearLastRow < 5 Then clearLastRow = 5
    ws.Range("A5:U" & CStr(clearLastRow)).ClearContents

    localStage = "resizing Coincidence Review table"
    If pairCount > 0 Then
        dataRows = pairCount
    Else
        dataRows = 1
    End If
    lo.Resize ws.Range("A4:U" & CStr(4 + dataRows))

    If pairCount = 0 Then
        localStage = "writing no-coincidence message"
        ws.Cells(5, 1).Value = _
            "No coincidence pairs - click Update Reviews after completing Review Setup"
        Exit Sub
    End If

    localStage = "writing coincidence pairs"
    outRow = 0
    For i = 1 To FindingCount - 1
        If IsRelaxedFinding(Findings(i)) Then
            For j = i + 1 To FindingCount
                If IsRelaxedFinding(Findings(j)) Then
                    If Findings(i).Alignment = Findings(j).Alignment And _
                       Findings(i).Parameter <> Findings(j).Parameter Then

                        overlapStart = GreaterOf(Findings(i).startCh, Findings(j).startCh)
                        overlapEnd = LesserOf(Findings(i).endCh, Findings(j).endCh)

                        If overlapEnd - overlapStart > 0.001 Then
                            outRow = outRow + 1
                            sheetRow = 4 + outRow
                            resultText = CombinationResult(Findings(i), Findings(j))

                            localStage = "writing Coincidence Review row " & CStr(outRow)
                            WriteCoincidenceWorksheetRow ws, sheetRow, outRow, _
                                                        Findings(i), Findings(j), _
                                                        overlapStart, overlapEnd, resultText
                        End If
                    End If
                End If
            Next j
        End If
    Next i

    Exit Sub

Failed:
    errNo = Err.Number
    errDesc = Err.description
    Err.Raise vbObjectError + 644, "WriteCoincidenceReview", _
              localStage & " | original error " & CStr(errNo) & ": " & errDesc
End Sub
Private Function IsPermittedRelaxation(ByRef f As FindingRec) As Boolean
    If f.hasStepAssessment Then
        IsPermittedRelaxation = (f.achievedSteps > 0 And f.achievedSteps <= f.permittedSteps)
    ElseIf f.Parameter = "Longitudinal gradient" Then
        IsPermittedRelaxation = (Not LEQ3(f.achieved, f.desirable) And LEQ3(f.achieved, f.permittedLimit))
    End If
End Function

Private Function IsRelaxedFinding(ByRef f As FindingRec) As Boolean
    If f.hasStepAssessment Then
        IsRelaxedFinding = (f.achievedSteps <> 0)
    ElseIf f.Parameter = "Longitudinal gradient" Then
        IsRelaxedFinding = Not LEQ3(f.achieved, f.desirable)
    End If
End Function


Private Function CombinationResult(ByRef a As FindingRec, ByRef b As FindingRec) As String
    Dim aWithin As Boolean
    Dim bWithin As Boolean
    Dim pairExcepted As Boolean

    aWithin = IsPermittedRelaxation(a)
    bWithin = IsPermittedRelaxation(b)

    pairExcepted = False
    If IsSSDHorizontalPair(a.Parameter, b.Parameter) Then
        If a.hasStepAssessment And b.hasStepAssessment Then
            If a.achievedSteps > 0 And b.achievedSteps > 0 And _
               a.achievedSteps <= 1 And b.achievedSteps <= 1 Then
                pairExcepted = True
            End If
        End If
    End If

    If pairExcepted Then
        If aWithin And bWithin Then
            CombinationResult = "PERMITTED - 2.12 exception"
        Else
            CombinationResult = "DEPARTURE - individual permitted relaxation exceeded"
        End If
        Exit Function
    End If

    If Not aWithin Or Not bWithin Then
        CombinationResult = "DEPARTURE - individual limit exceeded; 2.12 combination not permitted"
    Else
        CombinationResult = "DEPARTURE - CD 109 2.12"
    End If
End Function


Private Function IsSSDHorizontalPair(ByVal pA As String, ByVal pB As String) As Boolean
    IsSSDHorizontalPair = _
        ((pA = "Stopping sight distance" And pB = "Horizontal curvature") Or _
         (pA = "Horizontal curvature" And pB = "Stopping sight distance"))
End Function

Private Sub WriteCoincidenceWorksheetRow(ByVal ws As Worksheet, ByVal sheetRow As Long, _
                                         ByVal outRow As Long, ByRef a As FindingRec, ByRef b As FindingRec, _
                                         ByVal overlapStart As Double, ByVal overlapEnd As Double, _
                                         ByVal resultText As String)

    With ws
        .Cells(sheetRow, 1).Value = "PAIR-" & Format$(outRow, "000")
        .Cells(sheetRow, 2).Value = a.Alignment
        .Cells(sheetRow, 3).Value = a.FindingRef
        .Cells(sheetRow, 4).Value = a.Parameter
        .Cells(sheetRow, 5).Value = CompactValue(a.Parameter, a.achieved)
        .Cells(sheetRow, 6).Value = CompactValue(a.Parameter, a.desirable)
        .Cells(sheetRow, 7).Value = CompactValue(a.Parameter, a.permittedLimit)
        If a.hasStepAssessment Then .Cells(sheetRow, 8).Value = a.achievedSteps
        .Cells(sheetRow, 9).Value = a.ClauseBasis

        .Cells(sheetRow, 10).Value = b.FindingRef
        .Cells(sheetRow, 11).Value = b.Parameter
        .Cells(sheetRow, 12).Value = CompactValue(b.Parameter, b.achieved)
        .Cells(sheetRow, 13).Value = CompactValue(b.Parameter, b.desirable)
        .Cells(sheetRow, 14).Value = CompactValue(b.Parameter, b.permittedLimit)
        If b.hasStepAssessment Then .Cells(sheetRow, 15).Value = b.achievedSteps
        .Cells(sheetRow, 16).Value = b.ClauseBasis

        .Cells(sheetRow, 17).Value = overlapStart
        .Cells(sheetRow, 18).Value = overlapEnd
        .Cells(sheetRow, 19).Value = resultText
        .Cells(sheetRow, 20).Value = "CD 109 2.12"
    End With
End Sub

Private Function GreaterOf(ByVal a As Double, ByVal b As Double) As Double
    If a >= b Then
        GreaterOf = a
    Else
        GreaterOf = b
    End If
End Function

Private Function LesserOf(ByVal a As Double, ByVal b As Double) As Double
    If a <= b Then
        LesserOf = a
    Else
        LesserOf = b
    End If
End Function
Private Sub WriteDepartureSummary()
    Dim ws As Worksheet
    Dim wsCoin As Worksheet
    Dim lo As ListObject
    Dim coin As ListObject
    Dim r As Long
    Dim z As Long
    Dim outRow As Long
    Dim totalRows As Long
    Dim sheetRow As Long
    Dim coinSheetRow As Long
    Dim clearLastRow As Long
    Dim overlapStart As Double
    Dim overlapEnd As Double
    Dim statusText As String

    Set ws = ThisWorkbook.Worksheets(SH_DEPART)
    Set wsCoin = ThisWorkbook.Worksheets(SH_COINC)
    Set lo = ws.ListObjects(TBL_DEPART)
    Set coin = wsCoin.ListObjects(TBL_COINC)

    ' Count first so the table is resized once.
    totalRows = 0

    For r = 1 To FindingCount
        If Left$(Findings(r).ReviewStatus, 9) = "DEPARTURE" Then
            totalRows = totalRows + 1
        End If
    Next r

    For r = 1 To FindingCount
        If Is213Parameter(Findings(r).Parameter) And IsBelowDesirable(Findings(r)) Then
            For z = 1 To zoneCount
                If Zones(z).Alignment = Findings(r).Alignment And _
                   Zones(z).Parameter = Findings(r).Parameter Then

                    overlapStart = GreaterOf(Findings(r).startCh, Zones(z).startCh)
                    overlapEnd = LesserOf(Findings(r).endCh, Zones(z).endCh)

                    If overlapEnd - overlapStart > 0.000000001 Then
                        totalRows = totalRows + 1
                    End If
                End If
            Next z
        End If
    Next r

    For r = 1 To coin.ListRows.count
        coinSheetRow = 4 + r
        statusText = CStr(wsCoin.Cells(coinSheetRow, 19).Value2)
        If Left$(statusText, 9) = "DEPARTURE" Then
            totalRows = totalRows + 1
        End If
    Next r

    clearLastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    If clearLastRow < 5 Then clearLastRow = 5
    ws.Range("A5:P" & CStr(clearLastRow)).ClearContents

    If totalRows > 0 Then
        lo.Resize ws.Range("A4:P" & CStr(4 + totalRows))
    Else
        lo.Resize ws.Range("A4:P5")
        ws.Cells(5, 1).Value = _
            "No departures - click Update Reviews after completing Review Setup"
        Exit Sub
    End If

    outRow = 0

    For r = 1 To FindingCount
        If Left$(Findings(r).ReviewStatus, 9) = "DEPARTURE" Then
            outRow = outRow + 1
            sheetRow = 4 + outRow
            WriteIndividualDepartureWs ws, sheetRow, outRow, Findings(r)
        End If
    Next r

    ' CD 109 2.13 applies to SSD / crest curvature / sag curvature only.
    For r = 1 To FindingCount
        If Is213Parameter(Findings(r).Parameter) And IsBelowDesirable(Findings(r)) Then
            For z = 1 To zoneCount
                If Zones(z).Alignment = Findings(r).Alignment And _
                   Zones(z).Parameter = Findings(r).Parameter Then

                    overlapStart = GreaterOf(Findings(r).startCh, Zones(z).startCh)
                    overlapEnd = LesserOf(Findings(r).endCh, Zones(z).endCh)

                    If overlapEnd - overlapStart > 0.000000001 Then
                        outRow = outRow + 1
                        sheetRow = 4 + outRow
                        Write213DepartureWs ws, sheetRow, outRow, Findings(r), Zones(z), _
                                            overlapStart, overlapEnd
                    End If
                End If
            Next z
        End If
    Next r

    For r = 1 To coin.ListRows.count
        coinSheetRow = 4 + r
        statusText = CStr(wsCoin.Cells(coinSheetRow, 19).Value2)
        If Left$(statusText, 9) = "DEPARTURE" Then
            outRow = outRow + 1
            sheetRow = 4 + outRow
            Write212DepartureWs ws, sheetRow, outRow, wsCoin, coinSheetRow
        End If
    Next r
End Sub

Private Function Is213Parameter(ByVal p As String) As Boolean
    Is213Parameter = (p = "Stopping sight distance" Or p = "Crest curvature" Or p = "Sag curvature")
End Function

Private Function IsBelowDesirable(ByRef f As FindingRec) As Boolean
    If f.hasStepAssessment Then
        IsBelowDesirable = (f.achievedSteps <> 0)
    Else
        IsBelowDesirable = False
    End If
End Function

Private Sub WriteIndividualDepartureWs(ByVal ws As Worksheet, ByVal sheetRow As Long, _
                                       ByVal outRow As Long, ByRef f As FindingRec)
    Dim clauseText As String

    clauseText = f.ClauseBasis
    If InStr(1, f.ReviewStatus, "2.11", vbTextCompare) > 0 Then
        clauseText = "CD 109 2.11; " & clauseText
    End If

    With ws
        .Cells(sheetRow, 1).Value = "DEP-" & Format$(outRow, "000")
        .Cells(sheetRow, 2).Value = f.FindingRef
        .Cells(sheetRow, 4).Value = f.Alignment
        .Cells(sheetRow, 5).Value = "Individual limit"
        .Cells(sheetRow, 6).Value = f.Parameter
        .Cells(sheetRow, 7).Value = f.startCh
        .Cells(sheetRow, 8).Value = f.endCh
        .Cells(sheetRow, 9).Value = CompactValue(f.Parameter, f.achieved)
        .Cells(sheetRow, 10).Value = CompactValue(f.Parameter, f.desirable)
        .Cells(sheetRow, 11).Value = CompactValue(f.Parameter, f.permittedLimit)
        .Cells(sheetRow, 13).Value = f.ReviewStatus
        .Cells(sheetRow, 14).Value = clauseText
    End With
End Sub

Private Sub Write213DepartureWs(ByVal ws As Worksheet, ByVal sheetRow As Long, _
                                ByVal outRow As Long, ByRef f As FindingRec, ByRef z As ZoneRec, _
                                ByVal overlapStart As Double, ByVal overlapEnd As Double)
    Dim zoneText As String

    zoneText = "Zone Ch " & Format$(z.startCh, "0.000") & "-" & _
               Format$(z.endCh, "0.000") & "; 1.5 x SSD = " & _
               Format$(z.ssd15, "0.###") & " m"

    With ws
        .Cells(sheetRow, 1).Value = "DEP-" & Format$(outRow, "000")
        .Cells(sheetRow, 2).Value = f.FindingRef
        .Cells(sheetRow, 4).Value = f.Alignment
        .Cells(sheetRow, 5).Value = "Junction restriction"
        .Cells(sheetRow, 6).Value = f.Parameter
        .Cells(sheetRow, 7).Value = overlapStart
        .Cells(sheetRow, 8).Value = overlapEnd
        .Cells(sheetRow, 9).Value = CompactValue(f.Parameter, f.achieved)
        .Cells(sheetRow, 10).Value = CompactValue(f.Parameter, f.desirable)
        .Cells(sheetRow, 11).Value = "No relaxation permitted in immediate approach"
        .Cells(sheetRow, 13).Value = _
            "Below-desirable value lies within a clause 2.13 immediate approach"
        .Cells(sheetRow, 14).Value = "CD 109 2.13"
        .Cells(sheetRow, 15).Value = zoneText
        .Cells(sheetRow, 16).Value = z.junctionID & "; " & z.ClauseBasis
    End With
End Sub

Private Sub Write212DepartureWs(ByVal ws As Worksheet, ByVal sheetRow As Long, _
                                ByVal outRow As Long, ByVal wsCoin As Worksheet, _
                                ByVal coinSheetRow As Long)
    Dim overlapText As String

    overlapText = "Ch " & Format$(CDbl(wsCoin.Cells(coinSheetRow, 17).Value2), "0.000") & _
                  "-" & Format$(CDbl(wsCoin.Cells(coinSheetRow, 18).Value2), "0.000")

    With ws
        .Cells(sheetRow, 1).Value = "DEP-" & Format$(outRow, "000")
        .Cells(sheetRow, 2).Value = wsCoin.Cells(coinSheetRow, 3).Value2
        .Cells(sheetRow, 3).Value = wsCoin.Cells(coinSheetRow, 10).Value2
        .Cells(sheetRow, 4).Value = wsCoin.Cells(coinSheetRow, 2).Value2
        .Cells(sheetRow, 5).Value = "Coincident relaxations"
        .Cells(sheetRow, 6).Value = CStr(wsCoin.Cells(coinSheetRow, 4).Value2) & _
                                    " + " & CStr(wsCoin.Cells(coinSheetRow, 11).Value2)
        .Cells(sheetRow, 7).Value = wsCoin.Cells(coinSheetRow, 17).Value2
        .Cells(sheetRow, 8).Value = wsCoin.Cells(coinSheetRow, 18).Value2
        .Cells(sheetRow, 9).Value = wsCoin.Cells(coinSheetRow, 5).Value2
        .Cells(sheetRow, 10).Value = wsCoin.Cells(coinSheetRow, 6).Value2
        .Cells(sheetRow, 11).Value = _
            "Only SSD <=1 step + horizontal <=1 step is excepted"
        .Cells(sheetRow, 12).Value = wsCoin.Cells(coinSheetRow, 12).Value2
        .Cells(sheetRow, 13).Value = _
            "Relaxations overlap outside the clause 2.12 exception"
        .Cells(sheetRow, 14).Value = "CD 109 2.12"
        .Cells(sheetRow, 15).Value = overlapText
        .Cells(sheetRow, 16).Value = _
            "A " & wsCoin.Cells(coinSheetRow, 8).Value2 & _
            " step(s); B " & wsCoin.Cells(coinSheetRow, 15).Value2 & " step(s)"
    End With
End Sub

Private Sub UpdateAlignmentSummaries()
    Dim loSetup As ListObject
    Dim loReview As ListObject
    Dim loDep As ListObject
    Dim r As Long
    Dim alignmentName As String
    Dim relaxCount As Long
    Dim depCount As Long
    Dim findingCountForAlignment As Long

    Set loSetup = GetTable(SH_SETUP, TBL_SETUP)
    Set loReview = GetTable(SH_REVIEW, TBL_REVIEW)
    Set loDep = GetTable(SH_DEPART, TBL_DEPART)

    For r = 1 To loSetup.ListRows.count
        alignmentName = CStr(loSetup.DataBodyRange.Cells(r, 1).Value2)
        If Len(alignmentName) > 0 Then
            findingCountForAlignment = Application.WorksheetFunction.CountIf(loReview.ListColumns("Alignment").DataBodyRange, alignmentName)
            relaxCount = Application.WorksheetFunction.CountIfs(loReview.ListColumns("Alignment").DataBodyRange, alignmentName, _
                                                                 loReview.ListColumns("Review Status").DataBodyRange, "*RELAXATION*")
            depCount = Application.WorksheetFunction.CountIf(loDep.ListColumns("Alignment").DataBodyRange, alignmentName)
            loSetup.DataBodyRange.Cells(r, 6).Value = findingCountForAlignment & " findings | " & relaxCount & " relaxations | " & depCount & " departures"
        Else
            loSetup.DataBodyRange.Cells(r, 6).ClearContents
        End If
    Next r
End Sub

' ============================================================
' RULE LOOKUPS / COMPARISON HELPERS
' ============================================================

Private Function FindAlignmentSetting(ByVal alignmentName As String, ByRef speedCode As String, ByRef roadType As String, _
                                      ByRef profileName As String, ByRef includeText As String, ByRef setupStatus As String) As Boolean
    Dim lo As ListObject
    Dim r As Long

    Set lo = GetTable(SH_SETUP, TBL_SETUP)
    speedCode = vbNullString

    For r = 1 To lo.ListRows.count
        If CStr(lo.DataBodyRange.Cells(r, 1).Value2) = alignmentName Then
            profileName = CStr(lo.DataBodyRange.Cells(r, 2).Value2)
            roadType = CStr(lo.DataBodyRange.Cells(r, 3).Value2)
            includeText = CStr(lo.DataBodyRange.Cells(r, 4).Value2)
            setupStatus = CStr(lo.DataBodyRange.Cells(r, 5).Value2)
            FindAlignmentSetting = True
            Exit Function
        End If
    Next r
End Function

Private Function GetNumericSpeed(ByVal speedCode As String) As Long
    If Len(speedCode) >= 2 Then
        If IsNumeric(Left$(speedCode, Len(speedCode) - 1)) Then GetNumericSpeed = CLng(Left$(speedCode, Len(speedCode) - 1))
    End If
End Function

Private Function GetBand(ByVal speedCode As String) As String
    If Len(speedCode) > 0 Then GetBand = Right$(speedCode, 1)
End Function

Private Function SpeedIndex(ByVal speedValue As Long) As Long
    Select Case speedValue
        Case 120: SpeedIndex = 1
        Case 100: SpeedIndex = 2
        Case 85: SpeedIndex = 3
        Case 70: SpeedIndex = 4
        Case 60: SpeedIndex = 5
        Case 50: SpeedIndex = 6
        Case Else: SpeedIndex = 0
    End Select
End Function

Private Function SpeedFromIndex(ByVal idx As Long) As Long
    Select Case idx
        Case 1: SpeedFromIndex = 120
        Case 2: SpeedFromIndex = 100
        Case 3: SpeedFromIndex = 85
        Case 4: SpeedFromIndex = 70
        Case 5: SpeedFromIndex = 60
        Case 6: SpeedFromIndex = 50
    End Select
End Function

Private Function StepsTo50(ByVal speedValue As Long) As Long
    Dim idx As Long
    idx = SpeedIndex(speedValue)
    If idx > 0 Then StepsTo50 = 6 - idx
End Function

Private Function SpeedAfterSteps(ByVal speedValue As Long, ByVal steps As Long) As Long
    Dim idx As Long
    idx = SpeedIndex(speedValue)
    If idx = 0 Then Exit Function

    idx = idx + steps
    If idx > 6 Then idx = 6
    SpeedAfterSteps = SpeedFromIndex(idx)
End Function

Private Function GetDesignValue(ByVal speedValue As Long, ByVal parameterName As String) As Double
    Select Case speedValue
        Case 120
            Select Case parameterName
                Case "Stopping sight distance": GetDesignValue = 295
                Case "Horizontal curvature": GetDesignValue = 1020
                Case "Crest curvature": GetDesignValue = 182
                Case "Sag curvature": GetDesignValue = 37
            End Select
        Case 100
            Select Case parameterName
                Case "Stopping sight distance": GetDesignValue = 215
                Case "Horizontal curvature": GetDesignValue = 720
                Case "Crest curvature": GetDesignValue = 100
                Case "Sag curvature": GetDesignValue = 26
            End Select
        Case 85
            Select Case parameterName
                Case "Stopping sight distance": GetDesignValue = 160
                Case "Horizontal curvature": GetDesignValue = 510
                Case "Crest curvature": GetDesignValue = 55
                Case "Sag curvature": GetDesignValue = 20
            End Select
        Case 70
            Select Case parameterName
                Case "Stopping sight distance": GetDesignValue = 120
                Case "Horizontal curvature": GetDesignValue = 360
                Case "Crest curvature": GetDesignValue = 30
                Case "Sag curvature": GetDesignValue = 20
            End Select
        Case 60
            Select Case parameterName
                Case "Stopping sight distance": GetDesignValue = 90
                Case "Horizontal curvature": GetDesignValue = 255
                Case "Crest curvature": GetDesignValue = 17
                Case "Sag curvature": GetDesignValue = 13
            End Select
        Case 50
            Select Case parameterName
                Case "Stopping sight distance": GetDesignValue = 70
                Case "Horizontal curvature": GetDesignValue = 180
                Case "Crest curvature": GetDesignValue = 10
                Case "Sag curvature": GetDesignValue = 9
            End Select
    End Select
End Function

Private Function GetBaseSteps(ByVal parameterName As String, ByVal roadType As String, ByVal speedCode As String) As Long
    Dim band As String
    band = GetBand(speedCode)

    Select Case parameterName
        Case "Stopping sight distance"
            If roadType = "Motorway" Then
                If band = "A" Then GetBaseSteps = 1 Else GetBaseSteps = 2
            Else
                If band = "A" Then GetBaseSteps = 2 Else GetBaseSteps = 3
            End If

        Case "Horizontal curvature"
            If roadType = "Motorway" Then
                If band = "A" Then GetBaseSteps = 2 Else GetBaseSteps = 3
            Else
                If band = "A" Then GetBaseSteps = 3 Else GetBaseSteps = 4
            End If

        Case "Crest curvature"
            If roadType = "Motorway" Then
                If band = "A" Then GetBaseSteps = 1 Else GetBaseSteps = 2
            Else
                If band = "A" Then GetBaseSteps = 2 Else GetBaseSteps = 3
            End If

        Case "Sag curvature"
            If roadType = "Motorway" Then
                GetBaseSteps = 0
            ElseIf speedCode = "50B" Or speedCode = "60B" Or speedCode = "70B" Then
                GetBaseSteps = 2
            Else
                GetBaseSteps = 1
            End If
    End Select
End Function

Private Function BaseClauseForParameter(ByVal parameterName As String) As String
    Select Case parameterName
        Case "Stopping sight distance": BaseClauseForParameter = "CD 109 3.5 / Table 3.5"
        Case "Horizontal curvature": BaseClauseForParameter = "CD 109 4.5 / Table 4.5"
        Case "Crest curvature": BaseClauseForParameter = "CD 109 5.6 / Table 5.7"
        Case "Sag curvature": BaseClauseForParameter = "CD 109 5.9 / Table 5.9"
    End Select
End Function

Private Function GetContextAdjustment(ByVal alignmentName As String, ByVal parameterName As String, _
                                      ByVal startCh As Double, ByVal endCh As Double) As Long
    Dim lo As ListObject
    Dim r As Long

    Set lo = GetTable(SH_SSD_REVIEW, TBL_CONTEXTS)

    For r = 1 To lo.ListRows.count
        If CStr(lo.DataBodyRange.Cells(r, 6).Value2) = "Yes" And _
           CStr(lo.DataBodyRange.Cells(r, 10).Value2) = "OK" And _
           CStr(lo.DataBodyRange.Cells(r, 2).Value2) = alignmentName And _
           CStr(lo.DataBodyRange.Cells(r, 7).Value2) = parameterName Then

            If IsNumeric(lo.DataBodyRange.Cells(r, 3).Value2) And IsNumeric(lo.DataBodyRange.Cells(r, 4).Value2) Then
                If CDbl(lo.DataBodyRange.Cells(r, 3).Value2) < endCh And CDbl(lo.DataBodyRange.Cells(r, 4).Value2) > startCh Then
                    If IsNumeric(lo.DataBodyRange.Cells(r, 8).Value2) Then
                        GetContextAdjustment = GetContextAdjustment + CLng(lo.DataBodyRange.Cells(r, 8).Value2)
                    End If
                End If
            End If
        End If
    Next r
End Function

Private Function GetContextClauses(ByVal alignmentName As String, ByVal parameterName As String, _
                                   ByVal startCh As Double, ByVal endCh As Double) As String
    Dim lo As ListObject
    Dim r As Long
    Dim result As String
    Dim clauseText As String

    Set lo = GetTable(SH_SSD_REVIEW, TBL_CONTEXTS)

    For r = 1 To lo.ListRows.count
        If CStr(lo.DataBodyRange.Cells(r, 6).Value2) = "Yes" And _
           CStr(lo.DataBodyRange.Cells(r, 10).Value2) = "OK" And _
           CStr(lo.DataBodyRange.Cells(r, 2).Value2) = alignmentName And _
           CStr(lo.DataBodyRange.Cells(r, 7).Value2) = parameterName Then

            If IsNumeric(lo.DataBodyRange.Cells(r, 3).Value2) And IsNumeric(lo.DataBodyRange.Cells(r, 4).Value2) Then
                If CDbl(lo.DataBodyRange.Cells(r, 3).Value2) < endCh And CDbl(lo.DataBodyRange.Cells(r, 4).Value2) > startCh Then
                    clauseText = CStr(lo.DataBodyRange.Cells(r, 9).Value2)
                    If Len(result) > 0 Then result = result & ", "
                    result = result & "CD 109 " & clauseText
                End If
            End If
        End If
    Next r

    GetContextClauses = result
End Function

Private Function GetAchievedSteps(ByVal achieved As Double, ByVal speedValue As Long, ByVal parameterName As String) As Long
    Dim idx As Long
    Dim testIdx As Long
    Dim steps As Long
    Dim threshold As Double

    idx = SpeedIndex(speedValue)
    If idx = 0 Then
        GetAchievedSteps = -1
        Exit Function
    End If

    threshold = GetDesignValue(speedValue, parameterName)
    If GEQ3(achieved, threshold) Then
        GetAchievedSteps = 0
        Exit Function
    End If

    For testIdx = idx + 1 To 6
        steps = testIdx - idx
        threshold = GetDesignValue(SpeedFromIndex(testIdx), parameterName)

        If GEQ3(achieved, threshold) Then
            GetAchievedSteps = steps
            Exit Function
        End If
    Next testIdx

    GetAchievedSteps = -1
End Function

Private Sub GetGradientLimits(ByVal roadType As String, ByRef desirable As Double, ByRef permitted As Double)
    Select Case roadType
        Case "Motorway"
            desirable = 3
            permitted = 4
        Case "All-purpose dual carriageway"
            desirable = 4
            permitted = 8
        Case "All-purpose single carriageway"
            desirable = 6
            permitted = 8
    End Select
End Sub

Private Function GEQ3(ByVal a As Double, ByVal b As Double) As Boolean
    GEQ3 = (Application.WorksheetFunction.Round(a, 3) >= Application.WorksheetFunction.Round(b, 3))
End Function

Private Function LEQ3(ByVal a As Double, ByVal b As Double) As Boolean
    LEQ3 = (Application.WorksheetFunction.Round(a, 3) <= Application.WorksheetFunction.Round(b, 3))
End Function

Private Function CompactValue(ByVal parameterName As String, ByVal v As Double) As String
    Select Case parameterName
        Case "Horizontal curvature"
            CompactValue = "R" & Format$(v, "0.###")
        Case "Crest curvature", "Sag curvature"
            CompactValue = "K" & Format$(v, "0.###")
        Case "Stopping sight distance"
            CompactValue = Format$(v, "0.###") & " m"
        Case "Longitudinal gradient", "Superelevation", "SE edge grade variation", "SE minimum longitudinal gradient"
            CompactValue = Format$(v, "0.###") & "%"
        Case Else
            CompactValue = Format$(v, "0.###")
    End Select
End Function

Private Function CountZoneOverlaps(ByVal alignmentName As String, ByVal parameterName As String, _
                                   ByVal startCh As Double, ByVal endCh As Double) As Long
    Dim z As Long
    For z = 1 To zoneCount
        If Zones(z).Alignment = alignmentName And Zones(z).Parameter = parameterName Then
            If Zones(z).startCh < endCh And Zones(z).endCh > startCh Then
                CountZoneOverlaps = CountZoneOverlaps + 1
            End If
        End If
    Next z
End Function

Private Function GetAlignmentLimits(ByVal alignmentName As String, ByRef startCh As Double, ByRef endCh As Double) As Boolean
    Dim lo As ListObject
    Dim r As Long

    Set lo = GetTable(SH_ALIGN, TBL_ALIGN)

    For r = 1 To lo.ListRows.count
        If CStr(lo.DataBodyRange.Cells(r, 1).Value2) = alignmentName Then
            If IsNumeric(lo.DataBodyRange.Cells(r, 2).Value2) And IsNumeric(lo.DataBodyRange.Cells(r, 3).Value2) Then
                startCh = CDbl(lo.DataBodyRange.Cells(r, 2).Value2)
                endCh = CDbl(lo.DataBodyRange.Cells(r, 3).Value2)
                GetAlignmentLimits = True
            End If
            Exit Function
        End If
    Next r
End Function

' ============================================================
' LANDXML IMPORT
' ============================================================

Private Function IsWebPath(ByVal filePath As String) As Boolean
    Dim p As String

    p = LCase$(Trim$(filePath))
    IsWebPath = (Left$(p, 7) = "http://" Or Left$(p, 8) = "https://")
End Function

Private Function LocalFileExists(ByVal filePath As String) As Boolean
    Dim fso As Object

    If Len(Trim$(filePath)) = 0 Then Exit Function
    If IsWebPath(filePath) Then Exit Function

    On Error GoTo NotFound
    Set fso = CreateObject("Scripting.FileSystemObject")
    LocalFileExists = fso.FileExists(filePath)
    Exit Function

NotFound:
    LocalFileExists = False
End Function

Private Function FileNameFromPath(ByVal filePath As String) As String
    Dim s As String
    Dim p As Long
    Dim q As Long

    s = Trim$(filePath)
    If Len(s) = 0 Then Exit Function

    ' Remove URL query/fragment text before extracting the leaf name.
    q = InStr(1, s, "?", vbBinaryCompare)
    If q > 0 Then s = Left$(s, q - 1)

    q = InStr(1, s, "#", vbBinaryCompare)
    If q > 0 Then s = Left$(s, q - 1)

    ' Accept both Windows paths and HTTP/SharePoint URLs.
    s = Replace$(s, "\", "/")
    p = InStrRev(s, "/")

    If p > 0 Then
        FileNameFromPath = Mid$(s, p + 1)
    Else
        FileNameFromPath = s
    End If
End Function

Private Function DefaultXmlPath() As String
    Dim basePath As String

    basePath = ThisWorkbook.Path
    If Len(basePath) = 0 Then Exit Function

    ' ThisWorkbook.Path can be an https:// SharePoint / OneDrive URL.
    ' Dir$, FileSystemObject and Windows path concatenation are not valid for URLs.
    If IsWebPath(basePath) Then Exit Function

    DefaultXmlPath = basePath & Application.PathSeparator & "DMRB_Compliance.xml"
End Function

Private Function PickLandXMLFile() As String
    Dim fd As Object

    Set fd = Application.FileDialog(3)
    With fd
        .Title = "Select Civil 3D LandXML file"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "LandXML / XML", "*.xml"
        If Len(ThisWorkbook.Path) > 0 Then .InitialFileName = ThisWorkbook.Path & Application.PathSeparator

        If .Show = -1 Then PickLandXMLFile = CStr(.SelectedItems(1))
    End With
End Function


Private Function PickSuperelevationCSVFile() As String
    Dim fd As Object

    Set fd = Application.FileDialog(3)
    With fd
        .Title = "Select DMRB Superelevation Exporter CSV"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Superelevation CSV", "*.csv"
        If Len(ThisWorkbook.Path) > 0 Then .InitialFileName = ThisWorkbook.Path & Application.PathSeparator

        If .Show = -1 Then PickSuperelevationCSVFile = CStr(.SelectedItems(1))
    End With
End Function

Private Sub SaveReviewSettings(ByVal d As Object)
    Dim lo As ListObject
    Dim r As Long
    Dim aln As String

    Set lo = GetTable(SH_SETUP, TBL_SETUP)
    For r = 1 To lo.ListRows.count
        aln = Trim$(CStr(lo.DataBodyRange.Cells(r, 1).Value2))
        If Len(aln) > 0 Then
            d.Item(aln) = Array( _
                lo.DataBodyRange.Cells(r, 2).Value2, _
                lo.DataBodyRange.Cells(r, 3).Value2, _
                lo.DataBodyRange.Cells(r, 4).Value2)
        End If
    Next r
End Sub

Private Sub SaveSpeedZoneBands(ByVal d As Object)
    Dim lo As ListObject
    Dim r As Long
    Dim aln As String
    Dim band As String
    Dim stationValue As Variant
    Dim speedValue As Variant
    Dim key As String

    Set lo = GetTable(SH_DESIGN_SPEEDS, TBL_SPEED_ZONES)

    For r = 1 To lo.ListRows.count
        aln = Trim$(CStr(lo.DataBodyRange.Cells(r, 1).Value2))
        band = UCase$(Trim$(CStr(lo.DataBodyRange.Cells(r, 5).Value2)))
        stationValue = lo.DataBodyRange.Cells(r, 7).Value2
        speedValue = lo.DataBodyRange.Cells(r, 4).Value2

        If Len(aln) > 0 And (band = "A" Or band = "B") And IsNumeric(stationValue) And IsNumeric(speedValue) Then
            key = aln & "|" & Format$(CDbl(stationValue), "0.000000") & "|" & Format$(CDbl(speedValue), "0.###")
            d.Item(key) = band
        End If
    Next r
End Sub

Private Function SavedBand(ByVal d As Object, ByVal aln As String, ByVal stationValue As Double, _
                           ByVal speedValue As Double) As String
    Dim key As String
    key = aln & "|" & Format$(stationValue, "0.000000") & "|" & Format$(speedValue, "0.###")
    If d.Exists(key) Then SavedBand = CStr(d.Item(key))
End Function

Private Sub SyncSpeedZones(ByVal savedBands As Object)
    Dim loZones As ListObject
    Dim loA As ListObject
    Dim loS As ListObject
    Dim a As Long
    Dim r As Long
    Dim i As Long
    Dim j As Long
    Dim count As Long
    Dim totalRows As Long
    Dim outRow As Long
    Dim alignmentName As String
    Dim alnStart As Double
    Dim alnEnd As Double
    Dim stations() As Double
    Dim speeds() As Double
    Dim tStation As Double
    Dim tSpeed As Double
    Dim band As String

    Set loZones = GetTable(SH_DESIGN_SPEEDS, TBL_SPEED_ZONES)
    Set loA = GetTable(SH_ALIGN, TBL_ALIGN)
    Set loS = GetTable(SH_SPEEDS, TBL_SPEEDS)

    totalRows = 0
    For a = 1 To loA.ListRows.count
        alignmentName = Trim$(CStr(loA.DataBodyRange.Cells(a, 1).Value2))
        If Len(alignmentName) > 0 Then
            count = 0
            For r = 1 To loS.ListRows.count
                If CStr(loS.DataBodyRange.Cells(r, 2).Value2) = alignmentName And _
                   IsNumeric(loS.DataBodyRange.Cells(r, 3).Value2) And _
                   IsNumeric(loS.DataBodyRange.Cells(r, 4).Value2) Then count = count + 1
            Next r
            If count = 0 Then count = 1
            totalRows = totalRows + count
        End If
    Next a

    ResizeTableRows loZones, IIf(totalRows > 0, totalRows, 1)
    loZones.DataBodyRange.ClearContents

    outRow = 0
    For a = 1 To loA.ListRows.count
        alignmentName = Trim$(CStr(loA.DataBodyRange.Cells(a, 1).Value2))
        If Len(alignmentName) > 0 And IsNumeric(loA.DataBodyRange.Cells(a, 2).Value2) And _
           IsNumeric(loA.DataBodyRange.Cells(a, 3).Value2) Then

            alnStart = CDbl(loA.DataBodyRange.Cells(a, 2).Value2)
            alnEnd = CDbl(loA.DataBodyRange.Cells(a, 3).Value2)
            count = 0
            Erase stations
            Erase speeds

            For r = 1 To loS.ListRows.count
                If CStr(loS.DataBodyRange.Cells(r, 2).Value2) = alignmentName And _
                   IsNumeric(loS.DataBodyRange.Cells(r, 3).Value2) And _
                   IsNumeric(loS.DataBodyRange.Cells(r, 4).Value2) Then
                    count = count + 1
                    If count = 1 Then
                        ReDim stations(1 To 1)
                        ReDim speeds(1 To 1)
                    Else
                        ReDim Preserve stations(1 To count)
                        ReDim Preserve speeds(1 To count)
                    End If
                    stations(count) = CDbl(loS.DataBodyRange.Cells(r, 3).Value2)
                    speeds(count) = CDbl(loS.DataBodyRange.Cells(r, 4).Value2)
                End If
            Next r

            If count > 1 Then
                For i = 1 To count - 1
                    For j = i + 1 To count
                        If stations(j) < stations(i) Then
                            tStation = stations(i): stations(i) = stations(j): stations(j) = tStation
                            tSpeed = speeds(i): speeds(i) = speeds(j): speeds(j) = tSpeed
                        End If
                    Next j
                Next i
            End If

            If count = 0 Then
                outRow = outRow + 1
                loZones.DataBodyRange.Cells(outRow, 1).Value = alignmentName
                loZones.DataBodyRange.Cells(outRow, 2).Value = alnStart
                loZones.DataBodyRange.Cells(outRow, 3).Value = alnEnd
            Else
                For i = 1 To count
                    outRow = outRow + 1
                    loZones.DataBodyRange.Cells(outRow, 1).Value = alignmentName
                    If i = 1 Then
                        loZones.DataBodyRange.Cells(outRow, 2).Value = alnStart
                    Else
                        loZones.DataBodyRange.Cells(outRow, 2).Value = stations(i)
                    End If
                    If i < count Then
                        loZones.DataBodyRange.Cells(outRow, 3).Value = stations(i + 1)
                    Else
                        loZones.DataBodyRange.Cells(outRow, 3).Value = alnEnd
                    End If
                    loZones.DataBodyRange.Cells(outRow, 4).Value = speeds(i)
                    loZones.DataBodyRange.Cells(outRow, 7).Value = stations(i)
                    band = SavedBand(savedBands, alignmentName, stations(i), speeds(i))
                    If Len(band) > 0 Then loZones.DataBodyRange.Cells(outRow, 5).Value = band
                Next i
            End If
        End If
    Next a

    RefreshSpeedZoneDerivedValues
End Sub

Private Sub RefreshSpeedZoneDerivedValues()
    Dim lo As ListObject
    Dim r As Long
    Dim aln As String
    Dim band As String
    Dim speedValue As Variant
    Dim code As String
    Dim statusText As String

    Set lo = GetTable(SH_DESIGN_SPEEDS, TBL_SPEED_ZONES)
    If lo.DataBodyRange Is Nothing Then Exit Sub

    For r = 1 To lo.ListRows.count
        aln = Trim$(CStr(lo.DataBodyRange.Cells(r, 1).Value2))
        speedValue = lo.DataBodyRange.Cells(r, 4).Value2
        band = UCase$(Trim$(CStr(lo.DataBodyRange.Cells(r, 5).Value2)))
        code = vbNullString
        statusText = vbNullString

        If Len(aln) = 0 Then
            statusText = vbNullString
        ElseIf Not IsNumeric(speedValue) Then
            statusText = "NO C3D SPEED"
        ElseIf Not IsSupportedSpeed(CDbl(speedValue)) Then
            statusText = "UNSUPPORTED C3D SPEED"
        ElseIf band <> "A" And band <> "B" Then
            statusText = "BAND REQUIRED"
        Else
            code = Format$(CDbl(speedValue), "0") & band
            statusText = "OK"
        End If

        lo.DataBodyRange.Cells(r, 6).Value = code
        lo.DataBodyRange.Cells(r, 8).Value = statusText
    Next r
End Sub

Private Function IsSupportedSpeed(ByVal speedValue As Double) As Boolean
    Select Case CLng(speedValue)
        Case 50, 60, 70, 85, 100, 120
            IsSupportedSpeed = (Abs(speedValue - CLng(speedValue)) < 0.000001)
    End Select
End Function

Private Function AlignmentFallbackSpeedCode(ByVal alignmentName As String) As String
    AlignmentFallbackSpeedCode = vbNullString
End Function

Private Function ResolveReviewSpeed(ByVal alignmentName As String, ByVal startCh As Double, _
                                    ByVal endCh As Double, ByVal fallbackCode As String) As String
    Dim lo As ListObject
    Dim r As Long
    Dim zStart As Variant
    Dim zEnd As Variant
    Dim speedValue As Variant
    Dim band As String
    Dim code As String
    Dim speedN As Long
    Dim bestSpeed As Long
    Dim bestBand As String
    Dim pointRange As Boolean
    Dim bestPointStart As Double
    Dim havePoint As Boolean
    Dim overlapCount As Long
    Dim validOverlapCount As Long

    Set lo = GetTable(SH_DESIGN_SPEEDS, TBL_SPEED_ZONES)
    ResolveReviewSpeed = vbNullString

    If endCh < startCh Then SwapDoubleValues startCh, endCh
    pointRange = (Abs(endCh - startCh) <= 0.000001)

    If pointRange Then
        For r = 1 To lo.ListRows.count
            If Trim$(CStr(lo.DataBodyRange.Cells(r, 1).Value2)) = alignmentName Then
                zStart = lo.DataBodyRange.Cells(r, 2).Value2
                zEnd = lo.DataBodyRange.Cells(r, 3).Value2
                speedValue = lo.DataBodyRange.Cells(r, 4).Value2
                band = UCase$(Trim$(CStr(lo.DataBodyRange.Cells(r, 5).Value2)))
                If IsNumeric(zStart) And IsNumeric(zEnd) Then
                    If startCh >= CDbl(zStart) - 0.001 And startCh <= CDbl(zEnd) + 0.001 Then
                        If IsNumeric(speedValue) And IsSupportedSpeed(CDbl(speedValue)) And (band = "A" Or band = "B") Then
                            If Not havePoint Or CDbl(zStart) > bestPointStart + 0.000001 Then
                                bestPointStart = CDbl(zStart)
                                ResolveReviewSpeed = Format$(CDbl(speedValue), "0") & band
                                havePoint = True
                            End If
                        End If
                    End If
                End If
            End If
        Next r
        Exit Function
    End If

    For r = 1 To lo.ListRows.count
        If Trim$(CStr(lo.DataBodyRange.Cells(r, 1).Value2)) = alignmentName Then
            zStart = lo.DataBodyRange.Cells(r, 2).Value2
            zEnd = lo.DataBodyRange.Cells(r, 3).Value2
            speedValue = lo.DataBodyRange.Cells(r, 4).Value2
            band = UCase$(Trim$(CStr(lo.DataBodyRange.Cells(r, 5).Value2)))
            If IsNumeric(zStart) And IsNumeric(zEnd) Then
                If CDbl(zStart) < endCh - 0.000001 And CDbl(zEnd) > startCh + 0.000001 Then
                    overlapCount = overlapCount + 1
                    If IsNumeric(speedValue) And IsSupportedSpeed(CDbl(speedValue)) And (band = "A" Or band = "B") Then
                        validOverlapCount = validOverlapCount + 1
                        code = Format$(CDbl(speedValue), "0") & band
                        speedN = GetNumericSpeed(code)
                        If speedN > bestSpeed Then
                            bestSpeed = speedN
                            bestBand = band
                            ResolveReviewSpeed = code
                        ElseIf speedN = bestSpeed And speedN > 0 Then
                            If band = "A" And bestBand <> "A" Then
                                bestBand = "A"
                                ResolveReviewSpeed = code
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next r

    If overlapCount = 0 Or validOverlapCount <> overlapCount Then ResolveReviewSpeed = vbNullString
End Function

Private Function CountValidSpeedZonesOverlapping(ByVal alignmentName As String, ByVal startCh As Double, _
                                                 ByVal endCh As Double) As Long
    Dim lo As ListObject
    Dim r As Long
    Dim zStart As Variant
    Dim zEnd As Variant
    Dim speedValue As Variant
    Dim band As String

    Set lo = GetTable(SH_DESIGN_SPEEDS, TBL_SPEED_ZONES)
    If endCh < startCh Then SwapDoubleValues startCh, endCh

    For r = 1 To lo.ListRows.count
        If Trim$(CStr(lo.DataBodyRange.Cells(r, 1).Value2)) = alignmentName Then
            zStart = lo.DataBodyRange.Cells(r, 2).Value2
            zEnd = lo.DataBodyRange.Cells(r, 3).Value2
            speedValue = lo.DataBodyRange.Cells(r, 4).Value2
            band = UCase$(Trim$(CStr(lo.DataBodyRange.Cells(r, 5).Value2)))

            If IsNumeric(zStart) And IsNumeric(zEnd) And IsNumeric(speedValue) And _
               IsSupportedSpeed(CDbl(speedValue)) And (band = "A" Or band = "B") Then

                If CDbl(zStart) < endCh - 0.000001 And CDbl(zEnd) > startCh + 0.000001 Then
                    CountValidSpeedZonesOverlapping = CountValidSpeedZonesOverlapping + 1
                End If
            End If
        End If
    Next r
End Function

Private Function AppendSpeedInterfaceNote(ByVal clauseText As String, ByVal alignmentName As String, _
                                          ByVal startCh As Double, ByVal endCh As Double) As String
    If CountValidSpeedZonesOverlapping(alignmentName, startCh, endCh) > 1 Then
        If Len(clauseText) > 0 Then clauseText = clauseText & "; "
        clauseText = clauseText & "CD 109 2.9.2 - geometry spans multiple C3D design-speed zones; highest overlapping speed used"
    End If

    AppendSpeedInterfaceNote = clauseText
End Function

Private Sub SwapDoubleValues(ByRef a As Double, ByRef b As Double)
    Dim t As Double
    t = a
    a = b
    b = t
End Sub

Private Sub ResetRawTables()
    ResetRawTable SH_ALIGN, TBL_ALIGN
    ResetRawTable SH_SPEEDS, TBL_SPEEDS
    ResetRawTable SH_HORIZ, TBL_HORIZ
    ResetRawTable SH_PROFILES, TBL_PROFILES
    ResetRawTable SH_VERT, TBL_VERT
End Sub

Private Sub ResetRawTable(ByVal sheetName As String, ByVal tableName As String)
    Dim lo As ListObject

    Set lo = GetTable(sheetName, tableName)

    ' Do not resize the raw table before import. Excel can temporarily expose
    ' DataBodyRange as Nothing immediately after ListObject.Resize, which caused
    ' Error 91 after Clear / Start New Review had left a previously populated
    ' raw table ready for the next LandXML import.
    '
    ' Clearing the existing body is sufficient. Import routines write from row 5
    ' and ResizeRawTables sets the final table extent after the XML has been read.
    If Not lo.DataBodyRange Is Nothing Then
        lo.DataBodyRange.ClearContents
    End If
End Sub

Private Sub ResizeRawTables()
    ResizeRawTableToLastRow SH_ALIGN, TBL_ALIGN, "F"
    ResizeRawTableToLastRow SH_SPEEDS, TBL_SPEEDS, "D"
    ResizeRawTableToLastRow SH_HORIZ, TBL_HORIZ, "K"
    ResizeRawTableToLastRow SH_PROFILES, TBL_PROFILES, "D"
    ResizeRawTableToLastRow SH_VERT, TBL_VERT, "L"
End Sub

Private Sub ResizeRawTableToLastRow(ByVal sheetName As String, ByVal tableName As String, ByVal lastCol As String)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim lastRow As Long

    Set ws = ThisWorkbook.Worksheets(sheetName)
    Set lo = ws.ListObjects(tableName)

    lastRow = LastUsedRow(ws, "A")
    If lastRow < 5 Then lastRow = 5

    lo.Resize ws.Range("A4:" & lastCol & CStr(lastRow))
End Sub

Private Sub ImportAlignments(ByVal xmlDoc As Object)
    Dim ws As Worksheet
    Dim alns As Object
    Dim aln As Object
    Dim r As Long
    Dim s0 As Double
    Dim L As Double
    Dim importedSpeed As Variant
    Dim speedStation As Variant

    Set ws = ThisWorkbook.Worksheets(SH_ALIGN)
    Set alns = xmlDoc.SelectNodes("//lx:Alignments/lx:Alignment")
    r = 5

    For Each aln In alns
        s0 = NumberOrZero(XmlAttr(aln, "staStart"))
        L = NumberOrZero(XmlAttr(aln, "length"))
        importedSpeed = vbNullString
        speedStation = vbNullString
        GetSpeedStation aln, importedSpeed, speedStation

        ws.Cells(r, "A").Value = XmlAttr(aln, "name")
        ws.Cells(r, "B").Value = s0
        ws.Cells(r, "C").Value = s0 + L
        ws.Cells(r, "D").Value = L
        ws.Cells(r, "E").Value = importedSpeed
        ws.Cells(r, "F").Value = speedStation
        r = r + 1
    Next aln
End Sub

Private Sub GetSpeedStation(ByVal aln As Object, ByRef speedValue As Variant, ByRef stationValue As Variant)
    Dim featureNode As Object
    Dim prop As Object
    Dim lbl As String

    speedValue = vbNullString
    stationValue = vbNullString

    Set featureNode = aln.SelectSingleNode("./lx:Feature[@name='SpeedStation']")
    If featureNode Is Nothing Then Exit Sub

    For Each prop In featureNode.SelectNodes("./lx:Property")
        lbl = LCase$(XmlAttr(prop, "label"))
        If lbl = "speed" And IsNumeric(XmlAttr(prop, "value")) Then speedValue = CDbl(XmlAttr(prop, "value"))
        If lbl = "station" And IsNumeric(XmlAttr(prop, "value")) Then stationValue = CDbl(XmlAttr(prop, "value"))
    Next prop
End Sub

Private Sub ImportSpeedStations(ByVal xmlDoc As Object)
    Dim ws As Worksheet
    Dim aln As Object
    Dim featureNode As Object
    Dim prop As Object
    Dim r As Long
    Dim rawNo As Long
    Dim stationValue As Variant
    Dim speedValue As Variant
    Dim lbl As String

    Set ws = ThisWorkbook.Worksheets(SH_SPEEDS)
    r = 5
    rawNo = 0

    For Each aln In xmlDoc.SelectNodes("//lx:Alignments/lx:Alignment")
        For Each featureNode In aln.SelectNodes("./lx:Feature[@name='SpeedStation']")
            stationValue = vbNullString
            speedValue = vbNullString

            For Each prop In featureNode.SelectNodes("./lx:Property")
                lbl = LCase$(XmlAttr(prop, "label"))
                If lbl = "station" And IsNumeric(XmlAttr(prop, "value")) Then stationValue = CDbl(XmlAttr(prop, "value"))
                If lbl = "speed" And IsNumeric(XmlAttr(prop, "value")) Then speedValue = CDbl(XmlAttr(prop, "value"))
            Next prop

            rawNo = rawNo + 1
            ws.Cells(r, "A").Value = "SRAW-" & Format$(rawNo, "0000")
            ws.Cells(r, "B").Value = XmlAttr(aln, "name")
            ws.Cells(r, "C").Value = stationValue
            ws.Cells(r, "D").Value = speedValue
            r = r + 1
        Next featureNode
    Next aln
End Sub

Private Sub ImportHorizontal(ByVal xmlDoc As Object)
    Dim ws As Worksheet
    Dim aln As Object
    Dim coord As Object
    Dim ent As Object
    Dim r As Long
    Dim rawNo As Long
    Dim ch As Double
    Dim L As Double
    Dim t As String
    Dim radiusValue As Variant
    Dim direction As String
    Dim radiusStart As String
    Dim radiusEnd As String
    Dim finiteRadius As Double
    Dim aValue As Variant

    Set ws = ThisWorkbook.Worksheets(SH_HORIZ)
    r = 5

    For Each aln In xmlDoc.SelectNodes("//lx:Alignments/lx:Alignment")
        ch = NumberOrZero(XmlAttr(aln, "staStart"))
        Set coord = aln.SelectSingleNode("./lx:CoordGeom")

        If Not coord Is Nothing Then
            For Each ent In coord.ChildNodes
                If ent.NodeType = 1 Then
                    t = CStr(ent.baseName)

                    If t = "Line" Or t = "Spiral" Or t = "Curve" Then
                        rawNo = rawNo + 1
                        L = NumberOrZero(XmlAttr(ent, "length"))
                        radiusValue = vbNullString
                        direction = vbNullString
                        radiusStart = vbNullString
                        radiusEnd = vbNullString
                        aValue = vbNullString

                        If t = "Curve" Then
                            If IsNumeric(XmlAttr(ent, "radius")) Then radiusValue = CDbl(XmlAttr(ent, "radius"))
                            direction = RotationText(XmlAttr(ent, "rot"))
                        ElseIf t = "Spiral" Then
                            direction = RotationText(XmlAttr(ent, "rot"))
                            radiusStart = XmlAttr(ent, "radiusStart")
                            radiusEnd = XmlAttr(ent, "radiusEnd")
                            finiteRadius = FiniteRadiusFromText(radiusStart, radiusEnd)
                            If finiteRadius > 0 And L > 0 Then aValue = Sqr(finiteRadius * L)
                        End If

                        ws.Cells(r, "A").Value = "HRAW-" & Format$(rawNo, "0000")
                        ws.Cells(r, "B").Value = XmlAttr(aln, "name")
                        ws.Cells(r, "C").Value = t
                        ws.Cells(r, "D").Value = ch
                        ws.Cells(r, "E").Value = ch + L
                        ws.Cells(r, "F").Value = L
                        ws.Cells(r, "G").Value = radiusValue
                        ws.Cells(r, "H").Value = direction
                        ws.Cells(r, "I").Value = radiusStart
                        ws.Cells(r, "J").Value = radiusEnd
                        ws.Cells(r, "K").Value = aValue

                        ch = ch + L
                        r = r + 1
                    End If
                End If
            Next ent
        End If
    Next aln
End Sub

Private Function RotationText(ByVal rot As String) As String
    If LCase$(rot) = "cw" Then
        RotationText = "Right"
    ElseIf LCase$(rot) = "ccw" Then
        RotationText = "Left"
    End If
End Function

Private Function FiniteRadiusFromText(ByVal radiusStart As String, ByVal radiusEnd As String) As Double
    If IsNumeric(radiusStart) Then
        FiniteRadiusFromText = CDbl(radiusStart)
    ElseIf IsNumeric(radiusEnd) Then
        FiniteRadiusFromText = CDbl(radiusEnd)
    End If
End Function

Private Sub ImportProfilesAndVertical(ByVal xmlDoc As Object)
    Dim wsP As Worksheet
    Dim wsV As Worksheet
    Dim aln As Object
    Dim surfProfile As Object
    Dim layoutProfile As Object
    Dim rowP As Long
    Dim rowV As Long
    Dim rawNo As Long

    Set wsP = ThisWorkbook.Worksheets(SH_PROFILES)
    Set wsV = ThisWorkbook.Worksheets(SH_VERT)
    rowP = 5
    rowV = 5

    For Each aln In xmlDoc.SelectNodes("//lx:Alignments/lx:Alignment")
        For Each surfProfile In aln.SelectNodes("./lx:Profile/lx:ProfSurf")
            wsP.Cells(rowP, "A").Value = XmlAttr(aln, "name")
            wsP.Cells(rowP, "B").Value = XmlAttr(surfProfile, "name")
            wsP.Cells(rowP, "C").Value = "Surface"
            wsP.Cells(rowP, "D").Value = XmlAttr(surfProfile, "state")
            rowP = rowP + 1
        Next surfProfile

        For Each layoutProfile In aln.SelectNodes("./lx:Profile/lx:ProfAlign")
            wsP.Cells(rowP, "A").Value = XmlAttr(aln, "name")
            wsP.Cells(rowP, "B").Value = XmlAttr(layoutProfile, "name")
            wsP.Cells(rowP, "C").Value = "Layout"
            wsP.Cells(rowP, "D").Value = XmlAttr(layoutProfile, "state")
            rowP = rowP + 1

            ImportAllVerticalGeometry aln, layoutProfile, wsV, rowV, rawNo
        Next layoutProfile
    Next aln
End Sub

Private Sub ImportAllVerticalGeometry(ByVal aln As Object, ByVal profileNode As Object, ByVal ws As Worksheet, _
                                      ByRef rowOut As Long, ByRef rawNo As Long)
    Dim seq As Object
    Dim node As Object
    Dim nextNode As Object
    Dim i As Long
    Dim sta As Double
    Dim elev As Double
    Dim nextSta As Double
    Dim nextElev As Double
    Dim prevSta As Double
    Dim prevElev As Double
    Dim afterSta As Double
    Dim afterElev As Double
    Dim curveLength As Double
    Dim nextCurveLength As Double
    Dim gradeIn As Double
    Dim gradeOut As Double
    Dim tangentGrade As Double
    Dim deltaGrade As Double
    Dim kValue As Variant
    Dim curveType As String
    Dim startCh As Double
    Dim endCh As Double
    Dim firstProfileRow As Long

    firstProfileRow = rowOut
    Set seq = profileNode.SelectNodes("./*")

    For i = 0 To seq.Length - 1
        Set node = seq.Item(i)

        If node.NodeType = 1 Then
            If node.baseName = "ParaCurve" Or node.baseName = "CircCurve" Then
                If i > 0 And i < seq.Length - 1 Then
                    If ReadStationElevation(node, sta, elev) And _
                       FindPreviousProfilePoint(seq, i, prevSta, prevElev) And _
                       FindNextProfilePoint(seq, i, afterSta, afterElev) Then

                        curveLength = NumberOrZero(XmlAttr(node, "length"))

                        If sta <> prevSta And afterSta <> sta Then
                            gradeIn = (elev - prevElev) / (sta - prevSta) * 100
                            gradeOut = (afterElev - elev) / (afterSta - sta) * 100
                            deltaGrade = gradeOut - gradeIn

                            If Abs(deltaGrade) > 0.000000000001 Then kValue = Abs(curveLength / deltaGrade) Else kValue = vbNullString
                            If gradeOut < gradeIn Then curveType = "Crest" Else curveType = "Sag"

                            rawNo = rawNo + 1
                            ws.Cells(rowOut, "A").Value = "VRAW-" & Format$(rawNo, "0000")
                            ws.Cells(rowOut, "B").Value = XmlAttr(aln, "name")
                            ws.Cells(rowOut, "C").Value = XmlAttr(profileNode, "name")
                            ws.Cells(rowOut, "D").Value = curveType
                            ws.Cells(rowOut, "E").Value = sta - curveLength / 2
                            ws.Cells(rowOut, "F").Value = sta + curveLength / 2
                            ws.Cells(rowOut, "G").Value = curveLength
                            ws.Cells(rowOut, "H").Value = gradeIn
                            ws.Cells(rowOut, "I").Value = gradeOut
                            ws.Cells(rowOut, "J").Value = kValue
                            ws.Cells(rowOut, "K").Value = sta
                            ws.Cells(rowOut, "L").Value = elev
                            rowOut = rowOut + 1
                        End If
                    End If
                End If
            End If

            If i < seq.Length - 1 Then
                Set nextNode = seq.Item(i + 1)

                If ReadStationElevation(node, sta, elev) And ReadStationElevation(nextNode, nextSta, nextElev) Then
                    If nextSta > sta Then
                        curveLength = CurveLengthIfAny(node)
                        nextCurveLength = CurveLengthIfAny(nextNode)
                        startCh = sta + curveLength / 2
                        endCh = nextSta - nextCurveLength / 2

                        If endCh > startCh + 0.000000001 Then
                            tangentGrade = (nextElev - elev) / (nextSta - sta) * 100

                            rawNo = rawNo + 1
                            ws.Cells(rowOut, "A").Value = "VRAW-" & Format$(rawNo, "0000")
                            ws.Cells(rowOut, "B").Value = XmlAttr(aln, "name")
                            ws.Cells(rowOut, "C").Value = XmlAttr(profileNode, "name")
                            ws.Cells(rowOut, "D").Value = "Tangent"
                            ws.Cells(rowOut, "E").Value = startCh
                            ws.Cells(rowOut, "F").Value = endCh
                            ws.Cells(rowOut, "G").Value = endCh - startCh
                            ws.Cells(rowOut, "H").Value = tangentGrade
                            ws.Cells(rowOut, "I").Value = tangentGrade
                            rowOut = rowOut + 1
                        End If
                    End If
                End If
            End If
        End If
    Next i

    SortProfileVerticalRows ws, firstProfileRow, rowOut - 1
End Sub

Private Sub SortProfileVerticalRows(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    If lastRow < firstRow Then Exit Sub

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws.Range("E" & firstRow & ":E" & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending
        .SetRange ws.Range("A" & firstRow & ":L" & lastRow)
        .Header = xlNo
        .Apply
    End With
End Sub

Private Function CurveLengthIfAny(ByVal node As Object) As Double
    If node.NodeType = 1 Then
        If node.baseName = "ParaCurve" Or node.baseName = "CircCurve" Then CurveLengthIfAny = NumberOrZero(XmlAttr(node, "length"))
    End If
End Function

Private Function FindPreviousProfilePoint(ByVal seq As Object, ByVal startIndex As Long, ByRef sta As Double, ByRef elev As Double) As Boolean
    Dim j As Long
    For j = startIndex - 1 To 0 Step -1
        If ReadStationElevation(seq.Item(j), sta, elev) Then
            FindPreviousProfilePoint = True
            Exit Function
        End If
    Next j
End Function

Private Function FindNextProfilePoint(ByVal seq As Object, ByVal startIndex As Long, ByRef sta As Double, ByRef elev As Double) As Boolean
    Dim j As Long
    For j = startIndex + 1 To seq.Length - 1
        If ReadStationElevation(seq.Item(j), sta, elev) Then
            FindNextProfilePoint = True
            Exit Function
        End If
    Next j
End Function

Private Function ReadStationElevation(ByVal node As Object, ByRef sta As Double, ByRef elev As Double) As Boolean
    Dim txt As String
    Dim parts() As String

    On Error GoTo NoPoint
    txt = NormalizeSpaces(CStr(node.Text))
    If Len(txt) = 0 Then GoTo NoPoint

    parts = Split(txt, " ")
    If UBound(parts) >= 1 Then
        If IsNumeric(parts(0)) And IsNumeric(parts(1)) Then
            sta = CDbl(parts(0))
            elev = CDbl(parts(1))
            ReadStationElevation = True
            Exit Function
        End If
    End If

NoPoint:
    ReadStationElevation = False
End Function

Private Function NormalizeSpaces(ByVal txt As String) As String
    txt = Replace(txt, vbTab, " ")
    txt = Replace(txt, vbCr, " ")
    txt = Replace(txt, vbLf, " ")
    Do While InStr(txt, "  ") > 0
        txt = Replace(txt, "  ", " ")
    Loop
    NormalizeSpaces = Trim$(txt)
End Function

Private Function IsPreferredSuperelevationRow(ByVal lo As ListObject, ByVal rowIndex As Long) As Boolean
    Dim alignmentName As String
    Dim regionName As String
    Dim description As String
    Dim stationValue As Variant
    Dim otherStation As Variant
    Dim j As Long

    alignmentName = Trim$(CStr(lo.DataBodyRange.Cells(rowIndex, 3).Value2))
    regionName = Trim$(CStr(lo.DataBodyRange.Cells(rowIndex, 4).Value2))
    stationValue = lo.DataBodyRange.Cells(rowIndex, 5).Value2
    description = Trim$(CStr(lo.DataBodyRange.Cells(rowIndex, 6).Value2))

    If Len(alignmentName) = 0 Or Not IsNumeric(stationValue) Then Exit Function

    If description <> "Begin Alignment" And description <> "End Alignment" Then
        IsPreferredSuperelevationRow = True
        Exit Function
    End If

    If Len(regionName) > 0 Then
        IsPreferredSuperelevationRow = True
        Exit Function
    End If

    If description = "End Alignment" Then
        ' A generic End Alignment row is often a bookkeeping/reset row rather
        ' than a new superelevation control point. If Civil 3D has already given
        ' region-specific superelevation data on this alignment, do not use the
        ' generic end row to ramp the achieved xfall back toward its end value.
        ' Instead the last real superelevation xfall continues to the end unless
        ' another region-specific critical station says otherwise.
        For j = 1 To lo.ListRows.count
            If j <> rowIndex Then
                If Trim$(CStr(lo.DataBodyRange.Cells(j, 3).Value2)) = alignmentName And _
                   Len(Trim$(CStr(lo.DataBodyRange.Cells(j, 4).Value2))) > 0 Then

                    otherStation = lo.DataBodyRange.Cells(j, 5).Value2
                    If IsNumeric(otherStation) Then
                        If CDbl(otherStation) <= CDbl(stationValue) + 0.000001 Then Exit Function
                    End If
                End If
            End If
        Next j

        IsPreferredSuperelevationRow = True
        Exit Function
    End If

    ' Begin Alignment remains a useful fallback for the initial crossfall, but
    ' when a region-specific critical station exists at the same chainage the
    ' region-specific value takes precedence.
    For j = 1 To lo.ListRows.count
        If j <> rowIndex Then
            If Trim$(CStr(lo.DataBodyRange.Cells(j, 3).Value2)) = alignmentName And _
               Len(Trim$(CStr(lo.DataBodyRange.Cells(j, 4).Value2))) > 0 Then

                otherStation = lo.DataBodyRange.Cells(j, 5).Value2
                If IsNumeric(otherStation) Then
                    If Abs(CDbl(otherStation) - CDbl(stationValue)) <= 0.000001 Then Exit Function
                End If
            End If
        End If
    Next j

    IsPreferredSuperelevationRow = True
End Function

Private Sub GetRowCrossfalls(ByVal lo As ListObject, ByVal rowIndex As Long, _
                             ByRef leftXfall As Double, ByRef hasLeft As Boolean, _
                             ByRef rightXfall As Double, ByRef hasRight As Boolean)
    Dim v As Variant

    hasLeft = False
    hasRight = False

    v = lo.DataBodyRange.Cells(rowIndex, 9).Value2
    If IsNumeric(v) Then
        leftXfall = CDbl(v)
        hasLeft = True
    End If

    v = lo.DataBodyRange.Cells(rowIndex, 13).Value2
    If IsNumeric(v) Then
        rightXfall = CDbl(v)
        hasRight = True
    End If
End Sub

Private Function CrossfallRowsDiffer(ByVal leftA As Double, ByVal hasLeftA As Boolean, _
                                     ByVal rightA As Double, ByVal hasRightA As Boolean, _
                                     ByVal leftB As Double, ByVal hasLeftB As Boolean, _
                                     ByVal rightB As Double, ByVal hasRightB As Boolean) As Boolean
    If hasLeftA And hasLeftB Then
        If Abs(leftA - leftB) > 0.001 Then
            CrossfallRowsDiffer = True
            Exit Function
        End If
    End If

    If hasRightA And hasRightB Then
        If Abs(rightA - rightB) > 0.001 Then
            CrossfallRowsDiffer = True
            Exit Function
        End If
    End If
End Function

Private Sub GetCrossfallsAtChainageWithinProfile(ByVal alignmentName As String, ByVal chainage As Double, _
                                                   ByVal profileStart As Double, ByVal profileEnd As Double, _
                                                   ByRef leftXfall As Double, ByRef hasLeft As Boolean, _
                                                   ByRef rightXfall As Double, ByRef hasRight As Boolean)
    hasLeft = GetCrossfallSideAtChainageWithinProfile(alignmentName, chainage, profileStart, profileEnd, 9, leftXfall)
    hasRight = GetCrossfallSideAtChainageWithinProfile(alignmentName, chainage, profileStart, profileEnd, 13, rightXfall)
End Sub

Private Function GetCrossfallSideAtChainageWithinProfile(ByVal alignmentName As String, ByVal chainage As Double, _
                                                          ByVal profileStart As Double, ByVal profileEnd As Double, _
                                                          ByVal laneColumn As Long, ByRef xfall As Double) As Boolean
    Dim lo As ListObject
    Dim r As Long
    Dim sta As Variant
    Dim val As Variant
    Dim havePrev As Boolean
    Dim haveNext As Boolean
    Dim prevSta As Double
    Dim nextSta As Double
    Dim prevVal As Double
    Dim nextVal As Double
    Dim thisSta As Double
    Dim thisVal As Double

    If profileEnd < profileStart Then SwapDoubleValues profileStart, profileEnd

    Set lo = GetTable(SH_SUPER, TBL_SUPER)

    For r = 1 To lo.ListRows.count
        If IsPreferredSuperelevationRow(lo, r) Then
            If Trim$(CStr(lo.DataBodyRange.Cells(r, 3).Value2)) = alignmentName Then
                sta = lo.DataBodyRange.Cells(r, 5).Value2
                val = lo.DataBodyRange.Cells(r, laneColumn).Value2

                If IsNumeric(sta) And IsNumeric(val) Then
                    thisSta = CDbl(sta)
                    thisVal = CDbl(val)

                    If Abs(thisSta - chainage) <= 0.000001 Then
                        xfall = thisVal
                        GetCrossfallSideAtChainageWithinProfile = True
                        Exit Function
                    End If

                    ' A preceding critical station may legitimately sit before the
                    ' selected profile start: it establishes the crossfall carried
                    ' into the reviewed profile. A following critical station is
                    ' only allowed to drive interpolation when it is still inside
                    ' the selected profile. If it lies beyond the profile end, the
                    ' last achieved crossfall is carried forward instead.
                    If thisSta < chainage Then
                        If Not havePrev Or thisSta > prevSta Then
                            prevSta = thisSta
                            prevVal = thisVal
                            havePrev = True
                        End If
                    ElseIf thisSta > chainage And thisSta <= profileEnd + 0.000001 Then
                        If Not haveNext Or thisSta < nextSta Then
                            nextSta = thisSta
                            nextVal = thisVal
                            haveNext = True
                        End If
                    End If
                End If
            End If
        End If
    Next r

    If havePrev And haveNext Then
        If Abs(nextSta - prevSta) <= 0.000001 Then
            xfall = prevVal
        Else
            xfall = prevVal + (chainage - prevSta) * (nextVal - prevVal) / (nextSta - prevSta)
        End If
        GetCrossfallSideAtChainageWithinProfile = True
    ElseIf havePrev Then
        xfall = prevVal
        GetCrossfallSideAtChainageWithinProfile = True
    ElseIf haveNext Then
        xfall = nextVal
        GetCrossfallSideAtChainageWithinProfile = True
    End If
End Function

Private Sub CollectXfallBreakpoints(ByVal alignmentName As String, ByVal startCh As Double, ByVal endCh As Double, _
                                    ByRef points() As Double, ByRef pointCount As Long)
    Dim lo As ListObject
    Dim r As Long
    Dim sta As Variant
    Dim i As Long
    Dim j As Long
    Dim t As Double
    Dim isDup As Boolean

    If endCh < startCh Then SwapDoubleValues startCh, endCh

    pointCount = 2
    ReDim points(1 To 2)
    points(1) = startCh
    points(2) = endCh

    Set lo = GetTable(SH_SUPER, TBL_SUPER)

    For r = 1 To lo.ListRows.count
        If IsPreferredSuperelevationRow(lo, r) Then
            If Trim$(CStr(lo.DataBodyRange.Cells(r, 3).Value2)) = alignmentName Then
                sta = lo.DataBodyRange.Cells(r, 5).Value2

                If IsNumeric(sta) Then
                    If CDbl(sta) > startCh + 0.001 And CDbl(sta) < endCh - 0.001 Then
                        isDup = False
                        For i = 1 To pointCount
                            If Abs(points(i) - CDbl(sta)) <= 0.000001 Then
                                isDup = True
                                Exit For
                            End If
                        Next i

                        If Not isDup Then
                            pointCount = pointCount + 1
                            ReDim Preserve points(1 To pointCount)
                            points(pointCount) = CDbl(sta)
                        End If
                    End If
                End If
            End If
        End If
    Next r

    For i = 1 To pointCount - 1
        For j = i + 1 To pointCount
            If points(j) < points(i) Then
                t = points(i)
                points(i) = points(j)
                points(j) = t
            End If
        Next j
    Next i
End Sub

Private Function MinimumAbsoluteLinearXfall(ByVal xfallA As Double, ByVal xfallB As Double) As Double
    If Abs(xfallA) <= 0.0000001 Or Abs(xfallB) <= 0.0000001 Or _
       (xfallA < 0# And xfallB > 0#) Or (xfallA > 0# And xfallB < 0#) Then
        MinimumAbsoluteLinearXfall = 0#
    Else
        MinimumAbsoluteLinearXfall = LesserDouble(Abs(xfallA), Abs(xfallB))
    End If
End Function

Private Function CsvStationValue(ByVal sourceCell As Range) As Variant
    Dim txt As String
    txt = Trim$(CStr(sourceCell.Text))
    If Len(txt) = 0 Then
        CsvStationValue = vbNullString
        Exit Function
    End If

    If LCase$(Right$(txt, 1)) = "m" Then txt = Trim$(Left$(txt, Len(txt) - 1))
    If IsNumeric(txt) Then
        CsvStationValue = CDbl(txt)
    ElseIf IsNumeric(sourceCell.Value2) Then
        CsvStationValue = CDbl(sourceCell.Value2)
    Else
        CsvStationValue = vbNullString
    End If
End Function

Private Function CsvPlainNumber(ByVal sourceCell As Range) As Variant
    Dim txt As String
    txt = Trim$(CStr(sourceCell.Text))
    If Len(txt) = 0 Then
        CsvPlainNumber = vbNullString
    ElseIf IsNumeric(txt) Then
        CsvPlainNumber = CDbl(txt)
    ElseIf IsNumeric(sourceCell.Value2) Then
        CsvPlainNumber = CDbl(sourceCell.Value2)
    Else
        CsvPlainNumber = vbNullString
    End If
End Function

Private Function CsvSlopePercent(ByVal sourceCell As Range) As Variant
    Dim txt As String
    txt = Trim$(CStr(sourceCell.Text))

    If Len(txt) = 0 Then
        CsvSlopePercent = vbNullString
        Exit Function
    End If

    If InStr(1, txt, "%", vbTextCompare) > 0 Then
        txt = Replace(txt, "%", "")
        txt = Trim$(txt)
        If IsNumeric(txt) Then
            CsvSlopePercent = CDbl(txt)
            Exit Function
        End If
    End If

    If IsNumeric(sourceCell.Value2) Then
        If InStr(1, CStr(sourceCell.NumberFormat), "%", vbTextCompare) > 0 Then
            CsvSlopePercent = CDbl(sourceCell.Value2) * 100#
        Else
            CsvSlopePercent = CDbl(sourceCell.Value2)
        End If
    Else
        CsvSlopePercent = vbNullString
    End If
End Function

Private Sub SyncReviewSetup(ByVal savedSettings As Object)
    Dim lo As ListObject
    Dim loA As ListObject
    Dim importedCount As Long
    Dim dataRows As Long
    Dim r As Long
    Dim aln As String
    Dim v As Variant

    Set lo = GetTable(SH_SETUP, TBL_SETUP)
    Set loA = GetTable(SH_ALIGN, TBL_ALIGN)

    importedCount = TableNonBlankCount(loA, 1)
    dataRows = importedCount
    If dataRows < 30 Then dataRows = 30
    ResizeTableRows lo, dataRows
    lo.DataBodyRange.ClearContents

    For r = 1 To importedCount
        aln = CStr(loA.DataBodyRange.Cells(r, 1).Value2)
        lo.DataBodyRange.Cells(r, 1).Value = aln
        If savedSettings.Exists(aln) Then
            v = savedSettings.Item(aln)
            lo.DataBodyRange.Cells(r, 2).Value = v(0)
            lo.DataBodyRange.Cells(r, 3).Value = v(1)
            lo.DataBodyRange.Cells(r, 4).Value = IIf(Len(CStr(v(2))) > 0, v(2), "Yes")
        Else
            lo.DataBodyRange.Cells(r, 2).Value = SoleLayoutProfileForAlignment(aln)
            lo.DataBodyRange.Cells(r, 4).Value = "Yes"
        End If
    Next r
End Sub

Private Function SoleLayoutProfileForAlignment(ByVal alignmentName As String) As String
    Dim lo As ListObject
    Dim r As Long
    Dim count As Long
    Dim profileName As String

    Set lo = GetTable(SH_PROFILES, TBL_PROFILES)

    For r = 1 To lo.ListRows.count
        If CStr(lo.DataBodyRange.Cells(r, 1).Value2) = alignmentName And CStr(lo.DataBodyRange.Cells(r, 3).Value2) = "Layout" Then
            count = count + 1
            profileName = CStr(lo.DataBodyRange.Cells(r, 2).Value2)
        End If
    Next r

    If count = 1 Then SoleLayoutProfileForAlignment = profileName
End Function

Private Sub SyncSuperelevationSettings()
    Dim loSet As ListObject
    Dim loA As ListObject
    Dim a As Long
    Dim r As Long
    Dim targetRow As Long
    Dim found As Boolean
    Dim alignmentName As String
    Dim maxID As Long
    Dim idText As String
    Dim n As Long

    Set loSet = GetTable(SH_SUPER_SETUP, TBL_SUPER_SETTINGS)
    Set loA = GetTable(SH_ALIGN, TBL_ALIGN)

    For r = 1 To loSet.ListRows.count
        idText = CStr(loSet.DataBodyRange.Cells(r, 1).Value2)
        If Left$(idText, 3) = "SE-" Then
            If IsNumeric(Mid$(idText, 4)) Then
                n = CLng(Mid$(idText, 4))
                If n > maxID Then maxID = n
            End If
        End If
    Next r

    For a = 1 To loA.ListRows.count
        alignmentName = Trim$(CStr(loA.DataBodyRange.Cells(a, 1).Value2))
        If Len(alignmentName) > 0 Then
            found = False
            For r = 1 To loSet.ListRows.count
                If CStr(loSet.DataBodyRange.Cells(r, 2).Value2) = alignmentName Then
                    found = True
                    If Len(Trim$(CStr(loSet.DataBodyRange.Cells(r, 9).Value2))) = 0 Then loSet.DataBodyRange.Cells(r, 9).Value = "Proposed"
                    Exit For
                End If
            Next r

            If Not found Then
                targetRow = 0
                For r = 1 To loSet.ListRows.count
                    If Len(Trim$(CStr(loSet.DataBodyRange.Cells(r, 1).Value2))) = 0 Then
                        targetRow = r
                        Exit For
                    End If
                Next r
                If targetRow = 0 Then
                    loSet.ListRows.Add
                    targetRow = loSet.ListRows.count
                End If

                maxID = maxID + 1
                loSet.DataBodyRange.Cells(targetRow, 1).Value = "SE-" & Format$(maxID, "000")
                loSet.DataBodyRange.Cells(targetRow, 2).Value = alignmentName
                loSet.DataBodyRange.Cells(targetRow, 3).Value = loA.DataBodyRange.Cells(a, 2).Value2
                loSet.DataBodyRange.Cells(targetRow, 4).Value = loA.DataBodyRange.Cells(a, 3).Value2
                loSet.DataBodyRange.Cells(targetRow, 6).Value = "Yes"
                loSet.DataBodyRange.Cells(targetRow, 9).Value = "Proposed"
            End If
        End If
    Next a

    RefreshSuperSettingStatus
End Sub

Private Sub RefreshSuperSettingStatus()
    Dim lo As ListObject
    Dim r As Long
    Dim activeText As String
    Dim environment As String
    Dim assessmentBasis As String

    Set lo = GetTable(SH_SUPER_SETUP, TBL_SUPER_SETTINGS)
    If lo.DataBodyRange Is Nothing Then Exit Sub

    For r = 1 To lo.ListRows.count
        activeText = Trim$(CStr(lo.DataBodyRange.Cells(r, 6).Value2))
        environment = Trim$(CStr(lo.DataBodyRange.Cells(r, 5).Value2))
        assessmentBasis = Trim$(CStr(lo.DataBodyRange.Cells(r, 9).Value2))

        If Len(Trim$(CStr(lo.DataBodyRange.Cells(r, 1).Value2))) = 0 Then
            lo.DataBodyRange.Cells(r, 7).ClearContents
        ElseIf activeText <> "Yes" Then
            lo.DataBodyRange.Cells(r, 7).Value = "INACTIVE"
        ElseIf environment <> "Rural" And environment <> "Urban" Then
            lo.DataBodyRange.Cells(r, 7).Value = "ENVIRONMENT REQUIRED"
        ElseIf assessmentBasis <> "Proposed" And assessmentBasis <> "Existing retained" Then
            lo.DataBodyRange.Cells(r, 7).Value = "ASSESSMENT BASIS REQUIRED"
        Else
            lo.DataBodyRange.Cells(r, 7).Value = "OK"
        End If
    Next r
End Sub

Private Sub UpdateLists()
    Dim lo As ListObject
    Dim loA As ListObject
    Dim loP As ListObject
    Dim loRules As ListObject
    Dim alignCount As Long
    Dim layoutCount As Long
    Dim maxRows As Long
    Dim r As Long
    Dim p As Long
    Dim speedCodes As Variant
    Dim jTypes As Variant
    Dim firstRow As Long
    Dim firstCol As Long

    Set lo = GetTable(SH_LISTS, TBL_LISTS)
    Set loA = GetTable(SH_ALIGN, TBL_ALIGN)
    Set loP = GetTable(SH_PROFILES, TBL_PROFILES)
    Set loRules = GetTable(SH_REF, TBL_CONTEXT_RULES)

    ' The final column stores the defined relaxation-context choices.
    If lo.ListColumns.count < 12 Then
        firstRow = lo.HeaderRowRange.Row
        firstCol = lo.HeaderRowRange.Column
        lo.Resize lo.Parent.Range(lo.Parent.Cells(firstRow, firstCol), _
                                  lo.Parent.Cells(firstRow + lo.ListRows.count, firstCol + 11))
    End If

    lo.HeaderRowRange.Cells(1, 12).Value = "Relaxation Context Types"

    alignCount = TableNonBlankCount(loA, 1)

    For r = 1 To loP.ListRows.count
        If CStr(loP.DataBodyRange.Cells(r, 3).Value2) = "Layout" And Len(CStr(loP.DataBodyRange.Cells(r, 1).Value2)) > 0 Then
            layoutCount = layoutCount + 1
        End If
    Next r

    maxRows = 30
    If alignCount > maxRows Then maxRows = alignCount
    If layoutCount > maxRows Then maxRows = layoutCount
    If loRules.ListRows.count > maxRows Then maxRows = loRules.ListRows.count

    ResizeTableRows lo, maxRows
    lo.DataBodyRange.ClearContents

    lo.DataBodyRange.Cells(1, 1).Value = "Yes"
    lo.DataBodyRange.Cells(2, 1).Value = "No"

    For r = 1 To alignCount
        lo.DataBodyRange.Cells(r, 2).Value = loA.DataBodyRange.Cells(r, 1).Value2
    Next r

    p = 0
    For r = 1 To loP.ListRows.count
        If CStr(loP.DataBodyRange.Cells(r, 3).Value2) = "Layout" And Len(CStr(loP.DataBodyRange.Cells(r, 1).Value2)) > 0 Then
            p = p + 1
            lo.DataBodyRange.Cells(p, 3).Value = loP.DataBodyRange.Cells(r, 2).Value2
        End If
    Next r

    lo.DataBodyRange.Cells(1, 4).Value = "Motorway"
    lo.DataBodyRange.Cells(2, 4).Value = "All-purpose dual carriageway"
    lo.DataBodyRange.Cells(3, 4).Value = "All-purpose single carriageway"

    jTypes = Array( _
        "Minor road priority - no taper (Note 1)", _
        "Major road priority - no taper (Note 2)", _
        "At-grade diverge taper (Note 3)", _
        "At-grade diverge auxiliary lane (Note 4)", _
        "At-grade merge taper (Note 5)", _
        "Roundabout approach (Note 6)", _
        "Grade-separated diverge (Note 7)", _
        "Grade-separated merge (Note 8)")

    For r = 0 To UBound(jTypes)
        lo.DataBodyRange.Cells(r + 1, 5).Value = jTypes(r)
    Next r

    lo.DataBodyRange.Cells(1, 6).Value = "Increasing chainage"
    lo.DataBodyRange.Cells(2, 6).Value = "Decreasing chainage"
    lo.DataBodyRange.Cells(3, 6).Value = "Both"

    lo.DataBodyRange.Cells(1, 7).Value = "Increasing chainage"
    lo.DataBodyRange.Cells(2, 7).Value = "Decreasing chainage"

    speedCodes = Array("120A", "120B", "100A", "100B", "85A", "85B", "70A", "70B", "60A", "60B", "50A", "50B")
    For r = 0 To UBound(speedCodes)
        lo.DataBodyRange.Cells(r + 1, 8).Value = speedCodes(r)
    Next r

    lo.DataBodyRange.Cells(1, 9).Value = "A"
    lo.DataBodyRange.Cells(2, 9).Value = "B"

    lo.DataBodyRange.Cells(1, 10).Value = "Rural"
    lo.DataBodyRange.Cells(2, 10).Value = "Urban"

    lo.DataBodyRange.Cells(1, 11).Value = "Proposed"
    lo.DataBodyRange.Cells(2, 11).Value = "Existing retained"

    For r = 1 To loRules.ListRows.count
        If Len(Trim$(CStr(loRules.DataBodyRange.Cells(r, 1).Value2))) > 0 Then
            lo.DataBodyRange.Cells(r, 12).Value = loRules.DataBodyRange.Cells(r, 1).Value2
        End If
    Next r
End Sub

Private Sub RecordRefresh(ByVal filePath As String)
    With ThisWorkbook.Worksheets(SH_SETUP)
        .Range("G5").Value = FileNameFromPath(filePath)
        .Range("G6").Value = Now
        .Range("G6").NumberFormat = "yyyy-mm-dd hh:mm"
        .Range("G9").Value = TableNonBlankCount(GetTable(SH_ALIGN, TBL_ALIGN), 1)
    End With
End Sub

Private Sub RecordSuperelevationRefresh(ByVal filePath As String)
    With ThisWorkbook.Worksheets(SH_SETUP)
        .Range("G7").Value = FileNameFromPath(filePath)
        .Range("G8").Value = Now
        .Range("G8").NumberFormat = "yyyy-mm-dd hh:mm"
    End With
End Sub

Private Sub ApplyHelperVisibility()
    ThisWorkbook.Worksheets(SH_CONTROL).Visible = xlSheetVeryHidden
    ThisWorkbook.Worksheets(SH_LISTS).Visible = xlSheetVeryHidden
End Sub

Private Sub UnprotectRepairSheets()
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        ws.Unprotect Password:=""
        On Error GoTo 0
    Next ws
End Sub

Private Sub UnprotectRawSheets()
    Dim n As Variant

    For Each n In Array(SH_ALIGN, SH_SPEEDS, SH_HORIZ, SH_PROFILES, SH_VERT, SH_SUPER)
        ThisWorkbook.Worksheets(CStr(n)).Unprotect Password:=""
    Next n
End Sub

Private Function LastUsedRow(ByVal ws As Worksheet, ByVal col As String) As Long
    Dim r As Long
    r = ws.Cells(ws.Rows.count, col).End(xlUp).Row
    If r < 5 Then r = 4
    LastUsedRow = r
End Function

Private Function XmlAttr(ByVal node As Object, ByVal attrName As String) As String
    Dim a As Object
    On Error Resume Next
    Set a = node.Attributes.getNamedItem(attrName)
    On Error GoTo 0

    If a Is Nothing Then
        XmlAttr = vbNullString
    Else
        XmlAttr = CStr(a.Text)
    End If
End Function

Private Function NumberOrZero(ByVal txt As String) As Double
    If IsNumeric(txt) Then NumberOrZero = CDbl(txt)
End Function
