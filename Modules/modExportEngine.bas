Attribute VB_Name = "modExportEngine"
' ============================================================
' modExportEngine.bas
' Creation wbOut en valeurs, details, groupements, enregistrement
' ============================================================
Option Explicit

Public Sub ExportValuesCopy_WithoutLeads_ToBalanceFolder_V4()
    Dim folder As String, defaultName As String
    Dim pickedPath As Variant, finalPath As String
    Dim wb As Workbook, wbOut As Workbook
    Dim wsSrc As Worksheet, wsDest As Worksheet
    Dim delErr As String
    Dim firstCreated As Boolean
    Dim exportMode As eExportMode
    Dim oldAlerts As Boolean, oldScreen As Boolean

    oldAlerts = Application.DisplayAlerts
    oldScreen = Application.ScreenUpdating

    ' --- Chemin de sauvegarde ---
    folder = GetFolderFromPath(gBalancePath)
    If Len(Trim$(folder)) = 0 Then folder = ThisWorkbook.Path & Application.PathSeparator
    If Right$(folder, 1) <> Application.PathSeparator Then folder = folder & Application.PathSeparator
    defaultName = BuildDefaultFileName(gClient, gExercice, gVersion)
    If Len(defaultName) > 180 Then defaultName = Left$(defaultName, 180)

    pickedPath = PromptSaveAsPath_NoUI(folder & defaultName)
    If pickedPath = False Then GoTo CLEAN_EXIT

    finalPath = CStr(pickedPath)
    If LCase(Right(finalPath, 5)) <> ".xlsx" Then finalPath = finalPath & ".xlsx"

    ' Verifier que le fichier cible n'est pas deja ouvert
    For Each wb In Application.Workbooks
        If LCase(wb.FullName) = LCase(finalPath) Then
            MsgBox "Le fichier cible est deja ouvert :" & vbCrLf & finalPath & vbCrLf & "Fermez-le puis relancez.", vbExclamation
            GoTo CLEAN_EXIT
        End If
    Next wb

    ' Supprimer fichier existant si confirmation
    If FileExists(finalPath) Then
        If MsgBox("Un fichier existe deja :" & vbCrLf & finalPath & vbCrLf & vbCrLf & _
                  "Souhaitez-vous l'ecraser ?", vbQuestion + vbYesNo, "Fichier existant") <> vbYes Then
            GoTo CLEAN_EXIT
        End If
        If Not TryDeleteFile(finalPath, delErr) Then
            MsgBox "Impossible de supprimer le fichier existant." & vbCrLf & delErr, vbCritical
            GoTo CLEAN_EXIT
        End If
    End If

    Application.DisplayAlerts  = False
    Application.ScreenUpdating = False
    On Error GoTo EH

    ' --- Creer wbOut et copier les onglets en VALEURS ---
    Set wbOut = Workbooks.Add(xlWBATWorksheet)
    wbOut.Worksheets(1).Name = "TMP_DELETE"

    firstCreated = False
    exportMode = GetExportModeFromFrmLeadMeta()

    For Each wsSrc In ThisWorkbook.Worksheets
        If ShouldExportSheet(wsSrc.Name, exportMode) Then
            If Not firstCreated Then
                Set wsDest = wbOut.Worksheets(1)
                firstCreated = True
            Else
                Set wsDest = wbOut.Worksheets.Add(After:=wbOut.Worksheets(wbOut.Worksheets.Count))
            End If

            wsDest.Name = GetUniqueSheetName(wbOut, wsSrc.Name)
            CopySheetIntoExistingAsValuesAndFormats wsSrc, wsDest

            If wsSrc.Visible = xlSheetVisible Then
                wsDest.Visible = xlSheetVisible
            Else
                wsDest.Visible = xlSheetHidden
            End If
        End If
    Next wsSrc

    ' --- Preparer BS et SIG du wbOut ---
    On Error Resume Next
    Set wsDest = wbOut.Worksheets(SH_BS)
    On Error GoTo EH
    If Not wsDest Is Nothing Then PrepareBS_ForExport wbOut, wsDest
    PurgeZeroRowsBS wbOut

    Set wsDest = Nothing
    On Error Resume Next
    Set wsDest = wbOut.Worksheets(SH_SIG)
    On Error GoTo EH
    If Not wsDest Is Nothing Then PrepareSIG_ForExport wbOut, wsDest

    ' --- Nettoyer onglets superflus ---
    Application.DisplayAlerts = False
    DeleteSheetIfExists wbOut, "Param"
    DeleteSheetIfExists wbOut, "Mapping"
    DeleteSheetIfExists wbOut, "TMP_DELETE"
    Application.DisplayAlerts = True

    ' --- Onglets detail (BS_detail, SIG_detail) ---
    EnsureDetailedTabsGenerated wbOut

    ' --- Finalisation details ---
    Finalize_DetailSheets wbOut

    ' --- Supprimer colonnes formules BG (E:S) dans wbOut ---
    On Error Resume Next
    With wbOut.Worksheets(SH_BG)
        .Columns("E:S").Delete
    End With
    On Error GoTo EH

    ' --- Supprimer onglets selon mode export ---
    PruneExportSheetsByMode wbOut, exportMode

    ' --- Groupements lignes BS et SIG ---
    ApplyGrouping_WbOut wbOut

    ' --- Copier niveaux de regroupement depuis wbSource ---
    On Error Resume Next
    Set wsDest = wbOut.Worksheets(SH_BS)
    If Not wsDest Is Nothing Then CopyOutlines_FromSourceToOut ThisWorkbook.Worksheets(SH_BS), wsDest
    Set wsDest = Nothing
    Set wsDest = wbOut.Worksheets(SH_SIG)
    If Not wsDest Is Nothing Then CopyOutlines_FromSourceToOut ThisWorkbook.Worksheets(SH_SIG), wsDest
    On Error GoTo EH

    ' --- Groupement colonnes C:D sur BS et BS_detail ---
    EnsureCDColumnsGroupedCollapsed wbOut, SH_BS
    EnsureCDColumnsGroupedCollapsed wbOut, "BS_detail"

    ' --- Replier tous les groupes a niveau 1 ---
    CollapseAllGroupedSheetsToLevel1 wbOut

    ' --- Masquer grilles et finaliser affichage ---
    SetGridlinesOff_AllSheetsViews wbOut

    ' --- Enregistrer ---
    Application.DisplayAlerts = True
    wbOut.SaveAs Filename:=finalPath, FileFormat:=xlOpenXMLWorkbook, Local:=True
    Application.DisplayAlerts = False

    ' --- Ouvrir le wbOut, zoom 75% ---
    On Error Resume Next
    wbOut.Activate
    If wbOut.Windows.Count > 0 Then
        wbOut.Windows(1).Zoom = 75
        wbOut.Windows(1).DisplayGridlines = False
    End If
    On Error GoTo EH

CLEAN_EXIT:
    Application.DisplayAlerts  = oldAlerts
    Application.ScreenUpdating = oldScreen
    Exit Sub
EH:
    modKETrace.LogKE "ERROR " & Err.Number & " : " & Err.Description, "ExportValuesCopy"
    If Len(Trim$(Err.Description)) > 0 Then
        MsgBox "Erreur Export : " & Err.Number & vbCrLf & Err.Description, vbCritical
    End If
    Err.Clear
    Resume CLEAN_EXIT
End Sub

' ============================================================
' HELPERS INTERNES
' ============================================================
Private Sub PurgeZeroRowsBS(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim vE As Variant, vF As Variant

    On Error GoTo EH

    On Error Resume Next
    Set ws = wb.Worksheets(SH_BS)
    On Error GoTo EH
    If ws Is Nothing Then GoTo CleanExit

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then GoTo CleanExit

    For i = lastRow To 2 Step -1
        vE = ws.Cells(i, 5).Value
        vF = ws.Cells(i, 6).Value
        If vE = "" Or vF = "" Then GoTo NextRow
        If Not IsNumeric(vE) Or Not IsNumeric(vF) Then GoTo NextRow
        If CDbl(vE) = 0 And CDbl(vF) = 0 Then
            ws.Rows(i).Delete
        End If
NextRow:
    Next i

CleanExit:
    Exit Sub
EH:
    Debug.Print "PurgeZeroRowsBS error " & Err.Number & " : " & Err.Description
    Resume CleanExit
End Sub

Private Function ShouldExportSheet(ByVal sheetName As String, ByVal mode As eExportMode) As Boolean
    Dim nm As String
    nm = UCase$(Trim$(sheetName))
    ' Toujours exclure
    If nm = "LEADS" Or nm = "PARAM" Or nm = "MAPPING" Or nm = "ACCUEIL" Then Exit Function
    Select Case mode
        Case emAll
            ShouldExportSheet = True
        Case emFS, emLeads
            ShouldExportSheet = (nm = UCase$(SH_BG) Or nm = UCase$(SH_BS) Or nm = "BS_DETAIL")
        Case Else
            ShouldExportSheet = True
    End Select
End Function

Private Sub PruneExportSheetsByMode(ByVal wbOut As Workbook, ByVal mode As eExportMode)
    Dim ws As Worksheet, nm As String
    If mode = emAll Then Exit Sub
    Application.DisplayAlerts = False
    For Each ws In wbOut.Worksheets
        nm = UCase$(ws.Name)
        If nm <> UCase$(SH_BG) And nm <> UCase$(SH_BS) And nm <> "BS_DETAIL" Then
            On Error Resume Next
            ws.Delete
            On Error GoTo 0
        End If
    Next ws
    Application.DisplayAlerts = True
End Sub

Private Sub PrepareBS_ForExport(ByVal wbOut As Workbook, ByVal wsBS As Worksheet)
    wsBS.Rows(1).Hidden = True
    BSX_ApplyRowGroupings wsBS
    BSX_HideRowsWhereGHBlank_InRanges wsBS
End Sub

Private Sub PrepareSIG_ForExport(ByVal wbOut As Workbook, ByVal wsSIG As Worksheet)
    ' Pas de prep specifique SIG pour l'instant
End Sub

Public Sub BSX_HideRowsWhereGHBlank_InRanges(ByVal wsBS As Worksheet)
    BSX_HideRowsWhereGHBlank_Range wsBS, 13, 54
    BSX_HideRowsWhereGHBlank_Range wsBS, 59, 92
    BSX_HideRowsWhereGHBlank_Range wsBS, 97, 159
End Sub

Private Sub BSX_HideRowsWhereGHBlank_Range(ByVal wsBS As Worksheet, ByVal r1 As Long, ByVal r2 As Long)
    Dim r As Long
    For r = r1 To r2
        wsBS.Rows(r).Hidden = (BSX_IsBlank(wsBS.Cells(r, "G").Value) And BSX_IsBlank(wsBS.Cells(r, "H").Value))
    Next r
End Sub

Private Function BSX_IsBlank(ByVal v As Variant) As Boolean
    If IsEmpty(v) Then BSX_IsBlank = True: Exit Function
    BSX_IsBlank = (Len(Trim$(CStr(v))) = 0)
End Function

Public Sub BSX_ApplyRowGroupings(ByVal wsBS As Worksheet)
    On Error Resume Next
    wsBS.Rows("1:" & wsBS.Rows.Count).Ungroup
    wsBS.Cells.ClearOutline
    On Error GoTo 0
    wsBS.Rows("13:54").Group
    wsBS.Rows("59:92").Group
    wsBS.Rows("97:159").Group
    wsBS.Outline.SummaryRow = xlSummaryBelow
End Sub

Public Sub ApplyGrouping_WbOut(ByVal wb As Workbook)
    Dim bsRanges As Variant
    bsRanges = Array(Array(13, 54), Array(59, 90), Array(94, 149))
    GroupRowsIfSheetExists wb, SH_BS, bsRanges

    Dim sigRanges As Variant
    sigRanges = Array( _
        Array(11, 12), Array(14, 16), Array(19, 20), Array(23, 24), _
        Array(26, 30), Array(32, 37), Array(39, 42), Array(44, 48))
    GroupRowsIfSheetExists wb, SH_SIG, sigRanges
End Sub

Public Sub GroupRowsIfSheetExists(ByVal wb As Workbook, ByVal sheetName As String, ByVal ranges As Variant)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    Dim i As Long
    For i = LBound(ranges) To UBound(ranges)
        ws.Rows(CLng(ranges(i)(0)) & ":" & CLng(ranges(i)(1))).Group
    Next i
End Sub

Private Sub CopyOutlines_FromSourceToOut(ByVal wsSrc As Worksheet, ByVal wsDst As Worksheet)
    If wsSrc Is Nothing Or wsDst Is Nothing Then Exit Sub
    Dim lastR As Long, r As Long, lastC As Long, c As Long
    lastR = GetLastUsedRowSafe(wsSrc)
    lastC = GetLastUsedColSafe(wsSrc)
    If lastR > GetLastUsedRowSafe(wsDst) Then lastR = GetLastUsedRowSafe(wsDst)
    On Error Resume Next
    wsDst.Outline.SummaryRow = wsSrc.Outline.SummaryRow
    For r = 1 To lastR
        wsDst.Rows(r).OutlineLevel = wsSrc.Rows(r).OutlineLevel
        wsDst.Rows(r).Hidden = wsSrc.Rows(r).Hidden
    Next r
    For c = 1 To lastC
        wsDst.Columns(c).OutlineLevel = wsSrc.Columns(c).OutlineLevel
    Next c
    On Error GoTo 0
End Sub

Private Sub EnsureCDColumnsGroupedCollapsed(ByVal wb As Workbook, ByVal sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ws.Columns("C:D").Ungroup
    ws.Columns("C:D").Group
    ws.Outline.ShowLevels ColumnLevels:=1
    On Error GoTo 0
End Sub

Public Sub CollapseAllGroupedSheetsToLevel1(ByVal wb As Workbook)
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        On Error Resume Next
        ws.Outline.ShowLevels RowLevels:=1
        On Error GoTo 0
    Next ws
End Sub

Public Sub EnsureWorkingSheetsHidden(ByVal hideWorkSheets As Boolean)
    On Error Resume Next
    If hideWorkSheets Then
        ThisWorkbook.Worksheets(SH_BG).Visible  = xlVeryHidden
        ThisWorkbook.Worksheets(SH_BS).Visible  = xlVeryHidden
        ThisWorkbook.Worksheets(SH_MAP).Visible = xlVeryHidden
        ThisWorkbook.Worksheets(SH_HOME).Visible = xlSheetVisible
    Else
        ThisWorkbook.Worksheets(SH_BG).Visible  = xlSheetVisible
        ThisWorkbook.Worksheets(SH_BS).Visible  = xlSheetVisible
        ThisWorkbook.Worksheets(SH_MAP).Visible = xlSheetVisible
        ThisWorkbook.Worksheets(SH_HOME).Visible = xlSheetVisible
    End If
    On Error GoTo 0
End Sub

Public Sub SetGridlinesOff_AllSheetsViews(ByVal wb As Workbook)
    Dim w As Window, ws As Worksheet
    On Error Resume Next
    For Each w In wb.Windows
        w.DisplayGridlines = False
        For Each ws In wb.Worksheets
            ws.Activate
            w.DisplayGridlines = False
        Next ws
    Next w
    On Error GoTo 0
End Sub

Public Sub HideSheetByNameCI(ByVal wbOut As Workbook, ByVal targetName As String)
    Dim ws As Worksheet
    For Each ws In wbOut.Worksheets
        If LCase$(ws.Name) = LCase$(targetName) Then
            ws.Visible = xlSheetHidden
            Exit Sub
        End If
    Next ws
End Sub

Private Function GetExportModeFromFrmLeadMeta() As eExportMode
    If gExportMode = 0 Then gExportMode = emFS
    GetExportModeFromFrmLeadMeta = gExportMode
End Function
