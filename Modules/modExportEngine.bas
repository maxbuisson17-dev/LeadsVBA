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
    Dim errNum As Long, errDesc As String
    Dim firstCreated As Boolean
    Dim exportMode As eExportMode
    Dim oldAlerts As Boolean, oldScreen As Boolean

    oldAlerts = Application.DisplayAlerts
    oldScreen = Application.ScreenUpdating
    gLastExportSucceeded = False
    Set gLastExportedWorkbook = Nothing
    modKETrace.LogKE "START | gExportPDF=" & CStr(gExportPDF) & " | gBalancePath=" & gBalancePath, "ExportValuesCopy"

    ' --- Chemin de sauvegarde ---
    folder = GetFolderFromPath(gBalancePath)
    If Len(Trim$(folder)) = 0 Then folder = ThisWorkbook.Path & Application.PathSeparator
    If Right$(folder, 1) <> Application.PathSeparator Then folder = folder & Application.PathSeparator
    defaultName = BuildDefaultFileName(gClient, gExercice, gVersion)
    If Len(defaultName) > 180 Then defaultName = Left$(defaultName, 180)

    pickedPath = PromptSaveAsPath_NoUI(folder & defaultName)
    If VarType(pickedPath) = vbBoolean Then
        modKETrace.LogKE "PromptSaveAsPath_NoUI returned Boolean=False (cancel)", "ExportValuesCopy"
        If pickedPath = False Then GoTo CLEAN_EXIT
    Else
        modKETrace.LogKE "PromptSaveAsPath_NoUI returned path=" & CStr(pickedPath), "ExportValuesCopy"
    End If

    finalPath = CStr(pickedPath)
    If LCase(Right(finalPath, 5)) <> ".xlsx" Then finalPath = finalPath & ".xlsx"
    modKETrace.LogKE "finalPath=" & finalPath, "ExportValuesCopy"

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
    modKETrace.LogKE "Before Workbooks.Add", "ExportValuesCopy"
    Set wbOut = Workbooks.Add(xlWBATWorksheet)
    wbOut.Worksheets(1).Name = "TMP_DELETE"
    modKETrace.LogKE "wbOut created | Name=" & wbOut.Name, "ExportValuesCopy"

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

    ' --- Nettoyer onglets exclus (centralise via IsExcludedFromExport) ---
    Application.DisplayAlerts = False
    Dim wsClean As Worksheet
    For Each wsClean In wbOut.Worksheets
        If IsExcludedFromExport(UCase$(Trim$(wsClean.Name))) Then
            On Error Resume Next
            wsClean.Delete
            On Error GoTo 0
        End If
    Next wsClean
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

    ' --- Marquer les sections BS / BS_detail et neutraliser les sauts de page visibles ---
    RefreshBSMarkersAndPageBreaks wbOut

    ' --- Masquer grilles et finaliser affichage ---
    SetGridlinesOff_AllSheetsViews wbOut
    NormalizeAllSheetViews wbOut

    ' --- Activer BS comme onglet initial (sauvegarde avec BS actif -> ouverture sur BS) ---
    On Error Resume Next
    wbOut.Worksheets(SH_BS).Activate
    On Error GoTo EH

    ' --- Enregistrer ---
    modKETrace.LogKE "Before SaveAs | finalPath=" & finalPath, "ExportValuesCopy"
    Application.DisplayAlerts = True
    wbOut.SaveAs Filename:=finalPath, FileFormat:=xlOpenXMLWorkbook, Local:=True
    Application.DisplayAlerts = False
    Set gLastExportedWorkbook = wbOut
    gLastExportSucceeded = True
    modKETrace.LogKE "After SaveAs | SavedPath=" & finalPath, "ExportValuesCopy"

    ' --- Export PDF si demande ---
    If gExportPDF Then
        modKETrace.LogKE "Before ExportPDF_FromWbOut", "ExportValuesCopy"
        ExportPDF_FromWbOut wbOut, finalPath
        modKETrace.LogKE "After ExportPDF_FromWbOut", "ExportValuesCopy"
    End If

    ' --- Activer BS et ajuster la vue du wbOut en sortie ---
    ' ScreenUpdating doit etre True avant Activate pour que l'affichage bascule
    ' effectivement sur wbOut. Si on attend CLEAN_EXIT, Excel rafraichit sur
    ' la derniere fenetre visible (ThisWorkbook) et non sur wbOut.
    Application.DisplayAlerts  = oldAlerts
    Application.ScreenUpdating = oldScreen
    ActivateWorkbookOnBSView wbOut

CLEAN_EXIT:
    Application.DisplayAlerts  = oldAlerts
    Application.ScreenUpdating = oldScreen
    Exit Sub
EH:
    errNum = Err.Number
    errDesc = Err.Description
    modKETrace.LogKE "ERROR " & errNum & " : " & errDesc, "ExportValuesCopy"
    If Len(Trim$(errDesc)) > 0 Then
        MsgBox "Erreur Export : " & errNum & vbCrLf & errDesc, vbCritical
    End If
    Err.Clear
    Resume CLEAN_EXIT
End Sub

Public Sub ActivateWorkbookOnBSView(ByVal wbOut As Workbook)
    Dim wsBS As Worksheet
    Dim ws As Worksheet

    If wbOut Is Nothing Then Exit Sub

    On Error Resume Next
    Set wsBS = wbOut.Worksheets(SH_BS)
    If wsBS Is Nothing Then Exit Sub

    wbOut.Activate
    For Each ws In wbOut.Worksheets
        If ws.Visible = xlSheetVisible Then
            ws.Activate
            ws.DisplayPageBreaks = False
            If Not ActiveWindow Is Nothing Then
                ActiveWindow.View = xlNormalView
                ActiveWindow.DisplayGridlines = False
                ActiveWindow.ScrollRow = 1
                ActiveWindow.ScrollColumn = 1
            End If
            ws.Range("A1").Select
        End If
    Next ws

    wbOut.Activate
    wsBS.Activate
    wsBS.DisplayPageBreaks = False
    wsBS.Range("B2").Select

    If wbOut.Windows.Count > 0 Then
        With wbOut.Windows(1)
            .View = xlNormalView
            .Zoom = 75
            .DisplayGridlines = False
            .ScrollRow = 1
            .ScrollColumn = 1
        End With
    End If
    On Error GoTo 0
End Sub

' ============================================================
' EXPORT PDF
' Appele si gExportPDF = True.
' Feuilles exportees : SOMMAIRE (en premier si visible), BS, BS_detail,
' et onglets GestTable visibles (SIG, SIG_detail, CAF, BFR, TFT).
' Mise en page : A4 portrait, 1 page de large, zone depuis debut UsedRange jusqu a
'                la derniere colonne contenant des valeurs numeriques.
' Nom PDF : "[nom xlsx] - pdf.pdf" dans le meme dossier.
' ============================================================
Public Sub ExportPDF_FromWbOut(ByVal wbOut As Workbook, ByVal xlsxPath As String)
    Dim pdfPath       As String
    Dim delErr        As String
    Dim ws            As Worksheet
    Dim nm            As String
    Dim names()       As String
    Dim selectNames   As Variant
    Dim printableList As String
    Dim count         As Long
    Dim i             As Long
    Dim pageStartMap  As Object
    Dim currentPdfPage As Long
    Dim pageCountForSheet As Long
    Dim startCol      As Long
    Dim lastNumCol    As Long
    Dim firstPrintRow As Long
    Dim lastPrintRow  As Long
    Dim pArea         As String
    Dim expErr        As Long
    Dim expDesc       As String
    Dim setupErr      As Long
    Dim setupDesc     As String
    Dim errNum        As Long
    Dim errDesc       As String

    On Error GoTo EH
    Set pageStartMap = CreateObject("Scripting.Dictionary")

    ' Chemin PDF : meme dossier, meme nom de base
    If LCase$(Right$(xlsxPath, 5)) = ".xlsx" Then
        pdfPath = Left$(xlsxPath, Len(xlsxPath) - 5) & ".pdf"
    Else
        pdfPath = xlsxPath & ".pdf"
    End If
    modKETrace.LogKE "START | xlsxPath=" & xlsxPath & " | pdfPath=" & pdfPath, "ExportPDF_FromWbOut"

    If FileExists(pdfPath) Then
        If Not TryDeleteFile(pdfPath, delErr) Then
            Err.Raise vbObjectError + 1400, "ExportPDF_FromWbOut", "Impossible de supprimer le PDF existant : " & delErr
        End If
    End If

    ' Collecter SOMMAIRE, BS, BS_detail et GestTable visibles
    count = 0
    For Each ws In wbOut.Worksheets
        If ws.Visible = xlSheetVisible Then
            nm = LCase$(Trim$(ws.Name))
            If nm = "sommaire" Or nm = "bs" Or nm = "bs_detail" Or IsGestTableSheet(ws.Name) Then
                count = count + 1
            End If
        End If
    Next ws
    modKETrace.LogKE "Printable sheet count=" & CStr(count), "ExportPDF_FromWbOut"
    If count = 0 Then
        Err.Raise vbObjectError + 1401, "ExportPDF_FromWbOut", "Aucune feuille imprimable trouvee dans wbOut."
    End If

    ReDim names(0 To count - 1)
    ReDim selectNames(0 To count - 1)
    i = 0

    ' SOMMAIRE doit sortir en premiere page si l'onglet existe et est visible.
    For Each ws In wbOut.Worksheets
        If ws.Visible = xlSheetVisible Then
            nm = LCase$(Trim$(ws.Name))
            If nm = "sommaire" Then
                names(i) = ws.Name
                selectNames(i) = ws.Name
                If Len(printableList) > 0 Then printableList = printableList & ","
                printableList = printableList & ws.Name
                i = i + 1
                Exit For
            End If
        End If
    Next ws

    For Each ws In wbOut.Worksheets
        If ws.Visible = xlSheetVisible Then
            nm = LCase$(Trim$(ws.Name))
            If nm <> "sommaire" Then
                If nm = "bs" Or nm = "bs_detail" Or IsGestTableSheet(ws.Name) Then
                    names(i) = ws.Name
                    selectNames(i) = ws.Name
                    If Len(printableList) > 0 Then printableList = printableList & ","
                    printableList = printableList & ws.Name
                    i = i + 1
                End If
            End If
        End If
    Next ws
    modKETrace.LogKE "Printable sheets=[" & printableList & "]", "ExportPDF_FromWbOut"

    ' Mise en page (PrintCommunication=False evite erreurs driver pendant PageSetup)
    On Error Resume Next
    Application.PrintCommunication = False
    On Error GoTo EH
    currentPdfPage = 1

    For i = 0 To UBound(names)
        Set ws = wbOut.Worksheets(names(i))
        modKETrace.LogKE "PageSetup start | Sheet=" & ws.Name, "ExportPDF_FromWbOut"

        ws.Activate
        modKETrace.LogKE "Activate done | Sheet=" & ws.Name, "ExportPDF_FromWbOut"
        On Error Resume Next
        ws.Outline.ShowLevels RowLevels:=8, ColumnLevels:=8
        If Err.Number <> 0 Then
            setupErr = Err.Number
            setupDesc = Err.Description
            On Error GoTo EH
            modKETrace.LogKE "ShowLevels error | Sheet=" & ws.Name & " | Err=" & CStr(setupErr) & " | Desc=" & setupDesc, "ExportPDF_FromWbOut"
            Err.Raise setupErr, "ExportPDF_FromWbOut", "ShowLevels sur '" & ws.Name & "' : " & setupDesc
        End If
        On Error GoTo EH
        modKETrace.LogKE "ShowLevels done | Sheet=" & ws.Name, "ExportPDF_FromWbOut"

        If LCase$(ws.Name) = "bs" Or LCase$(ws.Name) = "bs_detail" Then
            On Error Resume Next
            ws.Columns("C:D").Hidden = True
            If Err.Number <> 0 Then
                setupErr = Err.Number
                setupDesc = Err.Description
                On Error GoTo EH
                modKETrace.LogKE "Hide C:D error | Sheet=" & ws.Name & " | Err=" & CStr(setupErr) & " | Desc=" & setupDesc, "ExportPDF_FromWbOut"
                Err.Raise setupErr, "ExportPDF_FromWbOut", "Hide C:D sur '" & ws.Name & "' : " & setupDesc
            End If
            On Error GoTo EH
            modKETrace.LogKE "Hide C:D done | Sheet=" & ws.Name, "ExportPDF_FromWbOut"
        End If

        startCol = 2
        modKETrace.LogKE "PDF startCol forced to B | Sheet=" & ws.Name, "ExportPDF_FromWbOut"
        lastNumCol = PDF_FindLastPrintableCol(ws)
        modKETrace.LogKE "lastPrintableCol=" & CStr(lastNumCol) & " | Sheet=" & ws.Name, "ExportPDF_FromWbOut"
        If lastNumCol < startCol Then lastNumCol = startCol
        firstPrintRow = PDF_FindFirstPrintableRow(ws, startCol, lastNumCol)
        lastPrintRow = PDF_FindLastPrintableRow(ws, startCol, lastNumCol)
        If firstPrintRow <= 0 Then firstPrintRow = 1
        If lastPrintRow < firstPrintRow Then lastPrintRow = firstPrintRow
        pArea = PDF_ColLtr(ws, startCol) & CStr(firstPrintRow) & ":" & PDF_ColLtr(ws, lastNumCol) & CStr(lastPrintRow)
        modKETrace.LogKE "PageSetup area | Sheet=" & ws.Name & " | PrintArea=" & pArea, "ExportPDF_FromWbOut"

        On Error Resume Next
        Err.Clear
        ws.PageSetup.Zoom = False
        With ws.PageSetup
            .PrintArea = pArea
            .Orientation = xlPortrait
            .PaperSize = xlPaperA4
            .FirstPageNumber = currentPdfPage
            If LCase$(Trim$(ws.Name)) = "sommaire" Then
                .LeftMargin = Application.CentimetersToPoints(2.5)
                .RightMargin = Application.CentimetersToPoints(2.5)
                .TopMargin = Application.CentimetersToPoints(4)
                .BottomMargin = Application.CentimetersToPoints(4)
                .CenterHorizontally = True
                .CenterVertically = True
            Else
                .LeftMargin = Application.CentimetersToPoints(1)
                .RightMargin = Application.CentimetersToPoints(1)
                .TopMargin = Application.CentimetersToPoints(1.5)
                .BottomMargin = Application.CentimetersToPoints(1.5)
                .CenterHorizontally = True
                .CenterVertically = False
            End If
            .HeaderMargin = Application.CentimetersToPoints(0.5)
            .FooterMargin = Application.CentimetersToPoints(0.5)
            .FitToPagesWide = 1
            .FitToPagesTall = 0
        End With
        setupErr = Err.Number
        setupDesc = Err.Description
        On Error GoTo EH

        If setupErr <> 0 Then
            modKETrace.LogKE "PageSetup error | Sheet=" & ws.Name & " | Err=" & CStr(setupErr) & " | Desc=" & setupDesc, "ExportPDF_FromWbOut"
            Err.Raise setupErr, "ExportPDF_FromWbOut", "PageSetup sur '" & ws.Name & "' : " & setupDesc
        End If

        pageStartMap(ws.Name) = currentPdfPage
        pageCountForSheet = GetPdfPageCountForSheet(ws)
        If pageCountForSheet < 1 Then pageCountForSheet = 1
        currentPdfPage = currentPdfPage + pageCountForSheet
        modKETrace.LogKE "PageSetup done | Sheet=" & ws.Name, "ExportPDF_FromWbOut"
        Set ws = Nothing
    Next i

    On Error Resume Next
    Application.PrintCommunication = True
    On Error GoTo EH

    On Error Resume Next
    Application.PrintCommunication = False
    On Error GoTo EH
    For i = 0 To UBound(names)
        Set ws = wbOut.Worksheets(names(i))
        ApplyHeaderBySheet ws, CLng(pageStartMap(ws.Name))
    Next i
    On Error Resume Next
    Application.PrintCommunication = True
    On Error GoTo EH

    modKETrace.LogKE "Before Sheets(selectNames).Select", "ExportPDF_FromWbOut"
    wbOut.Sheets(selectNames).Select

    On Error Resume Next
    Err.Clear
    modKETrace.LogKE "Before ExportAsFixedFormat", "ExportPDF_FromWbOut"
    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=pdfPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
    expErr = Err.Number
    expDesc = Err.Description
    On Error GoTo EH
    modKETrace.LogKE "After ExportAsFixedFormat | Err=" & CStr(expErr) & " | Desc=" & expDesc & " | Exists=" & CStr(FileExists(pdfPath)), "ExportPDF_FromWbOut"

    If Not FileExists(pdfPath) Then
        If expErr <> 0 Then
            Err.Raise expErr, "ExportPDF_FromWbOut", expDesc
        Else
            Err.Raise vbObjectError + 1402, "ExportPDF_FromWbOut", "ExportAsFixedFormat n'a retourne aucune erreur, mais aucun PDF n'a ete cree."
        End If
    End If

    ' Restaurer mise en page
    On Error Resume Next
    For i = 0 To UBound(names)
        Set ws = wbOut.Worksheets(names(i))
        If Not ws Is Nothing Then
            ws.PageSetup.PrintArea = ""
            ws.PageSetup.Zoom = 100
            ws.Outline.ShowLevels RowLevels:=1, ColumnLevels:=1
            If LCase$(ws.Name) = "bs" Or LCase$(ws.Name) = "bs_detail" Then
                ws.Columns("C:D").Hidden = True
            End If
        End If
        Set ws = Nothing
    Next i
    On Error GoTo EH

    wbOut.Worksheets(names(0)).Select
    modKETrace.LogKE "SUCCESS | pdfPath=" & pdfPath, "ExportPDF_FromWbOut"
    Exit Sub
EH:
    errNum = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    Application.PrintCommunication = True
    modKETrace.LogKE "ERROR PDF " & errNum & " : " & errDesc, "ExportPDF_FromWbOut"
    On Error GoTo 0
    MsgBox "Erreur export PDF (" & errNum & ") :" & vbCrLf & errDesc, vbCritical
End Sub

' Detecte la derniere colonne contenant au moins une donnee non vide
' en ignorant la colonne A pour la zone d'impression PDF.
Private Function PDF_FindLastPrintableCol(ByVal ws As Worksheet) As Long
    Dim lastC As Long, lastR As Long
    Dim c As Long, r As Long
    Dim v As Variant

    lastC = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    lastR = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If lastR < 1 Then lastR = 1
    If lastC < 2 Then
        PDF_FindLastPrintableCol = 2
        Exit Function
    End If

    For c = lastC To 2 Step -1
        For r = 1 To lastR
            v = ws.Cells(r, c).Value
            If Not IsError(v) Then
                If Not IsEmpty(v) Then
                    If Len(Trim$(CStr(v))) > 0 Then
                        PDF_FindLastPrintableCol = c
                        Exit Function
                    End If
                End If
            End If
        Next r
    Next c

    PDF_FindLastPrintableCol = 2
End Function

Private Function PDF_FindFirstPrintableRow(ByVal ws As Worksheet, ByVal startCol As Long, ByVal endCol As Long) As Long
    Dim firstR As Long
    Dim lastR As Long
    Dim r As Long
    Dim c As Long

    firstR = ws.UsedRange.Row
    lastR = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If firstR < 1 Then firstR = 1
    If lastR < firstR Then lastR = firstR

    For r = firstR To lastR
        For c = startCol To endCol
            If PDF_CellHasPrintableValue(ws.Cells(r, c).Value) Then
                PDF_FindFirstPrintableRow = r
                Exit Function
            End If
        Next c
    Next r

    PDF_FindFirstPrintableRow = firstR
End Function

Private Function PDF_FindLastPrintableRow(ByVal ws As Worksheet, ByVal startCol As Long, ByVal endCol As Long) As Long
    Dim firstR As Long
    Dim lastR As Long
    Dim r As Long
    Dim c As Long

    firstR = ws.UsedRange.Row
    lastR = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If firstR < 1 Then firstR = 1
    If lastR < firstR Then lastR = firstR

    For r = lastR To firstR Step -1
        For c = startCol To endCol
            If PDF_CellHasPrintableValue(ws.Cells(r, c).Value) Then
                PDF_FindLastPrintableRow = r
                Exit Function
            End If
        Next c
    Next r

    PDF_FindLastPrintableRow = firstR
End Function

Private Function PDF_CellHasPrintableValue(ByVal v As Variant) As Boolean
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    PDF_CellHasPrintableValue = (Len(Trim$(CStr(v))) > 0)
End Function

' Convertit un index de colonne en lettre(s) (ex. 8 -> "H").
Private Function PDF_ColLtr(ByVal ws As Worksheet, ByVal colIdx As Long) As String
    Dim addr As String
    addr = ws.Cells(1, colIdx).Address  ' ex: "$H$1"
    PDF_ColLtr = Mid$(addr, 2, InStr(2, addr, "$") - 2)
End Function

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

' Liste centralisee des onglets internes jamais exportes vers wbOut.
' C'est ici et uniquement ici qu'on ajoute/retire un onglet exclu.
Private Function IsExcludedFromExport(ByVal nm As String) As Boolean
    Select Case nm
        Case "LEADS", "PARAM", "MAPPING", "ACCUEIL", "CENTRAL", "CONSOL"
            IsExcludedFromExport = True
    End Select
End Function

Private Function ShouldExportSheet(ByVal sheetName As String, ByVal mode As eExportMode) As Boolean
    Dim nm As String
    nm = UCase$(Trim$(sheetName))
    If IsExcludedFromExport(nm) Then Exit Function
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

Private Sub RefreshBSMarkersAndPageBreaks(ByVal wbOut As Workbook)
    Dim unitLabel As String

    unitLabel = GetBSMarkerUnitLabel()
    ApplyBSMarkersAndPageBreaksToSheet wbOut, SH_BS, unitLabel
    ApplyBSMarkersAndPageBreaksToSheet wbOut, "BS_detail", unitLabel
End Sub

Private Sub ApplyBSMarkersAndPageBreaksToSheet(ByVal wbOut As Workbook, ByVal sheetName As String, ByVal unitLabel As String)
    Dim ws As Worksheet
    Dim markers As Range

    On Error Resume Next
    Set ws = wbOut.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Set markers = BuildBSMarkerRange(ws, unitLabel)
    DeleteSheetScopedNameIfExists ws, "BSMarker"
    If markers Is Nothing Then
        modKETrace.LogKE "BSMarker introuvable", "ApplyBSMarkersAndPageBreaksToSheet", ws.Name, wbOut.Name
        Exit Sub
    End If

    ws.Names.Add Name:="BSMarker", RefersTo:="=" & markers.Address(True, True, xlA1, True)
    ApplyBSPageBreaksFromMarker ws, markers
    ws.DisplayPageBreaks = False
    modKETrace.LogKE "BSMarker=" & markers.Address(False, False), "ApplyBSMarkersAndPageBreaksToSheet", ws.Name, wbOut.Name
End Sub

Private Function GetBSMarkerUnitLabel() As String
    If gGenerateInKE Then
        GetBSMarkerUnitLabel = "en KE"
    Else
        GetBSMarkerUnitLabel = "en euros"
    End If
End Function

Private Function BuildBSMarkerRange(ByVal ws As Worksheet, ByVal unitLabel As String) As Range
    Dim markerLabels As Variant
    Dim i As Long
    Dim foundCell As Range
    Dim markers As Range
    Dim seen As Object

    Set seen = CreateObject("Scripting.Dictionary")
    markerLabels = Array("ACTIF", "PASSIF", "COMPTE DE RESULTAT")

    For i = LBound(markerLabels) To UBound(markerLabels)
        Set foundCell = FindBSMarkerCell(ws, CStr(markerLabels(i)), unitLabel)
        If Not foundCell Is Nothing Then
            If Not seen.Exists(foundCell.Address(True, True, xlA1)) Then
                seen.Add foundCell.Address(True, True, xlA1), True
                If markers Is Nothing Then
                    Set markers = foundCell
                Else
                    Set markers = Union(markers, foundCell)
                End If
            End If
        End If
    Next i

    Set BuildBSMarkerRange = markers
End Function

Private Function FindBSMarkerCell(ByVal ws As Worksheet, ByVal markerLabel As String, ByVal unitLabel As String) As Range
    Dim cell As Range
    Dim foundCell As Range
    Dim normalizedMarker As String
    Dim normalizedTarget As String
    Dim normalizedFormula As String
    Dim normalizedValue As String

    If ws Is Nothing Then Exit Function
    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then Exit Function

    normalizedMarker = NormalizeBSMarkerText(markerLabel)
    normalizedTarget = NormalizeBSMarkerText(markerLabel & " - " & unitLabel)
    For Each cell In ws.UsedRange.Cells
        normalizedValue = NormalizeBSMarkerText(CStr(cell.Value))
        If normalizedValue = normalizedTarget _
           Or normalizedValue = normalizedMarker _
           Or Left$(normalizedValue, Len(normalizedMarker & " - ")) = normalizedMarker & " - " Then
            Set foundCell = cell
            Exit For
        End If

        normalizedFormula = NormalizeBSMarkerText(CStr(cell.Formula))
        If Len(normalizedFormula) > 0 Then
            If InStr(1, normalizedFormula, normalizedMarker, vbTextCompare) > 0 _
               And InStr(1, normalizedFormula, "PARAM!", vbTextCompare) > 0 _
               And InStr(1, normalizedFormula, "B7", vbTextCompare) > 0 Then
                Set foundCell = cell
                Exit For
            End If
        End If
    Next cell

    Set FindBSMarkerCell = foundCell
End Function

Private Function NormalizeBSMarkerText(ByVal rawText As String) As String
    Dim txt As String

    txt = UCase$(Trim$(rawText))
    txt = Replace(txt, vbCr, " ")
    txt = Replace(txt, vbLf, " ")
    txt = Replace(txt, "'", "")
    txt = Replace(txt, ChrW$(160), " ")
    txt = Replace(txt, ChrW$(8239), " ")
    txt = Replace(txt, ChrW$(8364), "E")
    txt = Replace(txt, ChrW$(201), "E")
    txt = Replace(txt, ChrW$(200), "E")
    txt = Replace(txt, ChrW$(202), "E")
    txt = Replace(txt, ChrW$(203), "E")
    txt = Replace(txt, ChrW$(192), "A")
    txt = Replace(txt, ChrW$(194), "A")
    txt = Replace(txt, ChrW$(196), "A")
    txt = Replace(txt, ChrW$(206), "I")
    txt = Replace(txt, ChrW$(207), "I")
    txt = Replace(txt, ChrW$(212), "O")
    txt = Replace(txt, ChrW$(214), "O")
    txt = Replace(txt, ChrW$(217), "U")
    txt = Replace(txt, ChrW$(219), "U")
    txt = Replace(txt, ChrW$(220), "U")
    txt = Replace(txt, ChrW$(199), "C")

    Do While InStr(txt, "  ") > 0
        txt = Replace(txt, "  ", " ")
    Loop

    NormalizeBSMarkerText = txt
End Function

Private Sub DeleteSheetScopedNameIfExists(ByVal ws As Worksheet, ByVal localName As String)
    On Error Resume Next
    ws.Names(localName).Delete
    On Error GoTo 0
End Sub

Private Sub ApplyHeaderBySheet(ByVal ws As Worksheet, Optional ByVal firstPdfPageNumber As Long = 1)
    Dim rightHeaderText As String
    Dim label           As String

    On Error Resume Next
    If ws Is Nothing Then Exit Sub

    label   = EscapeHeaderText(GetHeaderDisplayName(ws.Name))

    With ws.PageSetup
        .DifferentFirstPageHeaderFooter = False
        .OddAndEvenPagesHeaderFooter    = False
        .LeftHeader   = vbNullString
        .CenterHeader = vbNullString
        .RightHeader  = vbNullString
        .EvenPage.LeftHeader.Text = vbNullString
        .EvenPage.CenterHeader.Text = vbNullString
        .EvenPage.RightHeader.Text = vbNullString
        .FirstPage.LeftHeader.Text = vbNullString
        .FirstPage.CenterHeader.Text = vbNullString
        .FirstPage.RightHeader.Text = vbNullString
        .LeftFooter   = vbNullString
        .CenterFooter = vbNullString
        .RightFooter  = vbNullString

        Select Case UCase$(Trim$(ws.Name))
            Case "BS"
                .DifferentFirstPageHeaderFooter = True
                .OddAndEvenPagesHeaderFooter = True
                ConfigureThreePageHeaders ws.PageSetup, firstPdfPageNumber, "ACTIF", "PASSIF", "COMPTE DE RESULTAT"
            Case "BS_DETAIL"
                .DifferentFirstPageHeaderFooter = True
                .OddAndEvenPagesHeaderFooter = True
                ConfigureThreePageHeaders ws.PageSetup, firstPdfPageNumber, _
                    "ACTIF d" & ChrW$(233) & "taill" & ChrW$(233), _
                    "PASSIF d" & ChrW$(233) & "taill" & ChrW$(233), _
                    "COMPTE DE RESULTAT d" & ChrW$(233) & "taill" & ChrW$(233)
            Case Else
                rightHeaderText = BuildRightHeaderText(label)
                .RightHeader = rightHeaderText
        End Select
    End With
    On Error GoTo 0
End Sub

Private Function BuildRightHeaderText(ByVal label As String) As String
    Dim client As String
    Dim exoDate As String
    Dim metaLine As String

    client = Trim$(EscapeHeaderText(gClient))
    exoDate = EscapeHeaderText(GetFormattedExerciceForHeader())

    If Len(client) > 0 Then
        metaLine = client & " - " & exoDate
    Else
        metaLine = exoDate
    End If

    BuildRightHeaderText = "&16&B" & EscapeHeaderText(label) & "&B" & Chr$(10) & "&12 " & metaLine
End Function

Private Sub ConfigureThreePageHeaders(ByVal ps As PageSetup, ByVal firstPdfPageNumber As Long, ByVal firstLabel As String, ByVal secondLabel As String, ByVal thirdLabel As String)
    If ps Is Nothing Then Exit Sub

    ps.FirstPage.RightHeader.Text = BuildRightHeaderText(firstLabel)

    If (firstPdfPageNumber + 1) Mod 2 = 0 Then
        ps.EvenPage.RightHeader.Text = BuildRightHeaderText(secondLabel)
        ps.RightHeader = BuildRightHeaderText(thirdLabel)
    Else
        ps.RightHeader = BuildRightHeaderText(secondLabel)
        ps.EvenPage.RightHeader.Text = BuildRightHeaderText(thirdLabel)
    End If
End Sub

Private Function GetFormattedExerciceForHeader() As String
    Dim dExo As Date

    If IsValidExerciceDate_UI(gExercice, dExo) Then
        GetFormattedExerciceForHeader = Format$(dExo, "dd/mm/yyyy")
    Else
        GetFormattedExerciceForHeader = gExercice
    End If
End Function

Private Function GetHeaderDisplayName(ByVal sheetName As String) As String
    Select Case UCase$(Trim$(sheetName))
        Case "SOMMAIRE"
            GetHeaderDisplayName = "SOMMAIRE"
        Case "BS"
            GetHeaderDisplayName = "ACTIF"
        Case "BS_DETAIL"
            GetHeaderDisplayName = "ACTIF d" & ChrW$(233) & "taill" & ChrW$(233)
        Case "SIG"
            GetHeaderDisplayName = "Soldes interm" & ChrW$(233) & "diaires de gestion"
        Case "SIG_DETAIL"
            GetHeaderDisplayName = "Soldes interm" & ChrW$(233) & "diaires de gestion d" & ChrW$(233) & "taill" & ChrW$(233)
        Case "CAF"
            GetHeaderDisplayName = "Capacit" & ChrW$(233) & " d'autofinancement"
        Case "BFR"
            GetHeaderDisplayName = "Besoin en fond de roulement"
        Case "TFT"
            GetHeaderDisplayName = "Tableau de flux de tr" & ChrW$(233) & "sorerie"
        Case Else
            GetHeaderDisplayName = Replace(sheetName, "_detail", " d" & ChrW$(233) & "taill" & ChrW$(233))
    End Select
End Function

Private Function EscapeHeaderText(ByVal rawText As String) As String
    EscapeHeaderText = Replace(rawText, "&", "&&")
End Function

Private Function GetPdfPageCountForSheet(ByVal ws As Worksheet) As Long
    On Error Resume Next
    If ws Is Nothing Then Exit Function
    GetPdfPageCountForSheet = ws.HPageBreaks.Count + 1
    If Err.Number <> 0 Or GetPdfPageCountForSheet < 1 Then
        Err.Clear
        GetPdfPageCountForSheet = 1
    End If
    On Error GoTo 0
End Function

Private Sub ApplyBSPageBreaksFromMarker(ByVal ws As Worksheet, ByVal markers As Range)
    Dim rowMap As Object
    Dim cell As Range
    Dim rows() As Long
    Dim itemRows As Variant
    Dim i As Long
    Dim j As Long
    Dim tmp As Long

    If ws Is Nothing Or markers Is Nothing Then Exit Sub

    Set rowMap = CreateObject("Scripting.Dictionary")
    For Each cell In markers.Cells
        If Not rowMap.Exists(CStr(cell.Row)) Then rowMap.Add CStr(cell.Row), CLng(cell.Row)
    Next cell

    If rowMap.Count = 0 Then Exit Sub

    ws.ResetAllPageBreaks
    ReDim rows(1 To rowMap.Count)
    itemRows = rowMap.Items
    For i = 0 To rowMap.Count - 1
        rows(i + 1) = CLng(itemRows(i))
    Next i

    For i = LBound(rows) To UBound(rows) - 1
        For j = i + 1 To UBound(rows)
            If rows(j) < rows(i) Then
                tmp = rows(i)
                rows(i) = rows(j)
                rows(j) = tmp
            End If
        Next j
    Next i

    ' Le premier marker ouvre deja la premiere page.
    For i = LBound(rows) + 1 To UBound(rows)
        If rows(i) > 1 Then ws.HPageBreaks.Add Before:=ws.Rows(rows(i))
    Next i
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

Private Sub NormalizeAllSheetViews(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim w As Window

    If wb Is Nothing Then Exit Sub

    On Error Resume Next
    For Each w In wb.Windows
        w.View = xlNormalView
        w.DisplayGridlines = False
    Next w

    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            ws.Activate
            ws.DisplayPageBreaks = False
            If ActiveWindow Is Nothing Then GoTo NextSheet
            ActiveWindow.View = xlNormalView
            ActiveWindow.DisplayGridlines = False
            ws.Range("A1").Select
            ActiveWindow.ScrollRow = 1
            ActiveWindow.ScrollColumn = 1
        End If
NextSheet:
    Next ws
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

Public Sub EnsureWorkingSheetsHidden(ByVal bHide As Boolean)
    Dim ws As Worksheet
    Dim shNames As Variant
    Dim i As Long
    shNames = Array(SH_BG, SH_BS, SH_MAP, SH_SIG)
    On Error Resume Next
    For i = LBound(shNames) To UBound(shNames)
        Set ws = Nothing
        Set ws = ThisWorkbook.Worksheets(CStr(shNames(i)))
        If Not ws Is Nothing Then
            ws.Visible = IIf(bHide, xlSheetHidden, xlSheetVisible)
        End If
    Next i
    On Error GoTo 0
End Sub

Private Function GetExportModeFromFrmLeadMeta() As eExportMode
    If gExportMode = 0 Then gExportMode = emFS
    GetExportModeFromFrmLeadMeta = gExportMode
End Function
