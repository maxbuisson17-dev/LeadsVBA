Attribute VB_Name = "modOrchestrator"
' ============================================================
' modOrchestrator.bas - Point d'entree unique LeadsWizard
' Flux YES : balance directe -> CDC -> frmLeadMeta -> Generation
' Flux NO  : frmImportBalanceV6 -> CDC -> frmLeadMeta -> Generation
' ============================================================
Option Explicit

' ============================================================
' ENTRY POINT
' ============================================================
Public Sub Importer_BG_V4()
    ' Desactiver tous les add-ins au lancement (evite toute interference externe)
    DisableAllAddins

    Dim userChoice As VbMsgBoxResult
    userChoice = MsgBox("Importer une balance comparative N / N-1 ?", _
                        vbQuestion + vbYesNoCancel, "Import balance")
    Select Case userChoice
        Case vbYes
            gEntryPointWasYes = True
            RunYesFlow
        Case vbNo
            gEntryPointWasYes = False
            Dim ufImp As New frmImportBalanceV6
            ufImp.Show vbModal
            Unload ufImp
        Case Else
            ' Annuler
    End Select
End Sub

' ============================================================
' FLUX YES - balance unique 4 colonnes
' ============================================================
Private Sub RunYesFlow()
    On Error GoTo EH

    ' 1. Selectionner et charger la balance via modImportUnified
    '    gFullData est deja normalise 4 cols par PickAndLoadBalance
    If Not PickAndLoadBalance() Then GoTo CleanExit

    ' 2. Verifier que gFullData est exploitable
    If IsEmpty(gFullData) Or Not IsArray(gFullData) Then
        MsgBox "Impossible de normaliser la balance.", vbExclamation
        GoTo CleanExit
    End If
    gMaxAcctLen = ComputeMaxAccountLen(gFullData)

    ' 3. CDC
    BuildControlReportFromFullData

    ' 4. Navigation selon erreurs CDC
    If gOkToGenerate Then
        ' Importer dans BG source avant frmLeadMeta
        ImportIntoBG_FromFullData
        Dim ufMeta As New frmLeadMeta
        ufMeta.Show vbModal
        Unload ufMeta
    Else
        Dim ufErr As New frmBGError
        ufErr.Show vbModal
        Unload ufErr
    End If

CleanExit:
    Exit Sub
EH:
    MsgBox "Erreur flux OUI : " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

Private Function PickAndLoadBalance() As Boolean
    Dim f As Variant
    Dim info As String
    Dim arrRaw As Variant
    Dim arrNormalized As Variant

    On Error GoTo EH

    With Application.FileDialog(msoFileDialogFilePicker)
        .title = "Selectionner la balance (Compte / Libelle / Solde N / Solde N-1)"
        .Filters.Clear
        .Filters.Add "Fichiers supportes", "*.xlsx;*.xlsm;*.xls;*.csv;*.txt;*.dat"
        .AllowMultiSelect = False
        If Not .Show Then GoTo CleanExit
        f = .SelectedItems(1)
    End With
    gBalancePath = CStr(f)

    ' VBYES : tenter d'abord l'import comparative 4 colonnes.
    arrRaw = modImportUnified.ImportFile_ToBalance4Cols(gBalancePath, info)
    If Not modImportUnified.ImportUnified_ArrayHasRows4Cols(arrRaw) Then
        ' Fallback securise pour formats importables en 3 colonnes.
        arrRaw = modImportUnified.ImportFile_ToBalance3Cols(gBalancePath, info)
    End If

    If Not modImportUnified.ImportUnified_ArrayHasRows(arrRaw) Then
        MsgBox "Import impossible : aucune ligne exploitable." & vbCrLf & info, vbExclamation
        GoTo CleanExit
    End If

    arrNormalized = modLeadsWizard.Ensure4Cols(arrRaw)
    If IsEmpty(arrNormalized) Or Not IsArray(arrNormalized) Then
        MsgBox "Import impossible : normalisation 4 colonnes invalide.", vbExclamation
        GoTo CleanExit
    End If

    ' Affectation explicite avant CDC et import BG.
    gFullData = arrNormalized
    PickAndLoadBalance = True

CleanExit:
    Exit Function
EH:
    Debug.Print "PickAndLoadBalance error " & Err.Number & " : " & Err.Description
    PickAndLoadBalance = False
    Resume CleanExit
End Function

Public Function PickAndLoadPreview() As Boolean
    Dim f As Variant
    Dim info As String
    Dim arrRaw As Variant
    Dim arrPreview As Variant

    On Error GoTo EH

    With Application.FileDialog(msoFileDialogFilePicker)
        .title = "S"&ChrW$(233)&"lectionner un fichier Excel"
        .Filters.Clear
        .Filters.Add "Fichiers Excel", "*.xlsx;*.xlsm;*.xls"
        .AllowMultiSelect = False
        If Not .Show Then GoTo CleanExit
        f = .SelectedItems(1)
    End With

    gBalancePath = CStr(f)

    arrRaw = modImportUnified.ImportFile_ToBalance4Cols(gBalancePath, info)
    If Not modImportUnified.ImportUnified_ArrayHasRows4Cols(arrRaw) Then
        arrRaw = modImportUnified.ImportFile_ToBalance3Cols(gBalancePath, info)
    End If

    If Not modImportUnified.ImportUnified_ArrayHasRows(arrRaw) Then
        MsgBox "Import impossible : aucune ligne exploitable." & vbCrLf & info, vbExclamation
        GoTo CleanExit
    End If

    arrPreview = modLeadsWizard.Ensure4Cols(arrRaw)
    If IsEmpty(arrPreview) Or Not IsArray(arrPreview) Then
        MsgBox "Import impossible : previsualisation invalide.", vbExclamation
        GoTo CleanExit
    End If

    gPreviewData = arrPreview
    PickAndLoadPreview = True

CleanExit:
    Exit Function
EH:
    Debug.Print "PickAndLoadPreview error " & Err.Number & " : " & Err.Description
    PickAndLoadPreview = False
    Resume CleanExit
End Function

Public Function LoadFullDataCustomCols( _
    ByVal skipRows As Long, _
    ByVal colCompte As String, _
    ByVal colLib As String, _
    ByVal colN As String, _
    ByVal colN1 As String) As Boolean

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim arrSrc As Variant
    Dim arrMapped As Variant
    Dim idxCompte As Long
    Dim idxLib As Long
    Dim idxN As Long
    Dim idxN1 As Long

    On Error GoTo EH

    If Len(Trim$(gBalancePath)) = 0 Then GoTo CleanExit

    idxCompte = ColumnLetterToNumber(colCompte)
    idxLib = ColumnLetterToNumber(colLib)
    idxN = ColumnLetterToNumber(colN)
    idxN1 = ColumnLetterToNumber(colN1)

    If idxCompte <= 0 Or idxLib <= 0 Or idxN <= 0 Or idxN1 <= 0 Then GoTo CleanExit

    Set wb = Workbooks.Open(fileName:=gBalancePath, UpdateLinks:=0, ReadOnly:=True, IgnoreReadOnlyRecommended:=True, AddToMru:=False)
    Set ws = GetFirstNonEmptySheetPublic(wb)
    If ws Is Nothing Then GoTo CleanExit

    lastRow = modUtils.GetLastUsedRowSafe(ws)
    lastCol = modUtils.GetLastUsedColSafe(ws)
    If lastRow <= 0 Or lastCol <= 0 Then GoTo CleanExit

    arrSrc = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Value2
    If skipRows > 0 Then arrSrc = SliceArrayRows(arrSrc, skipRows + 1)
    If IsEmpty(arrSrc) Or Not IsArray(arrSrc) Then GoTo CleanExit

    arrMapped = modLeadsWizard.Balance_MapTo4Cols(arrSrc, idxCompte, idxLib, idxN, idxN1)
    If IsEmpty(arrMapped) Or Not IsArray(arrMapped) Then GoTo CleanExit

    gFullData = modLeadsWizard.TransformerBG_Array(arrMapped)
    LoadFullDataCustomCols = IsArray(gFullData)

CleanExit:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Exit Function
EH:
    Debug.Print "LoadFullDataCustomCols error " & Err.Number & " : " & Err.Description
    LoadFullDataCustomCols = False
    Resume CleanExit
End Function

Private Function ColumnLetterToNumber(ByVal colRef As String) As Long
    Dim i As Long
    Dim ch As String
    Dim txt As String

    txt = UCase$(Trim$(colRef))
    If Len(txt) = 0 Then Exit Function

    For i = 1 To Len(txt)
        ch = Mid$(txt, i, 1)
        If ch < "A" Or ch > "Z" Then Exit Function
        ColumnLetterToNumber = ColumnLetterToNumber * 26 + (Asc(ch) - Asc("A") + 1)
    Next i
End Function

Private Function GetFirstNonEmptySheetPublic(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Function

    For Each ws In wb.Worksheets
        If modUtils.GetLastUsedRowSafe(ws) > 0 And modUtils.GetLastUsedColSafe(ws) > 0 Then
            Set GetFirstNonEmptySheetPublic = ws
            Exit Function
        End If
    Next ws
End Function

Private Function SliceArrayRows(ByVal arr As Variant, ByVal startRow As Long) As Variant
    Dim lastRow As Long
    Dim lastCol As Long
    Dim r As Long
    Dim c As Long
    Dim outArr() As Variant
    Dim outRow As Long

    On Error GoTo EH

    If IsEmpty(arr) Or Not IsArray(arr) Then Exit Function
    lastRow = UBound(arr, 1)
    lastCol = UBound(arr, 2)
    If startRow < 1 Then startRow = 1
    If startRow > lastRow Then Exit Function

    ReDim outArr(1 To lastRow - startRow + 1, 1 To lastCol)
    outRow = 0
    For r = startRow To lastRow
        outRow = outRow + 1
        For c = 1 To lastCol
            outArr(outRow, c) = arr(r, c)
        Next c
    Next r

    SliceArrayRows = outArr
    Exit Function
EH:
    SliceArrayRows = Empty
End Function

' ============================================================
' POINT D'ENTREE apres frmBGError (cmdGenerate)
' gFullData est deja rempli et BG importe
' ============================================================
Public Sub RunAfterBGError()
    ImportIntoBG_FromFullData
    Dim ufMeta As New frmLeadMeta
    ufMeta.Show vbModal
    Unload ufMeta
End Sub

' ============================================================
' GENERATION - appele par frmLeadMeta.cmdOK
' Sequence stricte : Mapping -> Param -> Recalc -> [KE -> Recalc] -> Export -> Reset
' ============================================================
Public Sub RunGenerateLeads_V4()
    Dim dExo As Date
    Dim oldScreen As Boolean, oldEvents As Boolean
    Dim oldAlerts As Boolean, oldCalc As XlCalculation
    Dim exportSucceeded As Boolean

    If Not IsValidExerciceDate_UI(gExercice, dExo) Then
        MsgBox "La date d'exercice est invalide.", vbExclamation
        Exit Sub
    End If

    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    oldCalc = Application.Calculation

    On Error GoTo EH
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    ' 1. Formules de mapping dans BG (colonnes E:S)
    InjectMappingFormulas_BG

    ' 2. Metadonnees dans Param du wbSource
    ApplyMetaToSourceParamSheet ThisWorkbook

    ' 3. Recalcul complet wbSource
    ForceFullRecalc

    ' 4. KE : diviser C:D du BG source par 1000, puis recalculer
    '    APRES le recalc initial pour travailler sur valeurs numeriques
    If gGenerateInKE Then
        ApplyKEOnSourceBG
        ForceFullRecalc
    End If

    ' 5. Creer wbOut en valeurs, details, formatage, enregistrement
    Application.Calculation = xlCalculationAutomatic
    exportSucceeded = False
    ExportValuesCopy_WithoutLeads_ToBalanceFolder_V4

    If Err.Number = 0 Then
        MsgBox "Fichier g"&ChrW$(233)&"n"&ChrW$(233)&"r"&ChrW$(233)&" et enregistr"&ChrW$(233)&" avec succ"&ChrW$(232)&"s.", vbInformation
    End If
    If Err.Number = 0 Then exportSucceeded = gLastExportSucceeded

CleanExit:
    ResetSourceAfterExport
    ThisWorkbook.Saved = True
    Application.Calculation = oldCalc
    Application.DisplayAlerts = oldAlerts
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    If exportSucceeded Then
        MsgBox "Fichier généré et enregistré avec succès.", vbInformation
        ActivateWorkbookOnBSView gLastExportedWorkbook
    End If
    Exit Sub
EH:
    MsgBox "Erreur génération : " & Err.Number & vbCrLf & Err.Description, vbCritical
    Resume CleanExit
End Sub

' Division C:D du BG source par 1000 (valeurs uniquement, pas formules)
Private Sub ApplyKEOnSourceBG()
    Dim wsBG As Worksheet
    Dim lastRow As Long
    Dim rng As Range
    On Error GoTo EH
    Set wsBG = ThisWorkbook.Worksheets(SH_BG)
    lastRow = wsBG.Cells(wsBG.rows.count, "A").End(xlUp).Row
    If lastRow < BG_FIRST_ROW Then Exit Sub
    Set rng = wsBG.Range("C" & BG_FIRST_ROW & ":D" & lastRow)
    modKEScaling.DivideRangeByThousand rng
    modKETrace.LogKE "SOURCE BG divise /1000 | range=" & rng.Address(False, False), "ApplyKEOnSourceBG"
    Exit Sub
EH:
    modKETrace.LogKE "ERROR " & Err.Number & " : " & Err.Description, "ApplyKEOnSourceBG"
End Sub

' Reset wbSource apres export
Private Sub ResetSourceAfterExport()
    On Error Resume Next
    'ClearSourceBG_AtoD_ExceptHeader ThisWorkbook  ' Desactive temporairement
    ResetSourceParamPlaceholders ThisWorkbook
    gMetaAppliedToParam = False
    On Error GoTo 0
End Sub

' ============================================================
' UTILITAIRES
' ============================================================
Public Sub ForceFullRecalc()
    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFullRebuild
    DoEvents
    Application.Calculation = xlCalculationManual
End Sub

Public Function ComputeMaxAccountLen(ByVal arr As Variant) As Long
    Dim i As Long, digits As String
    On Error Resume Next
    For i = 1 To UBound(arr, 1)
        digits = KeepDigits(CStr(arr(i, 1)))
        If Len(digits) > ComputeMaxAccountLen Then ComputeMaxAccountLen = Len(digits)
    Next i
End Function

Public Sub SafeLogNoUI(ByVal msg As String, ByVal proc As String)
    Debug.Print proc & " | " & msg
    On Error Resume Next
    modKETrace.LogKE msg, proc
    On Error GoTo 0
End Sub

' ============================================================
' ADDINS - Desactivation de tous les add-ins au lancement
' Objectif : aucune interference externe pendant l'execution
' de la macro (copier/coller, PDF, clipboard...).
'
' Couvre deux categories :
'   - COMAddIns : plugins COM (Grammarly, DataSniper, SAP...)
'   - AddIns    : add-ins XLA/XLAM (DataSniper, Analysis ToolPak...)
'
' La routine est entierement silencieuse :
'   - aucun plantage si un add-in est absent ou protege
'   - aucun message affich�
'   - l'execution continue normalement dans tous les cas
' ============================================================
Private Sub DisableAllAddins()
    ' Desactive uniquement les COM Add-ins (plugins tiers charges en memoire :
    ' Grammarly, DataSniper, SAP, outils Office...).
    ' Les AddIns XLA/XLAM classiques (Analysis ToolPak, etc.) ne sont PAS touches :
    ' leur desactivation est persistante (registre) et peut casser ExportAsFixedFormat
    ' sur certaines configurations Excel qui utilisent un add-in pour le driver PDF.
    Dim ca As COMAddIn
    On Error Resume Next
    For Each ca In Application.COMAddIns
        If ca.Connect Then
            ca.Connect = False
            modKETrace.LogKE "COMAddIn desactive : " & ca.Description, "DisableAllAddins"
        End If
    Next ca
    On Error GoTo 0
End Sub

Public Function PromptSaveAsPath_NoUI(ByVal initialPath As String) As Variant
    On Error GoTo EH
    PromptSaveAsPath_NoUI = Application.GetSaveAsFilename( _
        InitialFileName:=initialPath, _
        FileFilter:="Classeur Excel (*.xlsx), *.xlsx", _
        title:="Enregistrer le fichier g&"ChrW$(233)&"n"&ChrW$(233)&"r"&ChrW$(233)&")
    Exit Function
EH:
    PromptSaveAsPath_NoUI = False
End Function
