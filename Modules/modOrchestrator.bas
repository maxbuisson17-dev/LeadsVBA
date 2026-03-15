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
        EnsureWorkingSheetsHidden False
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
        .Title = "Selectionner la balance (Compte / Libelle / Solde N / Solde N-1)"
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

' ============================================================
' POINT D'ENTREE apres frmBGError (cmdGenerate)
' gFullData est deja rempli et BG importe
' ============================================================
Public Sub RunAfterBGError()
    EnsureWorkingSheetsHidden False
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

    If Not IsValidExerciceDate_UI(gExercice, dExo) Then
        MsgBox "La date d'exercice est invalide.", vbExclamation
        Exit Sub
    End If

    oldScreen  = Application.ScreenUpdating
    oldEvents  = Application.EnableEvents
    oldAlerts  = Application.DisplayAlerts
    oldCalc    = Application.Calculation

    On Error GoTo EH
    Application.ScreenUpdating = False
    Application.EnableEvents   = False
    Application.DisplayAlerts  = False
    Application.Calculation    = xlCalculationManual

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
    ExportValuesCopy_WithoutLeads_ToBalanceFolder_V4

    If Err.Number = 0 Then
        MsgBox "Fichier genere et enregistre avec succes.", vbInformation
    End If

CleanExit:
    ResetSourceAfterExport
    EnsureWorkingSheetsHidden True
    ThisWorkbook.Saved = True
    Application.Calculation    = oldCalc
    Application.DisplayAlerts  = oldAlerts
    Application.EnableEvents   = oldEvents
    Application.ScreenUpdating = oldScreen
    Exit Sub
EH:
    MsgBox "Erreur generation : " & Err.Number & vbCrLf & Err.Description, vbCritical
    Resume CleanExit
End Sub

' Division C:D du BG source par 1000 (valeurs uniquement, pas formules)
Private Sub ApplyKEOnSourceBG()
    Dim wsBG As Worksheet
    Dim lastRow As Long
    Dim rng As Range
    On Error GoTo EH
    Set wsBG = ThisWorkbook.Worksheets(SH_BG)
    lastRow = wsBG.Cells(wsBG.Rows.Count, "A").End(xlUp).Row
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
    ClearSourceBG_AtoD_ExceptHeader ThisWorkbook
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

Public Function PromptSaveAsPath_NoUI(ByVal initialPath As String) As Variant
    On Error GoTo EH
    PromptSaveAsPath_NoUI = Application.GetSaveAsFilename( _
        InitialFileName:=initialPath, _
        FileFilter:="Classeur Excel (*.xlsx), *.xlsx", _
        Title:="Enregistrer le fichier genere")
    Exit Function
EH:
    PromptSaveAsPath_NoUI = False
End Function
