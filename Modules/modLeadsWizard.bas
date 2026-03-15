Attribute VB_Name = "modLeadsWizard"
' ============================================================
' modLeadsWizard.bas - MODULE STUBS UNIQUEMENT
' Les globals/constantes sont dans modGlobals.bas
' Ce module expose les alias pour compatibilite UserForms
' ============================================================
Option Explicit

' ALIAS STUBS - Redirection vers modules specialises
' Ces wrappers maintiennent la compatibilite avec les UserForms
' qui appellent modLeadsWizard.XXX
' ============================================================

Public Function ImporterFECtoBG(Optional ByRef outSelectedPath As String, Optional ByVal pathTxt As String = vbNullString) As Variant
    ' REFACTOR: import FEC en array 2D (Compte/Libelle/Solde), sans creation de classeur.
    ImporterFECtoBG = modBalanceCreator.ImporterFECtoBG(outSelectedPath, pathTxt)
End Function

Public Function FEC_ToBalanceArray_FromPath(ByVal pathTxt As String) As Variant
    ' REFACTOR: fa?ade centralis?e
    FEC_ToBalanceArray_FromPath = modBalanceCreator.FEC_ToBalanceArray_FromPath(pathTxt)
End Function

Public Function BalanceExcel_ToBalanceArray_FromPath(ByVal filePath As String, Optional ByRef outInfo As String) As Variant
    ' REFACTOR: fa?ade centralis?e
    BalanceExcel_ToBalanceArray_FromPath = modBalanceCreator.BalanceExcel_ToBalanceArray_FromPath(filePath, outInfo)
End Function

Public Function BGCompil_ArrayHasRows(ByVal arr As Variant) As Boolean
    BGCompil_ArrayHasRows = modBalanceCreator.BGCompil_ArrayHasRows(arr)
End Function

Public Function BGCompil_FromBalanceArrays( _
    ByVal balN As Variant, _
    ByVal balN1 As Variant, _
    Optional ByRef outMaxLen As Long, _
    Optional ByRef outInfo As String) As Variant

    Dim arrOut As Variant
    arrOut = modBalanceCreator.BGCompil_FromBalanceArrays(balN, balN1, outMaxLen, outInfo)
    BGCompil_FromBalanceArrays = NormalizeAccentsOnLabelColumn(arrOut)
End Function

Public Function Balance_DetectHeaderRow(ByVal arrSrc As Variant) As Boolean
    Balance_DetectHeaderRow = modBalanceCreator.Balance_DetectHeaderRow(arrSrc)
End Function

Public Function Balance_GuessColumns(ByVal arrSrc As Variant, ByRef idxCompte As Long, ByRef idxLib As Long, ByRef idxSoldeN As Long, ByRef idxSoldeN1 As Long) As Boolean
    Balance_GuessColumns = modBalanceCreator.Balance_GuessColumns(arrSrc, idxCompte, idxLib, idxSoldeN, idxSoldeN1)
End Function

Public Function Balance_MapTo4Cols(ByVal arrSrc As Variant, ByVal idxCompte As Long, ByVal idxLib As Long, ByVal idxSoldeN As Long, ByVal idxSoldeN1 As Long) As Variant
    Balance_MapTo4Cols = modBalanceCreator.Balance_MapTo4Cols(arrSrc, idxCompte, idxLib, idxSoldeN, idxSoldeN1)
End Function

Public Function TransformerBG_Array(ByVal arrIn As Variant) As Variant
    TransformerBG_Array = modBalanceCreator.TransformerBG_Array(arrIn)
End Function

Public Function NormalizeTo4Cols(ByVal arrBalance As Variant, ByVal mode As String, Optional ByVal fillN1WithZero As Boolean = True) As Variant
    NormalizeTo4Cols = modBalanceCreator.NormalizeTo4Cols(arrBalance, mode, fillN1WithZero)
End Function

Public Function DetectBalanceFormat_FromPath(ByVal filePath As String) As eBalanceFormat
    DetectBalanceFormat_FromPath = modBalanceCreator.DetectBalanceFormat_FromPath(filePath)
End Function

Public Function LoadBalanceArray_FromExcelPath(ByVal path As String) As Variant
    LoadBalanceArray_FromExcelPath = modBalanceCreator.LoadBalanceArray_FromExcelPath(path)
End Function

Public Function Ensure4Cols(ByVal arr As Variant) As Variant
    ' PATCH: garantit un Variant 2D (1..n,1..4) propre et normalise.
    Dim n As Long, m As Long, i As Long
    Dim tmp() As Variant

    On Error GoTo EH
    If IsEmpty(arr) Or Not IsArray(arr) Then Exit Function

    n = UBound(arr, 1)
    m = UBound(arr, 2)
    If n <= 0 Or m <= 0 Then Exit Function

    ReDim tmp(1 To n, 1 To 4)
    For i = 1 To n
        tmp(i, 1) = arr(i, 1)
        If m >= 2 Then tmp(i, 2) = arr(i, 2) Else tmp(i, 2) = vbNullString
        If m >= 3 Then tmp(i, 3) = arr(i, 3) Else tmp(i, 3) = 0#
        If m >= 4 Then tmp(i, 4) = arr(i, 4) Else tmp(i, 4) = 0#
    Next i

    Ensure4Cols = TransformerBG_Array(tmp)
    Exit Function
EH:
    Ensure4Cols = Empty
End Function

' ============================================================
' Workbook events
' ============================================================

Public Sub Workbook_BeforeClose(ByVal wbOut As Workbook)
    Dim oldAlerts As Boolean, oldEvents As Boolean, oldScreen As Boolean, oldCalc As XlCalculation
    oldAlerts = Application.DisplayAlerts
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    oldCalc = Application.Calculation
    On Error GoTo EH
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    ' KE : suppression d'?ventuels restes _KE (aucune cr?ation)
    DeleteSheetIfExists wbOut, "BS_KE"
    DeleteSheetIfExists wbOut, "SIG_KE"
    DeleteSheetIfExists wbOut, "BG_KE"
CLEAN_EXIT:
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    Exit Sub
EH:
    
    Resume CLEAN_EXIT
End Sub
