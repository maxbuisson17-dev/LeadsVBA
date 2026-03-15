Attribute VB_Name = "modSession"
' ============================================================
' modSession.bas
' Gestion du contexte session - reset entre deux lancements
' ============================================================
Option Explicit

Public Sub ResetImportContext()
    gPathN            = vbNullString
    gPathN1           = vbNullString
    gArrN             = Empty
    gArrN1            = Empty
    gArrCompiled      = Empty
    gFullData         = Empty
    gImportedN        = False
    gImportedN1       = False
    gImportParamsReady = False
    gExportMode       = emFS
    gMetaAppliedToParam = False
    gOkToGenerate     = False
    gHasNonBlockingIssue = False
    Set gControlRows  = New Collection
End Sub

Public Function HasBalanceN() As Boolean
    HasBalanceN = IsArray(gArrN) And gImportedN
End Function

Public Function HasBalanceN1() As Boolean
    HasBalanceN1 = IsArray(gArrN1) And gImportedN1
End Function

Public Function BalanceRowCount(ByVal arr As Variant) As Long
    On Error Resume Next
    If IsArray(arr) Then BalanceRowCount = UBound(arr, 1)
End Function
