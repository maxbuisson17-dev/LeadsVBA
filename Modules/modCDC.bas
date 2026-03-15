Attribute VB_Name = "modCDC"
' ============================================================
' modCDC.bas
' Controles de coherence sur la balance importee
' ============================================================
Option Explicit

Public Sub BuildControlReportFromFullData()
    Const EPS As Double = 1#   ' tolerance ? 1? demand?e
    Set gControlRows = New Collection
    gOkToGenerate = False
    gHasNonBlockingIssue = False
    If IsEmpty(gFullData) Or Not IsArray(gFullData) Then
        AddControlRow "Aucune donnee chargee", False, True
        Exit Sub
    End If
    Dim sumBalN As Double, sumBalN1 As Double
    Dim hasNonNumeric As Boolean
    Dim sum67N1 As Double
    Dim n As Long, i As Long, acc As String, digits As String
    n = UBound(gFullData, 1)
    ' ---- Test: balance contient > 4 comptes (non bloquant)
    AddControlRow "Balance contient plus de 4 comptes", (n > 4), False
    ' ---- Sommes balance N / N-1 + d?tection montants non num?riques
    For i = 1 To n
        acc = CStr(gFullData(i, 1))
        digits = KeepDigits(acc)
        ' montants non num?riques (vide accept?) => bloquant
        If Not IsEmpty(gFullData(i, 3)) And Len(Trim$(CStr(gFullData(i, 3)))) > 0 Then
            If Not IsNumeric(gFullData(i, 3)) Then hasNonNumeric = True
        End If
        If Not IsEmpty(gFullData(i, 4)) And Len(Trim$(CStr(gFullData(i, 4)))) > 0 Then
            If Not IsNumeric(gFullData(i, 4)) Then hasNonNumeric = True
        End If
        ' Sommes
        If IsNumeric(gFullData(i, 3)) Then sumBalN = sumBalN + CDbl(gFullData(i, 3))
        If IsNumeric(gFullData(i, 4)) Then sumBalN1 = sumBalN1 + CDbl(gFullData(i, 4))
        ' (6+7) en N-1 (non bloquant si =0)
        If Len(digits) > 0 Then
            If Left$(digits, 1) = "6" Or Left$(digits, 1) = "7" Then
                If IsNumeric(gFullData(i, 4)) Then sum67N1 = sum67N1 + CDbl(gFullData(i, 4))
            End If
        End If
    Next i
    ' ---- Regles bloquantes
    
    Dim isBalanced As Boolean
    isBalanced = (Abs(sumBalN) <= EPS And Abs(sumBalN1) <= EPS)
    AddControlRow "Balance equilibree : Somme N<>0 et Somme N-1<>0 (tolerance 1?)", isBalanced, True
    
    AddControlRow "Montants numeriques (vide accepte)", (Not hasNonNumeric), True
    ' ---- Non bloquant: si somme comptes 6 et 7 en N-1 = 0 => remonter erreur mais permettre generation
    Dim ok67 As Boolean
    ok67 = (Abs(sum67N1) > EPS)
    If Not ok67 Then gHasNonBlockingIssue = True
    AddControlRow "Somme comptes 6 et 7 en N-1 <> 0 (non bloquant)", ok67, False
    ' ---- Decision generation: uniquement erreurs bloquantes
    gOkToGenerate = Not HasBlockingFailure()
End Sub

Public Function HasBlockingFailure() As Boolean
    Dim i As Long
    If gControlRows Is Nothing Then Exit Function
    For i = 1 To gControlRows.Count
        ' gControlRows(i) = Array(test, okBool, criticalBool)
        If CBool(gControlRows(i)(2)) = True Then
            If CBool(gControlRows(i)(1)) = False Then
                HasBlockingFailure = True
                Exit Function
            End If
        End If
    Next i
End Function

Public Sub AddControlRow(ByVal testName As String, ByVal isOk As Boolean, ByVal isCritical As Boolean)
    If gControlRows Is Nothing Then Set gControlRows = New Collection
    gControlRows.Add Array(testName, isOk, isCritical)
End Sub
