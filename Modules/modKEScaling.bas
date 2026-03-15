Attribute VB_Name = "modKEScaling"
' ============================================================
' modKEScaling.bas
' Passage K euros : division par 1000 des soldes N et N-1
' LOGIQUE : si chkKE = True => on divise C:D du BG SOURCE par 1000
'           APRES le recalc initial, AVANT la creation du wbOut.
'           Le wbOut recoit donc des valeurs deja divisees.
' ============================================================
Option Explicit

' Divise une plage de cellules par 1000 (valeurs numeriques non vides)
Public Sub DivideRangeByThousand(ByVal rng As Range)
    If rng Is Nothing Then Exit Sub
    Dim data As Variant
    data = rng.Value
    Dim r As Long, c As Long, v As Variant
    If IsArray(data) Then
        For r = 1 To UBound(data, 1)
            For c = 1 To UBound(data, 2)
                v = data(r, c)
                If Not IsEmpty(v) And IsNumeric(v) Then
                    data(r, c) = CDbl(v) / 1000#
                End If
            Next c
        Next r
        rng.Value = data
    Else
        If Not IsEmpty(data) And IsNumeric(data) Then
            rng.Value = CDbl(data) / 1000#
        End If
    End If
End Sub
