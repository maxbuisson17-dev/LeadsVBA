Attribute VB_Name = "modDA_BG"
' ============================================================
' modDA_BG.bas
' Acces donnees - feuille BG (lecture/ecriture)
' ============================================================
Option Explicit

Public Sub ImportIntoBG_FromFullData()
    Dim wsBG As Worksheet: Set wsBG = ThisWorkbook.Worksheets(SH_BG)
    If IsEmpty(gFullData) Then Err.Raise vbObjectError + 100, , "Aucune donnee a importer : gFullData est vide."
    If Not IsArray(gFullData) Then Err.Raise vbObjectError + 101, , "Aucune donnee a importer : gFullData n'est pas un tableau."
    Dim n As Long, m As Long
    n = UBound(gFullData, 1)
    m = UBound(gFullData, 2)
    If n <= 0 Then Err.Raise vbObjectError + 102, , "Aucune donnee a importer (0 ligne)."
    If m < 4 Then Err.Raise vbObjectError + 104, , "4 colonnes attendues (A:D), trouve " & m & "."
    Dim arr As Variant
    arr = gFullData
    ' Normaliser comptes => digits only puis nombre (attention: si comptes tr?s longs, Double peut arrondir)
    Dim i As Long, raw As String, digits As String
    For i = 1 To n
        raw = CStr(arr(i, 1))
        digits = KeepDigits(raw)
        If digits = "" Then
            Err.Raise vbObjectError + 120, , "Compte invalide a la ligne " & i & " : '" & raw & "' (aucun chiffre detecte)."
        End If
        ' On garde en texte si trop long, mais ton template BG utilise GAUCHE() sur A => OK m?me en texte
        arr(i, 1) = digits
    Next i
    
    
    Dim lastRow As Long
    lastRow = BG_FIRST_ROW + n - 1
    Application.ScreenUpdating = False
    wsBG.Range(wsBG.Cells(BG_FIRST_ROW, COL_A), wsBG.Cells(wsBG.rows.Count, COL_D)).ClearContents
    wsBG.Range(wsBG.Cells(BG_FIRST_ROW, COL_A), wsBG.Cells(lastRow, COL_D)).value = arr
    ' A en texte num?rique (?vite arrondis)
    With wsBG.Range(wsBG.Cells(BG_FIRST_ROW, COL_A), wsBG.Cells(lastRow, COL_A))
        .NumberFormat = "@"
    End With
    
    wsBG.Range("E" & BG_FIRST_ROW).FormulaLocal = "=GAUCHE($A" & BG_FIRST_ROW & ";1)"
    wsBG.Range("E" & BG_FIRST_ROW).Copy
    wsBG.Range("E" & BG_FIRST_ROW & ":E" & lastRow).PasteSpecial xlPasteFormulas
    Application.CutCopyMode = False
    wsBG.Range("E" & BG_FIRST_ROW & ":E" & lastRow).value = wsBG.Range("E" & BG_FIRST_ROW & ":E" & lastRow).value
    wsBG.Range("E" & BG_FIRST_ROW & ":E" & lastRow).NumberFormat = "0"
    
    ' F:I = 2..5 premiers chiffres
    wsBG.Range("F" & BG_FIRST_ROW).FormulaLocal = "=GAUCHE($A" & BG_FIRST_ROW & ";2)"
    wsBG.Range("G" & BG_FIRST_ROW).FormulaLocal = "=GAUCHE($A" & BG_FIRST_ROW & ";3)"
    wsBG.Range("H" & BG_FIRST_ROW).FormulaLocal = "=GAUCHE($A" & BG_FIRST_ROW & ";4)"
    wsBG.Range("I" & BG_FIRST_ROW).FormulaLocal = "=GAUCHE($A" & BG_FIRST_ROW & ";5)"
    wsBG.Range("F" & BG_FIRST_ROW & ":I" & BG_FIRST_ROW).Copy
    wsBG.Range("F" & BG_FIRST_ROW & ":I" & lastRow).PasteSpecial xlPasteFormulas
    Application.CutCopyMode = False
    Dim rngFI As Range
    Set rngFI = wsBG.Range(wsBG.Cells(BG_FIRST_ROW, COL_F), wsBG.Cells(lastRow, COL_I))
    rngFI.value = rngFI.value
    rngFI.NumberFormat = "0"
    rngFI.value = wsBG.Evaluate("IF(" & rngFI.Address & "<>""""," & rngFI.Address & "*1,"""")")
    Application.ScreenUpdating = True
End Sub

Public Sub InjectMappingFormulas_BG()
    On Error GoTo EH
    Dim wsBG As Worksheet
    Set wsBG = ThisWorkbook.Worksheets(SH_BG)
    Dim lastRow As Long
    lastRow = wsBG.Cells(wsBG.rows.Count, "A").End(xlUp).Row
    If lastRow < BG_FIRST_ROW Then Exit Sub
    ' IMPORTANT : la colonne E doit exister et ?tre remplie (GAUCHE #1)
    ' sinon les formules ci-dessous (qui utilisent $E...) peuvent provoquer des soucis.
    If Len(Trim$(CStr(wsBG.Cells(BG_FIRST_ROW, "E").value))) = 0 Then
        wsBG.Range("E" & BG_FIRST_ROW).FormulaLocal = "=GAUCHE($A" & BG_FIRST_ROW & ";1)"
        wsBG.Range("E" & BG_FIRST_ROW).Copy
        wsBG.Range("E" & BG_FIRST_ROW & ":E" & lastRow).PasteSpecial xlPasteFormulas
        Application.CutCopyMode = False
        wsBG.Range("E" & BG_FIRST_ROW & ":E" & lastRow).value = wsBG.Range("E" & BG_FIRST_ROW & ":E" & lastRow).value
        wsBG.Range("E" & BG_FIRST_ROW & ":E" & lastRow).NumberFormat = "0"
    End If
    ' Plage Mapping (fixe chez toi, OK)
    Const MAP_RANGE As String = "Mapping!$A$1:$M$9000"
    Dim cas4 As String, cas5 As String, cas6 As String, cas7 As String
    Dim cas8 As String, cas9 As String, cas10 As String
    cas4 = "SIERREUR(RECHERCHEV($H" & BG_FIRST_ROW & ";" & MAP_RANGE & ";4;FAUX);" & _
           "SIERREUR(RECHERCHEV($G" & BG_FIRST_ROW & ";" & MAP_RANGE & ";4;FAUX);" & _
           "SIERREUR(RECHERCHEV($F" & BG_FIRST_ROW & ";" & MAP_RANGE & ";4;FAUX);" & _
           "RECHERCHEV($E" & BG_FIRST_ROW & ";" & MAP_RANGE & ";4;FAUX))))"
    cas5 = "SIERREUR(RECHERCHEV($H" & BG_FIRST_ROW & ";" & MAP_RANGE & ";5;FAUX);" & _
           "SIERREUR(RECHERCHEV($G" & BG_FIRST_ROW & ";" & MAP_RANGE & ";5;FAUX);" & _
           "SIERREUR(RECHERCHEV($F" & BG_FIRST_ROW & ";" & MAP_RANGE & ";5;FAUX);" & _
           "RECHERCHEV($E" & BG_FIRST_ROW & ";" & MAP_RANGE & ";5;FAUX))))"
    cas6 = "SIERREUR(RECHERCHEV($H" & BG_FIRST_ROW & ";" & MAP_RANGE & ";6;FAUX);" & _
           "SIERREUR(RECHERCHEV($G" & BG_FIRST_ROW & ";" & MAP_RANGE & ";6;FAUX);" & _
           "SIERREUR(RECHERCHEV($F" & BG_FIRST_ROW & ";" & MAP_RANGE & ";6;FAUX);" & _
           "RECHERCHEV($E" & BG_FIRST_ROW & ";" & MAP_RANGE & ";6;FAUX))))"
    cas7 = "SIERREUR(RECHERCHEV($H" & BG_FIRST_ROW & ";" & MAP_RANGE & ";7;FAUX);" & _
           "SIERREUR(RECHERCHEV($G" & BG_FIRST_ROW & ";" & MAP_RANGE & ";7;FAUX);" & _
           "SIERREUR(RECHERCHEV($F" & BG_FIRST_ROW & ";" & MAP_RANGE & ";7;FAUX);" & _
           "RECHERCHEV($E" & BG_FIRST_ROW & ";" & MAP_RANGE & ";7;FAUX))))"
    cas8 = "SIERREUR(RECHERCHEV($H" & BG_FIRST_ROW & ";" & MAP_RANGE & ";8;FAUX);" & _
           "SIERREUR(RECHERCHEV($G" & BG_FIRST_ROW & ";" & MAP_RANGE & ";8;FAUX);" & _
           "SIERREUR(RECHERCHEV($F" & BG_FIRST_ROW & ";" & MAP_RANGE & ";8;FAUX);" & _
           "RECHERCHEV($E" & BG_FIRST_ROW & ";" & MAP_RANGE & ";8;FAUX))))"
    cas9 = "SIERREUR(RECHERCHEV($H" & BG_FIRST_ROW & ";" & MAP_RANGE & ";9;FAUX);" & _
           "SIERREUR(RECHERCHEV($G" & BG_FIRST_ROW & ";" & MAP_RANGE & ";9;FAUX);" & _
           "SIERREUR(RECHERCHEV($F" & BG_FIRST_ROW & ";" & MAP_RANGE & ";9;FAUX);" & _
           "RECHERCHEV($E" & BG_FIRST_ROW & ";" & MAP_RANGE & ";9;FAUX))))"
    cas10 = "SIERREUR(RECHERCHEV($H" & BG_FIRST_ROW & ";" & MAP_RANGE & ";10;FAUX);" & _
            "SIERREUR(RECHERCHEV($G" & BG_FIRST_ROW & ";" & MAP_RANGE & ";10;FAUX);" & _
            "SIERREUR(RECHERCHEV($F" & BG_FIRST_ROW & ";" & MAP_RANGE & ";10;FAUX);" & _
            "RECHERCHEV($E" & BG_FIRST_ROW & ";" & MAP_RANGE & ";10;FAUX))))"
    Dim fI As String, fJ As String, fK As String, fL As String
    Dim fM As String, fN As String, fO As String, fP As String
    fI = "=SI($C" & BG_FIRST_ROW & ">0;" & cas4 & ";" & cas5 & ")"
    fJ = "=SI($C" & BG_FIRST_ROW & ">0;" & cas6 & ";" & cas7 & ")"
    fK = "=SI($C" & BG_FIRST_ROW & ">0;" & cas8 & ";" & cas9 & ")"
    fL = "=SI($C" & BG_FIRST_ROW & ">0;" & cas10 & ";" & cas10 & ")"
    fM = "=SI($D" & BG_FIRST_ROW & ">0;" & cas4 & ";" & cas5 & ")"
    fN = "=SI($D" & BG_FIRST_ROW & ">0;" & cas6 & ";" & cas7 & ")"
    fO = "=SI($D" & BG_FIRST_ROW & ">0;" & cas8 & ";" & cas9 & ")"
    fP = "=SI($D" & BG_FIRST_ROW & ">0;" & cas10 & ";" & cas10 & ")"
    ' Injection robuste : Formula2Local si dispo, sinon fallback FormulaLocal
    With wsBG
        .Range(.Cells(BG_FIRST_ROW, 9), .Cells(lastRow, 16)).NumberFormat = "General"
        On Error Resume Next
        .Range(.Cells(BG_FIRST_ROW, 9), .Cells(lastRow, 9)).Formula2Local = fI
        If Err.Number <> 0 Then Err.Clear: .Range(.Cells(BG_FIRST_ROW, 9), .Cells(lastRow, 9)).FormulaLocal = fI
        .Range(.Cells(BG_FIRST_ROW, 10), .Cells(lastRow, 10)).Formula2Local = fJ
        If Err.Number <> 0 Then Err.Clear: .Range(.Cells(BG_FIRST_ROW, 10), .Cells(lastRow, 10)).FormulaLocal = fJ
        .Range(.Cells(BG_FIRST_ROW, 11), .Cells(lastRow, 11)).Formula2Local = fK
        If Err.Number <> 0 Then Err.Clear: .Range(.Cells(BG_FIRST_ROW, 11), .Cells(lastRow, 11)).FormulaLocal = fK
        .Range(.Cells(BG_FIRST_ROW, 12), .Cells(lastRow, 12)).Formula2Local = fL
        If Err.Number <> 0 Then Err.Clear: .Range(.Cells(BG_FIRST_ROW, 12), .Cells(lastRow, 12)).FormulaLocal = fL
        .Range(.Cells(BG_FIRST_ROW, 13), .Cells(lastRow, 13)).Formula2Local = fM
        If Err.Number <> 0 Then Err.Clear: .Range(.Cells(BG_FIRST_ROW, 13), .Cells(lastRow, 13)).FormulaLocal = fM
        .Range(.Cells(BG_FIRST_ROW, 14), .Cells(lastRow, 14)).Formula2Local = fN
        If Err.Number <> 0 Then Err.Clear: .Range(.Cells(BG_FIRST_ROW, 14), .Cells(lastRow, 14)).FormulaLocal = fN
        .Range(.Cells(BG_FIRST_ROW, 15), .Cells(lastRow, 15)).Formula2Local = fO
        If Err.Number <> 0 Then Err.Clear: .Range(.Cells(BG_FIRST_ROW, 15), .Cells(lastRow, 15)).FormulaLocal = fO
        .Range(.Cells(BG_FIRST_ROW, 16), .Cells(lastRow, 16)).Formula2Local = fP
        If Err.Number <> 0 Then Err.Clear: .Range(.Cells(BG_FIRST_ROW, 16), .Cells(lastRow, 16)).FormulaLocal = fP
        On Error GoTo EH
    End With
    Exit Sub
EH:
    MsgBox "Erreur 1004 dans InjectMappingFormulas_BG" & vbCrLf & _
           "Detail : " & Err.Number & " - " & Err.Description, vbCritical
End Sub

Public Sub ClearSourceBG_AtoD_ExceptHeader(ByVal wbSrc As Workbook)
    Dim wsBG As Worksheet
    Dim lastRow As Long

    On Error GoTo EH
    Set wsBG = wbSrc.Worksheets(SH_BG)
    lastRow = wsBG.Cells(wsBG.rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then GoTo CleanExit

    wsBG.Range("A2:D" & lastRow).ClearContents

CleanExit:
    Exit Sub
EH:
    Debug.Print "ClearSourceBG_AtoD_ExceptHeader error " & Err.Number & " : " & Err.Description
    Resume CleanExit
End Sub

Public Sub ConvertUsedRangeToValues(ByVal ws As Worksheet)
    Dim ur As Range
    Set ur = ws.UsedRange
    If ur Is Nothing Then Exit Sub
    ur.value = ur.value
End Sub

Public Sub CopySheetIntoExistingAsValuesAndFormats(ByVal wsSrc As Worksheet, ByVal wsDest As Worksheet)
    wsDest.Cells.Clear
    Dim ur As Range
    Set ur = wsSrc.UsedRange
    If ur Is Nothing Then Exit Sub
    wsDest.Range("A1").Resize(ur.rows.Count, ur.Columns.Count).value = ur.value
    ur.Copy
    With wsDest.Range("A1")
        .PasteSpecial xlPasteFormats
        .PasteSpecial xlPasteColumnWidths
    End With
    Application.CutCopyMode = False
    ThisWorkbook.Saved = True
End Sub

Private Sub RebuildCopySheet(ByVal wb As Workbook, ByVal srcName As String, ByVal dstName As String)
    Dim wsSrc As Worksheet, wsDst As Worksheet
    On Error Resume Next
    Set wsSrc = wb.Worksheets(srcName)
    On Error GoTo 0
    If wsSrc Is Nothing Then Exit Sub
    
    DeleteSheetIfExists wb, dstName
    
    wsSrc.Copy After:=wsSrc
    Set wsDst = wsSrc.Next
    wsDst.name = dstName
    
    ConvertUsedRangeToValues wsDst
End Sub

Public Sub DeleteSheetIfExists(ByVal wb As Workbook, ByVal sName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sName)
    On Error GoTo 0
    If Not ws Is Nothing Then
        ws.Delete
    End If
End Sub

Public Sub RenameSheetIfExists(ByVal wb As Workbook, ByVal oldName As String, ByVal newName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(oldName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ws.name = newName
    On Error GoTo 0
End Sub
