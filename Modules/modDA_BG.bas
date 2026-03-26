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

    ' Garantir que E est remplie (fallback si appel standalone)
    If Len(Trim$(CStr(wsBG.Cells(BG_FIRST_ROW, "E").value))) = 0 Then
        wsBG.Range("E" & BG_FIRST_ROW).FormulaLocal = "=GAUCHE($A" & BG_FIRST_ROW & ";1)"
        wsBG.Range("E" & BG_FIRST_ROW).Copy
        wsBG.Range("E" & BG_FIRST_ROW & ":E" & lastRow).PasteSpecial xlPasteFormulas
        Application.CutCopyMode = False
        With wsBG.Range("E" & BG_FIRST_ROW & ":E" & lastRow)
            .value = .value
            .NumberFormat = "0"
        End With
    End If

    ' --- 1. Charger Mapping en memoire ---
    Dim wsMap As Worksheet
    Set wsMap = ThisWorkbook.Worksheets(SH_MAP)

    Dim mapLastRow As Long
    mapLastRow = wsMap.Cells(wsMap.rows.Count, "A").End(xlUp).Row

    ' Besoin jusqu'a la col O (index 15) pour la colonne S
    Dim mapLastCol As Long
    mapLastCol = wsMap.Cells(1, wsMap.Columns.Count).End(xlToLeft).Column
    If mapLastCol < 15 Then mapLastCol = 15

    If mapLastRow < 2 Then
        ' Mapping vide : effacer I:S et sortir
        wsBG.Range(wsBG.Cells(BG_FIRST_ROW, 9), wsBG.Cells(lastRow, 19)).ClearContents
        Exit Sub
    End If

    Dim mapArr As Variant
    mapArr = wsMap.Range(wsMap.Cells(1, 1), wsMap.Cells(mapLastRow, mapLastCol)).value

    ' --- 2. Dictionnaire : cle Mapping col A -> index de ligne dans mapArr ---
    ' (premiere occurrence uniquement, comme VLOOKUP)
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1  ' vbTextCompare

    Dim r As Long
    For r = 2 To mapLastRow
        Dim mapKey As String
        mapKey = Trim$(CStr(mapArr(r, 1)))
        If mapKey <> "" And Not dict.Exists(mapKey) Then
            dict(mapKey) = r
        End If
    Next r

    ' --- 3. Lire les colonnes C, D, E, F, G, H de BG en memoire ---
    Dim nRows As Long
    nRows = lastRow - BG_FIRST_ROW + 1

    Dim bgC As Variant: bgC = wsBG.Range(wsBG.Cells(BG_FIRST_ROW, 3), wsBG.Cells(lastRow, 3)).value
    Dim bgD As Variant: bgD = wsBG.Range(wsBG.Cells(BG_FIRST_ROW, 4), wsBG.Cells(lastRow, 4)).value
    Dim bgE As Variant: bgE = wsBG.Range(wsBG.Cells(BG_FIRST_ROW, 5), wsBG.Cells(lastRow, 5)).value
    Dim bgF As Variant: bgF = wsBG.Range(wsBG.Cells(BG_FIRST_ROW, 6), wsBG.Cells(lastRow, 6)).value
    Dim bgG As Variant: bgG = wsBG.Range(wsBG.Cells(BG_FIRST_ROW, 7), wsBG.Cells(lastRow, 7)).value
    Dim bgH As Variant: bgH = wsBG.Range(wsBG.Cells(BG_FIRST_ROW, 8), wsBG.Cells(lastRow, 8)).value

    ' --- 4. Calculer I:P et Q:S dans deux tableaux separes ---
    ' Deux tableaux distincts pour ecriture independante (evite les conflits
    ' avec les formules de type Formula2/tableau dynamique presentes dans le template)
    Dim outIP As Variant
    ReDim outIP(1 To nRows, 1 To 8)   ' I(1) J(2) K(3) L(4) M(5) N(6) O(7) P(8)
    Dim outQRS As Variant
    ReDim outQRS(1 To nRows, 1 To 3)  ' Q(1) R(2) S(3)

    Dim mx As Long: mx = UBound(mapArr, 2)
    Dim i As Long
    For i = 1 To nRows
        Dim kE As String, kF As String, kG As String, kH As String
        kE = Trim$(CStr(bgE(i, 1)))
        kF = Trim$(CStr(bgF(i, 1)))
        kG = Trim$(CStr(bgG(i, 1)))
        kH = Trim$(CStr(bgH(i, 1)))

        Dim valC As Double, valD As Double
        valC = 0: valD = 0
        If IsNumeric(bgC(i, 1)) Then valC = CDbl(bgC(i, 1))
        If IsNumeric(bgD(i, 1)) Then valD = CDbl(bgD(i, 1))

        ' Trouver la meilleure ligne Mapping avec fallback H -> G -> F -> E
        ' (meme logique que SIERREUR(RECHERCHEV($H...);SIERREUR(RECHERCHEV($G...);...)))
        Dim bestRow As Long
        bestRow = 0
        If kH <> "" And dict.Exists(kH) Then
            bestRow = dict(kH)
        ElseIf kG <> "" And dict.Exists(kG) Then
            bestRow = dict(kG)
        ElseIf kF <> "" And dict.Exists(kF) Then
            bestRow = dict(kF)
        ElseIf kE <> "" And dict.Exists(kE) Then
            bestRow = dict(kE)
        End If

        If bestRow = 0 Then
            ' Aucune cle trouvee -> tout vide
            Dim j As Integer
            For j = 1 To 8: outIP(i, j) = "": Next j
            outQRS(i, 1) = "": outQRS(i, 2) = "": outQRS(i, 3) = ""
        Else
            ' I:P
            outIP(i, 1) = GetMapColValue(mapArr, bestRow, IIf(valC > 0, 4, 5), mx)
            outIP(i, 2) = GetMapColValue(mapArr, bestRow, IIf(valC > 0, 6, 7), mx)
            outIP(i, 3) = GetMapColValue(mapArr, bestRow, IIf(valC > 0, 8, 9), mx)
            outIP(i, 4) = GetMapColValue(mapArr, bestRow, 10, mx)
            outIP(i, 5) = GetMapColValue(mapArr, bestRow, IIf(valD > 0, 4, 5), mx)
            outIP(i, 6) = GetMapColValue(mapArr, bestRow, IIf(valD > 0, 6, 7), mx)
            outIP(i, 7) = GetMapColValue(mapArr, bestRow, IIf(valD > 0, 8, 9), mx)
            outIP(i, 8) = GetMapColValue(mapArr, bestRow, 10, mx)
            ' Q:S
            outQRS(i, 1) = GetMapColValue(mapArr, bestRow, 12, mx)  ' Q = Mapping col L
            outQRS(i, 2) = GetMapColValue(mapArr, bestRow, 13, mx)  ' R = Mapping col M
            ' S : si I = "Actif" -> Mapping col N (14), sinon col O (15)
            If CStr(outIP(i, 1)) = "Actif" Then
                outQRS(i, 3) = GetMapColValue(mapArr, bestRow, 14, mx)
            Else
                outQRS(i, 3) = GetMapColValue(mapArr, bestRow, 15, mx)
            End If
        End If
    Next i

    ' --- 5. Ecriture en bloc ---
    ' I:P (cols 9:16) en une passe
    wsBG.Range(wsBG.Cells(BG_FIRST_ROW, 9), wsBG.Cells(lastRow, 16)).value = outIP
    ' Q:S (cols 17:19) : ClearContents d'abord pour lever toute formule/tableau
    ' dynamique existant dans le template, puis ecriture valeurs
    On Error Resume Next
    wsBG.Range(wsBG.Cells(BG_FIRST_ROW, 17), wsBG.Cells(lastRow, 19)).ClearContents
    On Error GoTo EH
    wsBG.Range(wsBG.Cells(BG_FIRST_ROW, 17), wsBG.Cells(lastRow, 19)).value = outQRS

    Exit Sub
EH:
    MsgBox "Erreur dans InjectMappingFormulas_BG" & vbCrLf & _
           "Detail : " & Err.Number & " - " & Err.Description, vbCritical
End Sub

' Retourne la valeur d'une cellule du tableau Mapping.
' Renvoie "" si l'indice de colonne depasse la largeur du tableau.
Private Function GetMapColValue(ByRef arr As Variant, ByVal rowIdx As Long, _
                                 ByVal colIdx As Long, ByVal maxCol As Long) As Variant
    If colIdx >= 1 And colIdx <= maxCol Then
        GetMapColValue = arr(rowIdx, colIdx)
    Else
        GetMapColValue = ""
    End If
End Function

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
    ' CORRECTIF colonne A GestTable :
    ' On colle a la position exacte du UsedRange source (en preservant son offset
    ' de ligne et de colonne). Si le UsedRange de la feuille source commence en B1
    ' (cas de SIG, CAF, BFR, TFT dont la col A est vide), la destination respecte
    ' cet offset : la col A reste vide dans wbOut, evitant un decalage qui cassait
    ' l'import de SIG_detail (les postes etaient lus dans la mauvaise colonne).
    wsDest.Cells.Clear
    Dim ur As Range
    Set ur = wsSrc.UsedRange
    If ur Is Nothing Then Exit Sub

    Dim destCell As Range
    Set destCell = wsDest.Cells(ur.Row, ur.Column)  ' preserves row/col offset

    destCell.Resize(ur.rows.Count, ur.Columns.Count).value = ur.value
    ur.Copy
    destCell.PasteSpecial xlPasteFormats
    destCell.PasteSpecial xlPasteColumnWidths
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
