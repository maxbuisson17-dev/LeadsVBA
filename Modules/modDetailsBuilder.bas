Attribute VB_Name = "modDetailsBuilder"
' ============================================================
' modDetailsBuilder.bas
' Generation onglets BS_detail et SIG_detail dans wbOut
' Les montants proviennent du BG du wbOut (deja en valeurs, deja divises si KE)
' BS_detail  : N en col E, N-1 en col F
' SIG_detail : N en col C, N-1 en col E
' ============================================================
Option Explicit

Public Sub EnsureDetailedTabsGenerated(ByVal wbOut As Workbook)
    Build_Details_V5 wbOut
    ' Figer les valeurs des onglets detail (ConvertUsedRangeToValues)
    ConvertDetailToValues wbOut, "BS_detail"
    ConvertDetailToValues wbOut, "SIG_detail"
End Sub

Private Sub ConvertDetailToValues(ByVal wbOut As Workbook, ByVal sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wbOut.Worksheets(sheetName)
    On Error GoTo 0
    If Not ws Is Nothing Then ConvertUsedRangeToValues ws
End Sub

Public Sub Build_Details_V5(ByVal wbOut As Workbook)
    Set gDetailRowRanges = CreateObject("Scripting.Dictionary")

    Dim wsBG As Worksheet
    On Error Resume Next
    Set wsBG = wbOut.Worksheets(SH_BG)
    On Error GoTo 0
    If wsBG Is Nothing Then Exit Sub

    Dim lastR As Long
    lastR = wsBG.Cells(wsBG.Rows.Count, "A").End(xlUp).Row
    If lastR < 2 Then Exit Sub

    Dim bgArr As Variant
    bgArr = wsBG.Range("A1:Q" & lastR).Value

    Dim mapBS As Object, mapSIG As Object
    Set mapBS  = CreateObject("Scripting.Dictionary")
    Set mapSIG = CreateObject("Scripting.Dictionary")
    BuildPosteMaps bgArr, mapBS, mapSIG

    If SheetExists(wbOut, SH_BS) Then
        BuildDetailFromBase wbOut, SH_BS, "BS_detail", mapBS, True
    End If
    If SheetExists(wbOut, SH_SIG) Then
        BuildDetailFromBase wbOut, SH_SIG, "SIG_detail", mapSIG, False
    End If
End Sub

Private Sub BuildPosteMaps(ByVal bgArr As Variant, ByVal mapBS As Object, ByVal mapSIG As Object)
    ' Colonnes BG : A=Compte, B=Libelle, C=Solde N, D=Solde N-1
    '               K=Poste BS N (col 11), O=Poste BS N-1 (col 15), Q=Poste SIG (col 17)
    Dim r As Long
    Dim compte As String, lib As String
    Dim soldeN As Variant, soldeN1 As Variant
    Dim posteK As String, posteO As String, posteQ As String

    For r = 2 To UBound(bgArr, 1)
        compte = Trim$(CStr(bgArr(r, 1)))
        lib    = Trim$(CStr(bgArr(r, 2)))
        soldeN  = bgArr(r, 3)
        soldeN1 = bgArr(r, 4)
        posteK = Trim$(CStr(bgArr(r, 11)))
        posteO = Trim$(CStr(bgArr(r, 15)))
        posteQ = Trim$(CStr(bgArr(r, 17)))

        If Len(posteK) > 0 Then AddMapEntry mapBS, posteK, compte, lib, soldeN,  Empty,  True,  False
        If Len(posteO) > 0 Then AddMapEntry mapBS, posteO, compte, lib, Empty,   soldeN1, False, True
        If Len(posteQ) > 0 Then AddMapEntry mapSIG, posteQ, compte, lib, soldeN, soldeN1, True,  True
    Next r
End Sub

Private Sub AddMapEntry(ByVal mapPoste As Object, ByVal poste As String, _
                        ByVal compte As String, ByVal lib As String, _
                        ByVal nVal As Variant, ByVal n1Val As Variant, _
                        ByVal setN As Boolean, ByVal setN1 As Boolean)
    Dim key As String
    key = compte & "|" & lib
    Dim dict As Object
    If Not mapPoste.Exists(poste) Then
        Set dict = CreateObject("Scripting.Dictionary")
        mapPoste.Add poste, dict
    Else
        Set dict = mapPoste(poste)
    End If
    Dim entry As Variant
    If dict.Exists(key) Then
        entry = dict(key)
    Else
        entry = Array(compte, lib, Empty, Empty)
    End If
    If setN  Then entry(2) = nVal
    If setN1 Then entry(3) = n1Val
    dict(key) = entry
End Sub

Private Sub BuildDetailFromBase(ByVal wbOut As Workbook, ByVal baseSheetName As String, _
                                ByVal detailName As String, ByVal mapPoste As Object, ByVal isBS As Boolean)
    Dim wsBase As Worksheet, wsDetail As Worksheet
    Dim lastRow As Long, r As Long
    Dim poste As String
    Dim dictPoste As Object

    On Error Resume Next
    Set wsBase = wbOut.Worksheets(baseSheetName)
    On Error GoTo 0
    If wsBase Is Nothing Or mapPoste Is Nothing Then Exit Sub

    Application.DisplayAlerts = False
    DeleteSheetIfExists wbOut, detailName
    Application.DisplayAlerts = True

    wsBase.Copy After:=wbOut.Worksheets(wbOut.Worksheets.Count)
    Set wsDetail = wbOut.Worksheets(wbOut.Worksheets.Count)
    wsDetail.Name = detailName
    wsDetail.Cells.ClearOutline

    lastRow = wsDetail.Cells(wsDetail.Rows.Count, "B").End(xlUp).Row
    If lastRow < 1 Then Exit Sub

    For r = lastRow To 1 Step -1
        poste = Trim$(CStr(wsDetail.Cells(r, "B").Value))
        If Len(poste) > 0 And mapPoste.Exists(poste) Then
            Set dictPoste = Nothing
            On Error Resume Next
            Set dictPoste = mapPoste(poste)
            On Error GoTo 0
            If Not dictPoste Is Nothing Then
                InsertDetailBlock wsDetail, r, dictPoste, isBS
            End If
        End If
    Next r
End Sub

Private Sub InsertDetailBlock(ByVal ws As Worksheet, ByVal posteRow As Long, _
                              ByVal dict As Object, ByVal isBS As Boolean)
    Dim cnt      As Long
    Dim firstRow As Long, lastRow As Long
    Dim outArr() As Variant
    Dim i As Long, k As Variant, entry As Variant

    If ws Is Nothing Or dict Is Nothing Then Exit Sub
    cnt = dict.Count
    If cnt <= 0 Then Exit Sub

    On Error GoTo InsertEH
    firstRow = posteRow + 1
    lastRow  = posteRow + cnt

    ' Sauvegarder la bordure superieure du sous-total AVANT insertion.
    ' Avant insertion, le sous-total est a firstRow. Apres, il sera a lastRow+1.
    ' On sauvegarde aussi la bordure basse du poste (separateur potentiel herite).
    Dim stStyle  As Variant, stColor As Long, stWeight As Variant
    On Error Resume Next
    stStyle  = ws.Rows(firstRow).Borders(xlEdgeTop).LineStyle
    stColor  = ws.Rows(firstRow).Borders(xlEdgeTop).Color
    stWeight = ws.Rows(firstRow).Borders(xlEdgeTop).Weight
    On Error GoTo InsertEH

    ws.Rows(firstRow & ":" & lastRow).Insert Shift:=xlDown
    ' Apres insertion : sous-total est desormais a lastRow+1

    ReDim outArr(1 To cnt, 1 To 6)
    i = 0
    For Each k In dict.Keys
        i = i + 1
        entry = dict(k)
        outArr(i, 2) = CStr(entry(0)) & " - " & CStr(entry(1))
        If isBS Then
            outArr(i, 5) = entry(2)
            outArr(i, 6) = entry(3)
        Else
            outArr(i, 3) = entry(2)
            outArr(i, 5) = entry(3)
        End If
    Next k

    ws.Range(ws.Cells(firstRow, 1), ws.Cells(lastRow, 6)).Value = outArr

    ' Nettoyer les bordures des lignes comptes inserees
    ' (ne touche pas la bordure basse de la derniere ligne inseree)
    ClearInsertedAccountRowBorders ws, firstRow, lastRow

    ' Restaurer la bordure superieure du sous-total (maintenant a lastRow+1)
    On Error Resume Next
    If Not IsError(stStyle) And CLng(stStyle) <> xlNone Then
        With ws.Rows(lastRow + 1).Borders(xlEdgeTop)
            .LineStyle = stStyle
            .Color     = stColor
            .Weight    = stWeight
        End With
    End If
    On Error GoTo InsertEH

    ws.Rows(firstRow & ":" & lastRow).Font.Size = 9
    ws.Rows(firstRow & ":" & lastRow).Font.Italic = True
    ' Decalage visuel : col B legèrement indentee par rapport au poste parent
    ws.Range(ws.Cells(firstRow, "B"), ws.Cells(lastRow, "B")).IndentLevel = 2
    RegisterDetailRange ws, firstRow, lastRow
    ws.Rows(firstRow & ":" & lastRow).Group
    Exit Sub

InsertEH:
    modKETrace.LogKE "WARN | sheet=" & ws.Name & " | row=" & posteRow & " | Err=" & Err.Number & " | " & Err.Description, "InsertDetailBlock"
    Err.Clear
End Sub

Private Sub ClearInsertedAccountRowBorders(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    ' Efface les bordures des lignes de comptes inserees.
    ' Principe :
    '   - Supprime la bordure haute du bloc (haut de firstRow)
    '   - Supprime les separations internes entre lignes comptes (xlInsideHorizontal)
    '   - NE TOUCHE PAS xlEdgeBottom de lastRow : c est le separateur visuel
    '     avec le sous-total situe en dessous. Le restaurer explicitement est
    '     gere par InsertDetailBlock via la sauvegarde/restauration.
    On Error Resume Next
    ws.Range(ws.Cells(firstRow, 1), ws.Cells(lastRow, 26)).Borders(xlEdgeTop).LineStyle          = xlNone
    ws.Range(ws.Cells(firstRow, 1), ws.Cells(lastRow, 26)).Borders(xlInsideHorizontal).LineStyle = xlNone
    On Error GoTo 0
End Sub

Private Sub RegisterDetailRange(ByVal ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long)
    If gDetailRowRanges Is Nothing Then Set gDetailRowRanges = CreateObject("Scripting.Dictionary")
    Dim col As Collection
    If gDetailRowRanges.Exists(ws.Name) Then
        Set col = gDetailRowRanges(ws.Name)
    Else
        Set col = New Collection
        gDetailRowRanges.Add ws.Name, col
    End If
    col.Add Array(r1, r2)
End Sub

Public Sub Finalize_DetailSheets(ByVal wbOut As Workbook)
    PlaceDetailSheetAfter wbOut, "BS_detail",  SH_BS
    PlaceDetailSheetAfter wbOut, "SIG_detail", SH_SIG
    HideSheetByNameCI wbOut, SH_BG
    HideSheetByNameCI wbOut, SH_MAP
End Sub

Private Sub PlaceDetailSheetAfter(ByVal wbOut As Workbook, ByVal detailName As String, ByVal afterName As String)
    Dim wsDetail As Worksheet, wsAfter As Worksheet
    On Error Resume Next
    Set wsDetail = wbOut.Worksheets(detailName)
    Set wsAfter  = wbOut.Worksheets(afterName)
    On Error GoTo 0
    If wsDetail Is Nothing Or wsAfter Is Nothing Then Exit Sub
    wsDetail.Move After:=wsAfter
End Sub
