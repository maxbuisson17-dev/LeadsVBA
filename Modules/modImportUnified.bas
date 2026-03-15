Attribute VB_Name = "modImportUnified"
Option Explicit

Public Function ImportFile_ToBalance3Cols(ByVal path As String, Optional ByRef outInfo As String) As Variant
    Dim ext As String
    Dim arr As Variant
    Dim info As String

    On Error GoTo EH

    path = Trim$(path)
    If Len(path) = 0 Then GoTo CleanExit

    ext = ImportUnified_Ext(path)
    Select Case ext
        Case "txt", "csv", "dat"
            arr = modBalanceCreator.FEC_ToBalanceArray_FromPath(path)
            If ImportUnified_ArrayHasRows(arr) Then
                outInfo = "Import FEC texte/CSV/DAT"
                ImportFile_ToBalance3Cols = ImportUnified_Normalize3Cols(arr)
                GoTo CleanExit
            End If

            arr = ImportUnified_LoadBalanceText(path, info)
            If ImportUnified_ArrayHasRows(arr) Then
                outInfo = info
                ImportFile_ToBalance3Cols = ImportUnified_Normalize3Cols(arr)
            End If

        Case "xls", "xlsx", "xlsm"
            arr = modBalanceCreator.LoadBalanceArray_FromExcelPath(path)
            If ImportUnified_ArrayHasRows(arr) Then
                outInfo = "Import balance Excel"
                ImportFile_ToBalance3Cols = ImportUnified_Normalize3Cols(arr)
                GoTo CleanExit
            End If

            arr = modBalanceCreator.FEC_ToBalanceArray_FromExcelPath(path, info)
            If ImportUnified_ArrayHasRows(arr) Then
                outInfo = info
                ImportFile_ToBalance3Cols = ImportUnified_Normalize3Cols(arr)
            End If

        Case Else
            outInfo = "Extension non supportee: " & ext
    End Select

CleanExit:
    Exit Function
EH:
    ImportFile_ToBalance3Cols = Empty
    outInfo = "Erreur import unifie: " & Err.Description
    Resume CleanExit
End Function

Public Function ImportFile_ToBalance4Cols(ByVal path As String, Optional ByRef outInfo As String) As Variant
    ' Retourne un tableau (1..n, 1..4) : Compte / Libelle / SoldeN / SoldeN1
    ' Utilise si le fichier source a >= 4 colonnes (balance comparative).
    Dim ext As String
    Dim arrRaw As Variant
    Dim outArr() As Variant
    Dim i As Long, n As Long
    Dim acc As String

    On Error GoTo EH

    ext = ImportUnified_Ext(path)
    Select Case ext
        Case "xls", "xlsx", "xlsm"
            arrRaw = ImportUnified_LoadRaw4ColsFromExcel(path)
        Case "txt", "csv", "dat"
            arrRaw = ImportUnified_LoadRaw4ColsFromText(path)
        Case Else
            outInfo = "Extension non supportee pour import 4 cols: " & ext
            GoTo CleanExit
    End Select

    If Not ImportUnified_ArrayHasRows4(arrRaw) Then GoTo CleanExit

    n = UBound(arrRaw, 1)
    ReDim outArr(1 To n, 1 To 4)
    Dim o As Long
    For i = 1 To n
        acc = ImportUnified_KeepDigits(CStr(arrRaw(i, 1)))
        If Len(acc) = 0 Then GoTo NextI4
        o = o + 1
        outArr(o, 1) = acc
        outArr(o, 2) = ImportUnified_SanitizeLabel(CStr(arrRaw(i, 2)))
        outArr(o, 3) = Round(ImportUnified_ParseDouble(arrRaw(i, 3)), 2)
        outArr(o, 4) = Round(ImportUnified_ParseDouble(arrRaw(i, 4)), 2)
NextI4:
    Next i

    If o = 0 Then GoTo CleanExit
    If o < n Then
        ImportFile_ToBalance4Cols = ImportUnified_ShrinkRows4(outArr, o)
    Else
        ImportFile_ToBalance4Cols = outArr
    End If
    outInfo = "Import balance comparative 4 colonnes (N / N-1)"

CleanExit:
    Exit Function
EH:
    ImportFile_ToBalance4Cols = Empty
    outInfo = "Erreur ImportFile_ToBalance4Cols: " & Err.Description
    Resume CleanExit
End Function

Public Function ImportUnified_ArrayHasRows4Cols(ByVal arr As Variant) As Boolean
    On Error GoTo EH
    ImportUnified_ArrayHasRows4Cols = IsArray(arr) And UBound(arr, 1) >= 1 And UBound(arr, 2) >= 4
    Exit Function
EH:
    ImportUnified_ArrayHasRows4Cols = False
End Function

Private Function ImportUnified_ShrinkRows4(ByVal arr As Variant, ByVal nOut As Long) As Variant
    Dim r As Long, c As Long
    Dim outArr() As Variant
    ReDim outArr(1 To nOut, 1 To 4)
    For r = 1 To nOut
        For c = 1 To 4
            outArr(r, c) = arr(r, c)
        Next c
    Next r
    ImportUnified_ShrinkRows4 = outArr
End Function

Public Function ImportUnified_IsSupportedPath(ByVal path As String) As Boolean
    Select Case ImportUnified_Ext(path)
        Case "txt", "csv", "dat", "xls", "xlsx", "xlsm"
            ImportUnified_IsSupportedPath = True
    End Select
End Function

Public Function ImportFile_DetectSourceColumnCount(ByVal path As String) As Long
    Dim ext As String
    Dim f As Integer
    Dim lineText As String
    Dim delim As String
    Dim parts As Variant
    Dim wb As Workbook
    Dim ws As Worksheet

    On Error GoTo EH
    path = Trim$(path)
    If Len(path) = 0 Then GoTo CleanExit

    ext = ImportUnified_Ext(path)
    Select Case ext
        Case "txt", "csv", "dat"
            f = FreeFile
            Open path For Input As #f
            Do While Not EOF(f)
                Line Input #f, lineText
                If Len(Trim$(lineText)) > 0 Then Exit Do
            Loop
            If Len(lineText) = 0 Then GoTo CleanExit
            delim = ImportUnified_DetectDelimiter(lineText)
            If Len(delim) = 0 Then GoTo CleanExit
            parts = Split(lineText, delim)
            ImportFile_DetectSourceColumnCount = UBound(parts) - LBound(parts) + 1

        Case "xls", "xlsx", "xlsm"
            Set wb = Workbooks.Open(fileName:=path, UpdateLinks:=0, ReadOnly:=True, IgnoreReadOnlyRecommended:=True, AddToMru:=False)
            Set ws = ImportUnified_FirstNonEmptySheet(wb)
            If ws Is Nothing Then GoTo CleanExit
            ImportFile_DetectSourceColumnCount = ImportUnified_LastUsedCol(ws)
    End Select

CleanExit:
    On Error Resume Next
    If f <> 0 Then Close #f
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Exit Function
EH:
    ImportFile_DetectSourceColumnCount = 0
    Resume CleanExit
End Function

Public Function ImportFile4Cols_ToBalance3Cols(ByVal path As String, ByVal mode As eBalance4ColsMode, Optional ByRef outInfo As String) As Variant
    Dim ext As String
    Dim arrRaw As Variant

    On Error GoTo EH
    ext = ImportUnified_Ext(path)

    Select Case ext
        Case "xls", "xlsx", "xlsm"
            arrRaw = ImportUnified_LoadRaw4ColsFromExcel(path)
        Case "txt", "csv", "dat"
            arrRaw = ImportUnified_LoadRaw4ColsFromText(path)
        Case Else
            GoTo CleanExit
    End Select

    If Not ImportUnified_ArrayHasRows4(arrRaw) Then GoTo CleanExit
    ImportFile4Cols_ToBalance3Cols = ImportUnified_Convert4To3(arrRaw, mode)
    If ImportUnified_ArrayHasRows(ImportFile4Cols_ToBalance3Cols) Then
        Select Case mode
            Case b4NN1: outInfo = "Balance 4 colonnes traitee en mode N/N-1"
            Case b4NN1_ColD: outInfo = "Balance 4 colonnes traitee en mode N/N-1 (Solde N en col D)"
            Case b4DebitCredit: outInfo = "Balance 4 colonnes traitee en mode Debit/Credit (C-D)"
        End Select
    End If

CleanExit:
    Exit Function
EH:
    ImportFile4Cols_ToBalance3Cols = Empty
    outInfo = "Erreur conversion 4->3: " & Err.Description
    Resume CleanExit
End Function

Private Function ImportUnified_LoadBalanceText(ByVal path As String, Optional ByRef outInfo As String) As Variant
    Dim f As Integer
    Dim lineText As String
    Dim headerLine As String
    Dim delim As String
    Dim headers() As String
    Dim parts() As String
    Dim idxAcc As Long
    Dim idxLib As Long
    Dim idxSolde As Long
    Dim idxDebit As Long
    Dim idxCredit As Long
    Dim maxIdx As Long
    Dim acc As String
    Dim lib As String
    Dim solde As Double
    Dim dict As Object
    Dim rec As Variant
    Dim keys As Variant
    Dim outArr() As Variant
    Dim i As Long
    Dim rowOut As Long

    On Error GoTo EH

    If Len(Trim$(path)) = 0 Then GoTo CleanExit

    f = FreeFile
    Open path For Input As #f

    Do While Not EOF(f)
        Line Input #f, lineText
        If Len(Trim$(lineText)) > 0 Then
            headerLine = lineText
            Exit Do
        End If
    Loop

    If Len(headerLine) = 0 Then GoTo CleanExit
    delim = ImportUnified_DetectDelimiter(headerLine)
    If Len(delim) = 0 Then GoTo CleanExit

    headers = Split(headerLine, delim)
    headers(LBound(headers)) = ImportUnified_StripBom(headers(LBound(headers)))

    ImportUnified_FindBalanceColumns headers, idxAcc, idxLib, idxSolde, idxDebit, idxCredit
    If idxAcc < 0 Then idxAcc = 0
    If idxLib < 0 Then idxLib = 1

    maxIdx = idxAcc
    If idxLib > maxIdx Then maxIdx = idxLib
    If idxSolde > maxIdx Then maxIdx = idxSolde
    If idxDebit > maxIdx Then maxIdx = idxDebit
    If idxCredit > maxIdx Then maxIdx = idxCredit

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbBinaryCompare

    Do While Not EOF(f)
        Line Input #f, lineText
        If Len(Trim$(lineText)) = 0 Then GoTo NextLine

        parts = Split(lineText, delim)
        If UBound(parts) < maxIdx Then GoTo NextLine

        acc = ImportUnified_KeepDigits(ImportUnified_GetField(parts, idxAcc))
        If Len(acc) = 0 Then GoTo NextLine
        lib = ImportUnified_SanitizeLabel(ImportUnified_GetField(parts, idxLib))

        If idxSolde >= 0 Then
            solde = ImportUnified_ParseDouble(ImportUnified_GetField(parts, idxSolde))
        Else
            solde = ImportUnified_ParseDouble(ImportUnified_GetField(parts, idxDebit)) - ImportUnified_ParseDouble(ImportUnified_GetField(parts, idxCredit))
        End If

        If dict.Exists(acc) Then
            rec = dict(acc)
            If Len(CStr(rec(0))) = 0 And Len(lib) > 0 Then rec(0) = lib
            rec(1) = CDbl(rec(1)) + solde
            dict(acc) = rec
        Else
            dict.Add acc, Array(lib, solde)
        End If
NextLine:
    Loop

    If dict.Count = 0 Then GoTo CleanExit

    keys = dict.keys
    ImportUnified_SortStringArray keys

    ReDim outArr(1 To dict.Count, 1 To 3)
    For i = LBound(keys) To UBound(keys)
        rowOut = (i - LBound(keys)) + 1
        rec = dict(keys(i))
        outArr(rowOut, 1) = CStr(keys(i))
        outArr(rowOut, 2) = CStr(rec(0))
        outArr(rowOut, 3) = CDbl(rec(1))
    Next i

    outInfo = "Import balance texte/CSV/DAT (fallback)"
    ImportUnified_LoadBalanceText = outArr

CleanExit:
    On Error Resume Next
    If f <> 0 Then Close #f
    On Error GoTo 0
    Exit Function
EH:
    ImportUnified_LoadBalanceText = Empty
    Resume CleanExit
End Function

Private Function ImportUnified_Normalize3Cols(ByVal arr As Variant) As Variant
    Dim outArr() As Variant
    Dim i As Long
    Dim n As Long

    On Error GoTo EH
    If Not ImportUnified_ArrayHasRows(arr) Then GoTo CleanExit

    n = UBound(arr, 1)
    ReDim outArr(1 To n, 1 To 3)
    For i = 1 To n
        outArr(i, 1) = ImportUnified_KeepDigits(CStr(arr(i, 1)))
        outArr(i, 2) = ImportUnified_SanitizeLabel(CStr(arr(i, 2)))
        outArr(i, 3) = Round(ImportUnified_ParseDouble(arr(i, 3)), 2)
    Next i
    ImportUnified_Normalize3Cols = outArr
    GoTo CleanExit
EH:
    Debug.Print "ImportUnified_Normalize3Cols error " & Err.Number & " : " & Err.Description
    ImportUnified_Normalize3Cols = Empty
    Resume CleanExit
CleanExit:
    Exit Function
End Function

Private Function ImportUnified_Ext(ByVal path As String) As String
    Dim p As Long
    path = Trim$(path)
    p = InStrRev(path, ".")
    If p <= 0 Or p = Len(path) Then Exit Function
    ImportUnified_Ext = LCase$(Mid$(path, p + 1))
End Function

Public Function ImportUnified_ArrayHasRows(ByVal arr As Variant) As Boolean
    On Error GoTo EH
    ImportUnified_ArrayHasRows = IsArray(arr) And UBound(arr, 1) >= 1 And UBound(arr, 2) >= 3
    GoTo CleanExit
EH:
    Debug.Print "ImportUnified_ArrayHasRows error " & Err.Number & " : " & Err.Description
    ImportUnified_ArrayHasRows = False
    Resume CleanExit
CleanExit:
    Exit Function
End Function

Private Function ImportUnified_DetectDelimiter(ByVal headerLine As String) As String
    Dim candidates As Variant
    Dim parts As Variant
    Dim i As Long
    Dim bestCols As Long
    Dim colCount As Long
    Dim d As String

    candidates = Array(vbTab, ";", "|", ",")
    For i = LBound(candidates) To UBound(candidates)
        d = CStr(candidates(i))
        parts = Split(headerLine, d)
        colCount = UBound(parts) - LBound(parts) + 1
        If colCount > bestCols Then
            bestCols = colCount
            ImportUnified_DetectDelimiter = d
        End If
    Next i
    If bestCols <= 1 Then ImportUnified_DetectDelimiter = vbNullString
End Function

Private Sub ImportUnified_FindBalanceColumns(ByRef headers() As String, ByRef idxAcc As Long, ByRef idxLib As Long, ByRef idxSolde As Long, ByRef idxDebit As Long, ByRef idxCredit As Long)
    Dim i As Long
    Dim h As String

    idxAcc = -1
    idxLib = -1
    idxSolde = -1
    idxDebit = -1
    idxCredit = -1

    For i = LBound(headers) To UBound(headers)
        h = ImportUnified_NormalizeHeader(CStr(headers(i)))
        Select Case h
            Case "compte", "comptenum", "comptenumero"
                If idxAcc < 0 Then idxAcc = i
            Case "libelle", "comptelib"
                If idxLib < 0 Then idxLib = i
            Case "solde", "solden", "solden1"
                If idxSolde < 0 Then idxSolde = i
            Case "debit", "totaldebit", "soldedebit"
                If idxDebit < 0 Then idxDebit = i
            Case "credit", "totalcredit", "soldecredit"
                If idxCredit < 0 Then idxCredit = i
        End Select
    Next i

    If idxSolde < 0 And idxDebit >= 0 And idxCredit >= 0 Then
        idxSolde = -1
    ElseIf idxSolde < 0 Then
        If UBound(headers) >= 2 Then idxSolde = 2
    End If
End Sub

Private Function ImportUnified_GetField(ByRef parts() As String, ByVal idx As Long) As String
    If idx < LBound(parts) Or idx > UBound(parts) Then Exit Function
    ImportUnified_GetField = CStr(parts(idx))
End Function

Private Function ImportUnified_KeepDigits(ByVal s As String) As String
    Dim i As Long
    Dim ch As String

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then ImportUnified_KeepDigits = ImportUnified_KeepDigits & ch
    Next i
End Function

Private Function ImportUnified_ParseDouble(ByVal v As Variant) As Double
    Dim s As String

    On Error GoTo EH
    If IsEmpty(v) Then Exit Function
    If IsNumeric(v) Then
        ImportUnified_ParseDouble = CDbl(v)
        Exit Function
    End If

    s = CStr(v)
    s = Replace(s, Chr$(160), " ")
    s = Replace(s, " ", "")
    s = Replace(s, "'", "")
    If InStr(1, s, ",", vbBinaryCompare) > 0 And InStr(1, s, ".", vbBinaryCompare) > 0 Then
        If InStrRev(s, ",") > InStrRev(s, ".") Then
            s = Replace(s, ".", "")
            s = Replace(s, ",", ".")
        Else
            s = Replace(s, ",", "")
        End If
    ElseIf InStr(1, s, ",", vbBinaryCompare) > 0 Then
        s = Replace(s, ",", ".")
    End If
    If IsNumeric(s) Then ImportUnified_ParseDouble = CDbl(s)
    GoTo CleanExit
EH:
    Debug.Print "ImportUnified_ParseDouble error " & Err.Number & " : " & Err.Description
    ImportUnified_ParseDouble = 0#
    Resume CleanExit
CleanExit:
    Exit Function
End Function

Private Function ImportUnified_StripBom(ByVal s As String) As String
    If Len(s) > 0 Then
        If AscW(Left$(s, 1)) = 65279 Then
            ImportUnified_StripBom = Mid$(s, 2)
            Exit Function
        End If
    End If
    ImportUnified_StripBom = s
End Function

Private Function ImportUnified_NormalizeHeader(ByVal s As String) As String
    Dim t As String
    t = LCase$(Trim$(ImportUnified_StripBom(s)))
    t = Replace(t, Chr$(160), "")
    t = Replace(t, " ", "")
    t = Replace(t, "_", "")
    t = Replace(t, "-", "")
    t = Replace(t, ChrW$(233), "e")
    t = Replace(t, ChrW$(232), "e")
    t = Replace(t, ChrW$(234), "e")
    t = Replace(t, ChrW$(235), "e")
    t = Replace(t, ChrW$(224), "a")
    t = Replace(t, ChrW$(226), "a")
    t = Replace(t, ChrW$(238), "i")
    t = Replace(t, ChrW$(239), "i")
    t = Replace(t, ChrW$(244), "o")
    t = Replace(t, ChrW$(246), "o")
    t = Replace(t, ChrW$(249), "u")
    t = Replace(t, ChrW$(251), "u")
    t = Replace(t, ChrW$(252), "u")
    t = Replace(t, ChrW$(231), "c")
    ImportUnified_NormalizeHeader = t
End Function

Private Function ImportUnified_SanitizeLabel(ByVal s As String) As String
    Dim t As String
    t = Trim$(s)
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")
    Do While InStr(1, t, "  ", vbBinaryCompare) > 0
        t = Replace(t, "  ", " ")
    Loop
    ImportUnified_SanitizeLabel = t
End Function

Private Sub ImportUnified_SortStringArray(ByRef arr As Variant)
    Dim i As Long
    Dim j As Long
    Dim tmp As String

    On Error GoTo EH
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If CStr(arr(j)) < CStr(arr(i)) Then
                tmp = CStr(arr(i))
                arr(i) = arr(j)
                arr(j) = tmp
            End If
        Next j
    Next i
    GoTo CleanExit
EH:
    Debug.Print "ImportUnified_SortStringArray error " & Err.Number & " : " & Err.Description
    Resume CleanExit
CleanExit:
    Exit Sub
End Sub

Private Function ImportUnified_LoadRaw4ColsFromExcel(ByVal path As String) As Variant
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim arr As Variant
    Dim startRow As Long
    Dim n As Long
    Dim i As Long
    Dim outArr() As Variant

    On Error GoTo EH
    Set wb = Workbooks.Open(fileName:=path, UpdateLinks:=0, ReadOnly:=True, IgnoreReadOnlyRecommended:=True, AddToMru:=False)
    Set ws = ImportUnified_FirstNonEmptySheet(wb)
    If ws Is Nothing Then GoTo CleanExit

    lastRow = ImportUnified_LastUsedRow(ws)
    If lastRow < 1 Then GoTo CleanExit

    arr = ws.Range("A1:D" & lastRow).Value2
    startRow = IIf(modBalanceCreator.Balance_DetectHeaderRow(arr), 2, 1)
    If startRow > UBound(arr, 1) Then GoTo CleanExit

    n = UBound(arr, 1) - startRow + 1
    ReDim outArr(1 To n, 1 To 4)
    For i = 1 To n
        outArr(i, 1) = arr(startRow + i - 1, 1)
        outArr(i, 2) = arr(startRow + i - 1, 2)
        outArr(i, 3) = arr(startRow + i - 1, 3)
        outArr(i, 4) = arr(startRow + i - 1, 4)
    Next i
    ImportUnified_LoadRaw4ColsFromExcel = outArr

CleanExit:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Exit Function
EH:
    ImportUnified_LoadRaw4ColsFromExcel = Empty
    Resume CleanExit
End Function

Private Function ImportUnified_LoadRaw4ColsFromText(ByVal path As String) As Variant
    Dim f As Integer
    Dim lineText As String
    Dim delim As String
    Dim rows As Collection
    Dim parts As Variant
    Dim vals(1 To 4) As Variant
    Dim i As Long
    Dim arr() As Variant
    Dim hasHeader As Boolean

    On Error GoTo EH

    Set rows = New Collection
    f = FreeFile
    Open path For Input As #f

    Do While Not EOF(f)
        Line Input #f, lineText
        If Len(Trim$(lineText)) > 0 Then
            delim = ImportUnified_DetectDelimiter(lineText)
            Exit Do
        End If
    Loop
    If Len(delim) = 0 Then GoTo CleanExit

    parts = Split(lineText, delim)
    hasHeader = ImportUnified_LineLooksLikeHeader(parts)

    If Not hasHeader Then
        For i = 1 To 4
            If i <= UBound(parts) - LBound(parts) + 1 Then vals(i) = CStr(parts(LBound(parts) + i - 1))
        Next i
        rows.Add Array(vals(1), vals(2), vals(3), vals(4))
    End If

    Do While Not EOF(f)
        Line Input #f, lineText
        If Len(Trim$(lineText)) = 0 Then GoTo NextL
        parts = Split(lineText, delim)
        If UBound(parts) - LBound(parts) + 1 < 4 Then GoTo NextL
        rows.Add Array(CStr(parts(LBound(parts))), CStr(parts(LBound(parts) + 1)), CStr(parts(LBound(parts) + 2)), CStr(parts(LBound(parts) + 3)))
NextL:
    Loop

    If rows.Count = 0 Then GoTo CleanExit
    ReDim arr(1 To rows.Count, 1 To 4)
    For i = 1 To rows.Count
        vals(1) = rows(i)(0): vals(2) = rows(i)(1): vals(3) = rows(i)(2): vals(4) = rows(i)(3)
        arr(i, 1) = vals(1)
        arr(i, 2) = vals(2)
        arr(i, 3) = vals(3)
        arr(i, 4) = vals(4)
    Next i
    ImportUnified_LoadRaw4ColsFromText = arr

CleanExit:
    On Error Resume Next
    If f <> 0 Then Close #f
    On Error GoTo 0
    Exit Function
EH:
    ImportUnified_LoadRaw4ColsFromText = Empty
    Resume CleanExit
End Function

Private Function ImportUnified_Convert4To3(ByVal arr4 As Variant, ByVal mode As eBalance4ColsMode) As Variant
    Dim outArr() As Variant
    Dim i As Long
    Dim o As Long
    Dim acc As String
    Dim lib As String
    Dim solde As Double

    On Error GoTo EH
    If Not ImportUnified_ArrayHasRows4(arr4) Then GoTo CleanExit

    ReDim outArr(1 To UBound(arr4, 1), 1 To 3)
    For i = 1 To UBound(arr4, 1)
        acc = ImportUnified_KeepDigits(CStr(arr4(i, 1)))
        If Len(acc) = 0 Then GoTo NextI
        lib = ImportUnified_SanitizeLabel(CStr(arr4(i, 2)))

        Select Case mode
            Case b4DebitCredit
                solde = ImportUnified_ParseDouble(arr4(i, 3)) - ImportUnified_ParseDouble(arr4(i, 4))
            Case b4NN1_ColD
                solde = ImportUnified_ParseDouble(arr4(i, 4))
            Case Else
                solde = ImportUnified_ParseDouble(arr4(i, 3))
        End Select

        o = o + 1
        outArr(o, 1) = acc
        outArr(o, 2) = lib
        outArr(o, 3) = Round(solde, 2)
NextI:
    Next i

    If o = 0 Then GoTo CleanExit
    If o < UBound(outArr, 1) Then
        ImportUnified_Convert4To3 = ImportUnified_ShrinkRows3(outArr, o)
    Else
        ImportUnified_Convert4To3 = outArr
    End If

CleanExit:
    Exit Function
EH:
    ImportUnified_Convert4To3 = Empty
    Resume CleanExit
End Function

Private Function ImportUnified_LineLooksLikeHeader(ByVal parts As Variant) As Boolean
    Dim p0 As String, p1 As String
    On Error GoTo EH
    p0 = ImportUnified_NormalizeHeader(CStr(parts(LBound(parts))))
    p1 = ImportUnified_NormalizeHeader(CStr(parts(LBound(parts) + 1)))
    If p0 = "compte" Or p0 = "comptenum" Or p0 = "comptenumero" Then ImportUnified_LineLooksLikeHeader = True
    If p1 = "libelle" Or p1 = "comptelib" Then ImportUnified_LineLooksLikeHeader = True
    Exit Function
EH:
    ImportUnified_LineLooksLikeHeader = False
End Function

Private Function ImportUnified_ArrayHasRows4(ByVal arr As Variant) As Boolean
    On Error GoTo EH
    ImportUnified_ArrayHasRows4 = IsArray(arr) And UBound(arr, 1) >= 1 And UBound(arr, 2) >= 4
    Exit Function
EH:
    ImportUnified_ArrayHasRows4 = False
End Function

Private Function ImportUnified_ShrinkRows3(ByVal arr As Variant, ByVal nOut As Long) As Variant
    Dim r As Long, c As Long
    Dim outArr() As Variant
    ReDim outArr(1 To nOut, 1 To 3)
    For r = 1 To nOut
        For c = 1 To 3
            outArr(r, c) = arr(r, c)
        Next c
    Next r
    ImportUnified_ShrinkRows3 = outArr
End Function

Private Function ImportUnified_FirstNonEmptySheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ImportUnified_LastUsedRow(ws) > 0 And ImportUnified_LastUsedCol(ws) > 0 Then
            Set ImportUnified_FirstNonEmptySheet = ws
            Exit Function
        End If
    Next ws
End Function

Private Function ImportUnified_LastUsedRow(ByVal ws As Worksheet) As Long
    Dim r As Range
    On Error Resume Next
    Set r = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    On Error GoTo 0
    If Not r Is Nothing Then ImportUnified_LastUsedRow = r.Row
End Function

Private Function ImportUnified_LastUsedCol(ByVal ws As Worksheet) As Long
    Dim r As Range
    On Error Resume Next
    Set r = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    On Error GoTo 0
    If Not r Is Nothing Then ImportUnified_LastUsedCol = r.Column
End Function
