Attribute VB_Name = "modBalanceCreator"
' ============================================================
' modBalanceCreator.bas
' Module unique de creation et normalisation de balances
' Fusionne : BalanceMapping + BGCompil + BGtoBG + FecBG + TransfoBG
' Remplace ces 5 modules - supprimer les .bas originals du projet
' ============================================================
Option Explicit

' ============================================================
' 1. DETECTION FORMAT BALANCE
' ============================================================

Public Function DetectBalanceFormat(ByVal ws As Worksheet) As eBalanceFormat
    Dim fmt As eBalanceFormat
    Dim bHdr As Boolean
    Dim accCol As Long, lblCol As Long, sldCol As Long, dbtCol As Long, crdCol As Long
    Dim lRow As Long, lCol As Long
    DetectColumns_ByRef ws, fmt, bHdr, accCol, lblCol, sldCol, dbtCol, crdCol, lRow, lCol
    DetectBalanceFormat = fmt
End Function

Public Function DetectBalanceFormat_FromPath(ByVal filePath As String) As eBalanceFormat
    Dim wb As Workbook
    Dim ws As Worksheet

    On Error GoTo EH

    Set wb = Workbooks.Open(fileName:=filePath, UpdateLinks:=0, ReadOnly:=True, IgnoreReadOnlyRecommended:=True, AddToMru:=False)
    Set ws = GetFirstNonEmptySheet(wb)
    If ws Is Nothing Then GoTo CleanExit

    DetectBalanceFormat_FromPath = DetectBalanceFormat(ws)

CleanExit:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Exit Function
EH:
    DetectBalanceFormat_FromPath = bfUnknown
    Resume CleanExit
End Function


' ============================================================
' 2. CHARGEMENT BALANCE DEPUIS FICHIER
' ============================================================

Public Function LoadBalanceArray_FromExcelPath(ByVal path As String) As Variant
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim arr As Variant

    On Error GoTo EH

    Set wb = Workbooks.Open(fileName:=path, UpdateLinks:=0, ReadOnly:=True, IgnoreReadOnlyRecommended:=True, AddToMru:=False)
    Set ws = GetFirstNonEmptySheet(wb)

    If ws Is Nothing Then GoTo CleanExit

    arr = TransformSheetToBalanceArray(ws)

    If IsEmpty(arr) Or Not IsArray(arr) Then
        LoadBalanceArray_FromExcelPath = Empty
    Else
        LoadBalanceArray_FromExcelPath = Normalize3Cols(arr)
    End If

CleanExit:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Exit Function
EH:
    LoadBalanceArray_FromExcelPath = Empty
    Resume CleanExit
End Function

Public Function BalanceExcel_ToBalanceArray_FromPath(ByVal filePath As String, Optional ByRef outInfo As String) As Variant
    Dim arr As Variant
    ' NOTE: fmt est de type eBalanceFormat (Enum global) pour eviter "Type defini par l'utilisateur non defini"
    Dim fmt As eBalanceFormat

    fmt = DetectBalanceFormat_FromPath(filePath)
    arr = LoadBalanceArray_FromExcelPath(filePath)

    If IsEmpty(arr) Or Not IsArray(arr) Then
        outInfo = "Format non reconnu: colonnes compte/libelle/solde introuvables"
        BalanceExcel_ToBalanceArray_FromPath = Empty
        Exit Function
    End If

    Select Case fmt
        Case bf3ColsSolde
            outInfo = "CAS A: 3 colonnes (solde direct) - aucun recalcul debit/credit"
        Case bf4ColsDebitCredit
            outInfo = "CAS B: 4 colonnes debit/credit - solde calcule"
        Case bfHeuristicMapped
            outInfo = "CAS C: mapping heuristique applique"
        Case Else
            outInfo = "Format detecte partiellement"
    End Select

    BalanceExcel_ToBalanceArray_FromPath = arr
End Function

Public Function ExcelBalance_ToBalanceArray_FromPath(ByVal pathXls As String) As Variant
    ExcelBalance_ToBalanceArray_FromPath = BalanceExcel_ToBalanceArray_FromPath(pathXls)
End Function

Public Function FEC_ToBalanceArray_FromPath(ByVal pathTxt As String) As Variant
    Dim f As Integer
    Dim lineText As String
    Dim delim As String
    Dim headers() As String
    Dim idxCompte As Long
    Dim idxLib As Long
    Dim idxDebit As Long
    Dim idxCredit As Long
    Dim dict As Object
    Dim parts() As String
    Dim acc As String
    Dim lib As String
    Dim solde As Double
    Dim rec As Variant
    Dim keys As Variant
    Dim outArr() As Variant
    Dim i As Long
    Dim outRow As Long
    Dim maxIdx As Long

    On Error GoTo EH

    If Len(Trim$(pathTxt)) = 0 Then Exit Function

    f = FreeFile
    Open pathTxt For Input As #f

    If EOF(f) Then GoTo CleanExit
    Line Input #f, lineText
    Do While Len(Trim$(lineText)) = 0 And Not EOF(f)
        Line Input #f, lineText
    Loop

    delim = DetectDelimiter2(lineText)
    If Len(delim) = 0 Then GoTo CleanExit

    headers = Split(lineText, delim)
    headers(LBound(headers)) = StripUtf8Bom(headers(LBound(headers)))
    FindFecColumns headers, idxCompte, idxLib, idxDebit, idxCredit
    If idxCompte < 0 Or idxLib < 0 Or idxDebit < 0 Or idxCredit < 0 Then GoTo CleanExit

    maxIdx = idxCompte
    If idxLib > maxIdx Then maxIdx = idxLib
    If idxDebit > maxIdx Then maxIdx = idxDebit
    If idxCredit > maxIdx Then maxIdx = idxCredit

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbBinaryCompare

    Do While Not EOF(f)
        Line Input #f, lineText
        If Len(Trim$(lineText)) = 0 Then GoTo NextLine

        parts = Split(lineText, delim)
        If UBound(parts) < maxIdx Then GoTo NextLine

        acc = KeepDigits3(GetField(parts, idxCompte))
        If Len(acc) = 0 Then GoTo NextLine

        lib = SanitizeLabel3(GetField(parts, idxLib))
        solde = ParseDoubleFR3(GetField(parts, idxDebit)) - ParseDoubleFR3(GetField(parts, idxCredit))

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
    SortStringArray3 keys

    ReDim outArr(1 To dict.Count, 1 To 3)
    For i = LBound(keys) To UBound(keys)
        outRow = (i - LBound(keys)) + 1
        rec = dict(keys(i))
        outArr(outRow, 1) = CStr(keys(i))
        outArr(outRow, 2) = CStr(rec(0))
        outArr(outRow, 3) = CDbl(rec(1))
    Next i

    FEC_ToBalanceArray_FromPath = outArr

CleanExit:
    On Error Resume Next
    If f <> 0 Then Close #f
    On Error GoTo 0
    Exit Function

EH:
    FEC_ToBalanceArray_FromPath = Empty
    Resume CleanExit
End Function

Public Function FEC_ToBalanceArray_FromExcelPath(ByVal pathXls As String, Optional ByRef outInfo As String) As Variant
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim headers() As String
    Dim delim As String
    Dim headerLine As String
    Dim idxCompte As Long
    Dim idxLib As Long
    Dim idxDebit As Long
    Dim idxCredit As Long
    Dim maxIdx As Long
    Dim dict As Object
    Dim r As Long
    Dim parts() As String
    Dim acc As String
    Dim lib As String
    Dim solde As Double
    Dim rec As Variant
    Dim keys As Variant
    Dim outArr() As Variant
    Dim i As Long
    Dim outRow As Long
    Dim parsedRows As Long

    On Error GoTo EH

    If Len(Trim$(pathXls)) = 0 Then Exit Function

    Set wb = Workbooks.Open(fileName:=pathXls, UpdateLinks:=0, ReadOnly:=True, IgnoreReadOnlyRecommended:=True, AddToMru:=False)
    Set ws = Fec_FirstNonEmptySheet(wb)
    If ws Is Nothing Then GoTo CleanExit

    lastRow = Fec_LastUsedRow(ws)
    lastCol = Fec_LastUsedCol(ws)
    If lastRow < 2 Or lastCol < 1 Then GoTo CleanExit

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbBinaryCompare

    ' Cas special: CSV importe dans Excel en une seule colonne.
    If lastCol = 1 Then
        headerLine = CStr(ws.Cells(1, 1).Value2)
        headerLine = StripUtf8Bom(headerLine)
        delim = DetectDelimiter2(headerLine)
        If Len(delim) > 0 Then
            headers = Split(headerLine, delim)
            headers(LBound(headers)) = StripUtf8Bom(headers(LBound(headers)))
            FindFecColumns headers, idxCompte, idxLib, idxDebit, idxCredit
            If idxCompte < 0 Or idxLib < 0 Or idxDebit < 0 Or idxCredit < 0 Then GoTo CleanExit

            maxIdx = idxCompte
            If idxLib > maxIdx Then maxIdx = idxLib
            If idxDebit > maxIdx Then maxIdx = idxDebit
            If idxCredit > maxIdx Then maxIdx = idxCredit

            For r = 2 To lastRow
                If Len(Trim$(CStr(ws.Cells(r, 1).Value2))) = 0 Then GoTo NextR_OneCol
                parts = Split(CStr(ws.Cells(r, 1).Value2), delim)
                If UBound(parts) < maxIdx Then GoTo NextR_OneCol

                acc = KeepDigits3(GetField(parts, idxCompte))
                If Len(acc) = 0 Then GoTo NextR_OneCol

                lib = SanitizeLabel3(GetField(parts, idxLib))
                solde = ParseDoubleFR3(GetField(parts, idxDebit)) - ParseDoubleFR3(GetField(parts, idxCredit))
                parsedRows = parsedRows + 1
                Fec_Accumulate dict, acc, lib, solde
NextR_OneCol:
            Next r
        End If
    End If

    ' Cas normal Excel: colonnes separees.
    If dict.Count = 0 Then
        ReDim headers(0 To lastCol - 1)
        For i = 1 To lastCol
            headers(i - 1) = CStr(ws.Cells(1, i).Value2)
        Next i
        headers(LBound(headers)) = StripUtf8Bom(headers(LBound(headers)))

        FindFecColumns headers, idxCompte, idxLib, idxDebit, idxCredit
        If idxCompte < 0 Or idxLib < 0 Or idxDebit < 0 Or idxCredit < 0 Then GoTo CleanExit

        For r = 2 To lastRow
            acc = KeepDigits3(CStr(ws.Cells(r, idxCompte + 1).Value2))
            If Len(acc) = 0 Then GoTo NextR_Normal

            lib = SanitizeLabel3(CStr(ws.Cells(r, idxLib + 1).Value2))
            solde = ParseDoubleFR3(ws.Cells(r, idxDebit + 1).Value2) - ParseDoubleFR3(ws.Cells(r, idxCredit + 1).Value2)
            parsedRows = parsedRows + 1
            Fec_Accumulate dict, acc, lib, solde
NextR_Normal:
        Next r
    End If

    If dict.Count = 0 Then GoTo CleanExit

    keys = dict.keys
    SortStringArray3 keys

    ReDim outArr(1 To dict.Count, 1 To 3)
    For i = LBound(keys) To UBound(keys)
        outRow = (i - LBound(keys)) + 1
        rec = dict(keys(i))
        outArr(outRow, 1) = CStr(keys(i))
        outArr(outRow, 2) = CStr(rec(0))
        outArr(outRow, 3) = CDbl(rec(1))
    Next i

    outInfo = "FEC Excel detecte : " & CStr(parsedRows) & " lignes, " & CStr(dict.Count) & " comptes"
    FEC_ToBalanceArray_FromExcelPath = outArr

CleanExit:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Exit Function
EH:
    FEC_ToBalanceArray_FromExcelPath = Empty
    Resume CleanExit
End Function

Public Function ImporterFECtoBG(Optional ByRef outSelectedPath As String, Optional ByVal pathTxt As String = vbNullString) As Variant
    Dim vPath As Variant
    Dim filePath As String
    Dim arr As Variant

    On Error GoTo EH

    filePath = Trim$(pathTxt)
    If Len(filePath) = 0 Then
        vPath = Application.GetOpenFilename("Fichiers texte (*.txt;*.csv;*.dat),*.txt;*.csv;*.dat", , "Selectionner le fichier FEC")
        If VarType(vPath) = vbBoolean Then
            ImporterFECtoBG = Empty
            GoTo CleanExit
        End If
        filePath = Trim$(CStr(vPath))
    End If

    If Len(filePath) = 0 Then
        ImporterFECtoBG = Empty
        GoTo CleanExit
    End If

    outSelectedPath = filePath

    arr = FEC_ToBalanceArray_FromPath(filePath)
    If Not FecBG_ArrayHasRows(arr) Then
        Err.Raise vbObjectError + 2301, "modBalanceCreator.ImporterFECtoBG", "Aucune donnee exploitable detectee dans le FEC."
    End If

    ImporterFECtoBG = arr
    GoTo CleanExit

EH:
    Debug.Print "ImporterFECtoBG error " & Err.Number & " : " & Err.Description
    Err.Raise Err.Number, "modBalanceCreator.ImporterFECtoBG", Err.Description
    Resume CleanExit
CleanExit:
    Exit Function
End Function

Public Function Balance_LoadFromExcelPath(ByVal path As String) As Variant
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim dataStartRow As Long
    Dim colAccount As Long
    Dim colLabel As Long
    Dim colSolde As Long
    Dim colDebit As Long
    Dim colCredit As Long
    Dim arrData As Variant
    Dim i As Long
    Dim rawAcc As String
    Dim key As String
    Dim lib As String
    Dim amount As Double
    Dim dict As Object
    Dim rec As Variant
    Dim keys As Variant
    Dim outArr() As Variant

    On Error GoTo EH

    Set wb = Workbooks.Open(fileName:=path, UpdateLinks:=0, ReadOnly:=True, IgnoreReadOnlyRecommended:=True, AddToMru:=False)
    Set ws = GetFirstNonEmptySheet(wb)
    If ws Is Nothing Then GoTo CleanExit

    lastRow = GetLastUsedRow(ws)
    lastCol = GetLastUsedCol(ws)
    If lastRow = 0 Or lastCol = 0 Then GoTo CleanExit

    DetectBalanceColumns ws, lastCol, dataStartRow, colAccount, colLabel, colSolde, colDebit, colCredit
    If colAccount = 0 Then colAccount = 1
    If colLabel = 0 Then colLabel = 2
    If colAccount > lastCol Then colAccount = 1
    If colLabel > lastCol Then colLabel = IIf(lastCol >= 2, 2, 1)
    If colSolde > lastCol Then colSolde = 0
    If colDebit > lastCol Then colDebit = 0
    If colCredit > lastCol Then colCredit = 0

    If lastRow < dataStartRow Then GoTo CleanExit

    arrData = ws.Range(ws.Cells(dataStartRow, 1), ws.Cells(lastRow, lastCol)).Value2
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbBinaryCompare

    For i = 1 To UBound(arrData, 1)
        rawAcc = CStr(Nz2(arrData(i, colAccount), ""))
        key = modUtils.KeepDigits(rawAcc)
        If Len(key) = 0 Then GoTo NextI

        lib = SanitizeLabel2(CStr(Nz2(arrData(i, colLabel), "")))

        If colSolde > 0 Then
            amount = ParseDoubleFR2(arrData(i, colSolde))
        ElseIf colDebit > 0 And colCredit > 0 Then
            amount = ParseDoubleFR2(arrData(i, colDebit)) - ParseDoubleFR2(arrData(i, colCredit))
        ElseIf lastCol >= 3 Then
            amount = ParseDoubleFR2(arrData(i, 3))
        Else
            amount = 0#
        End If

        If dict.Exists(key) Then
            rec = dict(key)
            If Len(CStr(rec(0))) = 0 And Len(lib) > 0 Then rec(0) = lib
            rec(1) = CDbl(rec(1)) + amount
            dict(key) = rec
        Else
            dict.Add key, Array(lib, amount)
        End If
NextI:
    Next i

    If dict.Count = 0 Then GoTo CleanExit

    keys = dict.keys
    ReDim outArr(1 To dict.Count, 1 To 3)
    For i = 0 To dict.Count - 1
        rec = dict(keys(i))
        outArr(i + 1, 1) = CStr(keys(i))
        outArr(i + 1, 2) = CStr(rec(0))
        outArr(i + 1, 3) = CDbl(rec(1))
    Next i

    Balance_LoadFromExcelPath = outArr

CleanExit:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Exit Function

EH:
    Balance_LoadFromExcelPath = Empty
    Resume CleanExit
End Function

Public Function Balance_LoadFromFecPath(ByVal pathTxt As String) As Variant
    Balance_LoadFromFecPath = FEC_ToBalanceArray_FromPath(pathTxt)
End Function


' ============================================================
' 3. TRANSFORMATION ET NORMALISATION
' ============================================================

Public Function TransformerBG_Array(ByVal arrIn As Variant) As Variant
    Dim n As Long
    Dim i As Long
    Dim outArr() As Variant
    Dim maxLen As Long
    Dim accDigits As String
    Dim lib As String
    Dim vN As Double
    Dim vN1 As Double

    On Error GoTo EH

    If IsEmpty(arrIn) Or Not IsArray(arrIn) Then
        TransformerBG_Array = Empty
        GoTo CleanExit
    End If

    n = UBound(arrIn, 1)
    If n <= 0 Then
        TransformerBG_Array = Empty
        GoTo CleanExit
    End If

    ReDim outArr(1 To n, 1 To 4)

    For i = 1 To n
        accDigits = modUtils.KeepDigits(CStr(NzAny(arrIn(i, 1), "")))
        If Len(accDigits) = 0 Then GoTo NextI
        If Len(accDigits) > maxLen Then maxLen = Len(accDigits)

        lib = SanitizeLabel_FR(CStr(NzAny(arrIn(i, 2), "")))
        vN = Round2(ParseDoubleFR_Safe(NzAny(arrIn(i, 3), 0#)))
        vN1 = Round2(ParseDoubleFR_Safe(NzAny(arrIn(i, 4), 0#)))

        outArr(i, 1) = accDigits
        outArr(i, 2) = lib
        outArr(i, 3) = vN
        outArr(i, 4) = vN1
NextI:
    Next i

    If maxLen < 1 Then maxLen = 1

    For i = 1 To n
        If Len(CStr(NzAny(outArr(i, 1), ""))) > 0 Then
            outArr(i, 1) = PadRightZeros(CStr(outArr(i, 1)), maxLen)
            outArr(i, 1) = AccountToNumberIfDigits(CStr(outArr(i, 1)))
        End If
    Next i

    TransformerBG_Array = RemoveEmptyAccountRows(outArr)
    GoTo CleanExit
EH:
    Debug.Print "TransformerBG_Array error " & Err.Number & " : " & Err.Description
    TransformerBG_Array = Empty
    Resume CleanExit
CleanExit:
    Exit Function
End Function

Public Sub TransformerBG(ByRef arrBalance As Variant, Optional ByVal wsTarget As Worksheet)
    Dim arrOut As Variant
    Dim n As Long

    arrOut = TransformerBG_Array(arrBalance)
    If IsEmpty(arrOut) Or Not IsArray(arrOut) Then Exit Sub

    arrBalance = arrOut

    If Not wsTarget Is Nothing Then
        n = UBound(arrOut, 1)
        wsTarget.Range("A1:D1").value = Array("Compte", "Libelle", "Solde N", "Solde N-1")
        wsTarget.Columns("A").NumberFormat = "@"
        wsTarget.Range("A2").Resize(n, 4).value = arrOut
        wsTarget.Range("C2:D" & n + 1).NumberFormat = "0.00"
    End If
End Sub

Public Function NormalizeTo4Cols(ByVal arrBalance As Variant, ByVal mode As String, Optional ByVal fillN1WithZero As Boolean = True) As Variant
    Dim n As Long
    Dim m As Long
    Dim i As Long
    Dim outArr() As Variant
    Dim useMode As String

    On Error GoTo EH

    If IsEmpty(arrBalance) Or Not IsArray(arrBalance) Then
        NormalizeTo4Cols = Empty
        GoTo CleanExit
    End If

    n = UBound(arrBalance, 1)
    m = UBound(arrBalance, 2)

    If n <= 0 Then
        NormalizeTo4Cols = Empty
        GoTo CleanExit
    End If

    useMode = UCase$(Trim$(mode))
    ReDim outArr(1 To n, 1 To 4)

    For i = 1 To n
        outArr(i, 1) = NzAny(arrBalance(i, 1), vbNullString)
        If m >= 2 Then outArr(i, 2) = NzAny(arrBalance(i, 2), vbNullString)

        If m >= 4 And (useMode = "DIRECT" Or useMode = "AUTO") Then
            outArr(i, 3) = NzAny(arrBalance(i, 3), 0#)
            outArr(i, 4) = NzAny(arrBalance(i, 4), 0#)
        ElseIf useMode = "N-1" Or useMode = "N1" Then
            outArr(i, 3) = 0#
            If m >= 3 Then
                outArr(i, 4) = NzAny(arrBalance(i, 3), 0#)
            Else
                outArr(i, 4) = 0#
            End If
        Else
            If m >= 3 Then
                outArr(i, 3) = NzAny(arrBalance(i, 3), 0#)
            Else
                outArr(i, 3) = 0#
            End If

            If m >= 4 And Not fillN1WithZero Then
                outArr(i, 4) = NzAny(arrBalance(i, 4), 0#)
            Else
                outArr(i, 4) = 0#
            End If
        End If
    Next i

    NormalizeTo4Cols = TransformerBG_Array(outArr)
    GoTo CleanExit
EH:
    Debug.Print "NormalizeTo4Cols error " & Err.Number & " : " & Err.Description
    NormalizeTo4Cols = Empty
    Resume CleanExit
CleanExit:
    Exit Function
End Function

Public Function TransformSheetToBalanceArray(ByVal ws As Worksheet) As Variant
    Dim fmt As eBalanceFormat
    Dim bHdr As Boolean
    Dim accCol As Long, lblCol As Long, sldCol As Long, dbtCol As Long, crdCol As Long
    Dim lRow As Long, lCol As Long
    Dim src As Variant
    Dim outArr() As Variant
    Dim i As Long
    Dim o As Long
    Dim startRow As Long
    Dim accRaw As String
    Dim accDigits As String
    Dim lib As String
    Dim solde As Double

    On Error GoTo EH

    If ws Is Nothing Then GoTo CleanExit

    DetectColumns_ByRef ws, fmt, bHdr, accCol, lblCol, sldCol, dbtCol, crdCol, lRow, lCol

    If fmt = bfUnknown Then
        MsgBox "Format non reconnu: colonnes compte/libelle/solde introuvables.", vbExclamation
        GoTo CleanExit
    End If

    If lRow < 1 Or lCol < 1 Then
        GoTo CleanExit
    End If

    src = ws.Range(ws.Cells(1, 1), ws.Cells(lRow, lCol)).Value2

    startRow = IIf(bHdr, 2, 1)
    If startRow > UBound(src, 1) Then
        GoTo CleanExit
    End If

    ReDim outArr(1 To UBound(src, 1) - startRow + 1, 1 To 3)

    For i = startRow To UBound(src, 1)
        accRaw = CStr(NzAny(src(i, accCol), vbNullString))
        accDigits = modUtils.KeepDigits(accRaw)
        If Len(accDigits) = 0 Then GoTo NextI

        lib = vbNullString
        If lblCol > 0 Then lib = CStr(NzAny(src(i, lblCol), vbNullString))

        Select Case fmt
            Case bf4ColsDebitCredit
                solde = ParseDoubleFR_Safe(NzAny(src(i, dbtCol), 0#)) - ParseDoubleFR_Safe(NzAny(src(i, crdCol), 0#))
            Case Else
                solde = ParseDoubleFR_Safe(NzAny(src(i, sldCol), 0#))
        End Select

        o = o + 1
        outArr(o, 1) = accDigits
        outArr(o, 2) = SanitizeLabel_FR(lib)
        outArr(o, 3) = Round2(solde)
NextI:
    Next i

    If o = 0 Then
        GoTo CleanExit
    End If

    If o < UBound(outArr, 1) Then
        TransformSheetToBalanceArray = ShrinkRows3(outArr, o)
    Else
        TransformSheetToBalanceArray = outArr
    End If
    GoTo CleanExit
CleanExit:
    Exit Function
EH:
    TransformSheetToBalanceArray = Empty
    Resume CleanExit
End Function

Public Function TransformerBG_FromPath(ByVal pathBalance As String, Optional ByVal pathComparatif As String = "") As Variant
    Dim arrN As Variant
    Dim arrN1 As Variant
    Dim arrN4 As Variant
    Dim arrN14 As Variant
    Dim arrCompil As Variant

    arrN = LoadBalanceArray_FromExcelPath(pathBalance)
    If IsEmpty(arrN) Or Not IsArray(arrN) Then
        TransformerBG_FromPath = Empty
        Exit Function
    End If

    arrN4 = NormalizeTo4Cols(arrN, "N", True)

    If Len(Trim$(pathComparatif)) = 0 Then
        TransformerBG_FromPath = arrN4
        Exit Function
    End If

    arrN1 = LoadBalanceArray_FromExcelPath(pathComparatif)
    If IsEmpty(arrN1) Or Not IsArray(arrN1) Then
        TransformerBG_FromPath = Empty
        Exit Function
    End If

    arrN14 = NormalizeTo4Cols(arrN1, "N-1", True)
    arrCompil = BGcompil.BGCompil_FromBalanceArrays(arrN4, arrN14)

    TransformerBG_FromPath = TransformerBG_Array(arrCompil)
End Function

Public Sub ApplyFormattingBalanceSheet(ByVal ws As Worksheet, ByVal detectedFormat As eBalanceFormat)
    Dim lastRow As Long
    If ws Is Nothing Then Exit Sub

    lastRow = GetLastUsedRow(ws)

    ws.Cells(1, 1).value = "Compte"
    ws.Cells(1, 2).value = "Libelle"
    ws.Cells(1, 3).value = "Solde"

    ws.Columns(1).NumberFormat = "@"
    If lastRow >= 2 Then ws.Range(ws.Cells(2, 3), ws.Cells(lastRow, 3)).NumberFormat = "0.00"

    ws.rows(1).Font.Bold = True
    ws.Columns("A:C").AutoFit
End Sub

Public Function Balance_Normalize(ByVal arr As Variant) As Variant
    Dim dict As Object
    Dim i As Long
    Dim n As Long
    Dim acc As String
    Dim lib As String
    Dim amount As Double
    Dim rec As Variant
    Dim keys As Variant
    Dim outArr() As Variant
    Dim maxLen As Long
    Dim has401Aux As Boolean
    Dim has411Aux As Boolean
    Dim sum401 As Double
    Dim sum411 As Double
    Dim c401 As String
    Dim c411 As String
    Dim outRow As Long

    On Error GoTo EH

    If Not BGCompil_ArrayHasRows(arr) Then
        Balance_Normalize = Empty
        Exit Function
    End If

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbBinaryCompare

    n = UBound(arr, 1)
    For i = 1 To n
        acc = modUtils.KeepDigits(CStr(Nz2(arr(i, 1), "")))
        If Len(acc) = 0 Then GoTo NextI

        If Len(acc) > maxLen Then maxLen = Len(acc)

        lib = SanitizeLabel2(CStr(Nz2(arr(i, 2), "")))
        amount = ParseDoubleFR2(Nz2(arr(i, 3), 0#))

        If Left$(acc, 3) = "401" And Len(acc) > 3 Then
            has401Aux = True
            sum401 = sum401 + amount
            GoTo NextI
        End If

        If Left$(acc, 3) = "411" And Len(acc) > 3 Then
            has411Aux = True
            sum411 = sum411 + amount
            GoTo NextI
        End If

        If dict.Exists(acc) Then
            rec = dict(acc)
            If Len(CStr(rec(0))) = 0 And Len(lib) > 0 Then rec(0) = lib
            rec(1) = CDbl(rec(1)) + amount
            dict(acc) = rec
        Else
            dict.Add acc, Array(lib, amount)
        End If
NextI:
    Next i

    If maxLen < 3 Then maxLen = 3

    If has401Aux Then
        c401 = "401" & String$(maxLen - 3, "0")
        If dict.Exists(c401) Then
            rec = dict(c401)
            rec(1) = CDbl(rec(1)) + sum401
            If Len(CStr(rec(0))) = 0 Then rec(0) = "Centralisation 401*"
            dict(c401) = rec
        Else
            dict.Add c401, Array("Centralisation 401*", sum401)
        End If
    End If

    If has411Aux Then
        c411 = "411" & String$(maxLen - 3, "0")
        If dict.Exists(c411) Then
            rec = dict(c411)
            rec(1) = CDbl(rec(1)) + sum411
            If Len(CStr(rec(0))) = 0 Then rec(0) = "Centralisation 411*"
            dict(c411) = rec
        Else
            dict.Add c411, Array("Centralisation 411*", sum411)
        End If
    End If

    If dict.Count = 0 Then
        Balance_Normalize = Empty
        Exit Function
    End If

    keys = dict.keys
    SortStringArray keys

    ReDim outArr(1 To dict.Count, 1 To 3)
    outRow = 0
    For i = LBound(keys) To UBound(keys)
        outRow = outRow + 1
        rec = dict(keys(i))
        outArr(outRow, 1) = CStr(keys(i))
        outArr(outRow, 2) = CStr(rec(0))
        outArr(outRow, 3) = CDbl(rec(1))
    Next i

    Balance_Normalize = outArr
    Exit Function

EH:
    Balance_Normalize = Empty
End Function


' ============================================================
' 4. MAPPING COLONNES BALANCE
' ============================================================

Public Function Balance_DetectHeaderRow(ByVal arrSrc As Variant) As Boolean
    Dim m As Long
    Dim j As Long
    Dim h As String

    On Error GoTo EH

    If IsEmpty(arrSrc) Or Not IsArray(arrSrc) Then Exit Function

    m = UBound(arrSrc, 2)
    For j = 1 To m
        h = NormalizeHeaderMap(CStr(ReadCellSafe(arrSrc, 1, j, vbNullString)))
        Select Case h
            Case "compte", "comptenum", "comptenumero", "comptelib", "libelle", "solde", "solden", "solden1", "debit", "credit", "totaldebit", "totalcredit"
                Balance_DetectHeaderRow = True
                Exit Function
        End Select
    Next j
    Exit Function
EH:
    Balance_DetectHeaderRow = False
End Function

Public Function Balance_GuessColumns( _
    ByVal arrSrc As Variant, _
    ByRef idxCompte As Long, _
    ByRef idxLib As Long, _
    ByRef idxSoldeN As Long, _
    ByRef idxSoldeN1 As Long) As Boolean

    Dim m As Long
    Dim j As Long
    Dim h As String
    Dim hasHeader As Boolean

    On Error GoTo EH

    If IsEmpty(arrSrc) Or Not IsArray(arrSrc) Then Exit Function
    m = UBound(arrSrc, 2)
    If m <= 0 Then Exit Function

    idxCompte = 0
    idxLib = 0
    idxSoldeN = 0
    idxSoldeN1 = 0

    hasHeader = Balance_DetectHeaderRow(arrSrc)

    If hasHeader Then
        For j = 1 To m
            h = NormalizeHeaderMap(CStr(ReadCellSafe(arrSrc, 1, j, vbNullString)))
            Select Case h
                Case "compte", "comptenum", "comptenumero"
                    If idxCompte = 0 Then idxCompte = j
                Case "comptelib", "libelle"
                    If idxLib = 0 Then idxLib = j
                Case "solden1", "solden-1", "soldenmoins1", "n1", "soldeanterior"
                    If idxSoldeN1 = 0 Then idxSoldeN1 = j
                Case "solde", "solden"
                    If idxSoldeN = 0 Then idxSoldeN = j
                Case "debit", "credit", "totaldebit", "totalcredit"
                    If idxSoldeN = 0 Then idxSoldeN = j
            End Select
        Next j
    End If

    If idxCompte = 0 Then idxCompte = 1
    If idxLib = 0 And m >= 2 Then idxLib = 2

    If idxSoldeN = 0 Then
        If m = 3 Then
            idxSoldeN = 3
        ElseIf m >= 4 Then
            idxSoldeN = 3
        End If
    End If

    If idxSoldeN1 = 0 Then
        If m >= 4 Then
            idxSoldeN1 = 4
        Else
            idxSoldeN1 = 0
        End If
    End If

    Balance_GuessColumns = (idxCompte > 0 And idxSoldeN > 0)
    Exit Function
EH:
    Balance_GuessColumns = False
End Function

Public Function Balance_MapTo4Cols( _
    ByVal arrSrc As Variant, _
    ByVal idxCompte As Long, _
    ByVal idxLib As Long, _
    ByVal idxSoldeN As Long, _
    ByVal idxSoldeN1 As Long) As Variant

    Dim hasHeader As Boolean
    Dim startRow As Long
    Dim srcRows As Long
    Dim outRows As Long
    Dim m As Long
    Dim i As Long
    Dim r As Long
    Dim outArr() As Variant

    On Error GoTo EH

    If IsEmpty(arrSrc) Or Not IsArray(arrSrc) Then
        Balance_MapTo4Cols = Empty
        Exit Function
    End If

    srcRows = UBound(arrSrc, 1)
    m = UBound(arrSrc, 2)
    If srcRows <= 0 Or m <= 0 Then
        Balance_MapTo4Cols = Empty
        Exit Function
    End If

    hasHeader = Balance_DetectHeaderRow(arrSrc)
    startRow = IIf(hasHeader, 2, 1)

    If startRow > srcRows Then
        Balance_MapTo4Cols = Empty
        Exit Function
    End If

    outRows = srcRows - startRow + 1
    ReDim outArr(1 To outRows, 1 To 4)

    r = 0
    For i = startRow To srcRows
        r = r + 1

        outArr(r, 1) = ReadCellSafe(arrSrc, i, idxCompte, vbNullString)

        If idxLib > 0 Then
            outArr(r, 2) = ReadCellSafe(arrSrc, i, idxLib, vbNullString)
        Else
            outArr(r, 2) = vbNullString
        End If

        If idxSoldeN > 0 Then
            outArr(r, 3) = ReadCellSafe(arrSrc, i, idxSoldeN, 0#)
        Else
            outArr(r, 3) = 0#
        End If

        If idxSoldeN1 > 0 Then
            outArr(r, 4) = ReadCellSafe(arrSrc, i, idxSoldeN1, 0#)
        Else
            outArr(r, 4) = 0#
        End If
    Next i

    Balance_MapTo4Cols = outArr
    Exit Function
EH:
    Balance_MapTo4Cols = Empty
End Function


' ============================================================
' 5. COMPILATION BALANCE N / N-1
' ============================================================

Public Function BGCompil_FromBalanceArrays( _
    ByVal balN As Variant, _
    ByVal balN1 As Variant, _
    Optional ByRef outMaxLen As Long, _
    Optional ByRef outInfo As String) As Variant

    Dim arrN3 As Variant
    Dim arrN13 As Variant
    Dim arrOut As Variant

    arrN3 = ConvertTo3ColsForCompile(balN, 3)
    arrN13 = ConvertTo3ColsForCompile(balN1, 4)
    arrOut = BGCompil_Compile(arrN3, arrN13)

    If BGCompil_ArrayHasRows(arrOut) Then
        outMaxLen = modOrchestrator.ComputeMaxAccountLen(arrOut)
        outInfo = "Compilation OK (" & CStr(UBound(arrOut, 1)) & " lignes)"
    Else
        outMaxLen = 0
        outInfo = "Compilation vide"
    End If

    BGCompil_FromBalanceArrays = arrOut
End Function

Public Function BGCompil_CompileFromBalances(ByVal arrN As Variant, ByVal arrN1 As Variant) As Variant
    BGCompil_CompileFromBalances = BGCompil_FromBalanceArrays(arrN, arrN1)
End Function

Public Function BGCompil_Compile(ByVal arrN As Variant, ByVal arrN1 As Variant) As Variant
    Dim nN As Variant
    Dim nN1 As Variant
    Dim dict As Object
    Dim i As Long
    Dim key As String
    Dim rec As Variant
    Dim keys As Variant
    Dim outArr() As Variant
    Dim outRow As Long
    Dim maxLen As Long
    Dim libN As String
    Dim libN1 As String

    On Error GoTo EH

    nN = Balance_Normalize(arrN)
    nN1 = Balance_Normalize(arrN1)

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbBinaryCompare

    If BGCompil_ArrayHasRows(nN) Then
        For i = 1 To UBound(nN, 1)
            key = CStr(nN(i, 1))
            If Len(key) = 0 Then GoTo NextN
            rec = Array(vbNullString, 0#, vbNullString, 0#)
            rec(0) = CStr(Nz2(nN(i, 2), ""))
            rec(1) = ParseDoubleFR2(Nz2(nN(i, 3), 0#))
            dict(key) = rec
            If Len(key) > maxLen Then maxLen = Len(key)
NextN:
        Next i
    End If

    If BGCompil_ArrayHasRows(nN1) Then
        For i = 1 To UBound(nN1, 1)
            key = CStr(nN1(i, 1))
            If Len(key) = 0 Then GoTo NextN1

            If dict.Exists(key) Then
                rec = dict(key)
            Else
                rec = Array(vbNullString, 0#, vbNullString, 0#)
            End If

            rec(2) = CStr(Nz2(nN1(i, 2), ""))
            rec(3) = ParseDoubleFR2(Nz2(nN1(i, 3), 0#))
            dict(key) = rec
            If Len(key) > maxLen Then maxLen = Len(key)
NextN1:
        Next i
    End If

    If dict.Count = 0 Then
        BGCompil_Compile = Empty
        Exit Function
    End If

    If maxLen < 1 Then maxLen = 1

    keys = dict.keys
    SortStringArray keys

    ReDim outArr(1 To dict.Count, 1 To 4)
    outRow = 0

    For i = LBound(keys) To UBound(keys)
        outRow = outRow + 1
        rec = dict(keys(i))

        libN = CStr(Nz2(rec(0), ""))
        libN1 = CStr(Nz2(rec(2), ""))

        outArr(outRow, 1) = PadRightZeros2(CStr(keys(i)), maxLen)
        If Len(libN) > 0 Then
            outArr(outRow, 2) = libN
        Else
            outArr(outRow, 2) = libN1
        End If
        outArr(outRow, 3) = CDbl(Nz2(rec(1), 0#))
        outArr(outRow, 4) = CDbl(Nz2(rec(3), 0#))
    Next i

    BGCompil_Compile = outArr
    Exit Function

EH:
    BGCompil_Compile = Empty
End Function

Public Function BGCompil_ArrayHasRows(ByVal arr As Variant) As Boolean
    On Error GoTo EH
    BGCompil_ArrayHasRows = IsArray(arr) And UBound(arr, 1) >= 1 And UBound(arr, 2) >= 3
    Exit Function
EH:
    BGCompil_ArrayHasRows = False
End Function


' ============================================================
' 6. UTILITAIRES PRIVES (version unique consolidee)
' ============================================================

Private Sub DetectColumns_ByRef(ByVal ws As Worksheet, _
    ByRef outFormat As eBalanceFormat, _
    ByRef outHasHeader As Boolean, _
    ByRef outAccountCol As Long, _
    ByRef outLabelCol As Long, _
    ByRef outSoldeCol As Long, _
    ByRef outDebitCol As Long, _
    ByRef outCreditCol As Long, _
    ByRef outLastRow As Long, _
    ByRef outLastCol As Long)

    Dim c As Long, startDataRow As Long, h As String
    Dim numScore() As Double, txtScore() As Double

    outFormat     = bfUnknown
    outHasHeader  = False
    outAccountCol = 0: outLabelCol  = 0
    outSoldeCol   = 0: outDebitCol  = 0: outCreditCol = 0
    outLastRow    = 0: outLastCol   = 0

    On Error GoTo EH

    outLastRow = GetLastUsedRow(ws)
    outLastCol = GetLastUsedCol(ws)
    If outLastRow = 0 Or outLastCol = 0 Then GoTo CleanExit

    ReDim numScore(1 To outLastCol)
    ReDim txtScore(1 To outLastCol)

    ' Detection entete
    outHasHeader = False
    For c = 1 To outLastCol
        h = NormalizeHeader_FR(CStr(NzAny(ws.Cells(1, c).Value2, "")))
        Select Case h
            Case "compte", "comptenum", "comptenumero", "comptelib", "libelle", _
                 "solde", "solden", "solden1", "debit", "credit", _
                 "totaldebit", "totalcredit", "soldedebit", "soldecredit"
                outHasHeader = True
                Exit For
        End Select
    Next c

    startDataRow = IIf(outHasHeader, 2, 1)

    If outHasHeader Then
        For c = 1 To outLastCol
            h = NormalizeHeader_FR(CStr(NzAny(ws.Cells(1, c).Value2, "")))
            Select Case h
                Case "compte", "comptenum", "comptenumero"
                    If outAccountCol = 0 Then outAccountCol = c
                Case "comptelib", "libelle"
                    If outLabelCol = 0 Then outLabelCol = c
                Case "solde", "solden", "solden1"
                    If outSoldeCol = 0 Then outSoldeCol = c
                Case "debit", "totaldebit", "soldedebit"
                    If outDebitCol = 0 Then outDebitCol = c
                Case "credit", "totalcredit", "soldecredit"
                    If outCreditCol = 0 Then outCreditCol = c
            End Select
        Next c
    End If

    Dim lim As Long
    lim = WorksheetFunction.Min(outLastRow, startDataRow + 200)
    For c = 1 To outLastCol
        numScore(c) = ColumnNumericScore(ws, c, startDataRow, lim)
        txtScore(c) = ColumnTextScore(ws, c, startDataRow, lim)
    Next c

    If outAccountCol = 0 Then outAccountCol = GuessAccountCol(ws, startDataRow, outLastRow, outLastCol)
    If outLabelCol  = 0 Then outLabelCol  = GuessLabelCol(outLastCol, outAccountCol, txtScore)

    If outSoldeCol = 0 And outDebitCol = 0 And outCreditCol = 0 Then
        If outLastCol = 3 Then
            outSoldeCol = 3
        Else
            GuessAmountCols outLastCol, outAccountCol, outLabelCol, numScore, outSoldeCol, outDebitCol, outCreditCol
        End If
    End If

    If outDebitCol > 0 And outCreditCol > 0 Then
        outFormat = bf4ColsDebitCredit
    ElseIf outSoldeCol > 0 Then
        If outLastCol = 3 Then
            outFormat = bf3ColsSolde
        Else
            outFormat = bfHeuristicMapped
        End If
    Else
        outFormat = bfUnknown
    End If

CleanExit:
    Exit Sub
EH:
    outFormat = bfUnknown
    Resume CleanExit
End Sub

Private Sub GuessAmountCols(ByVal lastCol As Long, ByVal accountCol As Long, ByVal labelCol As Long, ByRef numScore() As Double, ByRef soldeCol As Long, ByRef debitCol As Long, ByRef creditCol As Long)
    Dim c As Long
    Dim best1 As Long, best2 As Long
    Dim s1 As Double, s2 As Double

    For c = 1 To lastCol
        If c <> accountCol And c <> labelCol Then
            If numScore(c) > s1 Then
                s2 = s1: best2 = best1
                s1 = numScore(c): best1 = c
            ElseIf numScore(c) > s2 Then
                s2 = numScore(c): best2 = c
            End If
        End If
    Next c

    If best1 > 0 And best2 > 0 And s1 >= 0.5 And s2 >= 0.5 Then
        debitCol = best1
        creditCol = best2
    ElseIf best1 > 0 Then
        soldeCol = best1
    End If
End Sub

Private Function GuessAccountCol(ByVal ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long, ByVal lastCol As Long) As Long
    Dim c As Long
    Dim score As Double
    Dim best As Double

    For c = 1 To lastCol
        score = ColumnAccountLikeScore(ws, c, r1, WorksheetFunction.Min(r2, r1 + 200))
        If score > best Then
            best = score
            GuessAccountCol = c
        End If
    Next c

    If GuessAccountCol = 0 Then GuessAccountCol = 1
End Function

Private Function GuessLabelCol(ByVal lastCol As Long, ByVal accountCol As Long, ByRef txtScore() As Double) As Long
    Dim c As Long
    Dim best As Double

    For c = 1 To lastCol
        If c <> accountCol Then
            If txtScore(c) > best Then
                best = txtScore(c)
                GuessLabelCol = c
            End If
        End If
    Next c

    If GuessLabelCol = 0 Then
        If lastCol >= 2 Then GuessLabelCol = 2 Else GuessLabelCol = 1
    End If
End Function

Private Function ColumnAccountLikeScore(ByVal ws As Worksheet, ByVal col As Long, ByVal r1 As Long, ByVal r2 As Long) As Double
    Dim r As Long
    Dim n As Long
    Dim ok As Long
    Dim s As String

    For r = r1 To r2
        s = modUtils.KeepDigits(CStr(NzAny(ws.Cells(r, col).Value2, "")))
        If Len(s) > 0 Then
            n = n + 1
            If Len(s) >= 3 Then ok = ok + 1
        End If
    Next r

    If n > 0 Then ColumnAccountLikeScore = ok / n
End Function

Private Function ColumnNumericScore(ByVal ws As Worksheet, ByVal col As Long, ByVal r1 As Long, ByVal r2 As Long) As Double
    Dim r As Long
    Dim n As Long
    Dim ok As Long
    Dim v As Variant

    For r = r1 To r2
        v = ws.Cells(r, col).Value2
        If Len(Trim$(CStr(NzAny(v, "")))) > 0 Then
            n = n + 1
            If IsLikelyNumeric(v) Then ok = ok + 1
        End If
    Next r

    If n > 0 Then ColumnNumericScore = ok / n
End Function

Private Function ColumnTextScore(ByVal ws As Worksheet, ByVal col As Long, ByVal r1 As Long, ByVal r2 As Long) As Double
    Dim r As Long
    Dim n As Long
    Dim ok As Long
    Dim v As Variant
    Dim s As String

    For r = r1 To r2
        v = ws.Cells(r, col).Value2
        s = Trim$(CStr(NzAny(v, "")))
        If Len(s) > 0 Then
            n = n + 1
            If Not IsLikelyNumeric(v) Then ok = ok + 1
        End If
    Next r

    If n > 0 Then ColumnTextScore = ok / n
End Function

Private Function IsLikelyNumeric(ByVal v As Variant) As Boolean
    If IsNumeric(v) Then
        IsLikelyNumeric = True
    Else
        IsLikelyNumeric = (ParseDoubleFR_Safe(v) <> 0# Or InStr(1, CStr(v), "0", vbBinaryCompare) > 0)
    End If
End Function

Private Function RemoveEmptyAccountRows(ByVal arr As Variant) As Variant
    Dim n As Long
    Dim i As Long
    Dim o As Long
    Dim outArr() As Variant

    On Error GoTo EH

    n = UBound(arr, 1)
    ReDim outArr(1 To n, 1 To 4)

    For i = 1 To n
        If Len(CStr(NzAny(arr(i, 1), ""))) > 0 Then
            o = o + 1
            outArr(o, 1) = arr(i, 1)
            outArr(o, 2) = arr(i, 2)
            outArr(o, 3) = arr(i, 3)
            outArr(o, 4) = arr(i, 4)
        End If
    Next i

    If o = 0 Then
        RemoveEmptyAccountRows = Empty
        GoTo CleanExit
    End If

    RemoveEmptyAccountRows = ShrinkRows4(outArr, o)
    GoTo CleanExit
EH:
    Debug.Print "RemoveEmptyAccountRows error " & Err.Number & " : " & Err.Description
    RemoveEmptyAccountRows = Empty
    Resume CleanExit
CleanExit:
    Exit Function
End Function

Private Function Normalize3Cols(ByVal arr As Variant) As Variant
    Dim n As Long
    Dim i As Long
    Dim outArr() As Variant

    On Error GoTo EH

    n = UBound(arr, 1)
    ReDim outArr(1 To n, 1 To 3)

    For i = 1 To n
        outArr(i, 1) = modUtils.KeepDigits(CStr(NzAny(arr(i, 1), "")))
        outArr(i, 2) = SanitizeLabel_FR(CStr(NzAny(arr(i, 2), "")))
        outArr(i, 3) = Round2(ParseDoubleFR_Safe(NzAny(arr(i, 3), 0#)))
    Next i

    Normalize3Cols = RemoveEmptyAccountRows3(outArr)
    GoTo CleanExit
EH:
    Debug.Print "Normalize3Cols error " & Err.Number & " : " & Err.Description
    Normalize3Cols = Empty
    Resume CleanExit
CleanExit:
    Exit Function
End Function

Private Function RemoveEmptyAccountRows3(ByVal arr As Variant) As Variant
    Dim n As Long
    Dim i As Long
    Dim o As Long
    Dim outArr() As Variant

    On Error GoTo EH

    n = UBound(arr, 1)
    ReDim outArr(1 To n, 1 To 3)

    For i = 1 To n
        If Len(CStr(NzAny(arr(i, 1), ""))) > 0 Then
            o = o + 1
            outArr(o, 1) = arr(i, 1)
            outArr(o, 2) = arr(i, 2)
            outArr(o, 3) = arr(i, 3)
        End If
    Next i

    If o = 0 Then
        RemoveEmptyAccountRows3 = Empty
    Else
        RemoveEmptyAccountRows3 = ShrinkRows3(outArr, o)
    End If
    GoTo CleanExit
EH:
    Debug.Print "RemoveEmptyAccountRows3 error " & Err.Number & " : " & Err.Description
    RemoveEmptyAccountRows3 = Empty
    Resume CleanExit
CleanExit:
    Exit Function
End Function

Private Function ShrinkRows3(ByVal arr As Variant, ByVal nOut As Long) As Variant
    Dim r As Long
    Dim c As Long
    Dim outArr() As Variant

    ReDim outArr(1 To nOut, 1 To 3)
    For r = 1 To nOut
        For c = 1 To 3
            outArr(r, c) = arr(r, c)
        Next c
    Next r
    ShrinkRows3 = outArr
End Function

Private Function ShrinkRows4(ByVal arr As Variant, ByVal nOut As Long) As Variant
    Dim r As Long
    Dim c As Long
    Dim outArr() As Variant

    ReDim outArr(1 To nOut, 1 To 4)
    For r = 1 To nOut
        For c = 1 To 4
            outArr(r, c) = arr(r, c)
        Next c
    Next r
    ShrinkRows4 = outArr
End Function

Private Function AccountToNumberIfDigits(ByVal s As String) As Variant
    Dim i As Long
    Dim ch As String

    If Len(s) = 0 Then
        AccountToNumberIfDigits = vbNullString
        Exit Function
    End If

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch < "0" Or ch > "9" Then
            AccountToNumberIfDigits = s
            Exit Function
        End If
    Next i

    If Len(s) <= 15 Then
        AccountToNumberIfDigits = CDbl(s)
    Else
        AccountToNumberIfDigits = s
    End If
End Function

Private Sub DetectBalanceColumns(ByVal ws As Worksheet, ByVal lastCol As Long, ByRef dataStartRow As Long, ByRef colAccount As Long, ByRef colLabel As Long, ByRef colSolde As Long, ByRef colDebit As Long, ByRef colCredit As Long)
    Dim hasHeader As Boolean
    Dim c As Long
    Dim h As String

    hasHeader = False
    For c = 1 To lastCol
        h = NormalizeHeader2(CStr(Nz2(ws.Cells(1, c).Value2, "")))
        Select Case h
            Case "compte", "comptenum", "comptenumero", "comptelib", "libelle", "solde", "debit", "credit", "totaldebit", "totalcredit"
                hasHeader = True
                Exit For
        End Select
    Next c

    If hasHeader Then
        dataStartRow = 2
    Else
        dataStartRow = 1
    End If

    For c = 1 To lastCol
        h = NormalizeHeader2(CStr(Nz2(ws.Cells(1, c).Value2, "")))
        Select Case h
            Case "compte", "comptenum", "comptenumero"
                If colAccount = 0 Then colAccount = c
            Case "comptelib", "libelle"
                If colLabel = 0 Then colLabel = c
            Case "solde"
                If colSolde = 0 Then colSolde = c
            Case "debit", "totaldebit"
                If colDebit = 0 Then colDebit = c
            Case "credit", "totalcredit"
                If colCredit = 0 Then colCredit = c
        End Select
    Next c

    If colAccount = 0 Then colAccount = 1
    If colLabel = 0 Then colLabel = 2

    If colSolde = 0 Then
        If colDebit = 0 Then colDebit = 3
        If colCredit = 0 Then colCredit = 4
    End If
End Sub

Private Function ConvertTo3ColsForCompile(ByVal arrIn As Variant, ByVal preferredAmountCol As Long) As Variant
    Dim n As Long
    Dim m As Long
    Dim i As Long
    Dim amountCol As Long
    Dim outArr() As Variant

    On Error GoTo EH

    If Not IsArray(arrIn) Then
        ConvertTo3ColsForCompile = Empty
        Exit Function
    End If

    n = UBound(arrIn, 1)
    m = UBound(arrIn, 2)
    If n <= 0 Or m < 3 Then
        ConvertTo3ColsForCompile = Empty
        Exit Function
    End If

    amountCol = 3
    If m >= 4 Then
        If preferredAmountCol = 4 Then
            amountCol = 4
        Else
            amountCol = 3
        End If
    End If

    ReDim outArr(1 To n, 1 To 3)
    For i = 1 To n
        outArr(i, 1) = arrIn(i, 1)
        outArr(i, 2) = arrIn(i, 2)
        outArr(i, 3) = arrIn(i, amountCol)
    Next i

    ConvertTo3ColsForCompile = outArr
    Exit Function
EH:
    ConvertTo3ColsForCompile = Empty
End Function

Private Function Nz2(ByVal v As Variant, Optional ByVal fallback As Variant = "") As Variant
    If IsError(v) Then
        Nz2 = fallback
    ElseIf IsNull(v) Then
        Nz2 = fallback
    ElseIf IsEmpty(v) Then
        Nz2 = fallback
    Else
        Nz2 = v
    End If
End Function

Private Function ReadCellSafe(ByVal arr As Variant, ByVal r As Long, ByVal c As Long, ByVal fallback As Variant) As Variant
    On Error GoTo EH
    If c <= 0 Then
        ReadCellSafe = fallback
    Else
        ReadCellSafe = arr(r, c)
        If IsError(ReadCellSafe) Or IsNull(ReadCellSafe) Or IsEmpty(ReadCellSafe) Then
            ReadCellSafe = fallback
        End If
    End If
    Exit Function
EH:
    ReadCellSafe = fallback
End Function

Private Function NormalizeHeaderMap(ByVal s As String) As String
    Dim t As String

    t = LCase$(Trim$(CStr(s)))
    t = Replace(t, Chr$(160), vbNullString)
    t = Replace(t, " ", vbNullString)
    t = Replace(t, "_", vbNullString)
    t = Replace(t, "-", vbNullString)

    t = Replace(t, ChrW$(233), "e")
    t = Replace(t, ChrW$(232), "e")
    t = Replace(t, ChrW$(234), "e")
    t = Replace(t, ChrW$(235), "e")
    t = Replace(t, ChrW$(224), "a")
    t = Replace(t, ChrW$(226), "a")
    t = Replace(t, ChrW$(228), "a")
    t = Replace(t, ChrW$(238), "i")
    t = Replace(t, ChrW$(239), "i")
    t = Replace(t, ChrW$(244), "o")
    t = Replace(t, ChrW$(246), "o")
    t = Replace(t, ChrW$(249), "u")
    t = Replace(t, ChrW$(251), "u")
    t = Replace(t, ChrW$(252), "u")
    t = Replace(t, ChrW$(231), "c")

    NormalizeHeaderMap = t
End Function

Private Function BuildModifPath(ByVal srcPath As String) As String
    Dim dotPos As Long
    Dim basePath As String
    Dim ext As String

    dotPos = InStrRev(srcPath, ".")
    If dotPos > 0 Then
        basePath = Left$(srcPath, dotPos - 1)
        ext = Mid$(srcPath, dotPos + 1)
        BuildModifPath = basePath & "_MODIF." & ext
    Else
        BuildModifPath = srcPath & "_MODIF.xlsx"
    End If
End Function

Private Function GetFileFormatFromPath(ByVal filePath As String) As XlFileFormat
    Dim ext As String
    ext = LCase$(Mid$(filePath, InStrRev(filePath, ".") + 1))

    Select Case ext
        Case "xlsm"
            GetFileFormatFromPath = xlOpenXMLWorkbookMacroEnabled
        Case "xls"
            GetFileFormatFromPath = xlExcel8
        Case Else
            GetFileFormatFromPath = xlOpenXMLWorkbook
    End Select
End Function

Private Sub FindFecColumns(ByRef headers() As String, ByRef idxCompte As Long, ByRef idxLib As Long, ByRef idxDebit As Long, ByRef idxCredit As Long)
    Dim i As Long
    Dim h As String

    idxCompte = -1
    idxLib = -1
    idxDebit = -1
    idxCredit = -1

    For i = LBound(headers) To UBound(headers)
        h = NormalizeHeader3(headers(i))
        Select Case h
            Case "comptenum", "compte", "comptenumero"
                If idxCompte < 0 Then idxCompte = i
            Case "comptelib", "libelle"
                If idxLib < 0 Then idxLib = i
            Case "debit"
                If idxDebit < 0 Then idxDebit = i
            Case "credit"
                If idxCredit < 0 Then idxCredit = i
        End Select
    Next i
End Sub

Private Function DetectDelimiter2(ByVal headerLine As String) As String
    Dim candidates As Variant
    Dim i As Long
    Dim d As String
    Dim colCount As Long
    Dim bestCols As Long
    Dim parts As Variant

    candidates = Array(vbTab, ";", "|", ",")

    For i = LBound(candidates) To UBound(candidates)
        d = CStr(candidates(i))
        parts = Split(headerLine, d)
        colCount = UBound(parts) - LBound(parts) + 1
        If colCount > bestCols Then
            bestCols = colCount
            DetectDelimiter2 = d
        End If
    Next i

    If bestCols <= 1 Then DetectDelimiter2 = vbNullString
End Function

Private Function StripUtf8Bom(ByVal s As String) As String
    If Len(s) > 0 Then
        If AscW(Left$(s, 1)) = 65279 Then
            StripUtf8Bom = Mid$(s, 2)
            Exit Function
        End If
    End If
    StripUtf8Bom = s
End Function

Private Function GetField(ByRef parts() As String, ByVal idx As Long) As String
    If idx < LBound(parts) Or idx > UBound(parts) Then
        GetField = vbNullString
    Else
        GetField = CStr(parts(idx))
    End If
End Function

Private Sub Fec_Accumulate(ByVal dict As Object, ByVal acc As String, ByVal lib As String, ByVal solde As Double)
    Dim rec As Variant
    If dict.Exists(acc) Then
        rec = dict(acc)
        If Len(CStr(rec(0))) = 0 And Len(lib) > 0 Then rec(0) = lib
        rec(1) = CDbl(rec(1)) + solde
        dict(acc) = rec
    Else
        dict.Add acc, Array(lib, solde)
    End If
End Sub

Private Function Fec_FirstNonEmptySheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If Fec_LastUsedRow(ws) > 0 And Fec_LastUsedCol(ws) > 0 Then
            Set Fec_FirstNonEmptySheet = ws
            Exit Function
        End If
    Next ws
End Function

Private Function Fec_LastUsedRow(ByVal ws As Worksheet) As Long
    Dim r As Range
    On Error Resume Next
    Set r = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    On Error GoTo 0
    If r Is Nothing Then
        Fec_LastUsedRow = 0
    Else
        Fec_LastUsedRow = r.Row
    End If
End Function

Private Function Fec_LastUsedCol(ByVal ws As Worksheet) As Long
    Dim r As Range
    On Error Resume Next
    Set r = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    On Error GoTo 0
    If r Is Nothing Then
        Fec_LastUsedCol = 0
    Else
        Fec_LastUsedCol = r.Column
    End If
End Function


' ============================================================
' 7. HELPERS STRING ET PARSE (version canonique + aliases)
' ============================================================

Private Function GetFirstNonEmptySheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If GetLastUsedRow(ws) > 0 And GetLastUsedCol(ws) > 0 Then
            Set GetFirstNonEmptySheet = ws
            Exit Function
        End If
    Next ws
End Function

Private Function GetLastUsedRow(ByVal ws As Worksheet) As Long
    Dim r As Range
    On Error Resume Next
    Set r = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    On Error GoTo 0
    If r Is Nothing Then
        GetLastUsedRow = 0
    Else
        GetLastUsedRow = r.Row
    End If
End Function

Private Function GetLastUsedCol(ByVal ws As Worksheet) As Long
    Dim r As Range
    On Error Resume Next
    Set r = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    On Error GoTo 0
    If r Is Nothing Then
        GetLastUsedCol = 0
    Else
        GetLastUsedCol = r.Column
    End If
End Function

Private Function PadRightZeros(ByVal s As String, ByVal targetLen As Long) As String
    Dim t As String
    t = CStr(s)
    If targetLen < 1 Then targetLen = 1
    If Len(t) < targetLen Then t = t & String$(targetLen - Len(t), "0")
    PadRightZeros = t
End Function

Private Function PadRightZeros2(ByVal s As String, ByVal targetLen As Long) As String
    PadRightZeros2 = PadRightZeros(s, targetLen)
End Function

Private Function ParseDoubleFR_Safe(ByVal v As Variant) As Double
    Dim s As String
    Dim neg As Boolean
    Dim posComma As Long
    Dim posDot As Long
    Dim posDec As Long
    Dim decSep As String
    Dim intPart As String
    Dim fracPart As String
    Dim norm As String
    Dim digitsAfter As Long

    On Error GoTo EH

    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function
    If IsNumeric(v) Then
        ParseDoubleFR_Safe = CDbl(v)
        GoTo CleanExit
    End If

    s = Trim$(CStr(v))
    If Len(s) = 0 Then Exit Function

    s = Replace(s, Chr$(160), " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, " ", vbNullString)
    s = Replace(s, "'", vbNullString)

    If Left$(s, 1) = "(" And Right$(s, 1) = ")" Then
        neg = True
        s = Mid$(s, 2, Len(s) - 2)
    End If

    posComma = InStrRev(s, ",")
    posDot = InStrRev(s, ".")

    If posComma > 0 And posDot > 0 Then
        If posComma > posDot Then decSep = "," Else decSep = "."
    ElseIf posComma > 0 Then
        digitsAfter = Len(s) - posComma
        If digitsAfter = 3 And InStr(1, s, ",", vbBinaryCompare) = posComma Then decSep = vbNullString Else decSep = ","
    ElseIf posDot > 0 Then
        digitsAfter = Len(s) - posDot
        If digitsAfter = 3 And InStr(1, s, ".", vbBinaryCompare) = posDot Then decSep = vbNullString Else decSep = "."
    Else
        decSep = vbNullString
    End If

    If Len(decSep) > 0 Then
        posDec = InStrRev(s, decSep)
        intPart = Left$(s, posDec - 1)
        fracPart = Mid$(s, posDec + 1)
        intPart = Replace(intPart, ",", vbNullString)
        intPart = Replace(intPart, ".", vbNullString)
        fracPart = Replace(fracPart, ",", vbNullString)
        fracPart = Replace(fracPart, ".", vbNullString)
        norm = intPart & "." & fracPart
    Else
        norm = Replace(s, ",", vbNullString)
        norm = Replace(norm, ".", vbNullString)
    End If

    If neg And Left$(norm, 1) <> "-" Then norm = "-" & norm

    If Len(norm) = 0 Or norm = "-" Then
        ParseDoubleFR_Safe = 0#
    ElseIf IsNumeric(norm) Then
        ParseDoubleFR_Safe = CDbl(norm)
    End If
    GoTo CleanExit
EH:
    Debug.Print "ParseDoubleFR_Safe error " & Err.Number & " : " & Err.Description
    ParseDoubleFR_Safe = 0#
    Resume CleanExit
CleanExit:
    Exit Function
End Function

Private Function ParseDoubleFR2(ByVal v As Variant) As Double
    ParseDoubleFR2 = ParseDoubleFR_Safe(v)
End Function
Private Function ParseDoubleFR3(ByVal v As Variant) As Double
    ParseDoubleFR3 = ParseDoubleFR_Safe(v)
End Function

Private Function NormalizeHeader_FR(ByVal s As String) As String
    Dim t As String

    t = LCase$(Trim$(CStr(s)))
    t = ReplaceAccents_FR(t)
    t = Replace(t, Chr$(160), vbNullString)
    t = Replace(t, " ", vbNullString)
    t = Replace(t, "_", vbNullString)
    t = Replace(t, "-", vbNullString)

    NormalizeHeader_FR = t
End Function

Private Function NormalizeHeader2(ByVal s As String) As String
    NormalizeHeader2 = NormalizeHeader_FR(s)
End Function
Private Function NormalizeHeader3(ByVal s As String) As String
    NormalizeHeader3 = NormalizeHeader_FR(s)
End Function

Private Function ReplaceAccents_FR(ByVal s As String) As String
    Dim t As String
    t = s

    t = Replace(t, ChrW$(233), "e")
    t = Replace(t, ChrW$(232), "e")
    t = Replace(t, ChrW$(234), "e")
    t = Replace(t, ChrW$(235), "e")
    t = Replace(t, ChrW$(201), "E")
    t = Replace(t, ChrW$(200), "E")
    t = Replace(t, ChrW$(202), "E")
    t = Replace(t, ChrW$(203), "E")

    t = Replace(t, ChrW$(224), "a")
    t = Replace(t, ChrW$(226), "a")
    t = Replace(t, ChrW$(228), "a")
    t = Replace(t, ChrW$(192), "A")
    t = Replace(t, ChrW$(194), "A")
    t = Replace(t, ChrW$(196), "A")

    t = Replace(t, ChrW$(238), "i")
    t = Replace(t, ChrW$(239), "i")
    t = Replace(t, ChrW$(206), "I")
    t = Replace(t, ChrW$(207), "I")

    t = Replace(t, ChrW$(244), "o")
    t = Replace(t, ChrW$(246), "o")
    t = Replace(t, ChrW$(212), "O")
    t = Replace(t, ChrW$(214), "O")

    t = Replace(t, ChrW$(249), "u")
    t = Replace(t, ChrW$(251), "u")
    t = Replace(t, ChrW$(252), "u")
    t = Replace(t, ChrW$(217), "U")
    t = Replace(t, ChrW$(219), "U")
    t = Replace(t, ChrW$(220), "U")

    t = Replace(t, ChrW$(231), "c")
    t = Replace(t, ChrW$(199), "C")

    ReplaceAccents_FR = t
End Function

Private Function ReplaceAccents2(ByVal s As String) As String
    ReplaceAccents2 = ReplaceAccents_FR(s)
End Function
Private Function ReplaceAccents3(ByVal s As String) As String
    ReplaceAccents3 = ReplaceAccents_FR(s)
End Function

Private Function SanitizeLabel_FR(ByVal s As String) As String
    Dim t As String
    Dim i As Long
    Dim ch As String
    Dim code As Long

    t = ReplaceAccents_FR(CStr(s))
    t = Replace(t, "@", "")

    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        code = AscW(ch)
        If code >= 0 And code < 32 Then Mid$(t, i, 1) = " "
    Next i

    Do While InStr(1, t, "  ", vbBinaryCompare) > 0
        t = Replace(t, "  ", " ")
    Loop

    SanitizeLabel_FR = Trim$(t)
End Function

Private Function SanitizeLabel2(ByVal s As String) As String
    SanitizeLabel2 = SanitizeLabel_FR(s)
End Function
Private Function SanitizeLabel3(ByVal s As String) As String
    SanitizeLabel3 = SanitizeLabel_FR(s)
End Function

Private Sub SortStringArray(ByRef arr As Variant)
    Dim i As Long
    Dim j As Long
    Dim tmp As String

    On Error GoTo EH
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If CStr(arr(j)) < CStr(arr(i)) Then
                tmp = CStr(arr(i))
                arr(i) = CStr(arr(j))
                arr(j) = tmp
            End If
        Next j
    Next i
    Exit Sub
EH:
End Sub

Private Sub SortStringArray3(ByRef arr As Variant)
    SortStringArray arr
End Sub

Private Function KeepDigits3(ByVal s As String) As String
    KeepDigits3 = modUtils.KeepDigits(s)
End Function

Private Function FecBG_ArrayHasRows(ByVal arr As Variant) As Boolean
    FecBG_ArrayHasRows = BGCompil_ArrayHasRows(arr)
End Function

Private Function NzAny(ByVal v As Variant, ByVal fallback As Variant) As Variant
    If IsError(v) Then
        NzAny = fallback
    ElseIf IsNull(v) Then
        NzAny = fallback
    ElseIf IsEmpty(v) Then
        NzAny = fallback
    Else
        NzAny = v
    End If
End Function

Private Function Round2(ByVal d As Double) As Double
    Round2 = VBA.Round(d, 2)
End Function
