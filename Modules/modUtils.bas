Attribute VB_Name = "modUtils"
' ============================================================
' modUtils.bas
' Fonctions utilitaires pures sans dependance Excel
' ============================================================
Option Explicit

Public Function KeepDigits(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    s = Trim$(s)
    s = Replace(s, ChrW(160), "")
    s = Replace(s, " ", "")
    s = Replace(s, "'", "")
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "#" Then out = out & ch
    Next i
    KeepDigits = out
End Function

Public Function CleanFileName(ByVal s As String) As String
    Dim bad As Variant, i As Long, j As Long
    Dim ch As String, outName As String

    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If AscW(ch) >= 32 Then outName = outName & ch
    Next i

    bad = Array("/", ":", "*", "?", """", "<", ">", "|")
    For j = LBound(bad) To UBound(bad)
        outName = Replace(outName, bad(j), "_")
    Next j

    outName = Trim$(outName)
    Do While InStr(outName, "  ") > 0
        outName = Replace(outName, "  ", " ")
    Loop
    Do While Len(outName) > 0 And (Right$(outName, 1) = " " Or Right$(outName, 1) = ".")
        outName = Left$(outName, Len(outName) - 1)
    Loop
    If Len(outName) = 0 Then outName = "export"
    If Len(outName) > 180 Then outName = Left$(outName, 180)

    CleanFileName = outName
End Function

Public Function CleanFullPath(ByVal fullPath As String) As String
    Dim p As Long, folder As String, fileName As String
    p = InStrRev(fullPath, Application.PathSeparator)
    If p > 0 Then
        folder = Left$(fullPath, p)
        fileName = Mid$(fullPath, p + 1)
    Else
        folder = ""
        fileName = fullPath
    End If
    fileName = CleanFileName(fileName)
    Do While Right$(fileName, 1) = " " Or Right$(fileName, 1) = "."
        fileName = Left$(fileName, Len(fileName) - 1)
    Loop
    CleanFullPath = folder & fileName
End Function

Public Function BuildDefaultFileName(ByVal client As String, ByVal exo As String, ByVal ver As String) As String
    Dim d As Date
    Dim aamm As String, jjmmaaaa As String
    Dim safeClient As String, safeVer As String

    safeClient = CleanFileName(client)
    safeVer = CleanFileName(ver)

    If Not IsValidExerciceDate_UI(exo, d) Then
        BuildDefaultFileName = CleanFileName("XXXX " & safeClient & " - Leads au XXXXXXXX V" & safeVer & ".xlsx")
        Exit Function
    End If

    aamm = Format$(d, "yymm")
    jjmmaaaa = Format$(d, "ddmmyyyy")
    BuildDefaultFileName = CleanFileName(aamm & " " & safeClient & _
                           " - Leads au " & jjmmaaaa & _
                           " V" & safeVer & ".xlsx")
End Function

Public Function GetFolderFromPath(ByVal fullPath As String) As String
    Dim p As Long
    p = InStrRev(fullPath, Application.PathSeparator)
    If p > 0 Then
        GetFolderFromPath = Left$(fullPath, p)
    Else
        GetFolderFromPath = ""
    End If
End Function

Public Function IsValidExerciceDate_UI(ByVal s As String, ByRef outDate As Date) As Boolean
    Dim t As String
    Dim parts() As String
    Dim jj As Long, mm As Long, aa As Long

    On Error GoTo EH

    t = Trim$(s)
    t = Replace(t, vbTab, "")
    t = Replace(t, " ", "")
    t = Replace(t, ".", "/")
    t = Replace(t, "-", "/")
    Do While InStr(1, t, "//", vbBinaryCompare) > 0
        t = Replace(t, "//", "/")
    Loop
    If Len(t) = 0 Then GoTo CleanExit

    ' 1) JJ/MM/AA
    If InStr(1, t, "/", vbBinaryCompare) > 0 Then
        parts = Split(t, "/")
        If UBound(parts) = 2 Then
            If Len(parts(0)) = 2 And Len(parts(1)) = 2 And Len(parts(2)) = 2 Then
                If IsNumeric(parts(0)) And IsNumeric(parts(1)) And IsNumeric(parts(2)) Then
                    jj = CLng(parts(0))
                    mm = CLng(parts(1))
                    aa = YearFrom2Digits(CLng(parts(2)))
                    If TryBuildValidDate(jj, mm, aa, outDate) Then
                        IsValidExerciceDate_UI = True
                        GoTo CleanExit
                    End If
                End If
            End If

            ' 2) JJ/MM/AAAA
            If Len(parts(0)) = 2 And Len(parts(1)) = 2 And Len(parts(2)) = 4 Then
                If IsNumeric(parts(0)) And IsNumeric(parts(1)) And IsNumeric(parts(2)) Then
                    jj = CLng(parts(0))
                    mm = CLng(parts(1))
                    aa = CLng(parts(2))
                    If TryBuildValidDate(jj, mm, aa, outDate) Then
                        IsValidExerciceDate_UI = True
                        GoTo CleanExit
                    End If
                End If
            End If
        End If
    End If

    ' 3) JJMMAA
    If Len(t) = 6 And IsNumeric(t) And InStr(1, t, "/", vbBinaryCompare) = 0 Then
        jj = CLng(Left$(t, 2))
        mm = CLng(Mid$(t, 3, 2))
        aa = YearFrom2Digits(CLng(Right$(t, 2)))
        If TryBuildValidDate(jj, mm, aa, outDate) Then
            IsValidExerciceDate_UI = True
            GoTo CleanExit
        End If
    End If

    ' 4) JJMMAAAA
    If Len(t) = 8 And IsNumeric(t) And InStr(1, t, "/", vbBinaryCompare) = 0 Then
        jj = CLng(Left$(t, 2))
        mm = CLng(Mid$(t, 3, 2))
        aa = CLng(Right$(t, 4))
        If TryBuildValidDate(jj, mm, aa, outDate) Then
            IsValidExerciceDate_UI = True
            GoTo CleanExit
        End If
    End If

    ' 5) MMAA
    If Len(t) = 4 And IsNumeric(t) And InStr(1, t, "/", vbBinaryCompare) = 0 Then
        mm = CLng(Left$(t, 2))
        aa = YearFrom2Digits(CLng(Right$(t, 2)))
        jj = LastDayOfMonth(aa, mm)
        If TryBuildValidDate(jj, mm, aa, outDate) Then
            IsValidExerciceDate_UI = True
            GoTo CleanExit
        End If
    End If

    ' 6) MM/AA
    If InStr(1, t, "/", vbBinaryCompare) > 0 Then
        parts = Split(t, "/")
        If UBound(parts) = 1 Then
            If Len(parts(0)) = 2 And Len(parts(1)) = 2 Then
                If IsNumeric(parts(0)) And IsNumeric(parts(1)) Then
                    mm = CLng(parts(0))
                    aa = YearFrom2Digits(CLng(parts(1)))
                    jj = LastDayOfMonth(aa, mm)
                    If TryBuildValidDate(jj, mm, aa, outDate) Then
                        IsValidExerciceDate_UI = True
                        GoTo CleanExit
                    End If
                End If
            End If
        End If
    End If

CleanExit:
    Exit Function
EH:
    Debug.Print "IsValidExerciceDate_UI error " & Err.Number & " : " & Err.Description
    IsValidExerciceDate_UI = False
    Resume CleanExit
End Function

Private Function TryBuildValidDate(ByVal dd As Long, ByVal mm As Long, ByVal yyyy As Long, ByRef outDate As Date) As Boolean
    Dim dt As Date

    On Error GoTo EH
    If Not IsValidDayMonthYear(dd, mm, yyyy) Then GoTo CleanExit

    On Error Resume Next
    dt = CDate(DateSerial(yyyy, mm, dd))
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo EH
        GoTo CleanExit
    End If
    On Error GoTo EH

    outDate = dt
    TryBuildValidDate = True

CleanExit:
    Exit Function
EH:
    Debug.Print "TryBuildValidDate error " & Err.Number & " : " & Err.Description
    TryBuildValidDate = False
    Resume CleanExit
End Function
Public Function YearFrom2Digits(ByVal yy As Long) As Long
    ' Regle simple : 00-50 => 2000-2050 ; 51-99 => 1951-1999
    If yy < 0 Or yy > 99 Then
        YearFrom2Digits = yy
    ElseIf yy <= 50 Then
        YearFrom2Digits = 2000 + yy
    Else
        YearFrom2Digits = 1900 + yy
    End If
End Function

Public Function LastDayOfMonth(ByVal yyyy As Long, ByVal mm As Long) As Long
    ' DateSerial(yyyy, mm+1, 0) = dernier jour du mois mm
    LastDayOfMonth = Day(DateSerial(yyyy, mm + 1, 0))
End Function

Public Function IsValidMonthYear(ByVal mm As Long, ByVal yyyy As Long) As Boolean
    If yyyy < 1900 Or yyyy > 2100 Then Exit Function
    If mm < 1 Or mm > 12 Then Exit Function
    IsValidMonthYear = True
End Function

Public Function IsValidDayMonthYear(ByVal dd As Long, ByVal mm As Long, ByVal yyyy As Long) As Boolean
    If Not IsValidMonthYear(mm, yyyy) Then Exit Function
    If dd < 1 Or dd > 31 Then Exit Function

    Dim dt As Date
    dt = DateSerial(yyyy, mm, dd)

    ' Emp�che les d�bordements (32/01 => 01/02)
    If Day(dt) <> dd Or Month(dt) <> mm Or Year(dt) <> yyyy Then Exit Function

    IsValidDayMonthYear = True
End Function

Public Function FileExists(ByVal fullPath As String) As Boolean
    On Error Resume Next
    FileExists = (Len(Dir$(fullPath, vbNormal)) > 0)
    On Error GoTo 0
End Function

Public Function TryDeleteFile(ByVal fullPath As String, ByRef errMsg As String) As Boolean
    On Error GoTo EH
    If FileExists(fullPath) Then
        SetAttr fullPath, vbNormal
        Kill fullPath
    End If
    TryDeleteFile = True
    Exit Function
EH:
    errMsg = Err.Description
    TryDeleteFile = False
End Function

Public Function LastRowWithTextInCol(ByVal ws As Worksheet, ByVal ColLetter As String) As Long
    ' Cherche la derniere cellule contenant quelque chose (valeur ou formule non vide) dans une colonne.
    Dim rng As Range
    On Error Resume Next
    Set rng = ws.Columns(ColLetter).Find(What:="*", _
                                         LookIn:=xlFormulas, _
                                         LookAt:=xlPart, _
                                         SearchOrder:=xlByRows, _
                                         SearchDirection:=xlPrevious, _
                                         MatchCase:=False)
    On Error GoTo 0
    If rng Is Nothing Then
        LastRowWithTextInCol = 0
    Else
        LastRowWithTextInCol = rng.Row
    End If
End Function

Public Function SheetExists(ByVal wb As Workbook, ByVal sName As String) As Boolean
    On Error Resume Next
    SheetExists = Not wb.Worksheets(sName) Is Nothing
    On Error GoTo 0
End Function

Public Function GetUniqueSheetName(ByVal wb As Workbook, ByVal baseName As String) As String
    Dim nameTry As String, k As Long
    nameTry = Left$(baseName, 31)
    If Not SheetExists(wb, nameTry) Then
        GetUniqueSheetName = nameTry
        Exit Function
    End If
    k = 2
    Do
        nameTry = Left$(baseName, 28) & "_" & CStr(k)  ' garde <=31
        k = k + 1
    Loop While SheetExists(wb, nameTry)
    GetUniqueSheetName = nameTry
End Function

Public Function GetLastUsedRowSafe(ByVal ws As Worksheet) As Long
    Dim rng As Range
    On Error Resume Next
    Set rng = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    On Error GoTo 0
    If rng Is Nothing Then Exit Function
    GetLastUsedRowSafe = rng.Row
End Function

Public Function GetLastUsedColSafe(ByVal ws As Worksheet) As Long
    Dim rng As Range
    On Error Resume Next
    Set rng = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    On Error GoTo 0
    If rng Is Nothing Then Exit Function
    GetLastUsedColSafe = rng.Column
End Function

Public Function BSX_IsBlank(ByVal v As Variant) As Boolean
    If IsError(v) Then
        BSX_IsBlank = False
    Else
        BSX_IsBlank = (Len(Trim$(CStr(v))) = 0)
    End If
End Function

Public Function NormalizeAccents(ByVal txt As String) As String
    Dim t As String
    t = CStr(txt)
    t = Replace(t, "é", "�")
    t = Replace(t, "è", "�")
    t = Replace(t, "ê", "�")
    t = Replace(t, "ë", "�")
    t = Replace(t, "à", "�")
    t = Replace(t, "â", "�")
    t = Replace(t, "ä", "�")
    t = Replace(t, "ù", "�")
    t = Replace(t, "û", "�")
    t = Replace(t, "ü", "�")
    t = Replace(t, "ô", "�")
    t = Replace(t, "ö", "�")
    t = Replace(t, "ç", "�")
    t = Replace(t, "É", "�")
    t = Replace(t, "À", "�")
    t = Replace(t, "È", "�")
    t = Replace(t, "Ê", "�")
    t = Replace(t, "Ç", "�")
    t = Replace(t, "", "�")
    t = Replace(t, "€", "�")
    t = Replace(t, "’", "�")
    t = Replace(t, "–", "�")
    t = Replace(t, "—", "�")
    NormalizeAccents = t
End Function

Public Function NormalizeAccentsOnLabelColumn(ByVal arrIn As Variant) As Variant
    Dim arrOut As Variant
    Dim i As Long

    If IsEmpty(arrIn) Or Not IsArray(arrIn) Then
        NormalizeAccentsOnLabelColumn = arrIn
        Exit Function
    End If

    arrOut = arrIn
    On Error GoTo CleanExit
    For i = LBound(arrOut, 1) To UBound(arrOut, 1)
        arrOut(i, 2) = NormalizeAccents(CStr(arrOut(i, 2)))
    Next i

CleanExit:
    NormalizeAccentsOnLabelColumn = arrOut
End Function

Public Function LineStartsProcedure(ByVal lineText As String) As Boolean
    Dim normalized As String

    normalized = LCase$(Trim$(lineText))
    If Len(normalized) = 0 Then Exit Function

    If Left$(normalized, 11) = "public sub " Then LineStartsProcedure = True: Exit Function
    If Left$(normalized, 12) = "private sub " Then LineStartsProcedure = True: Exit Function
    If Left$(normalized, 11) = "friend sub " Then LineStartsProcedure = True: Exit Function
    If Left$(normalized, 16) = "public function " Then LineStartsProcedure = True: Exit Function
    If Left$(normalized, 17) = "private function " Then LineStartsProcedure = True: Exit Function
    If Left$(normalized, 16) = "friend function " Then LineStartsProcedure = True: Exit Function
    If Left$(normalized, 16) = "public property " Then LineStartsProcedure = True: Exit Function
    If Left$(normalized, 17) = "private property " Then LineStartsProcedure = True: Exit Function
    If Left$(normalized, 16) = "friend property " Then LineStartsProcedure = True: Exit Function
End Function

Public Function ExtractProcedureName(ByVal lineText As String) As String
    Dim tokens() As String
    Dim idx As Long

    tokens = Split(Trim$(lineText))
    If UBound(tokens) < 1 Then Exit Function

    For idx = LBound(tokens) To UBound(tokens)
        Select Case LCase$(tokens(idx))
            Case "sub", "function"
                If idx < UBound(tokens) Then
                    ExtractProcedureName = Trim$(Split(tokens(idx + 1), "(")(0))
                    Exit Function
                End If
            Case "property"
                If idx + 1 <= UBound(tokens) Then
                    Select Case LCase$(tokens(idx + 1))
                        Case "get", "let", "set"
                            If idx + 2 <= UBound(tokens) Then
                                ExtractProcedureName = Trim$(Split(tokens(idx + 2), "(")(0))
                                Exit Function
                            End If
                        Case Else
                            ExtractProcedureName = Trim$(Split(tokens(idx + 1), "(")(0))
                            Exit Function
                    End Select
                End If
        End Select
    Next idx
End Function

Public Function CountOccurrences(ByVal sourceText As String, ByVal targetText As String) As Long
    Dim startPos As Long
    Dim hitPos As Long

    If Len(targetText) = 0 Then Exit Function

    startPos = 1
    Do
        hitPos = InStr(startPos, sourceText, targetText, vbBinaryCompare)
        If hitPos = 0 Then Exit Do
        CountOccurrences = CountOccurrences + 1
        startPos = hitPos + Len(targetText)
    Loop
End Function


