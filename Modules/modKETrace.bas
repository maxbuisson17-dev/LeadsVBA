Attribute VB_Name = "modKETrace"
Option Explicit

'CHANGE: logger central pour tracer toutes les applications K€ (/1000)
Private Const KE_TRACE_FILE As String = "modLeadsV1_KE_trace.log"

Public Sub LogKE(ByVal msg As String, Optional ByVal procName As String = "", Optional ByVal sheetName As String = "", Optional ByVal wbName As String = "")
    Dim lineOut As String
    Dim ts As String
    Dim f As Integer
    Dim logPath As String

    On Error GoTo EH

    ts = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    lineOut = ts & " | " & IIf(Len(procName) > 0, procName, "-")
    If Len(sheetName) > 0 Then lineOut = lineOut & " | Sheet=" & sheetName
    If Len(wbName) > 0 Then lineOut = lineOut & " | Wb=" & wbName
    lineOut = lineOut & " | " & msg

    Debug.Print "[KE_TRACE] " & lineOut

    logPath = Environ$("TEMP")
    If Len(logPath) = 0 Then GoTo CleanExit
    If Right$(logPath, 1) <> "\" Then logPath = logPath & "\"
    logPath = logPath & KE_TRACE_FILE

    f = FreeFile
    Open logPath For Append As #f
    Print #f, lineOut
    Close #f

CleanExit:
    Exit Sub
EH:
    On Error Resume Next
    If f > 0 Then Close #f
    Debug.Print "[KE_TRACE] LogKE error " & Err.Number & " : " & Err.Description
    Resume CleanExit
End Sub

Public Sub LogWorkbookSheets(ByVal wb As Workbook, ByVal procName As String, Optional ByVal prefix As String = "")
    Dim ws As Worksheet
    Dim names As String

    On Error GoTo EH
    If wb Is Nothing Then
        LogKE prefix & "wb is Nothing", procName
        GoTo CleanExit
    End If

    For Each ws In wb.Worksheets
        If Len(names) > 0 Then names = names & ", "
        names = names & ws.name
    Next ws

    LogKE prefix & "Sheets(" & wb.Worksheets.Count & ")=[" & names & "]", procName, "", wb.name

CleanExit:
    Exit Sub
EH:
    LogKE "LogWorkbookSheets error " & Err.Number & " : " & Err.Description, procName
    Resume CleanExit
End Sub

Public Sub ResetKETraceFile()
    Dim logPath As String
    Dim f As Integer

    On Error GoTo EH
    logPath = Environ$("TEMP")
    If Len(logPath) = 0 Then GoTo CleanExit
    If Right$(logPath, 1) <> "\" Then logPath = logPath & "\"
    logPath = logPath & KE_TRACE_FILE

    f = FreeFile
    Open logPath For Output As #f
    Print #f, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | TRACE RESET"
    Close #f

CleanExit:
    Exit Sub
EH:
    On Error Resume Next
    If f > 0 Then Close #f
    Debug.Print "[KE_TRACE] ResetKETraceFile error " & Err.Number & " : " & Err.Description
    Resume CleanExit
End Sub

Public Sub PrintKEDivisionSummary()
    'CHANGE: synthese rapide de preuve (BS/BS_detail/SIG/SIG_detail)
    Dim logPath As String
    Dim f As Integer
    Dim lineText As String
    Dim cBS As Long, cBSDetail As Long, cSIG As Long, cSIGDetail As Long

    On Error GoTo EH

    logPath = Environ$("TEMP")
    If Len(logPath) = 0 Then GoTo CleanExit
    If Right$(logPath, 1) <> "\" Then logPath = logPath & "\"
    logPath = logPath & KE_TRACE_FILE

    If Len(Dir$(logPath, vbNormal)) = 0 Then
        Debug.Print "[KE_TRACE] No log file found: " & logPath
        GoTo CleanExit
    End If

    f = FreeFile
    Open logPath For Input As #f
    Do While Not EOF(f)
        Line Input #f, lineText
        If InStr(1, lineText, "ApplyKEDivision_BS", vbTextCompare) > 0 And _
           InStr(1, lineText, "APPLY /1000 start", vbTextCompare) > 0 Then cBS = cBS + 1

        If InStr(1, lineText, "ScaleBSDetailToKE", vbTextCompare) > 0 And _
           InStr(1, lineText, "APPLY /1000 start", vbTextCompare) > 0 Then cBSDetail = cBSDetail + 1

        If InStr(1, lineText, "ApplyKEDivision_SIG", vbTextCompare) > 0 And _
           InStr(1, lineText, "APPLY /1000 start", vbTextCompare) > 0 Then cSIG = cSIG + 1

        If InStr(1, lineText, "ScaleSIGDetailToKE", vbTextCompare) > 0 And _
           InStr(1, lineText, "APPLY /1000 start", vbTextCompare) > 0 Then cSIGDetail = cSIGDetail + 1
    Loop
    Close #f

    Debug.Print "[KE_TRACE] SUMMARY | BS=" & cBS & " | BS_detail=" & cBSDetail & " | SIG=" & cSIG & " | SIG_detail=" & cSIGDetail

CleanExit:
    Exit Sub
EH:
    On Error Resume Next
    If f > 0 Then Close #f
    Debug.Print "[KE_TRACE] PrintKEDivisionSummary error " & Err.Number & " : " & Err.Description
    Resume CleanExit
End Sub
