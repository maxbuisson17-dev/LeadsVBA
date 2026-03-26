Attribute VB_Name = "modDA_Param"
' ============================================================
' modDA_Param.bas
' Acces donnees - feuille Param et metadonnees export
' ============================================================
Option Explicit

Public Sub ApplyMetaToSourceParamSheet(ByVal wbSource As Workbook)
    Dim wsParam As Worksheet
    Dim dExo As Date
    Dim unitLabel As String

    On Error GoTo EH
    If wbSource Is Nothing Then Err.Raise vbObjectError + 1401, , "Classeur source introuvable."

    Set wsParam = wbSource.Worksheets("Param")
    If Not IsValidExerciceDate_UI(gExercice, dExo) Then Err.Raise vbObjectError + 1403, , "Date d'exercice invalide."

    If gGenerateInKE Then
        unitLabel = "en K€"
    Else
        unitLabel = "en euros"
    End If

    wsParam.Range("B3").value = gClient
    wsParam.Range("B4").value = dExo
    wsParam.Range("B4").NumberFormat = "dd/mm/yyyy"
    wsParam.Range("B6").value = gVersion
    wsParam.Range("B7").value = unitLabel

    wsParam.Calculate
    wbSource.Calculate
    gMetaAppliedToParam = True

    modKETrace.LogKE "Meta applied to Param | B3=" & CStr(wsParam.Range("B3").value) & _
                     " | B4=" & Format$(CDate(wsParam.Range("B4").value), "dd/mm/yyyy") & _
                     " | B6=" & CStr(wsParam.Range("B6").value) & _
                     " | B7=" & CStr(wsParam.Range("B7").value), "ApplyMetaToSourceParamSheet", "Param", wbSource.name
CleanExit:
    Exit Sub
EH:
    SafeLogNoUI "ApplyMetaToSourceParamSheet failed | Err=" & Err.Number & " | Desc=''" & Err.Description & "''", "ApplyMetaToSourceParamSheet"
    Err.Clear
    Resume CleanExit
End Sub

Public Sub ResetSourceParamPlaceholders(ByVal wbSource As Workbook)
    Dim wsParam As Worksheet
    On Error Resume Next
    Set wsParam = wbSource.Worksheets("Param")
    On Error GoTo 0
    If wsParam Is Nothing Then Exit Sub

    wsParam.Range("B3").value = "'@CLIENT"
    wsParam.Range("B4").value = "'@DATE"
    wsParam.Range("B6").value = "'@VERSION"
    wsParam.Range("B7").value = "'@euros"
End Sub

Public Sub WriteMetaToBS_Export_CDC(ByVal wsBS As Worksheet)
    'DEPRECATED: meta driven by Param sheet formulas
    Exit Sub
End Sub

Public Sub WriteMetaToSIG_Export_CDC(ByVal wsSIG As Worksheet)
    'DEPRECATED: meta driven by Param sheet formulas
    Exit Sub
End Sub

Public Function GetExportModeFromFrmLeadMeta() As eExportMode
    If gExportMode = 0 Then gExportMode = emFS
    GetExportModeFromFrmLeadMeta = gExportMode
End Function
