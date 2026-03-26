
Attribute VB_Name = "frmImportBalanceV6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Option Explicit

Private mCurrentPreview As Variant
Private m_PathN As String
Private m_PathN1 As String

Private Sub UserForm_Initialize()

    ModNovancesTheme.AppliquerTheme Me
    ModNovancesTheme.AjouterBarreTitre Me, "Import balances N et N-1"

    
    modSession.ResetImportContext

    SetControlProp "txtN_Path", "Value", ""
    SetControlProp "txtN1_Path", "Value", ""
    SetControlProp "txtN_Path", "Locked", True
    SetControlProp "txtN1_Path", "Locked", True

    SetControlProp "lblN_Preview", "Caption", "Apercu : -"
    SetControlProp "lblN_Period", "Caption", "Periode : -"
    SetControlProp "lblN1_Preview", "Caption", "Apercu : -"

    HideLegacyControls
    EnsureMappingControls
    EnsureBalanceLists
    UpdateBrowseButtons


    
End Sub

Private Sub optN_FEC_Click()
    ' Legacy control kept for compatibility only.
End Sub

Private Sub optN_BAL_Click()
    ' Legacy control kept for compatibility only.
End Sub

Private Sub optN1_FEC_Click()
    ' Legacy control kept for compatibility only.
End Sub

Private Sub optN1_BAL_Click()
    ' Legacy control kept for compatibility only.
End Sub

Private Sub UpdateBrowseButtons()
    SetControlProp "cmdN_Browse", "Enabled", True
    SetControlProp "cmdN1_Browse", "Enabled", True
End Sub

Private Sub cmdN_Browse_Click()
    BrowseAndImport True
End Sub

Private Sub cmdN1_Browse_Click()
    BrowseAndImport False
End Sub

Private Sub BrowseAndImport(ByVal isN As Boolean)
    Dim f As Variant
    Dim arrRaw As Variant
    Dim arr As Variant
    Dim rowCount As Long
    Dim path As String
    Dim idxC As Long, idxL As Long, idxSN As Long, idxSN1 As Long
    Dim info As String
    Dim srcColCount As Long
    Dim mode4 As eBalance4ColsMode

    On Error GoTo EH

    f = Application.GetOpenFilename( _
        "Tous fichiers supportes (*.txt;*.csv;*.dat;*.xls;*.xlsx;*.xlsm),*.txt;*.csv;*.dat;*.xls;*.xlsx;*.xlsm", , _
        IIf(isN, "Selectionner le fichier N", "Selectionner le fichier N-1"))

    If VarType(f) = vbBoolean Then GoTo CleanExit

    path = Trim$(CStr(f))
    If Len(path) = 0 Then GoTo CleanExit

    If Not IsSupportedImportExt(path) Then
        MsgBox "Extension non supportee : " & path, vbExclamation
        GoTo CleanExit
    End If

    srcColCount = modImportUnified.ImportFile_DetectSourceColumnCount(path)
    If srcColCount = 4 Then
        ' Choix type/colonne dans frmBalanceType
        mode4 = AskBalance4ColsMode()
        If mode4 = b4None Then GoTo CleanExit
        arrRaw = modImportUnified.ImportFile4Cols_ToBalance3Cols(path, mode4, info)
    Else
        ' 2/3 colonnes -> comportement existant
        arrRaw = modImportUnified.ImportFile_ToBalance3Cols(path, info)
    End If
    If Not IsArrayHasRows(arrRaw, 3) Then
        MsgBox "Import impossible : aucune ligne exploitable detectee.", vbExclamation
        GoTo CleanExit
    End If

    arr = EnsureHeaderRow3(arrRaw)
    If Not IsArrayHasRows(arr, 3) Then
        MsgBox "Import impossible : format de balance invalide.", vbExclamation
        GoTo CleanExit
    End If

    rowCount = UBound(arr, 1) - IIf(IsHeaderLikeRow3(arr), 1, 0)

    If isN Then
        m_PathN = path
        SetControlProp "txtN_Path", "Value", path

        gPathN = path
        gArrN = arr
        gImportedN = True
        gImportParams.SourceN = srcBalance
        gImportParams.pathN = path

        SetControlProp "lblN_Preview", "Caption", "Import OK (" & rowCount & " lignes)"
        SetControlProp "lblN_Period", "Caption", IIf(Len(info) > 0, info, "Periode : importee")

        mCurrentPreview = arr
        modLeadsWizard.Balance_GuessColumns arr, idxC, idxL, idxSN, idxSN1
        PopulateMappingCombos arr, idxC, idxL, idxSN, idxSN1
        FillListBoxWithBalance GetBalanceListBox(True), arr
    Else
        m_PathN1 = path
        SetControlProp "txtN1_Path", "Value", path

        gPathN1 = path
        gArrN1 = arr
        gImportedN1 = True
        gImportParams.SourceN1 = srcBalance
        gImportParams.pathN1 = path

        SetControlProp "lblN1_Preview", "Caption", "Import OK (" & rowCount & " lignes)"

        mCurrentPreview = arr
        modLeadsWizard.Balance_GuessColumns arr, idxC, idxL, idxSN, idxSN1
        PopulateMappingCombos arr, idxC, idxL, idxSN, idxSN1
        FillListBoxWithBalance GetBalanceListBox(False), arr
    End If

    gImportParamsReady = True
    GoTo CleanExit

EH:
    Debug.Print "BrowseAndImport error " & Err.Number & " : " & Err.Description
    MsgBox "Erreur pendant l'import : " & Err.Description, vbExclamation
    Resume CleanExit
CleanExit:
    Exit Sub
End Sub

Private Function AskBalance4ColsMode() As eBalance4ColsMode
    Dim uf As frmBalanceType

    On Error GoTo EH

    'CHANGE: frmImportBalanceV6 est deja affiche, ne pas forcer un Show modeless
    Me.Repaint
    DoEvents

    Set uf = New frmBalanceType
    uf.Show vbModal

    AskBalance4ColsMode = uf.SelectedMode

    'CHANGE: UX retour form appelant
    DoEvents
    Me.Repaint

CleanExit:
    On Error Resume Next
    If Not uf Is Nothing Then Unload uf
    On Error GoTo 0
    Exit Function
EH:
    Debug.Print "AskBalance4ColsMode error " & Err.Number & " : " & Err.Description
    AskBalance4ColsMode = b4None
    Resume CleanExit
End Function

Private Sub cmdContinue_Click()
    Dim idxC As Long, idxL As Long, idxSN As Long, idxSN1 As Long
    Dim gC As Long, gL As Long, gSN As Long, gSN1 As Long
    Dim arrN4 As Variant
    Dim arrN14 As Variant
    Dim arrCompil As Variant
    Dim arrFinal As Variant
    Dim outInfo As String
    Dim outMaxLen As Long

    On Error GoTo EH

    If Not modSession.HasBalanceN Then
        MsgBox "Importe d'abord la balance N.", vbExclamation
        GoTo CleanExit
    End If

    If Not modSession.HasBalanceN1 Then
        MsgBox "Importe d'abord la balance N-1.", vbExclamation
        GoTo CleanExit
    End If

    If Not ValidateCurrentSelections() Then GoTo CleanExit

    modLeadsWizard.Balance_GuessColumns gArrN, gC, gL, gSN, gSN1

    idxC = GetSelectedColumn("cboCompte", gC)
    idxL = GetSelectedColumn("cboLib", gL)
    idxSN = GetSelectedColumn("cboSoldeN", gSN)
    idxSN1 = GetSelectedColumn("cboSoldeN1", gSN1)

    arrN4 = modLeadsWizard.Balance_MapTo4Cols(gArrN, idxC, idxL, idxSN, 0)
    arrN4 = modLeadsWizard.TransformerBG_Array(arrN4)

    modLeadsWizard.Balance_GuessColumns gArrN1, gC, gL, gSN, gSN1
    If idxC <= 0 Then idxC = gC
    If idxL < 0 Then idxL = gL
    If idxSN1 = 0 Then
        If gSN1 > 0 Then
            idxSN1 = gSN1
        Else
            idxSN1 = gSN
        End If
    End If

    arrN14 = modLeadsWizard.Balance_MapTo4Cols(gArrN1, idxC, idxL, 0, idxSN1)
    arrN14 = modLeadsWizard.TransformerBG_Array(arrN14)

    arrCompil = modLeadsWizard.BGCompil_FromBalanceArrays(arrN4, arrN14, outMaxLen, outInfo)
    CleanLabelColumn2_NoAccents_InArray arrCompil
    If Not modLeadsWizard.BGCompil_ArrayHasRows(arrCompil) Then
        MsgBox "La compilation N / N-1 n'a produit aucune ligne.", vbExclamation
        GoTo CleanExit
    End If

    arrFinal = modLeadsWizard.Ensure4Cols(arrCompil)
    If IsEmpty(arrFinal) Or Not IsArray(arrFinal) Then
        MsgBox "Transformation finale impossible.", vbCritical
        GoTo CleanExit
    End If

    gArrCompiled = arrFinal
    gFullData = arrFinal
    gMaxAcctLen = modOrchestrator.ComputeMaxAccountLen(arrFinal)

    If Len(Trim$(gPathN)) > 0 Then
        gBalancePath = gPathN
    Else
        gBalancePath = gPathN1
    End If

    modCDC.BuildControlReportFromFullData

    ' --- Navigation identique a RunYesFlow ---
    If gOkToGenerate Then
        Unload Me
        modExportEngine.EnsureWorkingSheetsHidden False
        modDA_BG.ImportIntoBG_FromFullData
        Dim ufMeta As New frmLeadMeta
        ufMeta.Show vbModal
        Unload ufMeta
    Else
        Unload Me
        Dim ufErr As New frmBGError
        ufErr.Show vbModal
    End If
    GoTo CleanExit

EH:
    Debug.Print "cmdContinue_Click error " & Err.Number & " : " & Err.Description
    MsgBox "Erreur cmdContinue : " & Err.Description, vbCritical
    Resume CleanExit
CleanExit:
    Exit Sub
End Sub

Private Sub cmdProcess_Click()
    If Not ValidateCurrentSelections() Then Exit Sub
    cmdContinue_Click
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub HideLegacyControls()
    SetControlProp "chkDropZero", "Visible", False
    SetControlProp "chkRoundEuro", "Visible", False
    SetControlProp "chkToKE", "Visible", False
    SetControlProp "chkKeepAN", "Visible", False

    SetControlPropIfExists "optN_FEC", "Visible", False
    SetControlPropIfExists "optN_BAL", "Visible", False
    SetControlPropIfExists "optN1_FEC", "Visible", False
    SetControlPropIfExists "optN1_BAL", "Visible", False
    SetControlPropIfExists "fraN", "Visible", False
    SetControlPropIfExists "fraN1", "Visible", False
End Sub

Private Sub EnsureMappingControls()
    On Error Resume Next
    If Not Me.Controls("lstPreview") Is Nothing Then
        Me.Controls("lstPreview").Visible = False
    End If

    If Me.Controls("cboCompte") Is Nothing Then Me.Controls.Add "Forms.ComboBox.1", "cboCompte", True
    If Me.Controls("cboLib") Is Nothing Then Me.Controls.Add "Forms.ComboBox.1", "cboLib", True
    If Me.Controls("cboSoldeN") Is Nothing Then Me.Controls.Add "Forms.ComboBox.1", "cboSoldeN", True
    If Me.Controls("cboSoldeN1") Is Nothing Then Me.Controls.Add "Forms.ComboBox.1", "cboSoldeN1", True

    LayoutCombo "cboCompte", 360, 5600, 3000
    LayoutCombo "cboLib", 3600, 5600, 3000
    LayoutCombo "cboSoldeN", 6840, 5600, 3000
    LayoutCombo "cboSoldeN1", 10080, 5600, 3000
    On Error GoTo 0
End Sub

Private Sub EnsureBalanceLists()
    Dim LbN As MSForms.ListBox
    Dim LbN1 As MSForms.ListBox

    Set LbN = GetBalanceListBox(True)
    Set LbN1 = GetBalanceListBox(False)

    If LbN Is Nothing Then
        Set LbN = Me.Controls.Add("Forms.ListBox.1", "LstN", True)
        LbN.Left = 360
        LbN.Top = 2550
        LbN.Width = 7000
        LbN.Height = 2400
    End If

    If LbN1 Is Nothing Then
        Set LbN1 = Me.Controls.Add("Forms.ListBox.1", "LstN1", True)
        LbN1.Left = 7560
        LbN1.Top = 2550
        LbN1.Width = 7000
        LbN1.Height = 2400
    End If

    ApplyBalanceListStyle LbN
    ApplyBalanceListStyle LbN1
End Sub

Private Function GetBalanceListBox(ByVal isN As Boolean) As MSForms.ListBox
    If isN Then
        Set GetBalanceListBox = FirstExistingListBox(Array("LstN", "lstN", "lstPreviewN"))
    Else
        Set GetBalanceListBox = FirstExistingListBox(Array("LstN1", "lstN1", "lstPreviewN1"))
    End If
End Function

Private Function FirstExistingListBox(ByVal names As Variant) As MSForms.ListBox
    Dim i As Long
    Dim c As Control

    On Error Resume Next
    For i = LBound(names) To UBound(names)
        Set c = Me.Controls(CStr(names(i)))
        If Not c Is Nothing Then
            If TypeName(c) = "ListBox" Then
                Set FirstExistingListBox = c
                Exit Function
            End If
        End If
        Set c = Nothing
    Next i
    On Error GoTo 0
End Function

Private Sub ApplyBalanceListStyle(ByVal lb As MSForms.ListBox)
    If lb Is Nothing Then Exit Sub
    lb.Clear
    lb.ColumnCount = 3
    lb.ColumnHeads = False
    lb.ColumnWidths = GetBalanceColumnWidths()
End Sub

Private Function GetBalanceColumnWidths() As String
    Dim lbPreview As MSForms.ListBox
    Dim w As String
    Dim parts() As String

    GetBalanceColumnWidths = "120 pt;220 pt;120 pt"

    On Error Resume Next
    Set lbPreview = Me.Controls("lstPreview")
    On Error GoTo 0
    If lbPreview Is Nothing Then Exit Function

    w = Trim$(CStr(lbPreview.ColumnWidths))
    If Len(w) = 0 Then Exit Function

    parts = Split(w, ";")
    If UBound(parts) >= 2 Then
        GetBalanceColumnWidths = Trim$(parts(0)) & ";" & Trim$(parts(1)) & ";" & Trim$(parts(2))
    End If
End Function

Private Function GetBalanceArrayForN(ByVal path As String) As Variant
    Dim info As String
    GetBalanceArrayForN = modImportUnified.ImportFile_ToBalance3Cols(path, info)
End Function

Private Function GetBalanceArrayForN1(ByVal path As String) As Variant
    Dim info As String
    GetBalanceArrayForN1 = modImportUnified.ImportFile_ToBalance3Cols(path, info)
End Function

Private Function EnsureHeaderRow3(ByVal arr As Variant) As Variant
    Dim n As Long
    Dim i As Long
    Dim outArr() As Variant

    If Not IsArrayHasRows(arr, 3) Then Exit Function
    If IsHeaderLikeRow3(arr) Then
        EnsureHeaderRow3 = arr
        Exit Function
    End If

    n = UBound(arr, 1)
    ReDim outArr(1 To n + 1, 1 To 3)
    outArr(1, 1) = "Compte"
    outArr(1, 2) = "Libelle"
    outArr(1, 3) = "Solde"

    For i = 1 To n
        outArr(i + 1, 1) = arr(i, 1)
        outArr(i + 1, 2) = arr(i, 2)
        outArr(i + 1, 3) = arr(i, 3)
    Next i

    EnsureHeaderRow3 = outArr
End Function

Private Function IsHeaderLikeRow3(ByVal arr As Variant) As Boolean
    Dim h1 As String, h2 As String, h3 As String

    On Error GoTo EH
    h1 = NormalizeHeaderLocal(CStr(arr(1, 1)))
    h2 = NormalizeHeaderLocal(CStr(arr(1, 2)))
    h3 = NormalizeHeaderLocal(CStr(arr(1, 3)))

    If h1 = "compte" Or h1 = "comptenum" Or h1 = "comptenumero" Then
        If h2 = "libelle" Or h2 = "comptelib" Then
            If h3 = "solde" Or h3 = "debit" Or h3 = "credit" Then
                IsHeaderLikeRow3 = True
            End If
        End If
    End If

    GoTo CleanExit
EH:
    Debug.Print "IsHeaderLikeRow3 error " & Err.Number & " : " & Err.Description
    IsHeaderLikeRow3 = False
    Resume CleanExit
CleanExit:
    Exit Function
End Function

Private Function NormalizeHeaderLocal(ByVal s As String) As String
    Dim t As String
    t = LCase$(Trim$(s))
    t = Replace(t, Chr$(160), "")
    t = Replace(t, " ", "")
    t = Replace(t, "_", "")
    t = Replace(t, "-", "")
    t = Replace(t, ChrW$(233), "e")
    t = Replace(t, ChrW$(232), "e")
    t = Replace(t, ChrW$(234), "e")
    t = Replace(t, ChrW$(235), "e")
    t = Replace(t, ChrW$(231), "c")
    NormalizeHeaderLocal = t
End Function

Private Sub FillListBoxWithBalance(ByVal lb As MSForms.ListBox, ByVal arr As Variant)
    Dim i As Long
    Dim startRow As Long

    If lb Is Nothing Then Exit Sub

    ApplyBalanceListStyle lb

    If Not IsArrayHasRows(arr, 3) Then Exit Sub

    startRow = IIf(IsHeaderLikeRow3(arr), 2, 1)

    For i = startRow To UBound(arr, 1)
        lb.AddItem CStr(arr(i, 1))
        lb.List(lb.ListCount - 1, 1) = CStr(arr(i, 2))
        lb.List(lb.ListCount - 1, 2) = FormatBalanceAmount(arr(i, 3))
    Next i
End Sub

Private Function FormatBalanceAmount(ByVal v As Variant) As String
    If IsNumeric(v) Then
        FormatBalanceAmount = Format$(CDbl(v), "0.00")
    Else
        FormatBalanceAmount = CStr(v)
    End If
End Function

Private Function IsArrayHasRows(ByVal arr As Variant, Optional ByVal minCols As Long = 1) As Boolean
    On Error GoTo EH
    IsArrayHasRows = IsArray(arr) And UBound(arr, 1) >= 1 And UBound(arr, 2) >= minCols
    GoTo CleanExit
EH:
    Debug.Print "IsArrayHasRows error " & Err.Number & " : " & Err.Description
    IsArrayHasRows = False
    Resume CleanExit
CleanExit:
    Exit Function
End Function


Private Sub LayoutCombo(ByVal ctrlName As String, ByVal l As Double, ByVal t As Double, ByVal w As Double)
    On Error Resume Next
    With Me.Controls(ctrlName)
        .Left = l
        .Top = t
        .Width = w
        .Height = 20
        .Style = fmStyleDropDownList
    End With
    On Error GoTo 0
End Sub

Private Sub PopulateMappingCombos(ByVal arr As Variant, ByVal idxC As Long, ByVal idxL As Long, ByVal idxSN As Long, ByVal idxSN1 As Long)
    Dim m As Long
    Dim j As Long
    Dim header As String

    On Error GoTo EH
    If IsEmpty(arr) Or Not IsArray(arr) Then GoTo CleanExit

    m = UBound(arr, 2)

    ClearCombo "cboCompte"
    ClearCombo "cboLib"
    ClearCombo "cboSoldeN"
    ClearCombo "cboSoldeN1"

    AddComboItem "cboSoldeN1", "0 - Aucun"

    For j = 1 To m
        header = CStr(arr(1, j))
        AddComboItem "cboCompte", CStr(j) & " - " & ColLetter(j) & " - " & header
        AddComboItem "cboLib", CStr(j) & " - " & ColLetter(j) & " - " & header
        AddComboItem "cboSoldeN", CStr(j) & " - " & ColLetter(j) & " - " & header
        AddComboItem "cboSoldeN1", CStr(j) & " - " & ColLetter(j) & " - " & header
    Next j

    SetComboByColumn "cboCompte", idxC
    SetComboByColumn "cboLib", idxL
    SetComboByColumn "cboSoldeN", idxSN
    SetComboByColumn "cboSoldeN1", idxSN1
    GoTo CleanExit
EH:
    Debug.Print "PopulateMappingCombos error " & Err.Number & " : " & Err.Description
    Resume CleanExit
CleanExit:
    Exit Sub
End Sub

Private Function GetSelectedColumn(ByVal comboName As String, ByVal defaultIdx As Long) As Long
    Dim v As String
    Dim p As Long
    Dim s As String

    On Error GoTo EH

    v = CStr(GetControlValue(comboName, "Value", ""))
    If Len(v) = 0 Then
        GetSelectedColumn = defaultIdx
        GoTo CleanExit
    End If

    p = InStr(1, v, " - ", vbBinaryCompare)
    If p > 1 Then
        s = Left$(v, p - 1)
    Else
        s = v
    End If

    If IsNumeric(s) Then
        GetSelectedColumn = CLng(s)
    Else
        GetSelectedColumn = defaultIdx
    End If

    GoTo CleanExit
EH:
    Debug.Print "GetSelectedColumn error " & Err.Number & " : " & Err.Description
    GetSelectedColumn = defaultIdx
    Resume CleanExit
CleanExit:
    Exit Function
End Function

Private Sub ClearCombo(ByVal comboName As String)
    On Error Resume Next
    Me.Controls(comboName).Clear
    On Error GoTo 0
End Sub

Private Sub AddComboItem(ByVal comboName As String, ByVal txt As String)
    On Error Resume Next
    Me.Controls(comboName).AddItem txt
    On Error GoTo 0
End Sub

Private Sub SetComboByColumn(ByVal comboName As String, ByVal idx As Long)
    Dim i As Long
    Dim txt As String
    Dim target As String

    On Error Resume Next
    If idx <= 0 Then
        Me.Controls(comboName).value = "0 - Aucun"
        Exit Sub
    End If

    target = CStr(idx) & " - "
    For i = 0 To Me.Controls(comboName).ListCount - 1
        txt = CStr(Me.Controls(comboName).List(i))
        If Left$(txt, Len(target)) = target Then
            Me.Controls(comboName).ListIndex = i
            Exit For
        End If
    Next i
    On Error GoTo 0
End Sub

Private Function ColLetter(ByVal n As Long) As String
    Dim x As Long
    Dim s As String

    x = n
    Do While x > 0
        s = Chr$(((x - 1) Mod 26) + 65) & s
        x = (x - 1) \ 26
    Loop
    ColLetter = s
End Function

Private Sub SetCtrlValueIfExists(ByVal ctrlName As String, ByVal v As Variant)
    On Error Resume Next
    Dim c As MSForms.Control
    Set c = Me.Controls(ctrlName)
    If Not c Is Nothing Then
        CallByName c, "Value", VbLet, v
    End If
    On Error GoTo 0
End Sub
Private Sub SetControlProp(ByVal ctrlName As String, ByVal propName As String, ByVal propValue As Variant)
    Dim c As Control
    On Error Resume Next
    Set c = Me.Controls(ctrlName)
    If Not c Is Nothing Then
        CallByName c, propName, VbLet, propValue
    End If
    On Error GoTo 0
End Sub

Private Sub SetControlPropIfExists(ByVal ctrlName As String, ByVal propName As String, ByVal propValue As Variant)
    Dim c As Control
    On Error Resume Next
    Set c = Me.Controls(ctrlName)
    If Not c Is Nothing Then CallByName c, propName, VbLet, propValue
    On Error GoTo 0
End Sub

Private Function GetControlValue(ByVal ctrlName As String, ByVal propName As String, ByVal defaultValue As Variant) As Variant
    Dim c As Control
    On Error Resume Next
    Set c = Me.Controls(ctrlName)
    If c Is Nothing Then
        GetControlValue = defaultValue
    Else
        GetControlValue = CallByName(c, propName, VbGet)
    End If
    On Error GoTo 0
End Function

Private Function ValidateCurrentSelections() As Boolean
    Dim pathN As String
    Dim pathN1 As String

    pathN = Trim$(CStr(GetControlValue("txtN_Path", "Value", "")))
    pathN1 = Trim$(CStr(GetControlValue("txtN1_Path", "Value", "")))

    If Len(pathN) > 0 Then
        If Not IsSupportedImportExt(pathN) Then
            MsgBox "Chemin N non supporte : " & pathN, vbExclamation
            Exit Function
        End If
    End If

    If Len(pathN1) > 0 Then
        If Not IsSupportedImportExt(pathN1) Then
            MsgBox "Chemin N-1 non supporte : " & pathN1, vbExclamation
            Exit Function
        End If
    End If

    ValidateCurrentSelections = True
End Function

Private Function IsSupportedImportExt(ByVal path As String) As Boolean
    Dim ext As String
    ext = LCase$(ImportExt(path))

    Select Case ext
        Case "txt", "csv", "dat", "xls", "xlsx", "xlsm"
            IsSupportedImportExt = True
    End Select
End Function

Private Function ImportExt(ByVal path As String) As String
    Dim p As Long
    path = Trim$(path)
    p = InStrRev(path, ".")
    If p <= 0 Or p = Len(path) Then Exit Function
    ImportExt = Mid$(path, p + 1)
End Function

'========================================================
' Nettoyage libelles : corrige mojibake + enleve accents
'========================================================
Private Function CleanLabel_NoAccents(ByVal s As String) As String
    Dim t As String

    CleanLabel_NoAccents = CleanLabel_NoAccents_Core(s)
    Exit Function

    t = CStr(s)

    t = Replace(t, "Ã©", "e")
    t = Replace(t, "Ã¨", "e")
    t = Replace(t, "Ãª", "e")
    t = Replace(t, "Ã«", "e")

    t = Replace(t, "Ã ", "a")
    t = Replace(t, "Ã¢", "a")
    t = Replace(t, "Ã¤", "a")

    t = Replace(t, "Ã¹", "u")
    t = Replace(t, "Ã»", "u")
    t = Replace(t, "Ã¼", "u")

    t = Replace(t, "Ã´", "o")
    t = Replace(t, "Ã¶", "o")

    t = Replace(t, "Ã§", "c")

    t = Replace(t, "Ã‰", "E")
    t = Replace(t, "Ã€", "A")
    t = Replace(t, "Ãˆ", "E")
    t = Replace(t, "ÃŠ", "E")
    t = Replace(t, "Ã‡", "C")

    t = Replace(t, "â€™", "'")
    t = Replace(t, "â€˜", "'")
    t = Replace(t, "â€œ", """")
    t = Replace(t, "â€", """")
    t = Replace(t, "â€“", "-")
    t = Replace(t, "â€”", "-")
    t = Replace(t, "Â", "")
    t = Replace(t, "Â€", "EUR")

    t = Replace(t, "ï¿½", "e")

    t = Replace(t, "é", "e")
    t = Replace(t, "è", "e")
    t = Replace(t, "ê", "e")
    t = Replace(t, "ë", "e")

    t = Replace(t, "à", "a")
    t = Replace(t, "â", "a")
    t = Replace(t, "ä", "a")

    t = Replace(t, "ù", "u")
    t = Replace(t, "û", "u")
    t = Replace(t, "ü", "u")

    t = Replace(t, "ô", "o")
    t = Replace(t, "ö", "o")

    t = Replace(t, "î", "i")
    t = Replace(t, "ï", "i")

    t = Replace(t, "ç", "c")

    t = Replace(t, "É", "E")
    t = Replace(t, "È", "E")
    t = Replace(t, "Ê", "E")
    t = Replace(t, "Ë", "E")

    t = Replace(t, "À", "A")
    t = Replace(t, "Â", "A")
    t = Replace(t, "Ä", "A")

    t = Replace(t, "Ù", "U")
    t = Replace(t, "Û", "U")
    t = Replace(t, "Ü", "U")

    t = Replace(t, "Ô", "O")
    t = Replace(t, "Ö", "O")

    t = Replace(t, "Î", "I")
    t = Replace(t, "Ï", "I")

    t = Replace(t, "Ç", "C")

    t = Replace(t, ChrW(160), " ")
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    t = Trim$(t)

    CleanLabel_NoAccents = t
End Function

Private Function CleanLabel_NoAccents_Core(ByVal s As String) As String
    Dim t As String

    t = CStr(s)

    t = Replace(t, BuildChars(&HC3, &HA9), "e")
    t = Replace(t, BuildChars(&HC3, &HA8), "e")
    t = Replace(t, BuildChars(&HC3, &HAA), "e")
    t = Replace(t, BuildChars(&HC3, &HAB), "e")

    t = Replace(t, BuildChars(&HC3, &HA0), "a")
    t = Replace(t, BuildChars(&HC3, &HA2), "a")
    t = Replace(t, BuildChars(&HC3, &HA4), "a")

    t = Replace(t, BuildChars(&HC3, &HB9), "u")
    t = Replace(t, BuildChars(&HC3, &HBB), "u")
    t = Replace(t, BuildChars(&HC3, &HBC), "u")

    t = Replace(t, BuildChars(&HC3, &HB4), "o")
    t = Replace(t, BuildChars(&HC3, &HB6), "o")

    t = Replace(t, BuildChars(&HC3, &HA7), "c")

    t = Replace(t, BuildChars(&HC3, &H89), "E")
    t = Replace(t, BuildChars(&HC3, &H80), "A")
    t = Replace(t, BuildChars(&HC3, &H88), "E")
    t = Replace(t, BuildChars(&HC3, &H8A), "E")
    t = Replace(t, BuildChars(&HC3, &H87), "C")

    t = Replace(t, BuildChars(&HE2, &H20AC, &H2019), "'")
    t = Replace(t, BuildChars(&HE2, &H20AC, &H2018), "'")
    t = Replace(t, BuildChars(&HE2, &H20AC, &H201C), """")
    t = Replace(t, BuildChars(&HE2, &H20AC, &H201D), """")
    t = Replace(t, BuildChars(&HE2, &H20AC, &H2013), "-")
    t = Replace(t, BuildChars(&HE2, &H20AC, &H2014), "-")
    t = Replace(t, ChrW(&HC2), "")
    t = Replace(t, BuildChars(&HC2, &H20AC), "EUR")

    t = Replace(t, BuildChars(&HEF, &HBF, &HBD), "e")
    t = Replace(t, ChrW(&HFFFD), "e")

    t = Replace(t, ChrW(&HE9), "e")
    t = Replace(t, ChrW(&HE8), "e")
    t = Replace(t, ChrW(&HEA), "e")
    t = Replace(t, ChrW(&HEB), "e")

    t = Replace(t, ChrW(&HE0), "a")
    t = Replace(t, ChrW(&HE2), "a")
    t = Replace(t, ChrW(&HE4), "a")

    t = Replace(t, ChrW(&HF9), "u")
    t = Replace(t, ChrW(&HFB), "u")
    t = Replace(t, ChrW(&HFC), "u")

    t = Replace(t, ChrW(&HF4), "o")
    t = Replace(t, ChrW(&HF6), "o")

    t = Replace(t, ChrW(&HEE), "i")
    t = Replace(t, ChrW(&HEF), "i")

    t = Replace(t, ChrW(&HE7), "c")

    t = Replace(t, ChrW(&HC9), "E")
    t = Replace(t, ChrW(&HC8), "E")
    t = Replace(t, ChrW(&HCA), "E")
    t = Replace(t, ChrW(&HCB), "E")

    t = Replace(t, ChrW(&HC0), "A")
    t = Replace(t, ChrW(&HC2), "A")
    t = Replace(t, ChrW(&HC4), "A")

    t = Replace(t, ChrW(&HD9), "U")
    t = Replace(t, ChrW(&HDB), "U")
    t = Replace(t, ChrW(&HDC), "U")

    t = Replace(t, ChrW(&HD4), "O")
    t = Replace(t, ChrW(&HD6), "O")

    t = Replace(t, ChrW(&HCE), "I")
    t = Replace(t, ChrW(&HCF), "I")

    t = Replace(t, ChrW(&HC7), "C")

    t = Replace(t, ChrW(160), " ")
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    t = Trim$(t)

    CleanLabel_NoAccents_Core = t
End Function

Private Function BuildChars(ParamArray codePoints() As Variant) As String
    Dim i As Long

    For i = LBound(codePoints) To UBound(codePoints)
        BuildChars = BuildChars & ChrW(CLng(codePoints(i)))
    Next i
End Function

Private Sub CleanLabelColumn2_NoAccents_InArray(ByRef arr As Variant)
    Dim i As Long

    If IsEmpty(arr) Or Not IsArray(arr) Then Exit Sub
    On Error GoTo SafeExit

    For i = LBound(arr, 1) To UBound(arr, 1)
        arr(i, 2) = CleanLabel_NoAccents(arr(i, 2))
    Next i

SafeExit:
End Sub














