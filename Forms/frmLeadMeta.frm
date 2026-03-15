VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLeadMeta 
   ClientHeight    =   8076
   ClientLeft      =   132
   ClientTop       =   552
   ClientWidth     =   10284
   OleObjectBlob   =   "frmLeadMeta.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLeadMeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

'CHANGE: aucun controle cree dynamiquement, uniquement controles designer
Private mIsApplyingExportMode As Boolean
Private mSubmitInProgress As Boolean
Private mWarnedMissingKE As Boolean
Private mWarnedMissingExportControls As Boolean

Private Sub CheckBox1_Click()
End Sub

Private Sub UserForm_Initialize()
    'CHANGE: initialisation basee sur controles existants uniquement
    On Error GoTo EH

    ModNovancesTheme.AppliquerTheme Me
    ModNovancesTheme.AjouterBarreTitre Me, "Param�tres"

    SetCtrlCaptionIfExists "", "Informations dossier"

    If ControlExists("txtClient") Then SetCtrlValueIfExists "txtClient", gClient
    If ControlExists("txtExercice") Then SetCtrlValueIfExists "txtExercice", gExercice
    If ControlExists("txtVersion") Then SetCtrlValueIfExists "txtVersion", gVersion

    If ControlExists("chkKE") Then
        SetCtrlValueIfExists "chkKE", CBool(gGenerateInKE)
    Else
        WarnMissingChkKE
    End If

    If ExportModeControlsAvailable() Then
        EnsureExportModeDefault
    End If

CleanExit:
    Exit Sub
EH:
    Debug.Print "UserForm_Initialize error " & Err.Number & " : " & Err.Description
    Resume CleanExit
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim dExo As Date
    Dim vClient As String
    Dim vExercice As String
    Dim vVersion As String

    On Error GoTo EH

    If mSubmitInProgress Then GoTo CleanExit
    mSubmitInProgress = True
    SetCtrlEnabledIfExists "cmdOK", False

    If Not TryReadRequiredText("txtClient", "Client obligatoire.", vClient) Then GoTo CleanExit
    If Not TryReadRequiredText("txtExercice", "Exercice obligatoire (ex: 30/09/2025).", vExercice) Then GoTo CleanExit
    If Not TryReadRequiredText("txtVersion", "Version obligatoire (ex: 1).", vVersion) Then GoTo CleanExit

    If Not IsValidExerciceDate_UI(vExercice, dExo) Then
        MsgBox "Format d'exercice invalide.", vbExclamation
        SelectAllIfExists "txtExercice"
        GoTo CleanExit
    End If

    gClient = vClient
    gExercice = vExercice
    gVersion = vVersion

    'CHANGE: choix K? pilote uniquement via chkKE du designer
    If ControlExists("chkKE") Then
        gGenerateInKE = CBool(GetCtrlValueIfExists("chkKE", False))
    Else
        WarnMissingChkKE
    End If

    'CHANGE: mode export pilote par checkboxes designer, fallback emFS
    If ExportModeControlsAvailable() Then
        gExportMode = GetSelectedExportMode()
    Else
        WarnMissingExportControls
        gExportMode = emFS
    End If

    Unload Me
    modOrchestrator.RunGenerateLeads_V4
    GoTo CleanExit

EH:
    Debug.Print "cmdOK_Click error " & Err.Number & " : " & Err.Description
    Resume CleanExit
CleanExit:
    On Error Resume Next
    If Me.Visible Then SetCtrlEnabledIfExists "cmdOK", True
    On Error GoTo 0
    mSubmitInProgress = False
End Sub

Private Sub ChkAll_Click()
    ApplyExclusiveExportMode "ChkAll"
End Sub

Private Sub ChkFS_Click()
    ApplyExclusiveExportMode "ChkFS"
End Sub

Private Sub ChkLeads_Click()
    ApplyExclusiveExportMode "ChkLeads"
End Sub

Private Sub EnsureExportModeDefault()
    Dim hasAll As Boolean, hasFS As Boolean, hasLeads As Boolean

    If Not ExportModeControlsAvailable() Then Exit Sub

    hasAll = CBool(GetCtrlValueIfExists("ChkAll", False))
    hasFS = CBool(GetCtrlValueIfExists("ChkFS", False))
    hasLeads = CBool(GetCtrlValueIfExists("ChkLeads", False))

    If Not (hasAll Or hasFS Or hasLeads) Then
        ApplyExclusiveExportMode "ChkFS"
        Exit Sub
    End If

    If hasAll Then
        ApplyExclusiveExportMode "ChkAll"
    ElseIf hasFS Then
        ApplyExclusiveExportMode "ChkFS"
    Else
        ApplyExclusiveExportMode "ChkLeads"
    End If
End Sub

Private Sub ApplyExclusiveExportMode(ByVal activeName As String)
    If mIsApplyingExportMode Then Exit Sub
    If Not ExportModeControlsAvailable() Then Exit Sub

    On Error GoTo CleanExit
    mIsApplyingExportMode = True

    SetCtrlValueIfExists "ChkAll", (UCase$(activeName) = "CHKALL")
    SetCtrlValueIfExists "ChkFS", (UCase$(activeName) = "CHKFS")
    SetCtrlValueIfExists "ChkLeads", (UCase$(activeName) = "CHKLEADS")

    If Not CBool(GetCtrlValueIfExists("ChkAll", False)) _
       And Not CBool(GetCtrlValueIfExists("ChkFS", False)) _
       And Not CBool(GetCtrlValueIfExists("ChkLeads", False)) Then
        SetCtrlValueIfExists "ChkFS", True
    End If

CleanExit:
    mIsApplyingExportMode = False
End Sub

Private Function GetSelectedExportMode() As eExportMode
    If Not ExportModeControlsAvailable() Then
        GetSelectedExportMode = emFS
        Exit Function
    End If

    EnsureExportModeDefault

    If CBool(GetCtrlValueIfExists("ChkAll", False)) Then
        GetSelectedExportMode = emAll
    ElseIf CBool(GetCtrlValueIfExists("ChkLeads", False)) Then
        GetSelectedExportMode = emLeads
    Else
        GetSelectedExportMode = emFS
    End If
End Function

Private Function ExportModeControlsAvailable() As Boolean
    ExportModeControlsAvailable = ControlExists("ChkAll") And ControlExists("ChkFS") And ControlExists("ChkLeads")
End Function

Private Function TryReadRequiredText(ByVal ctrlName As String, ByVal emptyMessage As String, ByRef outValue As String) As Boolean
    If Not ControlExists(ctrlName) Then
        MsgBox "Le controle " & ctrlName & " est manquant dans frmLeadMeta.", vbCritical
        Exit Function
    End If

    outValue = Trim$(CStr(GetCtrlValueIfExists(ctrlName, "")))
    If Len(outValue) = 0 Then
        MsgBox emptyMessage, vbExclamation
        SetFocusIfExists ctrlName
        Exit Function
    End If

    TryReadRequiredText = True
End Function

Private Sub WarnMissingChkKE()
    If mWarnedMissingKE Then Exit Sub
    mWarnedMissingKE = True
    MsgBox "Le controle chkKE est manquant dans frmLeadMeta.", vbExclamation
End Sub

Private Sub WarnMissingExportControls()
    If mWarnedMissingExportControls Then Exit Sub
    mWarnedMissingExportControls = True
    MsgBox "Les controles ChkAll/ChkFS/ChkLeads sont manquants dans frmLeadMeta. Le mode FS sera utilise par defaut.", vbExclamation
End Sub

Private Function ResolveControl(ByVal ctrlName As String) As MSForms.Control
    Dim c As MSForms.Control
    Dim fr As MSForms.Frame

    On Error Resume Next
    Set c = Me.Controls(ctrlName)
    On Error GoTo 0

    If c Is Nothing Then
        On Error Resume Next
        Set fr = Me.Controls("Importemod")
        If fr Is Nothing Then Set fr = Me.Controls("Importmod")
        On Error GoTo 0

        If Not fr Is Nothing Then
            On Error Resume Next
            Set c = fr.Controls(ctrlName)
            On Error GoTo 0
        End If
    End If

    Set ResolveControl = c
End Function

Private Function ControlExists(ByVal ctrlName As String) As Boolean
    Dim c As MSForms.Control
    Set c = ResolveControl(ctrlName)
    ControlExists = Not c Is Nothing
End Function

Private Sub SetCtrlValueIfExists(ByVal ctrlName As String, ByVal v As Variant)
    On Error Resume Next
    Dim c As MSForms.Control
    Set c = ResolveControl(ctrlName)
    If Not c Is Nothing Then CallByName c, "Value", VbLet, v
    On Error GoTo 0
End Sub

Private Function GetCtrlValueIfExists(ByVal ctrlName As String, ByVal defaultValue As Variant) As Variant
    On Error Resume Next
    Dim c As MSForms.Control
    Set c = ResolveControl(ctrlName)
    If c Is Nothing Then
        GetCtrlValueIfExists = defaultValue
    Else
        GetCtrlValueIfExists = CallByName(c, "Value", VbGet)
    End If
    On Error GoTo 0
End Function

Private Sub SetCtrlCaptionIfExists(ByVal ctrlName As String, ByVal v As String)
    On Error Resume Next
    If Len(ctrlName) = 0 Then
        Me.caption = v
    Else
        Dim c As MSForms.Control
        Set c = ResolveControl(ctrlName)
        If Not c Is Nothing Then CallByName c, "Caption", VbLet, v
    End If
    On Error GoTo 0
End Sub

Private Sub SetFocusIfExists(ByVal ctrlName As String)
    On Error Resume Next
    Dim c As MSForms.Control
    Set c = ResolveControl(ctrlName)
    If Not c Is Nothing Then CallByName c, "SetFocus", VbMethod
    On Error GoTo 0
End Sub

Private Sub SelectAllIfExists(ByVal ctrlName As String)
    On Error Resume Next
    Dim c As MSForms.Control
    Set c = ResolveControl(ctrlName)
    If Not c Is Nothing Then
        CallByName c, "SelStart", VbLet, 0
        CallByName c, "SelLength", VbLet, Len(CStr(CallByName(c, "Text", VbGet)))
    End If
    On Error GoTo 0
End Sub

Private Sub SetCtrlEnabledIfExists(ByVal ctrlName As String, ByVal v As Boolean)
    On Error Resume Next
    Dim c As MSForms.Control
    Set c = ResolveControl(ctrlName)
    If Not c Is Nothing Then CallByName c, "Enabled", VbLet, v
    On Error GoTo 0
End Sub








