VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBGControl 
   ClientHeight    =   9336.001
   ClientLeft      =   204
   ClientTop       =   852
   ClientWidth     =   11784
   OleObjectBlob   =   "frmBGControl.frx":0000
End
Attribute VB_Name = "frmBGControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






'========================
' frmBGControl - V4 (CDC)
'========================
Option Explicit
Private mInit As Boolean

Private Sub Label1_Click()

End Sub

Private Sub UserForm_Initialize()
    mInit = True
        
    ModNovancesTheme.AppliquerTheme Me
    ModNovancesTheme.AjouterBarreTitre Me, "Balance import�e"
    
        
    SetupPreviewList
    DisableHeaderControls

    SetCtrlValueIfExists "chkIs4Cols", True
    gSkipRows = 0

    FillColumnCombos 30
    SetCtrlValueIfExists "cboColCompte", "A"
    SetCtrlValueIfExists "cboColLibelle", "B"
    SetCtrlValueIfExists "cboColSoldeN", "C"
    SetCtrlValueIfExists "cboColSoldeN1", "D"

    ToggleColumnMappingUI

    RefreshPreviewFromArray gPreviewData
    mInit = False
End Sub

Private Sub SetupPreviewList()
    With Me.lstPreview
        .Clear
        .ColumnCount = 4
        On Error Resume Next
        .ColumnWidths = "80 pt;220 pt;90 pt;90 pt"
        On Error GoTo 0
    End With
End Sub

Private Sub DisableHeaderControls()
    ' Parcours OUI: suppression du param�trage manuel des en-t�tes.
    SetControlPropIfExists "fraHeaders", "Visible", False
    SetControlPropIfExists "fraHeaders", "Enabled", False
    SetControlPropIfExists "optHeaderNo", "Visible", False
    SetControlPropIfExists "optHeaderNo", "Enabled", False
    SetControlPropIfExists "optHeaderYes", "Visible", False
    SetControlPropIfExists "optHeaderYes", "Enabled", False
End Sub

Private Sub RefreshPreviewFromArray(ByVal arr As Variant)
    Dim i As Long
    Dim rows As Long

    SetupPreviewList

    If IsEmpty(arr) Or Not IsArray(arr) Then Exit Sub

    rows = UBound(arr, 1)
    If rows > 200 Then rows = 200

    For i = 1 To rows
        Me.lstPreview.AddItem CStr(arr(i, 1))
        If UBound(arr, 2) >= 2 Then Me.lstPreview.List(Me.lstPreview.ListCount - 1, 1) = CStr(arr(i, 2))
        If UBound(arr, 2) >= 3 Then Me.lstPreview.List(Me.lstPreview.ListCount - 1, 2) = CStr(arr(i, 3))
        If UBound(arr, 2) >= 4 Then Me.lstPreview.List(Me.lstPreview.ListCount - 1, 3) = CStr(arr(i, 4))
    Next i
End Sub

Private Sub SetControlPropIfExists(ByVal ctrlName As String, ByVal propName As String, ByVal propValue As Variant)
    Dim c As Control
    On Error Resume Next
    Set c = Me.Controls(ctrlName)
    If Not c Is Nothing Then CallByName c, propName, VbLet, propValue
    Set c = Nothing
    On Error GoTo 0
End Sub

Private Sub SetCtrlValueIfExists(ByVal ctrlName As String, ByVal v As Variant)
    On Error Resume Next
    Dim c As MSForms.Control
    Set c = Me.Controls(ctrlName)
    If Not c Is Nothing Then
        CallByName c, "Value", VbLet, v
    End If
    On Error GoTo 0
End Sub

Private Function GetCtrlValueIfExists(ByVal ctrlName As String, ByVal defaultValue As Variant) As Variant
    On Error Resume Next
    Dim c As MSForms.Control
    Set c = Me.Controls(ctrlName)
    If c Is Nothing Then
        GetCtrlValueIfExists = defaultValue
    Else
        GetCtrlValueIfExists = CallByName(c, "Value", VbGet)
    End If
    On Error GoTo 0
End Function
Private Sub cmdChooseOther_Click()
    If Not modOrchestrator.PickAndLoadPreview() Then Exit Sub
    RefreshPreviewFromArray gPreviewData
End Sub

Private Sub cmdModify_Click()
    If Len(gBalancePath) > 0 Then Workbooks.Open gBalancePath
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub lstPreview_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
End Sub

Private Sub cmdContinue_Click()
    On Error GoTo EH

    Dim colCompte As String
    Dim colLib As String
    Dim colN As String
    Dim colN1 As String

    gSkipRows = 0

    If CBool(GetCtrlValueIfExists("chkIs4Cols", False)) Then
        colCompte = "A": colLib = "B": colN = "C": colN1 = "D"
    Else
        colCompte = CStr(GetCtrlValueIfExists("cboColCompte", ""))
        colLib = CStr(GetCtrlValueIfExists("cboColLibelle", ""))
        colN = CStr(GetCtrlValueIfExists("cboColSoldeN", ""))
        colN1 = CStr(GetCtrlValueIfExists("cboColSoldeN1", ""))

        If colCompte = "" Or colLib = "" Or colN = "" Or colN1 = "" Then
            MsgBox "Renseigne la correspondance des colonnes (Compte/Libell�/N/N-1).", vbExclamation
            GoTo CleanExit
        End If
    End If

    If Not modOrchestrator.LoadFullDataCustomCols(0, colCompte, colLib, colN, colN1) Then GoTo CleanExit
    If IsEmpty(gFullData) Or Not IsArray(gFullData) Then
        MsgBox "Erreur : gFullData n'a pas �t� charg�.", vbCritical
        GoTo CleanExit
    End If

    gFullData = modLeadsWizard.Ensure4Cols(gFullData)

    modExportEngine.EnsureWorkingSheetsHidden False
    modDA_BG.ImportIntoBG_FromFullData
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
    MsgBox "Erreur cmdContinue_Click : " & Err.Number & vbCrLf & Err.Description, vbCritical
    Resume CleanExit
CleanExit:
    Exit Sub
End Sub

Private Sub cmdProcess_Click()
    cmdContinue_Click
End Sub

Private Sub FillColumnCombos(ByVal maxCols As Long)
    Dim i As Long
    Dim colName As String

    Me.cboColCompte.Clear
    Me.cboColLibelle.Clear
    Me.cboColSoldeN.Clear
    Me.cboColSoldeN1.Clear

    For i = 1 To maxCols
        colName = ColumnLetter(i)
        Me.cboColCompte.AddItem colName
        Me.cboColLibelle.AddItem colName
        Me.cboColSoldeN.AddItem colName
        Me.cboColSoldeN1.AddItem colName
    Next i
End Sub

Private Function ColumnLetter(ByVal colNum As Long) As String
    ColumnLetter = Split(Cells(1, colNum).Address(True, False), "$")(0)
End Function

Private Sub ToggleColumnMappingUI()
    SetControlPropIfExists "fraColMap", "Enabled", (Not CBool(GetCtrlValueIfExists("chkIs4Cols", False)))
    SetControlPropIfExists "cboColCompte", "Enabled", (Not CBool(GetCtrlValueIfExists("chkIs4Cols", False)))
    SetControlPropIfExists "cboColLibelle", "Enabled", (Not CBool(GetCtrlValueIfExists("chkIs4Cols", False)))
    SetControlPropIfExists "cboColSoldeN", "Enabled", (Not CBool(GetCtrlValueIfExists("chkIs4Cols", False)))
    SetControlPropIfExists "cboColSoldeN1", "Enabled", (Not CBool(GetCtrlValueIfExists("chkIs4Cols", False)))
End Sub

Private Sub chkIs4Cols_Click()
    ToggleColumnMappingUI
End Sub




