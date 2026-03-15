VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBGError 
   ClientHeight    =   8436.001
   ClientLeft      =   144
   ClientTop       =   744
   ClientWidth     =   11784
   OleObjectBlob   =   "frmBGError.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBGError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit

Private Sub CmdContinuer_Click()
    ' M�me comportement que CmdGenerate
    cmdGenerate_Click
End Sub

Private Sub CmdRetour_Click()
    ' Retour vers l'�cran pr�c�dent selon l�entrypoint
    Unload Me

    If gEntryPointWasYes Then
        Dim ufCtrl As New frmBGControl
        ufCtrl.Show vbModal
    Else
        Dim ufImp As New frmImportBalanceV6
        ufImp.Show vbModal
        Unload ufImp
    End If
End Sub

Private Sub lstReport_Click()
End Sub

Private Sub RappErrotxt_Click()
End Sub

Private Sub UserForm_Initialize()

    ModNovancesTheme.AppliquerTheme Me
    ModNovancesTheme.AjouterBarreTitre Me, "Rapport d'erreurs"

    Dim i As Long
    Dim testName As String
    Dim isOk As Boolean
    Dim isCritical As Boolean

    On Error GoTo EH

    ClearListIfExists "lstReport"
    SetListPropIfExists "lstReport", "ColumnCount", 3
    SetListPropIfExists "lstReport", "ColumnWidths", "360 pt;45 pt;45 pt"

    If gControlRows Is Nothing Then
        AddListRow "lstReport", "Aucun controle disponible", "X", "!"
    Else
        For i = 1 To gControlRows.Count
            testName = CStr(gControlRows(i)(0))
            isOk = CBool(gControlRows(i)(1))
            isCritical = CBool(gControlRows(i)(2))
            AddListRow "lstReport", testName, IIf(isOk, "V", "X"), IIf(isCritical And Not isOk, "!", "")
        Next i
    End If

    SetCtrlEnabledIfExists "cmdGenerate", gOkToGenerate
    ' Optionnel : si tu veux aussi activer/d�sactiver CmdContinuer pareil
    SetCtrlEnabledIfExists "CmdContinuer", gOkToGenerate

    GoTo CleanExit

EH:
    Debug.Print "frmBGError.UserForm_Initialize error " & Err.Number & " : " & Err.Description
    Resume CleanExit

CleanExit:
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGenerate_Click()
    If Not gOkToGenerate Then
        MsgBox "Des erreurs bloquantes empechent la generation.", vbCritical
        Exit Sub
    End If
    Unload Me
    ' RunAfterBGError : EnsureWorkingSheetsHidden + ImportIntoBG + frmLeadMeta
    ' Meme comportement pour vbYes et vbNo.
    modOrchestrator.RunAfterBGError
End Sub

' ========= Helpers (inchang�s) =========

Private Sub SetCtrlEnabledIfExists(ByVal ctrlName As String, ByVal v As Boolean)
    On Error Resume Next
    Dim c As MSForms.Control
    Set c = Me.Controls(ctrlName)
    If Not c Is Nothing Then CallByName c, "Enabled", VbLet, v
    On Error GoTo 0
End Sub

Private Sub SetListPropIfExists(ByVal ctrlName As String, ByVal propName As String, ByVal propValue As Variant)
    On Error Resume Next
    Dim c As MSForms.Control
    Set c = Me.Controls(ctrlName)
    If Not c Is Nothing Then CallByName c, propName, VbLet, propValue
    On Error GoTo 0
End Sub

Private Sub ClearListIfExists(ByVal ctrlName As String)
    On Error Resume Next
    Dim c As MSForms.Control
    Set c = Me.Controls(ctrlName)
    If Not c Is Nothing Then CallByName c, "Clear", VbMethod
    On Error GoTo 0
End Sub

Private Sub AddListRow(ByVal ctrlName As String, ByVal col1 As String, ByVal col2 As String, ByVal col3 As String)
    On Error Resume Next
    Dim lb As Object
    Set lb = Me.Controls(ctrlName)
    If Not lb Is Nothing Then
        lb.AddItem col1
        lb.List(lb.ListCount - 1, 1) = col2
        lb.List(lb.ListCount - 1, 2) = col3
    End If
    On Error GoTo 0
End Sub

