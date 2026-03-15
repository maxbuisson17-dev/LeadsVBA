VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBalanceType 
   ClientHeight    =   5952
   ClientLeft      =   144
   ClientTop       =   744
   ClientWidth     =   7320
   OleObjectBlob   =   "frmBalanceType.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBalanceType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False












Option Explicit

Public SelectedMode As eBalance4ColsMode
Public SelectedColN As String

Private Sub UserForm_Initialize()
    On Error GoTo EH

    ModNovancesTheme.AppliquerTheme Me
    ModNovancesTheme.AjouterBarreTitre Me, "Format de la balance"
    Me.Caption = "Type balance 4 colonnes"

    ' Defaut: N/N-1 sur colonne C
    optBalNN1.Value = True
    optBalDebitCredit.Value = False
    optColC.Value = True
    optColD.Value = False
    ApplyColChoiceState
    UpdateSelectionFromUI

CleanExit:
    Exit Sub
EH:
    Debug.Print "frmBalanceType.UserForm_Initialize error " & Err.Number & " : " & Err.Description
    SelectedMode = b4None
    SelectedColN = ""
    Resume CleanExit
End Sub

Private Sub optBalNN1_Click()
    On Error GoTo EH

    If optBalNN1.Value Then optBalDebitCredit.Value = False
    ApplyColChoiceState
    UpdateSelectionFromUI

CleanExit:
    Exit Sub
EH:
    Debug.Print "optBalNN1_Click error " & Err.Number & " : " & Err.Description
    Resume CleanExit
End Sub

Private Sub optBalDebitCredit_Click()
    On Error GoTo EH

    If optBalDebitCredit.Value Then optBalNN1.Value = False
    ApplyColChoiceState
    UpdateSelectionFromUI

CleanExit:
    Exit Sub
EH:
    Debug.Print "optBalDebitCredit_Click error " & Err.Number & " : " & Err.Description
    Resume CleanExit
End Sub

Private Sub optColC_Click()
    On Error GoTo EH

    If optColC.Value Then optColD.Value = False
    UpdateSelectionFromUI

CleanExit:
    Exit Sub
EH:
    Debug.Print "optColC_Click error " & Err.Number & " : " & Err.Description
    Resume CleanExit
End Sub

Private Sub optColD_Click()
    On Error GoTo EH

    If optColD.Value Then optColC.Value = False
    UpdateSelectionFromUI

CleanExit:
    Exit Sub
EH:
    Debug.Print "optColD_Click error " & Err.Number & " : " & Err.Description
    Resume CleanExit
End Sub

Private Sub cmdContinue_Click()
    On Error GoTo EH

    UpdateSelectionFromUI
    Me.Hide

CleanExit:
    Exit Sub
EH:
    Debug.Print "cmdContinue_Click error " & Err.Number & " : " & Err.Description
    Resume CleanExit
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo EH

    SelectedMode = b4None
    SelectedColN = ""
    Me.Hide

CleanExit:
    Exit Sub
EH:
    Debug.Print "cmdCancel_Click error " & Err.Number & " : " & Err.Description
    Resume CleanExit
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error GoTo EH

    If CloseMode = vbFormControlMenu Then
        SelectedMode = b4None
        SelectedColN = ""
        Me.Hide
        Cancel = True
    End If

CleanExit:
    Exit Sub
EH:
    Debug.Print "UserForm_QueryClose error " & Err.Number & " : " & Err.Description
    Resume CleanExit
End Sub

Private Sub ApplyColChoiceState()
    Dim canChooseCol As Boolean

    On Error GoTo EH

    canChooseCol = CBool(optBalNN1.Value)
    SetOptionState optColC, canChooseCol
    SetOptionState optColD, canChooseCol

    If canChooseCol Then
        If Not optColC.Value And Not optColD.Value Then optColC.Value = True
    Else
        optColC.Value = False
        optColD.Value = False
    End If

CleanExit:
    Exit Sub
EH:
    Debug.Print "ApplyColChoiceState error " & Err.Number & " : " & Err.Description
    Resume CleanExit
End Sub

Private Sub UpdateSelectionFromUI()
    On Error GoTo EH

    If optBalDebitCredit.Value Then
        SelectedMode = b4DebitCredit
        SelectedColN = ""
    Else
        If optColD.Value Then
            SelectedMode = b4NN1_ColD
            SelectedColN = "D"
        Else
            SelectedMode = b4NN1
            SelectedColN = "C"
        End If
    End If

CleanExit:
    Exit Sub
EH:
    Debug.Print "UpdateSelectionFromUI error " & Err.Number & " : " & Err.Description
    SelectedMode = b4None
    SelectedColN = ""
    Resume CleanExit
End Sub

Private Sub SetOptionState(ByVal opt As MSForms.Control, ByVal enabledValue As Boolean)
    On Error GoTo EH

    CallByName opt, "Enabled", VbLet, enabledValue
    On Error Resume Next
    CallByName opt, "Locked", VbLet, Not enabledValue
    On Error GoTo EH

CleanExit:
    Exit Sub
EH:
    Debug.Print "SetOptionState error " & Err.Number & " : " & Err.Description
    Resume CleanExit
End Sub
