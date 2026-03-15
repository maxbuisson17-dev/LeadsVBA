Attribute VB_Name = "ModNovancesTheme"
'============================================================
' MODULE : ModNovancesTheme
' Charte graphique Novances Groupe pour UserForms VBA
' Import : VBE > Fichier > Importer un fichier > ce .bas
'============================================================

Option Explicit

' ---- PALETTE NOVANCES ----
Public Const NOVANCES_NAVY As Long = 2696220        ' RGB(28, 37, 65)  - Primaire
Public Const NOVANCES_BLUE As Long = 14375739       ' RGB(59, 91, 219) - Accent
Public Const NOVANCES_BLUE_LIGHT As Long = 15704435  ' RGB(83, 113, 239) - Accent hover
Public Const NOVANCES_WHITE As Long = 16777215       ' RGB(255, 255, 255)
Public Const NOVANCES_BG As Long = 15790320          ' RGB(240, 242, 245) - Fond
Public Const NOVANCES_CARD As Long = 16777215        ' RGB(255, 255, 255) - Cartes
Public Const NOVANCES_BORDER As Long = 15066597      ' RGB(213, 218, 229) - Bordures
Public Const NOVANCES_MUTED As Long = 7697781        ' RGB(117, 126, 141) - Texte secondaire
Public Const NOVANCES_TEXT As Long = 2696220          ' RGB(28, 37, 65)  - Texte principal

' ---- POLICES ----
Public Const FONT_TITLE As String = "Poppins"
Public Const FONT_BODY As String = "Poppins"
Public Const FONT_SIZE_TITLE As Single = 15
Public Const FONT_SIZE_LABEL As Single = 10
Public Const FONT_SIZE_INPUT As Single = 9.5
Public Const FONT_SIZE_BUTTON As Single = 9.5

'============================================================
' PROCEDURE PRINCIPALE : Appliquer le theme a un UserForm
'============================================================
Public Sub AppliquerTheme(frm As MSForms.UserForm)
    
    ' -- Style du formulaire --
    With frm
        .BackColor = NOVANCES_BG
        .Font.name = FONT_BODY
        .Font.Size = FONT_SIZE_LABEL
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = NOVANCES_BORDER
    End With
    
    ' -- Parcourir tous les controles --
    Dim ctrl As MSForms.Control
    For Each ctrl In frm.Controls
        StylerControle ctrl
    Next ctrl
    
End Sub

'============================================================
' STYLISER UN CONTROLE INDIVIDUEL
'============================================================
Private Sub StylerControle(ctrl As MSForms.Control)
    
    On Error Resume Next ' Certaines proprietes n'existent pas sur tous les controles
    
    Select Case TypeName(ctrl)
    
        Case "TextBox"
            With ctrl
                .BackColor = NOVANCES_CARD
                .BorderStyle = fmBorderStyleSingle
                .BorderColor = NOVANCES_BORDER
                .ForeColor = NOVANCES_TEXT
                .Font.name = FONT_BODY
                .Font.Size = FONT_SIZE_INPUT
                .SpecialEffect = fmSpecialEffectFlat
                .Height = 26
            End With
            
        Case "Label"
            With ctrl
                .BackStyle = fmBackStyleTransparent
                .ForeColor = NOVANCES_TEXT
                .Font.name = FONT_BODY
                .Font.Size = FONT_SIZE_LABEL
            End With
            ' Detection des titres (tag = "titre" ou hauteur > 20)
            If LCase(ctrl.Tag) = "titre" Then
                ctrl.Font.name = FONT_TITLE
                ctrl.Font.Size = FONT_SIZE_TITLE
                ctrl.ForeColor = NOVANCES_NAVY
            End If
            
        Case "CommandButton"
            With ctrl
                .Font.name = FONT_BODY
                .Font.Size = FONT_SIZE_BUTTON
                .Height = 32
            End With
            ctrl.BackStyle = fmBackStyleOpaque
            If IsCancelOrBackButton(ctrl.caption) Then
                ctrl.BackColor = RGB(192, 0, 0)
                ctrl.ForeColor = NOVANCES_WHITE
            Else
                ctrl.BackColor = NOVANCES_BLUE
                ctrl.ForeColor = NOVANCES_WHITE
            End If
            
        Case "CheckBox"
            With ctrl
                .BackStyle = fmBackStyleTransparent
                .ForeColor = NOVANCES_TEXT
                .Font.name = FONT_BODY
                .Font.Size = FONT_SIZE_LABEL
            End With
            
        Case "OptionButton"
            With ctrl
                .BackStyle = fmBackStyleTransparent
                .ForeColor = NOVANCES_TEXT
                .Font.name = FONT_BODY
                .Font.Size = FONT_SIZE_LABEL
            End With
            
        Case "ComboBox"
            With ctrl
                .BackColor = NOVANCES_CARD
                .BorderColor = NOVANCES_BORDER
                .ForeColor = NOVANCES_TEXT
                .Font.name = FONT_BODY
                .Font.Size = FONT_SIZE_INPUT
                .SpecialEffect = fmSpecialEffectFlat
                .Height = 26
            End With
            
        Case "ListBox"
            With ctrl
                .BackColor = NOVANCES_CARD
                .BorderColor = NOVANCES_BORDER
                .ForeColor = NOVANCES_TEXT
                .Font.name = FONT_BODY
                .Font.Size = FONT_SIZE_INPUT
                .SpecialEffect = fmSpecialEffectFlat
            End With
            
        Case "Frame"
            With ctrl
                .BackColor = NOVANCES_BG
                .BorderColor = NOVANCES_BORDER
                .ForeColor = NOVANCES_MUTED
                .Font.name = FONT_BODY
                .Font.Size = 8
                .SpecialEffect = fmSpecialEffectFlat
                .BorderStyle = fmBorderStyleSingle
            End With
            ' Styler les controles enfants du Frame
            Dim childCtrl As MSForms.Control
            For Each childCtrl In ctrl.Controls
                StylerControle childCtrl
            Next childCtrl
            
        Case "MultiPage"
            With ctrl
                .BackColor = NOVANCES_BG
                .ForeColor = NOVANCES_TEXT
                .Font.name = FONT_BODY
                .Font.Size = FONT_SIZE_LABEL
            End With
            
        Case "ScrollBar", "SpinButton"
            ctrl.BackColor = NOVANCES_BG
            ctrl.ForeColor = NOVANCES_BLUE
            
        Case "Image"
            ctrl.BackStyle = fmBackStyleTransparent
            ctrl.BorderColor = NOVANCES_BORDER
            ctrl.SpecialEffect = fmSpecialEffectFlat
            
    End Select
    
    On Error GoTo 0
    
End Sub

'============================================================
' BARRE DE TITRE PERSONNALISEE (bandeau navy en haut)
' Ajouter un Label nomme "lblTitleBar" en haut du form
'============================================================
Public Sub AjouterBarreTitre(frm As MSForms.UserForm, titre As String)

    Dim lbl As MSForms.label
    Dim lblText As MSForms.label
    Dim titleBarColor As Long

    ' Verifier si le label existe deja
    On Error Resume Next
    Set lbl = frm.Controls("lblTitleBar")
    Set lblText = frm.Controls("lblTitleBarText")
    On Error GoTo 0

    If lbl Is Nothing Then
        Set lbl = frm.Controls.Add("Forms.Label.1", "lblTitleBar")
    End If
    If lblText Is Nothing Then
        Set lblText = frm.Controls.Add("Forms.Label.1", "lblTitleBarText")
    End If

    titleBarColor = RGB(13, 34, 76)

    With lbl
        .Left = 0
        .Top = 0
        .Width = frm.InsideWidth
        .Height = 36
        .caption = vbNullString
        .BackColor = titleBarColor
        .BackStyle = fmBackStyleOpaque
        .ZOrder 0
    End With

    With lblText
        .Font.name = FONT_TITLE
        .Font.Size = 11
        .caption = titre
        .AutoSize = False
        .WordWrap = False
        .BackStyle = fmBackStyleTransparent
        .ForeColor = NOVANCES_WHITE

        ' IMPORTANT : on centre le TEXTE dans une zone large (bandeau - marges)
        .Left = 12
        .Width = frm.InsideWidth - 24
        If .Width < 10 Then .Width = frm.InsideWidth

        .TextAlign = fmTextAlignCenter
        .ZOrder 0
    End With

    FitTitle_NoWrap_Centered frm, lbl, lblText

End Sub
Private Function NormalizeButtonCaption(ByVal caption As String) As String
    NormalizeButtonCaption = LCase$(Trim$(caption))
End Function

Private Function IsCancelOrBackButton(ByVal caption As String) As Boolean
    Select Case NormalizeButtonCaption(caption)
        Case "annuler", "retour"
            IsCancelOrBackButton = True
    End Select
End Function

Private Sub CenterTitleLabel(frm As MSForms.UserForm, lbl As MSForms.label)
    If lbl Is Nothing Then Exit Sub

    lbl.Left = (frm.InsideWidth - lbl.Width) / 2
    If lbl.Left < 0 Then lbl.Left = 0
End Sub

Private Sub FitTitle_NoWrap_Centered(frm As MSForms.UserForm, titleBar As MSForms.label, lblText As MSForms.label)

    Const SIDE_MARGIN As Single = 12
    Const MIN_TITLE_FONT_SIZE As Single = 8

    Dim maxWidth As Single
    Dim measure As MSForms.label

    If frm Is Nothing Then Exit Sub
    If titleBar Is Nothing Then Exit Sub
    If lblText Is Nothing Then Exit Sub

    maxWidth = frm.InsideWidth - (SIDE_MARGIN * 2)
    If maxWidth <= 20 Then maxWidth = frm.InsideWidth

    ' Zone d'affichage : pleine largeur utile + texte centré
    lblText.Left = SIDE_MARGIN
    lblText.Width = maxWidth
    lblText.WordWrap = False
    lblText.TextAlign = fmTextAlignCenter

    ' Label temporaire pour mesurer la largeur réelle du texte (1 ligne)
    Set measure = frm.Controls.Add("Forms.Label.1", "lblTitleBarMeasure")
    With measure
        .Visible = False
        .AutoSize = True
        .WordWrap = False
        .caption = lblText.caption
        .Font.name = lblText.Font.name
        .Font.Size = lblText.Font.Size
    End With

    Do While measure.Width > maxWidth And lblText.Font.Size > MIN_TITLE_FONT_SIZE
        lblText.Font.Size = lblText.Font.Size - 0.5
        measure.Font.Size = lblText.Font.Size
        DoEvents
    Loop

    ' Centrage vertical dans le bandeau
    lblText.Top = titleBar.Top + ((titleBar.Height - lblText.Height) / 2)

    ' Nettoyage
    On Error Resume Next
    frm.Controls.Remove measure.name
    On Error GoTo 0

End Sub
'============================================================
' SEPARATEUR HORIZONTAL (ligne accent bleue)
'============================================================
Public Sub AjouterSeparateur(frm As MSForms.UserForm, topPos As Single, Optional largeur As Single = 0)
    
    Dim lbl As MSForms.label
    Set lbl = frm.Controls.Add("Forms.Label.1", "lblSep_" & Format(topPos, "000"))
    
    With lbl
        .caption = ""
        .Left = 12
        .Top = topPos
        If largeur = 0 Then largeur = frm.InsideWidth - 24
        .Width = largeur
        .Height = 2
        .BackColor = NOVANCES_BLUE
        .BackStyle = fmBackStyleOpaque
    End With
    
End Sub

'============================================================
' USAGE RAPIDE - Copier dans UserForm_Initialize :
'
'   Private Sub UserForm_Initialize()
'       ModNovancesTheme.AppliquerTheme Me
'       ModNovancesTheme.AjouterBarreTitre Me, "Mon Formulaire"
'   End Sub
'
' TAGS pour les boutons :
'   - (rien)      -> bouton bleu primaire
'   - "outline"   -> bouton bordure
'   - "ghost"     -> bouton transparent
'
' TAGS pour les labels :
'   - "titre"     -> style titre (Poppins, 14pt, navy)
'============================================================
