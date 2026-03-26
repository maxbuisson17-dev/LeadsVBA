Attribute VB_Name = "ZCopyCode"
Option Explicit

Sub Exporter_Tout_Le_Code_VBA()

    Dim chemin As String
    Dim comp As Object
    Dim fso As Object
    Dim dossierModules As String
    Dim dossierClasses As String
    Dim dossierForms As String

    ' === DOSSIER PRINCIPAL (� modifier si tu veux) ===
    chemin = "C:\Users\M.Buisson\Desktop\26 Leads project\VBA Claude"


    ' Cr�e l'objet FileSystem
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Cr�e les dossiers si absents
    If Not fso.FolderExists(chemin) Then fso.CreateFolder chemin

    dossierModules = chemin & "Modules\"
    dossierClasses = chemin & "Classes\"
    dossierForms = chemin & "Forms\"

    If Not fso.FolderExists(dossierModules) Then fso.CreateFolder dossierModules
    If Not fso.FolderExists(dossierClasses) Then fso.CreateFolder dossierClasses
    If Not fso.FolderExists(dossierForms) Then fso.CreateFolder dossierForms


    ' === EXPORT DE TOUS LES COMPOSANTS ===
    For Each comp In ThisWorkbook.VBProject.VBComponents

        Select Case comp.Type

            ' Modules standards (.bas)
            Case 1
                comp.Export dossierModules & comp.name & ".bas"

            ' Modules de classe (.cls)
            Case 2
                comp.Export dossierClasses & comp.name & ".cls"

            ' UserForms (.frm + .frx)
            Case 3
                comp.Export dossierForms & comp.name & ".frm"

            ' Feuilles / ThisWorkbook (non exportables proprement)
            Case 100
                ' On ignore volontairement

        End Select

    Next comp


    MsgBox "Export terminé !" & vbCrLf & _
           "Dossier : " & chemin, vbInformation

End Sub


