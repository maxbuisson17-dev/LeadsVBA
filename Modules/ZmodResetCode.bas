Attribute VB_Name = "ZmodResetCode"
Option Explicit

Public Sub Overwrite_All_VBA()
    ResetVBA_OverwriteAll_FromFolder
End Sub

Public Sub ResetVBA_OverwriteAll_FromFolder()
    Const ROOT As String = "C:\Users\M.Buisson\Desktop\26 Leads project\VBA Claude\"
    Dim keepModule As String
    Dim removed As Long
    Dim importedModules As Long
    Dim importedForms As Long
    Dim skipped As Long

    On Error GoTo EH

    keepModule = Reset_GetExecutingModuleName("ResetCode")
    Debug.Print "[ResetCode] ROOT=" & ROOT
    Debug.Print "[ResetCode] Running module kept during execution: " & keepModule

    removed = Reset_RemoveAllComponents(ThisWorkbook, keepModule)
    importedModules = Reset_ImportByKindRecursive(ThisWorkbook, ROOT, "module", keepModule, skipped)
    importedForms = Reset_ImportByKindRecursive(ThisWorkbook, ROOT, "form", keepModule, skipped)

    MsgBox "Reset termine." & vbCrLf & _
           "Supprimes: " & CStr(removed) & vbCrLf & _
           "Modules importes: " & CStr(importedModules) & vbCrLf & _
           "Forms importes: " & CStr(importedForms) & vbCrLf & _
           "Ignores: " & CStr(skipped), vbInformation
    GoTo CleanExit

EH:
    Debug.Print "ResetVBA_OverwriteAll_FromFolder error " & Err.Number & " : " & Err.Description
    MsgBox "Erreur reset: " & Err.Description, vbExclamation
    Resume CleanExit
CleanExit:
    Exit Sub
End Sub

Private Function Reset_RemoveAllComponents(ByVal wb As Workbook, ByVal keepName As String) As Long
    Dim vbComps As Object
    Dim comp As Object
    Dim i As Long

    Set vbComps = wb.VBProject.VBComponents

    For i = vbComps.Count To 1 Step -1
        Set comp = vbComps.item(i)
        If comp.Type = 1 Or comp.Type = 2 Or comp.Type = 3 Then
            If StrComp(comp.name, keepName, vbTextCompare) <> 0 Then
                On Error Resume Next
                vbComps.Remove comp
                If Err.Number = 0 Then
                    Reset_RemoveAllComponents = Reset_RemoveAllComponents + 1
                    Debug.Print "[ResetCode] Removed: " & comp.name
                Else
                    Debug.Print "[ResetCode] Remove failed: " & comp.name & " | " & Err.Number & " - " & Err.Description
                    Err.Clear
                End If
                On Error GoTo 0
            Else
                Debug.Print "[ResetCode] Kept running module: " & keepName
            End If
        End If
    Next i
End Function

Private Function Reset_ImportByKindRecursive( _
    ByVal wb As Workbook, _
    ByVal rootFolder As String, _
    ByVal kind As String, _
    ByVal keepName As String, _
    ByRef skippedCount As Long) As Long

    Dim vbComps As Object
    Dim files As Collection
    Dim fullPath As Variant
    Dim ext As String
    Dim baseName As String

    If Right$(rootFolder, 1) <> "\" Then rootFolder = rootFolder & "\"
    If Dir$(rootFolder, vbDirectory) = vbNullString Then
        Err.Raise vbObjectError + 4001, "ResetCode", "Dossier introuvable: " & rootFolder
    End If

    Set vbComps = wb.VBProject.VBComponents
    Set files = Reset_EnumFilesRecursive_FSO(rootFolder)

    For Each fullPath In files
        ext = LCase$(Mid$(CStr(fullPath), InStrRev(CStr(fullPath), ".") + 1))

        If Reset_FileMatchesKind(ext, kind) Then
            baseName = Reset_FileBaseName(CStr(fullPath))

            If StrComp(baseName, keepName, vbTextCompare) = 0 Then
                skippedCount = skippedCount + 1
                Debug.Print "[ResetCode] Skipped running module import: " & CStr(fullPath)

            ElseIf ext = "bas" And Reset_FileContainsFrmHeader(CStr(fullPath)) Then
                skippedCount = skippedCount + 1
                Debug.Print "[ResetCode] Skipped invalid .bas with .frm header: " & CStr(fullPath)

            Else
                Reset_ImportOneFile vbComps, CStr(fullPath)

                ' Important : après import d'un .frm, vérifier que c'est bien un MSForm
                If ext = "frm" Then
                    AssertIsUserForm wb, baseName
                End If

                Reset_ImportByKindRecursive = Reset_ImportByKindRecursive + 1
            End If
        End If
    Next fullPath
End Function

Private Function Reset_EnumFilesRecursive_FSO(ByVal rootFolder As String) As Collection
    Dim fso As Object
    Dim folder As Object
    Dim subFolder As Object
    Dim fil As Object
    Dim col As New Collection

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(rootFolder)

    ' fichiers du dossier courant
    For Each fil In folder.files
        col.Add fil.path
    Next fil

    ' récursif sur sous-dossiers
    For Each subFolder In folder.SubFolders
        Dim subCol As Collection
        Dim v As Variant

        Set subCol = Reset_EnumFilesRecursive_FSO(subFolder.path)
        For Each v In subCol
            col.Add v
        Next v
    Next subFolder

    Set Reset_EnumFilesRecursive_FSO = col
End Function

Private Function Reset_FileMatchesKind(ByVal ext As String, ByVal kind As String) As Boolean
    Select Case LCase$(kind)
        Case "module"
            Reset_FileMatchesKind = (ext = "bas" Or ext = "cls")
        Case "form"
            Reset_FileMatchesKind = (ext = "frm")
    End Select
End Function

Private Function Reset_FileContainsFrmHeader(ByVal fullPath As String) As Boolean
    Dim ff As Integer
    Dim txt As String

    On Error GoTo EH
    ff = FreeFile
    Open fullPath For Input As #ff
    txt = Input$(LOF(ff), ff)
    Close #ff

    txt = LCase$(txt)
    If InStr(1, txt, "version 5.00", vbBinaryCompare) > 0 Then Reset_FileContainsFrmHeader = True
    If InStr(1, txt, "begin {", vbBinaryCompare) > 0 Then Reset_FileContainsFrmHeader = True
    If InStr(1, txt, "oleobjectblob", vbBinaryCompare) > 0 Then Reset_FileContainsFrmHeader = True
    GoTo CleanExit
EH:
    Debug.Print "Reset_FileContainsFrmHeader error " & Err.Number & " : " & Err.Description
    On Error Resume Next
    If ff <> 0 Then Close #ff
    On Error GoTo 0
    Reset_FileContainsFrmHeader = False
    Resume CleanExit
CleanExit:
    Exit Function
End Function

Private Sub Reset_ImportOneFile(ByVal vbComps As Object, ByVal fullPath As String)
    Dim baseName As String
    Dim existing As Object

    baseName = Reset_FileBaseName(fullPath)

    On Error Resume Next
    Set existing = vbComps.item(baseName)
    On Error GoTo 0

    If Not existing Is Nothing Then
        If existing.Type = 1 Or existing.Type = 2 Or existing.Type = 3 Then
            On Error Resume Next
            vbComps.Remove existing
            If Err.Number <> 0 Then
                Debug.Print "[ResetCode] Remove existing failed: " & baseName & " | " & Err.Number & " - " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
        End If
    End If
    
    
    If LCase$(Right$(fullPath, 4)) = ".frm" Then
    Dim frxPath As String
    frxPath = Left$(fullPath, Len(fullPath) - 4) & ".frx"
    If Dir$(frxPath) = vbNullString Then
        Debug.Print "[ResetCode] WARNING: FRX manquant pour " & fullPath
        ' tu peux choisir : Skip au lieu d'importer
        ' Exit Sub
    End If
    End If

    
    
    On Error GoTo ImportErr
    vbComps.Import fullPath
    Debug.Print "[ResetCode] Imported: " & fullPath
    Exit Sub
ImportErr:
    Debug.Print "[ResetCode] Import failed: " & fullPath & " | " & Err.Number & " - " & Err.Description
    Err.Clear
End Sub

Private Function Reset_FileBaseName(ByVal fullPath As String) As String
    Dim f As String
    Dim p As Long

    f = Mid$(fullPath, InStrRev(fullPath, "\") + 1)
    p = InStrRev(f, ".")
    If p > 1 Then
        Reset_FileBaseName = Left$(f, p - 1)
    Else
        Reset_FileBaseName = f
    End If
End Function

Private Function Reset_GetExecutingModuleName(ByVal fallbackName As String) As String
    On Error GoTo EH
    Reset_GetExecutingModuleName = Application.VBE.ActiveCodePane.CodeModule.Parent.name
    GoTo CleanExit
EH:
    Debug.Print "Reset_GetExecutingModuleName error " & Err.Number & " : " & Err.Description
    Reset_GetExecutingModuleName = fallbackName
    Resume CleanExit
CleanExit:
    Exit Function
End Function

Private Sub AssertIsUserForm(ByVal wb As Workbook, ByVal formName As String)
    Dim comp As Object
    Set comp = wb.VBProject.VBComponents(formName)

    If comp.Type <> 3 Then
        Err.Raise vbObjectError + 513, , formName & " n'est pas un UserForm (Type=" & comp.Type & ")."
    End If
End Sub
