Attribute VB_Name = "Updater"
Option Explicit
Private Const UpdatePath As String = "L:\CCS\Documents\Billing\Billing Process Help\Add-Ins\CCS Excel Suite"
Private ceFolderDirectory As String
Public Sub Update()
    
    If MsgBox(Prompt:="Update CCS Excel Suite? If you clicked this button on accident, please click 'No'.", Title:="CCS Excel Suite: Update", Buttons:=vbYesNo) = vbNo Then Exit Sub
    
    Dim FilesCurrentlyInProject As Collection
    Set FilesCurrentlyInProject = GetFilesInProject
    
    'DeleteAllFilesInProject
    
    ceFolderDirectory = "C:\Users\" & Application.UserName & "\AppData\Roaming\CCS Excel Suite"
    
    Dim FileSysObject As Scripting.FileSystemObject
    Set FileSysObject = New Scripting.FileSystemObject
    
    If FileSysObject.GetFolder(UpdatePath).Files.Count = 0 Then
        MsgBox Prompt:="There are no files to import from " & UpdatePath & ".", Title:="CCS Excel Suite: Update", Buttons:=vbOKOnly
        Exit Sub
    End If
End Sub
Private Function GetFilesInProject() As Collection
    Set GetFilesInProject = New Collection
    Dim ProjectComponent As VBComponent
    
    For Each ProjectComponent In Application.ThisWorkbook.VBProject.VBComponents
        GetFilesInProject.Add ProjectComponent.Name
    Next ProjectComponent
End Function
Private Sub DeleteAllFilesInProject()
    Dim I As Long
    I = 1
    
    With Application.ThisWorkbook.VBProject.VBComponents
        While .Count > 3
            If Not .Item(I).Name = "Updater" And Not .Item(I).Name = "ThisWorkbook" And Not .Item(I).Name = "Sheet1" Then
                .Remove .Item(I)
            Else
                I = I + 1
            End If
        Wend
    End With
End Sub
Private Sub ImportProjectFiles()
    Dim objFile As File
    Dim FileName As String
    Dim FileExtension As String
    
    For Each objFile In FileSysObject.GetFolder(UpdatePath).Files
        FileName = objFile.Name
        FileExtension = Split(FileName, ".", -1, vbBinaryCompare)(1)
        FileName = Split(FileName, ".", -1, vbBinaryCompare)(0)
        If FileName = "Updater" Then FileName = vbNullString
        
        If FileExtension = "cls" Or FileExtension = "frm" Or FileExtension = "bas" Then
            If Not ComponentInProject(FileName) Then
                If MsgBox(Prompt:=FileName & " is not a part of Report Formatter v5.0. Add file to project anyway?", Title:="Report Formatter v5.0", Buttons:=vbYesNo) = vbYes Then
                    Application.ThisWorkbook.VBProject.VBComponents.Import objFile.Path
                End If
            Else
                If Not ReplaceFile(ModuleName:=FileName, FilePath:=objFile.Path) Then GoTo Failed
            End If
        Else
            'Check if file is one of the local files.
        End If
    Next objFile
End Sub
Private Sub CopyLooseFiles()

End Sub
