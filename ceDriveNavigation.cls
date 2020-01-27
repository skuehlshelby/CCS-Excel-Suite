VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ceDriveNavigation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private CurrentFolder As Folder
Private FSO As FileSystemObject

Public Function MostRecentSubfolder(ByVal ParentFolderPath As String, Optional NameMatchesPattern As String = "??.??") As String
    Dim ControlVar As Folder
    
    If FSO.GetFolder(ParentFolderPath).SubFolders.Count = 0 Then
        MostRecentSubfolder = CurrentFolder.Path
        Exit Function
    End If
    
    For Each ControlVar In FSO.GetFolder(ParentFolderPath).SubFolders
        If CurrentFolder Is Nothing Then Set CurrentFolder = ControlVar
        If ControlVar.DateCreated > CurrentFolder.DateCreated And ControlVar.Name Like NameMatchesPattern Then Set CurrentFolder = ControlVar
    Next ControlVar
    
    MostRecentSubfolder = CurrentFolder.Path
End Function

Public Function DaysOldFolder(ByVal FolderPath As String) As Long
    Dim tempFolder As Folder
    Set tempFolder = FSO.GetFolder(FolderPath)
    DaysOldFolder = Date - tempFolder.DateCreated
End Function

Public Function CreateTextFile(ByVal FilePath As String, ByVal OverWrite As Boolean) As Boolean
    CreateTextFile = False
    If Not DirectoryExists(FSO.GetParentFolderName(FilePath)) Then CreateFolder (FSO.GetParentFolderName(FilePath))
    If Not OverWrite Then
        If Not DirectoryExists(FilePath) Then
            FSO.CreateTextFile FilePath, OverWrite
            CreateTextFile = True
        End If
    Else
        FSO.CreateTextFile FilePath, OverWrite
        CreateTextFile = True
    End If
End Function

Public Sub CreateFolder(ByVal FolderPath As String)
    FSO.CreateFolder (FolderPath)
End Sub

Public Function DirectoryExists(ByVal Path As String) As Boolean
    If Path = vbNullString Then
        DirectoryExists = False
        Exit Function
    End If
    DirectoryExists = Dir(Path, vbDirectory) <> vbNullString
End Function

Private Sub Class_Initialize()
    Set FSO = New FileSystemObject
End Sub

Private Sub Class_Terminate()
    Set FSO = Nothing
End Sub
