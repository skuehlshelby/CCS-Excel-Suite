Attribute VB_Name = "Shutdown"
Option Explicit
Option Private Module
Public Sub CloseAllWorkBooks()
    Application.EnableEvents = False
    
    Dim ControlVar As Workbook
    
    For Each ControlVar In Application.Workbooks
        ControlVar.Close
    Next ControlVar
    
    Application.EnableEvents = True
End Sub

Public Sub CloseExcel()
    Dim CloseCommand As String
    CloseCommand = "C:\Users\" & Application.UserName & "\AppData\Roaming\CCS Excel Suite\CloseAllExcel.bat"
    Shell PathName:=CloseCommand, windowStyle:=vbHide
End Sub
