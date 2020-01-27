VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ceApplicationEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'@MemberAttribute VB_VarHelpID, -1
Public WithEvents AppEvent As Application
Attribute AppEvent.VB_VarHelpID = -1

Private Sub AppEvent_SheetActivate(ByVal Sh As Object)
    Params.Update
    Params.GetHeaderEndRow
End Sub

Private Sub AppEvent_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    If Settings.EntryRead("Highlight Selected Row And Column", False, "General", SettingsFileLocation) = True Then
        Target.Calculate
    End If
End Sub

Private Sub AppEvent_WorkbookActivate(ByVal Wb As Workbook)
    Params.Update
    Params.GetHeaderEndRow
    ReportType.Reset
End Sub

Private Sub AppEvent_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    If OpenWorkBookCount = 1 Then
        Globals.TerminateGlobals
        CloseAllWorkBooks
        Application.Quit
        CloseExcel
    End If
End Sub

Private Sub AppEvent_WorkbookOpen(ByVal Wb As Workbook)
    Start_CCS_Excel_Suite StartedBy:=App
End Sub

Private Function OpenWorkBookCount() As Long
    OpenWorkBookCount = 0
    
    Dim I As Long
    
    For I = 1 To Application.Workbooks.Count
        If Windows(I).Visible Then OpenWorkBookCount = OpenWorkBookCount + 1
        DoEvents
    Next I
End Function

