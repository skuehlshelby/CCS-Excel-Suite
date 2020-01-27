VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SetupForm 
   Caption         =   "CCS Excel Suite: Settings"
   ClientHeight    =   3390
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7755
   OleObjectBlob   =   "SetupForm.frx":0000
End
Attribute VB_Name = "SetupForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'@Folder("ReportFormatterV")
Private ColorIndex As Byte

Private Enum ColorChoices
    Highlighter = 65535
    Manilla = 13431551
    Peach = 14083324
    Mint = 14348258
    Blue = 16247773
    [_First] = Highlighter
    [_Last] = Blue
End Enum

Private ReportTags As Variant

Private Const RegExPropCode = "\b[Pp]rop[Cc]ode\b"
Private Const RegExMonth = "\b[Mm]{1,2}[ \.\-_]?[Dd]{0,2}[ \.\-_]?[Yy]{2,4}\b"
Private Const RegExReportName = "\b[Rr]eport[Nn]ame\b"

Private Sub Userform_Initialize()
    SetStartPosition
    
    PopulateListBox Me.lbSupportedReports, ceSettings.ReportsSupported
    
    ReportTags = ceSettings.ReportTags
    
    LoadSettings

    Me.SettingsTabs.Value = 0
    
    ColorIndex = 1
End Sub

Private Sub SetStartPosition()
    With New ceMousePosition
        Me.Top = .Top
        Me.Left = .Left
    End With
End Sub

Private Sub LoadSettings()
    With Me
        .cbAutoFreezeHeader.Value = ceSettings.AutoFreezeHeader
        .cbAddBlankRow.Value = ceSettings.AddBlankRowBetweenHeaderAndData
        .cbTriangulate.Value = ceSettings.HighlightActiveColumnandRow
        .lColor.BackColor = ceSettings.HighlightColor
        .tbMyProps.Value = ceSettings.MyProps
        .cbCreateIfNotFound.Value = ceSettings.CreateExpectedFoldersIfNotFound
        .SmartSaveNameFormat.Text = ceSettings.SmartSaveNamingConvention
        .SmartSaveNameFormat.ForeColor = CellFontGood
        .cbCloseAfterSelection.Value = ceSettings.CloseJumpToAfterFolderOpen
    End With
End Sub

Private Sub PopulateListBox(ByRef Box As MSForms.ListBox, ByVal List As Variant)
    Dim I As Long
    
    For I = LBound(List) To UBound(List)
        Box.AddItem (List(I))
    Next I
End Sub

Private Sub lbSupportedReports_Change()
    Me.tbTags.Value = ReportTags(Me.lbSupportedReports.ListIndex)
End Sub

Private Sub tbTags_Change()
    ReportTags(Me.lbSupportedReports.ListIndex) = Me.tbTags.Value
End Sub

Private Function GetListBoxSelected(ByVal Box As MSForms.ListBox) As String
    Dim I As Long
    For I = 0 To Box.ListCount - 1
        If Box.Selected(I) Then
            GetListBoxSelected = Box.List(I)
            Exit For
        End If
    Next I
End Function

Private Sub lGeneralOkayInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.lGeneralOkayInactive.Visible = False
End Sub

Private Sub lSmartSaveOkayInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.lSmartSaveOkayInactive.Visible = False
End Sub

Private Sub lReportIdentificationOkayInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.lReportIdentificationOkayInactive.Visible = False
End Sub

Private Sub Userform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    With Me
        .lGeneralOkayInactive.Visible = True
        .lSmartSaveOkayInactive.Visible = True
        .lReportIdentificationOkayInactive = True
    End With
End Sub

Private Sub SettingsTabs_MouseMove(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    With Me
        .lGeneralOkayInactive.Visible = True
        .lSmartSaveOkayInactive.Visible = True
        .lReportIdentificationOkayInactive.Visible = True
    End With
End Sub

Private Sub lColor_Click()
    Dim ColorArray As Variant
    ColorArray = Array(ColorChoices.Highlighter, ColorChoices.Manilla, ColorChoices.Peach, ColorChoices.Mint, ColorChoices.Blue)
    
    ColorIndex = ColorIndex + 1
    If ColorIndex > UBound(ColorArray) Then ColorIndex = LBound(ColorArray)
    
    Me.lColor.BackColor = ColorArray(ColorIndex)
End Sub

Private Sub SmartSaveNameFormat_Change()
    If SaveFormatGood(Me.SmartSaveNameFormat.Text) Then
        Me.SmartSaveNameFormat.ForeColor = CellFontGood
    Else
        Me.SmartSaveNameFormat.ForeColor = CellFontBad
    End If
End Sub
Private Function SaveFormatGood(UserInput As String) As Boolean
    
    Dim HasPropCode As Boolean
    Dim HasMonth As Boolean
    Dim HasReportName As Boolean
    Dim Matches As Object
    
    With CreateObject("VBScript.RegExp")
        
        .Global = True
        .IgnoreCase = False
        
        .Pattern = RegExPropCode
        Set Matches = .Execute(UserInput)
        HasPropCode = Matches.Count = 1
        
        .Pattern = RegExMonth
        Set Matches = .Execute(UserInput)
        HasMonth = Matches.Count = 1
        
        .Pattern = RegExReportName
        Set Matches = .Execute(UserInput)
        HasReportName = Matches.Count = 1
    End With
    
    SaveFormatGood = HasPropCode And HasMonth And HasReportName
End Function

Private Sub SmartSaveNameFormat_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not SaveFormatGood(Me.SmartSaveNameFormat.Text) Then
        Me.SmartSaveNameFormat.Text = Settings.EntryRead("Smart-Save Naming Convention", "PropCode ReportName MMDDYYYY", "Smart-Save & Jump-To", SettingsFilePath)
    End If
End Sub

Private Sub lGeneralOkayActive_Click()
    WriteSettings
    Me.Hide
End Sub

Private Sub WriteSettings()
    With ceSettings
        .AutoFreezeHeader = Me.cbAutoFreezeHeader.Value
        .AddBlankRowBetweenHeaderAndData = Me.cbAddBlankRow.Value
        .HighlightActiveColumnandRow = Me.cbTriangulate.Value
        .HighlightColor = Me.lColor.BackColor
        .MyProps = Me.tbMyProps.Value
        .CreateExpectedFoldersIfNotFound = Me.cbCreateIfNotFound.Value
        .SmartSaveNamingConvention = Me.SmartSaveNameFormat.Text
        .CloseJumpToAfterFolderOpen = Me.cbCloseAfterSelection.Value
    End With
End Sub

Private Sub UserForm_Terminate()
    ColorIndex = 0
    Erase ReportTags
End Sub
