VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ceSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private SettingsFileLocation As String
Private Settings As ceIniFile
'Global, pre-declared class made with hidden attributes. DO NOT attempt to make more instances of this class.

Public Property Get SettingsPath() As String
    SettingsPath = SettingsFileLocation
End Property

Public Property Get AutoFreezeHeader() As Boolean
    AutoFreezeHeader = Settings.EntryRead("Auto-Freeze Header", True, "General", SettingsFileLocation)
End Property

Public Property Let AutoFreezeHeader(AutoFreeze As Boolean)
    Settings.EntryWrite "Auto-Freeze Header", AutoFreeze, "General", SettingsFileLocation
End Property

Public Property Get AddBlankRowBetweenHeaderAndData() As Boolean
    AddBlankRowBetweenHeaderAndData = Settings.EntryRead("Add Blank Row Between Header and Data", True, "General", SettingsFileLocation)
End Property

Public Property Let AddBlankRowBetweenHeaderAndData(AddBlank As Boolean)
    Settings.EntryWrite "Add Blank Row Between Header and Data", AddBlank, "General", SettingsFileLocation
End Property

Public Property Get HighlightActiveColumnandRow() As Boolean
    HighlightActiveColumnandRow = Settings.EntryRead("Highlight Active Column and Row", False, "General", SettingsFileLocation)
End Property

Public Property Let HighlightActiveColumnandRow(Highlight As Boolean)
    Settings.EntryWrite "Highlight Active Column and Row", Highlight, "General", SettingsFileLocation
End Property

Public Property Get HighlightColor() As Long
    HighlightColor = Settings.EntryRead("Highlight Color", 16247773, "General", SettingsFileLocation)
End Property

Public Property Let HighlightColor(Color As Long)
    Settings.EntryWrite "Highlight Color", Color, "General", SettingsFileLocation
End Property

Public Property Get MyProps() As String
    MyProps = Settings.EntryRead("MyProps", "Please paste a comma separated list of prop codes.", "General", SettingsFileLocation)
End Property

Public Property Let MyProps(PropList As String)
    Settings.EntryWrite "MyProps", PropList, "General", SettingsFileLocation
End Property

Public Property Get CreateExpectedFoldersIfNotFound() As Boolean
    CreateExpectedFoldersIfNotFound = Settings.EntryRead("Create Expected Folders If Not Found", True, "Smart-Save & Jump-To", SettingsFileLocation)
End Property

Public Property Let CreateExpectedFoldersIfNotFound(Create As Boolean)
    Settings.EntryWrite "Create Expected Folders If Not Found", Create, "Smart-Save & Jump-To", SettingsFileLocation
End Property

Public Property Get SmartSaveNamingConvention() As String
    SmartSaveNamingConvention = Settings.EntryRead("Smart-Save Naming Convention", "PropCode ReportName MMDDYYYY", "Smart-Save & Jump-To", SettingsFileLocation)
End Property

Public Property Let SmartSaveNamingConvention(Convention As String)
    Settings.EntryWrite "Smart-Save Naming Convention", Convention, "Smart-Save & Jump-To", SettingsFileLocation
End Property

Public Property Get CloseJumpToAfterFolderOpen() As Boolean
    CloseJumpToAfterFolderOpen = Settings.EntryRead("Close Jump-To After Folder Open", True, "Smart-Save & Jump-To", SettingsFileLocation)
End Property

Public Property Let CloseJumpToAfterFolderOpen(CloseAfter As Boolean)
    Settings.EntryWrite "Close Jump-To After Folder Open", CloseAfter, "Smart-Save & Jump-To", SettingsFileLocation
End Property

Public Property Get ReportsSupported() As Variant
    ReportsSupported = Settings.SectionRead(False, "Report Identification", SettingsFileLocation)
End Property

Public Property Get ReportTags() As Variant
    Dim TempArray As Variant
    TempArray = Settings.SectionRead(True, "Report Identification", SettingsFileLocation)
    Dim I As Long
    For I = LBound(TempArray) To UBound(TempArray)
        TempArray(I) = Split(TempArray(I), "=", -1, vbTextCompare)(1)
    Next I
    ReportTags = TempArray
End Property

Public Property Get ReportTag(Report As String) As String
    ReportTag = Settings.EntryRead(Report, vbNullString, "Report Identification", SettingsFileLocation)
End Property

Public Function AddSupportedReport(Report As String, Tags As String)
    Settings.EntryWrite Report, Tags, "Report Identification", SettingsFileLocation
End Function

Private Sub Class_Initialize()
    Set Settings = New ceIniFile
    SettingsFileLocation = "C:\Users\" & Application.UserName & "\AppData\Roaming\CCS Excel Suite\CCSExcelSuiteSettings.ini"
End Sub

Private Sub Class_Terminate()
    Set Settings = Nothing
    SettingsFileLocation = vbNullString
End Sub

