Attribute VB_Name = "Setup"
'@Folder("ReportFormatterV")
Option Explicit
Option Private Module
Public Sub PerformSetup()
    With New ceDriveNavigation
        If .CreateTextFile(ceSettings.SettingsPath, False) Then CreateDefaults (ceSettings.SettingsPath)
    End With
    
    CCS_Excel_Suite_Change_Settings
End Sub
Private Sub CreateDefaults(ByVal SettingsFilePath As String)
    With ceSettings
        .AutoFreezeHeader = True
        .AddBlankRowBetweenHeaderAndData = True
        .HighlightActiveColumnandRow = False
        .HighlightColor = 16247773
        .MyProps = "cb81, cb82"
        .CreateExpectedFoldersIfNotFound = True
        .SmartSaveNamingConvention = "PropCode ReportName MMDDYYYY"
        .CloseJumpToAfterFolderOpen = True
        .AddSupportedReport "Transfer Export", "transfer"
        .AddSupportedReport "SHBBC Export", "shbbc"
        .AddSupportedReport "EDE SSRS", "edediscrepancyfiles"
        .AddSupportedReport "Discrepancy File", "discrepancy,descrepency,descrepancy,inmoveout,stepqc"
        .AddSupportedReport "Property Consumption", "propertyconsumption"
        .AddSupportedReport "Vacant QC", "vacantchargesqc"
        .AddSupportedReport "Vacant Holding Worksheet", "vacant"
        .AddSupportedReport "Utility Difference Report", "utilitydifference"
        .AddSupportedReport "Factored Occs QC", "factoredoccs"
        .AddSupportedReport "Resident Report", "resident"
    End With
End Sub
