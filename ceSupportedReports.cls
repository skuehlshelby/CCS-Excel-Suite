VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ceSupportedReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private vSupportedReports As Scripting.Dictionary
Private vReportType As String

Public Enum ceSupportedReport
    Transfer_Export
    SHBBC_Export
    EDE_SSRS
    Discrepancy_File
    Billing_Summary
    Property_Consumption
    Vacant_QC
    Vacant_Holding_Worksheet
    Utility_Difference_Report
    Factored_Occs_QC
    Resident_Report
    [_First] = Transfer_Export
    [_Last] = Resident_Report
End Enum

Public Property Get IsSupported() As Boolean
    IsSupported = Me.Name <> "Unknown"
End Property

Public Property Get Report(Name As ceSupportedReport) As String
    Report = CStr(vSupportedReports.Keys(Name))
End Property

Public Property Get Name() As String
    If vReportType = vbNullString Then vReportType = GetReportType
    Name = vReportType
End Property

Public Sub Reset()
    vReportType = vbNullString
End Sub

Public Function GetSupportedReports() As Collection
    Set GetSupportedReports = New Collection
    Dim TempArray As Variant
    TempArray = vSupportedReports.Keys
    Dim I As Long
    For I = LBound(TempArray) To UBound(TempArray)
        GetSupportedReports.Add TempArray(I)
    Next I
End Function

Public Function GetReportTags() As Collection
    Set GetReportTags = New Collection
    Dim I As Long
    For I = LBound(TempArray) To UBound(TempArray)
        GetReportTags.Add vSupportedReports.Item(I)
    Next I
End Function

Private Function GetReportType(Optional I As Long = 1) As String
    
    With ActiveWorkbook
        GetReportType = .ActiveSheet.Name & .Name & .ActiveSheet.Range("A1").Value2
    End With
    
    GetReportType = LettersOnly(GetReportType, True)
    
    GetReportType = LCase$(GetReportType)
    
    Dim ControlVar As Variant, TempArray As Variant
    
    For Each ControlVar In vSupportedReports.Keys
        
        TempArray = Split(vSupportedReports.Item(ControlVar), ",", -1, vbTextCompare)
        
        For I = LBound(TempArray) To UBound(TempArray)
            If InStr(1, GetReportType, TempArray(I), vbTextCompare) > 0 Then
                GetReportType = ControlVar
                Exit Function
            End If
        Next I
    Next ControlVar
    
    GetReportType = "Unknown"

End Function

Private Function LettersOnly(InputString As String, Optional ReplaceSpace As Boolean = False) As String

    LettersOnly = InputString
    
    Dim I As Long
    
    For I = 1 To Len(InputString)
        If Asc(Mid$(InputString, I, 1)) > IIf(ReplaceSpace, 31, 32) And Asc(Mid$(InputString, I, 1)) < 65 Then
            LettersOnly = Replace(LettersOnly, Mid$(InputString, I, 1), vbNullString, 1, -1, vbBinaryCompare)
        End If
    Next I
    
    LettersOnly = Trim$(LettersOnly)

End Function

Private Sub Class_Initialize()
    Set vSupportedReports = New Scripting.Dictionary
    
    With vSupportedReports
        .Add Item:="transfer", Key:="Transfer Export"
        .Add Item:="shbbc", Key:="SHBBC Export"
        .Add Item:="edediscrepancyfiles", Key:="EDE SSRS"
        .Add Item:="discrepancy,descrepency,descrepancy,inmoveout,stepqc", Key:="Discrepancy File"
        .Add Item:="summaryofutilities", Key:="Billing Summary"
        .Add Item:="propertyconsumption", Key:="Property Consumption"
        .Add Item:="vacantchargesqc", Key:="Vacant QC"
        .Add Item:="vacant", Key:="Vacant Holding Worksheet"
        .Add Item:="utilitydifference", Key:="Utility Difference Report"
        .Add Item:="factoredoccs", Key:="Factored Occs QC"
        .Add Item:="residents", Key:="Resident Report"
    End With
End Sub

Private Sub Class_Terminate()
    Set vSupportedReports = Nothing
    vReportType = vbNullString
End Sub

