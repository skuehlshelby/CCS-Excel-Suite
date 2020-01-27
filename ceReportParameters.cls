VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ceReportParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("ReportFormatterV")
Option Explicit
Option Base 1
Private UsedRangeAddress As Variant
Private pHeaderEndRow As Long

Private Sub Class_Initialize()
    Call Update
End Sub

Private Sub Class_Terminate()
    UsedRangeAddress = vbNullString
    pHeaderEndRow = 0
End Sub

Public Sub Update()
    ActiveSheet.UsedRange
    
    UsedRangeAddress = ActiveSheet.UsedRange.Address
    UsedRangeAddress = Replace(UsedRangeAddress, ":", vbNullString, 1, -1, vbTextCompare)
    UsedRangeAddress = Replace(UsedRangeAddress, "$", vbNullString, 1, 1, vbTextCompare)
    UsedRangeAddress = Split(UsedRangeAddress, "$", -1)
End Sub

Public Function GetHeaderEndRow(Optional ByVal ReportType As String = "Unknown") As Long
    
    If ActiveWindow.SplitRow <> 0 Then
        GetHeaderEndRow = ActiveWindow.SplitRow
        Exit Function
    End If
    
    Select Case ReportType
    
    Case "Discrepancy File"
        GetHeaderEndRow = 0
        Exit Function
    Case "EDE SSRS", "Property Consumption"
        GetHeaderEndRow = 4
        Exit Function
    Case "SHBBC Export", "Transfer Export", "Vacant Holding Worksheet", "Factored Occs QC", "Resident Report"
        GetHeaderEndRow = 1
        Exit Function
    Case "Utility Difference Report"
    Case "Vacant QC"
    Case Else
        GetHeaderEndRow = Application.WorksheetFunction.Max(FirstRowWithNumber - 1, 0)
    End Select
    
End Function

Private Function FirstRowWithNumber(Optional ByVal StopAtRow As Long = 16) As Long
    FirstRowWithNumber = 0
    
    Dim R As Range, C As Range
    
    For Each R In ActiveSheet.UsedRange.Rows
        For Each C In R.Columns
            If C.Address = C.MergeArea.Cells(1).Address Then
                If IsANumber(CStr(C.Value2)) Then
                    FirstRowWithNumber = C.Row
                    Exit For
                End If
            End If
        Next C
        If Not FirstRowWithNumber = 0 Then Exit For
        If R.Row >= StopAtRow Then Exit For
    Next R
End Function

Private Function IsANumber(ByVal TestString As String) As Boolean
    If TestString = vbNullString Then
        IsANumber = False
        Exit Function
    End If
    
    IsANumber = True
    
    Dim I As Long
    
    For I = 1 To Len(TestString)
        If Asc(Mid(TestString, I, 1)) < 48 Or Asc(Mid(TestString, I, 1)) > 57 Then
            IsANumber = False
            Exit For
        End If
    Next I
End Function

Public Sub FreezeHeader(Optional ByVal AddSpace As Boolean = True)
    Dim CurrentScrollRow As Long
    Dim CurrentScrollColumn As Long
    
    If ActiveWindow.SplitRow = 0 Then
        
    End If
End Sub

Public Sub TrimUsedRange(Optional I As Long = 1, Optional UsedRangeChanged As Boolean = False)
    
    I = LastRow
    
    While Application.WorksheetFunction.CountA(ActiveSheet.Cells(I, 1).EntireColumn) = 0 And I > FirstRow
        ActiveSheet.Cells(I, 1).EntireRow.Delete Shift:=xlShiftUp
        UsedRangeChanged = True
        I = I - 1
    Wend

    I = LastColumnNum
    
    While Application.WorksheetFunction.CountA(ActiveSheet.Cells(1, I).EntireColumn) = 0 And I > FirstColumnNum
        ActiveSheet.Cells(1, I).EntireColumn.Delete Shift:=xlShiftToLeft
        UsedRangeChanged = True
        I = I - 1
    Wend
  
    If UsedRangeChanged Then Call Update

End Sub

Private Function LetterToNumber(Letter As String) As Long
    While Len(Letter) > 1
        LetterToNumber = LetterToNumber + Asc(Left$(Letter, 1)) - 39
        Letter = Right$(Letter, Len(Letter) - 1)
    Wend
    LetterToNumber = LetterToNumber + Asc(Left$(Letter, 1)) - 64
End Function

Public Property Get FirstColumn() As String
    FirstColumn = UsedRangeAddress(0)
End Property

Public Property Get LastColumn() As String
    LastColumn = UsedRangeAddress(UBound(UsedRangeAddress) - 1)
End Property

Public Property Get FirstRow() As Long
    FirstRow = pHeaderEndRow + 1
    
    While Application.WorksheetFunction.CountA(ActiveSheet.Range(UsedRangeAddress(0) & FirstRow).EntireRow) = 0 And FirstRow < UsedRangeAddress(UBound(UsedRangeAddress))
        FirstRow = FirstRow + 1
    Wend
End Property

Public Property Get LastRow() As Long
    LastRow = UsedRangeAddress(UBound(UsedRangeAddress))
End Property

Public Property Get LastColumnNum() As Long
    LastColumnNum = LetterToNumber(LastColumn)
End Property

Public Property Get FirstColumnNum() As String
    FirstColumnNum = LetterToNumber(FirstColumn)
End Property

Public Property Get Width() As Long
    Width = LastColumnNum - FirstColumnNum + 1
End Property

Public Property Get Header() As Range
    Set Header = Nothing
    If pHeaderEndRow <> 0 Then Set Header = ActiveSheet.Range(UsedRangeAddress(0) & UsedRangeAddress(1), UsedRangeAddress(UBound(UsedRangeAddress) - 1) & pHeaderEndRow)
End Property

Public Property Get Data() As Range
    Set Data = ActiveSheet.Range(UsedRangeAddress(0) & FirstRow, UsedRangeAddress(UBound(UsedRangeAddress) - 1) & UsedRangeAddress(UBound(UsedRangeAddress)))
End Property
