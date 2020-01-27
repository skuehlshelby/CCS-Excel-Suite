Attribute VB_Name = "BundledFunctions"
Option Explicit
Public Function OnlyLetters(Str As String, Optional ReplaceSpace As Boolean = False) As String

    OnlyLetters = Str
    
    Dim I As Long
    
    For I = 1 To Len(Str)
        If Asc(Mid$(Str, I, 1)) > IIf(ReplaceSpace, 31, 32) And Asc(Mid$(Str, I, 1)) < 65 Then
            OnlyLetters = Replace(OnlyLetters, Mid$(Str, I, 1), vbNullString, 1, -1, vbBinaryCompare)
        End If
    Next I
    
    OnlyLetters = Trim$(OnlyLetters)

End Function
Public Function OnlyNums(Str As String, Optional ReplaceSpace As Boolean = False) As String
    OnlyNums = Str
    
    Dim I As Long
    
    For I = 1 To Len(Str)
        If Asc(Mid$(Str, I, 1)) > IIf(ReplaceSpace, 31, 32) And Asc(Mid$(Str, I, 1)) < 65 Then
            OnlyNums = Replace(OnlyNums, Mid$(Str, I, 1), vbNullString, 1, -1, vbBinaryCompare)
        End If
    Next I
    
    OnlyNums = Trim$(OnlyNums)
End Function
Public Function Constrain(ByVal Number As Double, ByVal Min As Double, ByVal Max As Double) As Double

    Constrain = Number
    
    With Application.WorksheetFunction
        Constrain = .Max(Constrain, Min)
        Constrain = .Min(Constrain, Max)
    End With

End Function

