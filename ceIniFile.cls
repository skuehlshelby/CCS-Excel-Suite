VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ceIniFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Win32 API declarations
Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare PtrSafe Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpApplicationName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare PtrSafe Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" (ByVal lpReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

' Some default values
Private Const defBufferSize As Long = 2048
Private Const defStripComments As Boolean = True

' Member variables
Private m_FileName As String
Private m_Section As String
Private m_StripComments As Boolean

' Public events
Public Event EnumIniSection(ByVal SectionName As String, ByVal FileName As String, Cancel As Boolean)
Public Event EnumIniSectionEntry(ByVal EntryName As String, ByVal SectionName As String, ByVal FileName As String, Cancel As Boolean)

' *************************************************
'  Initialization
' *************************************************
Private Sub Class_Initialize()
   m_StripComments = defStripComments
End Sub

Private Sub Class_Terminate()
   Me.Flush  'can't hurt
End Sub

' *************************************************
'  Public Properties
' *************************************************
Public Property Get FileName() As String
   FileName = m_FileName
End Property

Public Property Let FileName(ByVal IniFile As String)
   m_FileName = IniFile
End Property

Public Property Get Section() As String
   Section = m_Section
End Property

Public Property Let Section(ByVal IniSection As String)
   m_Section = IniSection
End Property

Public Property Get StripComments() As Boolean
   StripComments = m_StripComments
End Property

Public Property Let StripComments(ByVal NewSetting As Boolean)
   m_StripComments = NewSetting
End Property

' *************************************************
'  Public Methods
' *************************************************
Public Function EntryClear(ByVal EntryName As String, Optional ByVal Section As String, Optional ByVal FileName As String) As Boolean
   ' Update cached section and filename, if provided.
   Call UpdateFileDetails(Section, FileName)
   ' Write an empty string to entry.
   EntryClear = CBool(WritePrivateProfileString(m_Section, EntryName, "", m_FileName))
End Function

Public Function EntryDelete(ByVal EntryName As String, Optional ByVal Section As String, Optional ByVal FileName As String) As Boolean
   ' Update cached section and filename, if provided.
   Call UpdateFileDetails(Section, FileName)
   ' Totally zap this entry.
   EntryDelete = CBool(WritePrivateProfileString(m_Section, EntryName, vbNullString, m_FileName))
End Function

Public Function EntryRead(ByVal EntryName As String, Optional ByVal DefaultValue As String, Optional ByVal Section As String, Optional ByVal FileName As String) As String
   Dim Buffer As String
   Dim nRet As Long

   ' Update cached section and filename, if provided.
   Call UpdateFileDetails(Section, FileName)

   ' Establish buffer for return info.
   Buffer = Space$(defBufferSize)
   
   ' May need several attempts if entry is huge.
   Do
      nRet = GetPrivateProfileString(m_Section, EntryName, DefaultValue, Buffer, Len(Buffer), m_FileName)
      If nRet = (Len(Buffer) - 1) Then
         ' Need to expand buffer and try again.
         ' SDK Docs specify Len-1 for this case.
         Buffer = Space$(Len(Buffer) + defBufferSize)
      ElseIf nRet > 0 Then
         ' Trim null and remaining buffer.
         Buffer = Left$(Buffer, nRet)
         Exit Do
      Else
         ' Function failed.
         Buffer = ""
         Exit Do
      End If
   Loop
   
   ' Strip off everything from the semi-colon on?
   If m_StripComments Then
      nRet = InStr(Buffer, ";")
      If nRet = 1 Then
         Buffer = ""
      ElseIf nRet > 1 Then
         Buffer = Left$(Buffer, nRet - 1)
      End If
   End If
   
   ' Return results.
   EntryRead = Trim$(Buffer)
End Function

Public Function EntryWrite(ByVal EntryName As String, ByVal EntryValue As String, Optional ByVal Section As String, Optional ByVal FileName As String) As String
   ' Update cached section and filename, if provided.
   Call UpdateFileDetails(Section, FileName)
   ' Write this value as requested.
   Call WritePrivateProfileString(m_Section, EntryName, EntryValue, m_FileName)
End Function

Public Sub EnumSectionEntries(Optional ByVal Section As String, Optional ByVal FileName As String)
   Dim nRet As Long, I As Long
   Dim Buffer As String
   Dim arrEntry() As String
   Dim BailOut As Boolean

   ' Update cached section and filename, if provided.
   Call UpdateFileDetails(Section, FileName)

   ' Establish buffer for return info.
   Buffer = Space$(defBufferSize)
   
   ' May need several attempts if large
   ' number of entries.
   Do
      nRet = GetPrivateProfileString(m_Section, vbNullString, "", Buffer, Len(Buffer), m_FileName)
      If nRet = (Len(Buffer) - 2) Then
         ' Need to expand buffer and try again.
         ' SDK Docs specify Len-2 for this case.
         Buffer = Space$(Len(Buffer) + defBufferSize)
      Else
         Exit Do
      End If
   Loop
   
   ' Extract entry names.
   If ExtractStringsZ(Buffer, arrEntry()) Then
      ' Raise an event for each element in the array.
      For I = LBound(arrEntry) To UBound(arrEntry)
         ' Offer user chance to stop enumeration by bailing
         ' from loop if the last parameter is set to True.
         RaiseEvent EnumIniSectionEntry(arrEntry(I), m_Section, m_FileName, BailOut)
         If BailOut Then Exit For
      Next I
   End If
End Sub

Public Sub EnumSections(Optional ByVal FileName As String)
   Dim nRet As Long, I As Long
   Dim Buffer As String
   Dim arrSection() As String
   Dim BailOut As Boolean
   
   ' Update cached section and filename, if provided.
   Call UpdateFileDetails(, FileName)
   
   ' Establish buffer for return info.
   Buffer = Space$(defBufferSize)
   
   ' May need several attempts if large
   ' number of sections.
   Do
      nRet = GetPrivateProfileSectionNames(Buffer, Len(Buffer), m_FileName)
      If nRet = (Len(Buffer) - 2) Then
         ' Need to expand buffer and try again.
         ' SDK Docs specify Len-2 for this case.
         Buffer = Space$(Len(Buffer) + defBufferSize)
      Else
         Exit Do
      End If
   Loop
   
   ' Extract section names.
   If ExtractStringsZ(Buffer, arrSection()) Then
      ' Raise an event for each element in the array.
      For I = LBound(arrSection) To UBound(arrSection)
         ' Offer user chance to stop enumeration by bailing
         ' from loop if the last parameter is set to True.
         RaiseEvent EnumIniSection(arrSection(I), m_FileName, BailOut)
         If BailOut Then Exit For
      Next I
   End If
End Sub

Public Sub Flush(Optional ByVal FileName As String)
   ' Update cached section and filename, if provided.
   Call UpdateFileDetails(, FileName)
   ' To improve peceormance, Windows keeps a cached version of
   ' the most-recently accessed initialization file. If that
   ' filename is specified and the other three parameters are
   ' NULL, Windows flushes the cache.
   Call WritePrivateProfileString(vbNullString, vbNullString, vbNullString, m_FileName)
End Sub

   Public Function SectionDelete(Optional ByVal Section As String, Optional ByVal FileName As String) As Boolean
      ' Update cached section and filename, if provided.
      Call UpdateFileDetails(Section, FileName)
      ' Remove specific section from INI file.
      SectionDelete = CBool(WritePrivateProfileString(m_Section, vbNullString, vbNullString, m_FileName))
   End Function

Public Function SectionRead(Optional IncludeValues As Boolean = False, Optional ByVal Section As String, Optional ByVal FileName As String) As Variant
   ' This function will return an array filled with either
   ' the entry names (similar to EnumSectionEntries) or the
   ' complete entries (including the values for each).
   Dim nRet As Long, I As Long
   Dim Buffer As String
   Dim arrEntry() As String

   ' Update cached section and filename, if provided.
   Call UpdateFileDetails(Section, FileName)

   ' Establish buffer for return info.
   Buffer = Space$(defBufferSize)
   
   ' May need several attempts if large
   ' number of entries.
   Do
      ' Use different APIs depending on whether they
      ' want to include full entries, or just names.
      If IncludeValues Then
         nRet = GetPrivateProfileSection(m_Section, Buffer, Len(Buffer), m_FileName)
      Else
         nRet = GetPrivateProfileString(m_Section, vbNullString, "", Buffer, Len(Buffer), m_FileName)
      End If
         
      If nRet = (Len(Buffer) - 2) Then
         ' Need to expand buffer and try again.
         ' SDK Docs specify Len-2 for this case.
         Buffer = Space$(Len(Buffer) + defBufferSize)
      Else
         Exit Do
      End If
   Loop
   
   ' Extract entry names.
   If ExtractStringsZ(Buffer, arrEntry()) Then
      SectionRead = arrEntry
   End If
End Function

Public Function ToBoolean(ByVal EntryValue As String) As Boolean
   ' Interpret entry as either true or false.
   Select Case Trim$(UCase$(EntryValue))
      Case "YES", "Y", "TRUE", "T", "ON", "1", "-1"
         ToBoolean = True
      Case "NO", "N", "FALSE", "F", "OFF", "0"
         ToBoolean = False
      Case Else
         ToBoolean = False
   End Select
End Function

Public Function ToDouble(ByVal EntryValue As String) As Double
   ' Return interpreted floating point value of entry.
   On Error Resume Next
   ToDouble = CDbl(Val(EntryValue))
End Function

Public Function ToLong(ByVal EntryValue As String) As Long
   ' Return integral portion of entry.
   On Error Resume Next
   ToLong = CLng(Val(EntryValue))
End Function

' *************************************************
'  Private Methods
' *************************************************
Private Function ExtractStringsZ(Buffer As String, OutArray() As String) As Long
   Dim StartPos As Long
   Dim NullPos As Long
   Dim BuffLen As Long
   Dim Elements As Long
   
   ' Extract null terminated strings from large
   ' double-null terminated buffer.
   StartPos = 1
   Elements = 0
   BuffLen = Len(Buffer)
   
   ' Loop through buffer looking for nulls.
   Do While StartPos < BuffLen
      NullPos = InStr(StartPos, Buffer, vbNullChar)
      If NullPos = StartPos Then
         ' We've hit the double-null terminator.
         Exit Do
      Else
         ' Expand array, store new substring, and
         ' increment counters.
         ReDim Preserve OutArray(0 To Elements) As String
         OutArray(Elements) = Mid$(Buffer, StartPos, NullPos - StartPos)
         StartPos = NullPos + 1
         Elements = Elements + 1
      End If
   Loop
   
   ' Return number of substrings found.
   ExtractStringsZ = Elements
End Function

Private Sub UpdateFileDetails(Optional ByVal Section As String, Optional ByVal FileName As String)
   ' Cache new values if different from old ones.
   If Len(Section) Then
      m_Section = Section
   End If
   If Len(FileName) Then
      m_FileName = FileName
   End If
End Sub

