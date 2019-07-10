VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringClass"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'* Construct a new string by using the insertion point within a string
'* For use when typing to detect BACKSPACE, DEL etc
'*
'* string1          string to be manipulated
'* iInsert          insertion point in that string
'* char             ASCII character in question
'*
'* return value:
'* true if text inserted at end or False otherwise
Public Function ConstructString(String1 As String, iInsert As Integer, char As Integer) As Boolean
Dim i As Integer

  Call gTracer.xSet("ConstructString")
  ConstructString = False
  If iInsert = Len(String1) Then
    Select Case char
      Case 32 To 126 'valid character
        String1 = String1 & UCase(Chr(char))
        ConstructString = True
      Case 8 'backspace character
        If Len(String1) Then String1 = Left$(String1, Len(String1) - 1)
      Case -1 'del key pressed
    End Select
  Else 'text inserted/deleted not from end !?
    i = iInsert
    Select Case char
      Case 32 To 126 'valid character
        String1 = Left$(String1, i) & UCase(Chr(char)) & Right$(String1, Len(String1) - i)
      Case 8  'backspace
        If Len(String1) And (i <> 0) Then String1 = Left$(String1, i - 1) & Right$(String1, Len(String1) - i)
      Case -1
        If Len(String1) And (i <> 0) Then String1 = Left$(String1, i) & Right$(String1, Len(String1) - i)
    End Select
  End If
  Call gTracer.xReturn("ConstructString")
  
End Function

'* Prepares a string to use in SQL
'*
'* sourcestr - source string
'*
'* return value:
'* Valid SQL string
Public Function StrSQL(ByVal sourcestr As String) As String
  Dim i As Long
  Dim q As String
    
  Call gTracer.xSet("StrSQL")
  i = 1
  Do
    q = Mid$(sourcestr, i, 1)
    If q = "'" Then
      sourcestr = Left$(sourcestr, i) & "'" & Right$(sourcestr, Len(sourcestr) - i)
      i = i + 1
    End If
    i = i + 1
  Loop Until Len(q) = 0
  StrSQL = sourcestr
  Call gTracer.xReturn("StrSQL")
End Function

'* Duplicate a single character in a string - use StrSQL for SQL strings
'*
'* sourcestr - source string
'* sdupchar - character to be replaced
'*
'* return value:
'* the string with the duplicated character
Public Function StrDupChar(ByVal sourcestr As String, ByVal dupchar As String) As String
  Dim i As Long
  Dim q As String
    
  Call gTracer.xSet("StrDupChar")
  dupchar = Left$(dupchar, 1)
  If Len(dupchar) Then
    i = 1
    Do
      q = Mid$(sourcestr, i, 1)
      If q = dupchar Then
        sourcestr = Left$(sourcestr, i) & dupchar & Right$(sourcestr, Len(sourcestr) - i)
        i = i + 1
      End If
      i = i + 1
    Loop Until Len(q) = 0
  End If
  StrDupChar = sourcestr
  Call gTracer.xReturn("StrDupChar")
End Function



'* Pad a string at front or back with a character to a specified length
'*
'* String1        original string
'* sPad           padding character
'* Length         length of padded string
'* [bPadFront]    pad from front (default = false)
'*
'* return value:
'* The padded string
Public Function xStrPad(String1 As String, sPad As String, Length As Long, Optional bPadFront As Boolean = False) As String
  Dim padlen As Long
  Dim s As String
    
  Call gTracer.xSet("xStrPad")
  padlen = Length - Len(String1)
  If padlen > 0 Then
    s = String$(padlen, sPad)
    If bPadFront Then
      xStrPad = s & String1
    Else
      xStrPad = String1 & s
    End If
  Else
    xStrPad = Left$(String1, Length)
  End If
  Call gTracer.xReturn("xStrPad")
End Function

'* Same as InStr but starts from the end of the string
'* Start still references from the beginning of the string
'*
'* String1              string to be searched
'* String2              string to find in string
'* [Start]              start the search from position (default =1)
'* [Compare]            Compare method (default = vbBinaryCompare)
'*
'* return value:
'* The position in the string of the search string or zero if the string is not found
Public Function InStrBack(String1 As String, String2 As String, Optional Start As Long = 1, Optional Compare As VbCompareMethod = vbBinaryCompare) As Long
  Dim pos As Long
  Dim lastpos As Long
  
  Call gTracer.xSet("InStrBack")
  pos = InStr(Start, String1, String2, Compare)
  If pos > 0 Then
    Do
      lastpos = pos
      pos = InStr(pos + 1, String1, String2, Compare)
    Loop Until pos = 0
    pos = lastpos
  End If
  InStrBack = pos
  Call gTracer.xReturn("InStrBack")
End Function

'* Counts the occurences of a character / string in a string
'*
'* String1                string to be searched
'* Char                   character to be found in string
'* [bSingleChar]          single character search
'* [bCase]                Case sensitive search (default = true)
'*
'* return value:
'* Number of occurences of that string
Public Function CountChar(String1 As String, ByVal char As String, Optional bSingleChar As Boolean = True, Optional bCase As Boolean = True) As Long
  Dim l As Long, llen As Long
    
  Call gTracer.xSet("CountChar")
  If bSingleChar Then
    If Len(char) <> 1 Then
      CountChar = -1
      Exit Function
    End If
  End If
  
  llen = Len(char)
  If bCase Then
    For l = 1 To Len(String1) - (llen - 1)
      If Mid$(String1, l, llen) = char Then CountChar = CountChar + 1
    Next l
  Else
    char = UCase(char)
    For l = 1 To Len(String1) - (llen - 1)
      If UCase(Mid$(String1, l, llen)) = char Then CountChar = CountChar + 1
    Next l
  End If
  
  Call gTracer.xReturn("CountChar")
End Function


'* Extracts fields from a delimited string
'*
'* String1         string to extract FROM
'* Char            character delimiter
'* Fields          parameter fields to extract to in order
'*
'* return value:
'* Fields in Parameter array Fields
Public Sub xStrExtract(String1 As String, char As String, ParamArray Fields() As Variant)
  Dim tchar As String
  Dim i As Long
  Dim iStartPos As Long
  Dim iEndPos As Integer
  Dim sText As String
    
  Call gTracer.xSet("xStrExtract")
  tchar = Left$(char, 1)
  
  iStartPos = 1
  For i = LBound(Fields) To UBound(Fields)
    If iStartPos > Len(String1) Then GoTo endfunc
    iEndPos = InStr(iStartPos, String1, char, vbTextCompare)
    If iEndPos = 0 Then
      Fields(i) = Right$(String1, Len(String1) - iStartPos + 1)
      Exit For
    Else
      Fields(i) = Mid$(String1, iStartPos, iEndPos - iStartPos)
      iStartPos = iEndPos + 1
    End If
  Next i

endfunc:
  Call gTracer.xReturn("xStrExtract")
  
End Sub
'* Removes a character from a string
'*
'* String1              string to search
'* Char                 character to be removed
'*
'* return value:
'* none
' apf check
'RMC - this does not return anything
Public Function xStrRemoveChar(String1 As String, ByVal char As String) As Long
  Dim i As Long
  Dim sChar As String

  Call gTracer.xSet("xStrRemoveChar")
  i = 1
  Do
    sChar = Mid$(String1, i, 1)
    If sChar <> "" Then
      If sChar = char Then
        String1 = Left$(String1, i - 1) & Right$(String1, Len(String1) - i)
      Else
        i = i + 1
      End If
    End If
  Loop Until sChar = ""
  Call gTracer.xReturn("xStrRemoveChar")
  
End Function


'* Returns a numbered field from a string of fields seperated by tokens
'*
'* String1            source string
'* sToken             seperator token
'* TokenIndex         the index of the field to be extracted (1=Field1, 2=Field2)
'* [bTrim]            trim spaces off the field (default = false)
'*
'* return value:
'* the extracted string
Public Function GetStringToken(String1 As String, sToken As String, Optional TokenIndex As Long = 1, Optional bTrim As Boolean = False) As String
  Dim pos As Long
  Dim lpos As Long
  Dim i As Long
  
  ' sort out apf tokenize XXX  vbTab YYY
  Call gTracer.xSet("GetStringToken")
  If Len(String1) = 0 Then GoTo GetStringToken_End
  
  pos = 0: i = 0
  Do
    lpos = pos
    pos = InStr(pos + 1, String1, sToken, vbTextCompare)
    i = i + 1
  Loop Until (i = TokenIndex) Or (pos = 0)
  
  If pos = 0 Then
    If i <> TokenIndex Then
      GoTo GetStringToken_End
    End If
  End If
  
  If lpos = 0 Then
    If pos = 0 Then
      GetStringToken = String1
    Else
      GetStringToken = Left$(String1, pos - 1)
    End If
  ElseIf pos = 0 Then
    GetStringToken = Right$(String1, Len(String1) - lpos)
  ElseIf pos > 0 Then
    GetStringToken = Mid$(String1, lpos + 1, pos - lpos - 1)
  End If
  
  If bTrim Then GetStringToken = Trim$(GetStringToken)
  
GetStringToken_End:
  Call gTracer.xReturn("GetStringToken")
End Function

'* Finds the first occurence of a character in String1 which is not in String2
'*
'* String1           source string
'* String2           ignore character
'* [Start]           start point (Default = 1)
'* [Compare]         compare method (Default = vbBinaryCompare)
'* return value:
'* the position of the first character which is not String1,
'* or if String2 is not in String1 then Len(String1)+1
Public Function NotInStr(String1 As String, String2 As String, Optional Start As Long = 1, Optional Compare As VbCompareMethod = vbBinaryCompare) As Long
  Dim i As Long
  Dim s As String
  
  Call gTracer.xSet("NotInstr")
  If String1 = "" Or String2 = "" Then
    NotInStr = 0
  Else
    If Len(String2) > 1 Then
      For i = Start To Len(String1)
        s = Mid$(String1, i, 1)
        If InStr(1, String2, s, Compare) = 0 Then Exit For
      Next i
      'If i > Len(String1) Then
        'NotInStr = 0
      'Else
        NotInStr = i
      'End If
    Else
      s = String2
      For i = Start To Len(String1)
        If Mid$(String1, i, 1) <> s Then
          Exit For
        End If
      Next i
      NotInStr = i
    End If
  End If
  Call gTracer.xReturn("NotInStr")
End Function

'* GetDelimitedValue
'* Parses a string of the form  WWW,XXX  , YYYY, "ZZZ""ZZZ",,""
'* into 5 values WWW XXX YYYY ZZZ"ZZZ (empty string) (empty string)
'*
'* sValue               Value parsed using m_EscapeChar and m_Delimiter
'* Buffer               buffer to parse
'* offset               offset to begin parsing at
'*
'* return value
'* New offset in the buffer to parse
'* apf check
Public Function GetDelimitedValue(sValue As String, ByVal Buffer As String, ByVal offset As Long, Optional Delimiter As String = ",", Optional EscapeChar As String = """") As Long
  Dim lastspace As Long, cc As Long
  Dim q As String
  Dim bSkipSpace As Boolean ' skip leading and trailing space
  Dim bInEsc As Boolean
  
  sValue = "": bInEsc = False: bSkipSpace = True
  Do
    q = Mid$(Buffer, offset, 1)
    If Not (bSkipSpace And (q = " ")) Then
      bSkipSpace = False
      If q = EscapeChar Then
        If bInEsc And (Mid$(Buffer, offset + 1, 1) = EscapeChar) Then
          sValue = sValue & EscapeChar
          cc = cc + 1
          lastspace = cc
          offset = offset + 1
        Else
          bInEsc = Not bInEsc
          lastspace = cc
        End If
      ElseIf bInEsc Or (q <> Delimiter) Then
        sValue = sValue & q
        cc = cc + 1
        If bInEsc Or (q <> " ") Then lastspace = cc
      End If
    End If
    offset = offset + 1
  Loop While (Len(q) > 0) And (bInEsc Or (q <> Delimiter))
  If (lastspace > 0) And (lastspace < Len(sValue)) Then
    sValue = Left$(sValue, lastspace)
  End If
  GetDelimitedValue = offset
End Function


'* Returns a new string with the character char stripped from the left
'*
'* s                  source string
'* char               character to strip off, uses first character in string
'*
'* return value:
'* the trimmed string
'* if char is empty then the source string is returned
'*
'* note: see LTrim
Public Function LTrimChar(ByVal s As String, ByVal char As String) As String
  Dim c As String
  Dim i As Long
  
  Call gTracer.xSet("LTrimChar")
  If Len(s) < 1 Then GoTo LTrimChar_end
  If Len(char) < 1 Then
    LTrimChar = s
    GoTo LTrimChar_end
  End If
  
  c = Left$(char, 1)
  For i = 1 To Len(s)
    If c <> Mid$(s, i, 1) Then
      LTrimChar = Right$(s, Len(s) - i + 1)
      Exit For
    End If
  Next i
  
LTrimChar_end:
  Call gTracer.xReturn("LTrimChar")
End Function

'* Returns a new string with the character char stripped from the right
'*
'* s                  source string
'* char               character to strip off, uses first character in string
'*
'* return value:
'* the trimmed string
'* if char is empty then the source string is returned
Public Function RTrimChar(ByVal s As String, ByVal char As String) As String
  Dim c As String
  Dim i As Long
  
  Call gTracer.xSet("RTrimChar")
  If Len(s) < 1 Then GoTo RTrimChar_end
  If Len(char) < 1 Then
    RTrimChar = s
    GoTo RTrimChar_end
  End If
  
  c = Left$(char, 1)
  For i = Len(s) To 1 Step -1
    If Mid$(s, i, 1) <> c Then
      RTrimChar = Left$(s, i)
      Exit For
    End If
  Next i
  
RTrimChar_end:
  Call gTracer.xReturn("RTrimChar")
End Function

'* Returns a new string with the character char stripped from the left and right side
'* Similar to LTrim but does characters
'*
'* s                  source string
'* char               character to strip off, if length greater than one then take the left most character
'*
'* return value:
'* the trimmed string
'* if char is empty then the source string is returned
Public Function TrimChar(ByVal s As String, ByVal char As String) As String
  Call gTracer.xSet("TrimChar")
  TrimChar = LTrimChar(s, char)
  TrimChar = RTrimChar(TrimChar, char)
  Call gTracer.xReturn("TrimChar")
End Function

'apf fix
Public Function sReplaceChar(str As String, sFind, sReplace) As String
  Dim sStart As String
  Dim sEnd As String
  
  sReplaceChar = str
  Do While InStr(1, sReplaceChar, sFind) <> 0
    sStart = Left$(sReplaceChar, InStr(1, sReplaceChar, sFind) - 1)
    sEnd = Right$(sReplaceChar, Len(sReplaceChar) - InStr(1, sReplaceChar, sFind))
    sReplaceChar = sStart & sReplace & sEnd
  Loop
End Function

'* Split fullpath into directory (including backslash) and filename and extention (.ext)
'* returns true if could split
Public Sub SplitPath(ByVal sFullPath As String, Optional sDir As Variant, Optional sFile As Variant, Optional sExt As Variant)
  Dim p As Long, q As Long, tmp As Long
  
  On Error Resume Next
  Call gTracer.xSet("SplitPath")
  
  p = InStrBack(sFullPath, "\")
  If p = 0 Then p = InStrBack(sFullPath, ":")
  
  If (Not IsMissing(sDir)) And (p > 0) Then
    sDir = Left$(sFullPath, p - 1)
  End If
  
  q = InStrBack(sFullPath, ".")
  If Not IsMissing(sFile) Then
    p = p + 1
    If q > 0 Then
      tmp = q - p
    Else
      tmp = Len(sFullPath)
    End If
    sFile = Mid$(sFullPath, p, tmp)
  End If
  
  If Not IsMissing(sExt) And (q > 0) Then
    q = Len(sFullPath) - q + 1
    sExt = Right$(sFullPath, q)
  End If
  
SplitPath_end:
  Call gTracer.xReturn("SplitPath")
End Sub

Public Sub StringStub()
  Dim l As Long
  l = 5
End Sub


