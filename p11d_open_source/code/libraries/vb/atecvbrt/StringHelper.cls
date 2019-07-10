VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringHelper"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'##D Duplicate a single character in a string - use StrSQL for SQL strings
'##RV the string with the duplicated character
'##C STF
Public Function StrDupChar(ByVal String1 As String, ByVal CharDup As String) As String
  CharDup = Left$(CharDup, 1)
  StrDupChar = Replace(String1, CharDup, CharDup & CharDup, 1)
End Function

'##D Find a string in an array return offset of match
'##C STF
Public Function FindString(StrArray() As String, ByVal String1 As String, Optional ByVal NotFoundReturn As Long = -1) As Long
  Dim ub As Long, lb As Long, j As Long
  
  ub = UBound(StrArray)
  lb = LBound(StrArray)
  For j = lb To ub
    If StrComp(StrArray(j), String1, vbTextCompare) = 0 Then
      FindString = j
      Exit Function
    End If
  Next j
  FindString = NotFoundReturn
End Function


'##D Pad a string at front or back with a character to a specified length
'##V String1 original string
'##V Pad padding character
'##V Length length of padded string
'##V PadFront pad from front (default = false)
'##RV The padded string
'##C STF
Public Function xStrPad(ByVal String1 As String, ByVal Pad As String, ByVal Length As Long, Optional ByVal PadFront As Boolean = False) As String
  Dim padlen As Long
  Dim s As String
    
  padlen = Length - Len(String1)
  If padlen > 0 Then
    s = String$(padlen, Pad)
    If PadFront Then
      xStrPad = s & String1
    Else
      xStrPad = String1 & s
    End If
  Else
    xStrPad = Left$(String1, Length)
  End If
End Function

'##D Counts the occurences of a character / string in a string
'##V String1 string to be searched
'##V Char character to be found in string
'##V Compare Compare type
'##RV Number of occurences of that string
'##C STF
Public Function CountChar(String1 As String, ByVal Char As String, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As Long
  Dim p0 As Long, p1 As Long, ilen As Long

  ilen = Len(Char)
  If ilen > 0 Then
    p0 = 1
    Do
      p1 = InStr(p0, String1, Char, Compare)
      If p1 > 0 Then
        CountChar = CountChar + 1
        p0 = p1 + ilen
      End If
    Loop Until p1 = 0
  End If
End Function


'##D Construct a new string by using the insertion point within a string
'##V string1 string to be manipulated
'##V Insert insertion point in that string
'##V char ASCII character in question
'##LD Construct a new string by using the insertion point within a string
'##LD For use when typing to detect BACKSPACE, DEL etc
'##RV true if text inserted at end or False otherwise
'##C STF
Public Function ConstructString(String1 As String, ByVal Insert As Integer, ByVal Char As Integer) As Boolean
  Dim i As Integer

  String1 = ConstructString2(String1, ToUpper(Char), Insert, 0)
  ConstructString = ((Len(String1) - 1) = Insert)
End Function

'##C STF
Public Function ConstructString2(ByVal CurrentString As String, ByVal Char As Integer, ByVal SelStart As Integer, ByVal SelLength As Integer) As String
  Dim sLeftPart As String, sRightPart As String
    
  On Error Resume Next
  If Len(CurrentString) = 0 Then
    If IsPrint(Char) Then ConstructString2 = Chr$(Char)
    Exit Function
  End If
  sLeftPart = Mid$(CurrentString, 1, SelStart)
  sRightPart = Mid$(CurrentString, SelStart + SelLength + 1)
  If IsPrint(Char) Then
    If SelLength > 0 Then
      ConstructString2 = sLeftPart & Chr$(Char) & sRightPart
    Else
      If SelStart = 0 Then
        ConstructString2 = Chr$(Char) & CurrentString
      Else
        ConstructString2 = sLeftPart & Chr$(Char) & sRightPart
      End If
    End If
  ElseIf (Char = vbKeyBack) Or (Char = vbKeyDelete) Then
    If SelLength > 0 Then
      ConstructString2 = sLeftPart & sRightPart
    Else
      If SelStart > 0 Then
        ConstructString2 = Left$(sLeftPart, Len(sLeftPart) - 1) & sRightPart
      Else
        If Char = vbKeyDelete Then
          ConstructString2 = Mid$(CurrentString, 2)
        Else
          ConstructString2 = CurrentString
        End If
      End If
    End If
  End If
End Function

'##D remove multiple occurrances of characters in a string e.g "XX  X  Y" -> "X X Y"
'##V String1 String to be compressed
'##V Char Character to be removed when multiply occurring
'##RV Compressed String
'##C STF
Public Function CompressString(ByVal String1 As String, ByVal Char As String) As String
  Dim p0 As Long, p1 As Long
  Dim q0 As Long
  Dim s As String
  
  If Len(Char) = 0 Then
    s = String1
  Else
    p0 = 1
    Do
      p1 = InStr(p0, String1, Char)
      If p1 <> 0 Then
        s = s & Mid$(String1, p0, p1 - p0 + Len(Char))
        p0 = p1 + Len(Char)
        Do
          q0 = InStr(p0, String1, Char)
          If q0 = p0 Then
            p0 = q0 + Len(Char)
          Else
            Exit Do
          End If
        Loop Until False
      Else
        s = s & Mid$(String1, p0)
      End If
    Loop Until p1 = 0
  End If
  CompressString = s
End Function

'##D Adds Escape Characters to String
'##GCF
Public Function AddEscapeChars(ByVal s As String) As String
  If (InStr(1, s, ",") > 0) Or (InStr(1, s, """") > 0) Then
    If Not ((StrComp(Left$(s, 1), """") = 0) And (StrComp(Mid$(s, Len(s), 1), """") = 0)) Then
      s = """" & Replace(s, """", """""") & """"
    End If
  End If
  AddEscapeChars = s
End Function

'##C STF
Public Function ReplaceString(String1 As String, ByVal sFind As String, ByVal sReplace As String, Optional ByVal Compare As VbCompareMethod = vbTextCompare) As String
  ReplaceString = Replace(String1, sFind, sReplace, , , Compare)
End Function

'##D Replaces character in string with another
'##C STF
Public Function ReplaceChar(String1 As String, ByVal cFind As String, ByVal cReplace As String) As String
  ReplaceChar = Replace(String1, cFind, cReplace, , , vbTextCompare)
End Function

'##D Removes character from string
'##C STF
Public Function RemoveChar(String1 As String, ByVal cFind As String) As String
  RemoveChar = Replace(String1, cFind, "", , , vbTextCompare)
End Function

'##D Returns a numbered field from a string of fields seperated by tokens
'##V String1            source string
'##V Token             seperator token
'##V TokenIndex         the index of the field to be extracted (1=Field1, 2=Field2)
'##V TrimString Trim spaces off the field (default = false)
'##RV the extracted string
'##C STF
Public Function GetStringToken(String1 As String, ByVal Token As String, Optional ByVal TokenIndex As Long = 1, Optional ByVal TrimString As Boolean = False) As String
  Dim pos As Long
  Dim lpos As Long
  Dim i As Long
  
  ' sort out apf tokenize XXX  vbTab YYY
  If Len(String1) = 0 Then GoTo GetStringToken_End
  
  pos = 0: i = 0
  Do
    lpos = pos
    pos = InStr(pos + 1, String1, Token, vbTextCompare)
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
  
  If TrimString Then GetStringToken = Trim$(GetStringToken)
  
GetStringToken_End:
End Function

'##D Finds the first occurence of a character in String1 which is not in String2
'##V String1 source string
'##V String2 ignore character
'##V Start start point (Default = 1)
'##V Compare compare method (Default = vbBinaryCompare)
'##RV the position of the first character which is not String1 or if String2 is not in String1 then Len(String1)+1
Public Function NotInStr(String1 As String, ByVal String2 As String, Optional ByVal Start As Long = 1, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As Long
  NotInStr = NotInStrAny(String1, String2, Start, Compare)
End Function

'##D Parses a string of the form  WWW,XXX  , YYYY, "ZZZ""ZZZ",,"" into 5 values WWW XXX YYYY ZZZ"ZZZ (empty string) (empty string)
'##V DelimitedValue               Value parsed using m_EscapeChar and m_Delimiter
'##V DelimitedString buffer to parse
'##V offset               offset to begin parsing at
'##RV New offset in the buffer to parse, if > len(Buffer) then stop parsing string
'##C STF
Public Function GetDelimitedValue(DelimitedValue As String, DelimitedString As String, ByVal offset As Long, Optional ByVal Delimiter As String = ",", Optional ByVal EscapeChar As String = """") As Long
  GetDelimitedValue = GetDelimitedValueEx(DelimitedValue, DelimitedString, offset, True, Delimiter, EscapeChar)
End Function

'##D Returns number of delimited items in an array of strings
'##LD ValueArray is an uninitialised array of strings - Will be initialised as 1-based
'##LD DelimitedString - Obvious
'##LD Delimiter - Again obvious. Default is Comma
'##LD IgnoreBlank - Determines whether blank elements are stored. Default is False
'##LD Returns Number of Delimited Items, Zero on Error
'##LD apf IgnoreBlanks !!!
'##C STF
Public Function GetDelimitedValues(ValueArray As Variant, DelimitedString As String, Optional IgnoreBlank As Boolean = False, Optional ByVal TrimValues As Boolean = True, Optional ByVal Delimiter As String = ",", Optional ByVal EscapeChar As String = """") As Long
  GetDelimitedValues = GetDelimitedValuesEx(ValueArray, DelimitedString, IgnoreBlank, TrimValues, Delimiter, EscapeChar)
End Function
Public Function GetDelimitedValues2(ValueArray As Variant, DelimitedString As String, Optional IgnoreBlank As Boolean = False, Optional ByVal TrimValues As Boolean = True, Optional ByVal Delimiter As String = ",", Optional ByVal EscapeChar As String = """", Optional IncludeTrailing As Boolean = False) As Long
  GetDelimitedValues2 = GetDelimitedValuesEx(ValueArray, DelimitedString, IgnoreBlank, TrimValues, Delimiter, EscapeChar, IncludeTrailing)
End Function

'##C STF
Public Function GetDelimitedValueEx(DelimitedValue As String, DelimitedString As String, ByVal offset As Long, Optional ByVal TrimValues As Boolean = True, Optional ByVal Delimiter As String = ",", Optional ByVal EscapeChar As String = """") As Long
  Dim InSkipSpace As Boolean ' skip leading and trailing space
  Dim InESC As Boolean
  Dim AscEscapeChar As Integer, p As Integer
  Dim q As String, lastspace As Long, lenbuf As Long
  
  lenbuf = Len(DelimitedString)
  If offset < 1 Then Err.Raise ERR_INVALIDOFFSET, "GetDelimitedValueEx", "Offset must be in range 1 to Buffer Length"
  
  AscEscapeChar = Asc(EscapeChar)
  DelimitedValue = "": InESC = False
  lastspace = 0: InSkipSpace = TrimValues
  Do While offset <= lenbuf
    q = Mid$(DelimitedString, offset, 1)
    p = Asc(q)
    If Not (InSkipSpace And (p = 32)) Then  ' 32 = space
      InSkipSpace = False
      If p = AscEscapeChar Then
        If InESC And (Mid$(DelimitedString, offset + 1, 1) = EscapeChar) Then
          DelimitedValue = DelimitedValue & EscapeChar
          offset = offset + 1
        Else
          InESC = Not InESC
        End If
      ElseIf InESC Or (InStr(1, Delimiter, q, vbTextCompare) = 0) Then
        DelimitedValue = DelimitedValue & q
      Else
        Exit Do
      End If
      If (Not InESC) And (p = 32) Then
        If lastspace = 0 Then lastspace = Len(DelimitedValue) - 1
      Else
        lastspace = 0
      End If
    End If
    offset = offset + 1
  Loop
  If TrimValues And (lastspace > 0) And (lastspace < Len(DelimitedValue)) Then
    DelimitedValue = Left$(DelimitedValue, lastspace)
  End If
  GetDelimitedValueEx = offset + 1
End Function

Public Function GetDelimitedValuesEx(ValueArray As Variant, DelimitedString As String, Optional IgnoreBlank As Boolean = False, Optional ByVal TrimValues As Boolean = True, Optional ByVal Delimiter As String = ",", Optional ByVal EscapeChar As String = """", Optional ByVal IncludeTrailing As Boolean = False) As Long
  Const ARRAY_INCREMENT As Long = 64
  Dim offset As Long
  Dim tmp As String
  Dim ArrayCount As Long, ArrayMax As Long
  
  offset = 1
  Do While offset <= Len(DelimitedString)
    offset = GetDelimitedValueEx(tmp, DelimitedString, offset, TrimValues, Delimiter, EscapeChar)
    If Not (IgnoreBlank And (Len(tmp) = 0)) Then
      If ArrayCount = ArrayMax Then
        If ArrayMax = 0 Then
          ReDim ValueArray(1 To ARRAY_INCREMENT)
        Else
          ReDim Preserve ValueArray(1 To ArrayMax + ARRAY_INCREMENT)
        End If
        ArrayMax = ArrayMax + ARRAY_INCREMENT
      End If
      ArrayCount = ArrayCount + 1
      ValueArray(ArrayCount) = tmp
    End If
  Loop
  'RK 27/05/04 Allow for empty column at end of string
  'RK 27/05/04 BEGIN
  If IncludeTrailing Then
    If StrComp(Right(DelimitedString, 1), Delimiter, vbTextCompare) = 0 Then
      ArrayCount = ArrayCount + 1
      ValueArray(ArrayCount) = ""
    End If
  End If
  'RK 27/05/04 END
  If (Len(DelimitedString) = 0) And Not IgnoreBlank Then
    ArrayCount = 1
    ReDim ValueArray(1 To ArrayCount)
  Else
    If ArrayCount > 0 Then ReDim Preserve ValueArray(1 To ArrayCount)
  End If
  GetDelimitedValuesEx = ArrayCount
End Function
