Attribute VB_Name = "Strings"
Option Explicit

Public Function FullPathEx(ByVal sPath As String) As String
  If Len(sPath) > 0 Then
    FullPathEx = sPath
    If StrComp(right$(sPath, 1), "\", vbBinaryCompare) <> 0 Then
      FullPathEx = FullPathEx & "\"
    End If
  End If
End Function

'Public Function ReplaceCharEx(String1 As String, ByVal FindChar As String, ByVal ReplaceChar As String, ByVal Compare As VbCompareMethod) As String
'  Dim i As Long, rlen As Long, offset As Long
'
'  ReplaceCharEx = String1
'  rlen = Len(ReplaceChar)
'  offset = Len(FindChar)
'  If offset > 0 Then
'    i = 1:
'    Do
'      i = InStr(i, ReplaceCharEx, FindChar, Compare)
'      If i <> 0 Then
'        If (rlen = 1) And (offset = 1) Then
'          Mid$(ReplaceCharEx, i, 1) = ReplaceChar
'        Else
'          ReplaceCharEx = Left$(ReplaceCharEx, i - 1) & ReplaceChar & Mid$(ReplaceCharEx, i + offset)
'        End If
'        If rlen > 0 Then i = i + rlen
'      End If
'    Loop Until (i = 0) Or (i > Len(ReplaceCharEx))
'  End If
'End Function

' Split based on a delimiter ( and character in Delimiter )
Public Function GetDelimitedValueInt(DelimitedValue As String, buffer As String, ByVal offset As Long, ByVal TrimSpaces As Boolean, ByVal Delimiter As String, ByVal EscapeChar As String) As Long
  Dim InSkipSpace As Boolean ' skip leading and trailing space
  Dim InESC As Boolean
  Dim AscEscapeChar As Integer, p As Integer
  Dim q As String, lastspace As Long, lenbuf As Long
  
  lenbuf = Len(buffer)
  If offset < 1 Then Call Err.Raise(ERR_INVALIDOFFSET, "GetDelimitedValue", "Offset must be in range 1 to Buffer Length")
  
  AscEscapeChar = Asc(EscapeChar)
  DelimitedValue = "": InESC = False
  lastspace = 0: InSkipSpace = TrimSpaces
  Do While offset <= lenbuf
    q = Mid$(buffer, offset, 1)
    p = Asc(q)
    If Not (InSkipSpace And (p = 32)) Then  ' 32 = space
      InSkipSpace = False
      If p = AscEscapeChar Then
        If InESC And (Mid$(buffer, offset + 1, 1) = EscapeChar) Then
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
  If TrimSpaces And (lastspace > 0) And (lastspace < Len(DelimitedValue)) Then
    DelimitedValue = left$(DelimitedValue, lastspace)
  End If
  GetDelimitedValueInt = offset + 1
End Function

Public Function StrDupCharEx(String1 As String, CharDup As String) As String
  Dim q0 As Long
    
  StrDupCharEx = String1
  CharDup = left$(CharDup, 1)
  If (Len(CharDup) > 0) And (Len(StrDupCharEx) > 0) Then
    q0 = InStr(StrDupCharEx, CharDup)
    Do While q0 > 0
      StrDupCharEx = left$(StrDupCharEx, q0) & CharDup & Mid$(StrDupCharEx, q0 + 1)
      q0 = q0 + 2
      q0 = InStr(q0, StrDupCharEx, CharDup)
    Loop
  End If
End Function
  
Public Function DateStringEx2(ByVal v As Variant, ByVal DefaultValue As String, ByVal FormatString As String) As String
  
  On Error GoTo DateStringEx2_err
  If VarType(v) <> vbDate Then v = ConvertDateEx(v, CONVERT_DELIMITED, "DMY", "/", ":")
  If v = UNDATED Then
    DateStringEx2 = DefaultValue
  Else
    DateStringEx2 = Format$(v, FormatString)
  End If
DateStringEx2_end:
  Exit Function
  
DateStringEx2_err:
  DateStringEx2 = DefaultValue
  Resume DateStringEx2_end
End Function

Public Function ConstructStringEx(ByVal CurrentString As String, ByVal Char As Integer, ByVal SelStart As Integer, ByVal SelLength As Integer) As String
  Dim sLeftPart As String, sRightPart As String
    
  On Error Resume Next
  If Len(CurrentString) = 0 Then
    If IsPrint(Char) Then ConstructStringEx = Chr$(Char)
    Exit Function
  End If
  sLeftPart = Mid$(CurrentString, 1, SelStart)
  sRightPart = Mid$(CurrentString, SelStart + SelLength + 1)
  If IsPrint(Char) Then
    If SelLength > 0 Then
      ConstructStringEx = sLeftPart & Chr$(Char) & sRightPart
    Else
      If SelStart = 0 Then
        ConstructStringEx = Chr$(Char) & CurrentString
      Else
        ConstructStringEx = sLeftPart & Chr$(Char) & sRightPart
      End If
    End If
  ElseIf (Char = vbKeyBack) Or (Char = vbKeyDelete) Then
    If SelLength > 0 Then
      ConstructStringEx = sLeftPart & sRightPart
    Else
      If SelStart > 0 Then
        ConstructStringEx = left$(sLeftPart, Len(sLeftPart) - 1) & sRightPart
      Else
        If Char = vbKeyDelete Then
          ConstructStringEx = Mid$(CurrentString, 2)
        Else
          ConstructStringEx = CurrentString
        End If
      End If
    End If
  End If
End Function

Public Function ConvertDateEx(ByVal DateString As String, ByVal ConvType As DATECONVERT_TYPE, ByVal ConvStr As String, ByVal DateDelimit As String, ByVal TimeDelimit As String) As Date
  Dim i As Long, j As Long, p0 As Long, p1 As Long, ch As String
  Dim nyear As Integer, nmonth As Integer, nday As Integer
  Dim nhour As Integer, nminute As Integer, nsecond As Integer
  Dim d0 As Date

  On Error GoTo ConvertDateEx_err
  If (ConvType = CONVERT_FIXEDDATE) Or (ConvType = CONVERT_FIXEDDATETIME) Then
    i = InStr(1, ConvStr, "Y", vbTextCompare)
    j = InStrRev(ConvStr, "Y", , vbTextCompare)
    If (j - i) < 1 Then Err.Raise ERR_CONVERTDATE, "ConvertDate", "Format not complete: " & ConvStr & vbCr & "Year value must allow at least 2 digits"
    
    nyear = CLng(Mid$(DateString, i, j - i + 1))
    nyear = GetFullYear_CD(nyear, ConvStr)
        
    i = InStr(1, ConvStr, "D", vbTextCompare)
    j = InStrRev(ConvStr, "D", , vbTextCompare)
    If (i = 0) And (j = 0) Then
      nday = 1&
    Else
      If (j - i) < 1 Then Err.Raise ERR_CONVERTDATE, "ConvertDate", "Format not complete: " & ConvStr & vbCr & "Day value must allow 2 digits"
      nday = CLng(Mid$(DateString, i, j - i + 1))
    End If
  
    i = InStr(1, ConvStr, "M", vbTextCompare)
    j = InStrRev(ConvStr, "M", , vbTextCompare)
    If (j - i) < 1 Then Err.Raise ERR_CONVERTDATE, "ConvertDate", "Format not complete: " & ConvStr & vbCr & "Month value must allow 2 digits"
    nmonth = CLng(Mid$(DateString, i, j - i + 1))
    If ConvType = CONVERT_FIXEDDATETIME Then
      i = InStr(1, ConvStr, "H", vbTextCompare)
      j = InStrRev(ConvStr, "H", , vbTextCompare)
      If (j - i) < 1 Then Err.Raise ERR_CONVERTDATE, "ConvertDate", "Format not complete: " & ConvStr & vbCr & "Hour value must allow 2 digits"
      nhour = CLng(Mid$(DateString, i, j - i + 1))
     
      i = InStr(1, ConvStr, "N", vbTextCompare)
      j = InStrRev(ConvStr, "N", , vbTextCompare)
      If (j - i) < 1 Then Err.Raise ERR_CONVERTDATE, "ConvertDate", "Format not complete: " & ConvStr & vbCr & "Minute value must allow 2 digits"
      nminute = CLng(Mid$(DateString, i, j - i + 1))
           
      i = InStr(1, ConvStr, "S", vbTextCompare)
      j = InStrRev(ConvStr, "S", , vbTextCompare)
      If (j - i) >= 1 Then nsecond = CLng(Mid$(DateString, i, j - i + 1))
    End If
  ElseIf ConvType = CONVERT_DELIMITED Then
    i = 1: p0 = 1: p1 = 1
    nday = 1
    For i = 1 To Len(ConvStr)
      ch = UCase$(Mid$(ConvStr, i, 1))
      If InStr("HNS", ch) = 0 Then
        p1 = InStr(p0, DateString, DateDelimit, vbTextCompare)
      Else
        p1 = InStr(p0, DateString, TimeDelimit, vbTextCompare)
      End If
      If p1 = 0 Then p1 = InStr(p0, DateString, " ", vbTextCompare)
      If p1 = 0 Then
        If i = Len(ConvStr) Then p1 = Len(DateString) + 1
        If (p1 - p0) < 1 Then Err.Raise ERR_CONVERTDATE, "ConvertDate", "Format not complete: " & ConvStr
      End If
      Select Case ch
        Case "D"
            nday = CLng(Mid$(DateString, p0, p1 - p0))
        Case "M"
            nmonth = CLng(Mid$(DateString, p0, p1 - p0))
        Case "Y"
            nyear = CLng(Mid$(DateString, p0, p1 - p0))
            nyear = GetFullYear_CD(nyear, ConvStr)
        Case "H"
            nhour = CLng(Mid$(DateString, p0, p1 - p0))
        Case "N"
            nminute = CLng(Mid$(DateString, p0, p1 - p0))
        Case "S"
            nsecond = CLng(Mid$(DateString, p0, p1 - p0))
        Case Else
            Err.Raise ERR_CONVERTDATE, "ConvertDate", "Format contains invalid charactere: " & ConvStr & vbCr & "Date must have a Year, Month and Day order - DMY and HNS for Hours, minutes and seconds"
      End Select
      p0 = p1 + 1
    Next i
  Else
    If Not IsDate(DateString) Then Err.Raise ERR_CONVERTDATE, "ConvertDate", "Unknown Date format"
    d0 = CDate(DateString)
    nday = DatePart("d", d0)
    nmonth = DatePart("m", d0)
    nyear = DatePart("yyyy", d0)
    nhour = DatePart("h", d0)
    nminute = DatePart("n", d0)
    nsecond = DatePart("s", d0)
  End If
  d0 = DateSerialEx(nyear, nmonth, nday) + TimeSerial(nhour, nminute, nsecond)
  If (Day(d0) <> nday) Or (Month(d0) <> nmonth) Or (Year(d0) <> nyear) Then Err.Raise ERR_CONVERTDATE, "ConvertDate", "Failed to convert date. Converted to " & Format$(d0, "DD/MM/YYYY") & " (DD/MM/YYYY)"
  ConvertDateEx = d0
  Exit Function
  
ConvertDateEx_err:
  Err.Raise Err.Number, ErrorSourceEx(Err, "ConvertDateEx"), "Unable to convert string " & DateString & " to a date" & vbCrLf & Err.Description
End Function

Public Function GetDelimitedValuesEx(ValueArray As Variant, DelimitedString As String, ByVal IgnoreBlank As Boolean, ByVal TrimValues As Boolean, ByVal Delimiter As String, ByVal EscapeChar As String) As Long
  Const ARRAY_INCREMENT As Long = 64
  Dim offset As Long
  Dim tmp As String
  Dim ArrayCount As Long, ArrayMax As Long
  
  offset = 1
  Do While offset <= Len(DelimitedString)
    offset = GetDelimitedValueInt(tmp, DelimitedString, offset, TrimValues, Delimiter, EscapeChar)
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
  If (Len(DelimitedString) = 0) And Not IgnoreBlank Then
    ArrayCount = 1
    ReDim ValueArray(1 To ArrayCount)
  Else
    If ArrayCount > 0 Then ReDim Preserve ValueArray(1 To ArrayCount)
  End If
  GetDelimitedValuesEx = ArrayCount
End Function

Public Function GetIniKeyNamesExInternal(KeyNames As Variant, ByVal SectionName As String, ByVal IniFilePath As String) As Long
  Dim TempKeys() As String, nRet As Long
  
  nRet = GetIniKeyNamesInt(TempKeys, SectionName, IniFilePath)
  If nRet > 0 Then
    KeyNames = TempKeys
    GetIniKeyNamesExInternal = nRet
  End If
End Function

Public Function StrSQLEx(ByVal vString As Variant) As String
  If IsNull(vString) Then
    StrSQLEx = "Null"
  Else
    StrSQLEx = "'" & StrDupCharEx(CStr(vString), "'") & "'"
  End If
End Function


Public Function CompressStringEx(ByVal String1 As String, ByVal Char As String) As String
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
  CompressStringEx = s
End Function

