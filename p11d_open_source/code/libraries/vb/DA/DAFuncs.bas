Attribute VB_Name = "DAFuncs"
Option Explicit
Private dProgress As Double

Public Const dbQTable = 1024
Private mDaRSid As Long

'to be removed
Public Sub logerr(ByVal ErrText As String)
  Dim eType As errnumbersenum
  
  'If gDebugMode Then
    eType = ERR_ERROR + ERR_ALLOWIGNORE
  'Else
  '  eType = ERR_ERRORSILENT
  'End If
  
  Call ErrorMessage(eType, Err, "DADLL Error", "Standard DADLL Error", ErrText)


End Sub
Public Sub logmessage(ms As String)
  #If DEBUGVER Then
    If Not InErrorMessage Then Call ErrorMessage(ERR_INFOSILENT, Nothing, "Logmessage", "Logmessage", ms)
  #End If
End Sub

Public Sub AddList(ByVal text As String, List As Variant)
  Dim i As Long
  Dim errset As Boolean
  
  On Error GoTo AddList_err
  
  text = Trim$(text)
  If Len(text) = 0 Then Exit Sub
  If Not IsArrayInitialised(List) Then ReDim List(0)
  If Not InList(text, List) Then
    ReDim Preserve List(UBound(List) + 1)
    List(UBound(List)) = text
  End If
    
AddList_end:
  Exit Sub
AddList_err:
  Err.Raise ERR_ADD_LIST, "AddList", "An error occured adding '" & text & "' to a list." & vbCrLf & Err.Description
  Resume AddList_end
End Sub

Public Sub RemoveList(text As String, List As Variant)
  Dim i As Long
  Dim item As Long
  
  On Error GoTo RemoveList_Err
  
  text = Trim$(text)
  
  For i = LBound(List) To UBound(List)
    If StrComp(List(i), text, vbTextCompare) = 0 Then
      If Not (item = UBound(List)) Then
        List(item) = List(UBound(List))
      End If
      ReDim Preserve List(UBound(List) - 1)
      GoTo RemoveList_End
    End If
  Next i
  
RemoveList_End:
  Exit Sub
RemoveList_Err:
  Err.Raise ERR_REMOVE_LIST, "RemoveList", "An error occured removing '" & text & "' from a list." & vbCrLf & Err.Description
End Sub

Public Function InList(val As String, List As Variant) As Boolean
  Dim i As Long
  
  InList = False
  If IsArrayInitialised(List) Then
    For i = LBound(List) To UBound(List)
      If StrComp(val, List(i), vbTextCompare) = 0 Then
        InList = True
        Exit Function
      End If
    Next i
  End If
End Function

Public Function NextID() As String
  NextID = "_" + CStr(mDaRSid)
  mDaRSid = mDaRSid + 1
End Function

Public Function RightPart(text As String, pos As Long) As String
  RightPart = ""
  If IsNull(text) Then Exit Function
  If pos > Len(text) Then Exit Function
  RightPart = Right$(text, Len(text) - pos + 1)
End Function
Public Function InstrSQL(Startpos As Long, SearchString As String, SeekString As String, Compare As VbCompareMethod) As Long
' seekstring must not be inside quotes
Dim i As Long

  i = InStr(Startpos, SearchString, SeekString, Compare)
  If i Then
    Do While InQuotes(SearchString, i)
      i = InStr(i + 1, SearchString, SeekString, Compare)
      If i = 0 Then Exit Do
    Loop
  End If
  InstrSQL = i
  
End Function
Public Function InStrCollection(coll As Collection, s As String) As Boolean
Dim t As String

  On Error Resume Next
  t = coll(s)
  InStrCollection = (StrComp(t, s, vbTextCompare) = 0)
  
End Function

'Public Function GetNextPairSeparatedByString(ByVal SQLstr As String, ByVal Startpos As Long, sepstr() As String, endstr() As String, LeftofSep As String, RightofSep As String) As Long
'  Dim TempSepStartPos As Long
'  Dim SepStartPos As Long
'  Dim SepEndPos As Long
'  Dim i As Integer
'  Dim TempEndpos As Long
'  Dim Endpos As Long
'
'  Dim bStop As Boolean
'
'  On Error GoTo GetNextPairSeparatedByString_err:
'
'  GetNextPairSeparatedByString = 0: LeftofSep = "": RightofSep = ""
'  TempEndpos = 0
'  Endpos = Len(SQLstr)
'  ' establish position of next separator
'  For i = LBound(sepstr) To UBound(sepstr)
'    If Len(sepstr(i)) Then
'      ' should replace chr$(..) with " " throughout string before
'      TempSepStartPos = IIf(Len(sepstr(i)) - 1, InstrSQL(Startpos, SQLstr, sepstr(i), vbTextCompare), InStr(Startpos, SQLstr, sepstr(i), vbTextCompare))
'      If TempSepStartPos > 0 And SepStartPos > TempSepStartPos Then
'        SepStartPos = TempSepStartPos
'        SepEndPos = SepStartPos + Len(sepstr(i))
'      End If
'    End If
'  Next i
'  If SepStartPos = 0 Then GoTo GetNextPairSeparatedByString_end
'
'  ' establish position of next end beyond next separator
'  For i = LBound(endstr) To UBound(endstr)
'    If Len(endstr(i)) Then
'      TempEndpos = IIf(Len(endstr(i)) - 1, InstrSQL(SepEndPos, SQLstr, endstr(i), vbTextCompare), InStr(SepEndPos, SQLstr, endstr(i), vbTextCompare))
'      If TempEndpos > 0 Then Endpos = min(Endpos, TempEndpos)
'    End If
'  Next i
'  If SepEndPos >= Endpos Then bStop = True
'
'  If bStop Then
'    GetNextPairSeparatedByString = 0
'  Else
'    GetNextPairSeparatedByString = Endpos
'    LeftofSep = Trim(Mid$(SQLstr, Startpos + 1, SepStartPos - Startpos - 1))
'    RightofSep = Trim(Mid$(SQLstr, SepEndPos + 1, Endpos - SepEndPos - 1))
'  End If
'
'GetNextPairSeparatedByString_end:
'  Exit Function
'GetNextPairSeparatedByString_err:
'  Err.Raise ERR_DAPARSE, "GetNextPairSeparatedByString", "An error occurred parsing the string:" & vbCrLf & SQLstr & vbCrLf & Err.Description
'End Function
Public Function GetNextItemSeparatedByString(ByVal SQLstr As String, ByVal Startpos As Long, sepstr() As String, NextItem As String, sepItem As Long) As Long
  Dim i As Integer
  Dim TempEndpos As Long
  Dim Endpos As Long
  Dim NextStartpos  As Long
  Dim bStop As Boolean
  
  On Error GoTo GetNextItemSeparatedByString_err:
  
  GetNextItemSeparatedByString = 0: NextItem = ""
  TempEndpos = 0
  Endpos = Len(SQLstr)
  sepItem = 0
  ' establish position of next separator
  For i = LBound(sepstr) To UBound(sepstr)
    If Len(sepstr(i)) Then
      TempEndpos = InstrSQL(Startpos, SQLstr, sepstr(i), vbTextCompare)
      If TempEndpos > 0 And TempEndpos < Endpos Then
        Endpos = min(Endpos, TempEndpos)
        NextStartpos = TempEndpos + Len(sepstr(i))
        sepItem = i
      End If
    End If
  Next i
  
  GetNextItemSeparatedByString = NextStartpos
  
  If Endpos = Startpos Then GetNextItemSeparatedByString = 0
  If sepItem > 0 Then
    NextItem = Trim(Mid$(SQLstr, Startpos, Endpos - Startpos + IIf(Len(sepstr(sepItem)) > 1, 1, 0)))
  Else
    NextItem = Trim(Mid$(SQLstr, Startpos, Endpos - Startpos + 1))
  End If
  
  NextItem = RemoveChar(NextItem, "(")
  NextItem = RemoveChar(NextItem, ")")
  NextItem = RemoveChar(NextItem, "[")
  NextItem = RemoveChar(NextItem, "]")
  
GetNextItemSeparatedByString_end:
  Exit Function
GetNextItemSeparatedByString_err:
Resume
  Err.Raise ERR_DAPARSE, "GetNextItemSeparatedByString", "An error occurred parsing the string:" & vbCrLf & SQLstr & vbCrLf & Err.Description
End Function

Public Function SplitNextEqualsPair(ByVal sqltext As String, ByVal Startpos As Long, LeftofEquals As String, RightofEquals As String) As Long
  Dim equalsPos As Long
  Dim commaPos As Long
  Dim wherePos As Long
  Dim bStop As Boolean
  Dim IgnorePos As Long
  Dim IgnorePosStart As Long
  Dim IgnorePosEnd As Long
  Dim IgnoreCount As Long
  
  On Error GoTo SplitNextEqualsPair_err:
  
  SplitNextEqualsPair = 0: LeftofEquals = "": RightofEquals = ""
  
  ' must have = before a quote
  equalsPos = InStr(Startpos, sqltext, "=")
  If equalsPos = 0 Then GoTo SplitNextEqualsPair_end
  ' must ignore ,'s inside ' , ", ()
    
  commaPos = InStr(equalsPos, sqltext, ",", vbTextCompare)
  IgnorePos = InStr(equalsPos, sqltext, "'", vbTextCompare)
  If IgnorePos < commaPos And IgnorePos > 0 Then
    IgnorePos = InStr(IgnorePos + 1, sqltext, "'", vbTextCompare)
    commaPos = InStr(IgnorePos, sqltext, ",", vbTextCompare)
  End If
  IgnorePos = InStr(equalsPos, sqltext, Chr$(34), vbTextCompare)
  If IgnorePos < commaPos And IgnorePos > 0 Then
    IgnorePos = InStr(IgnorePos + 1, sqltext, Chr$(34), vbTextCompare)
    commaPos = InStr(IgnorePos, sqltext, ",", vbTextCompare)
  End If
  ' ignore's the horrible possibility of '(' and similar ...
  IgnoreCount = 0
  IgnorePosStart = InStr(equalsPos, sqltext, "(", vbTextCompare)
  If IgnorePosStart < commaPos And IgnorePosStart > 0 Then
    IgnoreCount = 1
    Do While IgnoreCount > 0
      IgnorePosEnd = InStr(IgnorePosStart + 1, sqltext, ")", vbTextCompare)
      IgnorePos = InStr(IgnorePosStart + 1, sqltext, "(", vbTextCompare)
      If (IgnorePosEnd < IgnorePos Or IgnorePos = 0) Then
        IgnoreCount = IgnoreCount - 1
        IgnorePosStart = IgnorePosEnd
      Else
        IgnoreCount = IgnoreCount + 1
        IgnorePosStart = IgnorePos
      End If
    Loop
    
    commaPos = InStr(IgnorePosEnd, sqltext, ",", vbTextCompare)
  End If
  
  
  'If commaPos = 0 Then
    wherePos = InStr(equalsPos, sqltext, "WHERE ", vbTextCompare)
    If wherePos > 0 Then
      If equalsPos < wherePos And (commaPos = 0 Or commaPos > wherePos) Then
        commaPos = wherePos
        bStop = True
      End If
    End If
  'End If
  
  If commaPos = 0 Then commaPos = InStr(equalsPos, sqltext, ";")
  If commaPos = 0 Then commaPos = InStr(equalsPos, sqltext, Chr$(0))
  If commaPos = 0 Then commaPos = Len(sqltext) + 1
  LeftofEquals = Trim(Mid$(sqltext, Startpos + 1, equalsPos - Startpos - 1))
  RightofEquals = Trim(Mid$(sqltext, equalsPos + 1, commaPos - equalsPos - 1))
  If bStop Then
    SplitNextEqualsPair = 0
  Else
    SplitNextEqualsPair = commaPos
  End If
    
SplitNextEqualsPair_end:
  Exit Function
SplitNextEqualsPair_err:
  Err.Raise ERR_DAPARSE, "SplitNextEqualsPair", "An error occurred parsing the string:" & vbCrLf & sqltext & vbCrLf & Err.Description
End Function

Public Function FieldName(ByVal text As String) As String
  Dim pos As Long
  
  On Error GoTo FieldName_Err
  
  pos = InStr(text, ".")
  If pos = 0 Then
    FieldName = text
  Else
    FieldName = RightPart(text, pos + 1)
  End If
  
  If Left$(FieldName, 1) = "[" Then
    FieldName = Mid$(FieldName, 2, Len(FieldName) - 2)
  End If
  
FieldName_End:
  Exit Function
  
FieldName_Err:
  FieldName = ""
  Err.Raise ERR_FIELD_NAME, "Field Name", "Error parsing the Field Name from the string:" & vbCrLf & text & vbCrLf & Err.Description
End Function

Public Function TableName(text As String) As String
  Dim pos As Long
  
  On Error GoTo TableName_Err
  
  pos = InStr(text, ".")
  If pos = 0 Then
     TableName = ""
  Else
    TableName = Trim(Left$(text, pos - 1))
  End If
  
  If Left$(TableName, 1) = "[" Then
    TableName = Mid$(TableName, 2, Len(TableName) - 2)
  End If

TableName_End:
  Exit Function

TableName_Err:
  TableName = ""
  Err.Raise ERR_TABLE_NAME, "Table Name", "Error parsing the Table Name from the string:" & vbCrLf & text & vbCrLf & Err.Description
End Function

Public Function xLng(val As Variant) As Long
  If IsNull(val) Then
    xLng = 0
  Else
    xLng = CLng(val)
  End If
End Function

Private Function IsArrayInitialised(arr As Variant) As Boolean
  Dim i As Long
  On Error GoTo IsArrayInitialised_err
  IsArrayInitialised = False
  i = LBound(arr)
  IsArrayInitialised = True
IsArrayInitialised_end:
  Exit Function
IsArrayInitialised_err:
  Resume IsArrayInitialised_end
End Function
Public Function InQuotes(s As String, pos As Long) As Boolean

  InQuotes = InSeparator(s, pos, Chr$(34)) Or InSeparator(s, pos, Chr$(39))
  
  'sdfghfj
End Function
Private Function InSeparator(s As String, pos As Long, sep As String) As Boolean
Dim instart As Long
Dim inend As Long
Dim b As Boolean
  
    
  b = False
  instart = InStr(1, s, sep, vbTextCompare)
  If instart Then
    Do While pos > instart And instart <> 0
      inend = InStr(instart + 1, s, sep, vbTextCompare)
      If inend = 0 Then
        ' non matching pairs
        ' raise an error
        Err.Raise ERR_SQL_PARSE, "Inseparator", "Invalid sql found with unmatched quotes, parsing " & vbCr & vbCr & s
      End If
      If pos >= instart And pos <= inend Then
        b = True
        Exit Do
      End If
      instart = InStr(inend + 1, s, sep, vbTextCompare)
    Loop
  End If
  InSeparator = b
End Function


Public Function ShowProgress(sMessage As String, sTitle As String, Optional dPortion As Double = 0) As Boolean
Dim i As Long

  On Error GoTo ShowProgress_Err

  If gShowPopMessages Then
    Call DisplayMessagePopup(Nothing, sMessage, sTitle)
  End If
  If Not gStatusBar Is Nothing And dPortion <> 0 Then
    If Not gStatusBar.StatusBar Is Nothing And (gStatusBar.StatMax > gStatusBar.StatMin) Then
      dProgress = dProgress + dPortion
      i = (gStatusBar.StatMax - gStatusBar.StatMin) * dProgress
      If CInt(i) <> 0 Then
        gStatusBar.StatusBar.SetStatus min(gStatusBar.StatMin + i, gStatusBar.StatMax)
      End If
    End If
  End If

ShowProgress_End:
  Exit Function

ShowProgress_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ShowProgress", "Error in ShowProgress", "Undefined error.")
  Resume ShowProgress_End
End Function

Public Sub ResetProgress()
  dProgress = 0
End Sub
