VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Parser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements ILibraryVersion

Public Enum PARSETYPE
  PARSETYPE_INVALID = 0
  PARSETYPE_TOKEN
  PARSETYPE_LINE
End Enum

Private Enum PARSEMODE
  TOKEN_SEARCH
  TOKEN_OPENREAD
  TOKEN_READ
  TOKEN_READPARAM
End Enum

Private m_ParseItems As ObjectHash
Private m_ParseSettings As IParseSettings
Private m_ParsePreProcess As IParsePreProcess
Private m_CurrentLine As Long  ' used only for single line parsing

Private Const MAXPARAMS As Long = 30
Private m_Parameters() As String

Private m_TokensOnly As Boolean
Private m_InEscape As Boolean

Private m_parsestatus As PARSESTATUS ' 0 ok, 1 abort current line, 2 abort parsing

' Parser specific constants
Private m_Parsetype As PARSETYPE
Private Const DELIMITER As String = ","
Private Const ESCCHAR As String = """"
Private Const TOKENPARAM As String = "="
Private TOKENOPEN As String
Private TOKENCLOSE As String
Private m_PostParse As String ' part of line after last parsed token

Private Sub Class_Initialize()
  Set m_ParseItems = New ObjectHash
  Set m_ParseSettings = New DefaultParseSettings
  Call RegisterLibrary(Me)
  ReDim m_Parameters(1 To MAXPARAMS) As String
End Sub

Private Sub Class_Terminate()
 Set m_ParseItems = Nothing
 Set m_ParseSettings = Nothing
 Set m_ParsePreProcess = Nothing
End Sub


Property Get InEscape() As Boolean
  InEscape = m_InEscape
End Property

Property Let InEscape(ByVal NewValue As Boolean)
  m_InEscape = NewValue
End Property

Public Sub AddParseItem(pItem As IParseItem)
    
  On Error GoTo AddParseItem_err
  Call m_ParseItems.Add(pItem, pItem.ActionStr)
AddParseItem_end:
  Exit Sub
  
AddParseItem_err:
  Call ErrorMessage(ERR_ERROR, Err, "AddParseItem", "ERR_ADDPARSEITEM", "Error adding Parse Item " & pItem.ActionStr)
  Resume AddParseItem_end
End Sub

Public Property Set ParseSettings(NewValue As IParseSettings)
  TOKENOPEN = ""
  TOKENCLOSE = ""
  m_Parsetype = PARSETYPE_INVALID
  Set m_ParseSettings = NewValue
  If Not m_ParseSettings Is Nothing Then
    If m_ParseSettings.ParserType = PARSETYPE_INVALID Then Err.Raise ERR_SETPARSESETTINGS, "ParseSettings", "Invalid Parse type set for ParseSettings"
    If m_ParseSettings.ParserType = PARSETYPE_LINE Then
      If (Len(m_ParseSettings.TOKENOPEN) > 0) Or Len(m_ParseSettings.TOKENCLOSE) > 0 Then
        Err.Raise ERR_SETPARSESETTINGS, "ParseSettings", "Line Parser cannot use token delimiters"
      End If
    End If
    If m_ParseSettings.ParserType = PARSETYPE_TOKEN Then
      If (Len(m_ParseSettings.TOKENOPEN) <> 1) Or (Len(m_ParseSettings.TOKENCLOSE) <> 1) Then
        Err.Raise ERR_SETPARSESETTINGS, "ParseSettings", "Token Parser delimiters must be 1 character long"
      End If
      If StrComp(m_ParseSettings.TOKENOPEN, m_ParseSettings.TOKENCLOSE, vbTextCompare) = 0 Then
        Err.Raise ERR_SETPARSESETTINGS, "ParseSettings", "Token Parser open and close delimiters cannot be the same character"
      End If
    End If
    TOKENOPEN = m_ParseSettings.TOKENOPEN
    TOKENCLOSE = m_ParseSettings.TOKENCLOSE
    m_Parsetype = m_ParseSettings.ParserType
  End If
  Call ParseReset
End Property

Public Property Get ParseSettings()
  Set ParseSettings = m_ParseSettings
End Property

Public Property Set ParsePreProcess(NewValue As IParsePreProcess)
  Set m_ParsePreProcess = NewValue
End Property

Public Property Let ParseTokensOnly(ByVal NewValue As Boolean)
  m_TokensOnly = NewValue
End Property

Public Function ParseLine(ByVal sline As String) As PARSESTATUS
  m_CurrentLine = m_CurrentLine + 1
  ParseLine = ParseLine_Int(sline, m_CurrentLine)
End Function

Private Function ParseLine_Int(sline As String, ByVal LineNum As Long) As PARSESTATUS
  Dim curParseObj As Variant
  Dim errorflag As Long, tlen As Long
  
  On Error GoTo ParseLine_Int_err
  Call xSet("ParseLine")
  errorflag = 0
  m_PostParse = ""
  If m_parsestatus = PARSESTATUS_ABORT Then GoTo ParseLine_Int_noparse
  m_parsestatus = PARSESTATUS_OK
  tlen = Len(LTrim$(sline))
  If (tlen > 0) And (Not m_ParsePreProcess Is Nothing) Then Call m_ParsePreProcess.PreProcessParseLine(sline, m_parsestatus, LineNum)
  If tlen = 0 Then
    m_PostParse = sline
    GoTo ParseLine_Int_noparse
  End If
    
  Set curParseObj = m_ParseSettings.StartParseLine
  If m_Parsetype = PARSETYPE_TOKEN Then
    Call ParseTokenLine(curParseObj, sline)  ' can change parsestatus
  ElseIf m_Parsetype = PARSETYPE_LINE Then
    Call ParseSingleLine(curParseObj, sline) ' can change parsestatus
  Else
    Call ECASE("ParseLine - invalid Parsetype:" & CStr(m_Parsetype))
    m_parsestatus = PARSESTATUS_ABORT
  End If
    
ParseLine_Int_end:
  Call m_ParseSettings.EndParseLine(curParseObj, m_parsestatus)
  
ParseLine_Int_noparse:
  ParseLine_Int = m_parsestatus
  Call xReturn("ParseLine")
  Exit Function
  
ParseLine_Int_err:
  m_parsestatus = PARSESTATUS_ABORTLINE
  Call ErrorMessage(ERR_ERROR + ERR_ALLOWIGNORE, Err, "ParseLine", "Error Parsing Line", "Unable to parse Line " & sline)
  errorflag = errorflag + 1
  If errorflag > 1 Then
    Resume ParseLine_Int_noparse
  Else
    Resume ParseLine_Int_end
  End If
End Function

Public Function ParseFile(FileName As String) As PARSESTATUS
  Dim iFileNum  As Integer
  Dim sline As String, sline2 As String ' allow <cr> between tokens
  Dim LineNum As Long
  
  On Error GoTo ParseFile_err
  Call xSet("ParseFile")
  iFileNum = -1
  m_parsestatus = PARSESTATUS_OK
  If m_ParseSettings Is Nothing Then Call Err.Raise(ERR_PARSESETTINGS, "ParseFile", "No ParseSettings defined")
  If Not FileExists(FileName) Then Call Err.Raise(ERR_PARSESETTINGS, "ParseFile", "File " & FileName & " does not exist")
  If m_parsestatus = PARSESTATUS_ABORT Then GoTo ParseFile_end
  
  iFileNum = FreeFile
  Open FileName For Input As iFileNum
  Do While (m_parsestatus <> PARSESTATUS_ABORT) And (Not EOF(iFileNum))
    Line Input #1, sline
    LineNum = LineNum + 1
    If m_Parsetype = PARSETYPE_LINE Then
      sline = Trim$(sline)
      If IsTokenLine(sline) Then
        m_parsestatus = ParseLine_Int(sline2, LineNum - 1)
        sline2 = sline
      ElseIf Len(sline) > 0 Then
        sline2 = sline2 & " " & sline
      End If
    Else
      m_parsestatus = ParseLine_Int(sline, LineNum)
    End If
  Loop
  If Len(sline2) > 0 Then m_parsestatus = ParseLine_Int(sline2, LineNum)
  
ParseFile_end:
  ParseFile = m_parsestatus
  If iFileNum > 0 Then Close #iFileNum
  Call xReturn("ParseFile")
  Exit Function

ParseFile_err:
  m_parsestatus = PARSESTATUS_ABORT
  Call ErrorMessage(ERR_ERROR + ERR_ALLOWIGNORE, Err, "ParseFile", "ERR_PARSEFILE", "Unable to parse file " & FileName)
  Resume ParseFile_end
End Function

Private Sub ParseTokenLine(curParseObj As Variant, sline As String)
  Dim linelen As Long, pToken As IParseItem
  Dim pre As String, token As String
  Dim mode As PARSEMODE, curparam As Long
  Dim nextOpen As Long, nextClose As Long, nextParam As Long, minStop As Long
  Dim nextOpen2 As Long, p As Long

  On Error GoTo ParseTokenLine_err
  Call xSet("ParseTokenLine")
  ' p - current position
  mode = TOKEN_SEARCH
  p = 1: nextOpen2 = -1: linelen = Len(sline)
  Do While p <= linelen
    If m_parsestatus >= PARSESTATUS_ABORTLINE Then GoTo ParseTokenLine_end
    curparam = 0: mode = TOKEN_SEARCH
    If (nextOpen2 = 0) Or (nextOpen2 >= p) Then
      nextOpen = nextOpen2
    Else
      nextOpen = InStr(p, sline, TOKENOPEN, vbBinaryCompare)
    End If
    If nextOpen = 0 Then
      m_PostParse = Mid$(sline, p)
      If m_TokensOnly And (Len(m_PostParse) > 0) Then Err.Raise ERR_PARSETOKEN, "ParseTokenLine", "Line contains unparsed tokens or text" & vbCrLf & "Current parse setting only allow valid tokens"
      GoTo ParseTokenLine_end
    End If
    mode = TOKEN_READ
    nextClose = InStr(nextOpen + 1, sline, TOKENCLOSE, vbBinaryCompare)
    If nextClose = 0 Then
      m_PostParse = Mid$(sline, p)
      GoTo ParseTokenLine_end
    End If
    nextParam = InStr(nextOpen + 1, sline, TOKENPARAM, vbBinaryCompare)
    minStop = nextClose
    If (nextParam > 0) And (nextParam < nextClose) Then minStop = nextParam
    nextOpen2 = InStr(nextOpen + 1, sline, TOKENOPEN, vbBinaryCompare)
    Do While nextOpen2 < minStop
      If nextOpen2 = 0 Then Exit Do
      nextOpen = nextOpen2
      nextOpen2 = InStr(nextOpen + 1, sline, TOKENOPEN, vbBinaryCompare)
    Loop
    token = Mid$(sline, nextOpen + 1, minStop - 1 - nextOpen)
    pre = Mid$(sline, p, nextOpen - p)
    Set pToken = m_ParseItems.Item(token, False)
    If (pToken Is Nothing) And m_InEscape Then GoTo skip_token
    p = minStop + 1
    If (Not pToken Is Nothing) And (minStop = nextParam) Then
      mode = TOKEN_READPARAM
      p = ReadParams(curparam, sline, p, TOKENCLOSE)
      p = p + 1
    End If
    m_parsestatus = DoAction(curParseObj, pre, token, pToken, m_Parameters, curparam)
skip_token:
  Loop

ParseTokenLine_end:
  Call xReturn("ParseTokenLine")
  Exit Sub

ParseTokenLine_err:
  m_parsestatus = PARSESTATUS_ABORTLINE
  Call ErrorMessage(ERR_ERROR + ERR_ALLOWIGNORE, Err, "ParseTokenLine", "Error parsing tokenized line", "Error parsing Line " & sline)
  Resume ParseTokenLine_end
End Sub

Private Sub ParseSingleLine(curParseObj As Variant, sline As String)
  Dim token As String, pToken As IParseItem, p As Long
  Dim curparam As Long, EscapeStatus As Boolean, s As String

  On Error GoTo ParseSingleLine_err
  Call xSet("ParseSingleLine")
  p = InStr(1, sline, TOKENPARAM, vbBinaryCompare)
  If p > 0 Then
    token = Trim$(Left$(sline, p - 1))
    Set pToken = m_ParseItems.Item(token, False)
  End If
  If Not pToken Is Nothing Then
    sline = sline & Chr$(0)
    p = ReadParams(curparam, sline, p + 1, Chr$(0), EscapeStatus)
  End If
  If p = 0 Then
    If EscapeStatus Then s = "Currently in Escape sequence"
    Err.Raise ERR_PARSEMODE, "ParseSingleLine", s
  End If
  m_parsestatus = DoAction(curParseObj, "", token, pToken, m_Parameters, curparam)
  
ParseSingleLine_end:
  Call xReturn("ParseSingleLine")
  Exit Sub
  
ParseSingleLine_err:
  m_parsestatus = PARSESTATUS_ABORTLINE
  Call ErrorMessage(ERR_ERROR + ERR_ALLOWIGNORE, Err, "ParseSingleLine", "Error parsing complete line", "Error parsing Line " & sline)
  Resume ParseSingleLine_end
End Sub

Public Property Get PostParseLine() As String
  PostParseLine = m_PostParse
End Property

Public Sub ParseReset()
  m_PostParse = ""
  m_parsestatus = PARSESTATUS_INVALID
  m_CurrentLine = 0
End Sub

Private Function DoAction(curParseObj As Variant, pre As String, token As String, ByVal pToken As IParseItem, params() As String, ByVal paramcount As Long) As PARSESTATUS
  Dim i As Long
  
  On Error GoTo DoAction_err
  Call xSet("DoAction")
  If pToken Is Nothing Then Set pToken = m_ParseItems.Item(token, False)
  If pToken Is Nothing Then Err.Raise ERR_PARSETOKEN, "DoAction", "Unable to find ParseItem for token " & token
    
  If (pToken.ActionParamCount >= 0) And (pToken.ActionParamCount <> paramcount) Then
    Err.Raise ERR_TOKENPARAMETERS, "DoAction", "Expected " & CStr(pToken.ActionParamCount) & " parameters found " & CStr(paramcount) & " parsing token " & token
  End If
  DoAction = pToken.OnAction(m_ParseSettings, curParseObj, pre, paramcount, params)

DoAction_end:
  Call xReturn("DoAction")
  Exit Function

DoAction_err:
  Dim s As String
  DoAction = PARSESTATUS_ABORTSTATEMENT
  s = TOKENOPEN & token & TOKENPARAM
  For i = 1 To paramcount
    s = s & params(i) & DELIMITER
  Next i
  s = Left$(s, Len(s) - 1) & TOKENCLOSE
  Call ErrorMessage(ERR_ERROR + ERR_ALLOWIGNORE, Err, "DoAction", "Error executing token action", "Error parsing token " & s)
  Resume DoAction_end
  Resume
End Function

Private Function IsTokenLine(String1 As String) As Boolean
  Dim p As Long
  
  p = InStr(1, String1, TOKENPARAM, vbBinaryCompare)
  If p > 0 Then IsTokenLine = (Not m_ParseItems.Item(Left$(String1, p - 1), False) Is Nothing)
End Function

'Private Function GetParseItemIndex(token As String) As Long
'  Dim i As Long, tlen As Long, thash As Long
'  #If DEBUGVER Then
'    Dim uCount As Long
'    Dim s As String
'  #End If
'
'  token = UCase$(token)
'  tlen = Len(token)
'  thash = GetHashValue(token)
'  For i = 1 To m_ParseItems.Count
'    If tlen = m_ParseItemsInfo(i).len Then
'      If thash = m_ParseItemsInfo(i).hash Then
'        #If DEBUGVER Then
'          uCount = uCount + 1
'        #End If
'        If StrComp(m_ParseItemsInfo(i).ItemStr, token, vbBinaryCompare) = 0 Then
'          GetParseItemIndex = i: Exit For
'        End If
'      End If
'    End If
'  Next i
'  #If DEBUGVER Then
'    If uCount > 1 Then
'      For i = 1 To m_ParseItems.Count
'        If tlen = m_ParseItemsInfo(i).len Then
'          If thash = m_ParseItemsHash(i) Then
'            s = vbCrLf & "token: " & m_ParseItemsStr(i) & s
'          End If
'        End If
'      Next i
'      Call ECASE("Parse Item Hash clash for tokens " & s)
'    End If
'  #End If
'End Function

Private Function ReadParams(curparam As Long, String1 As String, ByVal offset As Long, ByVal EndToken As String, Optional InEscape As Boolean) As Long
  Dim ch As String, nextch As String, i As Long, inesc As Boolean, linelen As Long
  Dim PrevNonBlankChar As String, skipnextch As Boolean

  Call xSet("ReadParams")
  inesc = False
  curparam = 1
  m_Parameters(curparam) = ""
  linelen = Len(String1)
  PrevNonBlankChar = TOKENPARAM
  nextch = Mid$(String1, offset, 1)
  For i = offset To linelen
    If m_parsestatus >= PARSESTATUS_ABORTLINE Then GoTo ReadParams_end
      
    ch = nextch
    nextch = Mid$(String1, i + 1, 1)
    If skipnextch Then
      skipnextch = False
      GoTo skipchar
    End If
    
    If ch = EndToken And Not inesc Then
      ReadParams = i
      Exit For
    ElseIf ch = DELIMITER And Not inesc Then
      curparam = curparam + 1
      m_Parameters(curparam) = ""
    ElseIf ch = ESCCHAR Then
      If Not inesc Then
        If (PrevNonBlankChar = DELIMITER) Or (PrevNonBlankChar = TOKENPARAM) Then
          inesc = True
        Else
          m_Parameters(curparam) = m_Parameters(curparam) & ch
        End If
      Else ' currently in escape sequence
        If nextch = ESCCHAR Then
          m_Parameters(curparam) = m_Parameters(curparam) & ch
          skipnextch = True
        Else
          inesc = False
        End If
      End If
    Else
      m_Parameters(curparam) = m_Parameters(curparam) & ch
    End If
skipchar:
    If (Len(ch) > 0) And (ch <> " ") Then PrevNonBlankChar = ch
  Next i
ReadParams_end:
 InEscape = inesc
 Call xReturn("ReadParams")
End Function

Private Property Get ILibraryVersion_Name() As String
  ILibraryVersion_Name = "Parser library"
End Property

Private Property Get ILibraryVersion_Version() As String
  ILibraryVersion_Version = App.Major & "." & App.Minor & "." & App.Revision
End Property
