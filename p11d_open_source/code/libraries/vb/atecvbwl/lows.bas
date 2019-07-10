Attribute VB_Name = "lows"
Option Explicit
'ALl EX FUNCTIONS ARE MAPPED TO AN EXTERNAL CLASS
Sub Main()
  'init variables
End Sub
Public Function AddRSField(ByVal fld As field, ByVal Name As String, ByVal FieldName As String, ByVal AsNode As Boolean) As NavigatorRSField
  Set AddRSField = New NavigatorRSField
  Call AddRSField.Init(fld, Name, FieldName, AsNode)
End Function

Public Function HTMLAttribEx(ByRef sAttribName As String, ByVal vValue As Variant, ByVal EscapeQuotes As Boolean) As String
  If VarType(vValue) = vbBoolean Then
    vValue = CLng(vValue)
  Else
    vValue = IsNullEx(vValue, "")
    If EscapeQuotes Then
      vValue = Replace(vValue, S_QUOT, "&quot;")
    Else
      If InStr(vValue, S_QUOT) > 0 Then Err.Raise ERR_HTMLATTR, "HTMLAttrib", "Attribute [" & vValue & "] cannot contain embedded quotes"
    End If
  End If
  If g_Debugging Then
    If StrComp(LCase$(sAttribName), sAttribName, vbBinaryCompare) <> 0 Then Err.Raise ERR_HTMLATTR, "HTMLAttribEx", "Attribute [" & sAttribName & "] should be lower case"
    If InStr(sAttribName, "=") > 0 Then Err.Raise ERR_HTMLATTR, "HTMLAttribEx", "Attribute [" & sAttribName & "] should not contain an equals sign"
  End If
  HTMLAttribEx = sAttribName & "=" & S_QUOT & vValue & S_QUOT & " "
End Function

'cadxx added AllowNoClose  - breaks compatability
Public Function ElementOpenEx(Name As String, Attributes As String, ByVal ShortClose As Boolean, Optional ByVal AllowNoClose As Boolean = False) As String
  Dim s As String

  s = "<" & LCase$(Name)
  Attributes = Trim$(Attributes)
  If Len(Attributes) > 0 Then s = s & " "
  s = s & Attributes

  If ShortClose Then
    s = s & "/"
  End If
  s = s & ">" & g_CRLF
  If g_Debugging And (Not ShortClose) And Not AllowNoClose Then
    ElementOpenEx = String$((g_Indent) * 2, " ") & s
    If g_Debugging Then g_Indent = g_Indent + 1
  Else
    ElementOpenEx = s
  End If
End Function
Public Function ElementCloseEx(Name As String) As String
  Dim s As String
  s = "</" & LCase$(Name) & ">" & g_CRLF
  If g_Debugging Then
    g_Indent = g_Indent - 1
    ElementCloseEx = String$(g_Indent * 2, " ") & s
  Else
    ElementCloseEx = s
  End If
  
End Function
Public Function FileToString(ByVal FileName As String) As String
  Dim fr As FileRead
  
  On Error GoTo FileToString_err
  Set fr = New FileRead
  If Not fr.OpenFile(FileName) Then Err.Raise ERR_FILETOSTRING, "FileToString", "Unable to open file '" & FileName & "'"
  Call fr.GetFile(FileToString)
  Set fr = Nothing
  Exit Function
  
FileToString_err:
  Err.Raise Err.Number, ErrorSourceComponentEx(Err, "FileToString", COMPONENT_NAME), Err.Description
End Function

Public Function ErrorSourceComponentEx(ByVal Err As ErrObject, FunctionName As String, AppExeName As String) As String
  ErrorSourceComponentEx = ErrorSource(Err, AppExeName & "." & FunctionName)
End Function

' This now deals with "hello]]>bye" as "<![CDATA[hello]]]><![CDATA[]>bye"
Public Function XMLTextEx(ByVal s As String) As String
  XMLTextEx = XMLTextRef(s)
End Function
Public Function XMLTextRef(ByRef s As String) As String
  Dim p0 As Long, p1 As Long
  
  If g_Debugging Then XMLTextRef = String$((g_Indent * 2), " ")
  If InStrAny(s, S_INVALID_XML_CHARS, 1, vbBinaryCompare) > 0 Then
    p0 = 1
    p1 = InStr(p0, s, CDATA_END)
    Do While p1 > 0
      XMLTextRef = XMLTextRef & CDATA_START & Mid$(s, p0, p1 - p0 + 1) & CDATA_END
      p0 = p1 + 1
      p1 = InStr(p0, s, CDATA_END)
    Loop
    XMLTextRef = XMLTextRef & CDATA_START & Mid$(s, p0) & CDATA_END & g_CRLF
  Else
    XMLTextRef = s
  End If
End Function

Private Function ReplaceXMLMetacharactersEx(ByVal sText As String) As String
  ReplaceXMLMetacharactersEx = ReplaceXMLMetacharactersRef(sText)
End Function

Private Function ReplaceXMLMetacharactersRef(ByRef sText As String) As String
  If InStrAny(sText, S_INVALID_XML_CHARS, 1, vbBinaryCompare) > 0 Then
    sText = Replace(sText, "&", "&amp;")
    sText = Replace(sText, "'", "&apos;")
    sText = Replace(sText, ">", "&gt;")
    sText = Replace(sText, "<", "&lt;")
    sText = Replace(sText, S_QUOT, "&quot;")
  End If
  ReplaceXMLMetacharactersRef = sText
End Function

' MPS, Updated AF 12/4/2005
Public Function ReplaceHTMLMetacharactersEx(ByVal sText As String, Optional ByVal replaceSpaces As Boolean = True) As String
  ReplaceHTMLMetacharactersEx = ReplaceHTMLMetacharactersRef(sText, replaceSpaces)
End Function

Public Function ReplaceHTMLMetacharactersRef(ByRef sText As String, ByVal replaceSpaces As Boolean) As String
  If replaceSpaces Then
    If InStrAny(sText, "<>""' ", 1, vbBinaryCompare) > 0 Then
      sText = Replace(sText, "<", "&lt;")
      sText = Replace(sText, ">", "&gt;")
      sText = Replace(sText, "'", "&#39;")
      sText = Replace(sText, """", "&quot;")
      sText = Replace(sText, " ", "&nbsp;")
    End If
  Else
    If InStrAny(sText, "<>""'", 1, vbBinaryCompare) > 0 Then
      sText = Replace(sText, "<", "&lt;")
      sText = Replace(sText, ">", "&gt;")
      sText = Replace(sText, "'", "&#39;")
      sText = Replace(sText, """", "&quot;")
    End If
  End If
  ReplaceHTMLMetacharactersRef = sText
End Function

'Required as xml attributes are case-sensitive
Public Function XMLAttribEx2(ByRef attrName As String, ByVal vValue As Variant) As String
  If VarType(vValue) = vbBoolean Then
    vValue = CLng(vValue)
  Else
    vValue = ReplaceXMLMetacharactersEx(IsNullEx(vValue, ""))
  End If
  XMLAttribEx2 = attrName & "=" & S_QUOT & vValue & S_QUOT & " "
End Function

Public Function XMLAttribEx(ByRef attrName As String, ByVal attrValue As String, ByVal add_trailing_space As Boolean, ByVal tolower_attrname As Boolean) As String
  attrValue = ReplaceXMLMetacharactersRef(attrValue)
  If tolower_attrname Then attrName = LCase$(attrName)
  XMLAttribEx = attrName & "=" & S_QUOT & attrValue & S_QUOT & IIf(add_trailing_space, " ", "")
End Function

'cadxx lower case attributes
Public Function HTMLCheckBoxEx(Optional ByVal id As String, Optional ByVal Checked As Boolean = False, Optional ByVal Disabled As Boolean = False, Optional ByVal Attributes As String) As String

  Dim sattributes As String
  On Error GoTo HTMLCheckBoxEx_err

  sattributes = Attributes & HTMLAttribEx("type", "checkbox", False)
  If Disabled Then sattributes = sattributes & HTMLAttribEx("disabled", "YES", False)
  If Checked Then sattributes = sattributes & HTMLAttribEx("checked", "YES", False)
  If Len(id) > 0 Then sattributes = sattributes & HTMLAttribEx("id", id, False) & HTMLAttribEx("name", id, False)
  HTMLCheckBoxEx = ElementOpenEx("INPUT", sattributes, True)

HTMLCheckBoxEx_end:
  Exit Function

HTMLCheckBoxEx_err:
 Err.Raise Err.Number, ErrorSourceComponentEx(Err, "HTMLCheckBoxEx", COMPONENT_NAME), Err.Description
End Function


'cadxx 'xml compatible
Public Function HTMLRadioButtonEx(Optional ByVal idGroup As String, Optional ByVal id As String, Optional ByVal Checked As Boolean = False, Optional ByVal Disabled As Boolean = False, Optional ByVal Attributes As String) As String
  'by ms
  Dim sattributes As String
  On Error GoTo HTMLRadioButtonEx_err

'cad changed to be xml compatible
  sattributes = Attributes & HTMLAttribEx("type", "radio", False)
  If Disabled Then sattributes = sattributes & HTMLAttribEx("disabled", "YES", False)
  If Checked Then sattributes = sattributes & HTMLAttribEx("checked", "YES", False)
  If Len(id) > 0 Then sattributes = sattributes & HTMLAttribEx("id", id, False) & HTMLAttribEx("name", idGroup, False)
  HTMLRadioButtonEx = ElementOpenEx("input", sattributes, True)

HTMLRadioButtonEx_end:
  Exit Function

HTMLRadioButtonEx_err:
 Err.Raise Err.Number, ErrorSourceComponentEx(Err, "HTMLRadioButtonEx", COMPONENT_NAME), Err.Description
 Resume
End Function

Public Function HTMLTextBoxEx(Optional ByVal id As String, Optional ByVal Value As String = "", Optional ByVal Disabled As Boolean = False, Optional ByVal Attributes As String) As String
  Dim sattributes As String
  On Error GoTo HTMLTextBoxEx_err

  sattributes = Attributes & HTMLAttribEx("type", "text", False)
  sattributes = sattributes & HTMLAttribEx("value", Value, False)
  If Disabled Then sattributes = sattributes & HTMLAttribEx("disabled", "YES", False)
  If Len(id) > 0 Then sattributes = sattributes & HTMLAttribEx("id", id, False) & HTMLAttribEx("name", id, False)
  HTMLTextBoxEx = ElementOpenEx("input", sattributes, True)

HTMLTextBoxEx_end:
  Exit Function

HTMLTextBoxEx_err:
 Err.Raise Err.Number, ErrorSourceComponentEx(Err, "HTMLTextBoxEx", COMPONENT_NAME), Err.Description
End Function

Public Function HTMLTextAreaEx(Optional ByVal id As String, Optional ByVal Rows As Long = 1, Optional ByVal Value As String = "", Optional ByVal Disabled As Boolean = False, Optional ByVal Attributes As String) As String
  
  Dim sattributes As String
  On Error GoTo HTMLTextAreaEx_err
  
  sattributes = sattributes & HTMLAttribEx("ROWS", CStr(Rows), False)
  If Disabled Then sattributes = sattributes & HTMLAttribEx("DISABLED", "YES", False)
  If Len(id) > 0 Then sattributes = sattributes & HTMLAttribEx("ID", id, False) & HTMLAttribEx("name", id, False)
  HTMLTextAreaEx = ElementOpenEx("TEXTAREA", sattributes, False) & Value & ElementCloseEx("TEXTAREA")
  
HTMLTextAreaEx_end:
  Exit Function

HTMLTextAreaEx_err:
 Err.Raise Err.Number, ErrorSourceComponentEx(Err, "HTMLTextAreaEx", COMPONENT_NAME), Err.Description
End Function

'cadxx xml compat
Public Function HTMLHiddenInputEx(Optional ByVal id As String, Optional ByVal Value As String = "", Optional ByVal Attributes As String) As String
  Dim sattributes As String
  On Error GoTo HTMLHiddenInputEx_err

  sattributes = Attributes & HTMLAttribEx("type", "hidden", False)
  sattributes = sattributes & HTMLAttribEx("value", Value, False)
  If Len(id) > 0 Then sattributes = sattributes & HTMLAttribEx("id", id, False) & HTMLAttribEx("name", id, False)
  HTMLHiddenInputEx = ElementOpenEx("input", sattributes, True)

HTMLHiddenInputEx_end:
  Exit Function

HTMLHiddenInputEx_err:
 Err.Raise Err.Number, ErrorSourceComponentEx(Err, "HTMLHiddenInputEx", COMPONENT_NAME), Err.Description
End Function

'cadxx xml compat
Public Function HTMLListBoxEx(Optional ByVal id As String, Optional ByVal ValueList As Variant, Optional ByVal ValueListDisplay As Variant, Optional ByVal Value As String = "", Optional ByVal Disabled As Boolean = False, Optional ByVal Attributes As String) As String
  Dim sattributes As String, sAttributesOption As String
  Dim OptionValue As String, OptionValueDisplay As String
  Dim s As String
  Dim i As Long
  On Error GoTo HTMLListBoxEx_err

  If IsMissing(ValueListDisplay) Or IsEmpty(ValueListDisplay) Then ValueListDisplay = ValueList

  sattributes = Attributes
  If Disabled Then sattributes = sattributes & HTMLAttribEx("disabled", "YES", False)
  If Len(id) > 0 Then sattributes = sattributes & HTMLAttribEx("id", id, False) & HTMLAttribEx("name", id, False)
  s = ElementOpenEx("select", sattributes, False)
  For i = LBound(ValueList) To UBound(ValueList)
    OptionValue = ValueList(i)
    OptionValueDisplay = ValueListDisplay(i)
    sAttributesOption = HTMLAttribEx("value", OptionValue, False)
    If Value = OptionValue Then sAttributesOption = sAttributesOption & HTMLAttribEx("selected", "YES", False)
    s = s & ElementOpenEx("option", sAttributesOption, True) & OptionValueDisplay
  Next i
  s = s & ElementCloseEx("select")

  HTMLListBoxEx = s

HTMLListBoxEx_end:
  Exit Function

HTMLListBoxEx_err:
 Err.Raise Err.Number, ErrorSourceComponentEx(Err, "HTMLListBoxEx", COMPONENT_NAME), Err.Description
End Function

'cadxx xml compat
Public Function HTMLButtonEx(Optional ByVal id As String, Optional ByVal Value As String = "", Optional ByVal Disabled As Boolean = False, Optional ByVal OnClickEvent As String, Optional ByVal Attributes As String) As String
  Dim sattributes As String
  On Error GoTo HTMLButtonEx_err

  sattributes = Attributes & HTMLAttribEx("type", "button", False)
  sattributes = sattributes & HTMLAttribEx("value", Value, False)
  sattributes = sattributes & HTMLAttribEx("onclick", OnClickEvent, False)
  If Disabled Then sattributes = sattributes & HTMLAttribEx("disabled", "YES", False)
  If Len(id) > 0 Then sattributes = sattributes & HTMLAttribEx("id", id, False) & HTMLAttribEx("name", id, False)
  HTMLButtonEx = ElementOpenEx("input", sattributes, True)

HTMLButtonEx_end:
  Exit Function

HTMLButtonEx_err:
 Err.Raise Err.Number, ErrorSourceComponentEx(Err, "HTMLButtonEx", COMPONENT_NAME), Err.Description
End Function

Public Sub DownLoadEx(ByVal Response As Response, Data As Variant, ByVal ContentType As String, ByVal FileName As String)
  Dim surl As String
  
  On Error GoTo DownLoadEx_ERR
    
  'Response.Buffer = True
  
  'Response.AddHeader "Content-Location", "fred.htm"
  Response.AddHeader "content-disposition", "attachment; filename=""" & FileName & """"
  Response.ContentType = ContentType
  Response.BinaryWrite (Data)
  
DownLoadEx_END:
  Exit Sub
DownLoadEx_ERR:
  Call Err.Raise(Err.Number, ErrorSourceComponentEx(Err, "DownLoadEx", COMPONENT_NAME), Err.Description)
End Sub

Public Sub BasicAuthenticationUserDetailsEX(ByRef UserName As String, ByRef Password As String, ByVal Request As Request)
  Dim sAuthorzation As String, s As String
  Dim iLen As Long
  Dim p0 As Long
  
  On Error GoTo ERR_ERR
    
  UserName = ""
  Password = ""
  sAuthorzation = Request.ServerVariables("HTTP_AUTHORIZATION")
  iLen = Len(sAuthorzation)
  If iLen = 0 Or Left$(sAuthorzation, 5) <> "Basic" Then GoTo ERR_END
  s = Replace(sAuthorzation, "Basic ", "")
  s = Base64Decode(s)
  iLen = Len(s)
  If StrComp(s, ":") = 0 Then GoTo ERR_END
  p0 = InStr(1, s, ":")
  UserName = Left$(s, p0 - 1)
  If p0 < iLen Then Password = Mid$(s, p0 + 1)
  
ERR_END:
  Exit Sub
ERR_ERR:
  Call Err.Raise(Err.Number, ErrorSourceComponentEx(Err, "BasicAuthenticationUserDetailsEX", COMPONENT_NAME), Err.Description)
  Resume
End Sub

Public Sub BasicAuthenticationInitEX(ByVal Response As Response, ByVal Realm As String, FailureHTML As String)
  
  On Error GoTo ERR_ERR
  
  Call Response.Clear
  Response.Buffer = True
  Response.Status = "401 Unauthorized"
  Call Response.AddHeader("WWW-Authenticate", "Basic realm=""" & Realm & """")
  Call Response.Write(FailureHTML)
    
ERR_END:
  Exit Sub
ERR_ERR:
  Call Err.Raise(Err.Number, ErrorSourceComponentEx(Err, "BasicAuthenticationInitEX", COMPONENT_NAME), Err.Description)
  Resume
End Sub

'cadxx new
Public Function IncludeStyleEx(ByVal WebGlobals As WebGlobals, ByVal StyleFileNoPath As String, Optional IncludeFile As Boolean = True, Optional RaiseErrorsOnMissingFile As Boolean = True) As String
  Dim sFile As String, s As String
  Dim src As String

  On Error GoTo ERR_ERR
  If IncludeFile Then
    src = HTMLAttribEx("href", WebGlobals.CSSDir & StyleFileNoPath, False)
    s = ElementOpenEx("link", HTMLAttribEx("type", "text/css", False) & HTMLAttribEx("rel", "stylesheet", False) & src, False)
    s = s & ElementCloseEx("link")
  Else
    s = ElementOpenEx("style", HTMLAttribEx("type", "text/css", False), False)
    sFile = WebGlobals.ServerRootDir & WebGlobals.CSSDir & StyleFileNoPath
    If Not FileExists(sFile) Then
      If RaiseErrorsOnMissingFile Then Call Err.Raise(ERR_INCLUDE_STYLE, "IncludeStyleEx", "The stylesheet '" & sFile & "' does not exist")
      GoTo ERR_END
    End If
    s = s & FileToString(WebGlobals.ServerRootDir & WebGlobals.CSSDir & StyleFileNoPath)
    s = s & ElementCloseEx("style")
  End If
  IncludeStyleEx = s
ERR_END:
  Exit Function
ERR_ERR:
  Call Err.Raise(ERR_INCLUDE_STYLE, ErrorSourceComponentEx(Err, "IncludeStyleEx", COMPONENT_NAME), Err.Description)
  Resume
End Function
'cadxx
Public Function HTMLRadioButtonEx2(ByVal Name As String, ByVal Value As String, Optional ByVal id As String, Optional ByVal Checked As Boolean = False, Optional ByVal Disabled As Boolean = False, Optional ByVal Caption As String = "", Optional CarridgeReturn As Boolean = True, Optional ByVal Attributes As String) As String
  Dim sattributes As String
  On Error GoTo HTMLRadioButtonEx_err

  'cad changed to be xml compatible
  sattributes = HTMLAttribEx("type", "radio", False) & HTMLAttribEx("name", Name, False) & HTMLAttribEx("value", Value, False) & sattributes
  If Len(id) > 0 Then sattributes = sattributes & HTMLAttribEx("id", id, False)
  If Disabled Then sattributes = sattributes & HTMLAttribEx("disabled", "YES", False)
  If Checked Then sattributes = sattributes & HTMLAttribEx("checked", "YES", False)
  HTMLRadioButtonEx2 = ElementOpenEx("input", sattributes, True)
  If Len(Caption) > 0 Then HTMLRadioButtonEx2 = HTMLRadioButtonEx2 & Caption
  If CarridgeReturn Then HTMLRadioButtonEx2 = HTMLRadioButtonEx2 & "<br>"

HTMLRadioButtonEx_end:
  Exit Function

HTMLRadioButtonEx_err:
 Err.Raise Err.Number, ErrorSourceComponentEx(Err, "HTMLRadioButtonEx", COMPONENT_NAME), Err.Description
 Resume
End Function

