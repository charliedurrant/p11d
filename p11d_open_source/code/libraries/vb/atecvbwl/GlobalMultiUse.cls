VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GlobalMultiUse"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const MAX_ERRORMESSAGE_TEXT As Long = 4096
Private Const MIN_ERRORMESSAGE_TEXT As Long = 512
Private Const XML_VERSION As String = "<?xml version=""1.0""?>"
Private Const XML_EM_BEGIN As String = "<errormessage>"
Private Const XML_EM_END As String = "</errormessage>"

Private Const XML_ENO_BEGIN As String = "<errno>"
Private Const XML_ENO_END As String = "</errno>"
Private Const XML_SRC_BEGIN As String = "<source>"
Private Const XML_SRC_END As String = "</source>"
Private Const XML_MSG_BEGIN As String = "<message>"
Private Const XML_MSG_END As String = "</message>"

Public Property Let Debugging(ByVal NewValue As Boolean)
  g_Debugging = NewValue
  g_CRLF = vbCrLf
End Property

Public Property Get Debugging() As Boolean
  Debugging = g_Debugging
End Property

Public Property Get CRLF() As String
  CRLF = g_CRLF
End Property

Public Function Attrib(ByVal Name As String, ByVal Value As Variant) As String
  Attrib = HTMLAttribEx(Name, Value, False)
End Function

Public Function HTMLAttrib(ByVal Name As String, ByVal Value As Variant, Optional ByVal EscapeQuotes As Boolean = False) As String
  HTMLAttrib = HTMLAttribEx(Name, Value, EscapeQuotes)
End Function

'cadxx new param Optional ByVal AllowNoClose As Boolean
Public Function ElementOpen(Name As String, Optional Attributes As String = "", Optional ByVal ShortClose As Boolean = False, Optional ByVal AllowNoClose As Boolean) As String
  ElementOpen = ElementOpenEx(Name, Attributes, ShortClose, AllowNoClose)
End Function

Public Property Get XMLHeader() As String
  XMLHeader = XML_VERSION
End Property

Public Function Style(ByVal Name As String, ByVal Value As String) As String
  Style = Name & ":" & Value & ";"
End Function

Public Function ElementClose(ByVal Name As String) As String
  ElementClose = ElementCloseEx(Name)
End Function

Public Function CheckBox(Optional ByVal Checked As Boolean = False, Optional ByVal id As String, Optional ByVal Attributes As String) As String
  Dim s As String
  s = ElementOpenEx("input", HTMLAttribEx("type", "CHECKBOX", False) & HTMLAttribEx("id", id, False) & HTMLAttribEx("name", id, False) & Attributes & IIf(Checked, "CHECKED", ""), True)
  CheckBox = s
End Function

Public Function TextBox(Optional ByVal Value As String, Optional ByVal id As String, Optional MaxLength As Long = -1, Optional ByVal Attributes As String) As String
  Dim s As String
  s = ElementOpenEx("INPUT", HTMLAttribEx("type", "TEXT", False) & HTMLAttribEx("value", Value, False) & HTMLAttribEx("id", id, False) & HTMLAttribEx("name", id, False) & Attributes & IIf(Checked, "CHECKED", ""), True)
  TextBox = s
End Function

Public Function XMLText(ByVal s As String) As String
  XMLText = XMLTextRef(s)
End Function

' Does not convert booleans,  pass in as string
Public Function XMLAttrib(ByVal attrName As String, ByVal attrValue As String, Optional ByVal add_trailing_space As Boolean = True, Optional ByVal tolower_attrname As Boolean = True) As String
  XMLAttrib = XMLAttribEx(attrName, attrValue, add_trailing_space, tolower_attrname)
End Function

' Converts booleans to 0/-1 & deals with Nulls - used in AO
Public Function XMLAttrib2(ByVal attrName As String, ByVal Value As Variant) As String
Attribute XMLAttrib2.VB_Description = "Converts booleans to 0/-1 & treats Null values as zero length strings"
  XMLAttrib2 = XMLAttribEx2(attrName, Value)
End Function

Public Function IncludeScript(ByVal WebGlobals As WebGlobals, ByVal ScriptFileNoPath As String, Optional IncludeFile As Boolean = True) As String
  Dim sFile As String, s As String
  Dim src As String
  
  On Error GoTo IncludeScript_ERR
  If IncludeFile Then src = HTMLAttribEx("src", WebGlobals.ScriptDir & ScriptFileNoPath, False)
  s = ElementOpenEx("script", HTMLAttribEx("type", "text/javascript", False) & HTMLAttribEx("language", "JavaScript", False) & src, False)
  If Not IncludeFile Then s = s & FileToString(WebGlobals.ServerRootDir & WebGlobals.ScriptDir & ScriptFileNoPath)
  s = s & ElementClose("script")
  IncludeScript = s
  
IncludeScript_END:
  Exit Function
IncludeScript_ERR:
  Call Err.Raise(ERR_INCLUDE_SCRIPT, ErrorSourceComponentEx(Err, "IncludeScript", COMPONENT_NAME), Err.Description)
  Resume
End Function

'cadxx new changed to call sub function
Public Function IncludeStyle(ByVal WebGlobals As WebGlobals, ByVal StyleFileNoPath As String, Optional IncludeFile As Boolean = True) As String
  On Error GoTo ERR_ERR

  IncludeStyle = IncludeStyleEx(WebGlobals, StyleFileNoPath, IncludeFile, True)

ERR_END:
  Exit Function
ERR_ERR:
  Call Err.Raise(ERR_INCLUDE_STYLE, ErrorSourceComponentEx(Err, "IncludeStyle", COMPONENT_NAME), Err.Description)
  Resume
End Function

'Public Function IncludeStyle(ByVal WebGlobals As WebGlobals, ByVal StyleFileNoPath As String, Optional IncludeFile As Boolean = True) As String
'  Dim sFile As String, s As String
'  Dim src As String
'
'  On Error GoTo IncludeStyle_ERR
'  If IncludeFile Then
'    src = HTMLAttribEx("href", WebGlobals.CSSDir & StyleFileNoPath)
'    s = ElementOpenEx("link", HTMLAttribEx("type", "text/css") & HTMLAttribEx("rel", "stylesheet") & src, False)
'    s = s & ElementClose("link")
'  Else
'    s = ElementOpenEx("style", HTMLAttribEx("type", "text/css"), False)
'    s = s & FileToString(WebGlobals.ServerRootDir & WebGlobals.CSSDir & StyleFileNoPath)
''    s = s & ElementClose("style")
'  End If
'  IncludeStyle = s
'IncludeStyle_END:
'  Exit Function
'IncludeStyle_ERR:
'  Call Err.Raise(ERR_INCLUDE_STYLE, ErrorSourceComponentEx(Err, "IncludeStyle", COMPONENT_NAME), Err.Description)
'  Resume
'End Function

Public Function ErrorSourceComponent(ByVal Err As ErrObject, ByVal FunctionName As String, ByVal AppExeName As String) As String
  ErrorSourceComponent = ErrorSourceComponentEx(Err, FunctionName, AppExeName)
End Function

Public Function RSToTable(ByVal rs As Recordset, Optional ByVal Attributes As String) As String
  Dim HTMLString As QString
  Dim i As Long

  On Error GoTo RSToTable_ERR
  Set HTMLString = New QString
  Call HTMLString.Append(ElementOpen("TABLE", HTMLAttribEx("CLASS", "ATEC_TABLE", False) & Attributes))
  If Not rs.EOF And Not rs.BOF Then
    Call HTMLString.Append(ElementOpen("TR", HTMLAttribEx("CLASS", "ATEC_TRH", False)))
    For i = 0 To rs.Fields.Count - 1
      Call HTMLString.Append(ElementOpen("TH", HTMLAttribEx("CLASS", "ATEC_TH", False)))
      Call HTMLString.Append(rs.Fields(i).Name)
      Call HTMLString.Append(ElementClose("TH"))
    Next i
    Call HTMLString.Append(ElementClose("TR"))
    Do While Not rs.EOF
      Call HTMLString.Append(ElementOpen("TR", HTMLAttribEx("CLASS", "ATEC_TRD", False)))
      For i = 0 To rs.Fields.Count - 1
        Call HTMLString.Append(ElementOpen("TD", HTMLAttribEx("CLASS", "ATEC_TD", False)))
        Call HTMLString.Append(rs.Fields(i).Value & "&nbsp;")
        Call HTMLString.Append(ElementClose("TD"))
      Next i
      Call HTMLString.Append(ElementClose("TR"))
      rs.MoveNext
    Loop
  End If
  Call HTMLString.Append(ElementClose("TABLE"))
  RSToTable = HTMLString
  Exit Function
  
  
RSToTable_ERR:
  Call Err.Raise(ERR_INCLUDE_STYLE, ErrorSourceComponentEx(Err, "RSToTable", COMPONENT_NAME), Err.Description)
  Resume
End Function

Public Function HTMLCheckBox(Optional ByVal id As String, Optional ByVal Checked As Boolean = False, Optional ByVal Disabled As Boolean = False, Optional ByVal Attributes As String) As String
  HTMLCheckBox = HTMLCheckBoxEx(id, Checked, Disabled, Attributes)
End Function

Public Function HTMLRadioButton(Optional ByVal idGroup As String, Optional ByVal id, Optional ByVal Checked As Boolean = False, Optional ByVal Disabled As Boolean = False, Optional ByVal Attributes As String) As String
  HTMLRadioButton = HTMLRadioButtonEx(idGroup, id, Checked, Disabled, Attributes)
End Function

Public Function HTMLTextBox(Optional ByVal id As String, Optional ByVal Value As String = "", Optional ByVal Disabled As Boolean = False, Optional ByVal Attributes As String) As String
  HTMLTextBox = HTMLTextBoxEx(id, Value, Disabled, Attributes)
End Function

Public Function HTMLTextArea(Optional ByVal id As String, Optional ByVal Value As String = "", Optional ByVal Rows As Long = 1, Optional ByVal Disabled As Boolean = False, Optional ByVal Attributes As String) As String
  HTMLTextArea = HTMLTextAreaEx(id, Value, Rows, Disabled, Attributes)
End Function

Public Function HTMLListBox(Optional ByVal id As String, Optional ByVal ValueList As Variant, Optional ByVal ValueListDisplay As Variant, Optional ByVal Value As String = "", Optional ByVal Disabled As Boolean = False, Optional ByVal Attributes As String) As String
  HTMLListBox = HTMLListBoxEx(id, ValueList, ValueListDisplay, Value, Disabled, Attributes)
End Function

Public Function HTMLHiddenInput(Optional ByVal id As String, Optional ByVal Value As String = "", Optional ByVal Attributes As String) As String
  HTMLHiddenInput = HTMLHiddenInputEx(id, Value, Attributes)
End Function

Public Function HTMLButton(Optional ByVal id As String, Optional ByVal Value As String = "", Optional ByVal Disabled As Boolean = False, Optional ByVal OnClickEvent As String, Optional ByVal Attributes As String) As String
  HTMLButton = HTMLButtonEx(id, Value, Disabled, OnClickEvent, Attributes)
End Function

' MPS, AF changed
Public Function ReplaceHTMLMetacharacters(ByVal sText As String) As String
  ReplaceHTMLMetacharacters = ReplaceHTMLMetacharactersRef(sText, False)
End Function

' AF changed,  also replace space with &nbsp; (non breaking space)
Public Function ReplaceHTMLMetacharactersNS(ByVal sText As String) As String
  ReplaceHTMLMetacharactersNS = ReplaceHTMLMetacharactersRef(sText, True)
End Function

Public Sub DownLoad(ByVal Response As Response, Data As Variant, ByVal ContentType As String, ByVal FileName As String)
  On Error GoTo DownLoad_ERR
  Call DownLoadEx(Response, Data, ContentType, FileName)
  
DownLoad_END:
  Exit Sub
DownLoad_ERR:
  Call Err.Raise(Err.Number, ErrorSourceComponentEx(Err, "DownLoad", COMPONENT_NAME), Err.Description)
End Sub

Public Sub BasicAuthenticationUserDetails(ByRef UserName As String, ByRef Password As String, ByVal Request As Request)
  
  On Error GoTo BasicAuthenticationUserDetails_ERR
  
  Call BasicAuthenticationUserDetailsEX(UserName, Password, Request)
  
BasicAuthenticationUserDetails_END:
  Exit Sub
BasicAuthenticationUserDetails_ERR:
  Call Err.Raise(Err.Number, ErrorSourceComponentEx(Err, "BasicAuthenticationUserDetails", COMPONENT_NAME), Err.Description)
End Sub

Public Sub BasicAuthenticationInit(ByVal Response As Response, ByVal Realm As String, FailureHTML As String)
  
  On Error GoTo BasicAuthenticationInit_ERR
  
  Call BasicAuthenticationInitEX(Response, Realm, FailureHTML)
  
BasicAuthenticationInit_END:
  Exit Sub
BasicAuthenticationInit_ERR:
  Call Err.Raise(Err.Number, ErrorSourceComponentEx(Err, "BasicAuthenticationInit", COMPONENT_NAME), Err.Description)
End Sub

'cadxx new
Public Function IncludeStandardStyles(ByVal WebGlobals As WebGlobals, Optional IncludeFile As Boolean = True) As String
  Dim s As String

  On Error GoTo IncludeStandardStyles_ERR
  
  s = IncludeStyleEx(WebGlobals, "abatec.css", IncludeFile, True)
  s = s & IncludeStyleEx(WebGlobals, "override.css", IncludeFile, True)
  s = s & IncludeStyleEx(WebGlobals, "product.css", IncludeFile, True)
  s = s & IncludeStyleEx(WebGlobals, "implementation.css", IncludeFile, True)

  IncludeStandardStyles = s

  Exit Function
IncludeStandardStyles_ERR:
  Call Err.Raise(Err.Number, ErrorSourceComponentEx(Err, "IncludeStandardStyles", COMPONENT_NAME), Err.Description)
End Function


'cadxx
Public Function HTMLRadioButton2(ByVal Name As String, ByVal Value As String, Optional ByVal id As String, Optional ByVal Checked As Boolean = False, Optional ByVal Disabled As Boolean = False, Optional ByVal Caption As String = "", Optional CarridgeReturn As Boolean = True, Optional ByVal Attributes As String) As String
  HTMLRadioButton2 = HTMLRadioButtonEx2(Name, Value, id, Checked, Disabled, Caption, CarridgeReturn, Attributes)
End Function

'cadxx new
Public Function ErrorPage(WebGlobals As WebGlobals, ErrDescription As String, ErrNumber As Long, ErrSource As String) As String
  Dim i As Long
  Dim v As Variant
  Dim qs As QString
  Dim weberrs As WebErrors

  Set qs = New QString
    Set weberrs = New WebErrors

    Err.Number = ErrNumber
    Err.Description = ErrDescription
    Err.Source = ErrSource
    Call weberrs.Add(Err)

    qs.Append ElementOpen("html")
    qs.Append ElementOpen("head")
    qs.Append ElementOpen("title") & "Error" & ElementClose("title")
    If Not WebGlobals Is Nothing Then
      qs.Append IncludeStandardStyles(WebGlobals, False)
      qs.Append IncludeScript(WebGlobals, "abatec_system.js")
      qs.Append IncludeScript(WebGlobals, "abatec_graphics.js")
    End If
    qs.Append ElementOpen("title") & "Error" & ElementClose("title")
    qs.Append ElementClose("head")
    qs.Append ElementOpen("body")
    
    qs.Append weberrs.HTML(, WebGlobals)

  qs.Append ElementClose("body")
  qs.Append ElementClose("html")

  ErrorPage = qs
  'ErrorString = err.Description & ":" & err.Number
End Function

' see schema databaseerror.xml
Public Function CreateXMLDBError(ByVal ErrNumber As Long, ByVal ErrSource As String, ByVal ErrMessage As String) As String
  Dim s0 As String, s1 As String, iLen As Long
  
  'MAX_ERRORMESSAGE_TEXT
  s0 = XML_VERSION & vbCrLf & XML_EM_BEGIN & vbCrLf & _
                        XML_ENO_BEGIN & ErrNumber & XML_ENO_END & vbCrLf & _
                        XML_SRC_BEGIN & ErrSource & XML_SRC_END & vbCrLf & _
                        XML_MSG_BEGIN
  s1 = XML_MSG_END & vbCrLf & XML_EM_END
  iLen = MAX_ERRORMESSAGE_TEXT - Len(s0) - Len(s1) - 1
  If iLen < MIN_ERRORMESSAGE_TEXT Then Err.Raise ERR_URLERROR, "CreateXMLDBError", "Cannot construct CreateXMLDBError as there is not enough space for the errormessage"
  If Len(ErrMessage) > iLen Then
    ErrMessage = XMLTextEx(Left$(ErrMessage, iLen - 64) & "...")
  Else
    ErrMessage = XMLTextRef(ErrMessage)
  End If
  CreateXMLDBError = s0 & ErrMessage & s1
End Function







