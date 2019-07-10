VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WebErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private m_Errors As Collection
Private m_WebGlobals As WebGlobals
' MPS
Private m_Basic As Boolean
Private m_BasicText As String

Private Sub Class_Initialize()
  Set m_Errors = New Collection
End Sub

Public Sub Clear()
  Set m_Errors = New Collection
End Sub
Public Property Get Count() As Long
  Count = m_Errors.Count
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
   Set NewEnum = m_Errors.[_NewEnum]
End Function

Public Property Let Basic(ByVal vValue As Boolean)
  m_Basic = vValue
End Property

Public Property Get Basic() As Boolean
  Basic = m_Basic
End Property

Public Property Let BasicText(ByVal vValue As String)
  m_BasicText = vValue
End Property

Public Property Get BasicText() As String
  BasicText = m_BasicText
End Property

Public Property Get Item(ByVal vIndex As Variant) As WebError
  On Error GoTo Item_err
  Set Item = m_Errors.Item(vIndex)
  Exit Property
Item_err:
  Err.Raise 380, "WebErrors.Item", "Invalid WebError index [" & vIndex & "]" & vbCrLf & Err.Description
End Property

Public Sub Add(ByVal ErrObj As ErrObject, Optional ByVal Description As String)
  Dim we As WebError
  
  Set we = New WebError
  we.Number = ErrObj.Number
  If Len(Description) = 0 Then
    we.Description = ErrObj.Description
  Else
    we.Description = Description & vbCrLf & ErrObj.Description
  End If
  we.Description = ADOSQLAnonymousDSN(we.Description)
  we.Source = ErrObj.Source
  m_Errors.Add we
End Sub

Public Property Get Text(Optional ByVal AllErrors As Boolean = False) As String
  Dim sErrorText As String
  Dim we As WebError
  Dim i As Long, eCount As Long
  
  For i = 1 To m_Errors.Count
    Set we = m_Errors.Item(i)
    If Not we.Logged Or AllErrors Then
      sErrorText = sErrorText & we.Description
      we.Logged = True
    End If
  Next i
  Text = sErrorText
End Property

' MPS
Public Property Get TextDetails(Optional ByVal AllErrors As Boolean = False) As String
  Dim sErrorTextDetails As String
  Dim we As WebError
  Dim i As Long, eCount As Long
  
  For i = 1 To m_Errors.Count
    Set we = m_Errors.Item(i)
    If Not we.Logged Or AllErrors Then
      sErrorTextDetails = sErrorTextDetails & "0x" & Right$("00000000" & Hex$(we.Number), 8) & ": " & we.Source & ";"
      we.Logged = True
    End If
  Next i
  TextDetails = sErrorTextDetails
End Property

'cadxx 'breaks compat
Public Property Get HTML(Optional ByVal LoggedErrors As Boolean = False, Optional ByVal WebGlobals As WebGlobals = Nothing) As String
  HTML = Me.HTML2(LoggedErrors, WebGlobals)
End Property

Public Property Get HTML2(Optional ByVal LoggedErrors As Boolean = False, Optional ByVal WebGlobals As WebGlobals = Nothing, Optional ByVal SummariseAll As Boolean = False) As String
  Dim qsError As QString, errid As String
  Dim we As WebError, bShowError As Boolean
  Dim i As Long

  Set qsError = New QString
  If m_Errors.Count > 0 Then
    If SummariseAll Then
      errid = "errorcallstack"
      qsError.Append "<DIV " & HTMLAttribEx("class", "div_error_callstack_header_server", False) & HTMLAttribEx("onclick", "flipDisplay(" & errid & ")", False) & ">"
      qsError.Append "Show error"
      qsError.Append "</DIV>"
      qsError.Append "<DIV " & HTMLAttribEx("id", errid, False) & HTMLAttribEx("class", "div_error_callstack_server", False) & ">"
    End If
    qsError.Append "<TABLE " & HTMLAttribEx("id", "errormessage", False) & HTMLAttribEx("class", "table_errormessage_server", False) & ">"
    qsError.Append "<TR><TD style=""background-color:black;color:white"">Error</TD></TR>"
    ' MPS
    If m_Basic Then
      qsError.Append "<TR><TD>"
      qsError.Append "<SPAN " & HTMLAttribEx("class", "span_error_description_server", False) & ">"
      qsError.Append ReplaceHTMLMetacharactersEx(m_BasicText)
      qsError.Append "</SPAN>"
      qsError.Append "</TD></TR>"
      qsError.Append "<TR><TD><BR></TD></TR>"
    Else
      For i = 1 To m_Errors.Count
        Set we = m_Errors.Item(i)
        If Not we.Logged Or LoggedErrors Then
          bShowError = True
          errid = "error_callstack_server" & i
          qsError.Append "<TR><TD>"
          qsError.Append "<SPAN " & HTMLAttribEx("class", "span_error_description_server", False) & ">"
          qsError.Append we.DescriptionHTML
          qsError.Append "</SPAN>"
          qsError.Append "</TD></TR>"
          qsError.Append "<TR><TD>"
          qsError.Append "<DIV " & HTMLAttribEx("class", "div_error_callstack_header_server", False) & HTMLAttribEx("onclick", "flipDisplay(" & errid & ")", False) & ">"
          qsError.Append "Show error details"
          qsError.Append "</DIV>"
          qsError.Append "<DIV " & HTMLAttribEx("id", errid, False) & HTMLAttribEx("class", "div_error_callstack_server", False) & ">"
          qsError.Append we.DetailsHTML
          qsError.Append "</DIV>"
          qsError.Append "</TD></TR>"
          qsError.Append "<TR><TD><BR></TD></TR>"
        End If
      Next i
    End If
    qsError.Append "</TABLE>"
    If SummariseAll Then qsError.Append "</DIV>"
    If Not WebGlobals Is Nothing Then
      If WebGlobals.IsPresent("Contact") Then
        qsError.Append WebGlobals.Item("Contact") & "<br>"
      End If
    End If
    HTML2 = qsError
  End If
End Property

Public Function ToURL(Optional ByVal AllErrors As Boolean = False) As String
  Dim sErrorURL As String
  Dim we As WebError
  Dim i As Long, eCount As Long

  ' MPS
  If m_Basic Then
    sErrorURL = "NUMBER" & "0" & "=" & "0" & _
                "&DESCRIPTION" & "0" & "=" & URLEncode(ReplaceHTMLMetacharactersEx(m_BasicText)) & _
                "&SOURCE" & "0" & "=" & URLEncode(ReplaceHTMLMetacharactersEx("..."))
    ToURL = URLEncode(sErrorURL)
    Exit Function
  End If
  
  For i = 1 To m_Errors.Count
    Set we = m_Errors.Item(i)
    If Not we.Logged Or AllErrors Then
      If Len(sErrorURL) = 0 Then
        sErrorURL = we.url(eCount)
      Else
        sErrorURL = sErrorURL & "&" & we.url(eCount)
      End If
      eCount = eCount + 1
    End If
  Next i
  ToURL = URLEncode(sErrorURL)
End Function

Public Sub FromURL(ByVal surl As String)
  Dim i As Long
  Dim p0 As Long, p1 As Long
  Dim sErrNum As String, sErrDesc As String, sErrSource As String
  Dim we As WebError
  
  On Error GoTo URLToErrors_err
  surl = URLDecode(surl)
  i = 0
  Do While True
    sErrNum = "&NUMBER" & i & "="
    sErrDesc = "&DESCRIPTION" & i & "="
    sErrSource = "&SOURCE" & i & "="
    
    p0 = InStr(1, surl, sErrNum, vbTextCompare)
    If p0 = 0 Then Exit Do
    p0 = p0 + Len(sErrNum)
    p1 = InStr(p0, surl, sErrDesc, vbTextCompare)
    If (p1 = 0) Then Err.Raise ERR_URLERROR, "URLToErrors", "Unable to parse a URL to an error"
    sErrNum = Mid$(surl, p0, p1 - p0)
    
    p0 = InStr(1, surl, sErrDesc, vbTextCompare)
    p0 = p0 + Len(sErrDesc)
    p1 = InStr(p0, surl, sErrSource, vbTextCompare)
    If (p1 = 0) Then Err.Raise ERR_URLERROR, "URLToErrors", "Unable to parse a URL to an error"
    sErrDesc = Mid$(surl, p0, p1 - p0)
    
    p0 = InStr(1, surl, sErrSource, vbTextCompare)
    p0 = p0 + Len(sErrSource)
    p1 = InStr(p0, surl, "&", vbTextCompare)
    If p1 = 0 Then p1 = Len(surl) + 1
    sErrSource = Mid$(surl, p0, p1 - p0)
    
    Set we = New WebError
    we.Number = CLng(sErrNum)
    we.Description = sErrDesc
    we.Source = sErrSource
    Call m_Errors.Add(we)
    i = i + 1
  Loop
  Exit Sub
  
URLToErrors_err:
  App.LogEvent "URLToErrors FAILED, Error: " & Err.Description & " Source: " & Err.Source, vbLogEventTypeError
End Sub

' see schema xmlerror.xml
Public Property Get XML() As String
  Dim qsError As QString
  Dim we As WebError
  Dim i As Long
  
  Set qsError = New QString
  qsError.Append "<errors>"
  For i = 1 To m_Errors.Count
    Set we = m_Errors.Item(i)
    qsError.Append "<error>"
    qsError.Append "<description>" & XMLTextEx(we.ErrorDescriptionXML) & "</description>"
    qsError.Append "<details>" & we.ErrorDetailsXML & "</details>"
    qsError.Append "</error>"
  Next i
  qsError.Append "</errors>"
  XML = qsError
  Set qsError = Nothing
End Property

' see schema databaseerror.xml
Public Sub AddXMLDBErrors(ByVal xmldom As DOMDocument)
  Dim we As WebError
  Dim sErrNum As String, sErrDesc As String, sErrSource As String
  Dim n As IXMLDOMNode
        
  On Error GoTo AddXMLDBErrors_err
  Set n = xmldom.selectSingleNode("errormessage/errno")
  sErrNum = n.Text
  Set n = xmldom.selectSingleNode("errormessage/message")
  sErrDesc = n.Text
  Set n = xmldom.selectSingleNode("errormessage/source")
  sErrSource = n.Text
  Set we = New WebError
  we.Number = CLng(sErrNum)
  we.Description = sErrDesc
  we.Source = sErrSource
  Call m_Errors.Add(we)
  Exit Sub
  
AddXMLDBErrors_err:
  ' cannot raise errors !!
  App.LogEvent "AddXMLDBErrors FAILED, Error: " & ADOSQLAnonymousDSN(Err.Description) & " Source: " & Err.Source, vbLogEventTypeError
End Sub


