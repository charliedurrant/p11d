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

Private Sub Class_Initialize()
  Set m_Errors = New Collection
End Sub

Public Property Get Count() As Long
  Count = m_Errors.Count
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

Public Property Get HTML(Optional ByVal LoggedErrors As Boolean = False) As String
  Dim qsError As QString, errid As String
  Dim we As WebError, bShowError As Boolean
  Dim i As Long
  
  Set qsError = New QString
  If m_Errors.Count > 0 Then
    qsError.Append "<TABLE " & AttribEx("id", "errormessage") & AttribEx("class", "errormessage_div_server") & ">"
  End If
  For i = 1 To m_Errors.Count
    Set we = m_Errors.Item(i)
    If Not we.Logged Or LoggedErrors Then
      bShowError = True
      errid = "errorcallstack" & i
      qsError.Append "<TR><TD>"
      qsError.Append "<SPAN " & AttribEx("class", "errordescription") & ">"
      qsError.Append we.DescriptionHTML
      qsError.Append "</SPAN>"
      qsError.Append "</TD></TR>"
      qsError.Append "<TR><TD>"
      qsError.Append "<DIV " & AttribEx("class", "diverrorcallstackheader") & AttribEx("onclick", "flipDisplay(" & errid & ")") & ">"
      qsError.Append "Show error details"
      qsError.Append "</DIV>"
      qsError.Append "<DIV " & AttribEx("id", errid) & AttribEx("class", "diverrorcallstack") & ">"
      qsError.Append we.DetailsHTML
      qsError.Append "</DIV>"
      qsError.Append "</TD></TR>"
      qsError.Append "<TR><TD><BR></TD></TR>"
    End If
  Next i
  qsError.Append "</TABLE>"
  HTML = qsError
End Property

Public Function ToURL(Optional ByVal AllErrors As Boolean = False) As String
  Dim sErrorURL As String
  Dim we As WebError
  Dim i As Long, eCount As Long
  
  For i = 1 To m_Errors.Count
    Set we = m_Errors.Item(i)
    If Not we.Logged Or AllErrors Then
      If Len(sErrorURL) = 0 Then
        sErrorURL = we.URL(eCount)
      Else
        sErrorURL = sErrorURL & "&" & we.URL(eCount)
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

Public Property Get XML() As String
  Dim qsError As QString
  Dim we As WebError
  Dim i As Long
  Dim eCount As Long
  Dim errid As String
  
  Set qsError = New QString
  qsError.Append "<ERRORS>"
  For i = 1 To m_Errors.Count
    Set we = m_Errors.Item(i)
    qsError.Append "<ERROR>"
    qsError.Append "<DESCRIPTION><![CDATA[" & we.ErrorDescriptionXML & "]]></DESCRIPTION>"
    qsError.Append "<DETAILS>" & we.ErrorDetailsXML & "</DETAILS>"
    qsError.Append "</ERROR>"
  Next i
  qsError.Append "</ERRORS>"
  XML = qsError.bstr
  Set qsError = Nothing
End Property
