VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VersionCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Private m_httpRequest As MSXML2.XMLHTTPRequest
Private m_xmlhttp As XMLHTTP60
   
Event Response(ByVal Response As String, lastAction As String)

Private Sub Class_Initialize()
  Set m_xmlhttp = New XMLHTTP60
End Sub

'OnReadyStateChange is set to be the default function in toos proc attribs
Private Sub Class_Terminate()
  Set m_xmlhttp = Nothing
End Sub
Public Sub Check()
  Dim xml As String

On Error GoTo err_Err

  Call m_xmlhttp.abort
  
  m_xmlhttp.OnReadyStateChange = Me
  xml = OnlineCheckXML("application_version_check", AppExeName, GetVersionString(False), "")
  Call m_xmlhttp.Open("POST", S_URL_AUTOMATIC_UPDATES, True)
  Call m_xmlhttp.send(xml)
  
err_End:
  Exit Sub
err_Err:
  p11d32.VersionCheckFinished = True
  Resume err_End
End Sub

Function OnReadyStateChange()
Attribute OnReadyStateChange.VB_UserMemId = 0
  Dim strHttpResponse As String
  Dim doc As DOMDocument60
  Dim node As IXMLDOMNode
  
  On Error GoTo err_Err
  
  If m_xmlhttp.ReadyState = 4 Then
    strHttpResponse = m_xmlhttp.responseText
    Set doc = New DOMDocument60
    doc.loadXML (strHttpResponse)
    If (doc.parseError.ErrorCode = 0) Then
      Set node = doc.selectSingleNode("automatic_updates/return")
      p11d32.VersionCheckResult = node.Text
      p11d32.VersionCheckFinished = True
     
    End If
  End If

err_End:
  Exit Function
err_Err:
  p11d32.VersionCheckFinished = True
  Resume err_End

End Function
