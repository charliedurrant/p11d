Attribute VB_Name = "EFilerLows"
Option Explicit


Public Sub Notify(ByVal iNotify As IBaseNotify, ByRef Message As String, Optional ByVal Current As Long = 0, Optional ByVal Max As Long = 0)
  If Not iNotify Is Nothing Then Call iNotify.Notify(Current, Max, Message)
End Sub
Public Function OnlineCheckXML(func As String, app As String, appVersion As String, companyName As String)
Dim s As String


s = "<?xml version='1.0'?>" & vbCrLf
  s = s & "<automatic_updates>" & vbCrLf
    s = s & "<function>" & vbCrLf
      s = s & "<name>" & vbCrLf
      s = s & func & vbCrLf
      s = s & "</name>" & vbCrLf
      s = s & "<parameters>" & vbCrLf
        s = s & "<application>" & vbCrLf
        s = s & app & vbCrLf
        s = s & "</application>" & vbCrLf
        s = s & "<version>" & vbCrLf
        s = s & appVersion & vbCrLf
        s = s & "</version>" & vbCrLf
        s = s & "<company_name><![CDATA[" & companyName & "]]></company_name>" & vbCrLf
      s = s & "</parameters>" & vbCrLf
    s = s & "</function>"
  s = s & "</automatic_updates>" & vbCrLf
  
  OnlineCheckXML = s
  
End Function

'component wraps the XML http object and returns a DOMDocument30
'specifically uses XMLHTTP60 for submitting as this works with PROXY servers, or seems to, testing was with Proxy+
Public Function SubmitEx(ByRef sResponseText As String, ByVal sURL As String, sData As String, bAsync As Boolean, bCheckProxy As Boolean, Optional iNotify As IBaseNotify, Optional NotifyMessage As String, Optional NotifyIntervalMilliseconds As Long = 500) As DOMDocument60
  'specifically use 40-60
  Dim xmlhttp As XMLHTTP60
  
  Dim sMessage As String
  On Error GoTo err_Err
  
  
  Set xmlhttp = New XMLHTTP60
  
  ' this should be in CheckNetwork()
  If bCheckProxy Then
    'if we do a "POST" first and there is a Proxy Server that requires authentication then it always fails
    'need to do a simple GET first
    sMessage = "Contacting " & S_GG_LIVE_SUBMITADDRESS & " to check for Proxy Server login"
    Call Notify(iNotify, sMessage)
      
    Call xmlhttp.Open("GET", S_GG_LIVE_SUBMITADDRESS, False)
    Call xmlhttp.send("a")
    If xmlhttp.status = 407 Then Call Err.Raise(ERR_FAILED_SUBMISSIONS, "SubmitEx", "Submission cancelled due to invalid proxy information")
    If xmlhttp.status <> 200 Then Call Err.Raise(ERR_FAILED_SUBMISSIONS, "SubmitEx", "Submission cancelled as gateway site is unavailable.  HTTP Status return code = [" & xmlhttp.status & "]")
  End If
  
'  Set xmlhttp = New XMLHTTP40
  Call Notify(iNotify, NotifyMessage)
  'cad....debugging to allow for schema validating, this is rubbish and this should be built in via a call back
  Call Notify(iNotify, sData, -1, -1)
  
  Call xmlhttp.Open("POST", sURL, bAsync)
  
  
  Call xmlhttp.send(sData)
  
  If (bAsync) Then
    While (xmlhttp.ReadyState <> 4)
      Call Sleep(NotifyIntervalMilliseconds)
      Call Notify(iNotify, NotifyMessage)
    Wend
  End If
  
  If (xmlhttp.status <> 200) Then Call Err.Raise(ERR_FAILED_SUBMISSIONS, "SubmitEx", "Submission to gateway site failed.  HTTP Status return code = [" & xmlhttp.status & "]")
      
  sResponseText = xmlhttp.responseText
  Set SubmitEx = lows.DOMDocumentNewEFiler()
  Call SubmitEx.loadXML(sResponseText)
      
err_End:
  Set xmlhttp = Nothing
  Exit Function
err_Err:
  Call Err.Raise(Err.Number, ErrorSource(Err, "SubmitEx"), Err.Description)
  Resume
End Function





