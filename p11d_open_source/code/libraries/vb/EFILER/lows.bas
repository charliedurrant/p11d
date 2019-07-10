Attribute VB_Name = "lows"
Option Explicit
Public Sub Notify(ByVal iNotify As IBaseNotify, ByRef Message As String, Optional ByVal Current As Long = 0, Optional ByVal Max As Long = 0)
  If Not iNotify Is Nothing Then Call iNotify.Notify(Current, Max, Message)
End Sub

'component wraps the XML http object and returns a DOMDocument30
'specifically uses XMLHTTP40 for submitting as this works with PROXY servers, or seems to, testing was with Proxy+
Public Function SubmitEx(ByRef sResponseText As String, ByVal sURL As String, sData As String, bAsync As Boolean, bCheckProxy As Boolean, Optional iNotify As IBaseNotify, Optional NotifyMessage As String, Optional NotifyIntervalMilliseconds As Long = 500) As DOMDocument30
  'specifically use 40
  Dim xmlhttp As XMLHTTP40
  Dim sMessage As String
  On Error GoTo err_err
  
  Set xmlhttp = New XMLHTTP40
  
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
    While (xmlhttp.readyState <> 4)
      Call Sleep(NotifyIntervalMilliseconds)
      Call Notify(iNotify, NotifyMessage)
    Wend
  End If
  
  If (xmlhttp.status <> 200) Then Call Err.Raise(ERR_FAILED_SUBMISSIONS, "SubmitEx", "Submission to gateway site failed.  HTTP Status return code = [" & xmlhttp.status & "]")
      
  sResponseText = xmlhttp.responseText
  Set SubmitEx = New DOMDocument30
  Call SubmitEx.loadXML(sResponseText)
      
err_end:
  Set xmlhttp = Nothing
  Exit Function
err_err:
  Call Err.Raise(Err.Number, ErrorSource(Err, "SubmitEx"), Err.Description)
  Resume
End Function




