VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EFilerGatewayComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mGatewayTest As Integer
Private mSubmitAddress As String
Private mPollAddress As String
Private mDeleteAddress As String
Private mSubmitType As SUBMISSION_TYPE

Friend Property Let TestSubmission(sNewValue As Boolean)
  If sNewValue = True Then
    mGatewayTest = 1
    mSubmitAddress = S_GG_TEST_SUBMITADDRESS
    mPollAddress = S_GG_TEST_POLLADDRESS
    mDeleteAddress = S_GG_TEST_DELETEADDRESS
  Else
    mGatewayTest = 0
    mSubmitAddress = S_GG_LIVE_SUBMITADDRESS
    mPollAddress = S_GG_LIVE_POLLADDRESS
    mDeleteAddress = S_GG_LIVE_DELETEADDRESS
  End If
  
End Property

Friend Property Get TestSubmission() As Boolean
  TestSubmission = IIf(mGatewayTest = 1, True, False)
End Property

Friend Property Let SubmitType(NewValue As SUBMISSION_TYPE)
  mSubmitType = NewValue
End Property

Friend Property Get SubmitType() As SUBMISSION_TYPE
  SubmitType = mSubmitType
End Property

Friend Function PostSubmission(ByVal SubmissionBody As String, ByVal MockSubmission As Boolean, ByVal Async As Boolean, ByVal iNotify As IBaseNotify, ByVal iNotifyIntervalMilliseconds As Long, ByVal bCheckProxy As Boolean) As DOMDocument60
  Dim SubmitDoc As DOMDocument60
  Dim MockDoc As DOMDocument60
  Dim sResponseText As String
  Dim sTempFile As String

  On Error GoTo PostSubmission_Err
  
  Set SubmitDoc = New DOMDocument60
  If Not SubmitDoc.loadXML(SubmissionBody) Then Err.Raise ERR_INVALID_XML, "PostSubmission", "Invalid XML for submission"
  If Not MockSubmission Then
    ' Post as a SUBMISSION_REQUEST and wait for response
    'was
    'Set PostSubmission = SubmitEx(sResponseText, mSubmitAddress, SubmitDoc.xml, Async, bCheckProxy, iNotify, "Submitting data to the gateway", iNotifyIntervalMilliseconds)
    'now
    Set PostSubmission = SubmitEx(sResponseText, mSubmitAddress, SubmissionBody, Async, bCheckProxy, iNotify, "Submitting data to the gateway", iNotifyIntervalMilliseconds)
    If Len(PostSubmission.xml) = 0 Then
      app.LogEvent "Invalid response from gateway.  Response text: " & vbCrLf & vbCrLf & sResponseText, vbLogEventTypeError
      Err.Raise ERR_GATEWAY, "PostSubmission", "A valid response was not returned from the Gateway, please check that the Gateway is operational.  See event log for details."
    End If
  Else
    ' send back a submit_acknowledgement
    Set MockDoc = New DOMDocument60
    MockDoc.loadXML (S_SAMPLEXML_SUBMISSION_ACKNOWLEDGEMENT)
    Set PostSubmission = MockDoc
  
  End If
 
 
PostSubmission_End:
  Set SubmitDoc = Nothing
  Set MockDoc = Nothing
  Exit Function
   
PostSubmission_Err:
  Err.Raise Err.Number, ErrorSource(Err, "PostSubmission"), Err.Description
  Resume
End Function

Friend Function PollSubmission(ByVal CorrID As String, ByVal MockSubmission As Boolean, ByVal ResponseEndPoint As String, ByVal bCheckProxy As Boolean) As DOMDocument60
  Dim httpreq As XMLHTTP60
  Dim PollDoc As DOMDocument60
  Dim currnode As IXMLDOMNode
  Dim MsgType As String, ErrMsgType As String, MsgText As String
  Dim Wait As Long, s As String
  Dim MockDoc As DOMDocument60 ' Efiler As Efiler,
  Dim sClass As String, sURI As String
  Dim sResponseText As String
On Error GoTo PollSubmission_Err

  If Len(ResponseEndPoint) = 0 Then ResponseEndPoint = mPollAddress
  
  Set PollDoc = lows.DOMDocumentNewEFiler()
  
  
  With PollDoc
    .loadXML (S_GG_ENVELOPE_KEYS_START & S_GG_ENVELOPE_KEYS_END & S_GG_ENVELOPE_END)
    
    Set currnode = .selectSingleNode(PayeOnlineSearchString("//GovTalkMessage/EnvelopeVersion"))
    currnode.Text = S_GG_ENVELOPE_VERSION
    Set currnode = .selectSingleNode(PayeOnlineSearchString("//GovTalkMessage/Header/MessageDetails/Transformation"))
    currnode.Text = S_GG_TRANSFORMATION
    Set currnode = .selectSingleNode(PayeOnlineSearchString("//GovTalkMessage/Header/MessageDetails/GatewayTest"))
    currnode.Text = mGatewayTest
    'FIX #50 - was optional
    Set currnode = .selectSingleNode(PayeOnlineSearchString("//GovTalkDetails/TargetDetails/Organisation"))
    currnode.Text = S_GG_ORGANISATION
                 
    ' global Government gateway settings, set by code
    'FIX #56 start
    If mSubmitType = CorporationTax Then
      sClass = S_GG_CT_CLASS
      sURI = S_GG_CT_URI
    ElseIf mSubmitType = P11D Then
      sClass = S_GG_P11D_CLASS
      sURI = S_GG_P11D_URI
    End If
    Set currnode = .selectSingleNode(PayeOnlineSearchString("//GovTalkMessage/Header/MessageDetails/Class"))
    currnode.Text = sClass
    Set currnode = .selectSingleNode(PayeOnlineSearchString("//GovTalkMessage/GovTalkDetails/ChannelRouting/Channel/URI")) 'moved from above
    currnode.Text = sURI 'mandatory if  channelrouting used
    'FIX #56 end
    Set currnode = .selectSingleNode(PayeOnlineSearchString("//GovTalkMessage/Header/MessageDetails/Qualifier"))
    currnode.Text = "poll"
    Set currnode = .selectSingleNode(PayeOnlineSearchString("//GovTalkMessage/Header/MessageDetails/Function"))
    currnode.Text = "submit"
    Set currnode = .selectSingleNode(PayeOnlineSearchString("//GovTalkMessage/Header/MessageDetails/CorrelationID"))
    currnode.Text = CorrID
    
    'pq this element seems to require a value if it is present
    'doesn't actually have a purpose
    Set currnode = .selectSingleNode(PayeOnlineSearchString("//GovTalkMessage/Header/SenderDetails/IDAuthentication/Authentication/Method"))
    currnode.Text = "clear"

  End With
   
  If Not MockSubmission Then
    'Post as a SUBMISSION_POLL and wait for response
    Set PollSubmission = SubmitEx(sResponseText, ResponseEndPoint, PollDoc.xml, False, bCheckProxy)
    
    If Len(PollSubmission.xml) = 0 Then
      app.LogEvent "Invalid response from gateway.  Response text: " & vbCrLf & vbCrLf & sResponseText, vbLogEventTypeError
      Err.Raise ERR_GATEWAY, "PostSubmission", "A valid response was not returned from the Gateway, please check that the Gateway is operational.  See event log for details."
    End If
  Else
    ' send back responsedoc xml
    Set MockDoc = lows.DOMDocumentNewEFiler()
    Call MockDoc.loadXML(S_SAMPLEXML_SUBMISSION_RESPONSE)
    Set PollSubmission = MockDoc
  End If

    
PollSubmission_End:
  Set httpreq = Nothing
  Set PollDoc = Nothing
  Set MockDoc = Nothing
  Set currnode = Nothing
  Exit Function
  
PollSubmission_Err:
  Err.Raise Err.Number, ErrorSource(Err, "PollSubmission"), Err.Description
  Resume PollSubmission_End
End Function
     
Friend Function DeleteSubmission(ByVal CorrID As String, ByVal MockSubmission As Boolean, ByVal ResponseEndPoint As String, ByVal bCheckProxy As Boolean) As DOMDocument60
  Dim httpreq As XMLHTTP60
  Dim DelDoc As DOMDocument60
  Dim currnode As IXMLDOMNode
  Dim MockDoc As DOMDocument60 ' Efiler As Efiler
  Dim sClass As String, sURI As String
  Dim sResponseText As String
  On Error GoTo DeleteSubmission_Err
  
  If Len(ResponseEndPoint) = 0 Then ResponseEndPoint = mDeleteAddress
  
  Set DelDoc = lows.DOMDocumentNewEFiler()
  
    
  With DelDoc
    .loadXML (S_GG_ENVELOPE_KEYS_START & S_GG_ENVELOPE_KEYS_END & S_GG_ENVELOPE_END) ' body block must be present and unpopulated
    
    ' gateway settings
    Set currnode = .selectSingleNode(lows.PayeOnlineSearchString("//GovTalkMessage/EnvelopeVersion"))
    currnode.Text = S_GG_ENVELOPE_VERSION
    'FIX #50 - was optional
    Set currnode = .selectSingleNode(PayeOnlineSearchString("//GovTalkDetails/TargetDetails/Organisation"))
    currnode.Text = S_GG_ORGANISATION
    
    ' header block settings
    'FIX #56
    If mSubmitType = CorporationTax Then
      sClass = S_GG_CT_CLASS
      sURI = S_GG_CT_URI
    ElseIf mSubmitType = P11D Then
      sClass = S_GG_P11D_CLASS
      sURI = S_GG_P11D_URI
    End If
    Set currnode = .selectSingleNode(PayeOnlineSearchString("//GovTalkMessage/Header/MessageDetails/Class"))
    currnode.Text = sClass
    Set currnode = .selectSingleNode(PayeOnlineSearchString("//GovTalkMessage/GovTalkDetails/ChannelRouting/Channel/URI"))
    currnode.Text = sURI 'formerly wasn't required for delete_request, now mandatory
    Set currnode = .selectSingleNode(PayeOnlineSearchString("//GovTalkMessage/Header/MessageDetails/Qualifier"))
    currnode.Text = "request"
    Set currnode = .selectSingleNode(PayeOnlineSearchString("//GovTalkMessage/Header/MessageDetails/Function"))
    currnode.Text = "delete"
    Set currnode = .selectSingleNode(PayeOnlineSearchString("//GovTalkMessage/Header/MessageDetails/CorrelationID"))
    currnode.Text = CorrID
    Set currnode = .selectSingleNode(PayeOnlineSearchString("//GovTalkMessage/Header/MessageDetails/Transformation"))
    currnode.Text = S_GG_TRANSFORMATION
    Set currnode = .selectSingleNode(PayeOnlineSearchString("//GovTalkMessage/Header/MessageDetails/GatewayTest"))
    currnode.Text = mGatewayTest
  
    'pq this element seems to require a value if it is present
    'doesn't actually have a purpose
    Set currnode = .selectSingleNode(PayeOnlineSearchString("//GovTalkMessage/Header/SenderDetails/IDAuthentication/Authentication/Method"))
    currnode.Text = "clear"
  End With
   
  If Not MockSubmission Then
    'Post as a DELETE_REQUEST
    
    Set DeleteSubmission = SubmitEx(sResponseText, ResponseEndPoint, DelDoc.xml, False, bCheckProxy)
    
    If Len(DeleteSubmission.xml) = 0 Then
      app.LogEvent "Invalid response from gateway.  Response text: " & vbCrLf & vbCrLf & sResponseText, vbLogEventTypeError
      Err.Raise ERR_GATEWAY, "PostSubmission", "A valid response was not returned from the Gateway, please check that the Gateway is operational.  See event log for details."
    End If
  Else
    ' send delete response back
    Set MockDoc = lows.DOMDocumentNewEFiler()
    MockDoc.loadXML (S_SAMPLEXML_DELETE_RESPONSE)
    Set DeleteSubmission = MockDoc
  End If
  
  
DeleteSubmission_End:
  Set httpreq = Nothing
  Set DelDoc = Nothing
  Set MockDoc = Nothing
  Set currnode = Nothing
  Exit Function
       
DeleteSubmission_Err:
  Err.Raise Err.Number, ErrorSource(Err, "DeleteSubmission"), Err.Description
  Resume
End Function
'removed by CAD as never called !!!!!
'Friend Function RequestStatus(ByVal ID As String, ByVal Password As String) As String
'  Dim xmldoc As DOMDocument
'  Dim ResponseDoc As DOMDocument
'  Dim currnode As IXMLDOMNode
'  Dim httpreq As XMLHTTP
'  Dim sClass As String, sURI As String
'
'  On Error GoTo RequestStatus_Err
'  Set xmldoc = New DOMDocument
'  'Debug.Print xmldoc.loadXML(S_GG_ENVELOPE_START & S_GG_ENVELOPE_END)
'
'  With xmldoc
'    .loadXML (GG_ENVELOPE_KEYS_START & S_GG_ENVELOPE_KEYS_END & S_GG_ENVELOPE_END)
'
'    ' global Government gateway settings, set by constants
'    Set currnode = .selectSingleNode("//GovTalkMessage/EnvelopeVersion")
'    currnode.Text = S_GG_ENVELOPE_VERSION
'    Set currnode = .selectSingleNode("//GovTalkMessage/Header/MessageDetails/Transformation")
'    currnode.Text = S_GG_TRANSFORMATION
'    Set currnode = .selectSingleNode("//GovTalkMessage/Header/SenderDetails/IDAuthentication/Authentication/Method")
'    If TestSubmission Then
'      currnode.Text = "clear"
'    Else
'      currnode.Text = "MD5"
'    End If
'    Set currnode = .selectSingleNode("//GovTalkMessage/Header/MessageDetails/GatewayTest")
'    currnode.Text = mGatewayTest
'    'FIX #50 - was optional
'    Set currnode = .selectSingleNode("//GovTalkDetails/TargetDetails/Organisation")
'    currnode.Text = S_GG_ORGANISATION
'
'    ' global Government gateway settings, set by calling application
'    Set currnode = .selectSingleNode("//GovTalkMessage/Header/SenderDetails/IDAuthentication/SenderID")
'    currnode.Text = ID
'    Set currnode = .selectSingleNode("//GovTalkMessage/Header/SenderDetails/IDAuthentication/Authentication/Value")
'    If TestSubmission Then
'      currnode.Text = Password
'    Else
'      currnode.Text = Base64HashString(Password)
'    End If
'
'    ' global Government gateway settings, set by code
'    'FIX #56
'    If mSubmitType = CorporationTax Then
'      sClass = S_GG_CT_CLASS
'      sURI = S_GG_CT_URI
'    ElseIf mSubmitType = P11D Then
'      sClass = S_GG_P11D_CLASS
'      sURI = S_GG_P11D_URI
'    End If
'    Set currnode = .selectSingleNode("//GovTalkMessage/Header/MessageDetails/Class")
'    currnode.Text = sClass
'    Set currnode = .selectSingleNode("//GovTalkMessage/GovTalkDetails/ChannelRouting/Channel/URI")
'    currnode.Text = sURI 'mandatory if  channelrouting used
'    Set currnode = .selectSingleNode("//GovTalkMessage/Header/MessageDetails/Qualifier")
'    currnode.Text = S_GG_STATUS_REQ_QUALIFIER
'    Set currnode = .selectSingleNode("//GovTalkMessage/Header/MessageDetails/Function")
'    currnode.Text = S_GG_STATUS_REQ_FUNCTION
'  End With
'
'  Set httpreq = New XMLHTTP
'  httpreq.Open "POST", mSubmitAddress, False
'  httpreq.send xmldoc
'  Set ResponseDoc = httpreq.responseXML
'
'  'NEED TO CHECK IF AN ERROR IS RETURNED
'  Set currnode = ResponseDoc.selectSingleNode("//StatusReport")
'  If currnode Is Nothing Then Err.Raise ERR_GATEWAY, "RequestStatus", "The response did not contain a Status Report"
'  RequestStatus = currnode.XML
'
'RequestStatus_End:
'  Set httpreq = Nothing
'  Set ResponseDoc = Nothing
'  Set xmldoc = Nothing
'  Set currnode = Nothing
'  Exit Function
'
'RequestStatus_Err:
'  Err.Raise Err.Number, ErrorSource(Err, "RequestStatus", App.EXEName), Err.Description
'  Resume
'End Function

Private Sub Class_Initialize()
  TestSubmission = False
End Sub

