Attribute VB_Name = "EFilerConst"
Option Explicit

Public Enum ATCEFILER_ERRORS
  ERR_UNKNOWN
  ERR_DSN
  ERR_WRONGKEY
  ERR_REQUIRED_PROPERTY
  ERR_INVALIDID
  ERR_INVALID_XML
  ERR_GATEWAY
  ERR_FAILED_SUBMISSIONS
  ERR_NETWORK
  ERR_INVALID_SCHEMAVERSION
  ERR_CORESETUP
  ERR_EFILERCHECK

End Enum

' SUBMISSION SITE CONSTANTS
' Live Constants

Public Const S_GG_LIVE_SUBMITADDRESS As String = "https://transaction-engine.tax.service.gov.uk/submission" '"https://secure.gateway.gov.uk/submission/ggsubmission.asp"
Public Const S_GG_LIVE_POLLADDRESS As String = "https://transaction-engine.tax.service.gov.uk/poll" '"https://secure.gateway.gov.uk/poll"
Public Const S_GG_LIVE_DELETEADDRESS As String = "https://secure.gateway.gov.uk/submission" 'P11D can't delete submissions
' Test constants
Public Const S_GG_TEST_SUBMITADDRESS As String = "https://test-transaction-engine.tax.service.gov.uk/submission" '"https://secure.dev.gateway.gov.uk/submission/ggsubmission.asp"
Public Const S_GG_TEST_POLLADDRESS As String = "https://test-transaction-engine.tax.service.gov.uk/poll" '"https://secure.dev.gateway.gov.uk/poll"
Public Const S_GG_TEST_DELETEADDRESS As String = "https://secure.dev.gateway.gov.uk/submission" 'P11D can't delete submissions

' CONSTANTS USED BY INLAND REVENUE
Public Const S_IR_ENVELOPE_KEYS_START As String = "<IRenvelope><IRheader><Keys>"
Public Const S_IR_ENVELOPE_KEYS_END_CT2002 As String = "</Keys><PeriodEnd></PeriodEnd><DefaultCurrency></DefaultCurrency><Manifest><Contains><Reference><Namespace></Namespace><SchemaVersion></SchemaVersion><TopElementName></TopElementName></Reference></Contains></Manifest></IRheader>" 'FIX #50 - CT currently uses a specific version of the Core Schema which is different to the one used by P11D
Public Const S_IR_ENVELOPE_KEYS_END_CT2004 As String = "</Keys><PeriodEnd></PeriodEnd><DefaultCurrency></DefaultCurrency><Manifest><Contains><Reference><Namespace></Namespace><SchemaVersion></SchemaVersion><TopElementName></TopElementName></Reference></Contains></Manifest><Sender></Sender></IRheader>"
Public Const S_IR_ENVELOPE_KEYS_END_P11D As String = "</Keys><PeriodEnd></PeriodEnd><DefaultCurrency></DefaultCurrency><Manifest><Contains><Reference><Namespace></Namespace><SchemaVersion></SchemaVersion><TopElementName></TopElementName></Reference></Contains></Manifest><Sender></Sender></IRheader>" 'FIX #50
Public Const S_IR_HEADER_END As String = "</IRenvelope>"
Public Const S_IR_TESTMESSAGE As String = "IR_TESTMESSAGE"  ' OPTIONAL
Public Const S_IR_DEFAULTCURRENCY As String = "GBP"

'CT constants
Public Const S_IR_CT2002_NAMESPACE As String = "http://www.govtalk.gov.uk/taxation/CT"
Public Const S_IR_CT2004_NAMESPACE As String = "http://www.govtalk.gov.uk/taxation/CT/2"
Public Const S_IR_CT_TOPELEMENTNAME As String = "CompanyTaxReturn"
Public Const S_GG_CT_CLASS As String = "IR-CT-CT600"
Public Const S_GG_CT_URI As String = "1003" 'FIX #56

'P11D constants
'Public Const S_IR_P11D_NAMESPACE As String = "http://www.govtalk.gov.uk/taxation/EXB"
Public Const S_IR_P11D_TOPELEMENTNAME As String = "ExpensesAndBenefits"
Public Const S_GG_P11D_CLASS As String = "IR-PAYE-EXB"
Public Const S_GG_P11D_URI As String = "0234" 'FIX #56

'Gateway envelope constants
'Email node removed from S_GG_ENVELOPE_KEYS_START as P11D submissions require that the node must contain a value if it is present.  We don't currently send an email address in any situation
Public Const S_GG_ENVELOPE_KEYS_START As String = "<GovTalkMessage xmlns='http://www.govtalk.gov.uk/CM/envelope'><EnvelopeVersion></EnvelopeVersion><Header><MessageDetails><Class></Class><Qualifier></Qualifier><Function></Function><CorrelationID/><Transformation></Transformation><GatewayTest></GatewayTest><GatewayTimestamp/></MessageDetails><SenderDetails><IDAuthentication><SenderID></SenderID><Authentication><Method></Method><Role></Role><Value></Value></Authentication></IDAuthentication></SenderDetails></Header><GovTalkDetails><Keys>"

Public Const S_GG_ENVELOPE_KEYS_END As String = "</Keys><TargetDetails><Organisation></Organisation></TargetDetails><ChannelRouting><Channel><URI></URI><Product></Product><Version></Version></Channel></ChannelRouting></GovTalkDetails><Body>"
Public Const S_GG_ENVELOPE_END As String = "</Body></GovTalkMessage>"
Public Const S_GG_ENVELOPE_VERSION As String = "2.0" 'FIX #50

Public Const S_GG_TRANSFORMATION As String = "XML" 'format message returned from gateway
Public Const S_GG_X509_CERTIFICATE As String = "X509_CERTIFICATE"  ' OPTIONAL
Public Const S_GG_ORGANISATION As String = "IR"
Public Const S_GG_STATUS_REQ_QUALIFIER As String = "request"
Public Const S_GG_STATUS_REQ_FUNCTION As String = "list"

' status constants
Public Const S_STATUS_NONE As String = ""
Public Const S_STATUS_NOT_SUBMITTED As String = "Submission has not been submitted"
Public Const S_STATUS_ERROR As String = "An error has occurred whilst submitting"
Public Const S_STATUS_ERROR_REQUIRES_DELETE As String = "Error requires to be deleted"
Public Const S_STATUS_SUBMISSION_ACKNOWLEDGEMENT As String = "The submission has been acknowledged by the gateway"
Public Const S_STATUS_SUBMISSION_RESPONSE As String = "The submission has been accepted by the gateway"
Public Const S_STATUS_DELETE_ACKNOWLEDGEMENT As String = "The delete request has been received by the gateway"
Public Const S_STATUS_DELETE_RESPONSE As String = "The submission has been cleaned up by the gateway"
Public Const S_STATUS_DELETE_REQUEST_CLIENT As String = "A Delete request has been made by client"
Public Const S_STATUS_CLIENT_DELETED As String = "The submission has been cleaned up by the client"
Public Const S_STATUS_COMPLETED As String = "Submission has been successfully completed"


'Mock submission xml
Public Const S_SAMPLEXML_SUBMISSION_ACKNOWLEDGEMENT As String = "<GovTalkMessage xmlns=""http://www.govtalk.gov.uk/CM/envelope""><EnvelopeVersion>2.0</EnvelopeVersion><Header><MessageDetails><Class>IR-CT-CT600</Class><Qualifier>acknowledgement</Qualifier><Function>submit</Function><CorrelationID>1DFLG4FFD03MD903SK767687867867</CorrelationID><ResponseEndPoint PollInterval=""2"">https://www.secure.gateway.gov.uk/poll</ResponseEndPoint><GatewayTimestamp>31-01-2001 10:20:18</GatewayTimestamp></MessageDetails><SenderDetails/></Header><GovTalkDetails><Keys/></GovTalkDetails><Body/></GovTalkMessage>"
Public Const S_SAMPLEXML_SUBMISSION_RESPONSE As String = "<GovTalkMessage xmlns=""http://www.govtalk.gov.uk/CM/envelope""><EnvelopeVersion>2.0</EnvelopeVersion><Header><MessageDetails><Class>IR-CT-CT600</Class><Qualifier>response</Qualifier><Function>submit</Function><CorrelationID>ABC123ABC123ABC123</CorrelationID><ResponseEndPoint PollInterval=""2"">https://www.secure.gateway.gov.uk/submission</ResponseEndPoint><GatewayTimestamp>31-01-2001 14:23:34</GatewayTimestamp></MessageDetails><SenderDetails/></Header><GovTalkDetails><Keys/></GovTalkDetails><Body><DepartmentDocument xmlns=""http://www.organisation.gov.uk/namespace""><Data>ABC</Data></DepartmentDocument ></Body></GovTalkMessage>"
Public Const S_SAMPLEXML_DELETE_RESPONSE As String = "<GovTalkMessage xmlns=""http://www.govtalk.gov.uk/CM/envelope""><EnvelopeVersion>2.0</EnvelopeVersion><Header><MessageDetails><Class>IR-CT-CT600</Class><Qualifier>response</Qualifier><Function>delete</Function><CorrelationID>GFKF895473DJ059347DJS</CorrelationID><ResponseEndPoint PollInterval=""2"">https://www.secure.gateway.gov.uk/submission</ResponseEndPoint><GatewayTimestamp>26-01-2001 14:23:52</GatewayTimestamp></MessageDetails><SenderDetails/></Header><GovTalkDetails><Keys/></GovTalkDetails><Body></Body></GovTalkMessage>"

'CT Schema Versions

Public Const S_SCHEMAVERSION_CT2004 As String = "2004-v1.0"


' Global init code goes here

Public Sub CloseRecordSet(ByVal rs As ADODB.Recordset)
  On Error Resume Next
  If Not rs Is Nothing Then rs.Close
End Sub

Public Sub CloseConnection(ByVal cn As ADODB.Connection)
  On Error Resume Next
  If Not cn Is Nothing Then cn.Close
End Sub


