VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Efiler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum SUBMISSION_TYPE
  [_ST_DEFAULT] = 0
  CorporationTax
  P11D
End Enum

Public Enum KEY_TYPE
  [_KT_DEFAULT] = 0
  UTR
  TaxOfficeNumber
  TaxOfficeReference
End Enum

Public Enum SUBMITTER_ROLE
  [_SR_DEFAULT] = 0
  Owner
  Agent
  'Witness
End Enum

Public Enum ARCHIVE_LEVEL
  Complete = 0 ' default
  NoAttachments
  None
End Enum

'FIX #50
Public Enum SENDER_TYPE
  [_SET_DEFAULT] = 0
  SET_COMPANY
  SET_AGENT
  SET_EMPLOYER
End Enum

Public AsyncXMLSubmission As Boolean
Public TestSubmission As Boolean
Public TidySubmissions As Boolean
Public CheckCanSubmitWithAbatec As Boolean
Public CheckProxy As Boolean

Public ErrResponseDoc As DOMDocument
'Public MockSubmission As Boolean

Private m_CompanyName As String
Private mXmlFilename As String
Private mDSN As String
Private mDA As DataAccess
Private mProceedSubmission As Boolean  ' Actually submit to gateway or just put in table.  Set to false if you want to be able to return to app quickly, or if no internet connection
Private mMockSub As Boolean ' do a mock submission, don't actually post to the gateway

' IR properties
Private mPeriodEnd As String
Private mPeriodEnd2 As Date

'Private mTransID As String 'optional
Private mEAddress As String
Private mProd As String
Private mVers As String
Private mSchVers As String
Private mSource As String
Private mSubmitType As SUBMISSION_TYPE
Private mSender As SENDER_TYPE 'FIX #50
Private mRole As SUBMITTER_ROLE
Private mArcLevel As ARCHIVE_LEVEL
Private mTimeOut As Long
Private mAmendedReturn As Boolean

'attachments
Private mAccName As String
Private mAccpdf() As Byte
Private mCompName As String
Private mComppdf() As Byte

'Keys
Private mKeys As Collection
Private m_NotifyIntervalMilliseconds As Long
Private Declare Function InternetGetConnectedState Lib "wininet" (ByRef lpdwFlags As Long, ByVal dwReserved As Long) As Long

Public Property Let XmlFilename(ByVal sNewValue As String)
  mXmlFilename = sNewValue
End Property

Public Property Get DSN() As String
  DSN = mDSN
End Property

Public Property Let DSN(ByVal sNewValue As String)
  If Not IsDSN(sNewValue) Then Err.Raise ERR_DSN, "DSN", "Invalid DSN [" & sNewValue & "]"
  mDSN = sNewValue
End Property

Private Property Get DA() As DataAccess
  If mDA Is Nothing Then
    Set mDA = New DataAccess
    mDA.DSN = mDSN
  End If
  Set DA = mDA
End Property

Public Property Get PeriodEnd() As String
  PeriodEnd = mPeriodEnd
End Property
Public Property Let PeriodEnd(ByVal sNewValue As String)
  'needs validation and/or datatype conversion.  Perhaps should be passed in as a date
  mPeriodEnd = CDate(sNewValue)
End Property
Public Property Get PeriodEnd2() As Date
  PeriodEnd = mPeriodEnd2
End Property
Public Property Let PeriodEnd2(ByVal NewValue As Date)
  mPeriodEnd2 = NewValue
  mPeriodEnd = DateString2(mPeriodEnd2, "", "YYYY-MM-DD")
End Property

Public Property Get ProceedSubmission() As Boolean
  ProceedSubmission = mProceedSubmission
End Property

Public Property Let ProceedSubmission(ByVal bNewValue As Boolean)
  mProceedSubmission = bNewValue
End Property

'could use TransactionID element to record SubmissionID
'Public Property Get TransactionID() As String
'  TransactionID = mTransID
'End Property
'
'Public Property Let TransactionID(ByVal sNewValue As String)
'  mTransID = sNewValue
'End Property

'DO NOT USE - left in for backwards compat
Public Property Get EmailAddress() As String
Attribute EmailAddress.VB_Description = "'DO NOT USE - left in for backwards compat"
  EmailAddress = mEAddress
End Property

'DO NOT USE - left in for backwards compat
Public Property Let EmailAddress(ByVal sNewValue As String)
  mEAddress = sNewValue
End Property

Public Property Get Product() As String
  Product = mProd
End Property

Public Property Let Product(ByVal sNewValue As String)
  mProd = sNewValue
End Property

Public Property Get ProductVersion() As String
  ProductVersion = mVers
End Property

Public Property Let ProductVersion(ByVal sNewValue As String)
  mVers = sNewValue
End Property

Public Property Get SchemaVersion() As String
  SchemaVersion = mSchVers
End Property

Public Property Let SchemaVersion(ByVal sNewValue As String)
  mSchVers = sNewValue
End Property

Public Property Get Source() As String
  Source = mSource
End Property

Public Property Let Source(ByVal sNewValue As String)
  mSource = sNewValue
End Property

Public Property Get SubmissionType() As SUBMISSION_TYPE
  SubmissionType = mSubmitType
End Property

Public Property Let SubmissionType(ByVal NewValue As SUBMISSION_TYPE)
  mSubmitType = NewValue
End Property

'FIX #50
Public Property Get Sender() As SENDER_TYPE
  Sender = mSender
End Property

'FIX #50
Public Property Let Sender(ByVal NewValue As SENDER_TYPE)
  mSender = NewValue
End Property

Public Property Get SubmitterRole() As SUBMITTER_ROLE
  SubmitterRole = mRole
End Property

Public Property Let SubmitterRole(ByVal NewValue As SUBMITTER_ROLE)
  mRole = NewValue
End Property

Public Property Get ArchiveLevel() As ARCHIVE_LEVEL
  ArchiveLevel = mArcLevel
End Property

Public Property Let ArchiveLevel(ByVal NewValue As ARCHIVE_LEVEL)
  mArcLevel = NewValue
End Property

Public Property Get TimeOut() As Long
  TimeOut = mTimeOut
End Property

Public Property Let TimeOut(ByVal lNewValue As Long)
  mTimeOut = lNewValue
End Property

Public Property Get MockSubmission() As Boolean
  MockSubmission = mMockSub
End Property

Public Property Let MockSubmission(ByVal bNewValue As Boolean)
  mMockSub = bNewValue
End Property


Public Sub AddKey(ByVal sKey As String, ByVal kType As KEY_TYPE)
  Dim s As String, KeyType As String
  Dim i As Integer
  
  On Error GoTo AddKey_Err
  
  'checks for correct key types for appropriate classes
  If mSubmitType = P11D Then
    If kType = TaxOfficeNumber Then
      KeyType = "TaxOfficeNumber"
      s = "<Key Type=""" & KeyType & """>" & sKey & "</Key>"
    ElseIf kType = TaxOfficeReference Then
      KeyType = "TaxOfficeReference"
      s = "<Key Type=""" & KeyType & """>" & sKey & "</Key>"
    Else
      Err.Raise ERR_WRONGKEY, "AddKey", "You are adding an incorrect key type"
    End If
  Else
    If kType = UTR Then
      KeyType = "UTR"
      s = "<Key Type=""" & KeyType & """>" & sKey & "</Key>"
    Else
      Err.Raise ERR_WRONGKEY, "AddKey", "You are adding an incorrect key type"
    End If
  End If
  
 'check if key exists in collection before adding it in
  For i = 1 To mKeys.Count
    If Not s = "" Then
      If StrComp(s, mKeys.Item(i), vbTextCompare) = 0 Then
        Err.Raise ERR_WRONGKEY, "AddKey", "The " & KeyType & " " & sKey & " has already been added. Please ensure you are entering the correct " & KeyType & "."
        Exit Sub
      End If
    End If
  Next i
  
  mKeys.Add s
  
AddKey_End:
  Exit Sub

AddKey_Err:
  Err.Raise Err.Number, ErrorSourceComponent(Err, "AddKey", App.EXEName), Err.Description
  Resume
End Sub
Private Sub CheckIEVersion()
  
'versions of schdocvw dll as at 05/05/2004 CAD
'  5.50.3825.1300  Internet Explorer 5.5 Developer Preview (Beta)
'  5.50.4030.2400  Internet Explorer 5.5 & Internet Tools Beta
'  5.50.4134.0100  Internet Explorer 5.5 for Windows Me (4.90.3000)
'  5.50.4134.0600  Internet Explorer 5.5
'  5.50.4308.2900  Internet Explorer 5.5 Advanced Security Privacy Beta
'  5.50.4522.1800  Internet Explorer 5.5 Service Pack 1
'  5.50.4807.2300  Internet Explorer 5.5 Service Pack 2
'  6.00.2462.0000  Internet Explorer 6 Public Preview (Beta)
'  6.00.2479.0006  Internet Explorer 6 Public Preview (Beta) Refresh
'  6.00.2600.0000  Internet Explorer 6 (Windows XP)
'  6.00.2800.1106  Internet Explorer 6 Service Pack 1 (Windows XP SP1)
'  6.00.2800.1278  Internet Explorer 6 Update v.01 Developer Preview (SP1b Beta)
'  6.0.2800.1314 Internet Explorer 6 Update v.04 Developer Preview (SP1b Beta)
'  6.00.3663.0000  Internet Explorer 6 for Microsoft Windows Server 2003 RC1
'  6.00.3718.0000  Internet Explorer 6 for Windows Server 2003 RC2
'  6.00.3790.0000  Internet Explorer 6 for Windows Server 2003 (released)

  Dim sSHDovVWDll As String
  Dim sVersion As String
  Dim sMessage As String
  Dim v As Variant
  Dim iCount As Long, i As Long
  Const SHDOCVW As String = "Shdocvw.dll"
  On Error GoTo err_err

  Dim s As String
  
  sSHDovVWDll = FullPath(GetSysDirectory) & SHDOCVW
  sMessage = ""
  If (Not FileExists(sSHDovVWDll)) Then
    sMessage = "Failed to find the file " & SHDOCVW & ". "
  Else
    sVersion = FileVersionQuery(sSHDovVWDll)
    If (Len(sVersion) = 0) Then
      sMessage = "Failed to determine the version of the file " & sSHDovVWDll & "."
    Else
      iCount = GetDelimitedValues(v, sVersion, True, True, ".")
      If (iCount) < 2 Then
        sMessage = "The version number of " & sSHDovVWDll & " is in an unrecognised format."
      End If
    
      If (CLng(v(1)) < 5) Or (CLng(v(1)) = 5 And (CLng(v(2)) < 50)) Then
        sMessage = "Your version of Internet Explorer is " & sVersion & "." & vbCrLf
        sMessage = sMessage & "For electronic filing the version must be 5.5 or greater." & vbCrLf & vbCrLf
        sMessage = sMessage & "Please go to http://www.microsoft.com and search for 'Internet Explorer'" & vbCrLf
        sMessage = sMessage & "or contact your local system administrator for help to upgrade."
      End If
    End If
  End If

  If (Len(sMessage) > 0) Then
    sMessage = sMessage & vbCrLf & vbCrLf & "Click OK to attempt submission or Cancel to exit."
    If MsgBox(sMessage, vbOKCancel Or vbQuestion, "Internet Explorer check") = vbCancel Then
       Call Err.Raise(ERR_REQUIRED_PROPERTY, "CheckIEVersion", "Submission cancelled due to failed Internet Explorer")
    End If
  End If
  
err_end:
  Exit Sub
err_err:
  Call Err.Raise(Err.Number, ErrorSourceComponent(Err, "CheckIEVersion", App.EXEName), Err.Description)
End Sub

'SHOULD NOT BE USED - left in for backwards compat with P11D
Public Function Submit2(ByRef SubmissionID As String, ByVal XMLBody As String, ByVal ID As String, ByVal Password As String, ByVal EntityID As String, ByVal EntityName As String, ByVal Username As String, Optional ByVal iNotify As IBaseNotify) As String
Attribute Submit2.VB_Description = "DO NOT USE - left in for backwards compat with P11D"
  
  Submit2 = Submit(SubmissionID, XMLBody, ID, Password, EntityID, EntityName, Username, "", iNotify)
End Function
Public Property Get CompanyName() As String
  CompanyName = m_CompanyName
End Property
Public Property Let CompanyName(ByVal NewValue As String)
  m_CompanyName = NewValue
End Property


Public Function Submit(ByRef SubmissionID As String, ByVal XMLBody As String, ByVal ID As String, ByVal Password As String, ByVal EntityID As String, ByVal EntityName As String, ByVal Username As String, Optional ByVal Email As String, Optional ByVal iNotify As IBaseNotify) As String
  Dim xmlDoc As DOMDocument
  Dim s As String, SubID As String, xmlFn As String
  Dim node As IXMLDOMNode
  On Error GoTo Submit_Err
  
  'NEED TO CHECK THAT ALL REQUIRED PROPERTIES HAVE BEEN SET BEFORE PROCEEDING.  ELSE RAISE ERROR PQ
  
  
  If Len(mPeriodEnd) = 0 Then Err.Raise ERR_REQUIRED_PROPERTY, "Submit", "PeriodEnd is a required property and has not been set."
  
  If Len(mDSN) = 0 Then Err.Raise ERR_REQUIRED_PROPERTY, "Submit", "DSN is a required property and has not been set."
  If Len(mSource) = 0 Then Err.Raise ERR_REQUIRED_PROPERTY, "Submit", "Source is a required property and has not been set."
  If Len(mSchVers) = 0 Then Err.Raise ERR_REQUIRED_PROPERTY, "Submit", "Schema Version is a required property and has not been set."
  If mSubmitType = [_ST_DEFAULT] Then Err.Raise ERR_REQUIRED_PROPERTY, "Submit", "Submission Type is a required property and has not been set."
  If mRole = [_SR_DEFAULT] Then Err.Raise ERR_REQUIRED_PROPERTY, "Submit", "Submitter Role is a required property and has not been set."
  If (mSubmitType = P11D And mSender = [_SET_DEFAULT]) Then Err.Raise ERR_REQUIRED_PROPERTY, "Submit", "Sender is a required property and has not been set." 'FIX #50
  If Len(Username) = 0 Then Err.Raise ERR_REQUIRED_PROPERTY, "Submit", "Username is a required property and has not been set."
      
  'check there is a network connection
  If mProceedSubmission Then Call CheckNetwork 'FIX #865

  Call CheckIEVersion
  If (Me.CheckCanSubmitWithAbatec) Then
    Call CheckCanSubmitWithAbatecEx(mProd, mVers, m_CompanyName)
  End If
    
  'check if there is an existing submission of uncertain status
  If Me.TidySubmissions Then
    If TidyExistingSubmissions(EntityID, iNotify) Then
      GoTo DEFAULT:
    End If
    s = "The submission has been cancelled owing to an unresolved earlier submission.  Please rectify this before attempting to submit again."
  Else
DEFAULT:
    Set xmlDoc = New DOMDocument
    If Not xmlDoc.loadXML(XMLBody) Then Err.Raise ERR_INVALID_XML, , "The xml supplied is not well formed."
    SubmissionID = ""
      'proceed with this submission
    If DA.IsAmendedSubmission(EntityID) Then mAmendedReturn = True
    If Len(SchemaVersion) = 0 Then Call Err.Raise(ERR_INVALID_SCHEMAVERSION, "Submit", "Blank schema version.")
    If Me.SubmissionType = CorporationTax Then
      If CLng(Left$(SchemaVersion, 4)) >= 2004 Then 'ct 2004 version = 2004-v1.0
        Set xmlDoc = WrapData_CT2004(xmlDoc, ID, Password)
      Else
        Set xmlDoc = WrapData(xmlDoc, ID, Password)
      End If
    Else
      Set xmlDoc = WrapData(xmlDoc, ID, Password)
    End If
    If Me.SubmissionType = P11D Then
      'remove the maifest node as not required
      Set node = xmlDoc.selectSingleNode("//Manifest")
      If Not node Is Nothing Then
        Call node.parentNode.removeChild(node)
      End If
    End If
    Set xmlDoc = Encrypt(xmlDoc)
    If (mPeriodEnd2 = UNDATED) Then
      SubID = DA.NewSubmission(EntityID, EntityName, mPeriodEnd, xmlDoc.xml, Username, Source, Email)
    Else
      SubID = DA.NewSubmission(EntityID, EntityName, mPeriodEnd2, xmlDoc.xml, Username, Source, Email)
    End If
    SubmissionID = SubID
    If mProceedSubmission Then
      
      s = ProcessSubmission(SubID, iNotify)
      If mSubmitType = CorporationTax Then
        Call SaveSubmission(SubID, xmlDoc)
      End If
    Else
      s = "Submission has been initialised"
    End If
  End If

Submit_End:
  Submit = s
  Set xmlDoc = Nothing
  Exit Function

Submit_Err:
  Err.Raise Err.Number, ErrorSourceComponent(Err, "Submit", App.EXEName), Err.Description
  Resume
End Function
Private Sub SaveSubmission(ByVal SubID As String, xmlDoc As DOMDocument)
Dim xmlSubmissionName As String
Dim rootNode As IXMLDOMNode
Dim pi As IXMLDOMProcessingInstruction

On Error GoTo SaveSubmission_Err

xmlSubmissionName = mXmlFilename & SubID & ".xml"
Set rootNode = xmlDoc.selectSingleNode("//GovTalkMessage")
Set pi = xmlDoc.createProcessingInstruction("xml-stylesheet", "type=""text/xsl"" href=""ss.xsl""")
Call xmlDoc.insertBefore(pi, rootNode)
Call xmlDoc.Save(xmlSubmissionName)

SaveSubmission_End:
Exit Sub
  
SaveSubmission_Err:
  Err.Raise Err.Number, ErrorSourceComponent(Err, "SaveSubmission", App.EXEName), Err.Description
  Resume
End Sub

Private Sub CheckCanSubmitWithAbatecEx(ByVal sApp As String, ByVal sVersion As String, ByVal CompanyName As String)
  Dim xml As DOMDocument
  Dim node As IXMLDOMNode
  Dim sURL As String
  Dim s As String
  Dim sError As String
  Dim sResponseText As String
  
On Error GoTo err_err
  
    
  sError = "Unable to verify whether the Revenue Online services connection is available." & vbCrLf
  sError = sError & "If you would like any additional information, please contact the " & sApp & " support team."
  
  Set xml = New DOMDocument
  
  s = "<?xml version='1.0'?>" & vbCrLf
  s = s & "<automatic_updates>" & vbCrLf
    s = s & "<function>" & vbCrLf
      s = s & "<name>" & vbCrLf
      s = s & "online_submission_valid" & vbCrLf
      s = s & "</name>" & vbCrLf
      s = s & "<parameters>" & vbCrLf
        s = s & "<application>" & vbCrLf
        s = s & sApp & vbCrLf
        s = s & "</application>" & vbCrLf
        s = s & "<version>" & vbCrLf
        s = s & sVersion & vbCrLf
        s = s & "</version>" & vbCrLf
        s = s & "<company_name>" & vbCrLf
        s = s & CompanyName & vbCrLf
        s = s & "</company_name>" & vbCrLf
      s = s & "</parameters>" & vbCrLf
    s = s & "</function>"
  s = s & "</automatic_updates>" & vbCrLf


  Call xml.loadXML(s)
  
  Set xml = SubmitEx(sResponseText, "http://www.abatec.deloitte.co.uk/AutomaticUpdates/default.asp", s, False, CheckProxy)
  If xml.parseError.errorCode <> 0 Then
    Call Err.Raise(ERR_FAILED_SUBMISSIONS, "CheckCanSubmitWithAbatecEx", sError)
  End If
  
  Set node = xml.selectSingleNode("automatic_updates/error")
  If Not (node Is Nothing) Then
    Call Err.Raise(ERR_FAILED_SUBMISSIONS, "CheckCanSubmitWithAbatecEx", node.Text)
  End If
  
  Set node = xml.selectSingleNode("automatic_updates/return")
  If (node Is Nothing) Then
    Call Err.Raise(ERR_FAILED_SUBMISSIONS, "CheckCanSubmitWithAbatecEx", sError)
  Else
    If (node.Text <> "1") Then
      Call Err.Raise(ERR_FAILED_SUBMISSIONS, "CheckCanSubmitWithAbatecEx", node.Text)
    End If
  End If
  
err_end:
  Exit Sub
err_err:
  If (Err.Number <> ERR_FAILED_SUBMISSIONS) Then
    Call Err.Raise(Err.Number, ErrorSourceComponent(Err, "CheckCanSubmitWithAbatecEx", App.EXEName), sError & vbCrLf & Err.Description)
  Else
    Call Err.Raise(Err.Number, ErrorSourceComponent(Err, "CheckCanSubmitWithAbatecEx", App.EXEName), Err.Description)
  End If
  Resume
End Sub
     

     
Private Function WrapData(ByVal XMLBody As DOMDocument, ByVal ID As String, ByVal Password As String) As DOMDocument
  Dim xmlDoc As DOMDocument
  Dim currnode As IXMLDOMNode
  Dim i As Integer, s As String, sKeys As String
  Dim nSpace As String, tElement As String, sClass As String, sURI As String
  Dim currEl As IXMLDOMElement
        
      
  On Error GoTo WrapData_Err
         
  'Wrap message body with GG Envelope incorporating the authentication and the IR Header
  
  'multiple keys
  For i = 1 To mKeys.Count
    s = s & mKeys(i)
  Next i
  sKeys = s
  
  Set xmlDoc = New DOMDocument
  xmlDoc.loadXML (S_GG_ENVELOPE_KEYS_START & sKeys & S_GG_ENVELOPE_KEYS_END & S_IR_ENVELOPE_KEYS_START & sKeys & IIf(mSubmitType = CorporationTax, S_IR_ENVELOPE_KEYS_END_CT2002, S_IR_ENVELOPE_KEYS_END_P11D) & Trim(XMLBody.xml) & S_IR_HEADER_END & S_GG_ENVELOPE_END)  'FIX #50
  
  With xmlDoc
    'create multiple elements at this point pq
        
    ' global Government gateway settings, set by constants
    Set currnode = .selectSingleNode("//GovTalkMessage/EnvelopeVersion")
    currnode.Text = S_GG_ENVELOPE_VERSION
    Set currnode = .selectSingleNode("//GovTalkMessage/Header/MessageDetails/Transformation")
    currnode.Text = S_GG_TRANSFORMATION
    Set currnode = .selectSingleNode("//" & "GovTalkMessage/Body/IRenvelope/IRheader/DefaultCurrency") ' optional
    currnode.Text = S_IR_DEFAULTCURRENCY
    'FIX #50 - was optional
    Set currnode = .selectSingleNode("//GovTalkDetails/TargetDetails/Organisation")
    currnode.Text = S_GG_ORGANISATION
    
    ' set by code
    Set currnode = .selectSingleNode("//GovTalkMessage/Header/SenderDetails/IDAuthentication/Authentication/Method")
    
    currnode.Text = "MD5"
    
    Set currnode = .selectSingleNode("//GovTalkMessage/Header/MessageDetails/GatewayTest")
    If TestSubmission Then
      currnode.Text = 1
    Else
      currnode.Text = 0
    End If
    
    'set nodes according to submission type, by code SECTION.
    If mSubmitType = CorporationTax Then
      nSpace = S_IR_CT2002_NAMESPACE
      tElement = S_IR_CT_TOPELEMENTNAME
      sClass = S_GG_CT_CLASS
      sURI = S_GG_CT_URI
    ElseIf mSubmitType = P11D Then
      nSpace = S_IR_P11D_NAMESPACE
      tElement = S_IR_P11D_TOPELEMENTNAME
      sClass = S_GG_P11D_CLASS
      sURI = S_GG_P11D_URI
    End If
    
    Set currnode = .selectSingleNode("//" & "GovTalkMessage/Body/IRenvelope/IRheader/Manifest/Contains/Reference/Namespace")
    currnode.Text = nSpace
    Set currEl = .selectSingleNode("//" & "GovTalkMessage/Body/IRenvelope")
    Call currEl.setAttribute("xmlns", nSpace)
    Set currnode = .selectSingleNode("//" & "GovTalkMessage/Body/IRenvelope/IRheader/Manifest/Contains/Reference/TopElementName")
    currnode.Text = tElement
    Set currnode = .selectSingleNode("//GovTalkMessage/Header/MessageDetails/Class")
    currnode.Text = sClass
    'FIX #56
    Set currnode = .selectSingleNode("//GovTalkMessage/GovTalkDetails/ChannelRouting/Channel/URI")
    currnode.Text = sURI 'mandatory if  channelrouting used
          
    ' global Government gateway settings, set by properties
    Set currnode = .selectSingleNode("//GovTalkMessage/GovTalkDetails/ChannelRouting/Channel/Version") ' optional
    currnode.Text = mVers
    Set currnode = .selectSingleNode("//GovTalkMessage/GovTalkDetails/ChannelRouting/Channel/Product") ' optional
    currnode.Text = mProd
    Set currnode = .selectSingleNode("//GovTalkMessage/Body/IRenvelope/IRheader/Manifest/Contains/Reference/SchemaVersion")
    currnode.Text = mSchVers
    Set currnode = .selectSingleNode("//GovTalkMessage/Header/SenderDetails/IDAuthentication/Authentication/Role")
    If mRole = Agent Then
      currnode.Text = "intermediary"
    ElseIf mRole = Owner Then
      currnode.Text = "principal"
    'ElseIf mRole = Witness Then
    '  currnode.Text = "witness"
    End If
    
    ' global Government gateway settings, set by calling application
    Set currnode = .selectSingleNode("//GovTalkMessage/Header/SenderDetails/IDAuthentication/SenderID")
    currnode.Text = ID
    Set currnode = .selectSingleNode("//GovTalkMessage/Header/SenderDetails/IDAuthentication/Authentication/Value")
    
    currnode.Text = Base64HashString(LCase$(Password))
    
    ' global Government gateway settings, set by code
    Set currnode = .selectSingleNode("//GovTalkMessage/Header/MessageDetails/Qualifier")
    currnode.Text = "request"
    Set currnode = .selectSingleNode("//GovTalkMessage/Header/MessageDetails/Function")
    currnode.Text = "submit"
        
    ' global Inland revenue settings, set by calling application
    Set currnode = .selectSingleNode("//" & "GovTalkMessage/Body/IRenvelope/IRheader/PeriodEnd")
    
    If (mPeriodEnd2 = UNDATED) Then
      currnode.Text = DateString2(mPeriodEnd, "", "YYYY-MM-DD")  'CAD this was as it was before, so left
    Else
      currnode.Text = DateString2(mPeriodEnd2, "", "YYYY-MM-DD")
    End If
    
    If mSubmitType = P11D Then 'FIX #50 - P11D currently uses a different schema to CT
      Set currnode = .selectSingleNode("//GovTalkMessage/Body/IRenvelope/IRheader/Sender")
      If mSender = SET_AGENT Then
        currnode.Text = "Agent"
      ElseIf mSender = SET_COMPANY Then
        currnode.Text = "Company"
      ElseIf mSender = SET_EMPLOYER Then
        currnode.Text = "Employer"
      End If
    End If
    
    'optional elements, commented for now.
'    Set currnode = .selectSingleNode("//GovTalkMessage/Header/SenderDetails/X509Certificate")
'    currnode.Text = S_GG_X509_CERTIFICATE
'    Set currnode = .selectSingleNode("//" & "GovTalkMessage/Header/MessageDetails/TransactionID") ' optional
'    currnode.Text = mTransID 'optional

    
    ' set CT-specific properties
    If mSubmitType = CorporationTax Then
      Set currnode = .selectSingleNode("//CompanyTaxReturn/@ReturnType")
      If currnode Is Nothing Then Err.Raise ERR_INVALID_XML, "WrapData", "ReturnType attribute missing from CompanyTaxReturn node"
      If mAmendedReturn Then
        currnode.Text = "amended"
      Else
        currnode.Text = "new"
      End If
      'attach files
      If Len(mAccName) > 0 Then
        Set currnode = .selectSingleNode("//CompanyTaxReturn/AttachedFiles/Attachment[@Type = ""accounts""]")
        currnode.Text = Base64EncodeByte(mAccpdf)
        Set currnode = currnode.selectSingleNode("./@Filename")
        currnode.Text = mAccName
      End If
      If Len(mCompName) > 0 Then
        Set currnode = .selectSingleNode("//CompanyTaxReturn/AttachedFiles/Attachment[@Type= ""computations""]")
        currnode.Text = Base64EncodeByte(mComppdf)
        Set currnode = currnode.selectSingleNode("./@Filename")
        currnode.Text = mCompName
      End If
    End If
    
  End With
  
  Set WrapData = xmlDoc
  
  'create a tmp file to put output to.  FOR DEBUG ONLY
'  Dim tempfile
'  tempfile = CreateTempFile(xmldoc.XML)
                

             
WrapData_End:
  Set xmlDoc = Nothing
  Set currnode = Nothing
  Set currEl = Nothing
  Exit Function
  
WrapData_Err:
  Err.Raise Err.Number, ErrorSourceComponent(Err, "WrapData", App.EXEName), Err.Description
  Resume
End Function
Private Function WrapData_CT2004(ByVal XMLBody As DOMDocument, ByVal ID As String, ByVal Password As String) As DOMDocument
  Dim xmlDoc As DOMDocument
  Dim currnode As IXMLDOMNode
  Dim i As Integer, s As String, sKeys As String
  Dim nSpace As String, tElement As String, sClass As String, sURI As String
  Dim currEl As IXMLDOMElement
        
      
  On Error GoTo WrapData_CT2004_Err
         
  'Wrap message body with GG Envelope incorporating the authentication and the IR Header
  
  'multiple keys
  For i = 1 To mKeys.Count
    s = s & mKeys(i)
  Next i
  sKeys = s
  
  Set xmlDoc = New DOMDocument
  xmlDoc.loadXML (S_GG_ENVELOPE_KEYS_START & sKeys & S_GG_ENVELOPE_KEYS_END & S_IR_ENVELOPE_KEYS_START & sKeys & IIf(mSubmitType = CorporationTax, S_IR_ENVELOPE_KEYS_END_CT2004, S_IR_ENVELOPE_KEYS_END_P11D) & Trim(XMLBody.xml) & S_IR_HEADER_END & S_GG_ENVELOPE_END)  'FIX #50
  'Call xmldoc.setProperty("SelectionLanguage", "XPath")
  With xmlDoc
    'create multiple elements at this point pq
        
    ' global Government gateway settings, set by constants
    Set currnode = .selectSingleNode("//GovTalkMessage/EnvelopeVersion")
    currnode.Text = S_GG_ENVELOPE_VERSION
    Set currnode = .selectSingleNode("//GovTalkMessage/Header/MessageDetails/Transformation")
    currnode.Text = S_GG_TRANSFORMATION
    Set currnode = .selectSingleNode("//" & "GovTalkMessage/Body/IRenvelope/IRheader/DefaultCurrency") ' optional
    currnode.Text = S_IR_DEFAULTCURRENCY
    'FIX #50 - was optional
    Set currnode = .selectSingleNode("//GovTalkDetails/TargetDetails/Organisation")
    currnode.Text = S_GG_ORGANISATION
    
    ' set by code
    Set currnode = .selectSingleNode("//GovTalkMessage/Header/SenderDetails/IDAuthentication/Authentication/Method")
    currnode.Text = "MD5"
    Set currnode = .selectSingleNode("//GovTalkMessage/Header/MessageDetails/GatewayTest")
    If TestSubmission Then
      currnode.Text = 1
    Else
      currnode.Text = 0
    End If
    
    'set nodes according to submission type, by code SECTION.
    If mSubmitType = CorporationTax Then
      nSpace = S_IR_CT2004_NAMESPACE
      tElement = S_IR_CT_TOPELEMENTNAME
      sClass = S_GG_CT_CLASS
      sURI = S_GG_CT_URI
    ElseIf mSubmitType = P11D Then
      nSpace = S_IR_P11D_NAMESPACE
      tElement = S_IR_P11D_TOPELEMENTNAME
      sClass = S_GG_P11D_CLASS
      sURI = S_GG_P11D_URI
    End If
    
    Set currnode = .selectSingleNode("//" & "GovTalkMessage/Body/IRenvelope/IRheader/Manifest/Contains/Reference/Namespace")
    currnode.Text = nSpace
    Set currEl = .selectSingleNode("//" & "GovTalkMessage/Body/IRenvelope")
    Call currEl.setAttribute("xmlns", nSpace)
    Set currnode = .selectSingleNode("//" & "GovTalkMessage/Body/IRenvelope/IRheader/Manifest/Contains/Reference/TopElementName")
    currnode.Text = tElement
    Set currnode = .selectSingleNode("//GovTalkMessage/Header/MessageDetails/Class")
    currnode.Text = sClass
    'FIX #56
    Set currnode = .selectSingleNode("//GovTalkMessage/GovTalkDetails/ChannelRouting/Channel/URI")
    currnode.Text = sURI 'mandatory if  channelrouting used
          
    ' global Government gateway settings, set by properties
    Set currnode = .selectSingleNode("//GovTalkMessage/GovTalkDetails/ChannelRouting/Channel/Version") ' optional
    currnode.Text = mVers
    Set currnode = .selectSingleNode("//GovTalkMessage/GovTalkDetails/ChannelRouting/Channel/Product") ' optional
    currnode.Text = mProd
    Set currnode = .selectSingleNode("//GovTalkMessage/Body/IRenvelope/IRheader/Manifest/Contains/Reference/SchemaVersion")
    currnode.Text = mSchVers
    Set currnode = .selectSingleNode("//GovTalkMessage/Header/SenderDetails/IDAuthentication/Authentication/Role")
    If mRole = Agent Then
      currnode.Text = "intermediary"
    ElseIf mRole = Owner Then
      currnode.Text = "principal"
    'ElseIf mRole = Witness Then
    '  currnode.Text = "witness"
    End If
    
    ' global Government gateway settings, set by calling application
    Set currnode = .selectSingleNode("//GovTalkMessage/Header/SenderDetails/IDAuthentication/SenderID")
    currnode.Text = ID
    Set currnode = .selectSingleNode("//GovTalkMessage/Header/SenderDetails/IDAuthentication/Authentication/Value")
    
    currnode.Text = Base64HashString(Password)
    
    ' global Government gateway settings, set by code
    Set currnode = .selectSingleNode("//GovTalkMessage/Header/MessageDetails/Qualifier")
    currnode.Text = "request"
    Set currnode = .selectSingleNode("//GovTalkMessage/Header/MessageDetails/Function")
    currnode.Text = "submit"
        
    ' global Inland revenue settings, set by calling application
    Set currnode = .selectSingleNode("//" & "GovTalkMessage/Body/IRenvelope/IRheader/PeriodEnd")
    
    If (mPeriodEnd2 = UNDATED) Then
      currnode.Text = DateString2(mPeriodEnd, "", "YYYY-MM-DD")  'CAD this was as it was before, so left
    Else
      currnode.Text = DateString2(mPeriodEnd2, "", "YYYY-MM-DD")
    End If
    
    'Sender has been added to CT2004 Schema
    Set currnode = .selectSingleNode("//GovTalkMessage/Body/IRenvelope/IRheader/Sender")
    If mSender = SET_AGENT Then
      currnode.Text = "Agent"
    ElseIf mSender = SET_COMPANY Then
      currnode.Text = "Company"
    ElseIf mSender = SET_EMPLOYER Then
      currnode.Text = "Employer"
    End If
    
    
    'optional elements, commented for now.
'    Set currnode = .selectSingleNode("//GovTalkMessage/Header/SenderDetails/X509Certificate")
'    currnode.Text = S_GG_X509_CERTIFICATE
'    Set currnode = .selectSingleNode("//" & "GovTalkMessage/Header/MessageDetails/TransactionID") ' optional
'    currnode.Text = mTransID 'optional

    
    ' set CT-specific properties
    If mSubmitType = CorporationTax Then
      'If the return type is "amended", elements marked with attribute @efiler='removeifamended'
      'will have to be removed from the submission
      Dim removeNodes As IXMLDOMNodeList
      Dim parentNode As IXMLDOMNode
      Set removeNodes = .selectNodes("//*[@eFiler='removeifamended']")
      For Each currnode In removeNodes
        Set parentNode = currnode.parentNode
        If mAmendedReturn Then
          Call currnode.parentNode.removeChild(currnode)
        Else
          currnode.Attributes.removeNamedItem ("eFiler")
        End If
      Next
      Set currnode = .selectSingleNode("//CompanyTaxReturn/@ReturnType")
      If currnode Is Nothing Then Err.Raise ERR_INVALID_XML, "WrapData_CT2004", "ReturnType attribute missing from CompanyTaxReturn node"
      If mAmendedReturn Then
        currnode.Text = "amended"
      Else
        currnode.Text = "new"
      End If
      'attach files
      
      If Len(mAccName) > 0 Then
        Set currnode = .selectSingleNode("//CompanyTaxReturn/AttachedFiles/Attachment[@Type = ""accounts""]")
        currnode.Text = Base64EncodeByte(mAccpdf)
        Set currnode = currnode.selectSingleNode("./@Filename")
        currnode.Text = mAccName
      End If
      If Len(mCompName) > 0 Then
        Set currnode = .selectSingleNode("//CompanyTaxReturn/AttachedFiles/Attachment[@Type= ""computations""]")
        currnode.Text = Base64EncodeByte(mComppdf)
        Set currnode = currnode.selectSingleNode("./@Filename")
        currnode.Text = mCompName
      End If
    End If
    
  End With
  
  Set WrapData_CT2004 = xmlDoc
  
  'create a tmp file to put output to.  FOR DEBUG ONLY
'  Dim tempfile
'  tempfile = CreateTempFile(xmldoc.XML)
                

             
WrapData_CT2004_End:
  Set xmlDoc = Nothing
  Set currnode = Nothing
  Set currEl = Nothing
  Exit Function
  
WrapData_CT2004_Err:
  Err.Raise Err.Number, ErrorSourceComponent(Err, "WrapData_CT2004", App.EXEName), Err.Description
  Resume
End Function
     

Private Function Encrypt(ByVal xmlDoc As DOMDocument) As DOMDocument

On Error GoTo Encrypt_Err
  Set Encrypt = xmlDoc
     
Encrypt_End:
  Exit Function
       
Encrypt_Err:
  Err.Raise Err.Number, ErrorSourceComponent(Err, "Encrypt", App.EXEName), Err.Description
  Resume
End Function
     
'changed from private to test this function as site was down, must resume after testing
Private Function ErrorResponseMsg(ByVal ResponseDoc As DOMDocument, ByRef TryAgain As Boolean, ByRef RequiresDelete As Boolean, Optional ByVal iNotify As IBaseNotify) As String
  Dim currnode As IXMLDOMNode
  Dim s As String, bFatal As Boolean
  Dim MsgText As String, DelText As String, ErrMsgType As String
  Dim ResponseEndPoint As String, CorrID As String
  Dim PollInterval As Integer, Wait As Long
  Dim ErrList As IXMLDOMNodeList, ErrNode As IXMLDOMNode
  Dim ErrRaisedby As String, ErrNum As String, ErrType As String
  Dim Errtext As String, ErrLocation As String, ErrDetails As String
  

On Error GoTo ErrorResponseMsg_Err
  Set currnode = ResponseDoc.selectSingleNode("//" & "GovTalkMessage/GovTalkDetails/GovTalkErrors/Error") 'use ixmlnodes

  MsgText = "The request has not been accepted."

  ErrMsgType = currnode.selectSingleNode("//Type").Text

  If StrComp(ErrMsgType, "fatal", vbTextCompare) = 0 Then
    bFatal = True
    s = MsgText & vbCrLf & vbCrLf & "Error type: " & ErrMsgType & vbCrLf & "Error number: " & currnode.selectSingleNode("//Number").Text & vbCrLf & currnode.selectSingleNode("//Text").Text
  End If

  If Not bFatal Then
    Set currnode = ResponseDoc.selectSingleNode("//GovTalkMessage/Header/MessageDetails/ResponseEndPoint")
    If Not currnode Is Nothing Then ResponseEndPoint = currnode.Text
    Set currnode = ResponseDoc.selectSingleNode("//GovTalkMessage/Header/MessageDetails/ResponseEndPoint/@PollInterval")
    If Not currnode Is Nothing Then PollInterval = currnode.Text
    Set currnode = ResponseDoc.selectSingleNode("//GovTalkMessage/Header/MessageDetails/CorrelationID") 'FH 30/4/03: Not present in messages from TPVS
    If Not currnode Is Nothing Then CorrID = currnode.Text

    If StrComp(ErrMsgType, "recoverable", vbTextCompare) = 0 Then
      Wait = PollInterval * 1000
      SleepEx (Wait)
      TryAgain = True
      s = ""
    ElseIf StrComp(ErrMsgType, "business", vbTextCompare) = 0 Then
      MsgText = MsgText & vbCrLf & vbCrLf & "Error type: " & ErrMsgType & vbCrLf & "Error number: " & currnode.selectSingleNode("//Number").Text & vbCrLf & currnode.selectSingleNode("//Text").Text & vbCrLf

      'business errors are given within the body of ResponseDoc
      Set ErrList = ResponseDoc.selectNodes("//GovTalkMessage/Body/ErrorResponse/Error")
      For Each ErrNode In ErrList
        ErrRaisedby = ErrNode.selectSingleNode("RaisedBy").Text
        ErrNum = ErrNode.selectSingleNode("Number").Text
        ErrType = ErrNode.selectSingleNode("Type").Text
        Errtext = ErrNode.selectSingleNode("Text").Text
        ErrLocation = ErrNode.selectSingleNode("Location").Text
        ErrDetails = "Type: " & ErrType & vbCrLf & "Number: " & ErrNum & vbCrLf & "Location: " & ErrLocation & vbCrLf & Errtext
        MsgText = MsgText & vbCrLf & ErrDetails & vbCrLf
      Next ErrNode
      's = MsgText & vbCrLf & Comm.DeleteSubmission(CorrID, ResponseEndPoint).xml
      RequiresDelete = True
      s = MsgText
    Else
      s = MsgText & vbCrLf & vbCrLf & "Error type: " & ErrMsgType & vbCrLf & "Error number: " & currnode.selectSingleNode("//Number").Text & vbCrLf & currnode.selectSingleNode("//Text").Text & vbCrLf
    End If
   End If

ErrorResponseMsg_End:
  ErrorResponseMsg = s
  Set currnode = Nothing
  Exit Function

ErrorResponseMsg_Err:
    Err.Raise Err.Number, ErrorSourceComponent(Err, "ErrorResponseMsg", App.EXEName), Err.Description
    Resume
End Function

Public Function AttachAccounts(ByVal Filename As String, ByVal vpdfFile As Variant) As Boolean

On Error GoTo AttachAccounts_Err
  AttachAccounts = False
  mAccName = Right(Filename, Len(Filename) - InStrRev(Filename, "\", , vbTextCompare)) 'remove dirs
  mAccpdf = vpdfFile
  
AttachAccounts_End:
  AttachAccounts = True
  Exit Function
     
AttachAccounts_Err:
    Err.Raise Err.Number, ErrorSourceComponent(Err, "AttachAccounts", App.EXEName), Err.Description
    Resume
End Function

Public Function AttachAccountsFromFilename(ByVal Filename As String) As Boolean
  Dim FileLength As Long, FileNumber As Long
  Dim btArr() As Byte
  
  On Error GoTo AttachAccountsFromFilename_Err
  FileNumber = FreeFile()
  Open Filename For Binary Access Read As #FileNumber
  FileLength = LOF(FileNumber)
  ReDim btArr(0 To FileLength - 1)
  Get #FileNumber, , btArr
  Close #FileNumber
  
  Call AttachAccounts(Filename, btArr)
  
AttachAccountsFromFilename_End:
  AttachAccountsFromFilename = True
  Exit Function
     
AttachAccountsFromFilename_Err:
    Err.Raise Err.Number, ErrorSourceComponent(Err, "AttachAccountsFromFilename", App.EXEName), Err.Description
    Resume
End Function

Public Function AttachComputation(ByVal Filename As String, ByVal vpdfFile As Variant) As Boolean
  
  On Error GoTo AttachComputation_Err
  AttachComputation = False
  mCompName = Right(Filename, Len(Filename) - InStrRev(Filename, "\", , vbTextCompare)) 'remove dirs
  mComppdf = vpdfFile

AttachComputation_End:
  AttachComputation = True
  Exit Function
     
AttachComputation_Err:
    Err.Raise Err.Number, ErrorSourceComponent(Err, "AttachComputation", App.EXEName), Err.Description
    Resume
End Function

Public Function AttachComputationFromFilename(ByVal Filename As String) As Boolean
  Dim FileLength As Long, FileNumber As Long
  Dim btArr() As Byte
  
  On Error GoTo AttachComputationFromFilename_Err
  FileNumber = FreeFile()
  Open Filename For Binary Access Read As #FileNumber
  FileLength = LOF(FileNumber)
  ReDim btArr(0 To FileLength - 1)
  Get #FileNumber, , btArr()
  Close #FileNumber
    
  Call AttachComputation(Filename, btArr)

AttachComputationFromFilename_End:
  AttachComputationFromFilename = True
  Exit Function
     
AttachComputationFromFilename_Err:
    Err.Raise Err.Number, ErrorSourceComponent(Err, "AttachComputationFromFilename", App.EXEName), Err.Description
    Resume
End Function

Private Sub Class_Initialize()
  Set mKeys = New Collection
  ProceedSubmission = True
  TidySubmissions = True
  m_NotifyIntervalMilliseconds = 500
  mTimeOut = 300000 '5 mins
  mPeriodEnd2 = UNDATED
  CheckProxy = True
End Sub
Public Property Get NotifyIntervalMilliseconds() As Long
  NotifyIntervalMilliseconds = m_NotifyIntervalMilliseconds
End Property
Public Property Let NotifyIntervalMilliseconds(ByVal NewValue As Long)
  If (NewValue < 1) Then Call Err.Raise(380, NotifyIntervalMilliseconds, "NotifyIntervalMilliseconds must be greater than 0")
  m_NotifyIntervalMilliseconds = NewValue
End Property
Private Sub SleepEx(Wait As Long)
  If (Wait > 0) Then
    Call Sleep(Wait)
  End If
End Sub
Private Function CardinalToOrdinal(ByVal i As Long) As String
  Dim s As String
  Dim sOrdinal As String
  s = CStr(i)
  
  Select Case True
    Case i Mod 100 = 11, i Mod 100 = 12, i Mod 100 = 13
      sOrdinal = "th"
    Case i Mod 10 = 1
      sOrdinal = "st"
    Case i Mod 10 = 2
      sOrdinal = "nd"
    Case i Mod 10 = 3
      sOrdinal = "rd"
    Case Else
      sOrdinal = "th"
  End Select
  
  CardinalToOrdinal = s & sOrdinal
End Function
Private Function ProcessSubmission(ByVal SubmissionID As String, Optional ByVal iNotify As IBaseNotify, Optional ByRef UserCancelled As Boolean) As String
  Dim status As String, PreviousStatus As String
  Dim Comm As GatewayComm
  Dim ResponseDoc As DOMDocument
  Dim currnode As IXMLDOMNode
  Dim MsgType As String, MsgText As String, Fn As String
  Dim Wait As Long, s As String
  Dim rs As Recordset
  Dim ResponseEndPoint As String, PollInterval As Integer, CorrelationID As String, SubmissionBody As String
  Dim bTryAgain As Boolean, bRequiresDelete As Boolean
  Dim StartTime As Long, CurrentTime As Long
  Dim frm As frmTimeOut
  Dim Retry As Boolean
  Dim iSubmissionAttempt As Long
  Dim iWaitLoops As Long
  Dim i As Long
  
  
  On Error GoTo ProcessSubmission_Err
  
  Call Notify(iNotify, "Processing submission...")
  
  StartTime = GetTicks
  
  'find out what the current status is
  Set rs = DA.SubmissionRS(SubmissionID)
  If rs.BOF And rs.EOF Then Err.Raise ERR_INVALIDID, "ProcessSubmission", "Invalid SubmissionID [" & SubmissionID & "]"  'pq err
  
  status = rs.Fields("status").Value
  CorrelationID = IsNullEx(rs.Fields("CorrelationID"), "")
  ResponseEndPoint = IsNullEx(rs.Fields("ResponseEndPoint"), "")
  PollInterval = IsNullEx(rs.Fields("PollInterval"), 0)
  SubmissionBody = rs.Fields("SubmissionBody").Value
  Call CloseRecordSet(rs)
  
  Set Comm = New GatewayComm
  Comm.TestSubmission = TestSubmission
  Comm.SubmitType = mSubmitType
  
  While (StrComp(status, S_STATUS_COMPLETED, vbTextCompare) <> 0) And (StrComp(status, S_STATUS_ERROR, vbTextCompare) <> 0)
    PreviousStatus = status
        
    If (iSubmissionAttempt = 0) Then
      Call Notify(iNotify, "Contacting the Goverment Gateway")
    Else
      Wait = PollInterval * 1000
      
      If (Wait > 0) Then
        iWaitLoops = Wait / NotifyIntervalMilliseconds
        If (iWaitLoops < 1) Then
          iWaitLoops = 1
        End If
        For i = 1 To iWaitLoops
          Call SleepEx(NotifyIntervalMilliseconds)
          Call Notify(iNotify, "Contacting the gateway for status, " & CardinalToOrdinal(iSubmissionAttempt) & " attempt.")
        Next
      End If
    End If
    
    'fire next step dependent on status
    Select Case status
      Case S_STATUS_SUBMISSION_ACKNOWLEDGEMENT
        If Not iNotify Is Nothing Then Call iNotify.Notify(0, 0, "Polling gateway...")
        Set ResponseDoc = Comm.PollSubmission(CorrelationID, mMockSub, ResponseEndPoint, CheckProxy)
      Case S_STATUS_SUBMISSION_RESPONSE
        If Not iNotify Is Nothing Then Call iNotify.Notify(0, 0, "Cleaning up submission...")
        Set ResponseDoc = Comm.DeleteSubmission(CorrelationID, mMockSub, ResponseEndPoint, CheckProxy)
      Case S_STATUS_DELETE_ACKNOWLEDGEMENT
        If Not iNotify Is Nothing Then Call iNotify.Notify(0, 0, "Cleaning up submission...")
        Set ResponseDoc = Comm.DeleteSubmission(CorrelationID, mMockSub, ResponseEndPoint, CheckProxy)
      Case S_STATUS_DELETE_REQUEST_CLIENT
        If Not iNotify Is Nothing Then Call iNotify.Notify(0, 0, "Cancelling submission...")
        Set ResponseDoc = Comm.DeleteSubmission(CorrelationID, mMockSub, ResponseEndPoint, CheckProxy)
      Case S_STATUS_ERROR_REQUIRES_DELETE
        If Not iNotify Is Nothing Then Call iNotify.Notify(0, 0, "Cancelling submission...")
        Set ResponseDoc = Comm.DeleteSubmission(CorrelationID, mMockSub, ResponseEndPoint, CheckProxy)
      Case S_STATUS_CLIENT_DELETED
        ' Fix #46
        s = "The submission was cancelled by the user and cannot be resumed.  Please make a new submission"
        GoTo ProcessSubmission_End
      Case S_STATUS_NONE
      Case S_STATUS_NOT_SUBMITTED 'new submission
        Set ResponseDoc = Comm.PostSubmission(SubmissionBody, mMockSub, AsyncXMLSubmission, iNotify, NotifyIntervalMilliseconds, CheckProxy)
    End Select
    
    'interpret response...
    Dim dom As DOMDocument
    Set dom = New DOMDocument
    Call dom.loadXML(ResponseDoc.xml)
    With ResponseDoc
      Set currnode = .selectSingleNode("//GovTalkMessage")
      If currnode Is Nothing Then
        App.LogEvent "Invalid response from gateway.  Response xml: " & vbCrLf & vbCrLf & ResponseDoc.xml, vbLogEventTypeError
        Err.Raise ERR_GATEWAY, "ProcessSubmission", "A valid response was not returned from the Gateway, please check that the Gateway is operational.  See event log for details."
      End If
      Set currnode = .selectSingleNode("//GovTalkMessage/Header/MessageDetails/Qualifier")
      MsgType = currnode.Text
      Set currnode = .selectSingleNode("//GovTalkMessage/Header/MessageDetails/Function")
      Fn = currnode.Text
      Set currnode = .selectSingleNode("//GovTalkMessage/Header/MessageDetails/ResponseEndPoint")
      ResponseEndPoint = currnode.Text
      Set currnode = .selectSingleNode("//GovTalkMessage/Header/MessageDetails/ResponseEndPoint/@PollInterval")
      PollInterval = currnode.Text
      Set currnode = .selectSingleNode("//GovTalkMessage/Header/MessageDetails/CorrelationID")
      CorrelationID = currnode.Text
    End With
      
    '...and update status
    Select Case LCase(MsgType)
      Case "error"
        'set global domdoc in case of error
        Set ErrResponseDoc = New DOMDocument
        Set ErrResponseDoc = ResponseDoc
        s = ErrorResponseMsg(ResponseDoc, False, False)
        If Not bTryAgain Then 'if TryAgain then do nothing, the while loop should make same call that caused this error
          If bRequiresDelete Then
            status = S_STATUS_ERROR_REQUIRES_DELETE
          Else
            status = S_STATUS_ERROR
          End If
          If StrComp(PreviousStatus, S_STATUS_ERROR_REQUIRES_DELETE, vbTextCompare) = 0 Then
            Call DA.UpdateStatus(SubmissionID, status, ResponseEndPoint, PollInterval, s, CorrelationID, True)
          Else
            Call DA.UpdateStatus(SubmissionID, status, ResponseEndPoint, PollInterval, s, CorrelationID)
          End If
        End If
      Case "response"
        Select Case LCase(Fn)
          Case "submit"
            status = S_STATUS_SUBMISSION_RESPONSE
            s = "Submission accepted, no further action required"
            Call DA.UpdateStatus(SubmissionID, status, ResponseEndPoint, PollInterval, s, CorrelationID)
            Call DA.ArchiveUpdate(mArcLevel, SubmissionID, mAccName, mCompName)
          Case "delete"
            status = S_STATUS_COMPLETED
            Call DA.UpdateStatus(SubmissionID, status, ResponseEndPoint, PollInterval, , CorrelationID)
        End Select
      Case "acknowledgement"
        Select Case LCase(Fn)
          Case "submit"
            s = "Submission has been acknowledged"
            status = S_STATUS_SUBMISSION_ACKNOWLEDGEMENT
            Call DA.UpdateStatus(SubmissionID, status, ResponseEndPoint, PollInterval, s, CorrelationID)
          Case "delete"
            status = S_STATUS_DELETE_ACKNOWLEDGEMENT
            Call DA.UpdateStatus(SubmissionID, status, ResponseEndPoint, PollInterval, , CorrelationID)
        End Select
      Case Else
        Err.Raise ERR_GATEWAY, "ProcessSubmission", "Unknown response type [" & MsgType & "] from gateway"
    End Select
    
    'reset timeout if status has changed
    If StrComp(PreviousStatus, status, vbTextCompare) <> 0 Then StartTime = GetTicks
    
    'Check to see if timeout limit has been reached
    CurrentTime = GetTicks
    If mTimeOut > -1 And mTimeOut < CurrentTime - StartTime Then
      ' open frmtimeOut here and give users the choice to continue or cancel submission
      Set frm = New frmTimeOut 'create a new instance of form
      frm.lblMessage.Caption = "The submission has reached the pre-defined time out period. Please press retry if you would like to continue polling for the next " & mTimeOut / 1000 & " seconds"
      frm.Show vbModal
      Retry = frm.Retry
      If Retry Then
        StartTime = GetTicks ' restart the start time
      Else
        s = "You have chosen to cancel this submission. Please note you will have to resume your submission at a later date."
        Call DA.UpdateStatus(SubmissionID, status, ResponseEndPoint, PollInterval, s, CorrelationID)
        UserCancelled = True
        GoTo ProcessSubmission_End
      End If
    End If
    iSubmissionAttempt = iSubmissionAttempt + 1
  Wend
    
ProcessSubmission_End:
  Set rs = Nothing
  Set Comm = Nothing
  Set frm = Nothing
  Set ResponseDoc = Nothing
  ProcessSubmission = s
  Exit Function
  
ProcessSubmission_Err:
  Err.Raise Err.Number, ErrorSourceComponent(Err, "ProcessSubmission", App.EXEName), Err.Description
  Resume
End Function



Public Function CheckSubmissions(Optional ByVal LastID As String, Optional ByVal Filter As String, Optional ByVal iNotify As IBaseNotify) As String
  Dim rs As Recordset, SubmissionID As String
        
  On Error GoTo CheckSubmissions_Err
  'send notification to calling application
  If Not iNotify Is Nothing Then Call iNotify.Notify(0, 0, "Checking submissions...")
  
  'check there is a network connection
  Call CheckNetwork 'FIX #865
  
  'get recordset of all submissions
  Set rs = DA.StalledSubmissions 'creates a read-only recordset by default
  If Not rs.BOF And Not rs.EOF Then
    Filter = "SubmissionID >" & gDBHelper.StrSQL(LastID)   '    find rs with > last guid
    rs.Filter = Filter
    If Not rs.EOF Then
      '    process next record (single)
      SubmissionID = rs.Fields("SubmissionID") ' return guid
    Else  'either lastID is the last record or no lastID entered, therefore take first record in.
      rs.Filter = ""
      SubmissionID = rs.Fields("SubmissionID")
    End If
    Call CloseRecordSet(rs)
    Call ProcessSubmission(SubmissionID)
    CheckSubmissions = SubmissionID
  End If
CheckSubmissions_End:
  Set rs = Nothing
  Exit Function
  
CheckSubmissions_Err:
  Err.Raise Err.Number, ErrorSourceComponent(Err, "CheckSubmissions", App.EXEName), Err.Description
  Resume
End Function

Public Function CancelSubmission(ByVal SubmissionID As String, Optional ByVal iNotify As IBaseNotify) As String
  Dim rs As Recordset, s As String, sCurrentStatus As String
      
  On Error GoTo CancelSubmission_Err
  If Not iNotify Is Nothing Then Call iNotify.Notify(0, 0, "Cancelling submission...")
  
  'find out what the current status is
  Set rs = DA.SubmissionRS(SubmissionID)
  If rs.BOF And rs.EOF Then Err.Raise ERR_INVALIDID, "ProcessSubmission", "Invalid SubmissionID [" & SubmissionID & "]"
  ' FIX #46
  sCurrentStatus = rs.Fields("Status")
  Select Case sCurrentStatus
    Case S_STATUS_NOT_SUBMITTED, S_STATUS_ERROR
      s = "The submission has been cancelled by the user."
    Case Else
      s = "The submission has been cancelled by the user but a copy has already been sent to the Revenue."
  End Select
  
  If DeleteRequired(rs) Then
    'check there is a network connection
    Call CheckNetwork 'FIX #865
    
    Call DA.UpdateStatus(SubmissionID, S_STATUS_DELETE_REQUEST_CLIENT)
    Call ProcessSubmission(SubmissionID, iNotify)
  End If
  Call CloseRecordSet(rs)
  
  Call DA.UpdateStatus(SubmissionID, S_STATUS_CLIENT_DELETED, , , s)

CancelSubmission_End:
  Set rs = Nothing
  CancelSubmission = s
  Exit Function
  
CancelSubmission_Err:
  Err.Raise Err.Number, ErrorSourceComponent(Err, "CancelSubmission", App.EXEName), Err.Description
  Resume CancelSubmission_End
  Resume
End Function

Private Function DeleteRequired(ByVal rs As Recordset) As Boolean
  Dim status As String
  
On Error GoTo DeleteRequired_Err
  
  status = rs.Fields("status").Value
  
  Select Case status
    Case S_STATUS_SUBMISSION_ACKNOWLEDGEMENT, S_STATUS_ERROR_REQUIRES_DELETE, S_STATUS_DELETE_ACKNOWLEDGEMENT
      DeleteRequired = True
    Case S_STATUS_SUBMISSION_RESPONSE, S_STATUS_NONE, S_STATUS_NOT_SUBMITTED, S_STATUS_DELETE_RESPONSE, S_STATUS_ERROR, S_STATUS_COMPLETED
      DeleteRequired = False
  End Select
  
DeleteRequired_End:
  Exit Function
  
DeleteRequired_Err:
  Err.Raise Err.Number, ErrorSourceComponent(Err, "DeleteRequired", App.EXEName), Err.Description
  Resume DeleteRequired_End
End Function


Public Function SubmissionExists(ByVal SubmissionID As String) As Boolean
  Dim rs As Recordset
  
  Set rs = DA.SubmissionRS(SubmissionID)
  SubmissionExists = Not (rs.EOF And rs.BOF)
End Function

Public Function CheckStatus(ByVal SubmissionID As String, Optional ByVal iNotify As IBaseNotify) As String
  Dim status As String, s As String
  Dim rs As Recordset
      
  On Error GoTo CheckStatus_Err
  If Not iNotify Is Nothing Then Call iNotify.Notify(0, 0, "Checking status...")
  
  'find out what the current status is
  Set rs = DA.SubmissionRS(SubmissionID)
  s = "Status: "
  If (rs.EOF And rs.BOF) Then
    s = s & "Submission '" & SubmissionID & "' does not exist"
  Else
    s = s & rs.Fields("Status") & vbCrLf & "Message: " & rs.Fields("Message") & vbCrLf & "Last Updated: " & rs.Fields("LastUpdated")
  End If
  Call CloseRecordSet(rs)
  
CheckStatus_End:
  CheckStatus = s
  Exit Function
  
CheckStatus_Err:
  Err.Raise Err.Number, ErrorSourceComponent(Err, "CheckStatus", App.EXEName), Err.Description
  Resume CheckStatus_End
End Function

Public Function CheckAllStatus(Optional ByVal iNotify As IBaseNotify) As String
  Dim s As String, MsgText As String
  Dim rs As Recordset
  Dim rsFilter As String
  
  On Error GoTo CheckAllStatus_Err
  
  If Not iNotify Is Nothing Then Call iNotify.Notify(0, 0, "Checking status for all submissions...")
  
  'find out what the current status is for all submissions in the table
  rsFilter = ""
  Set rs = DA.GetFilteredRecordSet(rsFilter) ' filter should include the 'Where' sql clause

  Do While Not rs.EOF
    If Not Len(MsgText) > 0 Then
      MsgText = "Submission ID: " & rs.Fields("SubmissionID") & vbCrLf & "Status: " & rs.Fields("Status") & vbCrLf & "Message: " & rs.Fields("Message") & vbCrLf & "Last Updated: " & rs.Fields("LastUpdated")
    Else
      MsgText = MsgText & vbCrLf & vbCrLf & "Submission ID: " & rs.Fields("SubmissionID") & vbCrLf & "Status: " & rs.Fields("Status") & vbCrLf & "Message: " & rs.Fields("Message") & vbCrLf & "Last Updated: " & rs.Fields("LastUpdated")
    End If
    rs.MoveNext
  Loop
  s = MsgText
  Call CloseRecordSet(rs)
  
CheckAllStatus_End:
  CheckAllStatus = s
  Set rs = Nothing
  Exit Function
  
CheckAllStatus_Err:
  Err.Raise Err.Number, ErrorSourceComponent(Err, "CheckAllStatus", App.EXEName), Err.Description
  Resume CheckAllStatus_End
End Function


Public Function ResumeSubmission(ByVal SubmissionID As String, Optional ByVal iNotify As IBaseNotify) As String
  Dim s As String
  
  On Error GoTo ResumeSubmission_Err
  If Not iNotify Is Nothing Then Call iNotify.Notify(0, 0, "Resuming submission...")

  'check there is a network connection
  Call CheckNetwork 'FIX #865
  
  ' resume submission for particular submissionID
  s = ProcessSubmission(SubmissionID)

ResumeSubmission_End:
  ResumeSubmission = s
  Exit Function

ResumeSubmission_Err:
  Err.Raise Err.Number, ErrorSourceComponent(Err, "ResumeSubmission", App.EXEName), Err.Description
  Resume ResumeSubmission_End
End Function
 
' New function using ABS file to do work - For future release
Private Function IsAvailableAbacus_Future(ByVal obj As Object) As Boolean
  'Dim ab As Abacus.Application
  'Dim errmsg As String
  'On Error GoTo IsAvailableAbacus_Future_Err
  
  'Set ab = obj.Application
  'Call ab.SetVariable("efiler_enabled", False)
  'Call ab.SetVariable("efiler_error_msg", "")
  'Call ab.Execute("SCHEDULERETRIEVE(""EFILER.ABS"",""check_enabled"")")
  'IsAvailableAbacus_Future = ab.Evaluate("+efiler_enabled")
  'errmsg = ab.Evaluate("+efiler_error_msg")
  'If Not IsAvailableAbacus_Future Then Err.Raise ERR_EFILERCHECK, "IsAvailableAbacus_Future", errmsg
  ' MS - remove before releasing
  'If True Then
  '  If IsAvailableAbacus_Future <> IsAvailableAbacus(obj) Then Call ECASE("Unexpected mismatch between Abacus functions")
  'End If
  'Exit Function
  
'IsAvailableAbacus_Future_Err:
  'Call ErrorMessage(ERR_ERROR, Err, "IsAvailableAbacus_Future", "Unable to eFile", "")
  'Err.Raise Err.Number, ErrorSourceComponent(Err, "IsAvailableAbacus_Future", App.EXEName), Err.Description
  'Resume
End Function

Public Function IsAvailableAbacus(ByVal obj As Object) As Boolean
  Dim NewAb As Boolean
  On Error GoTo IsAvailableAbacus_Err
  'If Not iNotify Is Nothing Then Call iNotify.Notify(0, 0, "Resuming submission...")
  
  'For new forms (2005), we need to use a different validation function
  If obj.Application.Evaluate("@NAMEEXISTS(&:[@]:SYS$VARS:,""CT600_05P1_SET"")") Then
    If obj.Application.Evaluate("@EMPTY(&:[@]:SYS$VARS:CT600_05P1_SET)") = 0 Then
      IsAvailableAbacus = IsAvailableAbacus_CT2005(obj)
      Exit Function
    End If
  End If

  'For new forms (2004), we need to use a different validation function
  If obj.Application.Evaluate("@NAMEEXISTS(&:[@]:SYS$VARS:,""CT600_04P1_SET"")") Then
    If obj.Application.Evaluate("@EMPTY(&:[@]:SYS$VARS:CT600_04P1_SET)") = 0 Then
      IsAvailableAbacus = IsAvailableAbacus_CT2004(obj)
      Exit Function
    End If
  End If
  
  NewAb = False
  
  Select Case UCase(obj.Pack)
    Case "CPS02C.ABS": IsAvailableAbacus = True
    Case "CPS03C.ABS": IsAvailableAbacus = True
    Case "CPS03B.ABS": IsAvailableAbacus = True
    ' #1094 mda added 01/03/04
    Case "CPS04C.ABS":
      IsAvailableAbacus = True
      NewAb = True
    Case "CPS04I.ABS":
      IsAvailableAbacus = True
      NewAb = True
    ' end #1094
    Case "CPS04B.ABS":
      IsAvailableAbacus = True
      NewAb = True
    Case "CPS05C.ABS", "CPS05I.ABS", "CPS05B.ABS":
      IsAvailableAbacus = True
    Case Else
      MsgBox "e-Filing this type of computation for this year end is not supported.  Please file your return in hard-copy.", vbCritical, "Online Services"
      IsAvailableAbacus = False
      Exit Function
  End Select
  
  
  If Not obj.Application.Evaluate("@EMPTY(&:[@]:SYS$VARS:CT600_02P1_SET)=1") Then
    If obj.Application.Evaluate("SAR_SUPPA") Then
      If Not NewAb Then
        MsgBox "e-Filing a CT600 with Supplementary Form A is not currently supported.  Please file your return in hard-copy.  An updated edition of e-Filing will be released presently that supports this functionality.", vbCritical, "Online Services"
        IsAvailableAbacus = False
        Exit Function
      Else
        If Not obj.Application.Evaluate("(@sheetexists(:[@]:," + Chr(34) + "CT600A_00P1" + Chr(34) + ")>0)") Or Not obj.Application.Evaluate("@getxmlname(&:[@]:CT600A_00P1:A2_1PART)=" + Chr(34) + "CT600A_Loan_Name1" + Chr(34)) Then
          MsgBox "To e-File you must update the CT600A form in this computation.  Please erase the CT600A (using Schedule, Erase from the menu options) and re-develop the CT600A.  You can then proceed.", vbCritical, "Online Services"
          IsAvailableAbacus = False
          Exit Function
        End If
        
        If obj.Application.Evaluate("@val(+:[@]:BOX_A1)>0") Then
          If obj.Application.Evaluate("@strlen(+:[@]:SYS$VARS:A50_SET)>0") Then
              If obj.Application.Evaluate("@VAL(+:[@]:[A50]:XMLVERSION)>0") Then
                  IsAvailableAbacus = True
              Else
                  MsgBox "e-Filing a CT600A, you will need to use a more recent version of the `Loans to participators` schedule.  Please copy your data to a working paper, erase the existing sheet, and re-develop it.  Please retry your submission when you have moved your analysis into the new schedule.", vbCritical, "Online Services"
                  IsAvailableAbacus = False
                  Exit Function
              End If
          Else
              MsgBox "e-Filing a CT600A, you are required to use the supporting Abacus Schedule `Loans to Participators`.  Please develop this schedule and enter your information.  Please retry your submission when you have analysed out your claim.", vbCritical, "Online Services"
              IsAvailableAbacus = False
              Exit Function
          End If
        End If
      End If
        
    End If
    If obj.Application.Evaluate("SAR_SUPPB") Then
      If Not NewAb Then
        MsgBox "e-Filing a CT600 with Supplementary Form B is not currently supported.  Please file your return in hard-copy.  An updated edition of e-Filing will be released presently that supports this functionality.", vbCritical, "Online Services"
        IsAvailableAbacus = False
        Exit Function
      Else
        If Not obj.Application.Evaluate("(@sheetexists(:[@]:," + Chr(34) + "CT600B_99P1" + Chr(34) + ")>0)") Or Not obj.Application.Evaluate("@getxmlname(&:[@]:CT600B_99P2:CFC_1N)=" + Chr(34) + "CT600B_Name1" + Chr(34)) Then
          MsgBox "To e-File you must update the CT600B form in this computation.  Please erase the CT600B (using Schedule, Erase from the menu options) and re-develop the CT600B.  You can then proceed.", vbCritical, "Online Services"
          IsAvailableAbacus = False
          Exit Function
        End If
        
        If obj.Application.Evaluate("@strlen(+:[@]:SYS$VARS:A39_SET)>0") Then
            If obj.Application.Evaluate("@VAL(+:[@]:[A39]:XMLVERSION)>0") Then
                IsAvailableAbacus = True
            Else
                MsgBox "e-Filing a CT600B, you will need to use a more recent version of the `Controlled foreign companies summary` schedule.  Please copy your data to a working paper, erase the existing sheet, and re-develop it.  Please retry your submission when you have moved your analysis into the new schedule.", vbCritical, "Online Services"
                IsAvailableAbacus = False
                Exit Function
            End If
        Else
            MsgBox "e-Filing a CT600B, you are required to use the supporting Abacus Schedule `Controlled foreign companies summary`.  Please develop this schedule and enter your information.  Please retry your submission when you have analysed out your claim.", vbCritical, "Online Services"
            IsAvailableAbacus = False
            Exit Function
        End If
      End If
    End If
    If obj.Application.Evaluate("SAR_SUPPC") Then
      
      If Not obj.Application.Evaluate("(@sheetexists(:[@]:," + Chr(34) + "CT600C_02P1" + Chr(34) + ")>0)#OR#(@sheetexists(:[@]:," + Chr(34) + "CT600C_01P1" + Chr(34) + ")>0)") Then
        MsgBox "To e-File you must update the CT600C form in this computation.  Please erase the CT600C (using Schedule, Erase from the menu options) and re-develop the CT600C.  You can then proceed.", vbCritical, "Online Services"
        IsAvailableAbacus = False
        Exit Function
      End If
      
      If obj.Application.Evaluate("@val(+:[@]:BOX_C1)>0") Then
        If obj.Application.Evaluate("@strlen(+:[@]:SYS$VARS:P6_SET)>0") Then
            If obj.Application.Evaluate("@VAL(+:[@]:[P6]:XMLVERSION)>0") Then
                IsAvailableAbacus = True
            Else
                MsgBox "e-Filing a CT600C with a claim to group relief, you will need to use a more recent version of the `Amounts claimed as group relief`.  Please copy your data to a working paper, erase the existing sheet, and re-develop it.  Please retry your submission when you have moved your analysis into the new schedule.", vbCritical, "Online Services"
                IsAvailableAbacus = False
                Exit Function
            End If
        Else
            MsgBox "e-Filing a CT600C with a claim to group relief, you are required to use the supporting Abacus Schedule `Amounts claimed as group relief`.  Please develop this schedule (from `P`) and enter your information.  Please retry your submission when you have analysed out your claim.", vbCritical, "Online Services"
            IsAvailableAbacus = False
            Exit Function
        End If
      End If
      
      If obj.Application.Evaluate("@val(+:[@]:BOX_C2)>0") Then
        If obj.Application.Evaluate("@strlen(+:[@]:SYS$VARS:P7_SET)>0") Then
            If obj.Application.Evaluate("@VAL(+:[@]:[P7]:XMLVERSION)>0") Then
                IsAvailableAbacus = True
            Else
                MsgBox "e-Filing a CT600C with a surrender of group relief, you will need to use a more recent version of the `Amounts surrendered as group relief`.  Please copy your data to a working paper, erase the existing sheet, and re-develop it.  Please retry your submission when you have moved your analysis into the new schedule.", vbCritical, "Online Services"
                IsAvailableAbacus = False
                Exit Function
            End If
        Else
            MsgBox "e-Filing a CT600C with a surrender of group relief, you are required to use the supporting Abacus Schedule `Amounts surrendered as group relief`.  Please develop this schedule (from `P`) and enter your information.  Please retry your submission when you have analysed out your claim.", vbCritical, "Online Services"
            IsAvailableAbacus = False
            Exit Function
        End If
      End If
      
      If obj.Application.Evaluate("@val(+:[@]:BOX_C3)>0") Then
        If obj.Application.Evaluate("@strlen(+:[@]:SYS$VARS:P8_SET)>0") Then
            If obj.Application.Evaluate("@VAL(+:[@]:[P8]:XMLVERSION)>0") Then
                IsAvailableAbacus = True
            Else
                MsgBox "e-Filing a CT600C with a claim to EUFT, you will need to use a more recent version of the `Amounts claimed as EUFT`.  Please copy your data to a working paper, erase the existing sheet, and re-develop it.  Please retry your submission when you have moved your analysis into the new schedule.", vbCritical, "Online Services"
                IsAvailableAbacus = False
                Exit Function
            End If
        Else
            MsgBox "e-Filing a CT600C with a claim to EUFT, you are required to use the supporting Abacus Schedule `Amounts claimed as EUFT`.  Please develop this schedule (from `P`) and enter your information.  Please retry your submission when you have analysed out your claim.", vbCritical, "Online Services"
            IsAvailableAbacus = False
            Exit Function
        End If
      End If
      If obj.Application.Evaluate("@val(+:[@]:BOX_C4)>0") Then
        If obj.Application.Evaluate("@strlen(+:[@]:SYS$VARS:P9_SET)>0") Then
            If obj.Application.Evaluate("@VAL(+:[@]:[P9]:XMLVERSION)>0") Then
                IsAvailableAbacus = True
            Else
                MsgBox "e-Filing a CT600C with a surrender of EUFT, you will need to use a more recent version of the `Amounts surrendered as group EUFT`.  Please copy your data to a working paper, erase the existing sheet, and re-develop it.  Please retry your submission when you have moved your analysis into the new schedule.", vbCritical, "Online Services"
                IsAvailableAbacus = False
                Exit Function
            End If
        Else
            MsgBox "e-Filing a CT600C with a surrender of EUFT, you are required to use the supporting Abacus Schedule `Amounts surrendered as group EUFT`.  Please develop this schedule (from `P`) and enter your information.  Please retry your submission when you have analysed out your claim.", vbCritical, "Online Services"
            IsAvailableAbacus = False
            Exit Function
        End If
      End If
    End If
    If obj.Application.Evaluate("SAR_SUPPD") Then
      If Not NewAb Then
        MsgBox "e-Filing a CT600 with Supplementary Form D is not currently supported.  Please file your return in hard-copy.  An updated edition of e-Filing will be released presently that supports this functionality.", vbCritical, "Online Services"
        IsAvailableAbacus = False
        Exit Function
      Else
        If Not obj.Application.Evaluate("(@sheetexists(:[@]:," + Chr(34) + "CT600D_99P1" + Chr(34) + ")>0)") Or Not obj.Application.Evaluate("@getxmlname(&:[@]:CT600D_99P1:SRD_OLAB)=" + Chr(34) + "CT600D_Declaration" + Chr(34)) Then
          MsgBox "To e-File you must update the CT600D form in this computation.  Please erase the CT600D (using Schedule, Erase from the menu options) and re-develop the CT600D.  You can then proceed.", vbCritical, "Online Services"
          IsAvailableAbacus = False
          Exit Function
        End If
      End If
    End If
    If obj.Application.Evaluate("SAR_SUPPE") Then
      If Not NewAb Then
        MsgBox "e-Filing a CT600 with Supplementary Form E is not currently supported.  Please file your return in hard-copy.  An updated edition of e-Filing will be released presently that supports this functionality.", vbCritical, "Online Services"
        IsAvailableAbacus = False
        Exit Function
      Else
        If Not obj.Application.Evaluate("(@sheetexists(:[@]:," + Chr(34) + "CT600E_02P1" + Chr(34) + ")>0)") Or Not obj.Application.Evaluate("@getxmlname(&:[@]:CT600E_02P1:SRE_CHREF)=" + Chr(34) + "CT600E_ClaimExemption_RepaymentReference" + Chr(34)) Then
          MsgBox "To e-File you must update the CT600E form in this computation.  Please erase the CT600E (using Schedule, Erase from the menu options) and re-develop the CT600E.  You can then proceed.", vbCritical, "Online Services"
          IsAvailableAbacus = False
          Exit Function
        End If
      End If
    End If
    If obj.Application.Evaluate("SAR_SUPPF") Then
      If Not NewAb Then
        MsgBox "e-Filing a CT600 with Supplementary Form F is not currently supported.  Please file your return in hard-copy.  An updated edition of e-Filing will be released presently that supports this functionality.", vbCritical, "Online Services"
        IsAvailableAbacus = False
        Exit Function
      Else
        If Not obj.Application.Evaluate("(@sheetexists(:[@]:," + Chr(34) + "CT600F_00P1" + Chr(34) + ")>0)") Or Not obj.Application.Evaluate("@getxmlname(&:[@]:CT600F_00P2:SRF_F1_YES)=" + Chr(34) + "CT600F_GroupElectionY" + Chr(34)) Then
          MsgBox "To e-File you must update the CT600F form in this computation.  Please erase the CT600F (using Schedule, Erase from the menu options) and re-develop the CT600F.  You can then proceed.", vbCritical, "Online Services"
          IsAvailableAbacus = False
          Exit Function
        End If
        
        If obj.Application.Evaluate("@strlen(+:[@]:SYS$VARS:A42_SET)>0") Then
            If obj.Application.Evaluate("@VAL(+:[@]:[A42]:XMLVERSION)>0") Then
                IsAvailableAbacus = True
            Else
                MsgBox "e-Filing a CT600F, you will need to use a more recent version of the 'Tonnage Tax` schedule.  Please copy your data to a working paper, erase the existing sheet, and re-develop it.  Please retry your submission when you have moved your analysis into the new schedule.", vbCritical, "Online Services"
                IsAvailableAbacus = False
                Exit Function
            End If
        Else
            MsgBox "e-Filing a CT600F, you are required to use the supporting Abacus Schedule `Tonnage Tax`.  Please develop this schedule (from `P`) and enter your information.  Please retry your submission when you have analysed out your claim.", vbCritical, "Online Services"
            IsAvailableAbacus = False
            Exit Function
        End If
      End If
    End If
    If obj.Application.Evaluate("SAR_SUPPG") Then
      If Not NewAb Then
        MsgBox "e-Filing a CT600 with Supplementary Form G is not currently supported.  Please file your return in hard-copy.  An updated edition of e-Filing will be released presently that supports this functionality.", vbCritical, "Online Services"
        IsAvailableAbacus = False
        Exit Function
      Else
        If Not obj.Application.Evaluate("(@sheetexists(:[@]:," + Chr(34) + "CT600G_02P1" + Chr(34) + ")>0)") Or Not obj.Application.Evaluate("@getxmlname(&:[@]:CT600G_02P2:SRG_INVEST_1NAME)=" + Chr(34) + "CT600G_CoName1" + Chr(34)) Then
          MsgBox "To e-File you must update the CT600G form in this computation.  Please erase the CT600G (using Schedule, Erase from the menu options) and re-develop the CT600G.  You can then proceed.", vbCritical, "Online Services"
          IsAvailableAbacus = False
          Exit Function
        End If
        
        If obj.Application.Evaluate("@strlen(+:[@]:SYS$VARS:A43_SET)>0") Then
            If obj.Application.Evaluate("@VAL(+:[@]:[A43]:XMLVERSION)>0") Then
                IsAvailableAbacus = True
            Else
                MsgBox "e-Filing a CT600G, you will need to use a more recent version of the 'Corporate venturing relief` schedule.  Please copy your data to a working paper, erase the existing sheet, and re-develop it.  Please retry your submission when you have moved your analysis into the new schedule.", vbCritical, "Online Services"
                IsAvailableAbacus = False
                Exit Function
            End If
        Else
            MsgBox "e-Filing a CT600G, you are required to use the supporting Abacus Schedule `Corporate venturing relief`.  Please develop this schedule (from `P`) and enter your information.  Please retry your submission when you have analysed out your claim.", vbCritical, "Online Services"
            IsAvailableAbacus = False
            Exit Function
        End If
        
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P2:SRG_INVEST_1NAME)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P2:SRG_INVEST_2NAME)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P2:SRG_INVEST_3NAME)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P2:SRG_INVEST_4NAME)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P2:SRG_INVEST_5NAME)>0") Then IsAvailableAbacus = False
        
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P2:SRG_INVEST_1NUM)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P2:SRG_INVEST_2NUM)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P2:SRG_INVEST_3NUM)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P2:SRG_INVEST_4NUM)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P2:SRG_INVEST_5NUM)>0") Then IsAvailableAbacus = False
        
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P2:SRG_SHARES_1NAME)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P2:SRG_SHARES_2NAME)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P2:SRG_SHARES_3NAME)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P2:SRG_SHARES_4NAME)>0") Then IsAvailableAbacus = False
        
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P2:SRG_SHARES_1DESC)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P2:SRG_SHARES_2DESC)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P2:SRG_SHARES_3DESC)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P2:SRG_SHARES_4DESC)>0") Then IsAvailableAbacus = False
        
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P2:SRG_DEF_1ISSNAME)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P2:SRG_DEF_2ISSNAME)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P2:SRG_DEF_3ISSNAME)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P2:SRG_DEF_4ISSNAME)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P2:SRG_DEF_5ISSNAME)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P2:SRG_DEF_6ISSNAME)>0") Then IsAvailableAbacus = False
        
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P2:SRG_DEF_1POSTNAME)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P2:SRG_DEF_2POSTNAME)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P2:SRG_DEF_3POSTNAME)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P2:SRG_DEF_4POSTNAME)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P2:SRG_DEF_5POSTNAME)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P2:SRG_DEF_6POSTNAME)>0") Then IsAvailableAbacus = False
        
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P2:SRG_DEF_1DISPDATE)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P2:SRG_DEF_2DISPDATE)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P2:SRG_DEF_3DISPDATE)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P2:SRG_DEF_4DISPDATE)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P2:SRG_DEF_5DISPDATE)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P2:SRG_DEF_6DISPDATE)>0") Then IsAvailableAbacus = False
        
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_INVEST_1PAR)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_INVEST_2PAR)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_INVEST_3PAR)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_INVEST_4PAR)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_INVEST_5PAR)>0") Then IsAvailableAbacus = False
        
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_INVEST_1DATE)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_INVEST_2DATE)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_INVEST_3DATE)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_INVEST_4DATE)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_INVEST_5DATE)>0") Then IsAvailableAbacus = False
        
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_INVEST_1SUBS)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_INVEST_2SUBS)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_INVEST_3SUBS)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_INVEST_4SUBS)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_INVEST_5SUBS)>0") Then IsAvailableAbacus = False
        
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_INVEST_1REL)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_INVEST_2REL)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_INVEST_3REL)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_INVEST_4REL)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_INVEST_5REL)>0") Then IsAvailableAbacus = False
        
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_INVEST_1VENTREF)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_INVEST_2VENTREF)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_INVEST_3VENTREF)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_INVEST_4VENTREF)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_INVEST_5VENTREF)>0") Then IsAvailableAbacus = False
        
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_SHARES_1NUM)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_SHARES_2NUM)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_SHARES_3NUM)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_SHARES_4NUM)>0") Then IsAvailableAbacus = False

        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_SHARES_1ISSDATE)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_SHARES_2ISSDATE)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_SHARES_3ISSDATE)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_SHARES_4ISSDATE)>0") Then IsAvailableAbacus = False
      
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_SHARES_1DISPDATE)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_SHARES_2DISPDATE)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_SHARES_3DISPDATE)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_SHARES_4DISPDATE)>0") Then IsAvailableAbacus = False
        
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_SHARES_1APSTART)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_SHARES_2APSTART)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_SHARES_3APSTART)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_SHARES_4APSTART)>0") Then IsAvailableAbacus = False
      
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_SHARES_1APEND)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_SHARES_2APEND)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_SHARES_3APEND)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_SHARES_4APEND)>0") Then IsAvailableAbacus = False
        
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_SHARES_1ALLOW)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_SHARES_2ALLOW)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_SHARES_3ALLOW)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_SHARES_4ALLOW)>0") Then IsAvailableAbacus = False
      
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_DEF_1AMOUNT)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_DEF_2AMOUNT)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_DEF_3AMOUNT)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_DEF_4AMOUNT)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_DEF_5AMOUNT)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_DEF_6AMOUNT)>0") Then IsAvailableAbacus = False
        
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_DEF_1NAME)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_DEF_2NAME)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_DEF_3NAME)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_DEF_4NAME)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_DEF_5NAME)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_DEF_6NAME)>0") Then IsAvailableAbacus = False
        
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_DEF_1ISSDATE)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_DEF_2ISSDATE)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_DEF_3ISSDATE)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_DEF_4ISSDATE)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_DEF_5ISSDATE)>0") Then IsAvailableAbacus = False
        If obj.Application.Evaluate("@strlen(+:[@]:CT600G_02P3:SRG_DEF_6ISSDATE)>0") Then IsAvailableAbacus = False
      
        If IsAvailableAbacus = False Then
            MsgBox "e-Filing a CT600G, you are required to use the supporting Abacus Schedule `Corporate venturing relief`.  Please remove all entries from the front of the CT600G Return by erasing the existing sheet, and re-develop it.  Please retry your submission when you have analysed out your claim.", vbCritical, "Online Services"
            Exit Function
        End If
      End If
    End If
    If obj.Application.Evaluate("SAR_SUPPH") Then
      MsgBox "e-Filing a CT600 with Supplementary Form H is not currently supported.  Please file your return in hard-copy.  An updated edition of e-Filing will be released presently that supports this functionality.", vbCritical, "Online Services"
      IsAvailableAbacus = False
      Exit Function
    End If
    If obj.Application.Evaluate("SAR_SUPPRF") Then
      MsgBox "e-Filing a CT600 with Supplementary Form I is not currently supported.  Please file your return in hard-copy.  An updated edition of e-Filing will be released presently that supports this functionality.", vbCritical, "Online Services"
      IsAvailableAbacus = False
      Exit Function
    End If
        
    If obj.Application.Evaluate("@sheetexists(:[@]:," + Chr(34) + "CT600_02P1" + Chr(34) + ")>0") Then
        If Not obj.Application.Evaluate("@strlen(+:[@]:[CT600_02P1]:XMLVERSION)>0") Then
            MsgBox "To e-File you must update the CT600 form in this computation.  Please erase the CT600 (using Schedule, Erase from the menu options) and re-develop the CT600.  You can then proceed.", vbCritical, "Online Services"
            IsAvailableAbacus = False
            Exit Function
        End If
    End If

    IsAvailableAbacus = True
  Else
    IsAvailableAbacus = False
    MsgBox "To e-file you must generate the CT600 for in the computation.  Please develop it from the 'P' Schedule.  You can then Proceed.", vbCritical, "Online Services"
    Exit Function
  End If
  
IsAvailableAbacus_End:
  Exit Function

IsAvailableAbacus_Err:
  IsAvailableAbacus = False
  Err.Raise Err.Number, ErrorSourceComponent(Err, "IsAvailableAbacus _obsolete", App.EXEName), Err.Description
  Resume IsAvailableAbacus_End
End Function

Public Function IsAvailableAbacus_CT2004(ByVal obj As Object) As Boolean
  Dim NewAb As Boolean
  On Error GoTo IsAvailableAbacus_CT2004_Err
  'If Not iNotify Is Nothing Then Call iNotify.Notify(0, 0, "Resuming submission...")
  
  
  Select Case UCase(obj.Pack)
    Case "CPS05C.ABS":
    Case "CPS05I.ABS":
    Case "CPS05B.ABS":
    ' #1094 mda added 01/03/04
    Case "CPS04C.ABS":
    Case "CPS04I.ABS":
    Case "CPS04B.ABS":
    Case Else
        MsgBox "e-Filing this type of computation for this year end is not supported.  Please file your return in hard-copy.", vbCritical, "Online Services"
        Exit Function
  End Select
  
  'Main form Test
  If Not obj.Application.Evaluate("@EMPTY(&:[@]:SYS$VARS:CT600_04P1_SET)=1") Then
    'Short Return - No validation for supplementary sheets required
    If obj.Application.Evaluate("+:[@]:P:RETURN_TYPE") = "Short" Then
      IsAvailableAbacus_CT2004 = True
      Exit Function
    End If

    'Supplementary Form A - Tests
    If obj.Application.Evaluate("SR_SUPPA") Then
      If obj.Application.Evaluate("@val(+:[@]:BOX_A1)>0") Then
        If obj.Application.Evaluate("@strlen(+:[@]:SYS$VARS:A50_SET)>0") Then
            If obj.Application.Evaluate("@VAL(+:[@]:[A50]:XMLVERSION)>0") Then
                IsAvailableAbacus_CT2004 = True
            Else
                MsgBox "e-Filing a CT600A, you will need to use a more recent version of the `Loans to participators` schedule.  Please copy your data to a working paper, erase the existing sheet, and re-develop it.  Please retry your submission when you have moved your analysis into the new schedule.", vbCritical, "Online Services"
                IsAvailableAbacus_CT2004 = False
                Exit Function
            End If
        Else
            MsgBox "e-Filing a CT600A, you are required to use the supporting Abacus Schedule `Loans to Participators`.  Please develop this schedule and enter your information.  Please retry your submission when you have analysed out your claim.", vbCritical, "Online Services"
            IsAvailableAbacus_CT2004 = False
            Exit Function
        End If
      End If
    End If
        
    'Supplementary Form B - Tests
    If obj.Application.Evaluate("SR_SUPPB") Then
      If obj.Application.Evaluate("@strlen(+:[@]:SYS$VARS:A39_SET)>0") Then
          If obj.Application.Evaluate("@VAL(+:[@]:[A39]:XMLVERSION)>0") Then
              IsAvailableAbacus_CT2004 = True
          Else
              MsgBox "e-Filing a CT600B, you will need to use a more recent version of the `Controlled foreign companies summary` schedule.  Please copy your data to a working paper, erase the existing sheet, and re-develop it.  Please retry your submission when you have moved your analysis into the new schedule.", vbCritical, "Online Services"
              IsAvailableAbacus_CT2004 = False
              Exit Function
          End If
      Else
          MsgBox "e-Filing a CT600B, you are required to use the supporting Abacus Schedule `Controlled foreign companies summary`.  Please develop this schedule and enter your information.  Please retry your submission when you have analysed out your claim.", vbCritical, "Online Services"
          IsAvailableAbacus_CT2004 = False
          Exit Function
      End If
    End If
    
    If obj.Application.Evaluate("SR_SUPPC") Then
      If obj.Application.Evaluate("@val(+:[@]:BOX_C1)>0") Then
        If obj.Application.Evaluate("@strlen(+:[@]:SYS$VARS:P6_SET)>0") Then
            If obj.Application.Evaluate("@VAL(+:[@]:[P6]:XMLVERSION)>0") Then
                IsAvailableAbacus_CT2004 = True
            Else
                MsgBox "e-Filing a CT600C with a claim to group relief, you will need to use a more recent version of the `Amounts claimed as group relief`.  Please copy your data to a working paper, erase the existing sheet, and re-develop it.  Please retry your submission when you have moved your analysis into the new schedule.", vbCritical, "Online Services"
                IsAvailableAbacus_CT2004 = False
                Exit Function
            End If
        Else
            MsgBox "e-Filing a CT600C with a claim to group relief, you are required to use the supporting Abacus Schedule `Amounts claimed as group relief`.  Please develop this schedule (from `P`) and enter your information.  Please retry your submission when you have analysed out your claim.", vbCritical, "Online Services"
            IsAvailableAbacus_CT2004 = False
            Exit Function
        End If
      End If
      
      If obj.Application.Evaluate("@val(+:[@]:BOX_C2)>0") Then
        If obj.Application.Evaluate("@strlen(+:[@]:SYS$VARS:P7_SET)>0") Then
            If obj.Application.Evaluate("@VAL(+:[@]:[P7]:XMLVERSION)>0") Then
                IsAvailableAbacus_CT2004 = True
            Else
                MsgBox "e-Filing a CT600C with a surrender of group relief, you will need to use a more recent version of the `Amounts surrendered as group relief`.  Please copy your data to a working paper, erase the existing sheet, and re-develop it.  Please retry your submission when you have moved your analysis into the new schedule.", vbCritical, "Online Services"
                IsAvailableAbacus_CT2004 = False
                Exit Function
            End If
        Else
            MsgBox "e-Filing a CT600C with a surrender of group relief, you are required to use the supporting Abacus Schedule `Amounts surrendered as group relief`.  Please develop this schedule (from `P`) and enter your information.  Please retry your submission when you have analysed out your claim.", vbCritical, "Online Services"
            IsAvailableAbacus_CT2004 = False
            Exit Function
        End If
      End If
      
      If obj.Application.Evaluate("@val(+:[@]:BOX_C3)>0") Then
        If obj.Application.Evaluate("@strlen(+:[@]:SYS$VARS:P8_SET)>0") Then
            If obj.Application.Evaluate("@VAL(+:[@]:[P8]:XMLVERSION)>0") Then
                IsAvailableAbacus_CT2004 = True
            Else
                MsgBox "e-Filing a CT600C with a claim to EUFT, you will need to use a more recent version of the `Amounts claimed as EUFT`.  Please copy your data to a working paper, erase the existing sheet, and re-develop it.  Please retry your submission when you have moved your analysis into the new schedule.", vbCritical, "Online Services"
                IsAvailableAbacus_CT2004 = False
                Exit Function
            End If
        Else
            MsgBox "e-Filing a CT600C with a claim to EUFT, you are required to use the supporting Abacus Schedule `Amounts claimed as EUFT`.  Please develop this schedule (from `P`) and enter your information.  Please retry your submission when you have analysed out your claim.", vbCritical, "Online Services"
            IsAvailableAbacus_CT2004 = False
            Exit Function
        End If
      End If
      
      If obj.Application.Evaluate("@val(+:[@]:BOX_C4)>0") Then
        If obj.Application.Evaluate("@strlen(+:[@]:SYS$VARS:P9_SET)>0") Then
            If obj.Application.Evaluate("@VAL(+:[@]:[P9]:XMLVERSION)>0") Then
                IsAvailableAbacus_CT2004 = True
            Else
                MsgBox "e-Filing a CT600C with a surrender of EUFT, you will need to use a more recent version of the `Amounts surrendered as group EUFT`.  Please copy your data to a working paper, erase the existing sheet, and re-develop it.  Please retry your submission when you have moved your analysis into the new schedule.", vbCritical, "Online Services"
                IsAvailableAbacus_CT2004 = False
                Exit Function
            End If
        Else
            MsgBox "e-Filing a CT600C with a surrender of EUFT, you are required to use the supporting Abacus Schedule `Amounts surrendered as group EUFT`.  Please develop this schedule (from `P`) and enter your information.  Please retry your submission when you have analysed out your claim.", vbCritical, "Online Services"
            IsAvailableAbacus_CT2004 = False
            Exit Function
        End If
      End If
    End If
    
    If obj.Application.Evaluate("SR_SUPPD") Then
        If Not obj.Application.Evaluate("(@sheetexists(:[@]:," + Chr(34) + "CT600D_04P1" + Chr(34) + ")>0)") Or Not obj.Application.Evaluate("@getxmlname(&:[@]:CT600D_04P1:SRD_OLAB)=" + Chr(34) + "CT600D_Declaration" + Chr(34)) Then
          MsgBox "To e-File you must update the CT600D form in this computation.  Please erase the CT600D (using Schedule, Erase from the menu options) and re-develop the CT600D.  You can then proceed.", vbCritical, "Online Services"
          IsAvailableAbacus_CT2004 = False
          Exit Function
        End If
    End If
    
    If obj.Application.Evaluate("SR_SUPPE") Then
      If Not obj.Application.Evaluate("(@sheetexists(:[@]:," + Chr(34) + "CT600E_04P1" + Chr(34) + ")>0)") Or Not obj.Application.Evaluate("@getxmlname(&:[@]:CT600E_04P1:SRE_CHREF)=" + Chr(34) + "CT600E_ClaimExemption_RepaymentReference" + Chr(34)) Then
        MsgBox "To e-File you must update the CT600E form in this computation.  Please erase the CT600E (using Schedule, Erase from the menu options) and re-develop the CT600E.  You can then proceed.", vbCritical, "Online Services"
        IsAvailableAbacus_CT2004 = False
        Exit Function
      End If
    End If
    
    If obj.Application.Evaluate("SR_SUPPF") Then
      If obj.Application.Evaluate("@strlen(+:[@]:SYS$VARS:A42_SET)>0") Then
          If obj.Application.Evaluate("@VAL(+:[@]:[A42]:XMLVERSION)>0") Then
              IsAvailableAbacus_CT2004 = True
          Else
              MsgBox "e-Filing a CT600F, you will need to use a more recent version of the 'Tonnage Tax` schedule.  Please copy your data to a working paper, erase the existing sheet, and re-develop it.  Please retry your submission when you have moved your analysis into the new schedule.", vbCritical, "Online Services"
              IsAvailableAbacus_CT2004 = False
              Exit Function
          End If
      Else
          MsgBox "e-Filing a CT600F, you are required to use the supporting Abacus Schedule `Tonnage Tax`.  Please develop this schedule (from `P`) and enter your information.  Please retry your submission when you have analysed out your claim.", vbCritical, "Online Services"
          IsAvailableAbacus_CT2004 = False
          Exit Function
      End If
    End If
    
    If obj.Application.Evaluate("SR_SUPPG") Then
      If Not obj.Application.Evaluate("(@sheetexists(:[@]:," + Chr(34) + "CT600G_04P1" + Chr(34) + ")>0)") Or Not obj.Application.Evaluate("@getxmlname(&:[@]:CT600G_04P2:SRG_INVEST_1NAME)=" + Chr(34) + "CT600G_CoName1" + Chr(34)) Then
        MsgBox "To e-File you must update the CT600G form in this computation.  Please erase the CT600G (using Schedule, Erase from the menu options) and re-develop the CT600G.  You can then proceed.", vbCritical, "Online Services"
        IsAvailableAbacus_CT2004 = False
        Exit Function
      End If
      
      If obj.Application.Evaluate("@strlen(+:[@]:SYS$VARS:A43_SET)>0") Then
          If obj.Application.Evaluate("@VAL(+:[@]:[A43]:XMLVERSION)>0") Then
              IsAvailableAbacus_CT2004 = True
          Else
              MsgBox "e-Filing a CT600G, you will need to use a more recent version of the 'Corporate venturing relief` schedule.  Please copy your data to a working paper, erase the existing sheet, and re-develop it.  Please retry your submission when you have moved your analysis into the new schedule.", vbCritical, "Online Services"
              IsAvailableAbacus_CT2004 = False
              Exit Function
          End If
      Else
          MsgBox "e-Filing a CT600G, you are required to use the supporting Abacus Schedule `Corporate venturing relief`.  Please develop this schedule (from `P`) and enter your information.  Please retry your submission when you have analysed out your claim.", vbCritical, "Online Services"
          IsAvailableAbacus_CT2004 = False
          Exit Function
      End If
      
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_INVEST_1NAME)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_INVEST_2NAME)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_INVEST_3NAME)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_INVEST_4NAME)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_INVEST_5NAME)>0") Then IsAvailableAbacus_CT2004 = False
      
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_INVEST_1NUM)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_INVEST_2NUM)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_INVEST_3NUM)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_INVEST_4NUM)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_INVEST_5NUM)>0") Then IsAvailableAbacus_CT2004 = False
      
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_SHARES_1NAME)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_SHARES_2NAME)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_SHARES_3NAME)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_SHARES_4NAME)>0") Then IsAvailableAbacus_CT2004 = False
      
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_SHARES_1DESC)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_SHARES_2DESC)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_SHARES_3DESC)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_SHARES_4DESC)>0") Then IsAvailableAbacus_CT2004 = False
      
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_DEF_1ISSNAME)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_DEF_2ISSNAME)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_DEF_3ISSNAME)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_DEF_4ISSNAME)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_DEF_5ISSNAME)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_DEF_6ISSNAME)>0") Then IsAvailableAbacus_CT2004 = False
      
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_DEF_1POSTNAME)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_DEF_2POSTNAME)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_DEF_3POSTNAME)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_DEF_4POSTNAME)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_DEF_5POSTNAME)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_DEF_6POSTNAME)>0") Then IsAvailableAbacus_CT2004 = False
      
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_DEF_1DISPDATE)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_DEF_2DISPDATE)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_DEF_3DISPDATE)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_DEF_4DISPDATE)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_DEF_5DISPDATE)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_DEF_6DISPDATE)>0") Then IsAvailableAbacus_CT2004 = False
      
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_1PAR)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_2PAR)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_3PAR)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_4PAR)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_5PAR)>0") Then IsAvailableAbacus_CT2004 = False
      
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_1DATE)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_2DATE)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_3DATE)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_4DATE)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_5DATE)>0") Then IsAvailableAbacus_CT2004 = False
      
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_1SUBS)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_2SUBS)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_3SUBS)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_4SUBS)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_5SUBS)>0") Then IsAvailableAbacus_CT2004 = False
      
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_1REL)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_2REL)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_3REL)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_4REL)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_5REL)>0") Then IsAvailableAbacus_CT2004 = False
      
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_1VENTREF)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_2VENTREF)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_3VENTREF)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_4VENTREF)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_5VENTREF)>0") Then IsAvailableAbacus_CT2004 = False
      
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_1NUM)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_2NUM)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_3NUM)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_4NUM)>0") Then IsAvailableAbacus_CT2004 = False

      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_1ISSDATE)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_2ISSDATE)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_3ISSDATE)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_4ISSDATE)>0") Then IsAvailableAbacus_CT2004 = False
    
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_1DISPDATE)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_2DISPDATE)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_3DISPDATE)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_4DISPDATE)>0") Then IsAvailableAbacus_CT2004 = False
      
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_1APSTART)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_2APSTART)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_3APSTART)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_4APSTART)>0") Then IsAvailableAbacus_CT2004 = False
    
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_1APEND)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_2APEND)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_3APEND)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_4APEND)>0") Then IsAvailableAbacus_CT2004 = False
      
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_1ALLOW)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_2ALLOW)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_3ALLOW)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_4ALLOW)>0") Then IsAvailableAbacus_CT2004 = False
    
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_DEF_1AMOUNT)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_DEF_2AMOUNT)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_DEF_3AMOUNT)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_DEF_4AMOUNT)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_DEF_5AMOUNT)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_DEF_6AMOUNT)>0") Then IsAvailableAbacus_CT2004 = False
      
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_DEF_1NAME)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_DEF_2NAME)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_DEF_3NAME)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_DEF_4NAME)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_DEF_5NAME)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_DEF_6NAME)>0") Then IsAvailableAbacus_CT2004 = False
      
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_DEF_1ISSDATE)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_DEF_2ISSDATE)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_DEF_3ISSDATE)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_DEF_4ISSDATE)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_DEF_5ISSDATE)>0") Then IsAvailableAbacus_CT2004 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_DEF_6ISSDATE)>0") Then IsAvailableAbacus_CT2004 = False
    
      If IsAvailableAbacus_CT2004 = False Then
          MsgBox "e-Filing a CT600G, you are required to use the supporting Abacus Schedule `Corporate venturing relief`.  Please remove all entries from the front of the CT600G Return by erasing the existing sheet, and re-develop it.  Please retry your submission when you have analysed out your claim.", vbCritical, "Online Services"
          Exit Function
      End If
    End If
        
    If obj.Application.Evaluate("@sheetexists(:[@]:," + Chr(34) + "CT600_04P1" + Chr(34) + ")>0") Then
        If Not obj.Application.Evaluate("@strlen(+:[@]:[CT600_04P1]:XMLVERSION)>0") Then
            MsgBox "To e-File you must update the CT600 form in this computation.  Please erase the CT600 (using Schedule, Erase from the menu options) and re-develop the CT600.  You can then proceed.", vbCritical, "Online Services"
            IsAvailableAbacus_CT2004 = False
            Exit Function
        End If
    End If

    IsAvailableAbacus_CT2004 = True
  Else
    IsAvailableAbacus_CT2004 = False
    MsgBox "To e-file you must generate the CT600 form in the computation.  Please develop it from the 'P' Schedule.  You can then Proceed.", vbCritical, "Online Services"
    Exit Function
  End If
  
IsAvailableAbacus_CT2004_End:
  Exit Function

IsAvailableAbacus_CT2004_Err:
  IsAvailableAbacus_CT2004 = False
  Err.Raise Err.Number, ErrorSourceComponent(Err, "IsAvailableAbacus_CT2004", App.EXEName), Err.Description
  Resume IsAvailableAbacus_CT2004_End
End Function
Public Function IsAvailableAbacus_CT2005(ByVal obj As Object) As Boolean
  Dim NewAb As Boolean
  On Error GoTo IsAvailableAbacus_CT2005_Err
  'If Not iNotify Is Nothing Then Call iNotify.Notify(0, 0, "Resuming submission...")
  
  
  Select Case UCase(obj.Pack)
    Case "CPS05C.ABS":
    Case "CPS05I.ABS":
    Case "CPS05B.ABS":
    Case Else
        MsgBox "e-Filing this type of computation for this year end is not supported.  Please file your return in hard-copy.", vbCritical, "Online Services"
        Exit Function
  End Select
  
  'Main form Test
  If Not obj.Application.Evaluate("@EMPTY(&:[@]:SYS$VARS:CT600_05P1_SET)=1") Then
    'Short Return - No validation for supplementary sheets required
    If obj.Application.Evaluate("+:[@]:P:RETURN_TYPE") = "Short" Then
      IsAvailableAbacus_CT2005 = True
      Exit Function
    End If

    'Supplementary Form A - Tests
    If obj.Application.Evaluate("SR_SUPPA") Then
      If obj.Application.Evaluate("@val(+:[@]:BOX_A1)>0") Then
        If obj.Application.Evaluate("@strlen(+:[@]:SYS$VARS:A50_SET)>0") Then
            If obj.Application.Evaluate("@VAL(+:[@]:[A50]:XMLVERSION)>0") Then
                IsAvailableAbacus_CT2005 = True
            Else
                MsgBox "e-Filing a CT600A, you will need to use a more recent version of the `Loans to participators` schedule.  Please copy your data to a working paper, erase the existing sheet, and re-develop it.  Please retry your submission when you have moved your analysis into the new schedule.", vbCritical, "Online Services"
                IsAvailableAbacus_CT2005 = False
                Exit Function
            End If
        Else
            MsgBox "e-Filing a CT600A, you are required to use the supporting Abacus Schedule `Loans to Participators`.  Please develop this schedule and enter your information.  Please retry your submission when you have analysed out your claim.", vbCritical, "Online Services"
            IsAvailableAbacus_CT2005 = False
            Exit Function
        End If
      End If
    End If
        
    'Supplementary Form B - Tests
    If obj.Application.Evaluate("SR_SUPPB") Then
      If obj.Application.Evaluate("@strlen(+:[@]:SYS$VARS:A39_SET)>0") Then
          If obj.Application.Evaluate("@VAL(+:[@]:[A39]:XMLVERSION)>0") Then
              IsAvailableAbacus_CT2005 = True
          Else
              MsgBox "e-Filing a CT600B, you will need to use a more recent version of the `Controlled foreign companies summary` schedule.  Please copy your data to a working paper, erase the existing sheet, and re-develop it.  Please retry your submission when you have moved your analysis into the new schedule.", vbCritical, "Online Services"
              IsAvailableAbacus_CT2005 = False
              Exit Function
          End If
      Else
          MsgBox "e-Filing a CT600B, you are required to use the supporting Abacus Schedule `Controlled foreign companies summary`.  Please develop this schedule and enter your information.  Please retry your submission when you have analysed out your claim.", vbCritical, "Online Services"
          IsAvailableAbacus_CT2005 = False
          Exit Function
      End If
    End If
    
    If obj.Application.Evaluate("SR_SUPPC") Then
      If obj.Application.Evaluate("@val(+:[@]:BOX_C1)>0") Then
        If obj.Application.Evaluate("@strlen(+:[@]:SYS$VARS:P6_SET)>0") Then
            If obj.Application.Evaluate("@VAL(+:[@]:[P6]:XMLVERSION)>0") Then
                IsAvailableAbacus_CT2005 = True
            Else
                MsgBox "e-Filing a CT600C with a claim to group relief, you will need to use a more recent version of the `Amounts claimed as group relief`.  Please copy your data to a working paper, erase the existing sheet, and re-develop it.  Please retry your submission when you have moved your analysis into the new schedule.", vbCritical, "Online Services"
                IsAvailableAbacus_CT2005 = False
                Exit Function
            End If
        Else
            MsgBox "e-Filing a CT600C with a claim to group relief, you are required to use the supporting Abacus Schedule `Amounts claimed as group relief`.  Please develop this schedule (from `P`) and enter your information.  Please retry your submission when you have analysed out your claim.", vbCritical, "Online Services"
            IsAvailableAbacus_CT2005 = False
            Exit Function
        End If
      End If
      
      If obj.Application.Evaluate("@val(+:[@]:BOX_C2)>0") Then
        If obj.Application.Evaluate("@strlen(+:[@]:SYS$VARS:P7_SET)>0") Then
            If obj.Application.Evaluate("@VAL(+:[@]:[P7]:XMLVERSION)>0") Then
                IsAvailableAbacus_CT2005 = True
            Else
                MsgBox "e-Filing a CT600C with a surrender of group relief, you will need to use a more recent version of the `Amounts surrendered as group relief`.  Please copy your data to a working paper, erase the existing sheet, and re-develop it.  Please retry your submission when you have moved your analysis into the new schedule.", vbCritical, "Online Services"
                IsAvailableAbacus_CT2005 = False
                Exit Function
            End If
        Else
            MsgBox "e-Filing a CT600C with a surrender of group relief, you are required to use the supporting Abacus Schedule `Amounts surrendered as group relief`.  Please develop this schedule (from `P`) and enter your information.  Please retry your submission when you have analysed out your claim.", vbCritical, "Online Services"
            IsAvailableAbacus_CT2005 = False
            Exit Function
        End If
      End If
      
      If obj.Application.Evaluate("@val(+:[@]:BOX_C3)>0") Then
        If obj.Application.Evaluate("@strlen(+:[@]:SYS$VARS:P8_SET)>0") Then
            If obj.Application.Evaluate("@VAL(+:[@]:[P8]:XMLVERSION)>0") Then
                IsAvailableAbacus_CT2005 = True
            Else
                MsgBox "e-Filing a CT600C with a claim to EUFT, you will need to use a more recent version of the `Amounts claimed as EUFT`.  Please copy your data to a working paper, erase the existing sheet, and re-develop it.  Please retry your submission when you have moved your analysis into the new schedule.", vbCritical, "Online Services"
                IsAvailableAbacus_CT2005 = False
                Exit Function
            End If
        Else
            MsgBox "e-Filing a CT600C with a claim to EUFT, you are required to use the supporting Abacus Schedule `Amounts claimed as EUFT`.  Please develop this schedule (from `P`) and enter your information.  Please retry your submission when you have analysed out your claim.", vbCritical, "Online Services"
            IsAvailableAbacus_CT2005 = False
            Exit Function
        End If
      End If
      
      If obj.Application.Evaluate("@val(+:[@]:BOX_C4)>0") Then
        If obj.Application.Evaluate("@strlen(+:[@]:SYS$VARS:P9_SET)>0") Then
            If obj.Application.Evaluate("@VAL(+:[@]:[P9]:XMLVERSION)>0") Then
                IsAvailableAbacus_CT2005 = True
            Else
                MsgBox "e-Filing a CT600C with a surrender of EUFT, you will need to use a more recent version of the `Amounts surrendered as group EUFT`.  Please copy your data to a working paper, erase the existing sheet, and re-develop it.  Please retry your submission when you have moved your analysis into the new schedule.", vbCritical, "Online Services"
                IsAvailableAbacus_CT2005 = False
                Exit Function
            End If
        Else
            MsgBox "e-Filing a CT600C with a surrender of EUFT, you are required to use the supporting Abacus Schedule `Amounts surrendered as group EUFT`.  Please develop this schedule (from `P`) and enter your information.  Please retry your submission when you have analysed out your claim.", vbCritical, "Online Services"
            IsAvailableAbacus_CT2005 = False
            Exit Function
        End If
      End If
    End If
    
    If obj.Application.Evaluate("SR_SUPPD") Then
        If Not obj.Application.Evaluate("(@sheetexists(:[@]:," + Chr(34) + "CT600D_04P1" + Chr(34) + ")>0)") Or Not obj.Application.Evaluate("@getxmlname(&:[@]:CT600D_04P1:SRD_OLAB)=" + Chr(34) + "CT600D_Declaration" + Chr(34)) Then
          MsgBox "To e-File you must update the CT600D form in this computation.  Please erase the CT600D (using Schedule, Erase from the menu options) and re-develop the CT600D.  You can then proceed.", vbCritical, "Online Services"
          IsAvailableAbacus_CT2005 = False
          Exit Function
        End If
    End If
    
    If obj.Application.Evaluate("SR_SUPPE") Then
      If Not obj.Application.Evaluate("(@sheetexists(:[@]:," + Chr(34) + "CT600E_05P1" + Chr(34) + ")>0)") Or Not obj.Application.Evaluate("@getxmlname(&:[@]:CT600E_05P1:SRE_CHREF)=" + Chr(34) + "CT600E_ClaimExemption_RepaymentReference" + Chr(34)) Then
        MsgBox "To e-File you must update the CT600E form in this computation.  Please erase the CT600E (using Schedule, Erase from the menu options) and re-develop the CT600E.  You can then proceed.", vbCritical, "Online Services"
        IsAvailableAbacus_CT2005 = False
        Exit Function
      End If
    End If
    
    If obj.Application.Evaluate("SR_SUPPF") Then
      If obj.Application.Evaluate("@strlen(+:[@]:SYS$VARS:A42_SET)>0") Then
          If obj.Application.Evaluate("@VAL(+:[@]:[A42]:XMLVERSION)>0") Then
              IsAvailableAbacus_CT2005 = True
          Else
              MsgBox "e-Filing a CT600F, you will need to use a more recent version of the 'Tonnage Tax` schedule.  Please copy your data to a working paper, erase the existing sheet, and re-develop it.  Please retry your submission when you have moved your analysis into the new schedule.", vbCritical, "Online Services"
              IsAvailableAbacus_CT2005 = False
              Exit Function
          End If
      Else
          MsgBox "e-Filing a CT600F, you are required to use the supporting Abacus Schedule `Tonnage Tax`.  Please develop this schedule (from `P`) and enter your information.  Please retry your submission when you have analysed out your claim.", vbCritical, "Online Services"
          IsAvailableAbacus_CT2005 = False
          Exit Function
      End If
    End If
    
    If obj.Application.Evaluate("SR_SUPPG") Then
      If Not obj.Application.Evaluate("(@sheetexists(:[@]:," + Chr(34) + "CT600G_04P1" + Chr(34) + ")>0)") Or Not obj.Application.Evaluate("@getxmlname(&:[@]:CT600G_04P2:SRG_INVEST_1NAME)=" + Chr(34) + "CT600G_CoName1" + Chr(34)) Then
        MsgBox "To e-File you must update the CT600G form in this computation.  Please erase the CT600G (using Schedule, Erase from the menu options) and re-develop the CT600G.  You can then proceed.", vbCritical, "Online Services"
        IsAvailableAbacus_CT2005 = False
        Exit Function
      End If
      
      If obj.Application.Evaluate("@strlen(+:[@]:SYS$VARS:A43_SET)>0") Then
          If obj.Application.Evaluate("@VAL(+:[@]:[A43]:XMLVERSION)>0") Then
              IsAvailableAbacus_CT2005 = True
          Else
              MsgBox "e-Filing a CT600G, you will need to use a more recent version of the 'Corporate venturing relief` schedule.  Please copy your data to a working paper, erase the existing sheet, and re-develop it.  Please retry your submission when you have moved your analysis into the new schedule.", vbCritical, "Online Services"
              IsAvailableAbacus_CT2005 = False
              Exit Function
          End If
      Else
          MsgBox "e-Filing a CT600G, you are required to use the supporting Abacus Schedule `Corporate venturing relief`.  Please develop this schedule (from `P`) and enter your information.  Please retry your submission when you have analysed out your claim.", vbCritical, "Online Services"
          IsAvailableAbacus_CT2005 = False
          Exit Function
      End If
      
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_INVEST_1NAME)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_INVEST_2NAME)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_INVEST_3NAME)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_INVEST_4NAME)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_INVEST_5NAME)>0") Then IsAvailableAbacus_CT2005 = False
      
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_INVEST_1NUM)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_INVEST_2NUM)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_INVEST_3NUM)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_INVEST_4NUM)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_INVEST_5NUM)>0") Then IsAvailableAbacus_CT2005 = False
      
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_SHARES_1NAME)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_SHARES_2NAME)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_SHARES_3NAME)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_SHARES_4NAME)>0") Then IsAvailableAbacus_CT2005 = False
      
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_SHARES_1DESC)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_SHARES_2DESC)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_SHARES_3DESC)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_SHARES_4DESC)>0") Then IsAvailableAbacus_CT2005 = False
      
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_DEF_1ISSNAME)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_DEF_2ISSNAME)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_DEF_3ISSNAME)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_DEF_4ISSNAME)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_DEF_5ISSNAME)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_DEF_6ISSNAME)>0") Then IsAvailableAbacus_CT2005 = False
      
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_DEF_1POSTNAME)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_DEF_2POSTNAME)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_DEF_3POSTNAME)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_DEF_4POSTNAME)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_DEF_5POSTNAME)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_DEF_6POSTNAME)>0") Then IsAvailableAbacus_CT2005 = False
      
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_DEF_1DISPDATE)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_DEF_2DISPDATE)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_DEF_3DISPDATE)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_DEF_4DISPDATE)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_DEF_5DISPDATE)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P2:SRG_DEF_6DISPDATE)>0") Then IsAvailableAbacus_CT2005 = False
      
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_1PAR)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_2PAR)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_3PAR)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_4PAR)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_5PAR)>0") Then IsAvailableAbacus_CT2005 = False
      
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_1DATE)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_2DATE)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_3DATE)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_4DATE)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_5DATE)>0") Then IsAvailableAbacus_CT2005 = False
      
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_1SUBS)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_2SUBS)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_3SUBS)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_4SUBS)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_5SUBS)>0") Then IsAvailableAbacus_CT2005 = False
      
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_1REL)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_2REL)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_3REL)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_4REL)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_5REL)>0") Then IsAvailableAbacus_CT2005 = False
      
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_1VENTREF)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_2VENTREF)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_3VENTREF)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_4VENTREF)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_INVEST_5VENTREF)>0") Then IsAvailableAbacus_CT2005 = False
      
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_1NUM)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_2NUM)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_3NUM)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_4NUM)>0") Then IsAvailableAbacus_CT2005 = False

      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_1ISSDATE)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_2ISSDATE)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_3ISSDATE)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_4ISSDATE)>0") Then IsAvailableAbacus_CT2005 = False
    
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_1DISPDATE)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_2DISPDATE)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_3DISPDATE)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_4DISPDATE)>0") Then IsAvailableAbacus_CT2005 = False
      
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_1APSTART)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_2APSTART)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_3APSTART)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_4APSTART)>0") Then IsAvailableAbacus_CT2005 = False
    
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_1APEND)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_2APEND)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_3APEND)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_4APEND)>0") Then IsAvailableAbacus_CT2005 = False
      
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_1ALLOW)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_2ALLOW)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_3ALLOW)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_SHARES_4ALLOW)>0") Then IsAvailableAbacus_CT2005 = False
    
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_DEF_1AMOUNT)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_DEF_2AMOUNT)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_DEF_3AMOUNT)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_DEF_4AMOUNT)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_DEF_5AMOUNT)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_DEF_6AMOUNT)>0") Then IsAvailableAbacus_CT2005 = False
      
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_DEF_1NAME)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_DEF_2NAME)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_DEF_3NAME)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_DEF_4NAME)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_DEF_5NAME)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_DEF_6NAME)>0") Then IsAvailableAbacus_CT2005 = False
      
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_DEF_1ISSDATE)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_DEF_2ISSDATE)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_DEF_3ISSDATE)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_DEF_4ISSDATE)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_DEF_5ISSDATE)>0") Then IsAvailableAbacus_CT2005 = False
      If obj.Application.Evaluate("@strlen(+:[@]:CT600G_04P3:SRG_DEF_6ISSDATE)>0") Then IsAvailableAbacus_CT2005 = False
    
      If IsAvailableAbacus_CT2005 = False Then
          MsgBox "e-Filing a CT600G, you are required to use the supporting Abacus Schedule `Corporate venturing relief`.  Please remove all entries from the front of the CT600G Return by erasing the existing sheet, and re-develop it.  Please retry your submission when you have analysed out your claim.", vbCritical, "Online Services"
          Exit Function
      End If
    End If
        
    If obj.Application.Evaluate("@sheetexists(:[@]:," + Chr(34) + "CT600_05P1" + Chr(34) + ")>0") Then
        If Not obj.Application.Evaluate("@strlen(+:[@]:[CT600_05P1]:XMLVERSION)>0") Then
            MsgBox "To e-File you must update the CT600 form in this computation.  Please erase the CT600 (using Schedule, Erase from the menu options) and re-develop the CT600.  You can then proceed.", vbCritical, "Online Services"
            IsAvailableAbacus_CT2005 = False
            Exit Function
        End If
    End If

    IsAvailableAbacus_CT2005 = True
  Else
    IsAvailableAbacus_CT2005 = False
    MsgBox "To e-file you must generate the CT600 form in the computation.  Please develop it from the 'P' Schedule.  You can then Proceed.", vbCritical, "Online Services"
    Exit Function
  End If
  
IsAvailableAbacus_CT2005_End:
  Exit Function

IsAvailableAbacus_CT2005_Err:
  IsAvailableAbacus_CT2005 = False
  Err.Raise Err.Number, ErrorSourceComponent(Err, "IsAvailableAbacus_CT2005", App.EXEName), Err.Description
  Resume IsAvailableAbacus_CT2005_End
End Function

' Uses UI - issue making code work in unattended mode
Private Function TidyExistingSubmissions(ByVal EntityID As String, Optional ByVal iNotify As IBaseNotify) As Boolean
  Dim status As String, s As String, bCancelled As Boolean
  Dim rs As Recordset, frm As frmTimeOut, Retry As Boolean
      
  On Error GoTo TidyExistingSubmissions_Err
  If Not iNotify Is Nothing Then Call iNotify.Notify(0, 0, "Checking existing submissions...")
  
  TidyExistingSubmissions = True
  'find out what the current status is
  Set rs = DA.StalledSubmissions("EntityID = " & gDBHelper.StrSQL(EntityID), "SubmissionDate", True)
  If Not (rs.BOF And rs.EOF) Then
    Set frm = New frmTimeOut
    frm.Caption = "Unresolved Submission"
    frm.lblMessage.Caption = "There is an existing submission which has not been resolved." & vbCrLf & vbCrLf & "Press 'Continue' to attempt to resolve it automatically, else press 'Cancel' to resolve it later."
    frm.Show vbModal
    Retry = frm.Retry
    If Retry Then
      While Not rs.EOF
        Call ProcessSubmission(rs.Fields("SubmissionID"), , bCancelled)
        If bCancelled Then
          TidyExistingSubmissions = False
          GoTo TidyExistingSubmissions_End
        End If
        rs.MoveNext
      Wend
      Set rs = DA.StalledSubmissions("EntityID = " & gDBHelper.StrSQL(EntityID))
      If Not (rs.BOF And rs.EOF) Then Err.Raise ERR_FAILED_SUBMISSIONS, "TidyExistingSubmissions", "Cannot submit because there is an existing submission which has not been resolved"
    Else
      TidyExistingSubmissions = False
    End If
  End If
  
  
TidyExistingSubmissions_End:
  Call CloseRecordSet(rs)
  Set frm = Nothing
  Set rs = Nothing
  Exit Function
  
TidyExistingSubmissions_Err:
  Err.Raise Err.Number, ErrorSourceComponent(Err, "TidyExistingSubmissions", App.EXEName), Err.Description
  Resume TidyExistingSubmissions_End
End Function

Public Function ValidateXML(ByVal XMLBody As String) As String
'  Dim xmldoc As DOMDocument, httpreq As xmlhttp, ResponseDoc As DOMDocument
'  Dim currnode  As IXMLDOMNode
'  Dim s As String, SubID As String, MsgType As String, bTryAgain As Boolean, bRequiresDelete As Boolean
'  Dim sResponseText As String
'
'  On Error GoTo ValidateXML_Err
'
'  'NEED TO CHECK THAT ALL REQUIRED PROPERTIES HAVE BEEN SET BEFORE PROCEEDING.  ELSE RAISE ERROR PQ
'  If Len(mPeriodEnd) = 0 Then Err.Raise ERR_REQUIRED_PROPERTY, "ValidateXML", "PeriodEnd is a required property and has not been set."
'  If Len(mSchVers) = 0 Then Err.Raise ERR_REQUIRED_PROPERTY, "ValidateXML", "Schema Version is a required property and has not been set."
'  If mSubmitType = 0 Then Err.Raise ERR_REQUIRED_PROPERTY, "ValidateXML", "Submission Type is a required property and has not been set."
'  If mRole = [_SR_DEFAULT] Then Err.Raise ERR_REQUIRED_PROPERTY, "ValidateXML", "ValidateXMLter Role is a required property and has not been set."
'
'  Set xmldoc = New DOMDocument
'  If Not xmldoc.loadXML(XMLBody) Then Err.Raise ERR_INVALID_XML, , "The xml supplied is not well formed."
'
'  Select Case SchemaVersion
'    Case S_SCHEMAVERSION_CT2004
'       Set xmldoc = WrapData_CT2004(xmldoc, "", "")
'    Case S_SCHEMAVERSION_CT2002 ' CT2002 & P11D
'    Case S_SCHEMAVERSION_P11D2003
'       Set xmldoc = WrapData(xmldoc, "", "")
'    Case Else
'       ' Unknown Schema Version error
'       Err.Raise ERR_INVALID_SCHEMAVERSION, "Submit", "Unknown schema version."
'  End Select
'
'  Set xmldoc = Encrypt(xmldoc)
'
'  Set ResponseDoc = SubmitEx(sResponseText, IIf(mSubmitType = CorporationTax, S_GG_TEST_TPVS_SITE_CT, S_GG_TEST_TPVS_SITE_P11D), xmldoc.xml, False, CheckProxy)
'  With ResponseDoc
'    Set currnode = .selectSingleNode("//GovTalkMessage/Header/MessageDetails/Qualifier")
'    MsgType = currnode.Text
'
'    Select Case LCase(MsgType)
'      Case "error"
'          s = ErrorResponseMsg(ResponseDoc, bTryAgain, bRequiresDelete)
'      Case "response"
'          s = "The body of your xml has successfully validated against the business rules."
'      Case Else
'          s = "Unexpected message type returned [" & MsgType & "].  Please try again."
'      End Select
'  End With
'
'
'ValidateXML_End:
'  ValidateXML = s
'  Set xmldoc = Nothing
'  Set ResponseDoc = Nothing
'  Set currnode = Nothing
'  Exit Function
'
'ValidateXML_Err:
'  Err.Raise Err.Number, ErrorSourceComponent(Err, "ValidateXML", App.EXEName), Err.Description
'  Resume
  Call Err.Raise(380, "ValidateXML", "Function is no longer valid")
End Function

'FIX #865
Private Sub CheckNetwork()
  Dim InetParam As Long
  Dim InetRtn As Boolean
  
  On Error GoTo CheckNetwork_Err
  If Not mMockSub Then
    InetRtn = InternetGetConnectedState(InetParam, 0)
    If Not InetRtn Then
      Err.Raise ERR_NETWORK, "CheckNetwork", "No internet connection detected - please ensure your internet connection is valid and try again"
    End If
  End If
  
  Exit Sub
  
CheckNetwork_Err:
  Err.Raise Err.Number, ErrorSourceComponent(Err, "CheckNetwork", App.EXEName), Err.Description
  Resume
End Sub


