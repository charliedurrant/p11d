VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_IDSMail As IDSMailInterface32.Server
Private m_IDSAddressBook As IDSMailInterface32.AddressBook
Private m_UseAddressBook As IDSM_AB_SEEK_TYPES
Private m_MailApplication As MAIL_APPLICATION
Private m_Success As Boolean 'Flag for successful completion of e-mail
Private m_OwnerForm As Form 'Parent form to associate Setup/Test forms
Private m_AutoDetect As Boolean
Private m_TestMailMode As Boolean
Private m_TestMailRecipient As String 'Default Recipient of test Emails
Private bIgnoreTestMailMode As Boolean 'Flag to handle "NO" requests to test email
'outlookredemption
'not late binding, redeption creates methods on the fly!?!
Private m_RedeptionApplication As Object
Private m_RedeptionNameSpace As Object
Private m_RedeptionSafeMailItem As Object

Private m_Attachments As Collection

Public Enum IDSM_AB_SEEK_TYPES
  IDSM_AB_SEEK_NONE = 0
  IDSM_AB_SEEK_FULL = 1
  IDSM_AB_SEEK_PARTIAL = 2
End Enum

Public Enum MAIL_APPLICATION
  MA_OTHER = 0
  MA_MICROSOFT_MAIL = 1
  MA_MICROSOFT_EXCHANGE = 2
  MA_LOTUS_CC_MAIL = 3
  MA_LOTUS_NOTES_VIM = 4
  MA_LOTUS_NOTES_API = 5
  MA_MICROSOFT_OUTLOOK = 6
  MA_NOVELL_GROUPWISE_V5_PLUS = 7
  MA_PEGASUS = 9
End Enum

Public Enum IDSM_MAIL_SYSTEM
  IDSM_SYS_UNKNOWN = 0
  IDSM_SYS_MAPI = 1
  IDSM_SYS_VIM = 2
  IDSM_SYS_MHS = 3
  IDSM_SYS_CSERVE = 5
  IDSM_SYS_SMTP_POP = 10
  IDSM_SYS_VINES = 20
  IDSM_SYS_ACTIVE_MESSAGING = 9
  IDSM_SYS_RAS = 12
  IDSM_SYS_NOTES = 11
  IDSM_SYS_MAPI_EXTENDED = 3 ' 'Used for outlook extended mapi as MHS not actually supported in IDSMail so can use its enum value
End Enum

Public Enum IDMS_MESSAGE_ITEM
  IDSM_ITM_TO = 1
  IDSM_ITM_CC = 2
  IDSM_ITM_BCC = 3
  IDSM_ITM_SUBJECT = 4
  IDSM_ITM_MESSAGE = 5
  IDSM_ITM_ATTACHMENT = 6
End Enum
Private Const IDSM_ITEM_FIRST_CLEARABLE_ITEM = IDSM_ITM_TO
Private Const IDSM_ITEM_LAST_CLEARABLE_ITEM = IDSM_ITM_ATTACHMENT

Private Enum OUTLOOK_RECIPIENT_TYPE
  olBCC = 3  'The recipient is specified in the BCC property of the Item.
  olCC = 2 'The recipient is specified in the CC property of the Item.
  olOriginator = 0  'Originator (sender) of the Item.
  olTo = 1
End Enum



Private Sub Class_Initialize()
  Set m_IDSMail = New IDSMailInterface32.Server
  m_IDSMail.objectKey = "HX4VEAB68HQ2"
  m_IDSMail.ErrorDisplay = False
  m_TestMailRecipient = S_DEFAULT_TEST_RECIPIENT
  bIgnoreTestMailMode = False
  Call RegistrySettings(REGISTRY_KEY_READ)
  Set m_Attachments = New Collection
End Sub
Private Sub RedeptionInit()
 Dim loExplorer As Object
 
 On Error GoTo err_err
 
 If (m_RedeptionApplication Is Nothing) Then
  Set m_RedeptionApplication = CreateObject("Outlook.Application")
  
  Set m_RedeptionNameSpace = m_RedeptionApplication.GetNamespace("MAPI")
  
 End If
err_end:
  Exit Sub
err_err:
End Sub


Private Sub RedeptionTerminate()
 If Not (m_RedeptionApplication Is Nothing) Then
  Set m_RedeptionApplication = Nothing
  Set m_RedeptionNameSpace = Nothing
 End If
End Sub

Private Sub Class_Terminate()
 Dim utils As Object
 
On Error GoTo err_err

 Call RegistrySettings(REGISTRY_KEY_WRITE)
 Set m_IDSMail = Nothing
 Set m_IDSAddressBook = Nothing
 Set m_RedeptionSafeMailItem = Nothing
 Set m_RedeptionNameSpace = Nothing
 
 
 If (Not m_RedeptionApplication Is Nothing) Then
   Set utils = CreateObject("Redemption.MAPIUtils")
   Call utils.Cleanup
   Set utils = Nothing
 
   Call m_RedeptionApplication.Quit
   Set m_RedeptionApplication = Nothing
 End If
 
err_end:
  Exit Sub
err_err:
  Resume err_end
End Sub
Private Property Get ILibraryVersion_Name() As String
  ILibraryVersion_Name = "MAIL library"
End Property

Private Property Get ILibraryVersion_Version() As String
  ILibraryVersion_Version = app.major & "." & app.minor & "." & app.Revision
End Property
Public Property Get RedeptionUsing() As Boolean
  RedeptionUsing = MailSystem = IDSM_MAIL_SYSTEM.IDSM_SYS_MAPI_EXTENDED
End Property

Private Function PreviouslyAddedAttachementsToVariantArray() As Variant
  If (m_Attachments.Count = 0) Then
    PreviouslyAddedAttachementsToVariantArray = Empty
    Exit Function
  End If
    
  Dim v() As Variant
  Dim i As Long, j As Long
  
  ReDim v(1 To m_Attachments.Count)
  For i = 1 To m_Attachments.Count
    v(i) = CStr(m_Attachments(i))
  Next
  j = m_Attachments.Count
  For i = j To 1 Step -1
    m_Attachments.Remove (i)
  Next
  PreviouslyAddedAttachementsToVariantArray = v
End Function
Public Sub MailSend(ByVal msgSubject As String, ByVal msgBody As String, ToList As Variant, Optional CCList As Variant, Optional BCCList As Variant, Optional Attachments As Variant)
  Dim vPreviouslyAddedAttachements As Variant
  Dim utils As Object
  Dim Btn As Object
  On Error GoTo MailSend_Err
  
  
  If (m_Success = False And Not m_TestMailMode) Then
    If Not bIgnoreTestMailMode Then Call ShowOptions
  End If
  
  If RedeptionUsing Then
    Call RedeptionInit
    If m_RedeptionSafeMailItem Is Nothing Then 'what a mess, RKs model is all up its arse
      Call NewMessage
    End If
  End If
  

  vPreviouslyAddedAttachements = PreviouslyAddedAttachementsToVariantArray()
  Call AddRecipients(ToList, CCList, BCCList)
  Call AddAttachments(Attachments)
  Call AddAttachments(vPreviouslyAddedAttachements)
  
  If (RedeptionUsing) Then
    m_RedeptionSafeMailItem.subject = msgSubject
    m_RedeptionSafeMailItem.HTMLBody = "<html><body><pre>" & msgBody & "</pre></body></html>"
    
    Const PR_InetMailOverrideFormat As Long = &H59020003
    Const ENCODING_PREFERENCE As Long = &H20000
    Const BODY_ENCODING_TEXT_AND_HTML As Long = &H100000
    Const ENCODING_MIME As Long = &H40000
    
    Const PR_MSG_EDITOR_FORMAT As Long = &H59090003
    Const EDITOR_FORMAT_PLAINTEXT As Long = 1
    Const EDITOR_FORMAT_HTML As Long = 2

    m_RedeptionSafeMailItem.Fields(PR_InetMailOverrideFormat) = ENCODING_PREFERENCE Or ENCODING_MIME Or BODY_ENCODING_TEXT_AND_HTML
    m_RedeptionSafeMailItem.Fields(PR_MSG_EDITOR_FORMAT) = EDITOR_FORMAT_HTML

    Call m_RedeptionSafeMailItem.send
    
    Set utils = CreateObject("Redemption.MAPIUtils")
    If (Not utils Is Nothing) Then Call utils.DeliverNow
    
    Set m_RedeptionSafeMailItem = Nothing
  Else
    m_IDSMail.subject = msgSubject
    m_IDSMail.Message = msgBody
    m_IDSMail.send
  End If
  
  
  'Set successful flag if at least one e-mail can be sent
  If Not m_Success Then
    m_Success = True
    Call RegistrySettings(REGISTRY_KEY_WRITE)
  End If
  
MailSend_End:
  Exit Sub
MailSend_Err:
  Set m_RedeptionSafeMailItem = Nothing
  Err.Raise Err.Number, Err.Source, Err.Description
  Resume MailSend_End
  Resume
End Sub
Private Sub RedemptionAddRecipient(ByVal email As String, recipientType As OUTLOOK_RECIPIENT_TYPE)
  Dim recipient As Object
  
  Set recipient = m_RedeptionSafeMailItem.Recipients.Add(email)
  If (recipient.Type <> recipientType) Then
    recipient.Type = recipientType
  End If
  
  'don't resolve
End Sub
Private Function AddRecipients(ToList As Variant, CCList As Variant, BCCList As Variant) As Boolean
  Dim i As Long
  Dim recipient As Object
  On Error GoTo AddRecipients_err
  
  'Elements could be passed individually or as an array
    If Not IsEmpty(ToList) Then
      If IsArray(ToList) Then
        For i = LBound(ToList) To UBound(ToList)
          If (RedeptionUsing) Then
            Call RedemptionAddRecipient(CStr(ToList(i)), OUTLOOK_RECIPIENT_TYPE.olTo)
          Else
            m_IDSMail.AddRecipientTo LookUpName(CStr(ToList(i)))
          End If
        Next i
      ElseIf VarType(ToList) = vbString Then
          If RedeptionUsing Then
            Call RedemptionAddRecipient(CStr(ToList), OUTLOOK_RECIPIENT_TYPE.olTo)
          Else
            m_IDSMail.AddRecipientTo LookUpName(CStr(ToList))
          End If
      End If
    End If
    If Not IsEmpty(CCList) Then
      If IsArray(CCList) Then
        For i = LBound(CCList) To UBound(CCList)
          If (RedeptionUsing) Then
            Call RedemptionAddRecipient(CStr(ToList(i)), OUTLOOK_RECIPIENT_TYPE.olCC)
          Else
            m_IDSMail.AddRecipientCc LookUpName(CStr(CCList(i)))
          End If
        Next i
      ElseIf VarType(CCList) = vbString Then
          If (RedeptionUsing) Then
            Call RedemptionAddRecipient(CStr(ToList), OUTLOOK_RECIPIENT_TYPE.olCC)
          Else
            m_IDSMail.AddRecipientCc LookUpName(CStr(CCList))
          End If
      End If
    End If
    If Not IsEmpty(BCCList) Then
      If IsArray(BCCList) Then
        For i = LBound(BCCList) To UBound(BCCList)
          If (RedeptionUsing) Then
            Call RedemptionAddRecipient(CStr(ToList(i)), OUTLOOK_RECIPIENT_TYPE.olBCC)
          Else
            m_IDSMail.AddRecipientBcc LookUpName(CStr(BCCList(i)))
          End If
        Next i
      ElseIf VarType(BCCList) = vbString Then
          If (RedeptionUsing) Then
            Call RedemptionAddRecipient(CStr(ToList), OUTLOOK_RECIPIENT_TYPE.olBCC)
          Else
            m_IDSMail.AddRecipientBcc LookUpName(CStr(BCCList))
          End If
      End If
    End If
    AddRecipients = True

AddRecipients_end:
  Exit Function

AddRecipients_err:
  AddRecipients = False
  Err.Raise Err.Number, Err.Source, Err.Description
  Resume AddRecipients_end
End Function

Private Function SetMailSystem() As Boolean
  On Error GoTo SetMailSystem_Err
  
  m_IDSMail.SetNativeSystem
  
SetMailSystem_End:
  Exit Function
  
SetMailSystem_Err:
  SetMailSystem = False
  'Err.Raise Err.Number, "SetMailSystem", Err.Description

End Function

Private Function AddAttachments(Attachments As Variant) As Boolean
  Dim i As Long
  On Error GoTo AddAttachments_err
  If (Not IsEmpty(Attachments)) And IsArray(Attachments) Then
    For i = LBound(Attachments) To UBound(Attachments)
      If (RedeptionUsing) Then
        Call m_RedeptionSafeMailItem.Attachments.Add(CStr(Attachments(i)))
      Else
        m_IDSMail.AddAttachment CStr(Attachments(i))
      End If
    Next i
  End If
  AddAttachments = True
 
AddAttachments_end:
  Exit Function
    
AddAttachments_err:
  AddAttachments = False
  Err.Raise Err.Number, Err.Source, Err.Description
  Resume AddAttachments_end
End Function
Public Property Get UseAddressBook() As IDSM_AB_SEEK_TYPES
  UseAddressBook = m_UseAddressBook
End Property

Public Property Let UseAddressBook(ByVal NewValue As IDSM_AB_SEEK_TYPES)
  If Not NewValue = IDSM_AB_SEEK_NONE Then
    If m_IDSAddressBook Is Nothing Then
      Set m_IDSAddressBook = New IDSMailInterface32.AddressBook
      m_IDSAddressBook.MailObject = m_IDSMail
    End If
  End If
  m_UseAddressBook = NewValue
  
End Property

Private Function LookUpName(sLookUpName As String) As String
  On Error GoTo LookupName_err
  Dim sDefaultProfile As String
  
  If m_UseAddressBook Then
    m_IDSMail.Login
    m_IDSAddressBook.OpenAddressBook
    m_IDSAddressBook.MoveFirst
    m_IDSAddressBook.SeekName m_UseAddressBook, sLookUpName
  
    If m_IDSAddressBook.NoMatch = True Then
      'MsgBox "Name not found." 'RK Debug/Raise custom error? or try name provided?
      LookUpName = sLookUpName
    Else
      'MsgBox "Found:" + m_IDSAddressBook.Name
      LookUpName = m_IDSAddressBook.Name
    End If
    LookUpName = m_IDSAddressBook.Name
  Else
    LookUpName = sLookUpName
  End If
 
LookupName_end:
  Exit Function
    
LookupName_err:
  LookUpName = False
  Err.Raise Err.Number, Err.Source, Err.Description
  Resume LookupName_end
End Function

Friend Property Get MailApplication() As MAIL_APPLICATION
  MailApplication = m_MailApplication
End Property

Friend Property Let MailApplication(ByVal NewValue As MAIL_APPLICATION)
  m_MailApplication = NewValue
End Property

Friend Property Get IDSMail() As IDSMailInterface32.Server 'RK Expose entire object?
  Set IDSMail = m_IDSMail
End Property

Friend Property Let MailSystem(ByVal NewValue As IDSM_MAIL_SYSTEM)
  m_IDSMail.MailSystem = NewValue
End Property
Friend Property Get MailSystem() As IDSM_MAIL_SYSTEM
  MailSystem = m_IDSMail.MailSystem
End Property
Friend Property Let Success(ByVal NewValue As Boolean)
  m_Success = NewValue
End Property
Friend Property Get Success() As Boolean
  Success = m_Success
End Property
Friend Property Let TestMailMode(ByVal NewValue As Boolean)
  m_TestMailMode = NewValue
End Property
Friend Property Get TestMailMode() As Boolean
  TestMailMode = m_TestMailMode
End Property
Public Property Set OwnerForm(ByVal NewValue As Object)
  If TypeOf NewValue Is Form Then 'RK raise error if not form?
    Set m_OwnerForm = NewValue
  End If
End Property
Public Property Let AutoDetect(ByVal NewValue As Boolean)
  If NewValue And Not m_Success Then
    m_Success = True
    'SendTestMail
  End If
  m_AutoDetect = NewValue
End Property
Public Property Get AutoDetect() As Boolean
  AutoDetect = m_AutoDetect
End Property
Friend Property Let LoginName(ByVal NewValue As String)
  m_IDSMail.LoginName = NewValue
End Property
Friend Property Get LoginName() As String
  LoginName = m_IDSMail.LoginName
End Property

Public Property Let TestMailRecipient(ByVal NewValue As String)
  m_TestMailRecipient = NewValue
End Property
Public Property Get TestMailRecipient() As String
  TestMailRecipient = m_TestMailRecipient
End Property

Friend Sub RegistrySettings(Action As RegistryKeyAction)
  On Error GoTo RegistrySettings_Err
    If Action = REGISTRY_KEY_WRITE Then
      Call SaveSetting(S_REGISTRY_APPNAME, S_REGISTRY_SECTION, "MailApplication", m_MailApplication)
      Call SaveSetting(S_REGISTRY_APPNAME, S_REGISTRY_SECTION, "MailSystem", m_IDSMail.MailSystem)
      Call SaveSetting(S_REGISTRY_APPNAME, S_REGISTRY_SECTION, "LoginName", m_IDSMail.LoginName)
      Call SaveSetting(S_REGISTRY_APPNAME, S_REGISTRY_SECTION, "UseAddressBook", m_UseAddressBook)
      Call SaveSetting(S_REGISTRY_APPNAME, S_REGISTRY_SECTION, "Success", m_Success)
    ElseIf Action = REGISTRY_KEY_READ Then
      m_MailApplication = GetSetting(S_REGISTRY_APPNAME, S_REGISTRY_SECTION, "MailApplication", MA_OTHER)
      m_IDSMail.MailSystem = GetSetting(S_REGISTRY_APPNAME, S_REGISTRY_SECTION, "MailSystem", IDSM_SYS_UNKNOWN)
      m_IDSMail.LoginName = GetSetting(S_REGISTRY_APPNAME, S_REGISTRY_SECTION, "LoginName", "")
      m_UseAddressBook = GetSetting(S_REGISTRY_APPNAME, S_REGISTRY_SECTION, "UseAddressBook", IDSM_AB_SEEK_NONE)
      m_Success = GetSetting(S_REGISTRY_APPNAME, S_REGISTRY_SECTION, "Success", 0)
    End If

RegistrySettings_End:
  Exit Sub

RegistrySettings_Err:
  Err.Raise Err.Number, Err.Source, Err.Description
  Resume RegistrySettings_End
End Sub
Public Sub ShowOptions()
  Dim frm As Form
  On Error GoTo ShowOptions_Err
    
    m_IDSMail.ErrorDisplay = True
    Set frm = New frmMailSetup
    Set frm.m_Mail = Me
    
    'If specified set as child form of application
    If Not m_OwnerForm Is Nothing Then
      Call frm.Show(vbModal, m_OwnerForm)
    Else
      Call frm.Show(vbModal)
    End If
 
ShowOptions_End:
  Exit Sub

ShowOptions_Err:
  Err.Raise Err.Number, Err.Source, Err.Description
  Resume ShowOptions_End
  Resume
End Sub

Public Sub AddAttachment(sAttachment As String)
  Call m_Attachments.Add(sAttachment)
End Sub

Public Sub ClearMessageItem(MessageItem As IDMS_MESSAGE_ITEM)
  On Error GoTo ClearMessageItem_Err
  
  Call m_IDSMail.ClearMessageItems(CInt(MessageItem))

ClearMessageItem_End:
  Exit Sub

ClearMessageItem_Err:
  Err.Raise Err.Number, Err.Source, Err.Description
  Resume ClearMessageItem_End
End Sub
Public Sub ClearAllMessageItems()
  Dim i As IDMS_MESSAGE_ITEM
  On Error GoTo ClearAllMessageItems_Err
  
  For i = IDSM_ITEM_FIRST_CLEARABLE_ITEM To IDSM_ITEM_LAST_CLEARABLE_ITEM
    Call m_IDSMail.ClearMessageItems(CInt(i))
  Next i

ClearAllMessageItems_End:
  Exit Sub

ClearAllMessageItems_Err:
  Err.Raise Err.Number, Err.Source, Err.Description
  Resume ClearAllMessageItems_End
End Sub
Public Sub AddRecipientTo(ToName As String)
  On Error GoTo AddRecipientTo_Err
  
  Call m_IDSMail.AddRecipientTo(ToName)

AddRecipientTo_End:
  Exit Sub

AddRecipientTo_Err:
  Err.Raise Err.Number, Err.Source, Err.Description
  Resume AddRecipientTo_End
End Sub
Public Sub AddRecipientCc(CCName As String)
  On Error GoTo AddRecipientCc_Err
  
  Call m_IDSMail.AddRecipientCc(CCName)

AddRecipientCc_End:
  Exit Sub

AddRecipientCc_Err:
  Err.Raise Err.Number, Err.Source, Err.Description
  Resume AddRecipientCc_End
End Sub
Public Sub AddRecipientBcc(BCCName As String)
  On Error GoTo AddRecipientBcc_Err
  
  Call m_IDSMail.AddRecipientBcc(BCCName)
  
AddRecipientBcc_End:
  Exit Sub

AddRecipientBcc_Err:
  Err.Raise Err.Number, Err.Source, Err.Description
  Resume AddRecipientBcc_End
End Sub
Private Sub SendTestMail()
  Dim bMsgBoxRes As VbMsgBoxResult
  Dim frm As frmMailSetup
  On Error GoTo SendTestMail_Err
  
    m_TestMailMode = True
    
    'Display MailSetup if not previously connected
    Set frm = New frmMailSetup
    Set frm.m_Mail = Me
    
    'Autodetect options
    Call frm.DetectDefaultSettings
    Call RegistrySettings(REGISTRY_KEY_WRITE)
    bMsgBoxRes = MsgBox("You have not used e-mail functionality before." & vbCrLf & _
                      "Do you wish to send a test mail to " & m_TestMailRecipient & "?", vbYesNo, "Test Email")
    If bMsgBoxRes = vbYes Then Call frm.SetupTestForm(m_OwnerForm)
    If bMsgBoxRes = vbNo Then bIgnoreTestMailMode = True
    m_TestMailMode = False
SendTestMail_End:
  Exit Sub

SendTestMail_Err:
  Err.Raise Err.Number, Err.Source, Err.Description
  Resume SendTestMail_End
  Resume
End Sub

Public Sub NewMessage()
  Dim oItem As Object
  On Error GoTo NewMessage_Err
  
  If RedeptionUsing Then
    Call RedeptionInit
    Set m_RedeptionSafeMailItem = CreateObject("Redemption.SafeMailItem") 'Create an instance of Redemption.SafeMailItem
    Set oItem = m_RedeptionApplication.CreateItem(0) 'Create a new message
    m_RedeptionSafeMailItem.Item = oItem
  Else
    Call m_IDSMail.NewMessage
  End If
  
  
NewMessage_End:
  Exit Sub

NewMessage_Err:
  Err.Raise Err.Number, Err.Source, Err.Description
  Resume NewMessage_End
End Sub



