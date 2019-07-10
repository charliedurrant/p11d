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
Attribute VB_Exposed = True
Option Explicit
Private m_UseDefaultProfile As Boolean 'RK to remove
Private m_UseAddressBook As IDSM_AB_SEEK_TYPES
Private m_IDSMail As IDSMailInterface32.Server
Private m_IDSAddressBook As IDSMailInterface32.AddressBook
Private m_MailApplication As MAIL_APPLICATIONS
Private m_LoginName As String
Private m_CanEmail As String 'Flag for successful completion of e-mail

Public Enum IDSM_AB_SEEK_TYPES
  IDSM_AB_SEEK_NONE = 0
  IDSM_AB_SEEK_FULL = 1
  IDSM_AB_SEEK_PARTIAL = 2
End Enum

Public Enum MAIL_APPLICATIONS
  MA_OTHER = 0
  MA_MICROSOFT_MAIL = 1
  MA_MICROSOFT_EXCHANGE_OUTLOOK = 2
  MA_LOTUS_CC_MAIL = 3
  MA_LOTUS_NOTES_VIM = 4
  MA_LOTUS_NOTES_API = 5
  MA_NOVELL_GROUPWISE_V4 = 6
  MA_NOVELL_GROUPWISE_V5_PLUS = 7
  MA_DAVINCI = 8
  MA_PEGASUS = 9
  MA_BANYAN_BEYONDMAIL = 10
  MA_SHARKMAIL = 11
End Enum

Public Enum IDSM_MAIL_SYSTEMS
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
End Enum



Private Sub Class_Initialize()
  m_UseDefaultProfile = False
  m_UseAddressBook = IDSM_AB_SEEK_NONE
  Set m_IDSMail = New IDSMailInterface32.Server
  'Set m_IDSAddressBook = New IDSMailInterface32.AddressBook
  m_IDSMail.objectKey = "EVALUATION"
  m_IDSMail.ErrorDisplay = False
  Call RegistrySettings(REGISTRY_KEY_READ)
    
  'Display MailSetup if debugging required
  If Not m_CanEmail Then
    m_IDSMail.ErrorDisplay = True
    Call frmMailSetup.Show(vbModal)
    'frmMailSetup.Start (Me)
  End If
End Sub

Private Sub Class_Terminate()
 Call RegistrySettings(REGISTRY_KEY_WRITE)
 Set m_IDSMail = Nothing
 Set m_IDSAddressBook = Nothing
End Sub

Public Sub ShowMAIL()
End Sub

Public Sub CloseMAIL()
End Sub

Private Property Get ILibraryVersion_Name() As String
  ILibraryVersion_Name = "MAIL library"
End Property

Private Property Get ILibraryVersion_Version() As String
  ILibraryVersion_Version = App.Major & "." & App.Minor & "." & App.Revision
End Property
Public Sub MailSend(ByVal msgSubject As String, ByVal msgBody As String, ToList As Variant, Optional CCList As Variant, Optional BCCList As Variant, Optional Attachments As Variant)
  
  On Error GoTo MailSend_Err
  
  If m_UseAddressBook Then
    Set m_IDSAddressBook = New IDSMailInterface32.AddressBook 'RK Should these be globals
    Set m_IDSAddressBook.MailObject = m_IDSMail
  End If
  
  Call SetMailSystem
  Call AddRecipients(ToList, CCList, BCCList)
  Call AddAttachments(Attachments)
  m_IDSMail.subject = msgSubject
  m_IDSMail.Message = msgBody
  m_IDSMail.Send

MailSend_End:
  Exit Sub

MailSend_Err:
  Err.Raise Err.Number, Err.Source, Err.Description
  Call RaiseMailErrors(Err)
  Resume MailSend_End
  Resume
End Sub

Private Function AddRecipients(ToList As Variant, CCList As Variant, BCCList As Variant) As Boolean
  Dim i As Long
  On Error GoTo AddRecipients_err
    
    If (Not IsEmpty(ToList)) And IsArray(ToList) Then
      For i = LBound(ToList) To UBound(ToList)
        m_IDSMail.AddRecipientTo LookUpName(CStr(ToList(i)))
      Next i
    End If
    If (Not IsEmpty(CCList)) And IsArray(CCList) Then
      For i = LBound(CCList) To UBound(CCList)
        m_IDSMail.AddRecipientCc LookUpName(CStr(CCList(i)))
      Next i
    End If
    If (Not IsEmpty(BCCList)) And IsArray(BCCList) Then
      For i = LBound(BCCList) To UBound(BCCList)
        m_IDSMail.AddRecipientBcc LookUpName(CStr(BCCList(i)))
      Next i
    End If
    AddRecipients = True

AddRecipients_end:
  Exit Function

AddRecipients_err:
  AddRecipients = False
  Err.Raise Err.Number, "AddRecipients", Err.Description
  Call RaiseMailErrors(Err)
End Function

Private Function SetMailSystem() As Boolean
  On Error GoTo SetMailSystem_Err
  'Interrogate registry or ini file
  m_IDSMail.SetNativeSystem
  
SetMailSystem_End:
  Exit Function
  
SetMailSystem_Err:
  SetMailSystem = False
  Err.Raise Err.Number, "SetMailSystem", Err.Description
End Function

Private Function AddAttachments(Attachments As Variant) As Boolean
  Dim i As Long
  On Error GoTo AddAttachments_err
  If (Not IsEmpty(Attachments)) And IsArray(Attachments) Then
    For i = LBound(Attachments) To UBound(Attachments)
      m_IDSMail.AddAttachment CStr(Attachments(i))
    Next i
  End If
  AddAttachments = True
 
AddAttachments_end:
  Exit Function
    
AddAttachments_err:
  AddAttachments = False
  Err.Raise Err.Number, "AddAttachments", Err.Description
End Function

Private Function SetDefaultProfile(ByRef IDSMail As IDSMailInterface32.Server) As Boolean
  On Error GoTo SetDefaultProfile_err
  Dim sDefaultProfile As String
  
  sDefaultProfile = IDSMail.DefaultProfile
  IDSMail.LoginName = sDefaultProfile
  IDSMail.Login
  SetDefaultProfile = True
 
SetDefaultProfile_end:
  Exit Function
    
SetDefaultProfile_err:
  SetDefaultProfile = False
  Err.Raise Err.Number, "SetDefaultProfile", Err.Description
End Function

Public Property Get UseDefaultProfile() As Boolean
  UseDefaultProfile = m_UseDefaultProfile
End Property

Public Property Let UseDefaultProfile(ByVal NewValue As Boolean)
  m_UseDefaultProfile = NewValue
End Property

Public Property Get UseAddressBook() As IDSM_AB_SEEK_TYPES
  UseAddressBook = m_UseAddressBook
End Property

Public Property Let UseAddressBook(ByVal NewValue As IDSM_AB_SEEK_TYPES)
  m_UseAddressBook = NewValue
End Property

Private Function LookUpName(sLookUpName As String) As String
  On Error GoTo LookupName_err
  Dim sDefaultProfile As String
  
  If m_UseAddressBook Then
    m_IDSMail.Login
    m_IDSAddressBook.OpenAddressBook
    m_IDSAddressBook.MoveFirst
    m_IDSAddressBook.SeekName m_UseAddressBook, LookUpName
  
    If m_IDSAddressBook.NoMatch = True Then
      MsgBox "Name not found." 'RK Debug/Raise custom error? or try name provided?
      LookUpName = sLookUpName
    Else
      MsgBox "Found:" + m_IDSAddressBook.Name
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
  Err.Raise Err.Number, "LookupName", Err.Description
  Call RaiseMailErrors(Err)
End Function


Public Sub RaiseMailErrors(MyError As ErrObject)
  Dim errDecoded_IDS As Long, errDecodedAPI_IDS As Long, errMessage_IDS As String
  On Error GoTo RaiseMailErrors_Err
  'RK how to raise errors??
  'Inital error
  Err.Raise Err.Number, Err.Source, Err.Description
  m_IDSMail.DecodeError MyError.Number, m_IDSMail.MailSystem, errDecoded_IDS, errDecodedAPI_IDS, errMessage_IDS
  'Decoded error
  Err.Raise CLng(errDecoded_IDS), "RaiseMailErrors", CStr(errMessage_IDS) 'RK Fix this
 
RaiseMailErrors_End:
  Exit Sub

RaiseMailErrors_Err:
  Err.Raise Err.Number, "RaiseMailErrors", Err.Description
  Resume RaiseMailErrors_End
  Resume
End Sub
Public Property Get MailApplication() As MAIL_APPLICATIONS
  MailApplication = m_MailApplication
End Property

Public Property Let MailApplication(ByVal NewValue As MAIL_APPLICATIONS)
  m_MailApplication = NewValue
End Property

Friend Property Get IDSMail() As IDSMailInterface32.Server 'RK Expose entire object?
  Set IDSMail = m_IDSMail
End Property

Public Property Let MailSystem(ByVal NewValue As IDSM_MAIL_SYSTEMS)
  m_IDSMail.MailSystem = NewValue
End Property
Public Property Get MailSystem() As IDSM_MAIL_SYSTEMS
  MailSystem = m_IDSMail.MailSystem
End Property
Public Property Let LoginName(ByVal NewValue As String)
  m_IDSMail.LoginName = NewValue
End Property
Public Property Get LoginName() As String
  LoginName = m_IDSMail.LoginName
End Property


Friend Sub RegistrySettings(Action As RegistryKeyAction)
  On Error GoTo RegistrySettings_Err
'RK conventions?
    If Action = REGISTRY_KEY_WRITE Then
      Call SaveSetting(S_REGISTRY_APPNAME, S_REGISTRY_SECTION, "MailApplication", m_MailApplication)
      Call SaveSetting(S_REGISTRY_APPNAME, S_REGISTRY_SECTION, "MailSystem", m_IDSMail.MailSystem)
      Call SaveSetting(S_REGISTRY_APPNAME, S_REGISTRY_SECTION, "LoginName", m_IDSMail.LoginName)
      Call SaveSetting(S_REGISTRY_APPNAME, S_REGISTRY_SECTION, "CanEmail", m_CanEmail)
    ElseIf Action = REGISTRY_KEY_READ Then
      m_MailApplication = GetSetting(S_REGISTRY_APPNAME, S_REGISTRY_SECTION, "MailApplication", MA_OTHER)
      m_IDSMail.MailSystem = GetSetting(S_REGISTRY_APPNAME, S_REGISTRY_SECTION, "MailSystem", IDSM_SYS_UNKNOWN)
      m_LoginName = GetSetting(S_REGISTRY_APPNAME, S_REGISTRY_SECTION, "LoginName", "")
      m_CanEmail = GetSetting(S_REGISTRY_APPNAME, S_REGISTRY_SECTION, "CanEmail", 0)
    End If

RegistrySettings_End:
  Exit Sub

RegistrySettings_Err:
  Err.Raise Err.Number, Err.Source, Err.Description
  Resume RegistrySettings_End
End Sub
