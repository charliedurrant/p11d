Attribute VB_Name = "MailSetup"
Option Explicit

Public Sub QueryMailApplication(ByRef IDSMail As IDSMailInterface32.Server, ByRef MyMailSystem As IDSM_MAIL_SYSTEM, ByRef MyMailApplication As MAIL_APPLICATION)
  Dim MyMailApplicationQuery As Long
  On Error GoTo QueryMailApplication_err
    
    'Determine notes applications (special case of method)
    If MyMailSystem = IDSM_SYS_VIM Then
      MyMailApplicationQuery = IDSMail.QueryMailSystem(IDSM_QUERY_NOTES_OR_CCMAIL)
      Select Case MyMailApplicationQuery
        Case IDSM_ENUM_CCMAIL
          MyMailApplication = MA_LOTUS_CC_MAIL
        Case IDSM_ENUM_NOTES
          MyMailApplication = MA_LOTUS_NOTES_VIM
        Case Else
          MyMailApplication = MA_OTHER
      End Select
    Else
      MyMailApplication = MyMailApplication 'No change
    End If
    
QueryMailApplication_end:
  Exit Sub
  
QueryMailApplication_err:
  MyMailApplication = MA_OTHER
  Resume QueryMailApplication_end
End Sub

Public Sub QueryDefaultProfile(ByRef IDSMail As IDSMailInterface32.Server, LoginName As String)
  On Error GoTo QueryDefaultProfile_err
      
  LoginName = IDSMail.DefaultProfile
  IDSMail.LoginName = LoginName
    
QueryDefaultProfile_end:
  Exit Sub
  
QueryDefaultProfile_err:
  IDSMail.LoginName = ""
  Err.Raise Err.Number, "QueryDefaultProfile", Err.Description
  Resume QueryDefaultProfile_end
End Sub
Public Sub GetDebugPaths(sGetDebugAppPathFile As String, sDebugOutputPathFile As String)
  Dim sSysDirectory As String
  On Error GoTo GetDebugPaths_err
      
  sSysDirectory = GetSysDirectory()
  sGetDebugAppPathFile = sSysDirectory & "\" & S_IDSMAIL_DEBUGAPP_DIR & "\" & S_IDSMAIL_DEBUGAPP_FILE
  sDebugOutputPathFile = sSysDirectory & "\" & S_IDSMAIL_DEBUGAPP_DIR & "\" & S_IDSMAIL_DEBUGAPP_OUTPUT_FILE
    
GetDebugPaths_end:
  Exit Sub
  
GetDebugPaths_err:
  Err.Raise Err.Number, "GetDebugPaths", Err.Description
  Resume GetDebugPaths_end
End Sub

'AM Core function
Function GetSysDirectory() As String
  Dim sRes As String
  Dim retval As Long
  
  On Error GoTo GetSysDirectory_err
  sRes = String$(TCSBUFSIZ, 0)
  retval = GetSystemDirectory(sRes, TCSBUFSIZ)
  If retval = 0 Then
    sRes = WINDIR
  Else
    sRes = Left$(sRes, retval)
  End If
  
GetSysDirectory_end:
  GetSysDirectory = sRes
  Exit Function
  
GetSysDirectory_err:
  Resume GetSysDirectory_end
End Function

