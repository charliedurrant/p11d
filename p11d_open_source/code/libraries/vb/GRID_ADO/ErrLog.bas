Attribute VB_Name = "ErrorLogging"
Option Explicit

#If DEBUGVER Then
Private m_InitCount As Long
Private m_file As Integer
Private m_filename As String

Public Sub InitAutoLog()
  On Error GoTo InitAutoLog_err
  m_InitCount = m_InitCount + 1
  If m_file = 0 Then
    m_file = FreeFile
    m_filename = AppPath & "\AUTO.LOG"
    Open m_filename For Append Lock Write As m_file
  End If
  
InitAutoLog_end:
  Exit Sub
    
InitAutoLog_err:
  Call CloseAutoLog
  Resume InitAutoLog_end
End Sub

Public Sub CloseAutoLog()
  On Error Resume Next
  m_InitCount = m_InitCount - 1
  If m_InitCount = 0 Then
    If m_file <> 0 Then Close #m_file
    m_file = 0
  End If
End Sub

Public Sub AutoLog(ByVal msg As String)
  If m_file <> 0 Then Print #m_file, GetNetUser(False) & vbTab & Format$(Now, "hh:mm dd/mm/yy") & vbTab & msg
End Sub
#Else
Public Sub CloseAutoLog()
End Sub

Public Sub InitAutoLog()
End Sub

Public Sub AutoLog(ByVal msg As String)
End Sub
#End If
