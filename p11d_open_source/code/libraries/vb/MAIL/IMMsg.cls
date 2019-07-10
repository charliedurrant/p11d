VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IMailMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum MessageAction
  MESSAGE_LEAVE = 0
  MESSAGE_DELETE
End Enum

'apf Byval att
Public Function Message(ByVal msgFrom As String, ByVal msgSubject As String, ByVal msgBody As String, Attachments As MailAttachments) As MessageAction
  
End Function
