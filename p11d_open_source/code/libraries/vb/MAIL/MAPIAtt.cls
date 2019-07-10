VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MailAttachment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum MAIL_ATTACHMENT_DATATYPE
  MAIL_ATT_DATA = mapData
  MAIL_ATT_EMBEDDED_OLE = mapEOLE
  MAIL_ATT_STATIC_OLE = mapSOLE
End Enum

Public DisplayName As String
Public FileName As String
Public MAILType As MAIL_ATTACHMENT_DATATYPE

Private Sub Class_Initialize()
  MAILType = MAIL_ATT_DATA
End Sub
