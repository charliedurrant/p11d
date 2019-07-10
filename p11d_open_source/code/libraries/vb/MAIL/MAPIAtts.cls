VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MailAttachments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mCol As Collection

Private Sub Class_Initialize()
  Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
  Set mCol = Nothing
End Sub

Public Function Add(ByVal DisplayName As String, ByVal FileName As String, Optional ByVal MAILType As MAIL_ATTACHMENT_DATATYPE = MAIL_ATT_DATA) As MailAttachment
  Dim att As MailAttachment
  
  If Not FileExists(FileName) Then Err.Raise ERR_MAIL_ATT_ADD, "AddAttachment", "file not found " & FileName
  Set att = New MailAttachment
  att.DisplayName = DisplayName
  att.FileName = FileName
  att.MAILType = MAILType
  Call mCol.Add(att, att.DisplayName)
  Set Add = att
End Function

Public Sub Remove(Index As Variant)
  Call mCol.Remove(Index)
End Sub

Public Sub RemoveAll()
  Set mCol = Nothing
  Set mCol = New Collection
End Sub

Public Function Item(Index As Variant) As MailAttachment
  Set Item = mCol(Index)
End Function

Public Property Get Count() As Long
  Count = mCol.Count
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mCol.[_NewEnum]
End Function

