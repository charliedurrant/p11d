VERSION 5.00
Begin VB.Form frmMAPIMail 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MAPI Mail"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtTo 
      Height          =   375
      Left            =   2730
      TabIndex        =   7
      Top             =   105
      Width           =   3255
   End
   Begin VB.TextBox txtSubjectFilter 
      Height          =   375
      Left            =   2745
      TabIndex        =   5
      Top             =   1020
      Width           =   3255
   End
   Begin VB.TextBox txtFromFilter 
      Height          =   375
      Left            =   2745
      TabIndex        =   2
      Top             =   540
      Width           =   3255
   End
   Begin VB.CommandButton cmdRetriveMail 
      Caption         =   "Retrieve"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1005
      Width           =   1215
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send Mail"
      Height          =   375
      Left            =   90
      TabIndex        =   0
      Top             =   555
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "To:"
      Height          =   255
      Left            =   1380
      TabIndex        =   8
      Top             =   165
      Width           =   1095
   End
   Begin VB.Label lblResults 
      Height          =   1695
      Left            =   345
      TabIndex        =   6
      Top             =   1620
      Width           =   5655
   End
   Begin VB.Label lblSubjectFilter 
      Caption         =   "Subject Filter:"
      Height          =   255
      Left            =   1425
      TabIndex        =   4
      Top             =   1140
      Width           =   1095
   End
   Begin VB.Label lblFromFilter 
      Caption         =   "From Filter:"
      Height          =   255
      Left            =   1425
      TabIndex        =   3
      Top             =   660
      Width           =   1095
   End
End
Attribute VB_Name = "frmMAPIMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Implements IMailMessage
'Implements IBaseNotify
'Private Const USE_NOTES As Boolean = True
'
'Private Sub cmdRetriveMail_Click()
'  If gMAPI Is Nothing Then Set gMAPI = New Mail
'  gMAPI.UseMAPIAPI = True
'  Call gMAPI.ForEachMessage(Me, True, Trim$(Me.txtFromFilter), Trim$(Me.txtSubjectFilter), Me)
'End Sub
'
'Private Sub cmdSend_Click()
'  If gMAPI Is Nothing Then Set gMAPI = New Mail
'  If USE_NOTES Then
'    gMAPI.UseNotesAPI = True
'    gMAPI.NotesiniFile = "G:\NOTES\NOTES.INI"
'    Call gMAPI.MailSend("Test Email", "Hello World", Me.txtTo.Text)
'  Else
'    Call gMAPI.ShowMAIL
'  End If
'End Sub
'
'Private Sub IBaseNotify_Notify(ByVal Current As Long, ByVal Max As Long, ByVal Message As String)
'  lblResults = "Message " & Current & " of " & Max
'End Sub
'
'Private Function IMailMessage_Message(ByVal msgFrom As String, ByVal msgSubject As String, ByVal msgBody As String, Attachments As TCSMAIL.MailAttachments) As TCSMAIL.MessageAction
'  Dim retval As Long
'  retval = MsgBox(msgFrom & vbCrLf & msgSubject, vbOKCancel + vbDefaultButton2, "Do you want to delete the following message?")
'  If retval = vbOK Then IMailMessage_Message = MESSAGE_DELETE
'End Function
