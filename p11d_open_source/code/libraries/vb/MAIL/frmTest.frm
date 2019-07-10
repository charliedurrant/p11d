VERSION 5.00
Begin VB.Form frmTest 
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   Picture         =   "frmTest.frx":0000
   ScaleHeight     =   8565
   ScaleWidth      =   4425
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTestRecipient 
      Height          =   375
      Left            =   1440
      TabIndex        =   23
      Text            =   "<USE DEFAULT>"
      Top             =   7920
      Width           =   1575
   End
   Begin VB.CommandButton cmdInstantiate 
      Caption         =   "Instantiate Mail Object"
      Height          =   375
      Left            =   2520
      TabIndex        =   21
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Frame fraTest 
      Caption         =   "Test options"
      Height          =   2895
      Left            =   120
      TabIndex        =   14
      Top             =   5520
      Width           =   4215
      Begin VB.CheckBox chkVariant 
         Caption         =   "Multiple names  (SPLIT/send name as variant)"
         Height          =   495
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   2415
      End
      Begin VB.CheckBox chkAutodetect 
         Caption         =   "Autodetect?"
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox chkClearMessageItems 
         Caption         =   "Clear Message items"
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   2055
      End
      Begin VB.CheckBox chkNewMail 
         Caption         =   "New Mail object each send"
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   2655
      End
      Begin VB.CheckBox chkClearRegistry 
         Caption         =   "Clear out registry entries"
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   2415
      End
      Begin VB.CheckBox chkNewMessage 
         Caption         =   "New Message?"
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label lblTestRecipient 
         Caption         =   "Test Recipient:"
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   2520
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdDestroyMail 
      Caption         =   "Destroy Mail Object"
      Height          =   375
      Left            =   2520
      TabIndex        =   13
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton cmdSetup 
      Caption         =   "Display S&etup"
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Frame fraUseAddressBook 
      Caption         =   "Use Address Book"
      Height          =   1335
      Left            =   120
      TabIndex        =   8
      Top             =   3600
      Width           =   2055
      Begin VB.OptionButton optAddressBook 
         Caption         =   "Partial Match"
         Height          =   435
         Index           =   2
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   1455
      End
      Begin VB.OptionButton optAddressBook 
         Caption         =   "Full match"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton optAddressBook 
         Caption         =   "Don't use"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.TextBox txtAttach 
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      Text            =   "C:\temp\mailtest.txt, C:\temp\mailtest2.txt, C:\temp\mailtest3.txt"
      Top             =   1320
      Width           =   2775
   End
   Begin VB.TextBox txtMessage 
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Text            =   "Enter message here"
      Top             =   1920
      Width           =   4215
   End
   Begin VB.TextBox txtSubject 
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Text            =   "Enter subject here"
      Top             =   720
      Width           =   3135
   End
   Begin VB.TextBox txtTo 
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Text            =   "RIchard kemp"
      Top             =   120
      Width           =   3135
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label lblAttach 
      Caption         =   "ATTACHMENT:"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblSubject 
      Caption         =   "SUBJECT:"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lblTo 
      Caption         =   "TO:"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1
Private TestMail As Mail
Const S_REGISTRY_APPNAME As String = "ATCMAIL"
Const S_REGISTRY_SECTION As String = "Settings"

Private Sub cmdDestroyMail_Click()
  Set TestMail = Nothing
End Sub

Private Sub cmdInstantiate_Click()
   InstantiateMail
End Sub

Private Sub cmdSend_Click()
Dim MailAttachments As Variant, i As Long
Dim ToList As Variant, Attachments As Object, AttachArray As Variant
On Error GoTo OnError


If chkVariant Then
  ToList = Split(txtTo.Text, ",")
Else
  ToList = txtTo.Text
End If

'Use address book
For i = optAddressBook.LBound To optAddressBook.UBound
  If optAddressBook(i).Value Then TestMail.UseAddressBook = i
Next i


AttachArray = Split(txtAttach.Text, ",")
Call TestMail.MailSend(txtSubject.Text, txtMessage.Text, ToList, , , AttachArray)

Exit Sub

OnEnd:
  Exit Sub

OnError:
  MsgBox Err.Description, vbCritical, Err.Number
  Resume OnEnd
  Resume
End Sub

Private Sub cmdSetup_Click()
'Mail recipient
If StrComp(txtTestRecipient.Text, "<USE DEFAULT>", vbTextCompare) Then
  TestMail.TestMailRecipient = txtTestRecipient.Text
End If
 
TestMail.ShowOptions

End Sub



Private Sub Form_Load()
  Set TestMail = New Mail
End Sub


Public Function FileExists(sFname As String) As Boolean
  Dim Attrs As Long
  On Error GoTo fileexists_err
  
  FileExists = False
  Attrs = GetAttr(sFname)
  If (Attrs And vbReadOnly) = 0 Then
      FileExists = True
  Else
      FileExists = False
  End If

fileexists_end:
  Exit Function
  
fileexists_err:
  FileExists = False
  Resume fileexists_end
End Function


Public Sub ClearRegEntries()
      Call SaveSetting(S_REGISTRY_APPNAME, S_REGISTRY_SECTION, "MailApplication", MA_OTHER)
      Call SaveSetting(S_REGISTRY_APPNAME, S_REGISTRY_SECTION, "MailSystem", IDSM_SYS_UNKNOWN)
      Call SaveSetting(S_REGISTRY_APPNAME, S_REGISTRY_SECTION, "LoginName", "")
      Call SaveSetting(S_REGISTRY_APPNAME, S_REGISTRY_SECTION, "UseAddressBook", IDSM_AB_SEEK_NONE)
      Call SaveSetting(S_REGISTRY_APPNAME, S_REGISTRY_SECTION, "Success", False)
End Sub
Public Sub InstantiateMail()
 If chkClearRegistry Then ClearRegEntries
 If TestMail Is Nothing Or chkNewMail Then Set TestMail = New Mail
 Set TestMail.OwnerForm = Me
 If chkClearMessageItems Then TestMail.ClearAllMessageItems
 If chkAutodetect Then TestMail.AutoDetect = True
 If chkClearMessageItems Then TestMail.ClearAllMessageItems
 If chkNewMessage Then TestMail.NewMessage
End Sub

