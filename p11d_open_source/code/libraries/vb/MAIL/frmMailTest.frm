VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMailTest 
   Caption         =   "Mail Test"
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   4530
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCC 
      Height          =   375
      Left            =   1200
      TabIndex        =   12
      Top             =   600
      Width           =   3135
   End
   Begin VB.CheckBox chkExcEnvir 
      Caption         =   "Exclude environment info?"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   5700
      Width           =   2175
   End
   Begin VB.TextBox txtTo 
      Height          =   375
      Left            =   1200
      TabIndex        =   9
      Top             =   120
      Width           =   3135
   End
   Begin VB.TextBox txtEnvir 
      Enabled         =   0   'False
      Height          =   1815
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   3840
      Width           =   4215
   End
   Begin MSComctlLib.StatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   6645
      Width           =   4530
      _ExtentX        =   7990
      _ExtentY        =   661
      Style           =   1
      SimpleText      =   "Mail test not conducted"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox chkAttachDebug 
      Caption         =   "Attach debug file?"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   5760
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Top             =   6120
      Width           =   1095
   End
   Begin VB.TextBox txtMessage 
      Height          =   1815
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "frmTest.frx":0000
      Top             =   1560
      Width           =   4215
   End
   Begin VB.TextBox txtSubject 
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   1080
      Width           =   3135
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label lblCC 
      Caption         =   "Cc:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblTo 
      Caption         =   "To:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblEnvir 
      Caption         =   "Environment Info To Be Appended To Message:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3600
      Width           =   3495
   End
   Begin VB.Label lblSubject 
      Caption         =   "Subject:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   855
   End
End
Attribute VB_Name = "frmMailTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public m_Mail As Mail
Public m_DebugOutputPathFile As String
Private m_EnvironmentInfo As String
Private m_AbatecSubject As String

Private Sub chkAttachDebug_Click()
  Dim i As Long, Attrs As Long
  Dim sOutput As String
    If FileExists(m_DebugOutputPathFile) Then
      sb.SimpleText = "Mail debug output found"
    Else
      sb.SimpleText = "No mail debug output found"
      chkAttachDebug.Enabled = False
      chkAttachDebug.Value = False
    End If
End Sub

Private Sub chkExcEnvir_Click()
  If chkExcEnvir.Value Then
    txtEnvir.Text = ""
  Else
    txtEnvir.Text = m_EnvironmentInfo
  End If
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdSend_Click()
Dim MailAttachments As Variant
Dim sSubject As String
Dim sMessage As String
Dim vCClist As Variant
Dim vTolist As Variant
On Error GoTo OnError

  'Set flag "Success" to False
  m_Mail.Success = False
  Call m_Mail.ClearAllMessageItems
  
  'Split multiple names
  If InStr(1, txtTo.Text, ",", vbTextCompare) Then
    vTolist = Split(txtTo.Text, ",")
  Else
    vTolist = txtTo.Text
  End If
  If InStr(1, txtCC.Text, ",", vbTextCompare) Then
    vCClist = Split(txtCC.Text, ",")
  Else
    vCClist = txtCC.Text
  End If
  
  'Force message header for abatec
  If InStr(1, txtTo.Text, m_Mail.TestMailRecipient, vbTextCompare) Then
    sSubject = m_AbatecSubject
  Else
    sSubject = txtSubject.Text
  End If
  
  'AM Message will include environment if not specifically checked not to
  sMessage = txtMessage.Text
  If Not chkExcEnvir.Value Then sMessage = sMessage & vbCrLf & vbCrLf & txtEnvir
  If chkAttachDebug.Value Then m_Mail.AddAttachment m_DebugOutputPathFile
  Call m_Mail.MailSend(sSubject, sMessage, vTolist, vCClist)
  
  'Set flag "Success" to True
  m_Mail.Success = True
  sb.SimpleText = "Mail test sent"
  
  'Reset message
  m_Mail.NewMessage
  
  'Shutdown form
  Unload Me
OnEnd:
  Exit Sub

OnError:
  m_Mail.Success = False
  sb.SimpleText = "Mail test not successfully completed"
  Resume Next
End Sub

Private Sub Form_Load()
  sb.SimpleText = "Loading form.."
  
  'AM Following functions for environment info (from TCSCore)
  Call MathInit
  m_EnvironmentInfo = GetEnvironmentInfo(Me.m_Mail)
  txtTo.Text = m_Mail.TestMailRecipient
  m_AbatecSubject = "ATCMail test:" & Now() 'RK include application name here
  txtSubject.Text = m_AbatecSubject
  txtEnvir.Text = m_EnvironmentInfo
  
  'Is Debug application present?
  If Len(m_DebugOutputPathFile) > 0 Then
    chkAttachDebug.Enabled = True
    chkAttachDebug.Value = vbChecked
  Else
    chkAttachDebug.Enabled = False
    chkAttachDebug.Value = vbUnchecked
  End If
End Sub

