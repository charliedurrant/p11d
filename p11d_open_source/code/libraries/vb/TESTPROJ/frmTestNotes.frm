VERSION 5.00
Begin VB.Form frmTestNotes 
   Caption         =   "Test Lotus Notes ( TCSNOTES )"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6615
   ScaleWidth      =   8685
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdKill 
      Caption         =   "Close Notes"
      Height          =   495
      Left            =   1920
      TabIndex        =   10
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox txtSendTo 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   1680
      Width           =   2775
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send mail"
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox txtPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Send to"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblINIFile 
      Caption         =   "Notes not initialised"
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label lblError 
      Caption         =   "Results:"
      Height          =   1335
      Left            =   0
      TabIndex        =   7
      Top             =   2880
      Width           =   4815
   End
   Begin VB.Label lblUserName 
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "password"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "username"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "ini file"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmTestNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mNotes As NotesDB

Private Sub OpenMailBox(ByVal NotesiniFile As String)
  Dim mailBox As String, mailServer As String, mailFile As String
  
  On Error GoTo OpenMailBox_err
  mailServer = GetIniEntry("NOTES", "MAILSERVER", "", NotesiniFile)
  If Len(mailServer) = 0 Then Err.Raise ERR_CORESETUP, "OpenMailBox", "Unable to find MAILSERVER entry in Notes ini file " & NotesiniFile
  mailFile = GetIniEntry("NOTES", "MAILFILE", "", NotesiniFile)
  If Len(mailFile) = 0 Then Err.Raise ERR_CORESETUP, "OpenMailBox", "Unable to find MAILFILE entry in Notes ini file " & NotesiniFile
  mailBox = mailServer & "!!" & mailFile
  Call mNotes.OpenDB(mailBox)
  Exit Sub
  
OpenMailBox_err:
  Err.Raise Err.Number, ErrorSource(Err, "OpenMailBox"), Err.Description
End Sub


Private Sub SetupNotes()
  On Error GoTo SetupNotes_err
  Me.lblError = ""
  If mNotes Is Nothing Then
    Set mNotes = New NotesDB
    Me.txtPassword = Trim$(Me.txtPassword)
    If Len(Me.txtPassword) > 0 Then Call mNotes.SetPassword(Me.txtPassword)
    Me.lblINIFile = mNotes.GetNotesIni
    Me.lblUserName = mNotes.GetUserName
    Call OpenMailBox(mNotes.GetNotesIni)
  End If
  Exit Sub
  
SetupNotes_err:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Private Sub cmdKill_Click()
  Set mNotes = Nothing
End Sub

Private Sub cmdSend_Click()
  Dim ne As NoteEdit
  On Error GoTo SendError_err
  Call SetupNotes
  Me.txtSendTo = Trim$(Me.txtSendTo)
  If Len(Me.txtSendTo) = 0 Then Err.Raise ERR_CORESETUP, "SendTo", "No SendTo specified"
  Set ne = mNotes.CreateMail()
  Call ne.SetField("SendTo", CStr(Me.txtSendTo))
  Call ne.SetField("Recipients", CStr(Me.txtSendTo))
  Call ne.SetField("Subject", "Test mail")
  Call ne.SetField("PostedDate", FormatDateTime(Now, vbGeneralDate))
  Call ne.SetField("LastName", mNotes.GetUserName)
  Call ne.SetField("Form", "Memo")
  Call ne.SetField("From", mNotes.GetUserName)
  Call ne.SendMail
  Call ne.CloseNote(True)
  Me.lblError = "Mail send successful"
  Exit Sub
  
SendError_err:
  Me.lblError = "Error: " & Err.Number & vbCrLf & Err.Description & vbCrLf & "Source: " & Err.Source
End Sub

