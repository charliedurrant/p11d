VERSION 5.00
Begin VB.Form frmLDAPMon 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "LDAP Monitor"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   5355
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMon 
      Caption         =   "Start Monitoring"
      Height          =   285
      Left            =   3105
      TabIndex        =   8
      Top             =   1125
      Width           =   2175
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1215
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   1125
      Width           =   1815
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   1215
      TabIndex        =   5
      Top             =   765
      Width           =   1815
   End
   Begin VB.TextBox txtTimeout 
      Height          =   285
      Left            =   1215
      TabIndex        =   3
      Text            =   "5000"
      Top             =   405
      Width           =   1815
   End
   Begin VB.TextBox txtDefault 
      Height          =   285
      Left            =   1215
      TabIndex        =   0
      Text            =   $"frmLDAPMon.frx":0000
      Top             =   45
      Width           =   4065
   End
   Begin VB.Label lblDebugContext 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   90
      TabIndex        =   9
      Top             =   1530
      Width           =   5145
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Password"
      Height          =   240
      Left            =   0
      TabIndex        =   6
      Top             =   1170
      Width           =   1140
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Username"
      Height          =   240
      Left            =   0
      TabIndex        =   4
      Top             =   810
      Width           =   1140
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Timeout"
      Height          =   240
      Left            =   0
      TabIndex        =   2
      Top             =   450
      Width           =   1140
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Default Context"
      Height          =   240
      Left            =   0
      TabIndex        =   1
      Top             =   90
      Width           =   1140
   End
End
Attribute VB_Name = "frmLDAPMon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MonProc As Check

Private Sub cmdMon_Click()
  
  If MonProc Is Nothing Then
    Me.txtDefault.Enabled = False
    Me.txtPassword.Enabled = False
    Me.txtTimeout.Enabled = False
    Me.txtUsername.Enabled = False
    cmdMon.Caption = "Stop Monitor"
    Set MonProc = New Check
    MonProc.Username = txtUsername
    MonProc.Password = txtPassword
    MonProc.SvrContext = txtDefault
    Set MonProc.Monfrm = Me
    Call StartTimer(txtTimeout, MonProc)
  Else
    Me.txtDefault.Enabled = True
    Me.txtPassword.Enabled = True
    Me.txtTimeout.Enabled = True
    Me.txtUsername.Enabled = True
    cmdMon.Caption = "Start Monitor"
    Call StopTimer(MonProc)
    Set MonProc.Monfrm = Nothing
    Set MonProc = Nothing
  End If
End Sub

Private Sub lblDebugContext_Change()
  Dim newValue As String
  
  newValue =
End Sub

Private Sub lblDebugContext_Click()

End Sub
