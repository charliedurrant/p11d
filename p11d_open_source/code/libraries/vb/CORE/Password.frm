VERSION 5.00
Begin VB.Form frmPassw 
   Caption         =   "QUERY_PASSWORD"
   ClientHeight    =   1935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1935
   ScaleWidth      =   5910
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   1500
      Width           =   1185
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   1500
      Width           =   1185
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   5160
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   810
      Width           =   600
   End
   Begin VB.Label lblContact 
      Caption         =   "Please contact TCS on 0171 438 3669"
      Height          =   300
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   4920
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblPrompt 
      Caption         =   "Warning - You are about to enter a system function."
      Height          =   555
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5700
   End
   Begin VB.Label lblInfoDate 
      Caption         =   "Please enter the password for today."
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   4920
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmPassw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mPasswordOK

Public Property Get PasswordOk() As Boolean
  PasswordOk = mPasswordOK
End Property
Public Property Let PasswordOk(bPasswordOK As Boolean)
  PasswordOk = mPasswordOK
End Property

Private Sub cmdOK_Click()
  mPasswordOK = True
  frmPassw.Visible = False
End Sub
Private Sub cmdCancel_Click()
  mPasswordOK = False
  frmPassw.Visible = False
End Sub

