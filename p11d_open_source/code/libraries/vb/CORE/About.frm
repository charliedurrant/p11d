VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   1680
      Width           =   1260
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblCopyright 
      Caption         =   "Copyright"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   5175
   End
   Begin VB.Label lblAppVersion 
      Alignment       =   2  'Center
      Caption         =   "Application Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   360
      Width           =   4695
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   2280
      TabIndex        =   5
      Top             =   1320
      Width           =   3015
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   2040
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   5235
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblContact 
      AutoSize        =   -1  'True
      Caption         =   "Contact Information"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   3885
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblApplication 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Application name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   375
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF12 And Shift = False Then
    If IsRunningInIDEEx Then
      Me.Visible = False
      Call ShowDebugPopupex
    End If
  End If
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

