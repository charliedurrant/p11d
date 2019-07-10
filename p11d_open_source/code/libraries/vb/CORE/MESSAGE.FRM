VERSION 5.00
Begin VB.Form frmMessage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Multi button dialog"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6675
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   6675
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fmeMessage 
      Height          =   2415
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5175
      Begin VB.Label lblMessage 
         Caption         =   "Your message should be here."
         Height          =   1935
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4815
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "&OK"
      Height          =   375
      Index           =   32767
      Left            =   5280
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lClicked As Long
Private Sub cmdArray_Click(Index As Integer)
  lClicked = Index + 1
  Unload Me
End Sub
