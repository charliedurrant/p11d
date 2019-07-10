VERSION 5.00
Object = "{F8D6DCD3-30A3-11D3-8C5F-0008C75A1F7A}#1.0#0"; "TCSSPLIT.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin TCSSPLIT.SPLIT SPLIT1 
      Height          =   3105
      Left            =   1350
      TabIndex        =   0
      Top             =   60
      Width           =   195
      _extentx        =   344
      _extenty        =   5477
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  Call Me.SPLIT1.Initialise(Me.hWnd, False)
End Sub
