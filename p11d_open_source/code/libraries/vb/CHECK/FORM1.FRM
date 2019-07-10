VERSION 5.00
Object = "*\ATCSCHECK.vbp"
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
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   2760
      Width           =   975
   End
   Begin TCSMultiCheck.tcsCheck tcsMultiStateCheckBox1 
      Height          =   195
      Left            =   1590
      TabIndex        =   0
      Top             =   300
      Width           =   210
      _ExtentX        =   344
      _ExtentY        =   344
      m_AllowClick    =   -1  'True
      ValidStates     =   10
      State           =   2
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Me.tcsMultiStateCheckBox1.ValidStates = CHECK_CROSS + CHECK_TICK
End Sub
