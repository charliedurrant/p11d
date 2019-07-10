VERSION 5.00
Object = "*\ATCSSplitter.vbp"
Begin VB.MDIForm MDI 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4620
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Align           =   4  'Align Right
      Height          =   4395
      Left            =   2625
      ScaleHeight     =   4335
      ScaleWidth      =   1935
      TabIndex        =   1
      Top             =   0
      Width           =   2000
      Begin TCSSplitter.TCSSPLIT TCSSPLIT2 
         Height          =   1455
         Left            =   960
         TabIndex        =   4
         Top             =   480
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   2566
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4395
      Left            =   0
      ScaleHeight     =   4395
      ScaleWidth      =   1275
      TabIndex        =   0
      Top             =   0
      Width           =   1275
      Begin TCSSplitter.TCSSPLIT TCSSPLIT1 
         Height          =   2415
         Left            =   1080
         TabIndex        =   3
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   4260
      End
      Begin VB.TextBox Text1 
         Height          =   4095
         Left            =   0
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   120
         Width           =   1095
      End
   End
End
Attribute VB_Name = "MDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
  Call TCSSPLIT1.Initialise(Me.hWnd, True)
  Call TCSSPLIT2.Initialise(Me.hWnd, True)
End Sub

Private Sub Picture1_Resize()
  Text1.Width = Picture1.Width - 100
End Sub
