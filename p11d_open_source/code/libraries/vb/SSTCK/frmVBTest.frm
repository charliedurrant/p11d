VERSION 5.00
Object = "*\AXIDBSTAT.vbp"
Object = "{770120E1-171A-436F-A3E0-4D51C1DCE486}#1.0#0"; "atc2stat.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   ScaleHeight     =   4320
   ScaleWidth      =   4920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Link to stat"
      Height          =   555
      Left            =   2850
      TabIndex        =   4
      Top             =   1290
      Width           =   1395
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Pop"
      Height          =   495
      Left            =   1530
      TabIndex        =   3
      Top             =   330
      Width           =   1275
   End
   Begin VB.CommandButton cmdPush 
      Caption         =   "Push"
      Height          =   525
      Left            =   90
      TabIndex        =   2
      Top             =   330
      Width           =   1185
   End
   Begin atc2Stat.TCSStatus stat 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   3945
      Width           =   4920
      _ExtentX        =   8678
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XIDBSTAT.IDBSTAT IDBSTAT1 
      Height          =   495
      Left            =   630
      TabIndex        =   0
      Top             =   1470
      Width           =   585
      _extentx        =   1032
      _extenty        =   873
   End
   Begin VB.Label lbl 
      Caption         =   "Label1"
      Height          =   465
      Left            =   240
      TabIndex        =   5
      Top             =   2220
      Width           =   3885
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Count As Long
Private Sub cmdPush_Click()
  Call Me.IDBSTAT1.PushStatus(STATUSID_1, "fred asd a", PI_HOUR_GLASS, vbHourglass)
  m_Count = m_Count + 1
  Call Me.IDBSTAT1.PushStatus(STATUSID_2, "fredasdasdsad asd a", PI_LIGHTENING Or PI_FLASH)
  m_Count = m_Count + 1
  lbl = "Count " & m_Count
End Sub

Private Sub Command1_Click()
  Call IDBSTAT1.PopStatus
  m_Count = m_Count - 1
  lbl = "Count " & m_Count
End Sub

Private Sub Command2_Click()
  Set IDBSTAT1.stat = Me.stat
End Sub

Private Sub Form_Load()
  Call Me.IDBSTAT1.DefaultStatus(STATUSID_1, "fred", PI_LIGHTENING)
  Call Me.IDBSTAT1.DefaultStatus(STATUSID_2, "fred", PI_INFO)
End Sub
