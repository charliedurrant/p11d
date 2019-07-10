VERSION 5.00
Begin VB.UserControl tcsMultiStateCheckBox 
   ClientHeight    =   900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   300
   ScaleHeight     =   900
   ScaleWidth      =   300
   Begin VB.PictureBox chkBox 
      Height          =   195
      Index           =   0
      Left            =   30
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   2
      Top             =   30
      Width           =   195
   End
   Begin VB.PictureBox chkBox 
      Height          =   195
      Index           =   2
      Left            =   30
      Picture         =   "tcsCheck.ctx":0000
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   1
      Top             =   555
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox chkBox 
      Height          =   195
      Index           =   1
      Left            =   30
      Picture         =   "tcsCheck.ctx":00CA
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   0
      Top             =   300
      Visible         =   0   'False
      Width           =   195
   End
End
Attribute VB_Name = "tcsMultiStateCheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Enum CHECK_STATE
  CHECK_EMPTY = 0
  CHECK_TICK = 1
  CHECK_CROSS = 2
End Enum

Dim MAXCHECK As Long
Dim MINCHECK As Long
Private m_State As Long
Private m_AllowClick As Boolean
Option Explicit
Private Sub ChangeCheck(curState As Long)
  Dim newState As Long
  Dim i As Long
  
  If curState = MAXCHECK Then
    curState = MINCHECK
  Else
    curState = curState + 1
  End If
  
  For i = MINCHECK To MAXCHECK
   If i = curState Then
     chkBox(i).Visible = True
   Else
     chkBox(i).Visible = False
   End If
  Next i
End Sub

Private Sub chkBox_Click(Index As Integer)
  If m_AllowClick Then Call ChangeCheck(CLng(Index))
End Sub

Private Sub UserControl_Initialize()
  Dim i As Long
  
  MAXCHECK = chkBox.UBound
  MINCHECK = chkBox.LBound
  For i = MINCHECK To MAXCHECK
    chkBox(i).Left = 0
    chkBox(i).Top = 0
    chkBox(i).Visible = False
    chkBox(i).Height = 195
    chkBox(i).Width = 195
  Next i
  
  chkBox(MINCHECK).Visible = True
End Sub

Public Property Let State(StateType As CHECK_STATE)
  m_State = StateType
  ChangeCheck (CLng(m_State - 1))
End Property

Public Property Get State() As CHECK_STATE
  State = m_State
End Property

Public Property Let AllowClick(AlCl As Boolean)
  m_AllowClick = AlCl
End Property

Public Property Get AllowClick() As Boolean
  AllowClick = m_AllowClick
End Property
