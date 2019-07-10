VERSION 5.00
Begin VB.Form frmErr 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   2865
   ClientLeft      =   945
   ClientTop       =   1785
   ClientWidth     =   7815
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2865
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExtra 
      Caption         =   "&Retry"
      Height          =   375
      Left            =   6360
      TabIndex        =   17
      Top             =   1920
      Width           =   1380
   End
   Begin VB.CheckBox chkIgnoreError 
      Caption         =   "Do not show this error again"
      Height          =   375
      Left            =   4920
      TabIndex        =   16
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Frame fraApp 
      Height          =   1065
      Left            =   90
      TabIndex        =   8
      Top             =   1740
      Width           =   4620
      Begin VB.Label lblApplication 
         Height          =   675
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   4395
      End
   End
   Begin VB.Frame fraDetails 
      Caption         =   "Details"
      Height          =   2475
      Left            =   120
      TabIndex        =   6
      Top             =   2880
      Width           =   7680
      Begin VB.CommandButton cmdCopyError 
         Caption         =   "Copy to Clipboard"
         Height          =   420
         Left            =   3675
         TabIndex        =   19
         Top             =   1875
         Width           =   1560
      End
      Begin VB.ListBox ErrorSource 
         Height          =   645
         Left            =   120
         TabIndex        =   15
         Top             =   1650
         Width           =   3495
      End
      Begin VB.TextBox txtPath 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "Path"
         Top             =   960
         Width           =   5100
      End
      Begin VB.ListBox lstStack 
         Height          =   1815
         Left            =   5280
         TabIndex        =   11
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label lblErrType 
         Caption         =   "Error Type"
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   5100
      End
      Begin VB.Label lblStack 
         Caption         =   "Function call stack:"
         Height          =   255
         Left            =   5325
         TabIndex        =   12
         Top             =   225
         Width           =   1965
      End
      Begin VB.Label lblFunction 
         Caption         =   "Function name"
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   5100
      End
      Begin VB.Label lblExeName 
         Caption         =   "Executable  name"
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   5100
      End
   End
   Begin VB.CommandButton cmdDetails 
      Caption         =   "&Details >>"
      Height          =   375
      Left            =   4875
      TabIndex        =   4
      Top             =   2400
      Width           =   1380
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   6360
      TabIndex        =   3
      Top             =   2400
      Width           =   1380
   End
   Begin VB.Frame fraContact 
      Height          =   525
      Left            =   75
      TabIndex        =   2
      Top             =   1125
      Width           =   7680
      Begin VB.Label lblHelp 
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   7440
      End
   End
   Begin VB.Frame fraErr 
      Height          =   1080
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   7680
      Begin VB.PictureBox picErrMsg 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   1935
         ScaleHeight     =   510
         ScaleWidth      =   915
         TabIndex        =   18
         Top             =   360
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label lblErrMsg 
         AutoSize        =   -1  'True
         Height          =   675
         Left            =   135
         TabIndex        =   1
         Top             =   240
         Width           =   7425
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmErr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const ERRFRM_INCREASE As Long = 2650
Public OtherButton As Boolean
Public ErrorNumber As Long
Public Message As String
Private mClipMessage As String

Public Property Let ClipMessage(ByVal NewValue As String)
  mClipMessage = Trim$(CompressStringEx(NewValue, vbCrLf))
End Property

Private Sub cmdCopyError_Click()
  If OpenClipboard(0) Then
    Call EmptyClipboard
    Call SetAnyClipboardDataEx(vbCFText, mClipMessage)
    Call CloseClipboard
  End If
End Sub

Private Sub cmdDetails_Click()
  Static bDetails As Boolean

  If bDetails Then
    Me.Height = Me.Height - ERRFRM_INCREASE
    Me.cmdDetails.Caption = "&Details >>"
    bDetails = False
  Else
    Me.Height = Me.Height + ERRFRM_INCREASE
    Me.cmdDetails.Caption = "&Details <<"
    bDetails = True
  End If
End Sub

Private Sub cmdExtra_Click()
  OtherButton = True
  Me.Visible = False
End Sub

Private Sub cmdOK_Click()
  Me.Visible = False
End Sub

Private Sub Form_Activate()
  If mForceErrorTopMost Then
    Call SetWindowZOrderEx(Me.hWnd, HWND_TOPMOST)
  Else
    Call SetForegroundWindow(Me.hWnd)
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF12 And Shift = False Then
    If IsRunningInIDEEx Then Call ShowDebugPopupex
  End If
End Sub

Private Sub Form_Load()
  OtherButton = False
End Sub

Private Sub Form_LostFocus()
  On Error Resume Next
  Call Me.SetFocus
End Sub

Private Sub picErrMsg_Paint()
  Call DrawFormattedString(Message, Me.picErrMsg.hWnd, Me.picErrMsg.ScaleX(Me.picErrMsg.Width, Me.picErrMsg.ScaleMode, vbPixels))
End Sub
