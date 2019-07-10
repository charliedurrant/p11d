VERSION 5.00
Begin VB.Form frmMSGOKCancel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Message Box"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label lblMsg 
      AutoSize        =   -1  'True
      Caption         =   "fgdfgdfgdfgdfgdfgdgdfgdfgdfgdfg"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   4125
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMSGOKCancel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mRetVal As Boolean
Private mTimeOut As Long

Private Sub cmdCancel_Click()
  mRetVal = False
  Hide
End Sub

Private Sub cmdOK_Click()
  mRetVal = True
  Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF12 And Shift = False Then
    If IsRunningInIDEEx Then Call ShowDebugPopupex
  End If
End Sub

Private Sub Form_Load()
  mRetVal = False
  mTimeOut = -1
End Sub

Public Function displayMsg(InForm As Form, Message As String, Title As String, OKText As String, CancelText As String, Optional ByVal TimeOut As Long = -1, Optional ByVal AlreadyLoaded As Boolean = False) As Boolean
  Dim frm As Form
  Dim x0 As Single, y0 As Single
  Dim d As Single
  On Error GoTo displayMsg_Err
  
  mTimeOut = TimeOut
  With frmMSGOKCancel
    If mTimeOut < 0 Then
      .cmdOK.Enabled = True
      .cmdOK.Visible = True
    Else
      .cmdOK.Enabled = False
      .cmdOK.Visible = False
    End If
    If (mTimeOut >= 0) Or (Len(CancelText) = 0) Then
      .cmdCancel.Enabled = False
      .cmdCancel.Visible = False
    Else
      .cmdCancel.Caption = CancelText
      .cmdCancel.Enabled = True
    End If
    If Len(OKText) = 0 Then OKText = "Ok"
    .cmdOK.Caption = OKText
    
    .lblMsg.left = 250
    .lblMsg.Width = (.Width - 500)
    .lblMsg = Message
    If mTimeOut < 0 Then
      d = .lblMsg.top + .lblMsg.Height + .cmdOK.Height + 750
    Else
      d = .lblMsg.top + .lblMsg.Height + 750
    End If
    If AlreadyLoaded Then
      If d > .Height Then .Height = d
    Else
      .Height = d
    End If
    Set frm = InForm
    If frm Is Nothing Then
      If Not vbg Is Nothing Then
        If Not vbg.Screen.ActiveForm Is Nothing Then Set frm = vbg.Screen.ActiveForm
      End If
      If frm Is Nothing Then Set frm = VB.Screen.ActiveForm
    End If
    If frm Is Nothing Then
      y0 = (Screen.Height - .Height) / 2
      x0 = (Screen.Width - .Width) / 2
    Else
      y0 = frm.top + (frm.Height - .Height) / 2
      x0 = frm.left + (frm.Width - .Width) / 2
    End If
    If (x0 < 0) Or (y0 < 0) Then
      y0 = (Screen.Height - .Height) / 2
      x0 = (Screen.Width - .Width) / 2
    End If
    .top = y0
    .left = x0
    .cmdCancel.top = .lblMsg.top + .lblMsg.Height + 200
    .cmdOK.top = .lblMsg.top + .lblMsg.Height + 200
    If Not .cmdCancel.Enabled Then .cmdOK.left = (.Width - .cmdOK.Width) / 2
    .Caption = Title
    If isMDI_Minimized() Then
      If InForm.WindowState = vbMinimized Then Err.Raise ERR_DISPLAYMSG
    End If
    If mTimeOut < 0 Then
      Call SetCursorEx(vbArrow, "")
      Call .Show(vbModal)
      Call ClearCursorEx(False)
    Else
      Call .Show
      'If Not AlreadyLoaded Then Call .ZOrder
      Call .Refresh
      If mTimeOut < NO_TIMEOUT Then
        Call SleepW32(mTimeOut * 1000)
        Call .Hide
      End If
    End If
  End With
  displayMsg = mRetVal
    
displayMsg_End:
  Exit Function

displayMsg_Err:
  displayMsg = False
  Resume displayMsg_End
End Function

