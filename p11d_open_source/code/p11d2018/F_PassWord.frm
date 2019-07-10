VERSION 5.00
Begin VB.Form F_PassWord 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   3015
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtConfirm 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   45
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   855
      Width           =   2895
   End
   Begin VB.TextBox txtPassWord 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   45
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   225
      Width           =   2895
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1890
      TabIndex        =   3
      Top             =   1305
      Width           =   1050
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   810
      TabIndex        =   2
      Top             =   1305
      Width           =   1005
   End
   Begin VB.Label lblEnter 
      Caption         =   "Please enter password"
      Height          =   195
      Left            =   45
      TabIndex        =   5
      Top             =   45
      Width           =   1725
   End
   Begin VB.Label lblConfirm 
      Caption         =   "Confirm password"
      Height          =   195
      Left            =   45
      TabIndex        =   4
      Top             =   675
      Width           =   1905
   End
End
Attribute VB_Name = "F_PassWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_PWM As PW_MODE
Private m_ok As Boolean
Private m_benEy As IBenefitClass
Public Function PassWord(ey As Employer, ByVal PWM As PW_MODE) As Boolean
  Dim sngGapButton As Single, sngGapForm As Single, sng As Single
  
  On Error GoTo PassWord_ERR
  
  
  If (ey Is Nothing) And (PWM = PWM_CHECK_CURRENT) Then Call Err.Raise(ERR_IS_NOTHING, "PassWord", "Checking password and employer is nothing.")
  
  Set m_benEy = ey

  Me.Caption = Me.Caption & " for: " & m_benEy.value(employer_Name_db)
  sngGapButton = cmdOK.Top - (txtConfirm.Height + txtConfirm.Top)
  sngGapForm = Me.ScaleHeight - (cmdOK.Top + cmdOK.Height)
  m_PWM = PWM
  
  Select Case PWM
    Case PW_MODE.PWM_CHECK_CURRENT
      txtConfirm.Visible = False
      lblConfirm.Visible = False
      cmdOK.Top = txtPassWord.Top + txtPassWord.Height + sngGapButton
      cmdCancel.Top = cmdOK.Top
      Me.Height = cmdOK.Top + cmdOK.Height + sngGapForm + (Me.Height - Me.ScaleHeight)
    Case PW_MODE.PWM_SET
      'basically display the crap
    Case Else
      Call ECASE("Invalid password mode, mode = " & PWM & " in function PassWord.")
      GoTo PassWord_END
  End Select
  
  Call SetCursor(vbArrow)
  
'  Me.Show vbModal
  Call p11d32.Help.ShowForm(Me, vbModal)
  PassWord = m_ok
  
PassWord_END:
  Call ClearCursor
  Exit Function
PassWord_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "PassWord", "Password", "Error analysing password.")
End Function
Private Sub SetTxtFocus(txt As TextBox)
End Sub

Private Function CheckInput() As Boolean
  On Error GoTo CheckInput_ERR
  Dim sPassWord As String
  Dim sConfirm As String
  Dim lErrNumber As Long
  Call xSet("CheckInput")
  
  sPassWord = Trim(txtPassWord.Text)
  sConfirm = Trim(txtConfirm.Text)
  
  Select Case m_PWM
    Case PWM_CHECK_CURRENT
      If (StrComp(m_benEy.value(employer_PassWord_db), sPassWord, vbTextCompare) <> 0) Then
        If (StrComp(S_PASSWORD_OVERIDE, sPassWord, vbTextCompare) <> 0) Then
          Call Err.Raise(ERR_PASSWORD, "CheckInput", "The password is incorrect, please try again.")
        End If
      End If
      CheckInput = True
    Case PWM_SET
      If StrComp(sPassWord, sConfirm, vbTextCompare) <> 0 Then
        Call Err.Raise(ERR_PASSWORD, "CheckInput", "The confirmation is not the same as the password, please try again.")
      End If
      Call PassWordWrite(m_benEy, sPassWord)
      CheckInput = True
  End Select
  
CheckInput_END:
  Call xReturn("CheckInput")
  Exit Function
CheckInput_ERR:
  
  lErrNumber = Err.Number
  Call ErrorMessage(ERR_ERROR, Err, "CheckInput", "Check Input", "Error in check input.")
  If lErrNumber = ERR_PASSWORD Then
    txtPassWord.SetFocus
    txtPassWord.SelStart = 0
    txtPassWord.SelLength = Len(txtPassWord.Text)
  End If
  Resume CheckInput_END
End Function

Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub cmdOK_Click()
  If CheckInput Then
    Me.Hide
    m_ok = True
  End If
End Sub

