VERSION 5.00
Begin VB.Form F_PrintCheckOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Check Before Print Options"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5430
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2768
      TabIndex        =   6
      Top             =   2100
      Width           =   1095
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1568
      TabIndex        =   5
      Top             =   2100
      Width           =   1095
   End
   Begin VB.OptionButton optCheckOptions 
      Caption         =   "Never"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   4
      Top             =   1725
      Width           =   4935
   End
   Begin VB.OptionButton optCheckOptions 
      Caption         =   "No, not this time"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   3
      Top             =   1350
      Width           =   4935
   End
   Begin VB.OptionButton optCheckOptions 
      Caption         =   "Yes, every time"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   975
      Width           =   4935
   End
   Begin VB.OptionButton optCheckOptions 
      Caption         =   "Yes"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   4935
   End
   Begin VB.Label lblPrintChecks 
      Caption         =   "Would you like to perform checks on the data before you print?"
      Height          =   255
      Left            =   150
      TabIndex        =   0
      Top             =   225
      Width           =   5535
   End
End
Attribute VB_Name = "F_PrintCheckOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bCancel As Boolean

Public Sub Start()
  
  On Error GoTo Start_ERR
  Me.Caption = S_DATA_CHECKER_WIZARD_NAME
  m_bCancel = False
  optCheckOptions_Click (-1)
  Call p11d32.Help.ShowForm(Me, vbModal)
Start_END:
  Exit Sub
Start_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "Start", "Start", "Error in Start of CompanyCarChecker.")
  Resume Start_END
  Resume
End Sub

Public Property Get Cancel() As Boolean
  Cancel = m_bCancel
End Property

Private Sub btnCancel_Click()
  m_bCancel = True
  Unload Me
End Sub

Private Sub btnOK_Click()
  Unload Me
End Sub

Private Sub optCheckOptions_Click(Index As Integer)
  
  Select Case Index
    Case -1
      Select Case p11d32.ReportPrint.CheckOptions
        Case YES_THIS_TIME_ONLY
          optCheckOptions(0) = True
        Case YES_ALWAYS
          optCheckOptions(1) = True
        Case NO_THIS_TIME_ONLY
          optCheckOptions(2) = True
        Case NEVER
          optCheckOptions(3) = True
        Case Else
          Call ECASE("Invalid Option, = " & p11d32.ReportPrint.CheckOptions)
      End Select
    Case Else
      Select Case Index
        Case 0
          p11d32.ReportPrint.CheckOptions = YES_THIS_TIME_ONLY
        Case 1
          p11d32.ReportPrint.CheckOptions = YES_ALWAYS
        Case 2
          p11d32.ReportPrint.CheckOptions = NO_THIS_TIME_ONLY
        Case 3
          p11d32.ReportPrint.CheckOptions = NEVER
      End Select
  End Select

End Sub
