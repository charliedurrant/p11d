VERSION 5.00
Begin VB.Form F_PayeOnlineWarning 
   ClientHeight    =   1440
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      ToolTipText     =   "Cancel PAYE Online submission"
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdProceed 
      Caption         =   "&Proceed"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      ToolTipText     =   "Proceed with PAYE Online submission"
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&View"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      ToolTipText     =   "View validation warnings and errors"
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblMessage 
      Height          =   735
      Left            =   1200
      TabIndex        =   3
      Top             =   120
      Width           =   3495
   End
   Begin VB.Image imgAlert_exclamation 
      Height          =   480
      Left            =   240
      Picture         =   "F_PayeOnlineWarning.frx":0000
      Top             =   120
      Width           =   465
   End
   Begin VB.Image imgAlert_info 
      Height          =   495
      Left            =   240
      Picture         =   "F_PayeOnlineWarning.frx":0C42
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "F_PayeOnlineWarning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Enum PAYE_MESSAGE_TYPE
  PMT_WARNINGS_ONLY
  PMT_WARNINGS_AND_ERRORS
End Enum

Private m_lMessageType As PAYE_MESSAGE_TYPE
Private m_Submit As Boolean
Public Function Start(lMessageType As PAYE_MESSAGE_TYPE) As Boolean
  
  Call SetCursor(vbIconPointer)
  m_lMessageType = lMessageType
  Select Case lMessageType
    Case PMT_WARNINGS_ONLY
      Me.Caption = "Warnings in submission!"
      Me.lblMessage = "Do you wish to view the warnings?"
      Me.imgAlert_exclamation.Visible = False
      Me.imgAlert_info.Visible = True
    Case PMT_WARNINGS_AND_ERRORS
      Me.Caption = "Warnings/Errors in submission!"
      Me.lblMessage = "There are validation errors that will prevent you from being able to submit." & vbCrLf & "Do you wish to view the warnings/errors?"
      Me.imgAlert_exclamation.Visible = True
      Me.imgAlert_info.Visible = False
      Me.cmdProceed.Visible = p11d32.PAYEonline.ViewProceedButtonIfErrors
      
  End Select
'  Me.Show vbModal
  Call p11d32.Help.ShowForm(Me, vbModal)
  Start = m_Submit
End Function


Private Sub cmdCancel_Click()
  m_Submit = False
  Unload Me
  
End Sub

Private Sub cmdProceed_Click()
  m_Submit = True
  Unload Me
End Sub

Private Sub cmdView_Click()
  Call p11d32.PAYEonline.Errors(PREPARE_REPORT, VET_PAYEONLINE_VALIDATION)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call ClearCursor
End Sub
