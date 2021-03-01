VERSION 5.00
Begin VB.Form F_PayeOnlineValidate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PAYE Online Validate"
   ClientHeight    =   9135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9660
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   9660
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   8280
      TabIndex        =   1
      Top             =   8640
      Width           =   1095
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   7080
      TabIndex        =   0
      Top             =   8640
      Width           =   1095
   End
   Begin P11D2020.PayeOnlineXMLVlaidatorControl validatorControl 
      Height          =   9015
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   15901
   End
End
Attribute VB_Name = "F_PayeOnlineValidate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ContinueSubmission As Boolean
Public Sub Init(validator As PayeOnlineXMLValidator, xmlPretty As String)
  Call Me.validatorControl.Init(validator, xmlPretty)
  Call Me.Show(1, F_PayeOnline)
End Sub

Private Sub btnCancel_Click()
  ContinueSubmission = False
  Me.Visible = False
End Sub

Private Sub btnOK_Click()
  ContinueSubmission = True
  Me.Visible = False
End Sub
