VERSION 5.00
Begin VB.Form frmTimeOut 
   Caption         =   "Time Out Dialog"
   ClientHeight    =   2115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3300
   LinkTopic       =   "Form1"
   ScaleHeight     =   2115
   ScaleWidth      =   3300
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRetry 
      Caption         =   "Continue Submit"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel Submit"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label lblMessage 
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmTimeOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Retry As Boolean

Private Sub cmdCancel_Click()
  Retry = False
  Me.Hide
End Sub

Private Sub cmdRetry_Click()
  Retry = True
  Me.Hide
End Sub

