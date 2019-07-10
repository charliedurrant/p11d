VERSION 5.00
Begin VB.Form F_Find 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3885
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   915
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   2790
      TabIndex        =   3
      Top             =   540
      Width           =   1050
   End
   Begin VB.CommandButton cmdFindNext 
      Caption         =   "Find &Next"
      Height          =   330
      Left            =   1710
      TabIndex        =   2
      Top             =   540
      Width           =   1005
   End
   Begin VB.TextBox txtFindWhat 
      Height          =   330
      Left            =   855
      TabIndex        =   0
      Top             =   135
      Width           =   2985
   End
   Begin VB.Label lblFindWhat 
      Caption         =   "Find what:"
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   180
      Width           =   735
   End
End
Attribute VB_Name = "F_Find"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public rt As RichTextBox
Private m_FindPos As Long
Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  m_FindPos = rt.SelStart
End Sub

Private Sub txtFindWhat_Change()
  m_FindPos = rt.SelStart
End Sub
