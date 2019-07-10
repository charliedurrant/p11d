VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find in Field: "
   ClientHeight    =   930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   930
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   3015
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdFindNext 
      Caption         =   "Find &Next"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdFindFirst 
      Caption         =   "&Find First"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtFind 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   4260
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_FieldName
Private m_ac As AutoClass
Private m_acol As AutoCol

Public Sub KillReferences()
  Set m_ac = Nothing
  Set m_acol = Nothing
  m_FieldName = ""
End Sub

Public Sub Start(ByVal ac As AutoClass, ByVal acol As AutoCol, ByVal FieldName As String)
  Set m_ac = ac
  Set m_acol = acol
  m_FieldName = FieldName
  Me.Show 1
End Sub

Private Sub cmdClose_Click()
  Me.Hide
End Sub

Private Sub cmdFindFirst_Click()
  Call m_ac.FindEx(m_acol, m_FieldName, FT_FINDFIRST)
End Sub

Private Sub cmdFindNext_Click()
  Call m_ac.FindEx(m_acol, m_FieldName, FT_FINDNEXT)
End Sub

Private Sub Form_Activate()
  txtFind.SetFocus
End Sub
