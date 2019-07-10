VERSION 5.00
Begin VB.Form F_PrintReports 
   Caption         =   "Print Reports"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   ScaleHeight     =   2565
   ScaleWidth      =   3750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmd_Cancel 
      Caption         =   "Cancel"
      Height          =   372
      Left            =   1560
      TabIndex        =   3
      Top             =   1800
      Width           =   852
   End
   Begin VB.CommandButton Cmd_OK 
      Caption         =   "OK"
      Height          =   372
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   852
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   240
      Style           =   1  'Simple Combo
      TabIndex        =   1
      Text            =   "Combo2"
      Top             =   960
      Width           =   1932
   End
   Begin VB.ComboBox Combo1 
      Height          =   288
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   1932
   End
End
Attribute VB_Name = "F_PrintReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_OK_Click()
  Me.Hide
  'Call printreports
End Sub
