VERSION 5.00
Object = "{80CAD458-7206-11D1-BAF2-00C04FA4C373}#5.0#0"; "TXTALIGN.OCX"
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   5355
   StartUpPosition =   3  'Windows Default
   Begin pTxtAlign.TXTAlign TXTAlign1 
      Height          =   2055
      Left            =   720
      TabIndex        =   4
      Top             =   960
      Width           =   2895
      _extentx        =   5106
      _extenty        =   3625
      text            =   ""
      fontsize        =   8.25
      fontname        =   "MS Sans Serif"
      font            =   "frmTest.frx":0000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Index           =   2
      Left            =   3120
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Index           =   1
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   1092
      Left            =   600
      TabIndex        =   3
      Top             =   4320
      Width           =   3972
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click(Index As Integer)
    frmTest.TXTAlign1.TXTAlign = Index
    Label1.Caption = frmTest.TXTAlign1.MaxLength
End Sub

