VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRepInterface 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Test Report Interface"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   4755
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   285
      Left            =   3555
      TabIndex        =   5
      Top             =   1125
      Width           =   1140
   End
   Begin MSComctlLib.StatusBar sbrTest 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   4
      Top             =   1485
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   423
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtValue 
      Height          =   330
      Left            =   1530
      TabIndex        =   3
      Top             =   1080
      Width           =   1950
   End
   Begin VB.Label lblCriteria 
      Caption         =   "Criteria SQL"
      Height          =   240
      Left            =   90
      TabIndex        =   6
      Top             =   765
      Width           =   4605
   End
   Begin VB.Label Label1 
      Caption         =   "Value"
      Height          =   240
      Left            =   45
      TabIndex        =   2
      Top             =   1125
      Width           =   1410
   End
   Begin VB.Label lblDataset 
      Caption         =   "Dataset"
      Height          =   240
      Left            =   90
      TabIndex        =   1
      Top             =   405
      Width           =   4605
   End
   Begin VB.Label lblSession 
      Caption         =   "Session"
      Height          =   240
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   4605
   End
End
Attribute VB_Name = "frmRepInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
  Me.Visible = False
  Unload Me
End Sub
