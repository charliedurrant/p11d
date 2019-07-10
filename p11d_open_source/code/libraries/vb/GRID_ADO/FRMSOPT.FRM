VERSION 5.00
Begin VB.Form frmSaveOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Save Settings"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3900
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   ScaleHeight     =   2205
   ScaleWidth      =   3900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   1740
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Caption         =   "Choose the settings to save"
      Height          =   1515
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.CheckBox chkSort 
         Caption         =   "Sort order"
         Height          =   495
         Left            =   1920
         TabIndex        =   4
         Top             =   900
         Width           =   1575
      End
      Begin VB.CheckBox chkFilter 
         Caption         =   "Filter"
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   900
         Width           =   1575
      End
      Begin VB.CheckBox chkOrder 
         Caption         =   "Column Order"
         Height          =   495
         Left            =   1920
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
      Begin VB.CheckBox chkWidth 
         Caption         =   "Column Widths"
         Height          =   495
         Left            =   360
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmSaveOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
  Me.Hide
  ClearCursor
End Sub

Private Sub Form_Activate()
  SetCursor vbArrow
End Sub
