VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MSM Scan"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdScan 
      Caption         =   "&Scan"
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   300
      Width           =   1035
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MSMPath As String = "I:\IShield\MSM"


Private Sub cmdScan_Click()
  Call ScanMSMs
End Sub


Private Sub ScanMSMs()
  Dim s As String
  
  s = Dir(MSMPath & "\IShield")
  Do While Len(s) > 0
    
    s = Dir
  Loop
  
End Sub
