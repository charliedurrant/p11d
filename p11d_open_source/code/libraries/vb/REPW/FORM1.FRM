VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2556
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   2556
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   432
      Left            =   540
      TabIndex        =   0
      Top             =   540
      Width           =   1932
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  Dim x As tcsReport.ReportWizard
  
  Set x = New tcsReport.ReportWizard
  Set x.ReportInterface = CreateObject("udmAdapt.UDM")
  x.ReportFileName = "R:\VB5\ADAPT\DATA\DELME98.ABD"
  Call x.StartReport
End Sub
