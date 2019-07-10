VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPrintDialog 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Print"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6195
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   375
      Left            =   3480
      TabIndex        =   14
      Top             =   1755
      Width           =   1230
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4800
      TabIndex        =   13
      Top             =   1755
      Width           =   1230
   End
   Begin VB.Frame fmeCopies 
      Caption         =   "Copies"
      Height          =   1590
      Left            =   3360
      TabIndex        =   8
      Top             =   45
      Width           =   2715
      Begin MSComCtl2.UpDown updCopies 
         Height          =   330
         Left            =   2160
         TabIndex        =   12
         Top             =   315
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Value           =   1
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtCopies"
         BuddyDispid     =   196612
         OrigLeft        =   2205
         OrigTop         =   360
         OrigRight       =   2445
         OrigBottom      =   600
         Max             =   1
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtCopies 
         Height          =   330
         Left            =   1665
         TabIndex        =   10
         Text            =   "1"
         Top             =   315
         Width           =   705
      End
      Begin VB.PictureBox picCopies 
         BorderStyle     =   0  'None
         Height          =   690
         Left            =   90
         Picture         =   "Printdlg.frx":0000
         ScaleHeight     =   690
         ScaleWidth      =   2220
         TabIndex        =   11
         Top             =   765
         Width           =   2220
      End
      Begin VB.Label Label1 
         Caption         =   "Number of &copies:"
         Height          =   285
         Left            =   180
         TabIndex        =   9
         Top             =   360
         Width           =   1545
      End
   End
   Begin VB.Frame fmePrint 
      Caption         =   "Print range"
      Height          =   1590
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   3300
      Begin MSComCtl2.UpDown updFrom 
         Height          =   330
         Left            =   1755
         TabIndex        =   15
         Top             =   600
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtFrom"
         BuddyDispid     =   196617
         OrigLeft        =   1920
         OrigTop         =   600
         OrigRight       =   2160
         OrigBottom      =   975
         Max             =   1
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtTo 
         Height          =   330
         Left            =   2400
         TabIndex        =   6
         Text            =   "1"
         Top             =   600
         Width           =   465
      End
      Begin VB.TextBox txtFrom 
         Height          =   330
         Left            =   1260
         TabIndex        =   4
         Text            =   "1"
         Top             =   600
         Width           =   465
      End
      Begin VB.OptionButton optPrint 
         Caption         =   "&Current Page"
         Height          =   330
         Index           =   2
         Left            =   135
         TabIndex        =   7
         Top             =   1080
         Width           =   2445
      End
      Begin VB.OptionButton optPrint 
         Caption         =   "Pages"
         Height          =   285
         Index           =   1
         Left            =   135
         TabIndex        =   2
         Top             =   675
         Width           =   780
      End
      Begin VB.OptionButton optPrint 
         Caption         =   "&All"
         Height          =   330
         Index           =   0
         Left            =   135
         TabIndex        =   1
         Top             =   225
         Width           =   2580
      End
      Begin MSComCtl2.UpDown updTo 
         Height          =   330
         Left            =   2910
         TabIndex        =   16
         Top             =   600
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtTo"
         BuddyDispid     =   196616
         OrigLeft        =   2205
         OrigTop         =   360
         OrigRight       =   2445
         OrigBottom      =   600
         Max             =   1
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lblTo 
         Caption         =   "&to:"
         Height          =   195
         Left            =   2085
         TabIndex        =   5
         Top             =   690
         Width           =   195
      End
      Begin VB.Label lblFrom 
         Caption         =   "&from:"
         Height          =   240
         Left            =   900
         TabIndex        =   3
         Top             =   720
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmPrintDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Cancel As Boolean

Private Sub cmdCancel_Click()
  Cancel = True
  Me.Visible = False
End Sub

Private Sub cmdPrint_Click()
  Cancel = False
  Me.Visible = False
End Sub

Private Sub Form_Load()
  Cancel = True
End Sub

Private Sub SetOptions()
  If optPrint(PAGES_RANGE).Value Then
    txtFrom.Enabled = True
    txtTo.Enabled = True
    updFrom.Enabled = True
    updTo.Enabled = True
    lblTo.Enabled = True
    lblFrom.Enabled = True
  Else
    txtFrom.Enabled = False
    txtTo.Enabled = False
    updFrom.Enabled = False
    updTo.Enabled = False
    lblTo.Enabled = False
    lblFrom.Enabled = False
  End If
End Sub

Private Sub optPrint_Click(Index As Integer)
  Call SetOptions
End Sub

Private Sub txtCopies_Validate(Cancel As Boolean)
  Dim l As Long
  On Error Resume Next
  Cancel = True
  If IsNumeric(txtCopies) Then
    l = CLng(txtCopies)
    Cancel = Not ((l > 0) And (l <= MAX_PAGE_COPIES))
  End If
End Sub

Private Sub txtFrom_Validate(Cancel As Boolean)
  Dim l As Long
  
  On Error Resume Next
  Cancel = True
  If IsNumeric(txtFrom.text) Then
    l = CLng(txtFrom)
    If (l >= 1) And (l <= ReportControl.Pages_N) And (l <= CLng(txtTo)) Then
      Cancel = False
    End If
  End If
End Sub

Private Sub txtTo_Validate(Cancel As Boolean)
  Dim l As Long
  
  On Error Resume Next
  Cancel = True
  If IsNumeric(txtTo.text) Then
    l = CLng(txtTo)
    If (l >= 1) And (l <= ReportControl.Pages_N) And (l >= CLng(txtFrom)) Then
      Cancel = False
    End If
  End If
End Sub
