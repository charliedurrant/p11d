VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Frm_FW 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fixed Width Source File"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7635
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Cmd_Clear 
      Caption         =   "Clear Breaks"
      Height          =   375
      Left            =   6240
      TabIndex        =   11
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Frame Fra_Omit 
      Caption         =   "Omit Rows"
      Height          =   612
      Left            =   360
      TabIndex        =   3
      Top             =   1920
      Width           =   5655
      Begin VB.TextBox Txt_Omit 
         Height          =   288
         Index           =   0
         Left            =   1320
         TabIndex        =   6
         Text            =   "0"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Txt_Omit 
         Height          =   288
         Index           =   1
         Left            =   3960
         TabIndex        =   5
         Text            =   "0"
         Top             =   240
         Width           =   735
      End
      Begin MSComCtl2.UpDown UpD_Omit 
         Height          =   288
         Index           =   0
         Left            =   2056
         TabIndex        =   4
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "Txt_Omit(0)"
         BuddyDispid     =   196611
         BuddyIndex      =   0
         OrigLeft        =   2280
         OrigTop         =   240
         OrigRight       =   2520
         OrigBottom      =   525
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpD_Omit 
         Height          =   288
         Index           =   1
         Left            =   4696
         TabIndex        =   7
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "Txt_Omit(1)"
         BuddyDispid     =   196611
         BuddyIndex      =   1
         OrigLeft        =   4920
         OrigTop         =   240
         OrigRight       =   5160
         OrigBottom      =   525
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label Lbl_Omit 
         Caption         =   "Header"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Lbl_Omit 
         Caption         =   "Footer"
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton Cmd_Next 
      Caption         =   "&Next >"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Back 
      Caption         =   "< &Back"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Cancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6120
      TabIndex        =   0
      Top             =   5760
      Width           =   1215
   End
   Begin atc2imp.FWCtrl FW_FWGrid 
      Height          =   2535
      Left            =   240
      TabIndex        =   12
      Top             =   2880
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   4471
   End
   Begin VB.Label Lbl_Instruct 
      Height          =   1695
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "Frm_FW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_ImpWiz As ImportWizard
Implements IImportForm

Private Sub Form_Load()
  Lbl_Instruct = "This screen lets you set field widths (column breaks)." & vbCrLf & vbCrLf & vbCrLf & _
                     "Vertical lines signify column breaks." & vbCrLf & vbCrLf & _
                     "   To CREATE a break line, left click at the desired position." & vbCrLf & _
                     "   To DELETE a break line, right click on the line." & vbCrLf & _
                     "   To MOVE a break line, left click and drag it."
  FW_FWGrid.OriginalWidth = FW_FWGrid.Width
  FW_FWGrid.OriginalHeight = FW_FWGrid.Height
End Sub

Private Property Get IImportForm_FormType() As IMPORT_GOTOFORM
  IImportForm_FormType = TCSIMP_FW
End Property

Private Property Set IImportForm_ImpWiz(RHS As ImportWizard)
  Set m_ImpWiz = RHS
End Property

Private Property Get IImportForm_ImpWiz() As ImportWizard
  Set IImportForm_ImpWiz = m_ImpWiz
End Property

Private Sub Cmd_Back_Click()
  Call m_ImpWiz.ReCalc_Src(Nothing)
  Call SwitchForm(Me, TCSIMP_SOURCE)
End Sub

Private Sub Cmd_Cancel_Click()
  Call SwitchForm(Me, TCSIMP_CANCEL)
End Sub

Private Sub Cmd_Clear_Click()
  Cmd_Clear.Enabled = False
  Me.FW_FWGrid.ClearBreaks
  Cmd_Clear.Enabled = True
End Sub

Private Sub Cmd_Next_Click()
  Call m_ImpWiz.ReCalc_FW2(Me, False)
  Call m_ImpWiz.ReCalc_Misc(False)
  Call SwitchForm(Me, TCSIMP_MISC)
End Sub

Private Sub Txt_Omit_LostFocus(Index As Integer)
  Call OmitLines(Me.Txt_Omit(0), Me.Txt_Omit(1), 0, m_ImpWiz.NumLines)
  Call m_ImpWiz.ReCalc_FW
End Sub

Private Sub UpD_Omit_Change(Index As Integer)
  Call Txt_Omit_LostFocus(Index)
End Sub
