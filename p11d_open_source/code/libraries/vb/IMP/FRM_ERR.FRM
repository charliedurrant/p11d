VERSION 5.00
Begin VB.Form Frm_Source 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Source Data File"
   ClientHeight    =   6372
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7632
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6372
   ScaleWidth      =   7632
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TCSIMP.FWCtrl FW_Source 
      Height          =   2052
      Left            =   240
      TabIndex        =   12
      Top             =   3600
      Visible         =   0   'False
      Width           =   7092
      _ExtentX        =   12510
      _ExtentY        =   3620
   End
   Begin VB.Frame Fra_Format 
      Caption         =   "Choose the format which describes your data:"
      Height          =   1575
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   7335
      Begin VB.OptionButton Opt_Format 
         Caption         =   "Fixed &Width - Fields are aligned in columns with spaces between each field"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   5895
      End
      Begin VB.OptionButton Opt_Format 
         Caption         =   "&Delimited - Characters such as comma or tab separate each field"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Value           =   -1  'True
         Width           =   5175
      End
      Begin VB.CommandButton Cmd_Spec 
         Caption         =   "Open Spec."
         Height          =   375
         Left            =   6000
         TabIndex        =   9
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Lbl_Spec 
         Caption         =   "Or press the Open Spec. button to open a file which contains the format specification for your data"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   5415
      End
   End
   Begin VB.CommandButton Cmd_OpenSource 
      Caption         =   "Open"
      Height          =   375
      Left            =   6120
      TabIndex        =   5
      Top             =   840
      Width           =   1215
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
   Begin VB.Label Lbl_SrcContents 
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3240
      Width           =   7095
   End
   Begin VB.Label Lbl_SourcePath 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Source File Path"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   5415
   End
   Begin VB.Label Lbl_SourceInst 
      Caption         =   "Source File Instructions"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   7095
   End
End
Attribute VB_Name = "Frm_Source"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_ImpWiz As ImportWizard
Implements IImportForm

Private Property Get IImportForm_FormType() As IMPORT_GOTOFORM
  IImportForm_FormType = TCSIMP_SOURCE
End Property

Private Property Set IImportForm_ImpWiz(RHS As ImportWizard)
  Set m_ImpWiz = RHS
End Property

Private Property Get IImportForm_ImpWiz() As ImportWizard
  Set IImportForm_ImpWiz = m_ImpWiz
End Property

Private Sub Cmd_Back_Click()
  Call m_ImpWiz.ReCalc_Dest
  Call SwitchForm(Me, TCSIMP_DEST)
End Sub

Private Sub Cmd_Cancel_Click()
  Call SwitchForm(Me, TCSIMP_CANCEL)
End Sub

Private Sub Cmd_Next_Click()
  Call m_ImpWiz.ReCalc_Src(Me)
  If m_ImpWiz.ImpParent.ImportType = IMPORT_DELIMITED Then
    Call m_ImpWiz.ReCalc_DLim(False)
    Call SwitchForm(Me, TCSIMP_DLIM)
  Else
    Call m_ImpWiz.ReCalc_FW
    Call SwitchForm(Me, TCSIMP_FW)
  End If
End Sub

Private Sub Cmd_OpenSource_Click()
  Call m_ImpWiz.OpenSourceFile(Me)
End Sub

Private Sub Cmd_Spec_Click()
  Call m_ImpWiz.LoadSpec
End Sub

Private Sub Opt_Format_Click(Index As Integer)
  If Index = 0 Then
    m_ImpWiz.ImpParent.ImportType = IMPORT_DELIMITED
  Else
    m_ImpWiz.ImpParent.ImportType = IMPORT_FIXED
  End If
  m_ImpWiz.ImpParent.HeaderCount = -1
End Sub
