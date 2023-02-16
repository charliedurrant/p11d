VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_Source 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Source Data File"
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
   Begin atc2imp.FWCtrl FW_Source 
      Height          =   1485
      Left            =   240
      TabIndex        =   12
      Top             =   4170
      Visible         =   0   'False
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   2619
   End
   Begin VB.Frame Fra_Format 
      Caption         =   "Choose the format which describes your data:"
      Height          =   2685
      Left            =   135
      TabIndex        =   7
      Top             =   1185
      Width           =   7335
      Begin VB.PictureBox panelRecentSpecs 
         BorderStyle     =   0  'None
         Height          =   1320
         Left            =   75
         ScaleHeight     =   1320
         ScaleWidth      =   7185
         TabIndex        =   13
         Top             =   1230
         Width           =   7185
         Begin VB.CommandButton cmdDeletePreviousSpec 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   6825
            TabIndex        =   15
            ToolTipText     =   "Delete"
            Top             =   105
            Width           =   300
         End
         Begin MSComctlLib.ListView listViewRecentSpecs 
            Height          =   1215
            Left            =   2085
            TabIndex        =   14
            Top             =   60
            Width           =   4725
            _ExtentX        =   8334
            _ExtentY        =   2143
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            Checkboxes      =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.Label lblRecentSpecs 
            Caption         =   "Or 'check' a spec previously used for this import type"
            Height          =   600
            Left            =   45
            TabIndex        =   16
            Top             =   60
            Width           =   2100
            WordWrap        =   -1  'True
         End
      End
      Begin VB.OptionButton Opt_Format 
         Caption         =   "Fixed &Width - Fields are aligned in columns with spaces between each field"
         Height          =   255
         Index           =   1
         Left            =   105
         TabIndex        =   11
         Top             =   540
         Width           =   5895
      End
      Begin VB.OptionButton Opt_Format 
         Caption         =   "&Delimited - Characters such as comma or tab separate each field"
         Height          =   255
         Index           =   0
         Left            =   105
         TabIndex        =   10
         Top             =   255
         Value           =   -1  'True
         Width           =   5175
      End
      Begin VB.CommandButton Cmd_Spec 
         Caption         =   "Open Spec."
         Height          =   375
         Left            =   5970
         TabIndex        =   9
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Lbl_Spec 
         Caption         =   "Or press the Open Spec. button to open a file which contains the format specification for your data"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   5730
      End
   End
   Begin VB.CommandButton Cmd_OpenSource 
      Caption         =   "Open"
      Height          =   375
      Left            =   6120
      TabIndex        =   5
      Top             =   675
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
      Height          =   195
      Left            =   150
      TabIndex        =   6
      Top             =   3930
      Width           =   7095
   End
   Begin VB.Label Lbl_SourcePath 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Source File Path"
      Height          =   465
      Left            =   255
      TabIndex        =   4
      Top             =   630
      Width           =   5685
   End
   Begin VB.Label Lbl_SourceInst 
      Caption         =   "Source File Instructions"
      Height          =   435
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
Private Sub cmdDeletePreviousSpec_Click()
  Dim lvi As ListItem
  Dim i As Long
  
  On Error GoTo err_err
  
  For i = listViewRecentSpecs.ListItems.Count To 1 Step -1
    Set lvi = listViewRecentSpecs.ListItems(i)
    If (lvi.Selected) Then
      Call listViewRecentSpecs.ListItems.Remove(i)
      Call m_ImpWiz.RemoveRecentSpec(lvi.Text)
    End If
  Next
  
err_end:
  Exit Sub
err_err:
  Call ErrorMessage(ERR_ERROR, Err, "DeletePreviousSpec", "Delete previous spec", Err.Description)
  Resume err_end
End Sub
Private Sub Form_Load()
  FW_Source.OriginalWidth = FW_Source.Width
  FW_Source.OriginalHeight = FW_Source.Height
  
  Dim ch As ColumnHeader
  Set ch = listViewRecentSpecs.ColumnHeaders.Add(, , "Spec")
  ch.Width = listViewRecentSpecs.Width
End Sub

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
  If Not m_ImpWiz.ReCalc_Dest Then Call SwitchForm(Me, TCSIMP_CANCEL, False)
  Call SwitchForm(Me, TCSIMP_DEST, False)
End Sub

Private Sub Cmd_Cancel_Click()
  Call SwitchForm(Me, TCSIMP_CANCEL, False)
End Sub
Private Property Get PreviousSelectedSpec() As String
  Dim lvi As ListItem
  
  PreviousSelectedSpec = ""
  For Each lvi In listViewRecentSpecs.ListItems
    If lvi.Checked Then
      PreviousSelectedSpec = lvi.Text
      Exit Property
    End If
    
  Next
  
End Property
Private Sub Cmd_Next_Click()
  If Len(PreviousSelectedSpec) > 0 Then
    m_ImpWiz.LoadSpec (PreviousSelectedSpec)
    Exit Sub
  End If
  Call m_ImpWiz.ReCalc_Src(Me)
  If m_ImpWiz.ImpParent.ImportType = IMPORT_DELIMITED Then
    Call m_ImpWiz.ReCalc_DLim(False)
    Call SwitchForm(Me, TCSIMP_DLIM, True)
  Else
    Call m_ImpWiz.ReCalc_FW
    Call SwitchForm(Me, TCSIMP_FW, True)
  End If
End Sub
Private Sub Cmd_OpenSource_Click()
  Call m_ImpWiz.OpenSourceFile(Me)
End Sub
Private Sub Cmd_Spec_Click()
  Call ClearPreviousSelectedSpecs
  Call m_ImpWiz.LoadSpec
End Sub
Private Sub ClearPreviousSelectedSpecs()
  Dim lvi As ListItem
  
  For Each lvi In listViewRecentSpecs.ListItems
    lvi.Checked = False
  Next
End Sub

Private Sub listViewRecentSpecs_ItemCheck(ByVal Item As MSComctlLib.ListItem)
  
  Dim lvi As ListItem
  
  If Not Item.Checked Then Exit Sub
  
  For Each lvi In listViewRecentSpecs.ListItems
    If Not Item Is lvi And lvi.Checked Then
      lvi.Checked = False
    End If
  Next
End Sub

Private Sub Opt_Format_Click(Index As Integer)
  If Index = 0 Then
    m_ImpWiz.ImpParent.ImportType = IMPORT_DELIMITED
  Else
    m_ImpWiz.ImpParent.ImportType = IMPORT_FIXED
  End If
  m_ImpWiz.ImpParent.HeaderCount = -1
End Sub
