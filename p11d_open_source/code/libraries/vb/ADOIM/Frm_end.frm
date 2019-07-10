VERSION 5.00
Object = "{AF27A9B5-A3F4-11D2-8DB7-00C04FA9DD6F}#1.2#0"; "tcsprog.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Frm_End 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import Data"
   ClientHeight    =   6375
   ClientLeft      =   4410
   ClientTop       =   3630
   ClientWidth     =   7635
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TCSPROG.TCSProgressBar PBar_Import 
      Height          =   240
      Left            =   1530
      TabIndex        =   18
      Top             =   1230
      Width           =   5835
      _cx             =   4204596
      _cy             =   4194727
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Min             =   0
      Max             =   100
      Value           =   0
      BarBackColor    =   12632256
      BarForeColor    =   8388608
      Appearance      =   1
      Style           =   0
      CaptionColor    =   0
      CaptionInvertColor=   16777215
      FillStyle       =   0
      FadeFromColor   =   0
      FadeToColor     =   16777215
      Caption         =   ""
      InnerCircle     =   0   'False
      Percentage      =   2
      Skew            =   0
      PictureOffsetTop=   0
      PictureOffsetLeft=   0
      Enabled         =   0   'False
      Increment       =   1
      TextAlignment   =   1
   End
   Begin VB.Frame fraImportAnalysis 
      Caption         =   "Import analysis"
      Height          =   1770
      Left            =   4410
      TabIndex        =   11
      Top             =   1485
      Visible         =   0   'False
      Width           =   2895
      Begin VB.Label lblLinesInError 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblLinesInError"
         Height          =   195
         Left            =   315
         TabIndex        =   17
         Top             =   495
         Width           =   2415
      End
      Begin VB.Label lblRecordsUpdated 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblRecordsUpdated"
         Height          =   195
         Left            =   315
         TabIndex        =   16
         Top             =   1500
         Width           =   2415
      End
      Begin VB.Label lblRecordsAdded 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblRecordsAdded"
         Height          =   195
         Left            =   315
         TabIndex        =   15
         Top             =   1260
         Width           =   2415
      End
      Begin VB.Label lblPostProcessingLinesInError 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblPostProcessingLinesInError"
         Height          =   195
         Left            =   315
         TabIndex        =   14
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label lblLinesOK 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblLinesOK"
         Height          =   195
         Left            =   315
         TabIndex        =   13
         Top             =   945
         Width           =   2415
      End
      Begin VB.Label lblFileLines 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblFileLines"
         Height          =   240
         Left            =   225
         TabIndex        =   12
         Top             =   225
         Width           =   2490
      End
      Begin VB.Line Line1 
         X1              =   90
         X2              =   2745
         Y1              =   1215
         Y2              =   1215
      End
      Begin VB.Line Line2 
         X1              =   2745
         X2              =   90
         Y1              =   450
         Y2              =   450
      End
   End
   Begin VB.CommandButton Cmd_ErrView 
      Caption         =   "View Errors"
      Height          =   375
      Left            =   300
      TabIndex        =   10
      Top             =   2400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_ErrPrint 
      Caption         =   "Print Errors"
      Height          =   375
      Left            =   1500
      TabIndex        =   9
      Top             =   2400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Another 
      Caption         =   "Import More"
      Height          =   375
      Left            =   4695
      TabIndex        =   7
      Top             =   5760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Import 
      Caption         =   "Import"
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_SaveSpec 
      Caption         =   "Save Spec"
      Height          =   375
      Left            =   1140
      TabIndex        =   3
      ToolTipText     =   "Press the Save Spec button to save the import specification into a file, for future use."
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
      Left            =   6000
      TabIndex        =   0
      Top             =   5760
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid FlG_Dest 
      Bindings        =   "Frm_end.frx":0000
      DragIcon        =   "Frm_end.frx":0017
      Height          =   2055
      Left            =   240
      TabIndex        =   2
      Top             =   3360
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   3625
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
      AllowBigSelection=   -1  'True
      FocusRect       =   0
      HighLight       =   0
      FillStyle       =   1
      SelectionMode   =   2
      AllowUserResizing=   1
   End
   Begin VB.Label Lbl_ImpProgress 
      Caption         =   "Import progress"
      Height          =   195
      Left            =   270
      TabIndex        =   8
      Top             =   1215
      Width           =   1215
   End
   Begin VB.Label Lbl_Info1 
      Height          =   1020
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   7095
   End
   Begin VB.Label Lbl_Dest 
      Caption         =   "Destination database:"
      Height          =   372
      Left            =   240
      TabIndex        =   5
      Top             =   3000
      Width           =   3312
   End
End
Attribute VB_Name = "Frm_End"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_ImpWiz As ImportWizard
Implements IImportForm

Private Sub Cmd_ErrPrint_Click()
  Call m_ImpWiz.ImpParent.ErrorFilter.PrintErrors
End Sub

Private Sub Cmd_ErrView_Click()
  Call m_ImpWiz.ImpParent.ErrorFilter.ViewErrors
End Sub

Private Property Get IImportForm_FormType() As IMPORT_GOTOFORM
  IImportForm_FormType = TCSIMP_END
End Property

Private Property Set IImportForm_ImpWiz(RHS As ImportWizard)
  Set m_ImpWiz = RHS
End Property

Private Property Get IImportForm_ImpWiz() As ImportWizard
  Set IImportForm_ImpWiz = m_ImpWiz
End Property

Private Sub Cmd_Back_Click()
  Call m_ImpWiz.ReCalc_Link
  Call SwitchForm(Me, TCSIMP_LINK)
End Sub

Private Sub Cmd_Cancel_Click()
  Call SwitchForm(Me, TCSIMP_CANCEL)
End Sub

Private Sub Cmd_Import_Click()
  Call m_ImpWiz.DoImport(Me.FlG_Dest, 0, False, True)
End Sub

Private Sub Cmd_SaveSpec_Click()
  Dim s As String
  
  s = FileSaveAsDlg("Choose a file in which to save the specification", "Import Specification Files (*.imp)|*.imp", m_ImpWiz.SpecPath)
  If Len(s) > 0 Then
    Call SplitPath(s, m_ImpWiz.SourcePath)
    Call m_ImpWiz.SaveSpec(s)
  End If
End Sub

Private Sub Cmd_Another_Click()
  m_ImpWiz.ImportAnother = True
  Call SwitchForm(Me, TCSIMP_CANCEL)
End Sub

