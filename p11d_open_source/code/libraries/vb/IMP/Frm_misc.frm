VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Frm_Misc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Field Manipulation"
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
      Left            =   6000
      TabIndex        =   0
      Top             =   5760
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid FlG_Source 
      DragIcon        =   "Frm_misc.frx":0000
      Height          =   2415
      Left            =   240
      TabIndex        =   4
      Top             =   3000
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   4260
      _Version        =   393216
      Cols            =   1
      AllowBigSelection=   -1  'True
      ScrollTrack     =   -1  'True
      Enabled         =   -1  'True
      FocusRect       =   0
      FillStyle       =   1
      SelectionMode   =   2
      AllowUserResizing=   1
   End
   Begin VB.HScrollBar HSc_Right 
      Height          =   495
      Left            =   7080
      TabIndex        =   16
      Top             =   3000
      Width           =   375
   End
   Begin VB.HScrollBar HSc_FGLeft 
      Height          =   492
      Left            =   120
      TabIndex        =   15
      Top             =   3000
      Width           =   612
   End
   Begin VB.Frame Fra_Format 
      Height          =   1212
      Left            =   240
      TabIndex        =   17
      Top             =   1440
      Width           =   7092
      Begin VB.CheckBox Chk_Factor 
         Height          =   315
         Left            =   240
         TabIndex        =   37
         Top             =   780
         Width           =   315
      End
      Begin VB.TextBox Txt_Factor 
         Height          =   315
         Left            =   2520
         TabIndex        =   36
         Top             =   780
         Width           =   975
      End
      Begin VB.ComboBox Cbo_TimeDLim 
         Height          =   288
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   720
         Width           =   1152
      End
      Begin VB.CheckBox Chk_4FigYear 
         Caption         =   "4 Figure Year ?"
         Height          =   492
         Left            =   4500
         TabIndex        =   27
         Top             =   480
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.ComboBox Cbo_DateDLim 
         Height          =   315
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   270
         Width           =   1152
      End
      Begin VB.CommandButton Cmd_FmtCancel 
         Caption         =   "(Not used now)  Cancel"
         Height          =   375
         Left            =   5640
         TabIndex        =   25
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Cmd_FmtOK 
         Caption         =   "Reformat"
         Height          =   375
         Left            =   5640
         TabIndex        =   24
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox Cbo_Boolean 
         Height          =   288
         Index           =   1
         Left            =   4080
         TabIndex        =   23
         Top             =   720
         Width           =   1092
      End
      Begin VB.ComboBox Cbo_Boolean 
         Height          =   288
         Index           =   0
         Left            =   4080
         TabIndex        =   21
         Top             =   360
         Width           =   1092
      End
      Begin VB.ComboBox Cbo_Date 
         Height          =   288
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   720
         Width           =   1272
      End
      Begin VB.ComboBox Cbo_Type 
         Height          =   288
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   360
         Width           =   1932
      End
      Begin VB.Label Lbl_Factor 
         BackStyle       =   0  'Transparent
         Caption         =   "Multiply imported values by "
         Height          =   375
         Left            =   540
         TabIndex        =   38
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Lbl_TimeDLim 
         Caption         =   "Time Delimiter"
         Height          =   372
         Left            =   2400
         TabIndex        =   31
         Top             =   660
         Width           =   672
      End
      Begin VB.Label Lbl_DateDLim 
         Caption         =   "Date Delimiter"
         Height          =   372
         Left            =   2400
         TabIndex        =   30
         Top             =   240
         Width           =   672
      End
      Begin VB.Label Lbl_DateFormat 
         Caption         =   "Format"
         Height          =   192
         Left            =   300
         TabIndex        =   29
         Top             =   780
         Width           =   552
      End
      Begin VB.Label Lbl_Boolean 
         Caption         =   "False"
         Height          =   252
         Index           =   1
         Left            =   3360
         TabIndex        =   22
         Top             =   720
         Width           =   492
      End
      Begin VB.Label Lbl_Boolean 
         Caption         =   "True"
         Height          =   252
         Index           =   0
         Left            =   3360
         TabIndex        =   20
         Top             =   360
         Width           =   492
      End
   End
   Begin VB.Frame Fra_Static 
      Caption         =   "Enter the details for the selected static field into the box below, then press the OK button"
      Height          =   1215
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Visible         =   0   'False
      Width           =   7095
      Begin VB.OptionButton Opt_Static 
         Caption         =   "Special Value"
         Height          =   315
         Index           =   1
         Left            =   360
         TabIndex        =   35
         Top             =   780
         Width           =   1275
      End
      Begin VB.OptionButton Opt_Static 
         Caption         =   "User Value"
         Height          =   315
         Index           =   0
         Left            =   360
         TabIndex        =   34
         Top             =   360
         Width           =   1275
      End
      Begin VB.ComboBox Cbo_Static 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   780
         Width           =   2475
      End
      Begin VB.CommandButton Cmd_Static 
         Caption         =   "OK"
         Height          =   375
         Left            =   5760
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Cmd_StaticCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   5760
         TabIndex        =   10
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Txt_Static 
         Height          =   285
         Left            =   1800
         TabIndex        =   8
         Top             =   360
         Width           =   2475
      End
      Begin VB.Label Lbl_Static3 
         Caption         =   "Static Value"
         Height          =   375
         Left            =   4740
         TabIndex        =   33
         Top             =   660
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Lbl_Static2 
         Caption         =   "Static Value"
         Height          =   375
         Left            =   4740
         TabIndex        =   9
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin VB.Frame Fra_DelStatic 
      Height          =   1215
      Left            =   240
      TabIndex        =   11
      Top             =   1440
      Visible         =   0   'False
      Width           =   7095
      Begin VB.CommandButton Cmd_DelYes 
         Caption         =   "Yes"
         Height          =   375
         Left            =   1800
         TabIndex        =   13
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton Cmd_DelNo 
         Caption         =   "No"
         Height          =   375
         Left            =   3960
         TabIndex        =   12
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Lbl_DelStatic 
         Caption         =   "Do you want to delete the selected static field ?"
         Height          =   255
         Left            =   1800
         TabIndex        =   14
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.Label Lbl_Misc1 
      Height          =   1152
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   7092
   End
   Begin VB.Label Lbl_Fields 
      Caption         =   "Modified source field contents:"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Menu mnuMisc 
      Caption         =   "Miscellaneous"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuMerge 
         Caption         =   "Merge"
      End
      Begin VB.Menu mnuMergeSpaces 
         Caption         =   "Merge with spaces"
      End
      Begin VB.Menu mnuSeparate 
         Caption         =   "Separate merged fields"
      End
      Begin VB.Menu mnuAddStatic 
         Caption         =   "Add static column"
      End
      Begin VB.Menu mnuDelStatic 
         Caption         =   "Delete static column"
      End
      Begin VB.Menu mnuFormat 
         Caption         =   "Format field"
      End
      Begin VB.Menu mnuCopyField 
         Caption         =   "Copy field"
      End
   End
End
Attribute VB_Name = "Frm_Misc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_ImpWiz As ImportWizard
Implements IImportForm

Private Sub Cbo_DateDLim_Click()
  Chk_4FigYear.Visible = (Cbo_DateDLim.Text = "{None}")
End Sub

'MPSMarch2
Private Sub Chk_Factor_Click()
  Lbl_Factor.Enabled = -Chk_Factor.Value
  Txt_Factor.Enabled = -Chk_Factor.Value
End Sub

Private Sub FlG_Source_SelChange()
  Static InSelChange As Boolean
  On Error Resume Next
  If Not InSelChange Then
    InSelChange = True
    Call m_ImpWiz.FormatField
    InSelChange = False
  End If
End Sub

Private Property Get IImportForm_FormType() As IMPORT_GOTOFORM
  IImportForm_FormType = TCSIMP_MISC
End Property

Private Property Set IImportForm_ImpWiz(RHS As ImportWizard)
  Set m_ImpWiz = RHS
End Property

Private Property Get IImportForm_ImpWiz() As ImportWizard
  Set IImportForm_ImpWiz = m_ImpWiz
End Property

Private Sub Cbo_Type_Click()
  Call TypeToScreen
End Sub

Private Sub TypeToScreen()
  
  On Error GoTo TypeToScreen_ERR
  Call xSet("TypeToScreen")
  If StrComp(Cbo_Type.Text, "DATE", vbTextCompare) = 0 Then
    Cbo_Date.Visible = True
    Chk_4FigYear.Visible = True
    Cbo_DateDLim.Visible = True
    Cbo_TimeDLim.Visible = True
    Lbl_DateFormat.Visible = True
    Lbl_DateDLim.Visible = True
    Lbl_TimeDLim.Visible = True
  Else
    Cbo_Date.Visible = False
    Chk_4FigYear.Visible = False
    Cbo_DateDLim.Visible = False
    Cbo_TimeDLim.Visible = False
    Lbl_DateFormat.Visible = False
    Lbl_DateDLim.Visible = False
    Lbl_TimeDLim.Visible = False
  End If
  Call Cbo_DateDLim_Click
  If Cbo_Type.Text = "BOOLEAN" Then
    Lbl_Boolean(0).Visible = True
    Lbl_Boolean(1).Visible = True
    Cbo_Boolean(0).Visible = True
    Cbo_Boolean(1).Visible = True
  Else
    Lbl_Boolean(0).Visible = False
    Lbl_Boolean(1).Visible = False
    Cbo_Boolean(0).Visible = False
    Cbo_Boolean(1).Visible = False
  End If
  'MPSMarch2
  If StrComp(Cbo_Type.Text, "LONG", vbTextCompare) = 0 Or StrComp(Cbo_Type.Text, "DOUBLE", vbTextCompare) = 0 Then
    Lbl_Factor.Visible = True
    Chk_Factor.Visible = True
    Txt_Factor.Visible = True
    Call m_ImpWiz.DisplayFactorFormatting
  Else
    Lbl_Factor.Visible = False
    Chk_Factor.Visible = False
    Txt_Factor.Visible = False
  End If
  
TypeToScreen_END:
  Call xReturn("TypeToScreen")
  Exit Sub
  
TypeToScreen_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "TypeToScreen", "Type To Screen", "Error placing the type to the screen.")
  Resume TypeToScreen_END
End Sub
Private Sub mnuAddStatic_Click()
  Call m_ImpWiz.CopyColumn(-1)
End Sub

Private Sub Cmd_Back_Click()
  If m_ImpWiz.ImpParent.ImportType = IMPORT_DELIMITED Then
    Call m_ImpWiz.ReCalc_DLim(False)
    Call SwitchForm(Me, TCSIMP_DLIM)
  Else
    Call m_ImpWiz.ReCalc_FW2(Nothing, False)
    Call SwitchForm(Me, TCSIMP_FW)
  End If
End Sub

Private Sub Cmd_Cancel_Click()
  Call SwitchForm(Me, TCSIMP_CANCEL)
End Sub

Private Sub mnuCopyField_Click()
  Call m_ImpWiz.CopyColumn(Me.FlG_Source.Coldata(Me.FlG_Source.Col))
End Sub

Private Sub mnuDelStatic_Click()
  Call m_ImpWiz.DeleteColumn(Me.FlG_Source.Coldata(Me.FlG_Source.Col))
  Call m_ImpWiz.ReCalc_Misc(False)
End Sub

Private Sub Cmd_DelNo_Click()
  Frm_Misc.Fra_DelStatic.Visible = False
  FlG_Source.Enabled = True
  Cmd_Cancel.Enabled = True
  Cmd_Back.Enabled = True
  Cmd_Next.Enabled = True
  Call m_ImpWiz.ReCalc_Misc(False)
End Sub

Private Sub Cmd_DelYes_Click()
     
  Call m_ImpWiz.DeleteColumn(Me.FlG_Source.Coldata(Me.FlG_Source.Col))
  Frm_Misc.Fra_DelStatic.Visible = False
  FlG_Source.Enabled = True
  Cmd_Cancel.Enabled = True
  Cmd_Back.Enabled = True
  Cmd_Next.Enabled = True
  Call m_ImpWiz.ReCalc_Misc(False)
End Sub

Private Sub Cmd_FmtCancel_Click()
  FlG_Source.Enabled = True
  Cmd_Cancel.Enabled = True
  Cmd_Back.Enabled = True
  Cmd_Next.Enabled = True
  Lbl_Misc1.Visible = True
  Fra_Format.Visible = False
End Sub

Private Sub Cmd_FmtOK_Click()
  Call m_ImpWiz.FormatFieldOK
End Sub

Private Sub mnuMerge_Click()
  'Call m_ImpWiz.MergeColumns("")
End Sub

Private Sub Cmd_Next_Click()
  Call m_ImpWiz.ReCalc_Link
  Call SwitchForm(Me, TCSIMP_LINK)
End Sub

Private Sub mnuSeparate_Click()
  'Call m_ImpWiz.SplitOK
End Sub

Private Sub Cmd_Static_Click()
  Call m_ImpWiz.SetStatic(Me.FlG_Source.Coldata(Me.FlG_Source.ColSel), Me.Txt_Static.Text)
  If Opt_Static(1).Value Then Call m_ImpWiz.SetSpecialFieldKey(FlG_Source.Coldata(Me.FlG_Source.ColSel), Cbo_Static.ItemData(Cbo_Static.ListIndex)) 'MPSMarch2
  Me.Fra_Static.Visible = False
  Me.FlG_Source.Enabled = True
  Me.Cmd_Cancel.Enabled = True
  Me.Cmd_Back.Enabled = True
  Me.Cmd_Next.Enabled = True
  Call m_ImpWiz.ReCalc_Misc(False)
End Sub

Private Sub Cmd_StaticCancel_Click()
  Call m_ImpWiz.DeleteColumn(Me.FlG_Source.Coldata(Me.FlG_Source.Col))
  Frm_Misc.Fra_Static.Visible = False
  FlG_Source.Enabled = True
  Cmd_Cancel.Enabled = True
  Cmd_Back.Enabled = True
  Cmd_Next.Enabled = True
  Call m_ImpWiz.ReCalc_Misc(False)
End Sub

Private Sub mnuFormat_Click()
  Call m_ImpWiz.FormatField
End Sub

Private Sub FlG_Source_DragDrop(Source As Control, x As Single, y As Single)
  Call m_ImpWiz.SetDragParams("-", "-", -1, -1, Me.Name, FlG_Source.Name, FlG_Source.MouseCol, FlG_Source.MouseRow)
  Call m_ImpWiz.ProcessDrag
End Sub

Private Sub FlG_Source_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 1 Then
    If FlG_Source.MouseRow < 2 Then
      If m_ImpWiz.AllowMoveColumns Then
        Call m_ImpWiz.SetDragParams(Me.Name, FlG_Source.Name, FlG_Source.MouseCol, FlG_Source.MouseRow, "*", "*", -1, -1)
        FlG_Source.Drag vbBeginDrag
      End If
    End If
  End If
  If Button = 2 Then
    If FlG_Source.Col = FlG_Source.ColSel Then
      FlG_Source.Col = FlG_Source.MouseCol
      FlG_Source.Row = FIXED_ROWCOUNT
      FlG_Source.RowSel = FlG_Source.Rows - 1
    End If
    Call m_ImpWiz.CreateMiscMenu
  End If
End Sub

Private Sub HSc_FGLeft_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  
  If Me.FlG_Source.LeftCol >= 2 Then
    Me.FlG_Source.LeftCol = Me.FlG_Source.LeftCol - 1
  End If
End Sub

Private Sub HSc_Right_DragOver(Source As Control, x As Single, y As Single, State As Integer)

  If Me.FlG_Source.LeftCol <= Me.FlG_Source.cols Then
    Me.FlG_Source.LeftCol = Me.FlG_Source.LeftCol + 1
  End If
End Sub


'MPSMarch2
Private Sub Opt_Static_Click(Index As Integer)
  Txt_Static.Enabled = Not (-Index)
  Cbo_Static.Enabled = (-Index)
End Sub


'MPSMarch2
' Add / Modify controls on Frm_Misc
