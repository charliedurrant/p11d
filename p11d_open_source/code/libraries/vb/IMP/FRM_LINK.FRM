VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Frm_Link 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Link Fields"
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
   Begin VB.CommandButton Cmd_ShowLinked 
      Caption         =   "Show Only Linked Fields"
      Height          =   252
      Left            =   5220
      TabIndex        =   11
      Top             =   2700
      Width           =   1932
   End
   Begin VB.ComboBox Cbo_Link 
      Height          =   288
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Finish 
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
      DragIcon        =   "Frm_link.frx":0000
      Height          =   2412
      Left            =   240
      TabIndex        =   4
      Top             =   3000
      Width           =   7092
      _ExtentX        =   12515
      _ExtentY        =   4260
      _Version        =   393216
      Cols            =   1
      AllowBigSelection=   -1  'True
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      FillStyle       =   1
      SelectionMode   =   2
   End
   Begin VB.Frame Fra_Link1 
      Height          =   1935
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   7095
      Begin MSFlexGridLib.MSFlexGrid FlG_DField 
         Bindings        =   "Frm_link.frx":014A
         DragIcon        =   "Frm_link.frx":0161
         Height          =   1212
         Left            =   5280
         TabIndex        =   6
         Top             =   600
         Width           =   1692
         _ExtentX        =   2990
         _ExtentY        =   2143
         _Version        =   393216
         Rows            =   0
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         AllowBigSelection=   -1  'True
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         FillStyle       =   1
         ScrollBars      =   2
         SelectionMode   =   2
         AllowUserResizing=   1
      End
      Begin MSFlexGridLib.MSFlexGrid FlG_DestInfo 
         DragIcon        =   "Frm_link.frx":02AB
         Height          =   1692
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   4752
         _ExtentX        =   8387
         _ExtentY        =   2990
         _Version        =   393216
         Cols            =   5
         AllowBigSelection=   -1  'True
         FocusRect       =   0
         FillStyle       =   1
         SelectionMode   =   1
      End
      Begin VB.Label Lbl_Link 
         Height          =   492
         Left            =   5280
         TabIndex        =   7
         Top             =   120
         Width           =   1572
      End
   End
   Begin VB.Label Lbl_Status 
      Height          =   615
      Left            =   300
      TabIndex        =   12
      Top             =   5520
      Width           =   3075
   End
   Begin VB.Label Lbl_Inst 
      Height          =   672
      Left            =   300
      TabIndex        =   10
      Top             =   120
      Width           =   7032
   End
   Begin VB.Label Lbl_DelimitedFile 
      Caption         =   "Source Fields:"
      Height          =   252
      Left            =   300
      TabIndex        =   3
      Top             =   2760
      Width           =   1272
   End
End
Attribute VB_Name = "Frm_Link"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_ImpWiz As ImportWizard

Public InScroll As Boolean

Implements IImportForm

Private Sub Cbo_Link_DragDrop(Source As Control, x As Single, y As Single)
  Call FlG_Source_DragDrop(Source, x, y)
End Sub

Private Sub FlG_DestInfo_SelChange()
  Call m_ImpWiz.DisplayDField
End Sub

Private Sub FlG_Source_SelChange()
  Static InSelChange As Boolean
  
  On Error Resume Next
    
  'Debug.Print "FlG_Source_SelChange"
  
  If Not InSelChange Then
    InSelChange = True
    FlG_Source.Redraw = False
    If FlG_Source.Col = 0 Then FlG_Source.Col = 1
    FlG_Source.ColSel = FlG_Source.Col
    FlG_Source.Row = LINK_FIXED_ROWCOUNT
    FlG_Source.RowSel = FlG_Source.Rows - 1
    FlG_Source.Redraw = True
    InSelChange = False
  End If
End Sub

Private Property Get IImportForm_FormType() As IMPORT_GOTOFORM
  IImportForm_FormType = TCSIMP_LINK
End Property

Private Property Set IImportForm_ImpWiz(RHS As ImportWizard)
  Set m_ImpWiz = RHS
End Property

Private Property Get IImportForm_ImpWiz() As ImportWizard
  Set IImportForm_ImpWiz = m_ImpWiz
End Property

Private Sub Cbo_Link_Click()
  Dim i As Long, j As Long
  Dim CboItem As Long
  Dim CboText As String
  Dim CurrCol As Long
  Dim CurrLinked As Boolean
  Static PrevCol As Long
  Static PrevCboText As String
  Static PrevCboLinked As Boolean
  
  CboItem = 0
  For i = 1 To Me.FlG_DestInfo.Rows - 1
    CboText = Me.Cbo_Link.List(Me.Cbo_Link.ListIndex)
    If Left$(CboText, 2) = "UN" Then CurrLinked = False Else CurrLinked = True
    j = InStr(CboText, " to ")
    If j > 0 Then
      CboText = Right$(CboText, Len(CboText) - j - 3)
      If CboText = Me.FlG_DestInfo.TextMatrix(i, 0) Then
        CboItem = i
        Exit For
      End If
    End If
  Next i
  CurrCol = Me.FlG_Source.ColSel
  'MsgBox CboText & "   " & CStr(CurrCol) & "   " & CurrLinked & vbCrLf & PrevCboText & "   " & CStr(PrevCol) & "   " & PrevCboLinked
  'If (CboText <> PrevCboText Or CurrCol <> PrevCol Or CurrLinked <> PrevCboLinked) Or (PrevCboText = "") Then
  If (CurrCol = PrevCol) And (CboText <> PrevCboText Or CurrLinked <> PrevCboLinked) Then
    'If CboItem = 0 Then
    If CurrLinked = False Then
      Call m_ImpWiz.SetDragParams(Me.Name, Me.FlG_Source.Name, CurrCol, 1, Me.Name, Me.FlG_DestInfo.Name, 1, CboItem)
    Else
      Call m_ImpWiz.SetDragParams(Me.Name, Me.FlG_DestInfo.Name, 1, CboItem, Me.Name, Me.FlG_Source.Name, Me.FlG_Source.ColSel, 1)
    End If
    Call m_ImpWiz.ProcessDrag
    Call m_ImpWiz.ReCalc_Link
  End If
  PrevCboText = CboText
  PrevCol = CurrCol
  PrevCboLinked = CurrLinked

End Sub
Private Sub Cmd_Back_Click()
  Call m_ImpWiz.ReCalc_Misc(False)
  Call SwitchForm(Me, TCSIMP_MISC)
End Sub

Private Sub Cmd_Cancel_Click()
  Call SwitchForm(Me, TCSIMP_CANCEL)
End Sub

Private Sub Cmd_Finish_Click()
  Call m_ImpWiz.ReCalc_End
  Call SwitchForm(Me, TCSIMP_END)
End Sub

Private Sub Cmd_ShowLinked_Click()
  Call m_ImpWiz.ToggleShowOnlyLinked
End Sub

Private Sub FlG_DestInfo_DragDrop(Source As Control, x As Single, y As Single)
  Call m_ImpWiz.SetDragParams("-", "-", -1, -1, Me.Name, FlG_DestInfo.Name, FlG_DestInfo.MouseCol, FlG_DestInfo.MouseRow)
  Call m_ImpWiz.ProcessDrag
End Sub

Private Sub FlG_DestInfo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If FlG_DestInfo.MouseRow = 0 Then
    Call m_ImpWiz.SetDragParams(Me.Name, FlG_DestInfo.Name, FlG_DestInfo.MouseCol, 1, "*", "*", -1, -1)
  Else
    Call m_ImpWiz.SetDragParams(Me.Name, FlG_DestInfo.Name, FlG_DestInfo.MouseCol, FlG_DestInfo.MouseRow, "*", "*", -1, -1)
  End If
  FlG_DestInfo.Drag vbBeginDrag
  FlG_DestInfo.RowSel = FlG_DestInfo.Row
  Call m_ImpWiz.DisplayDField
End Sub

Private Sub FlG_Source_DblClick()
  'Debug.Print "FlG_Source_DblClick"
  If Me.Cmd_ShowLinked.Caption = "Show Only Linked Fields" Then
    Call m_ImpWiz.SetDragParams(Me.Name, FlG_Source.Name, FlG_Source.MouseCol, FlG_Source.MouseRow, "*", "*", -1, -1)
    FlG_Source.Drag vbBeginDrag
    FlG_Source.ColSel = FlG_Source.Col
    Call m_ImpWiz.SetDragParams("-", "-", -1, -1, Me.Name, FlG_DestInfo.Name, FlG_DestInfo.MouseCol, FlG_DestInfo.MouseRow)
    'Call m_ImpWiz.ProcessDrag
    If m_ImpWiz.ProcessDrag Then Call m_ImpWiz.ReCalc_Link
  End If
End Sub

Private Sub FlG_Source_DragDrop(Source As Control, x As Single, y As Single)
  'Debug.Print "FlG_Source_DragDrop"
  If Me.Cmd_ShowLinked.Caption = "Show Only Linked Fields" Then
    FlG_Source.Col = FlG_Source.MouseCol
    Call m_ImpWiz.SetDragParams("-", "-", -1, -1, Me.Name, FlG_Source.Name, FlG_Source.MouseCol, FlG_Source.MouseRow)
    If m_ImpWiz.ProcessDrag Then
      Call m_ImpWiz.ReCalc_Link
    End If
  End If
End Sub

Private Sub FlG_Source_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  'Debug.Print "FlG_Source_MouseDown"
  Call m_ImpWiz.ReCalc_Link2(True)
End Sub

Private Sub FlG_Source_Scroll()
  On Error Resume Next
  
  'Debug.Print "FlG_Source_Scroll"
  
  InScroll = True
  Call m_ImpWiz.ReCalc_Link2(True)
  InScroll = False
End Sub


