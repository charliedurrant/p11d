VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Frm_RW 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Report Wizard"
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
   Begin VB.Frame Fra_Fields 
      BorderStyle     =   0  'None
      Height          =   5472
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   7392
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   432
         Left            =   3300
         TabIndex        =   7
         Top             =   1620
         Width           =   1452
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   432
         Left            =   3300
         TabIndex        =   6
         Top             =   960
         Width           =   1452
      End
      Begin MSFlexGridLib.MSFlexGrid FlG_Cols 
         Height          =   1812
         Left            =   300
         TabIndex        =   5
         Top             =   2760
         Width           =   5772
         _ExtentX        =   10181
         _ExtentY        =   3196
         _Version        =   65541
         FocusRect       =   0
         SelectionMode   =   2
      End
      Begin MSFlexGridLib.MSFlexGrid FlG_Fields 
         DragMode        =   1  'Automatic
         Height          =   1812
         Left            =   240
         TabIndex        =   4
         Top             =   180
         Width           =   2412
         _ExtentX        =   4255
         _ExtentY        =   3196
         _Version        =   65541
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
      Enabled         =   0   'False
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Exit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6120
      TabIndex        =   0
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Menu Mnu_RepWiz 
      Caption         =   "RepWizMenu"
      Visible         =   0   'False
      Begin VB.Menu Mnu_AddCol 
         Caption         =   "Add Column"
      End
      Begin VB.Menu Mnu_InsCol 
         Caption         =   "Insert Column"
      End
      Begin VB.Menu Mnu_DelCol 
         Caption         =   "Delete Column"
      End
   End
End
Attribute VB_Name = "Frm_RW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Option Explicit
''
''Private DragFromRow As Long
''Private DragFromCol As Long
''Private DragToRow As Long
''Private DragToCol As Long
''
''Private Sub Cmd_Exit_Click()
''  End
''End Sub
''
''Private Sub Command1_Click()
''  Dim i As Long, j As Long
''  FlG_Cols.Rows = 5
''  FlG_Cols.Cols = 5
''  Call FlG_Cols_SelChange
''  For i = 1 To 4
''    For j = 1 To 4
''      FlG_Cols.TextMatrix(i, j) = CStr(i) & ":" & CStr(j)
''    Next j
''  Next i
''End Sub
''
''Private Sub Command2_Click()
''  FlG_Cols.ColPosition(2) = 4
''End Sub
''
''Private Sub FlG_Cols_DragDrop(Source As Control, x As Single, y As Single)
''  DragToRow = FlG_Cols.MouseRow
''  DragToCol = FlG_Cols.MouseCol
''  If Source.Name = FlG_Fields.Name Then
''    MsgBox "Fields > Cols"
''  End If
''  If Source.Name = FlG_Cols.Name Then
''    If DragToCol = 0 Then DragToCol = 1
''    FlG_Cols.ColPosition(DragFromCol) = DragToCol
''    FlG_Cols.Col = DragToCol
''    Call FlG_Cols_SelChange
''  End If
''End Sub
''
''Private Sub FlG_Cols_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
''  If Button = 1 Then
''    If FlG_Cols.MouseCol >= FlG_Cols.FixedCols And FlG_Cols.MouseRow < FlG_Cols.FixedRows Then
''      DragFromRow = FlG_Cols.MouseRow
''      DragFromCol = FlG_Cols.MouseCol
''      FlG_Cols.Drag vbBeginDrag
''    End If
''  Else
''    If FlG_Cols.MouseCol >= FlG_Cols.FixedCols Then
''      PopupMenu Mnu_RepWiz
''    End If
''  End If
''End Sub
''
''Private Sub FlG_Cols_SelChange()
''  'Dim Col As Long, ColSel As Long
''  Static InChange As Boolean
''  If Not InChange Then
''    InChange = True
''    'If ColSel < Col Then
''    '  Col = FlG_Cols.ColSel
''    '  ColSel = FlG_Cols.Col
''    'Else
''    '  Col = FlG_Cols.Col
''    '  ColSel = FlG_Cols.ColSel
''    'End If
''    FlG_Cols.Redraw = False
''    FlG_Cols.Row = FlG_Cols.FixedRows
''    'FlG_Cols.Col = Col
''    'FlG_Cols.ColSel = ColSel
''    FlG_Cols.Col = FlG_Cols.Col  ' Uncomment this line restrict selection to one column only
''    FlG_Cols.RowSel = FlG_Cols.Rows - 1
''    FlG_Cols.Redraw = True
''    InChange = False
''  End If
''End Sub
''
''Private Sub FlG_Fields_DragDrop(Source As Control, x As Single, y As Single)
''  DragToRow = FlG_Fields.MouseRow
''  DragToCol = FlG_Fields.MouseCol
''  If Source.Name = FlG_Cols.Name Then
''    MsgBox "Cols > Fields"
''  End If
''End Sub
''
''Private Sub FlG_Fields_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
''  DragFromRow = FlG_Fields.MouseRow
''  DragFromCol = FlG_Fields.MouseCol
''  FlG_Fields.Drag vbBeginDrag
''End Sub
''
''Private Sub Mnu_AddCol_Click()
''  FlG_Cols.Cols = FlG_Cols.Cols + 1
''  FlG_Cols.Col = FlG_Cols.Cols - 1
''  Call FlG_Cols_SelChange
''End Sub
''
''Private Sub Mnu_DelCol_Click()
''  Dim Col As Long
''  Col = FlG_Cols.Col
''  FlG_Cols.ColPosition(FlG_Cols.Col) = FlG_Cols.Cols - 1
''  FlG_Cols.Cols = FlG_Cols.Cols - 1
''  If Col >= FlG_Cols.Cols Then FlG_Cols.Col = FlG_Cols.Cols - 1
''  Call FlG_Cols_SelChange
''End Sub
''
''Private Sub Mnu_InsCol_Click()
''  Dim Col As Long
''  Col = FlG_Cols.Col
''  FlG_Cols.Cols = FlG_Cols.Cols + 1
''  FlG_Cols.ColPosition(FlG_Cols.Cols - 1) = Col
''  FlG_Cols.Col = Col
''  Call FlG_Cols_SelChange
''End Sub
