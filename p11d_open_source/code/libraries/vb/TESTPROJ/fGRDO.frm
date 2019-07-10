VERSION 5.00
Object = "{4E0264F8-DBA6-449A-8A7D-CB15B1D00B0F}#1.1#0"; "ATC3GRIDDAO.OCX"
Begin VB.Form frmGridRDO 
   Caption         =   "RDO Grid Form"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6420
   ScaleWidth      =   8640
   Begin VB.CommandButton cmdSumColumn 
      Caption         =   "Sum Column"
      Height          =   375
      Left            =   6570
      TabIndex        =   10
      Top             =   4860
      Width           =   1575
   End
   Begin VB.CommandButton cmdPopulate 
      Caption         =   "Populate"
      Height          =   375
      Left            =   6600
      TabIndex        =   9
      Top             =   4365
      Width           =   1575
   End
   Begin VB.Frame fraAudit 
      Caption         =   "Audit Options"
      Enabled         =   0   'False
      Height          =   1215
      Left            =   240
      TabIndex        =   5
      Top             =   4320
      Width           =   2895
      Begin VB.OptionButton optAudit 
         Caption         =   "No Auditing"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.OptionButton optAudit 
         Caption         =   "Audit All Updates"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   540
         Width           =   2535
      End
      Begin VB.OptionButton optAudit 
         Caption         =   "Audit Changes Only"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   2535
      End
   End
   Begin VB.Frame fraGrid 
      Caption         =   "Grid Options"
      Enabled         =   0   'False
      Height          =   1215
      Left            =   3480
      TabIndex        =   1
      Top             =   4320
      Width           =   2895
      Begin VB.CheckBox chkGrid 
         Caption         =   "Allow Update"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1755
      End
      Begin VB.CheckBox chkGrid 
         Caption         =   "Allow Add New"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1575
      End
      Begin VB.CheckBox chkGrid 
         Caption         =   "Allow Delete"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1815
      End
   End
   Begin atc3GRID_DAO.AutoGridCtrl_RDO GridCtrl 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   6376
      AllowUpdate     =   -1  'True
   End
   Begin VB.Label lblSum 
      Caption         =   "0"
      Height          =   375
      Left            =   6570
      TabIndex        =   11
      Top             =   5355
      Width           =   1575
   End
End
Attribute VB_Name = "frmGridRDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mAuto As atc3GRID_DAO.AutoClass
Private cnGrid As rdoConnection
Private WithEvents mAutoGrid As atc3GRID_DAO.AutoGrid
Attribute mAutoGrid.VB_VarHelpID = -1


Private Sub chkGrid_Click(Index As Integer)
  If Not mAuto.Grid Is Nothing Then
    mAuto.Grid.AllowAddNew = (Me.chkGrid(0).Value = vbChecked)
    mAuto.Grid.AllowDelete = (Me.chkGrid(1).Value = vbChecked)
    mAuto.Grid.AllowUpdate = (Me.chkGrid(2).Value = vbChecked)
  End If
End Sub

Private Sub cmdPopulate_Click()
  Dim rs As rdoResultset, sql As String
  Dim i As Long
  
  For i = 0 To cnGrid.rdoTables.Count - 1
    Debug.Print cnGrid.rdoTables(i).Name
  Next
  If mAuto.Grid Is Nothing Then
    'sql = "SELECT * from Contacts2"
    sql = "SELECT * from AuditTable"
    
    
    
    'Call mAuto.AddFieldFormat("Company", "{NOCOPY}{DROPCOMBO=AB,BD,CA,DA}{SPLIT}")
    'Call mAuto.AddFieldFormat("Salary", "{FORMAT=""#,###.0 ""}{BUTTON}{GROUP}")
    'Call mAuto.AddFieldFormat("User", "{GROUP=H}")
    
    
    Set rs = cnGrid.OpenResultset(sql, rdOpenDynamic)
    Call mAuto.InitAutoDataRDO("TestAuto", rs, Me.GridCtrl)
    Call mAuto.AddFieldFormat("TableName", "{DROPQUERY=""select TableName from DropQueryTest where UserName=<%username%>""}")
    
    'Call mAuto.AddFieldFormat("Company", "{DROPCOMBO=Arthur Andersen,Rutlege Pharm,CA,DocuPront}")
    'Call mAuto.AddFieldFormat("Salary", "{FORMAT=""#,###.0 ""}")
    'Call mAuto.AddFieldFormat("XX", "{ONCHANGE}")
    
  End If
  Call mAuto.ShowGrid
  Set mAutoGrid = mAuto.Grid
  mAuto.Grid.AllowAddNew = False
  mAuto.Grid.AllowDelete = False
  mAuto.Grid.AllowUpdate = False
  Me.fraGrid.Enabled = True
  Me.fraAudit.Enabled = True
  Me.chkGrid(0).Value = vbUnchecked
  Me.chkGrid(1).Value = vbUnchecked
  Me.chkGrid(2).Value = vbUnchecked
End Sub

Private Sub cmdSumColumn_Click()
  Dim x As Double
  If Not mAuto.Grid Is Nothing Then
    x = mAuto.Grid.TotalColumn(Me.GridCtrl.Grid.col)
    Me.lblSum.Caption = x
  End If
End Sub

Private Sub Form_Load()
  Call AddStatic("DBName", "Test_Auto")
  Call AddStatic("DataSource2", "TestAutoDS2")
  Call AddStatic("ServerName", "londs3103")
  Call AddStatic("NetDriver", "")
  Call AddStatic("User", "sa")
  Call AddStatic("Password", "ukcentral")
  
  If mAuto Is Nothing Then
    Set mAuto = New atc3GRID_DAO.AutoClass
  End If
  DatabaseTarget = DB_TARGET_SQLSERVER
  Call RegisterDataSource(GetStatic("DataSource"), DSNAttributes("TestAuto Data Source", GetStatic("ServerName"), GetStatic("DBName"), GetStatic("NetDriver")))
  Set cnGrid = RDO_Connect(DSNConnectString(GetStatic("DataSource2"), GetStatic("User"), GetStatic("Password")), rdUseOdbc)
End Sub

Private Sub Form_Resize()
  Dim xWidth As Single, yHeight As Single
  Const TOP_BORDER As Single = 50
  Const BOTTON_BORDER As Single = (TOP_BORDER * 2) + 1300
  Const MIN_WIDTH As Single = 3800
  Const MIN_HEIGHT As Single = 1055
  
  If Me.WindowState <> vbMinimized Then
    xWidth = Me.ScaleWidth - (2 * TOP_BORDER)
    yHeight = Me.ScaleHeight - BOTTON_BORDER
    If (xWidth > MIN_WIDTH) And (yHeight > MIN_HEIGHT) Then
      Me.GridCtrl.Left = TOP_BORDER
      Me.GridCtrl.Top = TOP_BORDER
      Me.GridCtrl.Width = xWidth
      Me.GridCtrl.Height = yHeight
      Me.cmdPopulate.Left = Me.ScaleWidth - Me.cmdPopulate.Width - TOP_BORDER
      Me.cmdPopulate.Top = Me.ScaleHeight - BOTTON_BORDER + TOP_BORDER
      Me.cmdSumColumn.Left = Me.cmdPopulate.Left
      Me.cmdSumColumn.Top = Me.cmdPopulate.Top + Me.cmdPopulate.Height + TOP_BORDER
      Me.lblSum.Left = Me.cmdSumColumn.Left
      Me.lblSum.Top = Me.cmdSumColumn.Top + Me.cmdSumColumn.Height + TOP_BORDER
      Me.fraAudit.Left = TOP_BORDER
      Me.fraAudit.Top = Me.ScaleHeight - Me.fraAudit.Height - (TOP_BORDER * 2)
      Me.fraGrid.Left = fraAudit.Left + fraAudit.Width + TOP_BORDER
      Me.fraGrid.Top = Me.fraAudit.Top
    End If
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not (mAuto Is Nothing) Then
    Set mAutoGrid = Nothing
    mAuto.Kill
    Set mAuto = Nothing
  End If
End Sub

Private Sub Label1_Click()

End Sub

Private Sub optAudit_Click(Index As Integer)
  Dim ta As TestAudit
  If Not mAuto.Grid Is Nothing Then
    If Me.optAudit(0).Value Then
      Set mAuto.AuditInterface = Nothing
    ElseIf Me.optAudit(1).Value Then
      Set mAuto.AuditInterface = New TestAudit
    ElseIf Me.optAudit(2).Value Then
      Set ta = New TestAudit
      Set mAuto.AuditInterface = ta
      ta.ShowChangesOnly = True
    End If
  End If
End Sub

