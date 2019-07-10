VERSION 5.00
Object = "{4E0264F8-DBA6-449A-8A7D-CB15B1D00B0F}#1.1#0"; "ATC3GRIDDAO.OCX"
Begin VB.Form frmGrid 
   Caption         =   "Grid Form"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6420
   ScaleWidth      =   8640
   Begin atc3GRID_DAO.AutoGridCtrl_DAO gridctrl 
      Height          =   3480
      Left            =   90
      TabIndex        =   11
      Top             =   180
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   6138
      AllowUpdate     =   -1  'True
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Print Preview"
      Height          =   330
      Left            =   6885
      TabIndex        =   10
      Top             =   4005
      Width           =   1545
   End
   Begin VB.CommandButton cmdAddRecord 
      Caption         =   "Add Record"
      Height          =   330
      Left            =   6840
      TabIndex        =   9
      Top             =   4455
      Width           =   1545
   End
   Begin VB.Frame fraGrid 
      Caption         =   "Grid Options"
      Enabled         =   0   'False
      Height          =   1215
      Left            =   3240
      TabIndex        =   5
      Top             =   4080
      Width           =   2895
      Begin VB.CheckBox chkGrid 
         Caption         =   "Allow Delete"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1815
      End
      Begin VB.CheckBox chkGrid 
         Caption         =   "Allow Add New"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1575
      End
      Begin VB.CheckBox chkGrid 
         Caption         =   "Allow Update"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1755
      End
   End
   Begin VB.Frame fraAudit 
      Caption         =   "Audit Options"
      Enabled         =   0   'False
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Width           =   2895
      Begin VB.OptionButton optAudit 
         Caption         =   "Audit Changes Only"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   2535
      End
      Begin VB.OptionButton optAudit 
         Caption         =   "Audit All Updates"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   540
         Width           =   2535
      End
      Begin VB.OptionButton optAudit 
         Caption         =   "No Auditing"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   2535
      End
   End
   Begin VB.CommandButton cmdPopulate 
      Caption         =   "Populate"
      Height          =   375
      Left            =   6840
      TabIndex        =   0
      Top             =   4920
      Width           =   1575
   End
End
Attribute VB_Name = "frmGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mAuto As atc3GRID_DAO.AutoClass
Private WithEvents mAutoGrid As atc3GRID_DAO.AutoGrid
Attribute mAutoGrid.VB_VarHelpID = -1
Private mDB As dao.Database


Private Sub chkGrid_Click(Index As Integer)
  If Not mAuto.Grid Is Nothing Then
    mAuto.Grid.AllowAddNew = (Me.chkGrid(0).Value = vbChecked)
    mAuto.Grid.AllowDelete = (Me.chkGrid(1).Value = vbChecked)
    mAuto.Grid.AllowUpdate = (Me.chkGrid(2).Value = vbChecked)
  End If
End Sub

Private Sub cmdAddRecord_Click()
  Dim sql As String
  
  sql = "INSERT INTO Contacts ( FirstName, SurName, Company, DOB, Salary, [Year of Contact] ) VALUES('NewFirstName','NewSurname','Deloitte & Touche',#1/1/1999#,20000,1999)"
  Call mDB.Execute(sql)
  
End Sub

Private Sub cmdPopulate_Click()
  
  Dim rs As Recordset
  
  If mAuto.Grid Is Nothing Then
    Set mDB = InitDB(gwsMain, AppPath & "\" & "test.mdb", "Test Database Name")
    'Set rs = mDB.OpenRecordset("SELECT * from AuditTable ")
    Set rs = mDB.OpenRecordset("SELECT * from Contacts")
    
    Call mAuto.InitAutoData("TestAuto", rs, mDB.Name, Me.GridCtrl)
    Call mAuto.AddFieldFormat("Company", "{NOCOPY}{Caption=Company~DROPCOMBO}{DROPCOMBO=AB,CA,CB}")
    'Call mAuto.AddFieldFormat("Text3", "{NOCOPY}{Caption=Company~DROPLIST}{DROPLIST=AB,value_AB,CA,value_CA,CB,value_CB}")
    Call mAuto.AddFieldFormat("Text4", "{NOCOPY}{Caption=Company~DROP}{DROP=AB,value_AB,CA,value_CA,CB,value_CB}")
    Call mAuto.AddFieldFormat("Text2", "{CAPTION=Company~DROPQUERY}{DROPQUERY=""SELECT User, id FROM DropQueryTest""}")
    Call mAuto.AddFieldFormat("Salary", "{FORMAT=""#,###.0 ""}{BUTTON}")
    Call mAuto.AddFieldFormat("SelfEmployed", "{BOOLEAN}")
  End If
  mAuto.CheckBoxCross = False
  mAuto.PrintGroupHeader = False
  mAuto.SpanMultiplePages = True
  Call mAuto.ShowGrid
  Set mAutoGrid = mAuto.Grid
  mAuto.Grid.AllowAddNew = False
  mAuto.Grid.AllowDelete = False
  mAuto.Grid.AllowUpdate = False
  mAuto.ValidateDataTypes = TYPE_DATE
  Me.fraGrid.Enabled = True
  Me.fraAudit.Enabled = True
  Me.chkGrid(0).Value = vbUnchecked
  Me.chkGrid(1).Value = vbUnchecked
  Me.chkGrid(2).Value = vbUnchecked
End Sub

Private Sub cmdPreview_Click()
  Dim rep As Reporter
  
  Set rep = New Reporter
  Call rep.InitReport("Test Report", PREPARE_REPORT, PORTRAIT)
  rep.ReportFooter = "HELLO"
  mAuto.ReportHeader = "JJJJJJJJJ"
  Call mAuto.ShowReport(rep)
  Call rep.EndReport
  rep.PreviewReport
  'Call rep.ExportReport("C:\test.xls", EXPORT_EXCEL, True)
End Sub

Private Sub Command1_Click()
  Call mAuto.Grid.CopyColumnValue("SurName", "Salary")
End Sub

Private Sub Form_Load()
  If mAuto Is Nothing Then
    Set mAuto = New atc3GRID_DAO.AutoClass
  End If
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
      Me.cmdAddRecord.Left = Me.ScaleWidth - Me.cmdAddRecord.Width - TOP_BORDER
      Me.cmdAddRecord.Top = Me.cmdPopulate.Top + Me.cmdPopulate.Height + TOP_BORDER
      
      Me.cmdPreview.Left = Me.ScaleWidth - Me.cmdPreview.Width - TOP_BORDER
      Me.cmdPreview.Top = Me.cmdAddRecord.Top + Me.cmdAddRecord.Height + TOP_BORDER
            
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

Private Sub mAutoGrid_ButtonClick(ByVal AutoName As String, ByVal FieldName As String)
  Dim col As TrueDBGrid60.Column
  'MsgBox "Clicked " & FieldName
  'Set col = mAuto.Grid.TDBGrid.Columns(FieldName)
  If mAuto.Grid.AllowUpdate Then
    Set col = Me.GridCtrl.Grid.Columns.Item(mAuto.Grid.GetGridColIndex(FieldName))
    If col.Value = 0 Then
      col.Value = 10000
    Else
      col.Value = 0
    End If
  End If
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

Private Sub mAutoGrid_ValidateRow(Cancel As Boolean, ValidationError As String, ByVal AutoName As String, ByVal AutoColumns As Collection)
  Dim rs As Recordset
  
  'set rs = mAutoGrid.GetDropRecordset("CCN")
  
End Sub


