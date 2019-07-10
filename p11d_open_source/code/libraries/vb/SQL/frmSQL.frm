VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{4E0264F8-DBA6-449A-8A7D-CB15B1D00B0F}#1.1#0"; "atc3griddao.OCX"
Object = "{D7D47D2E-20A1-45D1-B08B-3A509726296E}#1.0#0"; "atc2split.OCX"
Begin VB.Form frmSQL 
   Caption         =   "Debug"
   ClientHeight    =   7110
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9480
   Icon            =   "frmSQL.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7110
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDetails 
      Height          =   825
      Left            =   2205
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   10
      Tag             =   "EQUALISER"
      Top             =   2565
      Width           =   6765
   End
   Begin VB.CheckBox chkAllowEditing 
      Caption         =   "Allow editing"
      Height          =   330
      Left            =   90
      TabIndex        =   9
      Tag             =   "LOCKB"
      Top             =   6660
      Width           =   1725
   End
   Begin MSComctlLib.ImageList iml 
      Left            =   2700
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSQL.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSQL.frx":041C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSQL.frx":052E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSQL.frx":0640
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSQL.frx":0B92
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSQL.frx":10E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSQL.frx":1636
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSQL.frx":1B88
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSQL.frx":1C9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSQL.frx":1DAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSQL.frx":22FE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Close"
      Height          =   420
      Left            =   8280
      TabIndex        =   6
      Tag             =   "LOCKBR"
      Top             =   6660
      Width           =   1140
   End
   Begin VB.CommandButton cmdExecute 
      Height          =   375
      Left            =   9000
      Picture         =   "frmSQL.frx":2850
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "LOCKR"
      Top             =   810
      Width           =   465
   End
   Begin VB.CommandButton cmdFileOpen 
      Height          =   375
      Left            =   9000
      Picture         =   "frmSQL.frx":2952
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "LOCKR"
      Top             =   180
      Width           =   465
   End
   Begin VB.TextBox txtFileName 
      Height          =   375
      Left            =   45
      TabIndex        =   2
      Tag             =   "EQUALISER"
      Top             =   180
      Width           =   8925
   End
   Begin VB.TextBox txtSQL 
      Height          =   1500
      Left            =   2205
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Tag             =   "EQUALISER"
      Top             =   810
      Width           =   6765
   End
   Begin ATC3Grid_DAO.AutoGridCtrl_DAO GridCtrl 
      Height          =   3030
      Left            =   2205
      TabIndex        =   0
      Tag             =   "EQUALISER,EQUALISEB"
      Top             =   3510
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   5345
      AllowUpdate     =   -1  'True
   End
   Begin ATC2SPLIT.SPLIT split 
      Height          =   5730
      Left            =   2025
      TabIndex        =   12
      Tag             =   "EQUALISEB"
      Top             =   810
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   10107
      MinBorderPixels =   50
   End
   Begin MSComctlLib.TreeView tvwStructure 
      Height          =   5775
      Left            =   45
      TabIndex        =   13
      Tag             =   "EQUALISEB"
      Top             =   810
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   10186
      _Version        =   393217
      Indentation     =   0
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "iml"
      Appearance      =   1
   End
   Begin VB.Label lblDetails 
      Caption         =   "Details"
      Height          =   195
      Left            =   2205
      TabIndex        =   11
      Tag             =   "EQUALISEB"
      Top             =   2385
      Width           =   1545
   End
   Begin VB.Label lblStructure 
      Caption         =   "Structure"
      Height          =   195
      Left            =   45
      TabIndex        =   8
      Tag             =   "LOCK"
      Top             =   585
      Width           =   1725
   End
   Begin VB.Label lblSQL 
      Caption         =   "Enter SQL statement"
      Height          =   240
      Left            =   2205
      TabIndex        =   7
      Tag             =   "LOCK"
      Top             =   585
      Width           =   1770
   End
   Begin VB.Label lblDataBase 
      Caption         =   "Database"
      Height          =   195
      Left            =   45
      TabIndex        =   5
      Tag             =   "LOCK"
      Top             =   0
      Width           =   2220
   End
End
Attribute VB_Name = "frmSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mDB As Database
Private mAuto As AutoClass
Private mAllowUpdate As Boolean

Private Const SQL_SELECT As String = "SELECT "
Private Const SQL_DELETE As String = "DELETE "
Private Const SQL_INSERT As String = "INSERT "
Private Const SQL_UPDATE As String = "UPDATE "
Private Const SQL_PARAMS As String = "PARAMETERS "
Private Const S_TABLES As String = "TABLES"
Private Const S_UNKNOWN As String = "UNKNOWN"

Private Const SNG_SPLIT_GAP As Single = 50
Public Enum NODE_TYPE
  NT_TABLE = 1
  NT_FIELD
  NT_PROPERTY
  NT_QUERY_SELECT
  NT_QUERY_UPDATE
  NT_QUERY_INSERT
  NT_QUERY_DELETE
  NT_UNKNOWN
  NT_SQL
  NT_FOLDER_CLOSED
  NT_FOLDER_OPEN
End Enum
'CAD convert split functions to singles and sort out cos it is rubbish

Private m_SQLDebug As SQLDebug
Private m_CR As clsFormResize

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  mAllowUpdate = True
  Call Split.Initialise(Me.hWnd, False)
  
  Set m_CR = New clsFormResize
  Call m_CR.InitResize(Me, 7515, 9600, VGA)
  
End Sub

Private Sub CreateAutoClass()
  If mAuto Is Nothing Then
    Set mAuto = New AutoClass
  End If
End Sub

Private Sub Kill()
 If Not (mAuto Is Nothing) Then
    Call mAuto.Kill
    Set mAuto = Nothing
  End If
End Sub

Private Sub OpenFile()
  Dim sFileAndPath As String
  
  On Error GoTo OpenFile_ERR
  sFileAndPath = FileOpenDlg(m_SQLDebug.FileCaption, m_SQLDebug.FileFilter, m_SQLDebug.CurrentDirectory)
  If Len(sFileAndPath) > 0 Then
    Set mDB = Nothing
    Call InitialiseForm(sFileAndPath, Nothing)
    If Not mDB Is Nothing Then
      Call SplitPath(sFileAndPath, sFileAndPath)
      m_SQLDebug.CurrentDirectory = sFileAndPath
    End If
  End If
  
OpenFile_END:
  Exit Sub
  
OpenFile_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "OpenFile", "Open File", "Error opening a file.")
  Resume OpenFile_END
End Sub

Private Sub chkAllowEditing_Click()
  Call SetEdits
End Sub

Private Sub cmdFileOpen_Click()
  Call OpenFile
End Sub

Private Sub Form_LostFocus()
  Call Me.SetFocus
End Sub

Private Sub Form_Resize()
  Call m_CR.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set m_SQLDebug = Nothing
  Call Kill
  Set mDB = Nothing
End Sub

Public Sub InitialiseForm(vDB As Variant, SQLDebug As SQLDebug)
  Dim DB As Database
  
  On Error GoTo InitialiseForm_ERR
  Call SetCursor
  Call Kill
  Call CreateAutoClass
  DoEvents
  If Not SQLDebug Is Nothing Then Set m_SQLDebug = SQLDebug
  If VarType(vDB) = vbString Then Set DB = InitDB(m_SQLDebug.WorkSpace, vDB, "SQL Debug", , m_SQLDebug.OpenExclusive)
  If VarType(vDB) = vbObject Then
    If TypeOf vDB Is Database Then Set DB = vDB
  End If
  If DB Is Nothing Then GoTo InitialiseForm_END
  Set mDB = DB
  txtFileName = DB.Name
  
InitialiseForm_END:
  Call FillTreeview
  Call ClearCursor
  Exit Sub
  
InitialiseForm_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "InitialiseForm", "Initialise Form", "Error initialising the SQLDebug form.")
  Resume InitialiseForm_END
  Resume
End Sub

Private Function NodeTypeToParentKey(ByVal NT As NODE_TYPE) As String
  Select Case NT
    Case NT_TABLE
      NodeTypeToParentKey = S_TABLES
    Case NT_QUERY_SELECT
      NodeTypeToParentKey = SQL_SELECT
    Case NT_QUERY_UPDATE
      NodeTypeToParentKey = SQL_UPDATE
    Case NT_QUERY_INSERT
      NodeTypeToParentKey = SQL_INSERT
    Case NT_QUERY_DELETE
      NodeTypeToParentKey = SQL_DELETE
    Case NT_UNKNOWN
      NodeTypeToParentKey = S_UNKNOWN
    Case Else
      Call ECASE("Invalid NODE_TYPE =" & NT)
  End Select
End Function

Private Sub FillTreeview()
  Dim td As TableDef
  Dim n As Node, nChild As Node
  Dim nodes As nodes, NodeImage As NODE_TYPE
  Dim qd As QueryDef
  Dim sql As String
  
  On Error GoTo FillTreeview_err
  mAllowUpdate = False
  
  tvwStructure.Enabled = False
  tvwStructure.nodes.Clear
  
  If mDB Is Nothing Then GoTo FillTreeview_end
  Set nodes = tvwStructure.nodes
  
  'add parent nodes
  Call nodes.Add(, , S_TABLES, "Tables", NT_FOLDER_CLOSED)
  
  For Each td In mDB.TableDefs
    If Not IsSysTable(td) Then
      Set n = nodes.Add(nodes(NodeTypeToParentKey(NT_TABLE)), tvwChild, , td.Name, NT_TABLE)
      sql = "Select * from [" & n.Text & "]"
      Set nChild = nodes.Add(n, tvwChild, , sql, NT_SQL)
      nChild.Tag = sql
    End If
  Next td
  
  
  Call nodes.Add(, , SQL_SELECT, "Select queries", NT_FOLDER_CLOSED)
  Call nodes.Add(, , SQL_INSERT, "Insert queries", NT_FOLDER_CLOSED)
  Call nodes.Add(, , SQL_DELETE, "Delete queries", NT_FOLDER_CLOSED)
  Call nodes.Add(, , SQL_UPDATE, "Update queries", NT_FOLDER_CLOSED)
  Call nodes.Add(, , S_UNKNOWN, "Other queries", NT_FOLDER_CLOSED)
  For Each qd In mDB.QueryDefs
    sql = qd.sql
    NodeImage = GetQueryType(sql)
    Set n = nodes.Add(nodes(NodeTypeToParentKey(NodeImage)), tvwChild, , qd.Name, NodeImage)
    Set nChild = nodes.Add(n, tvwChild, , ReplaceString(sql, vbCrLf, " "), NT_SQL)
    nChild.Tag = sql
  Next qd
  
FillTreeview_end:
  Me.tvwStructure.Enabled = True
  mAllowUpdate = True
  Exit Sub
  
FillTreeview_err:
  Call ErrorMessage(ERR_ERROR, Err, "FillTreeview", "Fill Treeview", "Error filling tree view with Table/Query details.")
  Resume FillTreeview_end
  Resume
End Sub


Private Sub ExecuteSQL()
  Dim qdTemp As QueryDef
  Dim status As String
  Dim nType As NODE_TYPE
  Dim sql As String, rs As Recordset
  Dim s As String, i As Long
  
  On Error GoTo ExecuteSQL_err
  Call SetCursor
  status = "Results:" & vbCrLf
  
  If txtSQL.SelLength > 0 Then
    sql = Trim$(txtSQL.SelText)
  Else
    sql = Trim$(Me.txtSQL)
  End If

  
  If Len(sql) > 0 Then
    Set qdTemp = mDB.CreateQueryDef("", sql)
    For i = 0 To (qdTemp.Parameters.Count - 1)
      s = InputBox$("Parameter " & qdTemp.Parameters(i).Name, "Enter value for parameter", "")
      qdTemp.Parameters(i).Value = s
    Next i
    If GetQueryType(sql) = NT_QUERY_SELECT Then
      ' sql select
      Set rs = qdTemp.OpenRecordset(dbOpenDynaset)
      Call DisplayGrid(rs, sql)
      status = status & "Recordset displayed"
    Else
      ' sql execute
      Call qdTemp.Execute
      status = status & "Records affected: " & mDB.RecordsAffected
      If Not mAuto.grid Is Nothing Then Call mAuto.grid.Refresh(True)
    End If
  Else
    status = status & "SQL string not specified"
  End If
  
ExecuteSQL_end:
  Me.txtDetails = status
  Call ClearCursor
  Exit Sub
  
ExecuteSQL_err:
  status = "Error:" & vbCrLf & Err.Description
  Resume ExecuteSQL_end
End Sub

Private Sub DisplayGrid(rs As Recordset, ByVal sql As String)
  On Error GoTo DisplayGrid_ERR
  
  If mAuto.grid Is Nothing Then
    If Not rs Is Nothing Then
      Call mAuto.InitAutoDataSQL("DebugSQLAuto", rs, sql, mDB.Name, Me.GridCtrl)
      Set mAuto.WorkSpace = m_SQLDebug.WorkSpace
    End If
  Else
    Call mAuto.SetNewRS(rs, , sql)
  End If
  Call mAuto.ShowGrid
  Call SetEdits
DisplayGrid_END:
  Exit Sub
DisplayGrid_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "DisplayGrid", "Display Grid", "Error displaying the grid DisplayGrid.")
  Resume DisplayGrid_END
  Resume
End Sub
Private Sub SetEdits()
  If Not mAuto.grid Is Nothing Then
    mAuto.grid.AllowAddNew = chkAllowEditing.Value
    mAuto.grid.AllowDelete = chkAllowEditing.Value
    mAuto.grid.AllowUpdate = chkAllowEditing.Value
  End If
End Sub
Private Sub cmdExecute_Click()
  Call ExecuteSQL
End Sub

Private Sub FillFields(nParent As Node)
  Dim td As TableDef
  Dim n As Node, nChild As Node
  Dim s As String
  Dim fld As Field
  
  On Error GoTo FillTables_err
  If nParent.Children > 1 Then GoTo FillTables_end
  Set td = mDB.TableDefs(nParent.Text)
  
  For Each fld In td.Fields
    Set n = tvwStructure.nodes.Add(nParent, tvwChild, , fld.Name, NT_FIELD)
    s = "Select [" & td.Name & "].[" & fld.Name & "] from " & td.Name
    Set nChild = tvwStructure.nodes.Add(n, tvwChild, , s, NT_SQL)
    nChild.Tag = s
  Next fld
  
FillTables_end:
  Exit Sub
  
FillTables_err:
  Call ErrorMessage(ERR_ALLOWRETRY, Err, "FillFields", "Fill Fields", "Error filling the fields.")
  Resume FillTables_end
End Sub

Private Sub FillProperties(nParent As Node)
  Dim td As TableDef
  Dim nParentsParent As Node
  Dim fld As Field
  
  On Error GoTo FillProperties_err
  If nParent.Children > 1 Then GoTo FillProperties_end
  Set nParentsParent = nParent.Parent
  Set td = mDB.TableDefs(nParentsParent.Text)
  Set fld = td.Fields(nParent.Text)
  
  Call tvwStructure.nodes.Add(nParent, tvwChild, , "Allow zero length: " & fld.AllowZeroLength, NT_PROPERTY)
  Call tvwStructure.nodes.Add(nParent, tvwChild, , "Default value: " & fld.DefaultValue, NT_PROPERTY)
  Call tvwStructure.nodes.Add(nParent, tvwChild, , "Required: " & fld.Required, NT_PROPERTY)
    
FillProperties_end:
  Exit Sub
  
FillProperties_err:
  Call ErrorMessage(ERR_ERROR, Err, "FillProperties", "Fill Properties", "Error filling the properties.")
  Resume FillProperties_end
End Sub


Private Sub split_FinishedSplit(ByVal MovementInTwips As Single)
   tvwStructure.Width = tvwStructure.Width + MovementInTwips
  
  txtDetails.Left = txtDetails.Left + MovementInTwips
  txtDetails.Width = txtDetails.Width - MovementInTwips
  
  txtSQL.Left = txtDetails.Left
  txtSQL.Width = txtDetails.Width
  
  lblDetails.Left = txtDetails.Left
  lblSQL.Left = txtDetails.Left
  
  Call m_CR.ReDoAspectRatios("txtSQL")
  Call m_CR.ReDoAspectRatios("txtDetails")
  GridCtrl.Left = txtDetails.Left
  GridCtrl.Width = GridCtrl.Width - MovementInTwips
  
  Call m_CR.ReDoAspectRatios("GridCtrl")
 
  
 
End Sub

Private Sub tvwStructure_Collapse(ByVal Node As MSComctlLib.Node)
  If Node.Image = NT_FOLDER_OPEN Then Node.Image = NT_FOLDER_CLOSED
End Sub

Private Sub tvwStructure_Expand(ByVal Node As MSComctlLib.Node)
  Select Case Node.Image
    
    Case NT_TABLE
      Call FillFields(Node)
    Case NT_FIELD
      Call FillProperties(Node)
    Case NT_FOLDER_CLOSED
      Node.Image = NT_FOLDER_OPEN
  End Select
End Sub

Private Sub tvwStructure_NodeClick(ByVal Node As MSComctlLib.Node)
  If Node.Image = NT_SQL Then txtSQL = Node.Tag
End Sub

Private Sub txtFileName_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then Call InitialiseForm(txtFileName.Text, Nothing)
End Sub

Private Function GetQueryType(ByVal sql As String) As NODE_TYPE
  Dim p As Long
  
  GetQueryType = NT_UNKNOWN
  sql = Trim$(sql)
  p = InStr(1, sql, SQL_PARAMS, vbTextCompare)
  If p = 1 Then
    p = InStr(1, sql, ";", vbTextCompare)
    If p > 0 Then sql = LTrimAny(Mid$(sql, p + 1), vbCrLf)
  End If
  If InStr(1, sql, SQL_SELECT, vbTextCompare) = 1 Then GetQueryType = NT_QUERY_SELECT: Exit Function
  If InStr(1, sql, SQL_INSERT, vbTextCompare) = 1 Then GetQueryType = NT_QUERY_INSERT: Exit Function
  If InStr(1, sql, SQL_DELETE, vbTextCompare) = 1 Then GetQueryType = NT_QUERY_DELETE: Exit Function
  If InStr(1, sql, SQL_UPDATE, vbTextCompare) = 1 Then GetQueryType = NT_QUERY_UPDATE: Exit Function
End Function
    
