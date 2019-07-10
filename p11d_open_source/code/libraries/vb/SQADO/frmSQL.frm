VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D7D47D2E-20A1-45D1-B08B-3A509726296E}#1.0#0"; "atc2split.OCX"
Object = "{29FBC4F2-EB78-49B6-B44C-B151CF1047D7}#1.1#0"; "atc3gridado.OCX"
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
   Begin atc3GRID_ADO.AutoGridCtrl_ADO GridCtrl 
      Height          =   2580
      Left            =   2205
      TabIndex        =   11
      Tag             =   "EQUALISER,EQUALISEB"
      Top             =   3510
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   4551
      AllowUpdate     =   -1  'True
   End
   Begin VB.TextBox txtDetails 
      Height          =   825
      Left            =   2205
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   9
      Tag             =   "EQUALISER"
      Top             =   2565
      Width           =   6765
   End
   Begin VB.CheckBox chkAllowEditing 
      Caption         =   "Allow editing"
      Height          =   330
      Left            =   90
      TabIndex        =   8
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
         NumListImages   =   17
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
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSQL.frx":2850
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSQL.frx":2962
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSQL.frx":2A74
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSQL.frx":2B86
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSQL.frx":2C98
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSQL.frx":2DAA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Close"
      Height          =   420
      Left            =   8280
      TabIndex        =   5
      Tag             =   "LOCKBR"
      Top             =   6660
      Width           =   1140
   End
   Begin VB.CommandButton cmdExecute 
      Height          =   375
      Left            =   9000
      Picture         =   "frmSQL.frx":2EBC
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "LOCKR"
      Top             =   810
      Width           =   465
   End
   Begin VB.CommandButton cmdConnection 
      Caption         =   "..."
      Height          =   375
      Left            =   9000
      Picture         =   "frmSQL.frx":2FBE
      TabIndex        =   2
      Tag             =   "LOCKR"
      Top             =   180
      Width           =   465
   End
   Begin VB.TextBox txtFileName 
      BackColor       =   &H8000000F&
      Height          =   375
      Left            =   45
      TabIndex        =   1
      Tag             =   "EQUALISER"
      Top             =   180
      Width           =   8925
   End
   Begin VB.TextBox txtSQL 
      Height          =   1500
      Left            =   2205
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Tag             =   "EQUALISER"
      Top             =   810
      Width           =   6765
   End
   Begin atc2SPLIT.SPLIT split 
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
      TabIndex        =   10
      Tag             =   "EQUALISEB"
      Top             =   2385
      Width           =   1545
   End
   Begin VB.Label lblStructure 
      Caption         =   "Structure"
      Height          =   195
      Left            =   45
      TabIndex        =   7
      Tag             =   "LOCK"
      Top             =   585
      Width           =   1725
   End
   Begin VB.Label lblSQL 
      Caption         =   "Enter SQL statement"
      Height          =   240
      Left            =   2205
      TabIndex        =   6
      Tag             =   "LOCK"
      Top             =   585
      Width           =   1770
   End
   Begin VB.Label lblDataBase 
      Caption         =   "Connection string"
      Height          =   195
      Left            =   45
      TabIndex        =   4
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


Private mConnection As Connection
Private mCatalog As Catalog
Private mTarget As DATABASE_TARGET
Private mAuto As AutoClass
Private mAllowUpdate As Boolean

Private Const SQL_SELECT As String = "SELECT "
Private Const SQL_DELETE As String = "DELETE "
Private Const SQL_INSERT As String = "INSERT "
Private Const SQL_UPDATE As String = "UPDATE "
Private Const SQL_PARAMS As String = "PARAMETERS "

Private Const S_INDEXES As String = "INDEXS"

Private Const S_PROCEDURES As String = "Procedures"
Private Const S_USERS As String = "Users"
Private Const S_VIEWS As String = "Views"
Private Const S_TABLES As String = "Tables"
Private Const S_UNKNOWN As String = "Unknown"
Private Const S_NOT_SUPPORTED As String = " not supported by OLE DB provider"
Private Const S_PARAMETERS As String = "{PARAMETERS}"

Private Const SNG_SPLIT_GAP As Single = 50
Public Enum NODE_TYPE
  NT_TABLES = 1
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
  NT_PROCEDURES
  NT_INDEXES
  NT_USERS
  NT_VIEWS
  NT_PARAMETERS
  NT_NOT_SUPPORTED
End Enum

Private m_SQLDebugADO As SQLDebugADO
Private m_CR As clsFormResize

Private Sub cmdConnection_Click()
  Call OpenConnection
  
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  mAllowUpdate = True
  Call split.Initialise(Me.hWnd, False)
  
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

Private Sub OpenConnection()
  Dim sConnectionString As String
  Dim sCurrentConnectionString As String
  
  On Error GoTo OpenConnection_ERR

  
  If Not mConnection Is Nothing Then sCurrentConnectionString = mConnection.ConnectionString
  
  sConnectionString = frmConnection.ShowConnection(mTarget, sCurrentConnectionString, m_SQLDebugADO)
  Set frmConnection = Nothing
  Call InitialiseForm(sConnectionString, m_SQLDebugADO, mTarget)
  
  
OpenConnection_END:
  Exit Sub
OpenConnection_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "OpenConnection", "Open Connection", "Error in OpenConnection")
  Resume OpenConnection_END
  Resume
End Sub

Private Sub chkAllowEditing_Click()
  Call SetEdits
End Sub

Private Sub cmdFileOpen_Click()
  
End Sub

Private Sub Form_LostFocus()
  Call Me.SetFocus
End Sub

Private Sub Form_Resize()
  Call m_CR.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set m_SQLDebugADO = Nothing
  Call Kill
  Set mConnection = Nothing
End Sub

Public Function InitialiseForm(vConnection As Variant, SQLDebugADO As SQLDebugADO, ByVal Target As DATABASE_TARGET) As String
  Dim conn As Connection
  
  On Error GoTo InitialiseForm_ERR
  
  Call SetCursor
  Call Kill
  Call CreateAutoClass
  mTarget = Target
  DoEvents
  If Not SQLDebugADO Is Nothing Then Set m_SQLDebugADO = SQLDebugADO
  
  If VarType(vConnection) = vbString Then
    frmSQL.txtFileName = vConnection
    Set conn = ADOConnect(vConnection, adUseClient)
  End If
  
  If VarType(vConnection) = vbObject Then
    If Not vConnection Is Nothing Then
      If TypeOf vConnection Is Connection Then Set conn = vConnection
    End If
  End If
  If conn Is Nothing Then
    GoTo InitialiseForm_END
  Else
    frmSQL.txtFileName = conn.ConnectionString
  End If
  
  Set mConnection = conn
  InitialiseForm = mConnection.ConnectionString
  
  Call FillTreeview
    
InitialiseForm_END:
  Call ClearCursor
  Exit Function
InitialiseForm_ERR:
  tvwStructure.nodes.Clear
  Call ErrorMessage(ERR_ERROR, Err, "InitialiseForm", "Initialise Form", "Error initialising the SQLDebug form.")
  Resume InitialiseForm_END
  Resume
End Function

Private Function NodeTypeToParentKey(ByVal NT As NODE_TYPE) As String
  Select Case NT
    Case NT_TABLES
      NodeTypeToParentKey = S_TABLES
    Case NT_VIEWS
      NodeTypeToParentKey = S_VIEWS
    Case NT_INDEXES
      NodeTypeToParentKey = S_INDEXES
    Case NT_USERS
      NodeTypeToParentKey = S_USERS
    Case NT_QUERY_DELETE
      NodeTypeToParentKey = SQL_DELETE
    Case NT_QUERY_INSERT
      NodeTypeToParentKey = SQL_INSERT
    Case NT_QUERY_SELECT
      NodeTypeToParentKey = SQL_SELECT
    Case NT_QUERY_UPDATE
      NodeTypeToParentKey = SQL_UPDATE
    Case Else
      Call ECASE("Invalid NODE_TYPE =" & NT)
  End Select
End Function
Private Sub AddTables(ByVal nodes As nodes)
  Dim table As table
  Dim Tables As Tables
  Dim sql As String
  Dim n As node
  Dim nTable As node
  Dim nChild As node
  
  On Error GoTo AddTables_ERR
  
  Call SetCursor
  
  Set nTable = nodes.Add(, , S_TABLES, "Tables", NT_FOLDER_CLOSED)
  
  Set Tables = mCatalog.Tables
  
  If AddErrorNode(nodes, nTable, Tables, S_TABLES) Then GoTo AddTables_END
  
  For Each table In Tables
    If StrComp(table.Type, "SYSTEM TABLE", vbBinaryCompare) <> 0 And StrComp(table.Type, "ACCESS TABLE", vbBinaryCompare) <> 0 Then
      Set n = nodes.Add(nTable, tvwChild, , table.Name, NT_TABLES)
      If mTarget = DB_TARGET_ORACLE Then
        sql = "Select * from " & table.Name
      Else
        sql = "Select * from [" & table.Name & "]"
      End If
      
      Set nChild = nodes.Add(n, tvwChild, , sql, NT_SQL)
      nChild.Tag = sql
    End If
  Next
    
AddTables_END:
  Call ClearCursor
  Exit Sub
AddTables_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "AddTables", "Add Tables", "Error adding tables.")
  Resume AddTables_END
  Resume
End Sub

Private Sub AddFieldProperties(ByVal nodes As nodes, nParent As node)
  Dim table As table
  Dim nParentsParent As node
  Dim column As column
  Dim properties As properties
  Dim property As property

  On Error GoTo AddFieldProperties_ERR

  Call SetCursor

  If nParent.Children > 1 Then GoTo AddFieldProperties_END

  Set nParentsParent = nParent.Parent
  Set table = mCatalog.Tables(nParentsParent.Text)
  Set column = table.Columns(nParent.Text)
      
  Call nodes.Add(nParent, tvwChild, , "Defined size:" & column.DefinedSize, NT_PROPERTY)
  Call nodes.Add(nParent, tvwChild, , "Numeric scale:" & column.NumericScale, NT_PROPERTY)
  Call nodes.Add(nParent, tvwChild, , "Precision:" & column.Precision, NT_PROPERTY)
  Call nodes.Add(nParent, tvwChild, , "Type:" & TypeToString(column.Type), NT_PROPERTY)
  
  Set properties = column.properties
  
  If AddErrorNode(nodes, nParent, properties, "Extended Properties") Then GoTo AddFieldProperties_END
  
  For Each property In properties
    Call nodes.Add(nParent, tvwChild, , property.Name & ":" & PropertyValue(property), NT_PROPERTY)
  Next
  
  
AddFieldProperties_END:
  Call ClearCursor
  Exit Sub
AddFieldProperties_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "AddFieldProperties", "Add Field Properties", "Error filling the properties.")
  Resume AddFieldProperties_END
  Resume
End Sub
Private Function PropertyValue(ByVal property As property)
  On Error Resume Next
  PropertyValue = "Value" & S_NOT_SUPPORTED
  PropertyValue = property.Value
End Function
Private Function TypeToString(ByVal dt As DataTypeEnum) As String
  Dim s As String
  
  Select Case dt
    Case adTinyInt
      s = "adTinyInt"
    Case adSmallInt
      s = "adSmallInt"
    Case adInteger
      s = "adInteger"
    Case adBigInt
      s = "adBigInt"
    Case adUnsignedTinyInt
      s = "adUnsignedTinyInt"
    Case adUnsignedSmallInt
      s = "adUnsignedSmallInt"
    Case adUnsignedInt
      s = "adUnsignedInt"
    Case adUnsignedBigInt
      s = "adUnsignedBigInt"
    Case adSingle
      s = "adSingle"
    Case adDouble
      s = "adDouble"
    Case adCurrency
      s = "adCurrency"
    Case adDecimal
      s = "adDecimal"
    Case adNumeric
      s = "adNumeric"
    Case adBoolean
      s = "adBoolean"
    Case adUserDefined
      s = "adUserDefined"
    Case adVariant
      s = "adVariant"
    Case adGUID
      s = "adGuid"
    Case adDBDate
      s = "adDBDate"
    Case adDBTime
      s = "adDBTime"
    Case adDBTimeStamp
      s = "adDBTimestamp"
    Case adBSTR
      s = "adBSTR"
    Case adChar
      s = "adChar"
    Case adVarChar
      s = "adVarChar"
    Case adLongVarChar
      s = "adLongVarChar"
    Case adWChar
      s = "adWChar"
    Case adVarWChar
      s = "adVarWChar"
    Case adLongVarWChar
      s = "adLongVarWChar"
    Case adBinary
      s = "adBinary"
    Case adVarBinary
      s = "adVarBinary"
    Case adLongVarBinary
      s = "adLongVarBinary"
    Case Else
      s = S_UNKNOWN
  End Select
  
  TypeToString = s
  
End Function
Private Sub AddFields(ByVal nodes As nodes, ByVal nParent As node)
  Dim sql As String
  Dim n As node, nChild As node
  Dim Columns As Columns
  Dim table As table
  Dim column As column
  
  On Error GoTo AddFields_ERR
  
  Call SetCursor
  
  If nParent.Children > 1 Then GoTo AddFields_END
  
  
  Set table = mCatalog.Tables(nParent.Text)
    
  Set Columns = table.Columns
  
  If AddErrorNode(nodes, nParent, Columns, "Fields") Then GoTo AddFields_END
  
  For Each column In Columns
    Set n = nodes.Add(nParent, tvwChild, , column.Name, NT_FIELD)
    If mTarget = DB_TARGET_ORACLE Then
      sql = "select " & column.Name & " from " & nParent.Text
    Else
      sql = "select [" & column.Name & "] from " & "[" & nParent.Text & "]"
    End If
    Set nChild = tvwStructure.nodes.Add(n, tvwChild, , sql, NT_SQL)
    nChild.Tag = sql
  Next
  
AddFields_END:
  Call ClearCursor
  Exit Sub
AddFields_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "AddFields", "Add Fields", "Error adding fields.")
  Resume AddFields_END
  Resume
End Sub
Private Function CommandNotSupported(command As command, ByVal o As Object) As Boolean
  On Error Resume Next
  Set command = o.command
  CommandNotSupported = (command Is Nothing)
End Function

Private Sub AddParameters(ByVal nodes As nodes, ByVal nParent As node)
  Dim sql As String
  Dim Procedure As Procedure
  Dim column As column
  Dim parameter As parameter, Parameters As Parameters
  Dim command As command
  Dim nSQL As node
  Dim sParams As String, sParam As String
  
  Dim i As Long
  
  On Error GoTo AddParameters_ERR
  
  Call SetCursor
  
  If nParent.Children > 1 Then GoTo AddParameters_END
  
  Set nSQL = nParent.Child
  
  Set Procedure = mCatalog.procedures(nParent.Tag)
    
  If CommandNotSupported(command, Procedure) Then
    Call nodes.Add(nParent, tvwChild, , "Parameters " & S_NOT_SUPPORTED, NT_NOT_SUPPORTED)
    GoTo AddParameters_END
  End If
  
  Set Parameters = command.Parameters
  
  If AddErrorNode(nodes, nParent, Parameters, "Parameters") Then GoTo AddParameters_END
  
  For i = 0 To Parameters.Count - 1
    Set parameter = Parameters.Item(i)
    sParam = parameter.Name & ":" & TypeToString(parameter.Type)
    If mTarget <> DB_TARGET_JET Then
      sParams = sParams & sParam
      If i < Parameters.Count Then sParams = sParams & ","
    End If
    Call nodes.Add(nParent, tvwChild, , sParam, NT_PARAMETERS)
  Next
  
  If mTarget <> DB_TARGET_JET Then
    sql = ReplaceString(nSQL.Tag, S_PARAMETERS, sParams, vbBinaryCompare)
    nSQL.Tag = sql
    sql = ReplaceString(nSQL.Text, S_PARAMETERS, sParams, vbBinaryCompare)
    nSQL.Text = sql
  End If
  
AddParameters_END:
  Call ClearCursor
  Exit Sub
AddParameters_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "AddParameters", "Add Parameters", "Error adding parameters.")
  Resume AddParameters_END
  Resume
End Sub

Private Function CollectionNotSupported(ByVal c As Variant) As Boolean
  On Error Resume Next
  CollectionNotSupported = True
  CollectionNotSupported = (c.Count < 0)
End Function
Private Function AddErrorNode(ByVal nodes As nodes, ByVal nParent As node, ByVal c As Variant, ByVal sCollectionType As String) As Boolean
  If CollectionNotSupported(c) Then
    AddErrorNode = True
    Call nodes.Add(nParent, tvwChild, , sCollectionType & S_NOT_SUPPORTED, NT_NOT_SUPPORTED)
  End If
End Function
Private Sub AddViews(ByVal nodes As nodes)
  Dim view As view
  Dim views As views
  Dim n As node, nChild As node, nViews As node
  Dim command As command
  Dim nParent As node
  Dim sql As String
  
  On Error GoTo AddViews_ERR
  
  Call SetCursor
  
  If mTarget = DB_TARGET_JET Then GoTo AddViews_END
    
  Set nViews = nodes.Add(, , S_VIEWS, S_VIEWS, NT_FOLDER_CLOSED)
  
  Set views = mCatalog.views
  
  If AddErrorNode(nodes, nViews, views, S_VIEWS) Then GoTo AddViews_END
  
  For Each view In views
    Set n = nodes.Add(nViews, tvwChild, , view.Name, NT_VIEWS)
    Set nChild = nodes.Add(n, tvwChild, , , NT_SQL)
    If mTarget = DB_TARGET_ORACLE Then
      sql = "Select * from " & view.Name
    Else
      sql = "Select * from [" & view.Name & "]"
    End If
    
    nChild.Text = sql
    nChild.Tag = sql
  Next
    
AddViews_END:
  Call ClearCursor
  Exit Sub
AddViews_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "AddViews", "Add Views", "Error in AddViews")
  Resume AddViews_END
  Resume
End Sub
Private Sub AddUsers(ByVal nodes As nodes)
  Dim user As user
  Dim users As users
  Dim n As node, nUsers As node
  Dim nParent As node
  
  
  On Error GoTo AddUsers_ERR
  
  Call SetCursor
  
  If mTarget = DB_TARGET_JET Then Exit Sub
    
  Set nUsers = nodes.Add(, , S_USERS, S_USERS, NT_FOLDER_CLOSED)
  
  Set users = mCatalog.users
  
  If AddErrorNode(nodes, nUsers, users, S_USERS) Then GoTo AddUsers_END
  
  For Each users In users
    Set n = nodes.Add(nUsers, tvwChild, , user.Name, NT_USERS)
  Next
    
  
AddUsers_END:
  Call ClearCursor
  Exit Sub
AddUsers_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "AddUsers", "Add Users", "Error in AddUsers")
  Resume AddUsers_END
  Resume
End Sub

Private Sub AddIndexes(ByVal nodes As nodes, ByVal nParent As node)
  Dim index As index
  Dim indexes As indexes
  Dim nIndex As node
  Dim table As table
  Dim sColumn As String
  Dim property As property
  Dim properties As properties
  Dim i As Long
  
  On Error GoTo AddIndexes_ERR
  
  Call SetCursor
  
  'If mTarget = DB_TARGET_JET Then Exit Sub
    
  Set table = mCatalog.Tables(nParent.Text)
  Set indexes = table.indexes
  
  If AddErrorNode(nodes, nParent, indexes, "Indexes") Then GoTo AddIndexes_END
  
  For Each index In indexes
    Set nIndex = nodes.Add(nParent, tvwChild, , index.Name, NT_INDEXES)
    sColumn = ""
    
    For i = 0 To index.Columns.Count - 1
      sColumn = sColumn & index.Columns(i).Name
      If i < index.Columns.Count - 1 Then
        sColumn = sColumn & ","
      End If
    Next
    If Len(sColumn) Then sColumn = "(" & sColumn & ")"
    
    Call nodes.Add(nIndex, tvwChild, , "Columns:" & sColumn, NT_PROPERTY)
    Call nodes.Add(nIndex, tvwChild, , "Clustered:" & index.Clustered, NT_PROPERTY)
    Call nodes.Add(nIndex, tvwChild, , "IndexNulls:" & index.IndexNulls, NT_PROPERTY)
    Call nodes.Add(nIndex, tvwChild, , "PrimaryKey:" & index.PrimaryKey, NT_PROPERTY)
    
    
    Set properties = index.properties
    
    If AddErrorNode(nodes, nIndex, properties, "Properties") Then GoTo AddIndexes_END
    
    For Each property In properties
      Call nodes.Add(nIndex, tvwChild, , property.Name & ":" & PropertyValue(property), NT_PROPERTY)
    Next
    
    
  Next
    
  
AddIndexes_END:
  Call ClearCursor
  
  Exit Sub
AddIndexes_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "AddIndexes", "Add Indexes", "Error in AddIndexes")
  Resume AddIndexes_END
  Resume
End Sub

Private Sub AddProcedures(ByVal nodes As nodes)
  Dim Procedure As Procedure
  Dim procedures As procedures
  Dim sql As String
  Dim n As node, nChild As node, nProcedures As node
  Dim command As command
  Dim NodeImage As NODE_TYPE
  Dim sName As String
  
  On Error GoTo AddProcedures_ERR
  
  Call SetCursor
  
  Set procedures = mCatalog.procedures
  
  If mTarget = DB_TARGET_JET Then
    'jet is fully adox compliant
    Call nodes.Add(, , SQL_SELECT, "Select queries", NT_FOLDER_CLOSED)
    Call nodes.Add(, , SQL_INSERT, "Insert queries", NT_FOLDER_CLOSED)
    Call nodes.Add(, , SQL_DELETE, "Delete queries", NT_FOLDER_CLOSED)
    Call nodes.Add(, , SQL_UPDATE, "Update queries", NT_FOLDER_CLOSED)
    Call nodes.Add(, , S_UNKNOWN, "Other queries", NT_FOLDER_CLOSED)
        
    For Each Procedure In procedures
      Set command = Procedure.command
      sql = command.CommandText
      NodeImage = GetQueryType(sql)
      Set n = nodes.Add(nodes(NodeTypeToParentKey(NodeImage)), tvwChild, , Procedure.Name, NodeImage)
      n.Tag = Procedure.Name
      Set nChild = nodes.Add(n, tvwChild, , ReplaceString(sql, vbCrLf, " "), NT_SQL)
      nChild.Tag = sql
    Next
    
  Else
    Set nProcedures = nodes.Add(, , S_PROCEDURES, S_PROCEDURES, NT_FOLDER_CLOSED)
    
    If AddErrorNode(nodes, nProcedures, procedures, S_PROCEDURES) Then GoTo AddProcedures_END
    
    For Each Procedure In procedures
      Set n = nodes.Add(nProcedures, tvwChild, , , NT_PROCEDURES)
      
      Select Case mTarget
        Case DATABASE_TARGET.DB_TARGET_ORACLE
          sName = Procedure.Name
          
           sql = "BEGIN " & vbCrLf & _
                 sName & "(" & S_PARAMETERS & ");" & vbCrLf & _
                 "END;"
                 
        Case DATABASE_TARGET.DB_TARGET_SQLSERVER
           sName = Left(Procedure.Name, InStr(1, Procedure.Name, ";") - 1)
          sql = "EXEC [" & sName & "] " & S_PARAMETERS
      End Select
      
      n.Text = sName
      n.Tag = Procedure.Name
      Set nChild = nodes.Add(n, tvwChild, , ReplaceString(sql, vbCrLf, " "), NT_SQL)
      nChild.Tag = sql
    Next
  End If
    
AddProcedures_END:
  Call ClearCursor
  Exit Sub
AddProcedures_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "AddProcedures", "Add Procedures", "Error adding procedures.")
  Resume AddProcedures_END
  Resume
End Sub


Private Sub FillTreeview()
  Dim nodes As nodes
  
  On Error GoTo FillTreeview_err
  
  Call SetCursor
  
  mAllowUpdate = False
  tvwStructure.Enabled = False
  tvwStructure.nodes.Clear
  
  If mConnection Is Nothing Then GoTo FillTreeview_end
  
  Set nodes = tvwStructure.nodes
  Set mCatalog = New Catalog
  Set mCatalog.ActiveConnection = mConnection
  
  Call AddTables(nodes)
  Call AddProcedures(nodes)
  Call AddViews(nodes)
  Call AddUsers(nodes)
  
FillTreeview_end:
  Me.tvwStructure.Enabled = True
  mAllowUpdate = True
  Call ClearCursor
  Exit Sub
FillTreeview_err:
  Call ErrorMessage(ERR_ERROR, Err, "FillTreeview", "Fill Treeview", "Error filling tree view with Table/Query details.")
  Resume FillTreeview_end
  Resume
End Sub


Private Sub ExecuteSQL()
  Dim status As String
  Dim sql As String, rs As Recordset

  On Error GoTo ExecuteSQL_err
  
  Call SetCursor
  
  Me.txtDetails = ""
    
  status = "Results:" & vbCrLf
  
  If txtSQL.SelLength > 0 Then
    sql = Trim$(txtSQL.SelText)
  Else
    sql = Trim$(Me.txtSQL)
  End If
  
  
  If Not mAuto Is Nothing Then
    Call Kill
    Call CreateAutoClass
  End If
  
  If Len(sql) > 0 Then
    Set rs = mConnection.Execute(sql)
    If Not rs Is Nothing Then
      If rs.Fields.Count > 0 Then
        Call DisplayGrid(rs)
        status = status & "Recordset displayed"
      Else
        status = status & "Statement executed"
      End If
      
    End If
  Else
    status = status & "SQL string not specified"
  End If

ExecuteSQL_end:
  Me.txtDetails = status
  Call ClearCursor
  Exit Sub
ExecuteSQL_err:
  
  If Not mConnection Is Nothing Then
    status = GetADOError(Err, mConnection)
  Else
    status = Err.Description
  End If
  status = "Error:" & vbCrLf & status
  Resume ExecuteSQL_end
End Sub

Private Sub DisplayGrid(rs As Recordset)
  On Error GoTo DisplayGrid_ERR
  
  If mAuto.grid Is Nothing Then
    If Not rs Is Nothing Then
     Call mAuto.InitAutoData("DebugSQLAuto", rs, Me.GridCtrl)
    End If
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

Private Sub tvwStructure_Collapse(ByVal node As MSComctlLib.node)
  If node.Image = NT_FOLDER_OPEN Then node.Image = NT_FOLDER_CLOSED
End Sub

Private Sub tvwStructure_Expand(ByVal node As MSComctlLib.node)
  Select Case node.Image
   'CAD
    Case NT_TABLES
      Call AddFields(Me.tvwStructure.nodes, node)
      Call AddIndexes(Me.tvwStructure.nodes, node)
    Case NT_FIELD
      Call AddFieldProperties(Me.tvwStructure.nodes, node)
    Case NT_PROCEDURES, NT_QUERY_DELETE, NT_QUERY_INSERT, NT_QUERY_SELECT, NT_QUERY_UPDATE
      Call AddParameters(Me.tvwStructure.nodes, node)
    Case NT_FOLDER_CLOSED
      node.Image = NT_FOLDER_OPEN
  End Select
End Sub

Private Sub tvwStructure_NodeClick(ByVal node As MSComctlLib.node)
  If node.Image = NT_SQL Then txtSQL = node.Tag
End Sub

Private Sub txtFileName_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then Call InitialiseForm(txtFileName, Nothing, mTarget)
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
    
