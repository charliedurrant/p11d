VERSION 5.00
Object = "{2B9F5150-5EED-4657-82EC-33E5E52ACF54}#1.0#0"; "atc2where.OCX"
Begin VB.Form frmWhere 
   Caption         =   "TCSWhere testing"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6435
   ScaleWidth      =   7545
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdSetSql 
      Caption         =   "Set SQL"
      Height          =   465
      Left            =   2340
      TabIndex        =   13
      Top             =   4860
      Width           =   2355
   End
   Begin atc2WhereControl.tcsWhere tcsWhere1 
      Height          =   2595
      Left            =   0
      TabIndex        =   12
      Top             =   1215
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   4577
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show SQL"
      Height          =   360
      Left            =   90
      TabIndex        =   10
      Top             =   4815
      Width           =   1800
   End
   Begin VB.Frame fmeNewCriteria 
      Caption         =   "Define more criteria"
      Height          =   1185
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8445
      Begin VB.CommandButton cmdAddCondition 
         Caption         =   "Add Condition"
         Height          =   345
         Left            =   6090
         TabIndex        =   6
         Top             =   210
         Width           =   1980
      End
      Begin VB.TextBox txtValue 
         Height          =   345
         Left            =   5415
         TabIndex        =   5
         Top             =   660
         Width           =   2685
      End
      Begin VB.ComboBox cmbCondition 
         Height          =   315
         Left            =   3060
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   675
         Width           =   2250
      End
      Begin VB.ComboBox cmbColumn 
         Height          =   315
         Left            =   750
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   660
         Width           =   2190
      End
      Begin VB.OptionButton optOR 
         Caption         =   "&Or"
         Height          =   240
         Left            =   90
         TabIndex        =   2
         Top             =   750
         Width           =   540
      End
      Begin VB.OptionButton optAND 
         Caption         =   "&And"
         Height          =   315
         Left            =   90
         TabIndex        =   1
         Top             =   405
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.Label lblValue 
         Caption         =   "&Value:"
         Height          =   210
         Left            =   5415
         TabIndex        =   9
         Top             =   405
         Width           =   1230
      End
      Begin VB.Label lblCondition 
         Caption         =   "&Condition:"
         Height          =   195
         Left            =   3030
         TabIndex        =   8
         Top             =   435
         Width           =   1200
      End
      Begin VB.Label lblColumn 
         Caption         =   "Co&lumn:"
         Height          =   195
         Left            =   780
         TabIndex        =   7
         Top             =   360
         Width           =   1320
      End
   End
   Begin VB.Label lblSQL 
      Caption         =   "SQL"
      Height          =   900
      Left            =   45
      TabIndex        =   11
      Top             =   5400
      Width           =   8595
   End
End
Attribute VB_Name = "frmWhere"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private db As Database

Private Sub cmbCondition_Click()
  If Len(cmbCondition.Text) > 0 Then cmdAddCondition.Enabled = True
End Sub

Private Sub cmdAddCondition_Click()
  If optAND Then
    Call tcsWhere1.AddCondition(cmbColumn.Text, cmbCondition.ItemData(cmbCondition.ListIndex), txtValue, LOGICAL_AND)
  Else
    Call tcsWhere1.AddCondition(cmbColumn.Text, cmbCondition.ItemData(cmbCondition.ListIndex), txtValue, LOGICAL_OR)
  End If
End Sub

Private Sub cmdSetSql_Click()
  tcsWhere1.SQLLogic = lblSQL.Caption
End Sub

Private Sub cmdShow_Click()
  lblSQL.Caption = tcsWhere1.sql
End Sub

Private Sub Form_Load()
  Set db = InitDB(gwsMain, AppPath & "\" & "test.mdb", "Test Database Name")
  initColumns
End Sub
Private Function initColumns()
  Dim td As TableDef
  Dim i As Long
  
  On Error GoTo initColumns_err
  cmbColumn.Clear
  
  Set td = db.TableDefs("Contacts")
  For i = 0 To td.Fields.Count - 1
    Call cmbColumn.AddItem(td.Fields(i).Name)
    cmbColumn.ItemData(cmbColumn.NewIndex) = td.Fields(i).Type
    Call tcsWhere1.AddField(td.Fields(i).Name, DAOtoDatatype(td.Fields(i).Type))
    i = i + 1
  Next i
  cmbColumn.ListIndex = 0
  
initColumns_end:
  Exit Function
initColumns_err:
  Call ErrorMessage(ERR_ALLOWIGNORE, Err, "initColumns", "Set columns", "An error occurred setting the column choices.")
  Resume Next
End Function

Private Sub cmbColumn_Click()
  Call initConditions
End Sub

Private Function initConditions()
  Dim fldType As DataTypeEnum
  
  On Error GoTo initConditions_err
  
  cmbCondition.Clear
  cmdAddCondition.Enabled = False
  txtValue.Text = ""
  fldType = cmbColumn.ItemData(cmbColumn.ListIndex)
  Select Case fldType
    Case dbDouble, dbLong
      txtValue.Enabled = True
      cmbCondition.AddItem ("Greater than")
      cmbCondition.ItemData(cmbCondition.NewIndex) = NUM_GREATER_THAN
      cmbCondition.AddItem ("Less than")
      cmbCondition.ItemData(cmbCondition.NewIndex) = NUM_LESS_THAN
      cmbCondition.AddItem ("Equal to")
      cmbCondition.ItemData(cmbCondition.NewIndex) = NUM_EQUAL_TO
      cmbCondition.AddItem ("Greater than/Equal to")
      cmbCondition.ItemData(cmbCondition.NewIndex) = NUM_GREATER_OR_EQUAL
      cmbCondition.AddItem ("Less than/Equal to")
      cmbCondition.ItemData(cmbCondition.NewIndex) = NUM_LESS_OR_EQUAL
      cmbCondition.AddItem ("Not equal to")
      cmbCondition.ItemData(cmbCondition.NewIndex) = NUM_NOT_EQUAL
      cmbCondition.AddItem ("Is empty")
      cmbCondition.ItemData(cmbCondition.NewIndex) = NUM_ISEMPTY
    Case dbText
      txtValue.Enabled = True
      cmbCondition.AddItem ("Contains")
      cmbCondition.ItemData(cmbCondition.NewIndex) = STR_CONTAINS
      cmbCondition.AddItem ("Begins with")
      cmbCondition.ItemData(cmbCondition.NewIndex) = STR_BEGINS
      cmbCondition.AddItem ("Ends with")
      cmbCondition.ItemData(cmbCondition.NewIndex) = STR_ENDS
      cmbCondition.AddItem ("Equals")
      cmbCondition.ItemData(cmbCondition.NewIndex) = STR_EQUALS
      cmbCondition.AddItem ("Does not include")
      cmbCondition.ItemData(cmbCondition.NewIndex) = STR_NOT_INCLUDE
      cmbCondition.AddItem ("Is empty")
      cmbCondition.ItemData(cmbCondition.NewIndex) = STR_ISEMPTY
    Case dbDate
      txtValue.Enabled = True
      cmbCondition.AddItem ("Is on")
      cmbCondition.ItemData(cmbCondition.NewIndex) = DT_ON
      cmbCondition.AddItem ("Is after")
      cmbCondition.ItemData(cmbCondition.NewIndex) = DT_AFTER
      cmbCondition.AddItem ("Is before")
      cmbCondition.ItemData(cmbCondition.NewIndex) = DT_BEFORE
      cmbCondition.AddItem ("is not on")
      cmbCondition.ItemData(cmbCondition.NewIndex) = DT_NOT_ON
      cmbCondition.AddItem ("Is empty")
      cmbCondition.ItemData(cmbCondition.NewIndex) = DT_ISEMPTY
    Case dbBoolean
      txtValue.Enabled = False
      cmbCondition.AddItem ("Is true")
      cmbCondition.ItemData(cmbCondition.NewIndex) = BOOL_TRUE
      cmbCondition.AddItem ("Is false")
      cmbCondition.ItemData(cmbCondition.NewIndex) = BOOL_FALSE
    Case Else
      Call ECASE("Unsupported field type")
  End Select
  
initConditions_end:
  Exit Function
initConditions_err:
  Call ErrorMessage(ERR_ERROR, Err, "initConditions", "Load conditions", "An error occurred loading a list of conditions for column '" & cmbColumn.Text & "'.")
  Resume Next
End Function

