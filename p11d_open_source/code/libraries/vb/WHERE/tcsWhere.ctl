VERSION 5.00
Begin VB.UserControl tcsWhere 
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9045
   ScaleHeight     =   3195
   ScaleWidth      =   9045
   Begin VB.CommandButton cmdParse 
      Caption         =   "Parse"
      Height          =   330
      Left            =   7065
      TabIndex        =   6
      Top             =   1845
      Width           =   645
   End
   Begin VB.TextBox txtSQL 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      Height          =   555
      Left            =   90
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "txtSQL"
      Top             =   2520
      Visible         =   0   'False
      Width           =   8880
   End
   Begin VB.ListBox lstConditions 
      CausesValidation=   0   'False
      Height          =   1035
      Left            =   45
      TabIndex        =   1
      Top             =   315
      Width           =   8520
   End
   Begin VB.TextBox txtConditions 
      Height          =   360
      Left            =   45
      TabIndex        =   0
      Top             =   1800
      Width           =   6765
   End
   Begin VB.Label lblSQL 
      Caption         =   "SQL:"
      Height          =   270
      Left            =   135
      TabIndex        =   4
      Top             =   2250
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Label lblLogic 
      AutoSize        =   -1  'True
      Caption         =   "Logic:"
      Height          =   195
      Left            =   135
      TabIndex        =   3
      Top             =   1590
      Width           =   435
   End
   Begin VB.Label lblConditions 
      AutoSize        =   -1  'True
      Caption         =   "Conditions:"
      Height          =   195
      Left            =   90
      TabIndex        =   2
      Top             =   45
      Width           =   780
   End
End
Attribute VB_Name = "tcsWhere"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum TCSWHERE_LOGICAL_OPERATOR
  LOGICAL_NONE = 0
  LOGICAL_AND
  LOGICAL_OR
End Enum

Public Enum TCSWHERE_CONDITIONS
  CONDITION_NONE = 0
  NUM_GREATER_THAN
  NUM_LESS_THAN
  NUM_EQUAL_TO
  NUM_GREATER_OR_EQUAL
  NUM_LESS_OR_EQUAL
  NUM_NOT_EQUAL
  NUM_ISEMPTY
  STR_CONTAINS
  STR_BEGINS
  STR_ENDS
  STR_EQUALS
  STR_NOT_INCLUDE
  STR_ISEMPTY
  DT_ON
  DT_AFTER
  DT_BEFORE
  DT_NOT_ON
  DT_ISEMPTY
  BOOL_TRUE
  BOOL_FALSE
End Enum

Private mRootClause As whereClause
Private mFields As Collection

Public Sub ClearFields()
  Set mFields = New Collection
End Sub

Public Sub AddField(ByVal FieldName As String, ByVal dType As DATABASE_FIELD_TYPES)
  Dim ifld As New FieldDetails
  
  ifld.Name = FieldName
  ifld.DataType = dType
  Call mFields.Add(ifld, ifld.Name)
End Sub

Public Property Get InternalFormat() As String
  Dim s As String
  If Not mRootClause Is Nothing Then
    s = mRootClause.InternalConditions & CONDITION_SEP
    s = s & mRootClause.Internal
    InternalFormat = s
  End If
End Property

Public Property Let InternalFormat(ByVal NewValue As String)
  If Not mRootClause Is Nothing Then
    Call mRootClause.Kill
    Set mRootClause = Nothing
  End If
  NewValue = Trim$(NewValue)
  Set mRootClause = CreateClauseFromInternal(NewValue)
  Call RefreshControl
End Property

Public Property Get SQL() As String
  If Not mRootClause Is Nothing Then
    SQL = mRootClause.OutputSQL(False)
  End If
End Property

Public Property Get SQLLogic() As String
  If Not mRootClause Is Nothing Then
    SQLLogic = mRootClause.OutputSQL(True)
  End If
End Property

Private Function SetSQLLogic(NewValue As String) As Boolean
  Dim SavedFormat As String
  Dim Conds As Collection
    
  SetSQLLogic = True
  SavedFormat = Me.InternalFormat
  Set Conds = New Collection
  If Not mRootClause Is Nothing Then
    Call FillConditions(mRootClause, Conds)
    Call mRootClause.Kill
    Set mRootClause = Nothing
  End If
  NewValue = Trim$(NewValue)
  Set mRootClause = CreateConditionTreeSQL(Conds, NewValue)
  If mRootClause Is Nothing Then
    Set mRootClause = CreateClauseFromInternal(SavedFormat)
    SetSQLLogic = False
  End If
  Call RefreshControl
End Function

Public Property Let SQLLogic(ByVal NewValue As String)
  Call SetSQLLogic(NewValue)
End Property

Private Sub RefreshControl()
  If mRootClause Is Nothing Then
    lstConditions.Clear
    txtConditions = ""
    txtSQL = ""
  Else
    Call mRootClause.RebaseConditions(True)
    Call mRootClause.RebaseConditions(False)
    Call mRootClause.OutputConditionList(lstConditions, True)
    Call lstConditions.Clear
    Call mRootClause.OutputConditionList(lstConditions, False)
    txtConditions = mRootClause.OutputSQL(True)
    txtSQL = mRootClause.OutputSQL(False)
  End If
End Sub
Public Sub ClearConditions()
  If Not mRootClause Is Nothing Then
    Call mRootClause.Kill
    Set mRootClause = Nothing
  End If
  Call RefreshControl
End Sub

Public Function AddCondition(ByVal FieldName As String, ByVal Operator As TCSWHERE_CONDITIONS, ByVal Value As Variant, ByVal LogicalOperator As TCSWHERE_LOGICAL_OPERATOR)
  Dim fi As FieldDetails
  Dim cond As whereCondition, wCL As whereClause

  On Error GoTo AddCondition_err
  Set fi = mFields.Item(FieldName)
  Set cond = New whereCondition
  If cond.Init(fi, Operator, Value) Then
    Set wCL = New whereClause
    Set wCL.Value = cond
    Set mRootClause = MergeClauses(mRootClause, mRootClause, wCL, LogicalOperator)
  End If
  Call RefreshControl
    
AddCondition_end:
  Exit Function
  
AddCondition_err:
  Err.Raise Err.Number, "AddCondition", "Unable to Add condition for field " & FieldName & vbCrLf & Err.Description
End Function

Private Sub cmdParse_Click()
  Dim s As String
  
  s = txtConditions
  If Not SetSQLLogic(s) Then
    txtConditions = s
  End If
End Sub

Private Sub lstConditions_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim sItem As String, p As Long
  
  If KeyCode = vbKeyDelete Then
    If lstConditions.ListIndex >= 0 Then
      sItem = lstConditions.List(lstConditions.ListIndex)
      p = InStr(sItem, vbTab)
      If p > 0 Then
        sItem = Left$(sItem, p - 1)
        Set mRootClause = DeleteCondition(sItem, mRootClause)
        Call RefreshControl
      End If
    End If
  End If
End Sub


Private Sub UserControl_Initialize()
  
  lblConditions.Left = 90
  lstConditions.Left = 90
  lblLogic.Left = 90
  txtConditions.Left = 90
  lblConditions.Top = 45
  txtConditions.Top = 285
  
  Set mFields = New Collection
  Call RefreshControl
End Sub

Public Function FillConditionCombo(ConditionCombo As Object, ByVal FieldName As String)
  Dim fi As FieldDetails
  
  On Error GoTo FillConditionCombo_err
  Set fi = mFields.Item(FieldName)
  ConditionCombo.Clear
  Select Case fi.DataType
    Case TYPE_DOUBLE, TYPE_LONG
      ConditionCombo.AddItem ("Greater than")
      ConditionCombo.ItemData(ConditionCombo.NewIndex) = NUM_GREATER_THAN
      ConditionCombo.AddItem ("Less than")
      ConditionCombo.ItemData(ConditionCombo.NewIndex) = NUM_LESS_THAN
      ConditionCombo.AddItem ("Equal to")
      ConditionCombo.ItemData(ConditionCombo.NewIndex) = NUM_EQUAL_TO
      ConditionCombo.AddItem ("Greater than/Equal to")
      ConditionCombo.ItemData(ConditionCombo.NewIndex) = NUM_GREATER_OR_EQUAL
      ConditionCombo.AddItem ("Less than/Equal to")
      ConditionCombo.ItemData(ConditionCombo.NewIndex) = NUM_LESS_OR_EQUAL
      ConditionCombo.AddItem ("Not equal to")
      ConditionCombo.ItemData(ConditionCombo.NewIndex) = NUM_NOT_EQUAL
      ConditionCombo.AddItem ("Is empty")
      ConditionCombo.ItemData(ConditionCombo.NewIndex) = NUM_ISEMPTY
    Case TYPE_STR
      ConditionCombo.AddItem ("Contains")
      ConditionCombo.ItemData(ConditionCombo.NewIndex) = STR_CONTAINS
      ConditionCombo.AddItem ("Begins with")
      ConditionCombo.ItemData(ConditionCombo.NewIndex) = STR_BEGINS
      ConditionCombo.AddItem ("Ends with")
      ConditionCombo.ItemData(ConditionCombo.NewIndex) = STR_ENDS
      ConditionCombo.AddItem ("Equals")
      ConditionCombo.ItemData(ConditionCombo.NewIndex) = STR_EQUALS
      ConditionCombo.AddItem ("Does not include")
      ConditionCombo.ItemData(ConditionCombo.NewIndex) = STR_NOT_INCLUDE
      ConditionCombo.AddItem ("Is empty")
      ConditionCombo.ItemData(ConditionCombo.NewIndex) = STR_ISEMPTY
    Case TYPE_DATE
      ConditionCombo.AddItem ("Is on")
      ConditionCombo.ItemData(ConditionCombo.NewIndex) = DT_ON
      ConditionCombo.AddItem ("Is after")
      ConditionCombo.ItemData(ConditionCombo.NewIndex) = DT_AFTER
      ConditionCombo.AddItem ("Is before")
      ConditionCombo.ItemData(ConditionCombo.NewIndex) = DT_BEFORE
      ConditionCombo.AddItem ("is not on")
      ConditionCombo.ItemData(ConditionCombo.NewIndex) = DT_NOT_ON
      ConditionCombo.AddItem ("Is empty")
      ConditionCombo.ItemData(ConditionCombo.NewIndex) = DT_ISEMPTY
    Case TYPE_BOOL
      ConditionCombo.AddItem ("Is true")
      ConditionCombo.ItemData(ConditionCombo.NewIndex) = BOOL_TRUE
      ConditionCombo.AddItem ("Is false")
      ConditionCombo.ItemData(ConditionCombo.NewIndex) = BOOL_FALSE
    Case Else
      Call ECASE("Unsupported field type")
  End Select
  
FillConditionCombo_end:
  Exit Function
  
FillConditionCombo_err:
  Call ErrorMessage(ERR_ERROR, Err, "FillConditionCombo", "Load conditions", "An error occurred loading a list of conditions for column '" & FieldName & "'.")
  Resume Next
End Function

Private Sub UserControl_Resize()
  Static InResize As Boolean
  Const MIN_WIDTH As Single = 1500
  Const MIN_HEIGHT As Single = 2000
  Dim uHeight As Single, uWidth As Single
  
  On Error Resume Next
  If InResize Then Exit Sub
  InResize = True
  uHeight = Max(MIN_HEIGHT, UserControl.Height)
  uWidth = Max(MIN_WIDTH, UserControl.Width)
  
  lstConditions.Height = uHeight - (lstConditions.Top + 45 + lblLogic.Height + txtConditions.Height)
  lblLogic.Top = lstConditions.Height + lstConditions.Top + 80
  txtConditions.Top = lblLogic.Top + lblLogic.Height + 45
  lstConditions.Width = uWidth - 180
  txtConditions.Width = lstConditions.Width - cmdParse.Width - 350
  cmdParse.Left = txtConditions.Left + txtConditions.Width + 200
  cmdParse.Top = txtConditions.Top
  InResize = False
End Sub
