VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AutoGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'enums
Public Enum COLUMN_VALUES
  COLUMN_CAPTION
  COLUMN_CAPTION_CLEAN
  COLUMN_DATAFIELD
  COLUMN_HIDE
End Enum

Public Enum COPY_TYPE
  CT_UNKNOWN = 0
  CT_ROW
  CT_COL
  CT_SEL
  
End Enum

Public Enum POPUP_MENU_ACTIONS
  SORT_ASC
  SORT_DESC
  MOVE_COLUMN
  INSERT_COLUMN
  
  MENU_COPY
  MENU_COPYPASTE_ROW
  MENU_CUT
  MENU_PASTE
  MENU_FIND_INCOL
  AUTOSIZE_COLUMN
  AUTOSIZEALL_COLUMNS
  
  FILTER_SELECTION_INC
  FILTER_SELECTION_EX
  FILTER_WIZARD
  FILTER_REMOVE
  
  COMMIT_EDIT
  MENU_DEBUG
End Enum

Public Enum POPUP_MENU_SHOWFLAGS
  SHOW_ALWAYS = 1
  VALID_ROW = 4
  VALID_ROWSELECTION = 6
  VALID_VISIBLE = &H100000
  VALID_MASK = &HFFFFF
End Enum

Private Enum GRID_MESSAGES
  MSG_REFRESH_GRID = 128
  MSG_KEYPRESS_DELAY = 129
  MSG_COMBO_RESELECT = 130
  MSG_RESTORE_GRID_POSITION = 131
  MSG_FORCE_REFRESH_GRID = 132
End Enum

Public Enum AUTOWIDTH_SETTINGS
  AW_INCLUDE_CAPTIONS = 1
  AW_RESIZE_NOSMALLER = 2
End Enum


Private mUpdateType As GRIDEDIT_TYPE
  
'Constants
Private Const DTAGRID_SCALE As Single = 0.8

'Variables

Public Event OnColumnChange(ByVal AutoName As String, ByVal FieldName As String)
Public Event ButtonClick(ByVal AutoName As String, ByVal FieldName As String)
Public Event ValidateRow(Cancel As Boolean, ValidationError As String, ByVal AutoName As String, ByVal AutoColumns As Collection)

Public ParentAC As AutoClass
Private m_AutoGrid_ADO As Object
Private WithEvents ADODataGrid As Adodc
Attribute ADODataGrid.VB_VarHelpID = -1
Private WithEvents grid As TDBGrid
Attribute grid.VB_VarHelpID = -1

Private m_NewRowBMK As Variant
Private m_NewRowRefresh As Boolean
Private m_DelayedKeyPressEvent  As Boolean

Private m_GridAllowAddNew As Boolean
Private m_GridAllowUpdate As Boolean
Private m_GridAllowDelete As Boolean

Private m_GridActive As Boolean
Private m_ShowValidateErrors As Boolean
Private m_SkipError As Boolean
Private m_ForcePopError As Boolean
Private m_InDisplayError As Boolean
Private m_UnboundColumns As Boolean
Private m_InGridRefresh As Boolean

Private m_Popped As Boolean

Private m_SaveComboValue As String

Private m_LastEditControl As Long

' Default values used when cell is blanked out/ ADO providers need valid values
Private m_BlankCellAsNull As Boolean

' Cut/Paste column ReOrder
Private StartCol As Long
Private EndCol As Long
Private m_DisplayOrder() As Long

Private mNotifyIF As IBaseNotify
Private mNotifyDisable As Boolean

Private m_PastePrivateFormat As Boolean

' OnDelete Restore position
Private m_FirstRowBMK As Variant
Private m_RowBMK As Variant

'Auto Width helpers
Private mDropDownWidth  As Single
Private mCheckBoxWidth As Single
Private mAutoWidthSettings As AUTOWIDTH_SETTINGS

'Support for Audit interface
Private mAuditIF As IAutoAuditRecordADO
Private mAuditRS As Recordset
Private mAuditRSFields() As Variant
Private mAuditRSValues() As Variant
Private mAuditType  As GRIDEDIT_TYPE
Private m_CopyAllColumnValues As Boolean

Private Sub Class_Initialize()
  m_InGridRefresh = False
  m_ShowValidateErrors = True
  mNotifyDisable = False
  m_DelayedKeyPressEvent = True
  m_BlankCellAsNull = True
  m_NewRowRefresh = True
  Call ResetState
  Debug.Print "  AutoGrid init"
End Sub

Private Sub Class_Terminate()
  m_GridActive = False
  If Not ParentAC Is Nothing Then ParentAC.Kill
  Set ParentAC = Nothing
  Set m_AutoGrid_ADO = Nothing
  Set grid = Nothing
  Debug.Print "  AutoGrid terminate"
End Sub

Friend Sub ResetState()
  m_LastEditControl = -2 ' not true or false
End Sub

Friend Sub Kill()
  Set ParentAC = Nothing
  Set GridControl = Nothing
  Set AuditInterface = Nothing
  Set NotifyInterface = Nothing

  Set m_AutoGrid_ADO = Nothing
  Set ADODataGrid = Nothing
  Set grid = Nothing
  Set mNotifyIF = Nothing

  ' Support for Audit interface
  Set mAuditIF = Nothing
  Set mAuditRS = Nothing
  m_GridActive = False
End Sub

Private Sub TryGridSetFocus()
  On Error Resume Next
  Call grid.SetFocus
End Sub

Friend Property Get AuditInterface() As IAutoAuditRecordADO
  Set AuditInterface = mAuditIF
End Property

Friend Property Set AuditInterface(ByVal NewValue As IAutoAuditRecordADO)
  Set mAuditIF = NewValue
End Property

Friend Property Get NotifyInterface() As IBaseNotify
  Set NotifyInterface = mNotifyIF
End Property

Friend Property Set NotifyInterface(ByVal NewValue As IBaseNotify)
  Set mNotifyIF = NewValue
End Property

Public Property Set GridControl(ByVal NewValue As Object)
  Set m_AutoGrid_ADO = Nothing
  If Not NewValue Is Nothing Then
    If StrComp(TypeName(NewValue), "AutoGridCtrl_ADO") = 0 Then Set m_AutoGrid_ADO = NewValue
  End If
  If m_AutoGrid_ADO Is Nothing Then
    Set ADODataGrid = Nothing
    Set grid = Nothing
  Else
    Call SetGrid
  End If
End Property

Public Property Get GridControl() As Object
  Set GridControl = m_AutoGrid_ADO
End Property

Public Property Get TDBGrid() As Object
  Set TDBGrid = grid
End Property

Public Property Get Recordset() As Recordset
  Set Recordset = ADODataGrid.Recordset
End Property

Public Property Get ShowValidationErrors() As Boolean
  ShowValidationErrors = m_ShowValidateErrors
End Property

Public Property Let ShowValidationErrors(ByVal NewValue As Boolean)
  m_ShowValidateErrors = NewValue
End Property

Public Property Let BlankCellAsNull(ByVal NewValue As Boolean)
  Dim ColSet As TrueOleDBGrid60.Column
  
  m_BlankCellAsNull = NewValue
  If Not grid Is Nothing Then
    For Each ColSet In grid.Columns
      ColSet.ConvertEmptyCell = IIf(m_BlankCellAsNull, 1, 0)
    Next ColSet
  End If
End Property

Private Sub CheckGridDate()
  On Error GoTo CheckGridDate_err
  If CDate("01/06/1999") <> TryConvertDate("01/06/1999") Then Err.Raise ERR_GRIDDATES, "CheckGrisDate", "The Auto grid expects the machine date format to be DD/MM/YYYY"
CheckGridDate_end:
  Exit Sub
  
CheckGridDate_err:
  Call ErrorMessage(ERR_ERROR + ERR_ALLOWIGNORE, Err, "InitGrid", "Auto Grid Date warning", "Machine date format incorrect")
  Resume CheckGridDate_end
End Sub

Public Sub InitGrid()
  Dim rs As Recordset, dFld As field
  Dim acol As AutoCol, dUpdateable As Boolean
  Dim i As Long, ComboCount As Long
  
  On Error GoTo InitGrid_err
  StartCol = -1
  m_GridActive = False
  m_UnboundColumns = False
  If ParentAC.CheckGridDate Then Call CheckGridDate
  dUpdateable = IsDataUpdateable()
  If Not SetGrid Then Err.Raise ERR_SETGRID, "InitGrid", "Auto grid type is not recognised. TCSADORS and TCSGRID are out of sync"
  
  If Not dUpdateable Then
    grid.AllowUpdate = False
    grid.AllowAddNew = False
    grid.AllowDelete = False
  End If
  m_GridAllowAddNew = grid.AllowAddNew
  m_GridAllowUpdate = grid.AllowUpdate
  m_GridAllowDelete = grid.AllowDelete
  If ParentAC.Count > 0 Then ReDim m_DisplayOrder(1 To ParentAC.Count) As Long
  For i = 1 To ParentAC.Count
    Set acol = ParentAC.Item(i)
    m_DisplayOrder(i) = CLng(GetIniEntry("Order_" & ParentAC.AutoName, CStr(i), "0"))
    If (m_DisplayOrder(i) < 1) Or (m_DisplayOrder(i) > ParentAC.Count) Then m_DisplayOrder(i) = 0
    acol.ColWidth = CLng(GetIniEntry(ParentAC.AutoName, acol.DataField, "-1"))
    If (acol.DropType And DROPDOWN_BOUND) = DROPDOWN_BOUND Then acol.QueryCombo = m_AutoGrid_ADO.AddDataCombo
    If acol.UnboundColumn Then m_UnboundColumns = True
  Next i
  Call EnsureDisplayOrderValid
  Call ReBindColumns
  Call GetPrimaryKeyADO(ParentAC.AutoColumns, ADODataGrid.Recordset)
  Call SetAutoWidthDefaults
  Exit Sub
  
InitGrid_err:
  ParentAC.SaveColumns = False
  Err.Raise Err.Number, ErrorSource(Err, "InitGrid"), Err.Description
End Sub

Public Property Let AutoWidthSetting(ByVal NewValue As AUTOWIDTH_SETTINGS)
  mAutoWidthSettings = NewValue
End Property

Private Sub SetAutoWidthDefaults()
  Dim pic As PictureBox
  
  Set pic = m_AutoGrid_ADO.picTest
  mDropDownWidth = pic.ScaleX(GetSystemMetrics(SM_CXHSCROLL), vbPixels, pic.ScaleMode)
  mCheckBoxWidth = 270
  If mAutoWidthSettings = 0 Then mAutoWidthSettings = AW_INCLUDE_CAPTIONS
End Sub

Public Sub AutoWidth(Optional ByVal ColIndex As Integer = -1)
  Dim pic As PictureBox
  Dim ColWidths() As Single, cWidth As Single
  Dim ColSet As TrueOleDBGrid60.Column, acol As AutoCol
  Dim vbmk As Variant
  Dim iRow As Long, j As Long, k As Long
  Const COLWIDTH_OFFSET As Single = (GRID_MINCOLWIDTH * 2)
  
  On Error GoTo AutoWidth_err
  Call SetCursor
  If grid.Columns.Count = 0 Then GoTo AutoWidth_end
  Call CreateGridSplit(False, False)
  Set pic = m_AutoGrid_ADO.picTest
  
  If ColIndex >= 0 Then
    ReDim ColWidths(ColIndex To ColIndex)
  Else
    ReDim ColWidths(0 To (grid.Columns.Count - 1))
  End If
    
  For k = 0 To (grid.Splits.Count - 1)
    grid.Split = k
    For j = 0 To (grid.Columns.Count - 1)
      Set ColSet = grid.Columns(j)
      If ColSet.visible Then
        If GetAColByGridIndex(j, acol) Then
          If Not acol.Hide Then
            Set pic.Font = ColSet.Font
            If (ColIndex = j) Or (ColIndex < 0) Then
              cWidth = 0
              If mAutoWidthSettings And AW_INCLUDE_CAPTIONS Then cWidth = pic.TextWidth(ColSet.Caption)
              If mAutoWidthSettings And AW_RESIZE_NOSMALLER Then
                If (ColSet.Width - COLWIDTH_OFFSET) > cWidth Then cWidth = ColSet.Width - COLWIDTH_OFFSET
              End If
              ColWidths(j) = cWidth
            End If
          End If
        End If
      End If
    Next j
  Next k
  
  Set ColSet = Nothing
  If ColIndex >= 0 Then Set ColSet = grid.Columns(ColIndex)
      
  If Not ColSet Is Nothing Then Set pic.Font = ColSet.Font
  ' retrieve First Grid bookmark
  iRow = 0
  Do While iRow < grid.VisibleRows
    vbmk = grid.RowBookmark(iRow)
    If Len(vbmk) = 0 Then GoTo NEXT_ROW
    If ColIndex >= 0 Then
      If GetAColByGridIndex(ColIndex, acol) Then
        If acol.IsCheckBox Then
          cWidth = mCheckBoxWidth
        Else
          cWidth = pic.TextWidth(Trim$(ColSet.CellText(vbmk)))
          If acol.DropType <> DROPDOWN_NONE Then cWidth = cWidth + mDropDownWidth
          If cWidth > ColWidths(ColIndex) Then ColWidths(ColIndex) = cWidth
        End If
      End If
    Else
      For k = 0 To (grid.Splits.Count - 1)
        grid.Split = k
        For j = 0 To grid.Columns.Count - 1
          Set ColSet = grid.Columns(j)
          If ColSet.visible Then
            If GetAColByGridIndex(j, acol) Then
              If Not acol.Hide Then
                Set pic.Font = ColSet.Font
                If acol.IsCheckBox Then
                  cWidth = mCheckBoxWidth
                Else
                  cWidth = pic.TextWidth(Trim$(ColSet.CellText(vbmk)))
                  If acol.DropType <> DROPDOWN_NONE Then cWidth = cWidth + mDropDownWidth
                End If
                If cWidth > ColWidths(j) Then ColWidths(j) = cWidth
              End If
            End If
          End If
        Next j
      Next k
    End If
NEXT_ROW:
    iRow = iRow + 1
  Loop
  
  For k = 0 To (grid.Splits.Count - 1)
    grid.Split = k
    For j = 0 To grid.Columns.Count - 1
      Set ColSet = grid.Columns(j)
      If ColSet.visible Then
        If (ColIndex = j) Or (ColIndex < 0) Then
          If ColWidths(j) > 0 Then
            ColSet.Width = ColWidths(j) + COLWIDTH_OFFSET
            ColSet.visible = True
          End If
        End If
      End If
    Next j
  Next k
AutoWidth_end:
  Call CreateGridSplit(False, True)
  Call ClearCursor
  Exit Sub
  
AutoWidth_err:
  Resume AutoWidth_end
End Sub

Private Function SetGrid() As Boolean
  On Error Resume Next
  If grid Is Nothing And (Not m_AutoGrid_ADO Is Nothing) Then
    Set grid = m_AutoGrid_ADO.grid
    Set ADODataGrid = m_AutoGrid_ADO.GridDataControl
    If Not AutoControlRegistered Then Call RegisterLibrary(m_AutoGrid_ADO)
    AutoControlRegistered = True
  End If
  SetGrid = Not (grid Is Nothing)
End Function

Private Sub EnsureDisplayOrderValid()
  Dim i As Long, j As Long, IndexUsed() As Boolean
  
  ' make sure each column is only referenced once
  If ParentAC.Count <= 0 Then Exit Sub
  ReDim IndexUsed(1 To ParentAC.Count) As Boolean
  For i = 1 To ParentAC.Count
    If m_DisplayOrder(i) <> 0 Then
      For j = (i + 1) To ParentAC.Count
        If m_DisplayOrder(i) = m_DisplayOrder(j) Then m_DisplayOrder(j) = 0
      Next j
      IndexUsed(m_DisplayOrder(i)) = True
    End If
  Next i
  
  ' make sure every column is references
  For i = 1 To ParentAC.Count
    If Not IndexUsed(i) Then
      If m_DisplayOrder(i) = 0 Then
        m_DisplayOrder(i) = i
      Else     ' Find First Unused index + use
        For j = 1 To ParentAC.Count
          If m_DisplayOrder(j) = 0 Then
            m_DisplayOrder(j) = i
            Exit For
          End If
        Next j
      End If
      IndexUsed(i) = True
    End If
  Next i
End Sub

Private Sub ADODataGrid_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  Dim cur As Long, max As Long
  On Error Resume Next
  cur = pRecordset.AbsolutePosition
  max = pRecordset.RecordCount
  If cur = adPosUnknown Then
    cur = 0: max = 0
  ElseIf cur = adPosBOF Then
    cur = 0
  ElseIf cur = adPosEOF Then
    cur = max
  End If
  ADODataGrid.Caption = CStr(cur) & " of " & CStr(max)
  If (Not mNotifyIF Is Nothing) And (Not mNotifyDisable) Then Call mNotifyIF.Notify(cur, max, "Reposition")
End Sub

'Private Sub RestoreColumnFormat(ByVal ColIndex As Integer)
'  Dim aCol As AutoCol, ColSet As TrueOleDBGrid60.Column
'
'  If GetAColByGridIndex(ColIndex, aCol) Then
'    If (aCol.dbDataType = TYPE_DOUBLE) Or (aCol.dbDataType = TYPE_LONG) Then
'      Set ColSet = Grid.Columns(ColIndex)
'      ColSet.NumberFormat = aCol.Format
'    End If
'  End If
'End Sub

Private Sub Grid_AfterColUpdate(ByVal ColIndex As Integer)
  Dim acol As AutoCol
  
  'Call RestoreColumnFormat(ColIndex)
  If GetAColByGridIndex(ColIndex, acol) Then
    If acol.OnChangeEvent Then
      RaiseEvent OnColumnChange(ParentAC.AutoName, acol.DataField)
    End If
  End If
End Sub

Private Sub CheckDropDown(ByVal acol As AutoCol)
  Dim sField As String
  On Error GoTo CheckDropDown_err
  
  If acol Is Nothing Then Call GetAColByGridIndex(grid.Col, acol)
  sField = "Unknown"
  If IsDataUpdateable() And acol.RefreshRequired Then
    sField = acol.GridCaptionClean
    acol.RefreshRequired = Not RefreshSingleDropDown(acol, False)
  End If
  Exit Sub
  
CheckDropDown_err:
  Call ErrorMessage(ERR_ERROR, Err, "CheckDropDown", "Failed to refresh field dropdown", "Unable to refresh dropdown in field " & sField)
  If Not acol Is Nothing Then acol.RefreshRequired = True
End Sub

Private Sub Grid_ButtonClick(ByVal ColIndex As Integer)
  Dim acol As AutoCol
  
  If m_GridActive Then
    If GetAColByGridIndex(ColIndex, acol) Then
      Call CheckDropDown(acol)
      If acol.HasButton Then RaiseEvent ButtonClick(ParentAC.AutoName, acol.DataField)
    End If
  End If
End Sub

'Private Sub Grid_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
'  Dim aCol As AutoCol, ColSet As TrueOleDBGrid60.Column
'
'  If GetAColByGridIndex(ColIndex, aCol) Then
'    If (aCol.dbDataType = TYPE_DOUBLE) Or (aCol.dbDataType = TYPE_LONG) Then
'      Set ColSet = Grid.Columns(ColIndex)
'      ColSet.NumberFormat = ""
'      ColSet.Text = ColSet.Value
'    End If
'  End If
'End Sub

Private Function UpdateGrid(Optional ByVal CheckOnLoseFocus As Boolean = False) As Boolean
  On Error GoTo UpdateGrid_err
  
  UpdateGrid = True
  If grid.DataChanged And Not m_InDisplayError Then
    If CheckOnLoseFocus And Not ParentAC.UpdateOnLoseFocus Then Exit Function
    Call grid.Update
  End If
  Exit Function
  
UpdateGrid_err:
  UpdateGrid = False
  Call DisplayGridError
End Function

Private Sub Grid_LostFocus()
  Static InLoseFocus As Boolean
  
  If InLoseFocus Then Exit Sub
  InLoseFocus = True
  Call UpdateGrid(True)
  
  InLoseFocus = False
  Exit Sub
End Sub

Private Sub Grid_OnAddNew()
  If m_NewRowRefresh Then m_NewRowBMK = ADODataGrid.Recordset.Bookmark
  Call SetGridDefaults   'Call Grid.Refresh causes ghost dates to appear 31/12/99
End Sub

Private Function IsValidGridBookmark(ByVal vbmk As Variant) As Boolean
  On Error GoTo IsValidGridBookmark_err
  grid.Bookmark = m_NewRowBMK
  IsValidGridBookmark = True
  Exit Function
  
IsValidGridBookmark_err:
  IsValidGridBookmark = False
End Function

Private Sub RefreshCurrentRow()
  Dim cbmk As Variant
  Dim anMode As AddNewModeConstants
  
  On Error GoTo RefreshCurrentRow_err
  If Not m_NewRowRefresh Then Exit Sub
  If IsEmpty(m_NewRowBMK) Then
    Call ECASE("RefreshCurrentRow called where NewRowBMK is empty")
  Else
    anMode = grid.AddNewMode
    cbmk = grid.Bookmark
    If IsValidGridBookmark(m_NewRowBMK) Then
      grid.Bookmark = m_NewRowBMK
      Call ADODataGrid.Recordset.Resync(adAffectCurrent, adResyncAllValues)
      If anMode = dbgAddNewCurrent Then
        grid.MoveLast
        grid.CurrentCellVisible = True
        grid.Row = grid.Row + 1
      ElseIf Not IsNull(cbmk) Then
        grid.Bookmark = cbmk
      End If
    End If
    Call TryGridSetFocus
    m_NewRowBMK = Empty
  End If
  Exit Sub
  
RefreshCurrentRow_err:
  Call ErrorMessage(ERR_ERROR, Err, "RefreshCurrentRow", "Unable to refresh Current row", "Error refreshing current row")
  Resume
End Sub

Private Sub Grid_PostEvent(ByVal MsgId As Integer)
  Dim acol As AutoCol, rOffset As Long
  
  Select Case MsgId
    Case MSG_FORCE_REFRESH_GRID
      Call RefreshGrid(True, True, False, True)
    Case MSG_REFRESH_GRID
      Call RefreshCurrentRow
    Case MSG_KEYPRESS_DELAY
      m_DelayedKeyPressEvent = False
      Call GridSpeedKey(0, True)
      m_DelayedKeyPressEvent = True
    Case MSG_COMBO_RESELECT
      grid.Text = m_SaveComboValue
      SendKeys "{ENTER}"
    Case MSG_RESTORE_GRID_POSITION
      If Not IsNull(m_FirstRowBMK) Then grid.FirstRow = m_FirstRowBMK
      If Not IsNull(m_RowBMK) Then grid.Bookmark = m_RowBMK
      grid.CurrentCellVisible = True
      Call TryGridSetFocus
  End Select
End Sub

Private Sub Grid_UnboundColumnFetch(Bookmark As Variant, ByVal Col As Integer, Value As Variant)
  Dim sColName As String
  Dim acol As AutoCol, v As Variant
    
  On Error GoTo UnboundColumnFetch_err
  If GetAColByGridIndex(Col, acol) Then
    sColName = acol.GridCaptionClean
    Value = GetCalculatedValue(Me, Nothing, acol, acol.OnUpdateCalcValue, acol.OnUpdateCalcValueType, Bookmark, True)
  End If
  Exit Sub
  
UnboundColumnFetch_err:
  Err.Raise Err.Number, ErrorSource(Err, "UnboundColumnFetch"), Err.Description
End Sub

Private Sub Grid_FetchCellTips(ByVal SplitIndex As Integer, ByVal ColIndex As Integer, ByVal RowIndex As Long, CellTip As String, ByVal FullyDisplayed As Boolean, ByVal TipStyle As StyleDisp)
  Dim acol As AutoCol
  
  CellTip = ""
  If GetAColByGridIndex(ColIndex, acol) Then
    If Len(acol.ToolTip) Then CellTip = acol.ToolTip
  End If
End Sub

Private Sub SetGridDefaults()
  Static InSetDefault As Boolean
  Dim i As Integer, acol As AutoCol, v As Variant
  Dim ColSetTo As TrueOleDBGrid60.Column

  On Error GoTo SetGridDefaults_err
  If InSetDefault Then Exit Sub
  InSetDefault = True
  For i = 1 To ParentAC.Count
    Set acol = ParentAC.Item(i)
    If acol.GridColumn >= 0 Then
      Set ColSetTo = grid.Columns(acol.GridColumn)
      If Not IsEmpty(acol.OnAddNewDefault) And ((Len(ColSetTo.Text) = 0) Or IsEmpty(ColSetTo.Value)) Then
        v = GetCalculatedValue(Me, Nothing, acol, acol.OnAddNewDefault, acol.OnAddNewDefaultType)
        ColSetTo.Value = v
        ColSetTo.Text = v
      End If
    End If
  Next i
  
SetGridDefaults_end:
  InSetDefault = False
  Exit Sub
  
SetGridDefaults_err:
  Call AbortGridEdit
  Call ErrorMessage(ERR_ERROR, Err, "SetGridDefaults", "Unable to set default values for grid", "Error setting default values")
  Resume SetGridDefaults_end
End Sub

Private Sub SetAColCellValues(ByVal rs As Recordset)
  Dim errstring As String
  Dim v As Variant
  Dim i As Long, ColSet As TrueOleDBGrid60.Column, acol As AutoCol
  
  On Error GoTo SetAColCellValues_err
  For i = 1 To ParentAC.Count
    Set acol = ParentAC.Item(i)
    acol.CellValue = Empty
    If acol.UnboundColumn Then GoTo next_column
    If (acol.GridColumn >= 0) And Not m_PastePrivateFormat Then
      Set ColSet = grid.Columns(acol.GridColumn)
      If (Not IsEmpty(ColSet.Value)) And (Len(ColSet.Value) > 0) Then
        acol.CellValue = GetGridTypedValueDefault(ColSet.Value, acol.dbDataType, Empty)
        If IsEmpty(acol.CellValue) And (((2 ^ (acol.dbDataType + 1)) And ParentAC.ValidateDataTypes) <> 0) Then
          Err.Raise ERR_INVALID_CELLVALUE, "SetAColCellValues", "The value '" & ColSet.Value & "' is not a valid entry in the column " & acol.GridCaptionClean & "." & vbCrLf & "The value must be of type " & DataTypeName(acol.dbDataType)
        End If
      End If
    Else
      If Not rs Is Nothing Then acol.CellValue = rs.Fields(acol.DataField)
    End If
next_column:
    If IsNull(acol.CellValue) Then acol.CellValue = Empty
  Next i
  If Len(errstring) > 0 Then Err.Raise ERR_INVALID_CELLVALUES, "SetAColCellValues", Err.Description
  Exit Sub
  
SetAColCellValues_err:
  If Err.Number = ERR_INVALID_CELLVALUE Then
    errstring = errstring & Err.Description & vbCrLf
    Resume next_column
  End If
  Err.Raise Err.Number, ErrorSource(Err, "SetAColCellValues"), Err.Description
End Sub

Private Sub Grid_BeforeUpdate(Cancel As Integer)
  Dim i As Integer, acol As AutoCol, aColSrc As AutoCol
  Dim ColSet As TrueOleDBGrid60.Column
  Dim FieldName As String
  Dim rs As Recordset, v As Variant
  Static iLevel As Long
     
  On Error GoTo Grid_BeforeUpdate_err
  iLevel = iLevel + 1
  Call AutoLog("BEGIN BF " & iLevel)
  m_SkipError = False
  m_ForcePopError = False
 
  ' bmk = Grid.Bookmark note grid bookmark does not!! refer to the current record in addnew buffer
  ' If IsEmpty(bmk) Then Err.Raise ERR_GRID_BEFOREUPDATE, "Grid_BeforeUpdate", "Unable to obtain a bookmark for the current row"
  Set rs = ADODataGrid.Recordset
  If rs.EditMode = adEditAdd Then
    mUpdateType = GRID_ADDNEW
  Else
    mUpdateType = GRID_EDIT
  End If
  Call SetAColCellValues(rs)
    
  ' Audit Updates - prepare record current values of recordset
  Call RecordAuditValues(rs, False, False)
  
  ' calculated values not derived + booleans
  For i = 1 To ParentAC.Count
    Set acol = ParentAC.Item(i)
    v = Empty
    If Not acol.UnboundColumn Then
      If mUpdateType = GRID_ADDNEW Then
        If Not IsEmpty(acol.OnAddNewCalcValue) Then
          If IsSimpleCalc(acol.OnAddNewCalcValueType) Then v = GetCalculatedValue(Me, Nothing, acol, acol.OnAddNewCalcValue, acol.OnAddNewCalcValueType)
        End If
      ElseIf mUpdateType = GRID_EDIT Then
        If Not IsEmpty(acol.OnUpdateCalcValue) Then
          If IsSimpleCalc(acol.OnUpdateCalcValueType) Then v = GetCalculatedValue(Me, Nothing, acol, acol.OnUpdateCalcValue, acol.OnUpdateCalcValueType)
        End If
      End If
      If Not IsEmpty(v) Then
        rs.Fields(acol.DataField) = v
        acol.CellValue = v
      End If
      
      If acol.GridColumn >= 0 Then
        v = Empty
        If acol.dbDataType = TYPE_BOOL Then
          If IsEmpty(acol.CellValue) Then
            v = acol.OnNullValue
          Else
            v = CBoolean(acol.CellValue)
          End If
        End If
        If (acol.dbDataType = TYPE_DOUBLE) Or (acol.dbDataType = TYPE_LONG) Or (acol.dbDataType = TYPE_DATE) Then
          If IsEmpty(acol.CellValue) Then v = acol.OnNullValue
        End If
        If Not IsEmpty(v) Then
          Set ColSet = grid.Columns(acol.GridColumn)
          'APF 18/9/2002 UPDATED
          If Not IsNull(v) Then ColSet.Value = v
          acol.CellValue = v
        End If
      End If
    End If
  Next i

  ' derived calculated values
  For i = 1 To ParentAC.Count
    Set acol = ParentAC.Item(i)
    v = Empty
    If Not acol.UnboundColumn Then
      If mUpdateType = GRID_ADDNEW Then
        If Not IsEmpty(acol.OnAddNewCalcValue) Then
          If Not IsSimpleCalc(acol.OnAddNewCalcValueType) Then v = GetCalculatedValue(Me, Nothing, acol, acol.OnAddNewCalcValue, acol.OnAddNewCalcValueType)
        End If
      ElseIf mUpdateType = GRID_EDIT Then
        If Not IsEmpty(acol.OnUpdateCalcValue) Then
          If Not IsSimpleCalc(acol.OnUpdateCalcValueType) Then v = GetCalculatedValue(Me, Nothing, acol, acol.OnUpdateCalcValue, acol.OnUpdateCalcValueType)
        End If
      End If
      If Not IsEmpty(v) Then
        rs.Fields(acol.DataField) = v
        acol.CellValue = v
      End If
    End If
  Next i

  ' CopyFromField - set the value of another field get Source Field from grid first
  If mUpdateType = GRID_ADDNEW Then
    For i = 1 To ParentAC.Count
      Set acol = ParentAC.Item(i)
      If Len(acol.CopyFromField) > 0 Then
        Set aColSrc = GetAColByKey(acol.CopyFromField)
        If Not aColSrc Is Nothing Then
          v = aColSrc.CellValue
        Else ' old behaviour - can be static database field apf see Bug entry 26/2/99
          v = rs.Fields(acol.CopyFromField)
        End If
        rs.Fields(acol.DataField) = v
        acol.CellValue = v
      End If
    Next i
  End If
  Call ValidateGridRow
  
Grid_BeforeUpdate_end:
  iLevel = iLevel - 1
  Call AutoLog("END BF " & iLevel)
  Exit Sub
  
Grid_BeforeUpdate_err:
  Cancel = True
  m_SkipError = True
  If ((Err.Number <> ERR_GRIDVALIDATE) Or m_ShowValidateErrors) And (Err.Number <> ERR_GRIDVALIDATEIGNORE) Then
    m_ForcePopError = True
    m_SkipError = False
    Call ErrorMessagePush(Err)
  End If
  Call AutoLog("AFTER ERROR " & iLevel & " " & m_SkipError & " " & Cancel)
  Resume Grid_BeforeUpdate_end
  Resume
End Sub

Private Sub ValidateGridRow()
  Dim Cancel As Boolean, ValidationError As String
  Dim acol As AutoCol
  Dim dType As DATABASE_FIELD_TYPES, v As Variant, i As Long
  Dim errstring  As String, dbType As DATABASE_FIELD_TYPES
  
  On Error GoTo ValidRow_DataTypes_err
  
  For i = 1 To ParentAC.Count
    Set acol = ParentAC.Item(i)
    If (acol.GridColumn >= 0) And Not acol.UnboundColumn Then
      v = acol.CellValue
      dbType = acol.dbDataType
      If Not ((dbType <> TYPE_STR) And (Len(v) = 0)) Then Call GetTypedValue(v, dbType)
      If Not IsEmpty(acol.MinValidValue) Then
        If v < acol.MinValidValue Then Err.Raise ERR_GRIDVALIDATE, "ValidateGridRow", "The Value in Column '" & acol.GridCaptionClean & _
                                                 "' must be equal to or greater than " & GetValueAsString(acol.MinValidValue, dbType)
      End If
      If Not IsEmpty(acol.MaxValidValue) Then
        If v > acol.MaxValidValue Then Err.Raise ERR_GRIDVALIDATE, "ValidateGridRow", "The Value in Column '" & acol.GridCaptionClean & _
                                                 "' must be equal to or less than " & GetValueAsString(acol.MaxValidValue, dbType)
      End If
    End If
  Next i
  RaiseEvent ValidateRow(Cancel, ValidationError, ParentAC.AutoName, ParentAC.AutoColumns)
  If Cancel Then
    If StrComp(ValidationError, "IGNORE_ERROR", vbTextCompare) = 0 Then Err.Raise ERR_GRIDVALIDATEIGNORE, Err.Source, ValidationError
    Err.Raise ERR_GRIDVALIDATE, Err.Source, ValidationError
  End If
  Exit Sub
  
ValidRow_DataTypes_err:
  If (Err.Number = ERR_GRIDVALIDATE) Or (Err.Number = ERR_GRIDVALIDATEIGNORE) Then Err.Raise Err.Number, ErrorSource(Err, "ValidateGridRow"), Err.Description
  If Not acol Is Nothing Then
    errstring = "The Column '" & acol.GridCaptionClean & "' cannot contain the value: " & IsNullEx(v, "(Null)") & vbCrLf & _
                "Column is of type " & DataTypeName(dbType)
  Else
    errstring = "Grid Row could not be validated"
  End If
  Err.Raise ERR_GRIDVALIDATE, "ValidateGridRow", errstring
End Sub

Private Function ConfirmDelete() As Boolean
  ConfirmDelete = True
  If Len(ParentAC.ConfirmDeleteText) > 0 Then
    ConfirmDelete = DisplayMessage(m_AutoGrid_ADO.ContainerForm, ParentAC.ConfirmDeleteText, "Confirm Delete", "Yes", "No")
  End If
End Function
' single record delete only
Private Sub Grid_BeforeDelete(Cancel As Integer)
  If Not ConfirmDelete() Then
    Cancel = True
    m_SkipError = True
    Exit Sub
  End If
  ' record current values of recordset
  Call RecordAuditValues(ADODataGrid.Recordset, True, False)
End Sub

Private Sub Grid_AfterDelete()
  ' record current values of recordset
  If Not mAuditIF Is Nothing Then Call FireAuditRecordUpdated(mAuditIF.AuditLevel)
End Sub

Private Sub Grid_AfterUpdate()
  Dim i As Long, ColSet As TrueOleDBGrid60.Column, acol As AutoCol
  Dim rs As Recordset, aLevel As GRIDAUDIT_LEVELS
  Dim OldNotifyDisable  As Boolean
  
  On Error GoTo AfterUpdate_err
  OldNotifyDisable = mNotifyDisable
  mNotifyDisable = True
  If Not mAuditIF Is Nothing Then ' record current values of recordset
    aLevel = mAuditIF.AuditLevel
    If (aLevel = GRIDAUDIT_FULL) Or (aLevel = GRIDAUDIT_MINIMUM) Then
      Call RecordAuditValues(ADODataGrid.Recordset, False, True)
    End If
    Call FireAuditRecordUpdated(aLevel)
  End If
  If (mUpdateType = GRID_ADDNEW) And m_NewRowRefresh Then Call grid.PostMsg(MSG_REFRESH_GRID)
  mUpdateType = GRID_NOEDIT
  mNotifyDisable = OldNotifyDisable
  Exit Sub
  
AfterUpdate_err:
  mNotifyDisable = OldNotifyDisable
  Err.Raise Err.Number, ErrorSource(Err, "AfterUpdate"), Err.Description
  Resume
End Sub

Private Sub Grid_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
  Dim acol As AutoCol
  On Error GoTo ColResize_err
    
  ParentAC.SaveColumns = True
  If GetAColByGridIndex(ColIndex, acol) Then Call AutoSizeSingleDropDown(acol)
  Exit Sub
  
ColResize_err:
  Call ErrorMessage(ERR_ERROR + ERR_ALLOWIGNORE, Err, "ColResize", "Non Fatal error resizing column", "Error in Column resize")
End Sub

Private Function GetColumnName(ByVal table As String, ByVal field As String) As String
  Dim acol As AutoCol, i As Long
  
  On Error GoTo GetColumnName_err
  GetColumnName = field
  For i = 1 To ParentAC.Count
    Set acol = ParentAC.Item(i)
    If StrComp(acol.SourceTable, table, vbTextCompare) = 0 Then
      If StrComp(acol.SourceField, field, vbTextCompare) = 0 Then
        GetColumnName = acol.GridCaptionClean
        Exit For
      End If
    End If
  Next i
  
GetColumnName_end:
  Exit Function
  
GetColumnName_err:
  Resume GetColumnName_end
End Function

Public Function GetValidationString(errstring As String) As Boolean
  Const ValidationString As String = "Valid:"
  Dim p0 As Long
  
  p0 = InStr(1, errstring, ValidationString, vbTextCompare)
  If p0 = 1 Then
    errstring = Mid$(errstring, p0 + Len(ValidationString))
    GetValidationString = True
  Else
    GetValidationString = False
  End If
End Function

Private Sub AbortGridEdit()
  Dim i As Integer
  
  On Error Resume Next
  If grid.DataChanged Then
    grid.DataChanged = False
    i = grid.Col + 1
    If i >= grid.Columns.Count Then i = 0
    grid.Col = i
    grid.Col = 0
    grid.DataChanged = False
  End If
End Sub

Private Sub DebugCurrentGridRow()
  Dim tCol As TrueOleDBGrid60.Column
  Dim s As String
  
  For Each tCol In grid.Columns
    s = tCol.Name & " = "
    If IsNull(tCol.Value) Then
      s = s & "(Null)"
    ElseIf IsEmpty(tCol.Value) Then
      s = s & "(Empty)"
    Else
      s = s & "'" & tCol.Value & "' (" & VarType(tCol.Value) & ")"
    End If
    Debug.Print s
  Next tCol
End Sub

Private Sub DisplayGridError(Optional ByVal sErrorTitle As String, Optional ByVal sErrorSource As String)
  Const FieldString As String = "field '"
  Const SetForString As String = "set for '"
  Const PK_ValidationError As String = "Index or primary key can't contain a Null value."
  Const PK_ValidationError_ora As String = "Non-nullable column cannot be updated to Null."
  Const PKdup_ValidationError As String = "The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship.  Change the data in the field or fields that contain duplicate data, remove the index, or redefine the index to permit duplicate entries and try again."
  Const PKdup_ValidationError_ora As String = "ORA-00001: unique constraint"
  Const Grid_UpdateFailedMessage As String = "Unable to complete update of Grid row. The changes you have made will not be saved without further input from you. Choose either abort edit, in which case your changes to this row will be lost, or choose Ok, in which case you can complete this row"
  Dim PKErr As Boolean, PKDupErr As Boolean
  Dim fLen As Long, i As Long
  Dim cn As Connection
  Dim errno As Long
  Dim errortext0 As String, errortext1 As String, errstring As String
  Dim field As String, table As String
  Dim pkcolumns As String
  Dim p0 As Long, p1 As Long, q As Long
  
  If m_InDisplayError Then Exit Sub
  m_InDisplayError = True
  If Not m_SkipError Then
    If Len(sErrorTitle) = 0 Then sErrorTitle = "Unable to complete update of row"
    If Len(sErrorSource) = 0 Then
      Err.Source = ErrorSource(Err, "Grid_Error")
    Else
      Err.Source = ErrorSource(Err, sErrorSource & ";Grid_Error")
    End If
  End If
  If Not m_ForcePopError Then Call ErrorMessagePush(Err)
  On Error Resume Next
  If Not m_SkipError Then
    Call DebugCurrentGridRow
    If m_ForcePopError Then
      RetryCancelCaption = "&Abort Edit"
      If ErrorMessagePop(ERR_ERROR + ERR_ALLOWCANCEL, Err, "", sErrorTitle, Grid_UpdateFailedMessage) Then
        Call AbortGridEdit
      End If
      RetryCancelCaption = ""
    Else
      If Len(grid.ErrorText) > 0 Then
        errno = 0
        errortext0 = grid.ErrorText
        errstring = errortext0 & vbCrLf
      End If
      If grid.Errors.Count > 0 Then
        errno = grid.Errors.Item(0).Number
        For i = 0 To (grid.Errors.Count - 1)
          errortext1 = grid.Errors.Item(i).Description
          If (Len(errortext1) > 0) And (StrComp(errortext0, errortext1, vbTextCompare) <> 0) Then
            errstring = errstring & vbCrLf & errortext1
          End If
        Next i
      End If
      PKErr = (StrComp(Left$(errstring, Len(PK_ValidationError)), PK_ValidationError, vbTextCompare) = 0) Or (StrComp(Left$(errstring, Len(PK_ValidationError_ora)), PK_ValidationError_ora, vbTextCompare) = 0)
      PKDupErr = (StrComp(Left$(errstring, Len(PKdup_ValidationError)), PKdup_ValidationError, vbTextCompare) = 0) Or (StrComp(Left$(errstring, Len(PKdup_ValidationError_ora)), PKdup_ValidationError_ora, vbTextCompare) = 0) Or (errno = &H80040E2F)
      
      If PKErr Or PKDupErr Then
        ' apf derived columns PK's should name calc columns
        pkcolumns = GetDerivedPKColumnCaptionsEx(ParentAC, Me, mUpdateType)
        If PKErr Then
          errstring = PK_ValidationError
          If Len(pkcolumns) > 0 Then errstring = "The following columns must contain a value:" & vbCrLf & vbCrLf & pkcolumns & vbCrLf
        Else
          errstring = PKdup_ValidationError
          If Len(pkcolumns) > 0 Then errstring = "A combination of the entries in the following columns must be unique:" & vbCrLf & vbCrLf & pkcolumns & vbCrLf
        End If
        Call ErrorMessagePop(ERR_ERROR, Err, "", sErrorTitle, "An error has occurred updating the current row." & vbCr & vbCr & errstring)
      Else
        fLen = Len(FieldString)
        p0 = InStr(1, errstring, FieldString, vbTextCompare)
        If p0 = 0 Then
          p0 = InStr(1, errstring, SetForString, vbTextCompare)
          fLen = Len(SetForString)
        End If
        If p0 > 0 Then p1 = InStr(p0 + fLen, errstring, "'", vbTextCompare)
        If (p0 > 0) And (p1 > 0) Then
          field = Mid$(errstring, p0 + fLen, p1 - p0 - fLen)
          q = InStr(field, ".")
          If q > 1 Then
            table = Left$(field, q - 1)
            field = Mid$(field, q + 1)
            field = GetColumnName(table, field)
          End If
          errstring = Left$(errstring, p0 - 1) & "Column '" & field & "'" & Mid$(errstring, p1 + 1)
          Call ErrorMessagePop(ERR_ERROR, Err, "", "Validation failure", "Unable to validate the current row." & vbCr & errstring)
        Else
          If GetValidationString(errstring) Then
            Call ErrorMessagePop(ERR_ERROR, Err, "", "Validation failure", "Unable to validate the current row." & vbCr & errstring)
          Else
            Call ErrorMessagePop(ERR_ERROR, Err, "", "Grid Error", "An error has occurred in the grid." & vbCr & errstring)
          End If
        End If
      End If
    End If
  End If
  m_SkipError = False
  m_ForcePopError = False
  m_InDisplayError = False
End Sub

Private Sub Grid_Error(ByVal DataError As Integer, Response As Integer)
  Call DisplayGridError
  Response = vbDataErrContinue
End Sub

Private Sub Grid_HeadClick(ByVal ColIndex As Integer)
  ParentAC.SaveColumns = True
End Sub

Private Sub DisplayFastKey(ByVal FastKey As String)
  If Len(FastKey) > 0 Then FastKey = "Fast Key: " & FastKey
   m_AutoGrid_ADO.FastKeyLabel = FastKey
End Sub

Private Function IsValidKeyCode(ByVal KeyCode As Integer, ByVal Shift As Integer) As Boolean
  If (KeyCode = 0) Or ((Shift And vbAltMask) = vbAltMask) Or ((Shift And vbCtrlMask) = vbCtrlMask) Then Exit Function
  If (KeyCode = vbKeyReturn) Or (KeyCode = vbKeyTab) Or (KeyCode = vbKeyUp) Or (KeyCode = vbKeyDown) Or (KeyCode = vbKeyLeft) Or (KeyCode = vbKeyRight) Then Exit Function
  IsValidKeyCode = True
End Function

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim acol As AutoCol
  Static inKeyDown As Boolean, LastKeyEsc As Boolean
  
  On Error Resume Next
  If inKeyDown Then Exit Sub
  inKeyDown = True
  If (KeyCode = vbKeyUp) Or (KeyCode = vbKeyDown) Then
    If IsComboOpen(grid.hwnd) Then
      m_SaveComboValue = ""
    Else
      If (Shift And vbAltMask) = vbAltMask Then Call CheckDropDown(Nothing)
    End If
  End If
  If KeyCode = vbKeyEscape Then
    If LastKeyEsc And grid.DataChanged Then Call AbortGridEdit
    LastKeyEsc = True
  Else
    LastKeyEsc = False
  End If
  If (Shift = 0) And (KeyCode = vbKeyDelete) Then
    Call DeleteSelection
    KeyCode = 0
  End If
  If (((Shift And vbShiftMask) = vbShiftMask) And (KeyCode = vbKeyInsert)) Or _
     (((Shift And vbCtrlMask) = vbCtrlMask) And (KeyCode = vbKeyV)) Then
    Call PasteSelection
    KeyCode = 0
  End If
  
  If (Shift And vbCtrlMask) = vbCtrlMask Then
    If (KeyCode = vbKeyInsert) Or (KeyCode = vbKeyC) Then
      Call CopySelection
      KeyCode = 0
    End If
    If (KeyCode = vbKeyDelete) Or (KeyCode = vbKeyX) Then
      Call CutSelection
      KeyCode = 0
    End If
  End If
  If (Shift = 0) And (KeyCode = vbKeyF5) And ParentAC.EnableF5Copy Then
    If UpdateGrid Then Call CopyPasteGridValue(ADODataGrid.Recordset, grid, ParentAC.AutoColumns)
    KeyCode = 0
  End If
  If (Shift = 0) And (KeyCode = vbKeyF4) And ParentAC.EnableF4Copy Then
    If UpdateGrid Then Call CopyPasteCurrentGridRow(ADODataGrid.Recordset, grid, ParentAC.AutoColumns)
  End If
  If (Shift = 0) And (KeyCode = vbKeyF9) And ParentAC.EnableF9Refresh Then
    Call RefreshGrid(False, True, False, False)
  End If
  If IsValidKeyCode(KeyCode, Shift) Then Call CheckDropDown(Nothing)
  inKeyDown = False
End Sub

Private Sub Grid_ComboSelect(ByVal ColIndex As Integer)
  Dim acol As AutoCol
  
  If GetAColByGridIndex(ColIndex, acol) Then
    acol.FastKey = ""
    acol.FastKeyPress = ""
    Call DisplayFastKey(acol.FastKey)
  End If
End Sub

Private Sub GridSpeedKey(KeyAscii As Integer, Optional ByVal bGridKeyPressDelay As Boolean = False)
  Dim SetSelText As Boolean, uKeyAscii As Integer
  Dim vItem As ValueItem, sValue As String, fLen As Long, s As String, vValue As Variant
  Dim acol As AutoCol, ColSet As TrueOleDBGrid60.Column, NewValue As Integer
  Dim cCtrl As TrueOleDBGrid60.TDBDropDown, rs As Recordset, vbmk As Variant
    
  On Error GoTo GridSpeedKey_err
  If GetAColByGridIndex(grid.Col, acol) Then
    If KeyAscii = vbKeyReturn Then
      If IsComboOpen(grid.hwnd) And (Len(m_SaveComboValue) > 0) Then Call grid.PostMsg(MSG_COMBO_RESELECT)
      Exit Sub
    End If
    If acol.IsCheckBox Then
      If (KeyAscii = vbKeySpace) Or (KeyAscii = 0) Then
        Set ColSet = grid.Columns(grid.Col)
        If ColSet.Locked Or (Not grid.AllowUpdate) Then Exit Sub
        If IsNumeric(ColSet.Value) Then
          NewValue = Not (ColSet.Value = -1)
        Else
          NewValue = False
        End If
        ColSet.Value = NewValue
        'ColSet.Text = ""
      End If
      KeyAscii = 0
      Exit Sub
    End If
    If ((acol.DropType And DROPDOWN_LIST) = DROPDOWN_LIST) Or _
       ((acol.DropType And DROPDOWN_COMBO) = DROPDOWN_COMBO) Then
      m_SaveComboValue = ""
      Set ColSet = grid.Columns(grid.Col)
      If ColSet.Locked Or (Not grid.AllowUpdate) Then Exit Sub
      If Not bGridKeyPressDelay Then
        If (acol.DropType And DROPDOWN_COMBO) = DROPDOWN_COMBO Then
          acol.FastKey = Left$(grid.Text, grid.SelStart)
        End If
        If KeyAscii = vbKeyBack Then
          If Len(acol.FastKeyPress) > 0 Then acol.FastKeyPress = Left$(acol.FastKeyPress, Len(acol.FastKeyPress) - 1)
          If (acol.DropType And DROPDOWN_COMBO) = DROPDOWN_COMBO Then Exit Sub
          If Len(acol.FastKey) > 0 Then acol.FastKey = Left$(acol.FastKey, Len(acol.FastKey) - 1)
        ElseIf IsPrint(KeyAscii) Then
          acol.FastKey = acol.FastKey & Chr$(ToUpper(KeyAscii))
          acol.FastKeyPress = acol.FastKeyPress & Chr$(KeyAscii)
        ElseIf KeyAscii = vbKeyEscape Then
          acol.FastKey = ""
          acol.FastKeyPress = ""
        End If
        'Debug.Print "BEGIN " & aCol.FastKey & ":" & aCol.FastKeyPress; ":" & Grid.Text
        If (acol.DropType And DROPDOWN_LIST) = DROPDOWN_LIST Then
          Call DisplayFastKey(acol.FastKey)
          If KeyAscii <> vbKeyEscape Then KeyAscii = 0
        End If
      End If
      fLen = Len(acol.FastKey)
      If fLen = 0 Then Exit Sub
      If (acol.QueryCombo > 0) And (ColSet.ValueItems.Count = 0) Then
        Set cCtrl = m_AutoGrid_ADO.Combo(acol.QueryCombo)
        If cCtrl.DataSource Is Nothing Then Exit Sub
        vbmk = cCtrl.Bookmark
        Set rs = cCtrl.DataSource
        rs.Find rs.Fields(0).Name & " like " & StrSQL(acol.FastKey & "*"), 0, adSearchForward
        If Not rs.EOF Then
          cCtrl.FirstRow = rs.Bookmark
          ColSet.Text = cCtrl.Text
        End If
      Else
        For Each vItem In ColSet.ValueItems
          If ((acol.DropType And DROPDOWN_LIST) = DROPDOWN_LIST) And ColSet.ValueItems.Translate Then
            vValue = vItem.DisplayValue
          Else
            vValue = vItem.Value
          End If
          s = Left$(vValue, fLen)
          If StrComp(s, acol.FastKey, vbTextCompare) = 0 Then
            sValue = vItem.Value
            If IsComboOpen(grid.hwnd) And ((acol.DropType And DROPDOWN_LIST) = DROPDOWN_LIST) Then
              ColSet.Text = vValue
              m_SaveComboValue = sValue
            Else
              ColSet.Text = sValue
            End If
            'ColSet.Text = vItem.DisplayValue
            'ColSet.Value = vItem.Value
            SetSelText = True
            Exit For
          End If
        Next vItem
      End If
      'Debug.Print "DropDown: " & ColSet.Text & "(" & ColSet.Value & ")"
      'If Not SetSelText Then Grid.SelText = ""
      If (acol.DropType And DROPDOWN_LIST) = DROPDOWN_LIST Then
        If IsComboOpen(grid.hwnd) Then
          'Debug.Print "Combo Open: " & ColSet.Text & " : " & Grid.Text
        End If
      ElseIf (acol.DropType And DROPDOWN_COMBO) = DROPDOWN_COMBO Then
        If Not SetSelText Then
          If IsPrint(KeyAscii) Then
            ColSet.Text = Left$(acol.FastKeyPress, Len(acol.FastKeyPress) - 1)
          Else
            ColSet.Text = acol.FastKeyPress
          End If
          grid.Text = ColSet.Text
          Call DisplayFastKey("")
        Else
          KeyAscii = 0
          grid.SelText = ""
          ColSet.Text = sValue
          grid.Text = ColSet.Text
          grid.SelLength = 0
          grid.SelStart = fLen
          grid.SelLength = Len(ColSet.Text) - fLen
          Call DisplayFastKey(acol.FastKey)
        End If
      End If
      'Debug.Print "END   " & aCol.FastKey & ":" & Grid.Text
    End If
  End If
GridSpeedKey_end:
  Exit Sub
  
GridSpeedKey_err:
  Debug.Print "GridSpeedKey Error: " & Err.Description
  Resume GridSpeedKey_end
  Resume
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
  Call GridSpeedKey(KeyAscii)
End Sub

Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Dim acol As AutoCol, ColCheckBox As Boolean
  Dim MouseCol As Long, CurRow As Long, HeaderHeight As Single
  Dim UMenus As Collection, um As UserPopupMenu, NextMenu As Long
  Dim GridState As POPUP_MENU_SHOWFLAGS, ActiveMenu As Boolean
  
  On Error GoTo Grid_MouseDown_err
  If CheckMouseButton(VK_RBUTTON) Then
    MouseCol = grid.ColContaining(x)
    If MouseCol <> -1 Then grid.Col = MouseCol
    Load frmPopupMenu
    CurRow = grid.RowContaining(Y)
    frmPopupMenu.mnuSortAsc.Enabled = True
    frmPopupMenu.mnuSortDesc.Enabled = True
    HeaderHeight = grid.HeadLines * (grid.HeadFont.Size * 1440) / 72
    GridState = SHOW_ALWAYS
    
    ' If in Edit then do not change current row (assume no current row)
    If grid.DataChanged Then MouseCol = -1
    
    ' No Valid Column implies no Valid Row
    If MouseCol = -1 Then CurRow = -1
    If (Y > grid.RowHeight) And (CurRow <> -1) Then
      ColCheckBox = False
      Call GetAColByGridIndex(grid.Col, acol)
      If Not acol Is Nothing Then ColCheckBox = acol.IsCheckBox
      grid.Row = CurRow
      frmPopupMenu.mnuFilterSelectionEx.visible = True
      frmPopupMenu.mnuFilterSelectionInc.visible = True

      frmPopupMenu.mnuFilterSelectionEx.Enabled = (Len(grid.SelText) > 0) Or ColCheckBox Or (Len(grid.Text) = 0)
      frmPopupMenu.mnuFilterSelectionInc.Enabled = (Len(grid.SelText) > 0) Or ColCheckBox Or (Len(grid.Text) = 0)
      'frmPopupMenu.mnuFilterS.Visible = True
      'frmPopupMenu.mnuFilterS.Enabled = False  'Not (m_AutoGrid_ADO Is Nothing)
        
      frmPopupMenu.mnuCopyCol.visible = False
      frmPopupMenu.mnuInsCol.visible = False
      
      frmPopupMenu.mnuAutoSize.visible = False
      frmPopupMenu.mnuAutoSizeAll.visible = False
          
      'APF Call frmFilter.FillRCMenu(Me)
      frmPopupMenu.sep3.visible = False
      
      frmPopupMenu.sep4.visible = True
      frmPopupMenu.mnuCopyRow.visible = True
      frmPopupMenu.mnuCutRow.visible = True
      frmPopupMenu.mnuPasteRow.visible = True
      
      frmPopupMenu.sep5.visible = False
      frmPopupMenu.mnuFind.visible = True
      If Len(grid.SelText) > 0 Then
        GridState = GridState + VALID_ROWSELECTION
      Else
        GridState = GridState + VALID_ROW
      End If
    Else
      frmPopupMenu.mnuFilterSelectionEx.visible = False
      frmPopupMenu.mnuFilterSelectionInc.visible = False
      
      frmPopupMenu.mnuSortAsc.visible = (MouseCol <> -1)
      frmPopupMenu.mnuSortDesc.visible = (MouseCol <> -1)
      frmPopupMenu.mnuCopyCol.visible = (MouseCol <> -1)
      frmPopupMenu.mnuInsCol.visible = (MouseCol <> -1)
      frmPopupMenu.mnuRemoveFilter.visible = Not grid.DataChanged
      frmPopupMenu.mnuAutoSize.visible = (MouseCol <> -1)
      frmPopupMenu.mnuAutoSizeAll.visible = True
      
      frmPopupMenu.sep2.visible = (MouseCol <> -1) And (Not grid.DataChanged)
      
      frmPopupMenu.mnuCopyCol.Enabled = (Y <= HeaderHeight) And (Not m_UnboundColumns)
      frmPopupMenu.mnuInsCol.Enabled = (Y <= HeaderHeight) And (StartCol >= 0) And (Not m_UnboundColumns)
      
      frmPopupMenu.mnuAutoSize.Enabled = (Y <= HeaderHeight)
      frmPopupMenu.mnuAutoSizeAll.Enabled = (Y > HeaderHeight) Or (MouseCol = -1)
        
      frmPopupMenu.sep3.visible = (MouseCol <> -1)
      
      frmPopupMenu.sep4.visible = (MouseCol <> -1)
      frmPopupMenu.mnuCopyRow.visible = (MouseCol <> -1)
      frmPopupMenu.mnuCutRow.visible = False
      frmPopupMenu.mnuPasteRow.visible = False
      
      If (grid.VisibleRows = 0) Or (MouseCol = -1) Then
        'frmPopupMenu.mnuFilterS.Visible = False
        frmPopupMenu.mnuSortAsc.Enabled = False
        frmPopupMenu.mnuSortDesc.Enabled = False
        frmPopupMenu.mnuFind.visible = False
        frmPopupMenu.sep5.visible = False
      Else
        'frmPopupMenu.mnuFilterS.Visible = True
        'frmPopupMenu.mnuFilterS.Enabled = Not (m_AutoGrid_ADO Is Nothing)
        frmPopupMenu.mnuFind.visible = True
        frmPopupMenu.mnuFind.Enabled = (Not (m_AutoGrid_ADO Is Nothing)) And (Not grid.Columns(MouseCol).ValueItems.Translate) Or True 'apf
        frmPopupMenu.sep5.visible = True
      End If
    End If
    
    frmPopupMenu.mnuDebugSep.visible = ParentAC.DebugMode
    frmPopupMenu.mnuDebugMenu.visible = ParentAC.DebugMode
    If ParentAC.DebugMode Then
      If MouseCol >= 0 Then
        frmPopupMenu.mnuDebugMenu.Caption = "Show Format"
      Else
        frmPopupMenu.mnuDebugMenu.Caption = "Show SQL"
      End If
      frmPopupMenu.mnuDebugMenu.Tag = MouseCol
      frmPopupMenu.mnuDebugMenu.Enabled = (Y <= HeaderHeight)
    End If
    
    frmPopupMenu.mnuAbortEdit.visible = grid.DataChanged
    frmPopupMenu.mnuCommitEdit.visible = grid.DataChanged
    frmPopupMenu.sep6.visible = grid.DataChanged And (MouseCol <> -1)
      
    frmPopupMenu.mnuCopyRow.Enabled = (Len(grid.SelText) > 0) Or (CurRow = -1)
    frmPopupMenu.mnuCutRow.Enabled = grid.AllowUpdate And (Len(grid.SelText) > 0)
    frmPopupMenu.mnuPasteRow.Enabled = grid.AllowUpdate And IsCellPaste()
       
    frmPopupMenu.UserSep.visible = False
    Set UMenus = ParentAC.UserPopupMenus
      
    If Not UMenus Is Nothing Then
      frmPopupMenu.UserSep.visible = (UMenus.Count > 0)
      If ParentAC.UserPopupMenuHandler Is Nothing Then GridState = 0
      For Each um In UMenus
        NextMenu = frmPopupMenu.mnuUserMenus.UBound
        NextMenu = NextMenu + 1
        Load frmPopupMenu.mnuUserMenus(NextMenu)
        frmPopupMenu.mnuUserMenus(NextMenu).Caption = um.Caption
        ActiveMenu = ((um.ShowFlags And VALID_MASK And GridState) = (um.ShowFlags And VALID_MASK))
        If ((um.ShowFlags And VALID_VISIBLE) = VALID_VISIBLE) Then
          frmPopupMenu.mnuUserMenus(NextMenu).visible = ActiveMenu
        Else
          frmPopupMenu.mnuUserMenus(NextMenu).visible = True
        End If
        frmPopupMenu.mnuUserMenus(NextMenu).Enabled = ActiveMenu
      Next um
    End If
    
    m_Popped = True
    Set frmPopupMenu.AGrid = Me
    Call frmPopupMenu.PopupMenu(frmPopupMenu.mnuGrid)
    Unload frmPopupMenu
  End If
  
Grid_MouseDown_end:
  Exit Sub
  
Grid_MouseDown_err:
  ' 6145 = Invalid Column Index, 6148 = Invalid Row Number
  'If (Err.Number = 6145) Or (Err.Number = 6148) Then Resume Grid_MouseDown_end
  Call ErrorMessage(ERR_ERROR, Err, "AutoGrid_MouseDown", "Display Popup Menu", "Unable to display right mouse popup menu")
  Resume Grid_MouseDown_end
End Sub

Private Function MoveColumns(ByVal InsertAtColumn As Long) As Boolean
  Dim acol As AutoCol, nGridCols As Long, bRebaseDone(1 To 3) As Boolean
  Dim i As Long, j As Long, dTemp As Long
  
  On Error GoTo MoveColumns_Err
  Call xSet("MoveColumns")
  If InsertAtColumn > StartCol Then InsertAtColumn = InsertAtColumn - 1
  If (InsertAtColumn >= StartCol) And (InsertAtColumn <= EndCol) Then Err.Raise ERR_MOVECOLUMNS, "MoveColumns", "Cannot move into the selected column range"
  ' rebase column references InsertAtColumn, StartCol, EndCol to columns in m_DisplayOrder
   nGridCols = 0
   For i = 1 To ParentAC.Count
    Set acol = ParentAC.Item(i)
    If Not acol.Hide Then
      If (Not bRebaseDone(1)) And (nGridCols = InsertAtColumn) Then
        InsertAtColumn = i
        bRebaseDone(1) = True
      End If
      If (Not bRebaseDone(2)) And (nGridCols = StartCol) Then
        StartCol = i
        bRebaseDone(2) = True
      End If
      If (Not bRebaseDone(3)) And (nGridCols = EndCol) Then
        EndCol = i
        bRebaseDone(3) = True
      End If
      nGridCols = nGridCols + 1
    End If
  Next i
  If InsertAtColumn > StartCol Then ' move up
    For i = EndCol To StartCol Step -1
      dTemp = m_DisplayOrder(i)
      For j = i To (InsertAtColumn - 1)
        m_DisplayOrder(j) = m_DisplayOrder(j + 1)
      Next j
      m_DisplayOrder(InsertAtColumn) = dTemp
      InsertAtColumn = InsertAtColumn - 1
    Next i
  Else ' move down
    For i = StartCol To EndCol
      dTemp = m_DisplayOrder(i)
      For j = i To (InsertAtColumn + 1) Step -1
        m_DisplayOrder(j) = m_DisplayOrder(j - 1)
      Next j
      m_DisplayOrder(InsertAtColumn) = dTemp
      InsertAtColumn = InsertAtColumn + 1
    Next i
  End If
  Call ReBindColumns
  MoveColumns = True
  
MoveColumns_end:
  Call xReturn("MoveColumns")
  Exit Function
  
MoveColumns_Err:
  Call ErrorMessage(ERR_ERROR, Err, "MoveColumns", "Error moving columns", "Unable to move columns (" & StartCol & " to " & EndCol & ") to position " & InsertAtColumn)
  Resume MoveColumns_end
End Function

Public Sub MenuAction(ByVal pm As POPUP_MENU_ACTIONS)
  m_Popped = True
  Call PopupMenuAction(pm)
End Sub

Friend Sub PopupMenuAction(ByVal pm As POPUP_MENU_ACTIONS)
  Dim i As Long

  On Error GoTo PopupMenuAction_err
  If m_Popped Then
    Select Case pm
      Case AUTOSIZE_COLUMN
        Call AutoWidth(grid.Col)
      Case AUTOSIZEALL_COLUMNS
        Call AutoWidth
      Case SORT_ASC
        Call SortColumn(True)
      Case SORT_DESC
        Call SortColumn(False)
      Case MOVE_COLUMN
        If grid.SelStartCol >= 0 Then
          StartCol = grid.SelStartCol
          EndCol = grid.SelEndCol
        End If
      Case INSERT_COLUMN
        If StartCol > 0 Then
          If MoveColumns(grid.Col) Then StartCol = -1
        End If
      Case MENU_COPY
        Call CopySelection
      Case MENU_COPYPASTE_ROW
        Call CopyPasteCurrentGridRow(ADODataGrid.Recordset, grid, ParentAC.AutoColumns)
      Case MENU_CUT
        Call CutSelection
      Case FILTER_SELECTION_INC
        Call FilterBySelection(True)
      Case FILTER_SELECTION_EX
        Call FilterBySelection(False)
      Case FILTER_WIZARD
        Call ShowFilterWizard
      Case FILTER_REMOVE
        Call ParentAC.RemoveFilterSort
      Case MENU_FIND_INCOL
        Call FindInCol
      Case MENU_PASTE
        Call PasteSelection
      Case COMMIT_EDIT
        Call UpdateGrid
      Case MENU_DEBUG
        Call ShowColumnDebug(frmPopupMenu.mnuDebugMenu.Tag)
    End Select
    m_Popped = False
  End If
  Exit Sub
  
PopupMenuAction_err:
  Call ErrorMessage(ERR_ERROR, Err, "PopupMenuAction", "Execute Popup menu action", "Failed to execute Popup Menu action")
End Sub

Private Function IsFieldUpdateable(ByVal acol As AutoCol) As Boolean
  IsFieldUpdateable = IsFieldUpdateableADO(ADODataGrid.Recordset, acol.DataField)
End Function

Private Sub ClearGridColumns()
  grid.ClearFields
  Do While grid.Columns.Count > 0
    If grid.Columns.Count > 1 Then
      Call grid.Columns.Remove(1)
    ElseIf grid.Columns.Count = 1 Then
      Call grid.Columns.Remove(0)
    End If
  Loop
End Sub

Friend Sub RefreshUpdateMode(Optional ByVal ForceDDRefresh As Boolean = False)
  Dim bRequireEmptyRows As Boolean
  If (Len(ParentAC.Filter) > 0) Or (Len(ParentAC.Sort) > 0) Then
    grid.AllowAddNew = False
    grid.AllowUpdate = False
    grid.AllowDelete = False
  Else
    grid.AllowAddNew = m_GridAllowAddNew
    grid.AllowUpdate = m_GridAllowUpdate
    grid.AllowDelete = m_GridAllowDelete
  End If
  If ForceDDRefresh Then Call RefreshDropDowns
End Sub

Public Sub ReDisplayColumns()
  Call xSet("ReDisplayColumns")
  Call ReBindColumns
  Call xReturn("ReDisplayColumns")
End Sub

Private Sub CreateGridSplit(ByVal FullRefresh As Boolean, ByVal InResize As Boolean)
  Dim lSplit As TrueOleDBGrid60.Split, rSplit As TrueOleDBGrid60.Split, SplitCol As Long
  Dim acol As AutoCol, ColSet As TrueOleDBGrid60.Column, i As Long
    
  On Error GoTo CreateGridSplit_err
  Set lSplit = grid.Splits.Item(0)
  If FullRefresh Then
    Do While grid.Splits.Count > 1
      Call grid.Splits.Remove(0)
    Loop
    For i = 1 To ParentAC.Count
      Set acol = ParentAC.Item(i)
      If acol.GridColumn >= 0 Then
        If acol.SplitColumn Then
          Set rSplit = grid.Splits.Add(1)
          SplitCol = acol.GridColumn + 1
          grid.Split = (grid.Splits.Count - 1)
          Exit For
        End If
      End If
    Next i
  Else
    If grid.Splits.Count > 1 Then
      For i = 1 To ParentAC.Count
        Set acol = ParentAC.Item(i)
        If acol.GridColumn >= 0 Then
          If acol.SplitColumn Then
            Set rSplit = grid.Splits.Item(1)
            SplitCol = acol.GridColumn + 1
            Exit For
          End If
        End If
      Next i
    End If
  End If
  
  If Not rSplit Is Nothing Then
    grid.Split = (grid.Splits.Count - 1)
    ' split 0 LHS, 1 RHS
    lSplit.SizeMode = dbgNumberOfColumns
    lSplit.Size = SplitCol
    For i = 0 To (lSplit.Columns.Count - 1)
      Set ColSet = lSplit.Columns(i)
      If InResize Then ColSet.Width = grid.Columns.Item(ColSet.ColIndex).Width
      ColSet.visible = (i < SplitCol)
      ColSet.AllowSizing = ColSet.visible
    Next i
    For i = 0 To (rSplit.Columns.Count - 1)
      Set ColSet = rSplit.Columns(i)
      If InResize Then ColSet.Width = grid.Columns.Item(ColSet.ColIndex).Width
      ColSet.visible = (i >= SplitCol)
      ColSet.AllowSizing = ColSet.visible
    Next i
    rSplit.RecordSelectors = False
    grid.TabAcrossSplits = True
  End If
  Exit Sub
  
CreateGridSplit_err:
  Err.Raise Err.Number, ErrorSource(Err, "CreateGridSplit"), Err.Description
End Sub

Friend Sub ReBindColumns(Optional ByVal doUpdateModeRefresh As Boolean = True)
  Dim i As Long, id As Integer, CharCount As Long
  Dim acol As AutoCol, ColSet As TrueOleDBGrid60.Column

  On Error GoTo ReBindColumns_err
  m_GridActive = False
  Call ClearGridColumns
  grid.CellTips = dbgNoCellTips
  
  'Get Display order
  For i = 1 To ParentAC.Count
    Set acol = ParentAC.Item(m_DisplayOrder(i))
    acol.GridColumn = -1
    If acol.ForceCreateColumn Or (Not acol.Hide) Then
      id = grid.Columns.Count
      Set ColSet = grid.Columns.Add(id)
      acol.GridColumn = id
      CharCount = CountChar(acol.GridCaption, vbCr, vbBinaryCompare) + 1
      If CharCount > grid.HeadLines Then grid.HeadLines = CharCount
      ColSet.Caption = acol.GridCaption
      If Not acol.UnboundColumn Then ColSet.DataField = acol.DataField
      ColSet.Alignment = acol.Alignment
      ColSet.Locked = (Not IsFieldUpdateable(acol)) Or acol.NoEdit
      ColSet.visible = True
      ColSet.ConvertEmptyCell = IIf(m_BlankCellAsNull, 1, 0)
      If Len(acol.ToolTip) > 0 Then
        grid.CellTipsDelay = 500
        grid.CellTips = dbgFloating
      End If
      If acol.Hide Then acol.ColWidth = 0
      If acol.ColWidth > 0 Then
        ColSet.Width = acol.ColWidth
      ElseIf acol.ColWidth = 0 Then
        ColSet.visible = False
      End If
      If acol.CustomFormat Then
        ColSet.NumberFormat = "FormatText Event"
      Else
        ColSet.NumberFormat = acol.Format
      End If
      ColSet.Button = acol.HasButton
      If acol.HasButton And acol.IsCheckBox Then Err.Raise ERR_GRIDFORMAT, "RebindColumns", "Cannot have the CheckBox and Button format on the same column" & vbCrLf & "column [" & acol.GridCaptionClean & "]" & vbCrLf
      If acol.HasButton Then ColSet.ButtonPicture = LoadResPicture(acol.ButtonPictureIndex, vbResBitmap)
      'FIX CAD1
      ColSet.ForeColor = acol.ForeColor
      ColSet.BackColor = acol.BackColor
    End If
  Next i
  Call grid.HoldFields
  Call CreateGridSplit(True, False)
  grid.ReBind
  grid.RowHeight = 0
  'Grid.ExtendRightColumn = True
  m_GridActive = True
  If doUpdateModeRefresh Then Call RefreshUpdateMode(True)
  Exit Sub
  
ReBindColumns_err:
  Err.Raise Err.Number, ErrorSource(Err, "ReBindColumns"), Err.Description
  Resume
End Sub
Private Sub Grid_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
  Dim acol As AutoCol

  On Error GoTo Grid_FormatText_err
  If GetAColByGridIndex(ColIndex, acol) Then Value = acol.GetFormattedValue(Value)
  Exit Sub

Grid_FormatText_err:
  Value = "Error in Grid FormatText " & Err.Description
End Sub

Private Sub RefreshDropDowns()
  Dim i As Long
  Dim acol As AutoCol
 
  On Error GoTo RefreshDropDowns_err
  For i = 1 To ParentAC.Count
    Set acol = ParentAC.Item(i)
    If acol.IsCheckBox Or (((acol.DropType And DROPDOWN_LIST) = DROPDOWN_LIST) And (acol.TranslateDropDown <> 0)) Then
      Call RefreshSingleDropDown(acol, False)
    Else
      acol.RefreshRequired = True
    End If
  Next i
  Set acol = Nothing
  Exit Sub
  
RefreshDropDowns_err:
  Err.Raise Err.Number, ErrorSource(Err, "RefreshDropDowns"), Err.Description
  Resume
End Sub

Private Function RefreshSingleDropDown(ByVal acol As AutoCol, ByVal ShowButtonOnly As Boolean) As Boolean
  Dim rs As Recordset, cn As Connection
  Dim i As Long, j As Long, dItems As Long, cCount As Long
  Dim ColSet As TrueOleDBGrid60.Column
  Dim vItem As ValueItem, doResize As Boolean
  Dim errstring As String
 
  On Error GoTo RefreshSingleDropDown_err
  If m_InGridRefresh Then Exit Function
  doResize = False
  acol.RefreshRequired = False
  Set cn = ADODataGrid.Recordset.ActiveConnection
  If (acol.GridColumn >= 0) And ((acol.DropType <> DROPDOWN_NONE) Or acol.IsCheckBox) Then
    Set ColSet = grid.Columns(acol.GridColumn)
    If acol.IsCheckBox Then
      'APF 18/9/2002 replaced
      ColSet.ValueItems.Presentation = dbgCheckBox
      ColSet.Alignment = dbgCenter
'      Call ColSet.ValueItems.Clear
'      Set vItem = New ValueItem
'      vItem.DisplayValue = LoadResPicture(TICK_BMP, vbResBitmap)
'      vItem.Value = -1
'      Call ColSet.ValueItems.Add(vItem)
'
'      Set vItem = New ValueItem
'      If ParentAC.CheckBoxCross Then
'        vItem.DisplayValue = LoadResPicture(CROSS_BMP, vbResBitmap)
'      Else
'        vItem.DisplayValue = LoadResPicture(CROSS_BLANK_BMP, vbResBitmap)
'      End If
'      vItem.Value = 0
'      Call ColSet.ValueItems.Add(vItem)
'
'      ColSet.ValueItems.Translate = True
'      ColSet.ValueItems.CycleOnClick = True
'      ColSet.ValueItems.Presentation = dbgNormal
'      ColSet.Alignment = dbgCenter
      acol.TranslateDropDown = -1
    ElseIf ((acol.DropType And DROPDOWN_LIST) = DROPDOWN_LIST) Or _
       ((acol.DropType And DROPDOWN_COMBO) = DROPDOWN_COMBO) Then
      doResize = True
      Call ColSet.ValueItems.Clear
      If ShowButtonOnly Then
        Set vItem = New ValueItem
        Call ColSet.ValueItems.Add(vItem)
        ColSet.ValueItems.Presentation = dbgComboBox
        acol.RefreshRequired = True
      Else
        If (acol.DropType And DROPDOWN_QUERY) = 0 Then  ' values only
          j = 1
          Do While j <= acol.NoParams
            Set vItem = New ValueItem
            If acol.DropType = DROPDOWN_COMBO Then
              vItem.Value = acol.GetParam(j)
              j = j + 1
            Else
              vItem.DisplayValue = acol.GetParam(j)
              vItem.Value = acol.GetParam(j + 1)
              j = j + 2
            End If
            Call ColSet.ValueItems.Add(vItem)
          Loop
          'ColSet.ValueItems.CycleOnClick = True
          cCount = 2  ' transloate if appropriate
        Else   ' query to non bound drop down
          Set rs = GetComboRecordsetADO(cn, acol)
          If rs Is Nothing Then GoTo RefreshSingleDropDown_end
          cCount = rs.Fields.Count
          Do While Not rs.EOF
            Set vItem = New ValueItem
            If (acol.DropType And DROPDOWN_COMBO) = DROPDOWN_COMBO Then
              vItem.Value = rs.Fields(0).Value
            Else
              If cCount > 1 Then
                vItem.DisplayValue = rs.Fields(0).Value
                vItem.Value = rs.Fields(1).Value
              Else
                vItem.Value = rs.Fields(0).Value
              End If
            End If
            Call ColSet.ValueItems.Add(vItem)
            rs.MoveNext
          Loop
          Set rs = Nothing
        End If
        'ColSet.ValueItems.Validate = True
        If ((acol.DropType And DROPDOWN_LIST) = DROPDOWN_LIST) And (cCount > 1) Then
          ColSet.ValueItems.Translate = True
          acol.TranslateDropDown = -1
        Else
          ColSet.ValueItems.Translate = False
          acol.TranslateDropDown = 0
        End If
        ColSet.ValueItems.Presentation = dbgComboBox
        
        If acol.MaxDropDownItems > 0 Then
          dItems = acol.MaxDropDownItems
        Else
          dItems = ParentAC.MaxDropDownItems
        End If
        If dItems > ColSet.ValueItems.Count Then
          dItems = ColSet.ValueItems.Count
          If dItems <= 0 Then dItems = 1
        End If
        ColSet.ValueItems.MaxComboItems = dItems
        'ColSet.ValueItems.CycleOnClick = ((aCol.DropType And DROPDOWN_LIST) = DROPDOWN_LIST)
      End If
    Else
      doResize = True
      Call DoQueryCombos(cn, acol, ColSet)
    End If
  End If
  If doResize Then Call AutoSizeSingleDropDown(acol)
  RefreshSingleDropDown = True
      
RefreshSingleDropDown_end:
  Set acol = Nothing
  Exit Function
  
RefreshSingleDropDown_err:
  If Not acol Is Nothing Then
    If acol.IsCheckBox Then
      errstring = "Error displaying the CheckBox for column [" & acol.GridCaptionClean & "]" & vbCrLf
    Else
      errstring = "Error refreshing the drop down for column [" & acol.GridCaptionClean & "]" & vbCrLf
      If (acol.DropType And DROPDOWN_QUERY) = 0 Then
        errstring = errstring & "Static values based drop down" & vbCrLf
      Else
        errstring = errstring & "Query based drop down: [" & acol.QueryName & "]" & vbCrLf
      End If
      errstring = errstring & "Drop down type: " & GetDropDownTypeAsString(acol.DropType) & vbCrLf
    End If
  End If
  Err.Raise Err.Number, ErrorSource(Err, "RefreshSingleDropDown"), errstring & Err.Description
  Resume
End Function

Private Sub AutoSizeSingleDropDown(ByVal acol As AutoCol)
  Dim ColSet As TrueOleDBGrid60.Column
  Dim cCtrl As TrueOleDBGrid60.TDBDropDown
  Dim tWidth As Single, FirstColumn As Boolean
  Dim j As Long
  
  On Error GoTo AutoSizeSingleDropDown_err
  If acol.GridColumn >= 0 Then
    Set ColSet = grid.Columns(acol.GridColumn)
    If Len(ColSet.DropDown) > 0 Then
      Set cCtrl = m_AutoGrid_ADO.Combo(acol.QueryCombo)
      tWidth = 0
      FirstColumn = True
      For j = 0 To (cCtrl.Columns.Count - 1)
        If cCtrl.Columns(j).visible Then
          If FirstColumn Then
            cCtrl.Columns(j).Width = ColSet.Width
            FirstColumn = False
          End If
          tWidth = tWidth + cCtrl.Columns(j).Width
        Else
          cCtrl.Columns(j).Width = 0
        End If
      Next j
      cCtrl.Width = tWidth + 350 ' width + scrollbar
    End If
  End If
  Exit Sub
  
AutoSizeSingleDropDown_err:
  Err.Raise Err.Number, ErrorSource(Err, "AutoSizeSingleDropDown"), Err.Description
End Sub

Private Sub DoQueryCombos(ByVal cn As Connection, ByVal acol As AutoCol, ByVal ColSet As TrueOleDBGrid60.Column)
  Dim FirstField As Long, VisibleFields As Long
  Dim cCtrl As TrueOleDBGrid60.TDBDropDown, rs As Recordset
  Dim dItems As Long, j As Long
      
  If acol.MaxDropDownItems > 0 Then
    dItems = acol.MaxDropDownItems
  Else
    dItems = ParentAC.MaxDropDownItems
  End If
  
  Set cCtrl = m_AutoGrid_ADO.Combo(acol.QueryCombo)
  cCtrl.ColumnHeaders = False
  cCtrl.DataField = acol.BoundDataField
  cCtrl.Font = grid.Font
  Set rs = GetComboRecordsetADO(cn, acol)
  If Not rs Is Nothing Then
    Set cCtrl.DataSource = rs
    If rs.Fields.Count > 1 Then
      cCtrl.Columns(0).visible = False
      FirstField = 1
      VisibleFields = (rs.Fields.Count - 1)
    Else
      FirstField = 0
      VisibleFields = 1
    End If
    cCtrl.ListField = rs.Fields(FirstField).Name
    'cCtrl.Width = ColSet.Width * VisibleFields
    If dItems > rs.RecordCount Then
      dItems = rs.RecordCount
      If dItems <= 0 Then dItems = 1
    End If
    cCtrl.Height = (grid.RowHeight * dItems) + 25
    cCtrl.ExtendRightColumn = True
    ColSet.DropDown = cCtrl.Name
  End If
End Sub
          
Public Function GetDropRecordset(ByVal DataField As String) As Recordset
  Dim acol As AutoCol, vbmk As Variant
  Dim tdbgDrop As TrueOleDBGrid60.TDBDropDown
  
  Set acol = GetAColByKey(DataField)
  If Not acol Is Nothing Then
    Set tdbgDrop = m_AutoGrid_ADO.Combo(acol.QueryCombo)
    vbmk = tdbgDrop.Bookmark
    Set GetDropRecordset = tdbgDrop.DataSource
    If Not IsEmpty(vbmk) Then GetDropRecordset.Bookmark = vbmk
  End If
End Function

Private Function GetComboRecordsetADO(ByVal cn As Connection, ByVal acol As AutoCol) As Recordset
  Dim DynamicParams As Boolean, IsSQLStringParam As Boolean
  Dim rsGrid As Recordset
  Dim rsCmd As Command, rsParam As Parameter
  Dim i As Long, p As Long
  Dim ColIndex As Integer, sql As String
  Dim rsEmpty As Boolean, rsDoQuery As Boolean
  Dim param As String, paramValue As Variant
    
  On Error GoTo GetComboRecordsetADO_err
  ' apf should be dynamic ReadOnly Cursor
  DynamicParams = False
  Set rsCmd = New Command
  sql = acol.QueryName
  IsSQLStringParam = acol.IsSQLString
  rsDoQuery = True
    
  If acol.NoParams > 0 Then
    Set rsGrid = ADODataGrid.Recordset
    rsEmpty = rsGrid.EOF And rsGrid.BOF
    For i = 1 To acol.NoParams
      paramValue = Empty
      param = acol.GetParam(i)
      p = InStr(1, param, "<")
      If p = 1 Then
        paramValue = Mid$(param, 2, Len(param) - 2)
      Else
        ColIndex = GetGridColIndex(param)
        If ColIndex <> -1 Then
          paramValue = grid.Columns(ColIndex).Value
          If IsEmpty(paramValue) And Not rsEmpty Then
            paramValue = rsGrid.Fields(param)
          End If
          DynamicParams = True
        Else
          rsDoQuery = Not rsEmpty
          If Not rsEmpty Then paramValue = rsGrid.Fields(param)
        End If
      End If
      If paramValue = Empty Then Err.Raise ERR_GETCOMBORS, "GetComboRecordsetADO", "Parameter [" & param & "] could not be found"
      
      If IsSQLStringParam Then
        sql = ReplaceString(sql, BEGIN_PARAM & param & END_PARAM, GetSQLValueAsString(paramValue, acol.dbDataType), vbTextCompare)
      Else
        Set rsParam = rsCmd.CreateParameter("Parameter " & i, , adParamInput, , paramValue)
        rsCmd.Parameters.Append rsParam
        
      End If
    Next i
  End If
  
  If rsDoQuery Then
    rsCmd.CommandText = sql
    rsCmd.CommandType = adCmdText
    Set rsCmd.ActiveConnection = cn
    Set GetComboRecordsetADO = rsCmd.Execute(, , adCmdText)
  End If
  
  If DynamicParams Then
     acol.DropType = acol.DropType Or DROPDOWN_DYNAMIC
   Else
     acol.DropType = acol.DropType Or DROPDOWN_STATIC
  End If
  
GetComboRecordsetADO_end:
  Exit Function
  
GetComboRecordsetADO_err:
  Set GetComboRecordsetADO = Nothing
  Call ErrorMessage(ERR_ERROR, Err, "GetComboRecordsetADO", "Initialising drop down", "Unable to initailise drop down on column '" & acol.GridCaptionClean & "'." & vbCrLf & "The resultset to support the drop down could not be created.")
  Resume GetComboRecordsetADO_end
  Resume
End Function

' apf should be dynamic ReadOnly Cursor
Private Function GetComboSQL(ByVal acol As AutoCol) As String
  GetComboSQL = acol.QueryName
End Function

Public Sub Refresh(Optional ByVal ForceRefresh As Boolean = False)
  Call xSet("Refresh")
  Call RefreshGrid(True, ForceRefresh, ForceRefresh, False)
  Call xReturn("Refresh")
End Sub

Private Function IsDeletedRecord(ByVal rs As Recordset) As Boolean
  IsDeletedRecord = False
  If Not (rs.EOF Or rs.BOF) Then
    IsDeletedRecord = ((rs.Status And adRecDeleted) = adRecDeleted) Or ((rs.Status And adRecDBDeleted) = adRecDBDeleted)
  End If
End Function

Friend Sub RefreshGrid(ByVal SetFocus As Boolean, ByVal ForceDataRefresh As Boolean, ByVal ForceDataRebind As Boolean, ByVal InPostMessage As Boolean)
  Dim OldNotifyDisable As Boolean
  Dim RowOffset As Long, ColOffset As Long
  Dim rs As Recordset, tmpv As Variant
    
  On Error GoTo RefreshGrid_err
  If m_InGridRefresh Then Exit Sub
  OldNotifyDisable = mNotifyDisable
  mNotifyDisable = True
  m_InGridRefresh = True
  Call ResetState
  If Not grid Is Nothing Then
    If grid.visible Then
      Call UpdateGrid
      m_FirstRowBMK = Null
      m_RowBMK = Null
      If Len(grid.FirstRow) <> 0 Then
        m_FirstRowBMK = grid.FirstRow
        m_RowBMK = grid.Bookmark
      End If
      If ForceDataRefresh Then
        If Not ADODataGrid Is Nothing Then
          Set rs = ADODataGrid.Recordset
          If InPostMessage Then Debug.Print "Process delayed post"
          If IsDeletedRecord(rs) Then
            Debug.Print "Post Refresh"
            If Not InPostMessage Then Call grid.PostMsg(MSG_FORCE_REFRESH_GRID)
            GoTo RefreshGrid_end
          End If
        End If
        Call GetCurrentRowColOffset(ColOffset, RowOffset, rs)
        ParentAC.RequeryOnRefresh = True
        ParentAC.SortFilterRefresh = True
        Call ParentAC.SetFilterSort
      End If
      grid.Refresh
      If ForceDataRefresh Then
        grid.CurrentCellVisible = True
        If SetFocus Then Call TryGridSetFocus
        Call SetCurrentPosition(ColOffset, RowOffset, grid.Bookmark)
        If ForceDataRebind Then Call ReBindColumns(False)
      Else
        tmpv = ADODataGrid.Recordset.Sort
        ADODataGrid.Recordset.Sort = tmpv
        Debug.Print "F: " & IsNullEx(m_FirstRowBMK, "(null)") & " R: " & IsNullEx(m_RowBMK, "(null)")
        Call grid.PostMsg(MSG_RESTORE_GRID_POSITION)
      End If
    End If
  End If
  
RefreshGrid_end:
  m_InGridRefresh = False
  If ForceDataRefresh Then Call RefreshDropDowns
  mNotifyDisable = OldNotifyDisable
  Exit Sub
  
RefreshGrid_err:
  Call DisplayGridError
  Resume RefreshGrid_end
  Resume
End Sub

Private Sub GetCurrentRowColOffset(ColOffset As Long, RowOffset As Long, ByVal rs As Recordset)
  ColOffset = 0
  RowOffset = 0
  If Not grid.BOF Then
    ColOffset = grid.Col
    If Not rs Is Nothing Then RowOffset = (rs.AbsolutePosition + 1) * -1
    Do
      RowOffset = RowOffset + 1
    Loop Until Not IsNull(grid.GetBookmark(RowOffset)) Or (RowOffset >= 0)
    If RowOffset > 0 Then RowOffset = 0
  End If
  RowOffset = (RowOffset * -1)
End Sub

Private Sub SetCurrentPosition(ByVal ColOffset As Long, ByVal RowOffset As Long, ByVal vCurRow As Variant)
  On Error Resume Next
  grid.Col = ColOffset
  Call grid.MoveRelative(RowOffset, vCurRow)
End Sub

Property Let Font(ByVal NewValue As String)
  On Error Resume Next
  grid.Font.Name = NewValue
  grid.HeadFont.Name = NewValue
  ADODataGrid.Font.Name = NewValue
End Property

Property Let FontSize(ByVal NewValue As Long)
  On Error Resume Next
  grid.Font.Size = NewValue
  grid.HeadFont.Size = NewValue
  grid.RowHeight = 0
  ADODataGrid.Font.Size = NewValue
End Property

Property Let AllowUpdate(ByVal NewValue As Boolean)
  If NewValue And Not IsDataUpdateable() Then Call Err.Raise(380, "AllowUpdate", "Data not updateable!")
  grid.AllowUpdate = NewValue
  m_GridAllowUpdate = NewValue
End Property

Property Get AllowUpdate() As Boolean
  AllowUpdate = grid.AllowUpdate
End Property

Property Let AllowAddNew(ByVal NewValue As Boolean)
  If NewValue And Not IsDataUpdateable() Then Call Err.Raise(380, "AllowUpdate", "Data not updateable!")
  grid.AllowAddNew = NewValue
  m_GridAllowAddNew = NewValue
End Property

Property Get AllowAddNew() As Boolean
  AllowAddNew = grid.AllowAddNew
End Property

Property Let AllowDelete(ByVal NewValue As Boolean)
  If NewValue And Not IsDataUpdateable() Then Call Err.Raise(380, "AllowUpdate", "Data not updateable!")
  grid.AllowDelete = NewValue
  m_GridAllowDelete = NewValue
End Property

Property Get AllowDelete() As Boolean
  AllowDelete = grid.AllowDelete
End Property

Public Sub Find(ByVal FieldName As String)
  Call ParentAC.Find(FieldName)
  grid.SelText = grid.Text
  grid.SelLength = Len(grid.Text)
  Call TryGridSetFocus
End Sub

Public Sub SortColumn(ByVal Asc As Boolean)
  Dim DataField As String
  Dim i As Integer
  Dim sSort As String
  
  If grid.SelStartCol >= 0 Then
    For i = grid.SelStartCol To grid.SelEndCol
      DataField = ColKey(i)
      If Len(DataField) > 0 Then
        sSort = sSort & "[" & DataField & "] " & IIf(Asc, "ASC", "DESC")
        If i <> grid.SelEndCol Then sSort = sSort & ","
      End If
    Next i
    ParentAC.Sort = sSort
  Else
    DataField = ColKey(grid.Col)
    If Len(DataField) > 0 Then
      ParentAC.Sort = "[" & DataField & "] " & IIf(Asc, "ASC", "DESC")
    End If
  End If
End Sub

Public Sub FindInCol()
  Dim acol As AutoCol
  If GetAColByGridIndex(grid.Col, acol) Then Call Find(acol.DataField)
End Sub

Private Function GetAColByGridIndex(ByVal GridColumn As Integer, acol As AutoCol) As Boolean
  GetAColByGridIndex = GetAColByGridIndexEx(GridColumn, acol, ParentAC)
End Function

Friend Function GetAColByKey(ByVal ColKey As Variant) As AutoCol
  Dim acol As AutoCol
  
  On Error Resume Next
  Set acol = ParentAC.Item(ColKey)
  Set GetAColByKey = acol
End Function

Public Function SetColumnValue(ByVal ColKey As Variant, ByVal ValueItem As COLUMN_VALUES, ByVal NewValue As Variant) As Boolean
  Dim acol As AutoCol
  
  Set acol = GetAColByKey(ColKey)
  If Not acol Is Nothing Then
    Select Case ValueItem
      Case COLUMN_HIDE
        acol.Hide = CBool(NewValue)
      Case Else
        Call ECASE("Unknown Set Value item: " & CStr(ValueItem))
    End Select
  End If
End Function

Public Function GetColumnValue(ByVal ColKey As Variant, ByVal ValueItem As COLUMN_VALUES) As Variant
  Dim acol As AutoCol
  
  GetColumnValue = Empty
  Set acol = GetAColByKey(ColKey)
  If Not acol Is Nothing Then
    Select Case ValueItem
      Case COLUMN_CAPTION
        GetColumnValue = acol.GridCaption
      Case COLUMN_CAPTION
        GetColumnValue = acol.GridCaption
      Case COLUMN_CAPTION_CLEAN
        GetColumnValue = acol.GridCaptionClean
      Case COLUMN_DATAFIELD
        GetColumnValue = acol.DataField
      Case COLUMN_HIDE
        GetColumnValue = acol.Hide
      Case Else
        Call ECASE("Unknown Get Value item: " & CStr(ValueItem))
    End Select
  End If
End Function

Public Function CurrentColumn() As Integer
  If grid.SelStartCol >= 0 Then
    CurrentColumn = grid.SelStartCol
  Else
    CurrentColumn = grid.Col
  End If
End Function

Public Sub FilterBySelection(ByVal IncludeSelection As Boolean)
  Dim acol As AutoCol, ColName As String, ColIndex As Integer
  Dim FilterStr As String
  
  On Error GoTo FilterBySelection_err
  ColIndex = grid.Col
  If Not GetAColByGridIndex(ColIndex, acol) Then Err.Raise ERR_NOFILTER, "FilterBySelection", "Invalid Column selected"
  ColName = acol.GridCaptionClean
  If (Len(grid.SelText) <> Len(grid.Text)) And (acol.dbDataType = TYPE_STR) Then
    FilterStr = FilterStr & WhereClause(acol.DataField, grid.SelText, TYPE_STR, FILTER_LIKE_ALL, , IncludeSelection)
  Else
    FilterStr = FilterStr & WhereClause(acol.DataField, grid.Columns(ColIndex).Value, acol.dbDataType, FILTER_NONE, grid.Text, IncludeSelection)
  End If
  ParentAC.Filter = FilterStr
FilterBySelection_end:
  Exit Sub
  
FilterBySelection_err:
  Call ErrorMessage(ERR_ERROR, Err, "FilterBySelection", "Filter by Selection", "Error filtering column '" & ColName & "' by selection.")
  Resume FilterBySelection_end
  Resume
End Sub

Private Sub CopySelection(Optional ByVal CopyType As COPY_TYPE = CT_UNKNOWN)
  
  m_CopyAllColumnValues = False
  If CopyType = CT_UNKNOWN Then
    CopyType = CT_SEL
    If grid.SelStartCol >= 0 Then
      CopyType = CT_COL
    ElseIf grid.SelBookmarks.Count > 0 Then
      CopyType = CT_ROW
    End If
  End If
  
  If CopyType = CT_ROW Then
    Call CopyClipBoardRow
    m_CopyAllColumnValues = True
  ElseIf CopyType = CT_COL Then
    Call CopyClipBoardColumn
  Else
    If Len(grid.SelText) > 0 Then Call SetAnyClipboardData(grid.SelText)
  End If
End Sub

Private Sub CutSelection(Optional ByVal CopyType As COPY_TYPE = CT_UNKNOWN)
  Dim i As Integer
  
  If grid.AllowDelete And IsDataUpdateable() Then
    If CopyType = CT_UNKNOWN Then
      CopyType = CT_SEL
      If grid.SelStartCol >= 0 Then
        CopyType = CT_COL
      ElseIf grid.SelBookmarks.Count > 0 Then
        CopyType = CT_ROW
      End If
    End If
    If CopyType = CT_ROW Then
      Call CopySelection(CopyType)
      Call DeleteSelection
    ElseIf CopyType = CT_COL Then
      Call Beep
      Call ErrorMessage(ERR_INFO, Err, "Cut", "ERR_CUT", "You cannot cut a column of data in this version of AutoGrid")
    Else
      Call CopySelection(CopyType)
      grid.SelText = ""
    End If
  End If
End Sub

Private Function IsCellPaste() As Boolean
  Dim cString As String
  cString = GetAnyClipboardData(vbCFText)
  IsCellPaste = (Len(cString) > 0) And (InStr(cString, vbCrLf) = 0)
End Function

Private Sub FillGridRowsFromClipboard(cString As String, ByVal nClipRows As Long, ByVal nClipColumns As Long)
  Dim p0 As Long, p1 As Long
  Dim sRow As String, vRowColumns As Variant
  Dim i As Long, j As Long, k As Long
  Dim acol As AutoCol, rs As Recordset
  
  On Error GoTo FillGridRowsFromClipboard_ERR
  
  If m_AutoGrid_ADO Is Nothing Then GoTo FillGridRowsFromClipboard_END
  
  p0 = 1
  p1 = InStr(p0, cString, vbCrLf, vbBinaryCompare)
  If p1 = 0 Then p1 = Len(cString) + 1
  
  Set rs = ADODataGrid.Recordset.Clone
  For i = 1 To grid.SelBookmarks.Count
    rs.Bookmark = grid.SelBookmarks(i - 1)
    
    sRow = Mid$(cString, p0, p1 - p0)
    k = 0
    vRowColumns = Split(sRow, vbTab)
    For j = 1 To ParentAC.Count
      Set acol = ParentAC.Item(m_DisplayOrder(j))
      If acol.ClipboardColumn Then
        rs.Fields(acol.DataField).Value = GetTypedValueDefault(vRowColumns(k), acol.dbDataType, Null)
        k = k + 1
        If k = nClipColumns Then Exit For
      End If
    Next j
    If nClipRows <> 1 Then
      p0 = p1 + Len(vbCrLf)
      p1 = InStr(p0, cString, vbCrLf, vbBinaryCompare)
    End If
    rs.Update
  Next i
FillGridRowsFromClipboard_END:
  Exit Sub
  
FillGridRowsFromClipboard_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "FillGridRowsFromClipboard", "Fill Grid Rows From Clipboard", "Error pasting column values from clipboard")
  Resume FillGridRowsFromClipboard_END
  Resume
End Sub
Private Function CountHeaderCols(cString As String) As Long
  Dim sHeaders As String
  Dim i As Long
  
  On Error GoTo CountHeaderCols_ERR
  i = InStr(1, cString, vbCrLf)
  If i = 0 Then
    sHeaders = cString
  Else
    sHeaders = Left$(cString, i - 1)
  End If
  CountHeaderCols = CountChar(sHeaders, vbTab) + 1

CountHeaderCols_END:
  Exit Function
  
CountHeaderCols_ERR:
  Err.Raise Err.Number, ErrorSource(Err, "CountHeaderCols"), Err.Description
End Function

Public Sub PasteColumnValue(ParamArray Columns())
  Dim i As Integer
  Dim vColumns As Variant
  Dim acol As AutoCol
  
  On Error GoTo PasteColumnValue_ERR
  Call SetCursor
  Call SetClipboardColumn(ParentAC.AutoColumns, False)
  
  vColumns = Columns
  i = GetColumnsFromArray(vColumns)
  If i = -1 Then GoTo PasteColumnValue_END
  Call PasteSelection
  
PasteColumnValue_END:
  Call ClearCursor
  Exit Sub
  
PasteColumnValue_ERR:
  Call ErrorMessage(ERR_ERROR, Err, ErrorSource(Err, "PasteColumnValue"), "Paste Column Value", "Unable to paste column value")
  Resume PasteColumnValue_END
  Resume
End Sub


Private Sub PasteSelection()
  Dim cString  As String, ColSet As TrueOleDBGrid60.Column
  Dim acol As AutoCol, p As Long
  Dim nClipRows As Long, nColumns As Long, nClipColumns As Long
  
  On Error GoTo PasteSelectedRows_err
  If Not IsDataUpdateable() Then Err.Raise ERR_PASTE, "PasteSelection", "Cannot Paste into a grid that is not updateable"
  
  ' deal with column based paste
  If grid.SelBookmarks.Count > 0 And (Not m_CopyAllColumnValues) Then
    If Not grid.AllowUpdate Then Err.Raise ERR_PASTE, "PasteSelection", "Cannot Paste over rows into a grid that does not support Update"
        
    nColumns = CountCopyColumns(ParentAC)
    If nColumns = 0 Then
      ' deal with fillrows functionality
      cString = GetAnyClipboardData(AutoClipHandle)
      If Len(cString) > 0 Then nClipRows = CountChar(cString, vbCrLf, vbBinaryCompare) - 1
      If Len(cString) = 0 Then
        cString = GetAnyClipboardData(vbCFText)
        nClipRows = CountChar(cString, vbCrLf, vbBinaryCompare)
        If (nClipRows = 0) And (Len(cString) > 0) Then nClipRows = 1
      End If
      If Len(cString) > 0 Then
        If (nClipRows <> 1) And (nClipRows <> grid.SelBookmarks.Count) Then Err.Raise ERR_PASTE, "PasteSelection", "Cannot Paste into the grid." & vbCrLf & "The number of rows on clipboard (" & nClipRows & ") must be equal to the number of rows selected (" & grid.SelBookmarks.Count & ")"
        ' set aCol.bClipColumn based on 1st row of private format
        nColumns = GetColumns(ParentAC, CountHeaderCols(cString))
        If nColumns < 0 Then GoTo PasteSelectedRows_end
      End If
    End If
    
    nColumns = CountCopyColumns(ParentAC)
    cString = GetAnyClipboardData(vbCFText)
    If nColumns = 0 Then Err.Raise ERR_PASTE, "PasteSelection", "Unexpected error in paste, no columns marked for column based paste"
    If Len(cString) > 0 Then
      nClipRows = CountChar(cString, vbCrLf, vbBinaryCompare)
      If nClipRows = 0 Then nClipRows = 1
      If (nClipRows <> 1) And (nClipRows <> grid.SelBookmarks.Count) Then Err.Raise ERR_PASTE, "PasteSelection", "Cannot Paste into the grid, number of rows on clipboard (" & nClipRows & ") must be equal to the number of rows selected (" & grid.SelBookmarks.Count & ")"
      nClipColumns = CountHeaderCols(cString)
      If nColumns <> nClipColumns Then Err.Raise ERR_PASTE, "PasteSelection", "Cannot Paste into the grid. The number of columns on clipboard (" & nClipColumns & ") must be equal to the number of columns selected in the grid"
      Call FillGridRowsFromClipboard(cString, nClipRows, nClipColumns)
    End If
  Else
    cString = GetAnyClipboardData(AutoClipHandle)
    If Len(cString) > 0 Then
      If CountChar(cString, vbTab, vbBinaryCompare) > 1 Then
        If Not grid.AllowAddNew Then Err.Raise ERR_PASTE, "PasteSelection", "Cannot Paste rows into a grid that does not support AddNew"
        If CountChar(cString, vbCrLf, vbBinaryCompare) > 1 Then
          Call AddGridRowsFromClipboard(cString, True)
          Exit Sub
        End If
      End If
    End If
    
    cString = GetAnyClipboardData(vbCFText)
    ' deal with single cell paste from Excel
    p = InStr(cString, vbCrLf)
    If (p > 0) And ((p + Len(vbCrLf) - 1) = Len(cString)) Then
      cString = Left$(cString, Len(cString) - Len(vbCrLf))
      p = 0
    End If
    If (p > 0) Or (InStr(cString, vbTab) > 0) Then
      If Not grid.AllowAddNew Then Err.Raise ERR_PASTE, "PasteSelection", "Cannot Paste rows into a grid that does not support AddNew"
      Call AddGridRowsFromClipboard(cString, False)
      Exit Sub
    End If
    If Len(cString) > 0 Then grid.SelText = cString
  End If
  
PasteSelectedRows_end:
  Call SetClipboardColumn(ParentAC.AutoColumns, False)
  Exit Sub
  
PasteSelectedRows_err:
  Call ErrorMessage(ERR_ERROR, Err, "PasteSelection", "Error Pasting rows into grid", "Unable to paste rows into grid")
  Resume PasteSelectedRows_end
  Resume
End Sub

Public Sub CloseGrid(ByVal Repaint As Boolean)
  If Not grid Is Nothing Then Call grid.Close(Repaint)
End Sub

Private Sub AddGridRowsFromClipboard(String1 As String, ByVal PrivateFormat As Boolean)
  Dim ColSet As TrueOleDBGrid60.Column, i As Long, k As Long, cSplit As Integer
  Dim acol As AutoCol
  Dim rs As Recordset, FieldName As String
  Dim Headers() As String, FieldValues() As String, sRow As String, sValue As String
  Dim icol As Long, iRow As Long, MaxCol As Long, maxrow As Long
  Dim p0 As Long, p1 As Long, p As Long, pend As Long, q0 As Long, q1 As Long
  Dim v As Variant, dType As DATABASE_FIELD_TYPES, vbmk As Variant
  Dim bRefreshRequired As Boolean, rCount  As Long
  
  On Error GoTo SetRow_err
  Call SetCursor
  grid.Enabled = False
  If grid.AddNewMode = dbgAddNewCurrent Then
    'apf 25/9 fixed visiblerows number
    If grid.VisibleRows > 1 Then grid.Row = grid.Row - 1
  End If
  
  If PrivateFormat Then
    Set rs = ADODataGrid.Recordset.Clone
    p = InStr(String1, vbCrLf)
    MaxCol = CountChar(Left$(String1, p), vbTab)
    If MaxCol < 1 Then Call Err.Raise(ERR_PASTEROW, "SetRow", "Invalid row headers, cannot paste form clipboard")
    ReDim Preserve Headers(1 To MaxCol) As String
  
    p0 = 1: p1 = InStr(1, String1, vbTab)
    pend = InStr(1, String1, vbCrLf)
    If p1 > 0 And pend > 0 Then
      icol = 0
      Do
        icol = icol + 1
        Headers(icol) = Mid$(String1, p0, p1 - p0)
        p1 = p1 + 1: p0 = p1
        p1 = InStr(p0, String1, vbTab, vbBinaryCompare)
      Loop While p1 < pend
      maxrow = CountChar(String1, vbCrLf) - 1
      ReDim Preserve FieldValues(1 To (MaxCol * maxrow)) As String
      icol = 0
      For iRow = 1 To maxrow
        p0 = pend + 2
        For icol = 1 To MaxCol
          FieldValues(((iRow - 1) * MaxCol) + icol) = Mid$(String1, p0, p1 - p0)
          p1 = p1 + 1: p0 = p1
          p1 = InStr(p0, String1, vbTab, vbBinaryCompare)
        Next icol
        pend = InStr(pend + 1, String1, vbCrLf, vbBinaryCompare)
      Next iRow
      
      'APF 25/9/2002 updated to act like grid paste
      For iRow = 1 To maxrow
        If Not rs Is Nothing Then rs.AddNew
        
        ' set addnew calculated values not derived + booleans
        For i = 0 To (rs.Fields.Count - 1)
          Set acol = GetAColByKey(rs.Fields.Item(i).Name)
          If Not acol Is Nothing Then
            v = Empty
            If Not acol.UnboundColumn Then
              If Not IsEmpty(acol.OnAddNewDefault) Then
                v = GetCalculatedValue(Me, Nothing, acol, acol.OnAddNewDefault, acol.OnAddNewDefaultType)
              ElseIf Not IsEmpty(acol.OnAddNewCalcValue) Then
                If IsSimpleCalc(acol.OnAddNewCalcValueType) Then v = GetCalculatedValue(Me, Nothing, acol, acol.OnAddNewCalcValue, acol.OnAddNewCalcValueType)
              End If
              If Not IsEmpty(v) Then rs.Fields(i).Value = v
            End If
          End If
        Next i
        For icol = 1 To MaxCol
          FieldName = Headers(icol)
          v = FieldValues(((iRow - 1) * MaxCol) + icol)
          Call UpdateFieldValue(rs, FieldName, v)
        Next icol
        
        ' derived calculated values
        For i = 0 To (rs.Fields.Count - 1)
          Set acol = GetAColByKey(rs.Fields.Item(i).Name)
          If Not acol Is Nothing Then
            v = Empty
            If Not acol.UnboundColumn Then
              If Not IsEmpty(acol.OnAddNewCalcValue) Then
                If Not IsSimpleCalc(acol.OnAddNewCalcValueType) Then v = GetCalculatedValue(Me, Nothing, acol, acol.OnAddNewCalcValue, acol.OnAddNewCalcValueType)
              End If
              If Not IsEmpty(v) Then rs.Fields(i).Value = v
            End If
          End If
        Next i
        If Not rs Is Nothing Then rs.Update
      Next iRow
      bRefreshRequired = True
    End If
  Else ' standard format from excel etc..
    ' vbmk = Grid.RowBookmark(cRow)
    grid.MoveLast
    grid.Row = grid.Row + 1
    p0 = 1
    Do
      rCount = rCount + 1
      p1 = InStr(p0, String1, vbCrLf)
      If p1 > 0 Then
        sRow = Mid$(String1, p0, p1 - p0)
      Else
        sRow = Mid$(String1, p0)
      End If
      If Len(sRow) > 0 Then
        q0 = 1
        'o If Not rs Is Nothing Then rs.AddNew
        'o If Not rsRDO Is Nothing Then rsRDO.AddNew
        ' new code
        cSplit = grid.Split
        For k = 0 To (grid.Splits.Count - 1)
          grid.Split = k
          For i = 0 To grid.Columns.Count - 1
            Set ColSet = grid.Columns(i)
            If (ColSet.Width >= GRID_MINCOLWIDTH) And ColSet.visible Then
              ' insert into columns in order
              q1 = InStr(q0, sRow, vbTab)
              If q1 > 0 Then
                sValue = Mid$(sRow, q0, q1 - q0)
              Else
                sValue = Mid$(sRow, q0)
              End If
              'o If Not rs Is Nothing Then rs.Fields(ColSet.DataField) = sValue
              'o If Not rsRDO Is Nothing Then rsRDO.rdoColumns(ColSet.DataField) = sValue
              If IsPasteField(ADODataGrid.Recordset, ParentAC.AutoColumns, ColSet.DataField) Then grid.Columns(i).Value = sValue
              If q1 = 0 Then GoTo NEXT_ROW
              q0 = q1 + 1
            End If
          Next i
        Next k
NEXT_ROW:
        If (p1 <> 0) Or (rCount <> 1) Then
          m_NewRowRefresh = False
          If Not m_InDisplayError Then grid.Update
          bRefreshRequired = True
          grid.Row = grid.Row + 1
        End If
      End If
      p0 = p1 + 2
    Loop Until p1 = 0
  End If
  
SetRow_end:
  grid.Enabled = True
  m_NewRowRefresh = True
  Call ClearCursor
  If bRefreshRequired Then
    Call RefreshGrid(True, True, False, False)
  Else
    'Call TryGridSetFocus
  End If
  Exit Sub
  
SetRow_err:
  If (Err.Number = 3164) Or (Err.Number = 3421) Then Resume Next
  bRefreshRequired = True
  Call ErrorMessagePush(Err)
  Call AbortGridEdit
  If Not rs Is Nothing Then
    If rs.EditMode <> adEditNone Then rs.CancelUpdate
  End If
  Call DisplayGridError("Paste into row", "AddGridRowsFromClipboard")
  'Call ErrorMessagePop(ERR_ERROR, Err, "SetRow", "Paste into row", "Row could not be pasted")
  Resume SetRow_end
  Resume
End Sub

Private Sub UpdateFieldValue(ByVal rs As Recordset, FieldName As Variant, v As Variant)
  Dim dType As DATABASE_FIELD_TYPES
  
  On Error GoTo UpdateFieldValue_ERR
  
  If Not rs Is Nothing Then
    If IsFieldPresentADO(rs.Fields, FieldName) Then
      If IsFieldUpdateableADO(rs, FieldName) Then
        dType = ADOtoDatatype(rs.Fields(FieldName).Type)
        If (dType = TYPE_STR) And (Len(v) = 0) Then v = Null
        Debug.Print "Field: " & FieldName & " Value: " & v
        rs.Fields(FieldName).Value = GetGridTypedValueDefault(v, dType, Null)
      End If
    End If
  End If

UpdateFieldValue_END:
  Exit Sub
UpdateFieldValue_ERR:
  Call Err.Raise(ERR_PASTE, Err, "UpdateFieldValue", "Update Field Value", "Error updating a field value.")
  Resume
End Sub

Public Sub SelectAll()
   Call m_AutoGrid_ADO.SelectAll
End Sub

Private Sub CopyClipBoardRow()
  On Error GoTo CopyClipBoardRow_err
  Call SetCursor
  Call SetClipboardColumn(ParentAC.AutoColumns, True)
  Call CopyClipboardRowEx(ADODataGrid.Recordset.Clone, grid, ParentAC.AutoColumns)
  Call SetClipboardColumn(ParentAC.AutoColumns, False)
  
CopyClipBoardRow_end:
  Call ClearCursor
  Exit Sub
  
CopyClipBoardRow_err:
  Resume CopyClipBoardRow_end
End Sub

Private Function CopyClipBoardColumn()
  Dim rs As Recordset
  Dim i As Integer, FieldKey As String, s As String

  On Error Resume Next
  Call SetCursor
  Set rs = ADODataGrid.Recordset.Clone
  If rs.EOF And rs.BOF Then Exit Function
  rs.MoveFirst
  Do
    For i = grid.SelStartCol To grid.SelEndCol
      FieldKey = ColKey(i)
      If Len(FieldKey) > 0 Then
        s = s & rs.Fields(FieldKey) & vbTab
      End If
    Next i
    s = s & vbCrLf
    rs.MoveNext
  Loop Until rs.EOF
  Call SetAnyClipboardData(s)
  Call ClearCursor
End Function

Private Function TotalColumnAsString(ByRef acol As AutoCol, Optional ByVal ColIndex As Integer = -1) As String
  Dim v As Variant
    
  On Error GoTo TotalColumnAsString_err
  If ColIndex < 0 Then ColIndex = grid.Col
  If GetAColByGridIndex(ColIndex, acol) Then
    If IsNumberField(acol.dbDataType) Then
      TotalColumnAsString = Format$(TotalColumn(ColIndex), acol.Format)
    End If
  End If
  Exit Function
  
TotalColumnAsString_err:
  Call ErrorMessage(ERR_ERROR, Err, "TotalColumnAsString", "Column Total", "Error calculation column total")
End Function

Public Sub TotalColumnDisplay()
  Dim ColTotal As String, acol As AutoCol
    
  ColTotal = TotalColumnAsString(acol)
  If Len(ColTotal) > 0 Then
    Call DisplayMessage(Nothing, "Sum = " & ColTotal, ParentAC.AutoName & " - " & acol.GridCaptionClean, "Ok", "")
  End If
End Sub
Public Function CopyColumnValue(ParamArray Columns()) As Boolean
  Dim i As Integer
  Dim vColumns As Variant
  Dim acol As AutoCol
  
  On Error GoTo CopyColumnValue_ERR
  Call SetCursor
  Call SetClipboardColumn(ParentAC.AutoColumns, False)
  
  vColumns = Columns
  i = GetColumnsFromArray(vColumns)
  If i = -1 Then
    i = GetColumns(ParentAC, -1)
    If i = -1 Then GoTo CopyColumnValue_END
  End If
  m_CopyAllColumnValues = False
  
  Call CopyClipboardRowEx(ADODataGrid.Recordset.Clone, grid, ParentAC.AutoColumns)
  CopyColumnValue = True
  
CopyColumnValue_END:
  Call ClearCursor
  Exit Function
  
CopyColumnValue_ERR:
  Call ErrorMessage(ERR_ERROR, Err, ErrorSource(Err, "CopyColumnValue"), "Copy Column Value", "Error in CopyColumnValue")
  Resume CopyColumnValue_END
  Resume
End Function

Private Function GetColumnsFromArray(vColumns As Variant) As Long
  Dim i As Long
  Dim j As Long
  Dim acol As AutoCol
  
  On Error GoTo GetColumnsFromArray_ERR
  
  GetColumnsFromArray = -1
  If Not IsArrayEx(vColumns) Then GoTo GetColumnsFromArray_END
  
  For i = LBound(vColumns) To UBound(vColumns)
    j = j + 1
    Select Case VarType(vColumns(i))
      Case vbString
        Set acol = GetAColByKey(vColumns(i))
        acol.ClipboardColumn = True
      Case vbLong, vbInteger
        Call GetAColByGridIndex(grid.Col, acol)
        acol.ClipboardColumn = True
      Case Else
        Call Err.Raise(ERR_INVALID_VAR_TYPE, "GetColumnsFromArray", "Invalid Column value datatype passed")
    End Select
  Next i
  GetColumnsFromArray = j
  
GetColumnsFromArray_END:
  Exit Function
  
GetColumnsFromArray_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "GetColumnsFromArray", "Select column", "Unable to select column for column based copy/paste")
  Resume GetColumnsFromArray_END
End Function
Public Function TotalColumnFormat(Optional ByVal ColIndex As Integer = -1) As String
  TotalColumnFormat = TotalColumnAsString(Nothing, ColIndex)
End Function
Public Function TotalColumn(ByVal ColIndex As Integer) As Variant
  Dim Col As Integer
  Dim rs As Recordset
  Dim FieldName As String, acol As AutoCol
  
  On Error GoTo TotalColumn_err
  TotalColumn = Empty
  If ColIndex < 0 Then ColIndex = grid.Col
  If GetAColByGridIndex(ColIndex, acol) Then
    If IsNumberField(acol.dbDataType) Then
      FieldName = acol.DataField
      Set rs = ADODataGrid.Recordset.Clone
      rs.Filter = ADODataGrid.Recordset.Filter
      If Not (rs.BOF And rs.EOF) Then
        rs.MoveFirst
        Do
          TotalColumn = TotalColumn + IsNullEx(rs.Fields(FieldName), 0)
          rs.MoveNext
        Loop Until rs.EOF
      End If
    End If
  End If
  
TotalColumn_end:
  Call TryGridSetFocus
  Set rs = Nothing
  Exit Function
  
TotalColumn_err:
  TotalColumn = vbNullString
  Resume TotalColumn_end
End Function

Public Property Get ColKey(ByVal Index As Integer) As String
  On Error GoTo colkey_err
  ColKey = grid.Columns(Index).DataField
colkey_end:
  Exit Property
  
colkey_err:
  Call ErrorMessage(ERR_ERROR, Err, "ColKey", "Invalid Grid Column", "There is column displayed in the grid which the program can not find." & vbCr & vbCr & "This is unlikely and needs investigating.")
  Resume colkey_end
End Property

Private Sub DeleteSelection()
  Dim rConnectGrid As Boolean, vRecs() As Variant, bMax As Long
  Dim i As Long, rs As Recordset, bDeleteFailed As Boolean, bDoRefresh As Boolean
  
  On Error GoTo DeleteSelection_err
  Call SetCursor
  bDoRefresh = False
  bDeleteFailed = False
  m_RowBMK = Null
  If grid.SelBookmarks.Count > 0 Then
    If grid.AllowDelete And IsDataUpdateable() Then
      bDoRefresh = True
      grid.Bookmark = grid.SelBookmarks(0)
      'm_FirstRowBMK = grid.FirstRow
      m_RowBMK = grid.GetBookmark(-1)
      If IsNull(m_RowBMK) Then
        i = 0
        Do While True
          i = i + 1
          m_RowBMK = grid.GetBookmark(i)
          If IsNull(m_RowBMK) Then Exit Do
          If grid.IsSelected(m_RowBMK) = -1 Then Exit Do
        Loop
      End If
      bMax = grid.SelBookmarks.Count - 1
      ReDim vRecs(0 To bMax)
      For i = 0 To bMax
        vRecs(i) = grid.SelBookmarks(i)
        'If vRecs(i) = m_FirstRowBMK Then m_FirstRowBMK = Null
      Next i
      If bMax > 0 Then
        If Not ConfirmDelete() Then GoTo DeleteSelection_end
        grid.Close False
        rConnectGrid = True
      End If
      Set rs = ADODataGrid.Recordset.Clone
      rs.Filter = vRecs
      If rConnectGrid Then Call RecordAuditValuesAll(rs, True)
      rs.Delete adAffectGroup
    End If
  Else
    If grid.AllowUpdate Then
      If grid.SelLength = 0 Then grid.SelLength = 1
      If Len(grid.SelText) > 0 Then grid.SelText = ""
    End If
  End If
  
DeleteSelection_end:
  If bDoRefresh Then
    If rConnectGrid Then
      Call grid.ReOpen(m_RowBMK)
      If Not mAuditIF Is Nothing Then Call FireAuditRecordUpdated(mAuditIF.AuditLevel)
    Else
      If Not IsNull(m_RowBMK) Then grid.Bookmark = m_RowBMK
    End If
    Call RefreshGrid(True, bDeleteFailed, False, False)
  End If
  
DeleteSelection_end_error:
  Call ClearCursor
  Exit Sub

DeleteSelection_err:
  If bDeleteFailed Then Resume DeleteSelection_end_error
  If Not m_SkipError Then
    Call ErrorMessage(ERR_ERROR, Err, "DeleteSelection", "Deleting selection", "The program has encountered an error while deleting the selected rows/text.")
    m_SkipError = False
  End If
  bDeleteFailed = True
  Resume DeleteSelection_end
  Resume
End Sub


Public Sub SaveWidths()
  Dim ac As AutoCol, ColSet As TrueOleDBGrid60.Column
  Dim k As Long, j As Long
  Dim FieldName As String
  
  On Error GoTo SaveWidths_err
  For Each ac In ParentAC.AutoColumns
    ac.ColWidth = 0
  Next ac
  
  For k = 0 To (grid.Splits.Count - 1)
    grid.Split = k
    For j = 0 To (grid.Columns.Count - 1)
      Set ColSet = grid.Columns(j)
      If ColSet.visible Then
        ParentAC.Item(ColKey(j)).ColWidth = ColSet.Width
      End If
    Next j
  Next k
  
SaveWidths_end:
  Exit Sub
SaveWidths_err:
  Resume SaveWidths_end
End Sub

Public Property Get Display(ByVal Index As Integer) As Long
  Display = m_DisplayOrder(Index)
End Property

' converts field name into grid column or -1 if not in grid
Public Function GetGridColIndex(DataField As String) As Integer
  Dim i As Long
  
  GetGridColIndex = -1
  For i = 0 To grid.Columns.Count - 1
    If StrComp(DataField, grid.Columns.Item(i).DataField, vbTextCompare) = 0 Then
      GetGridColIndex = i
      Exit For
    End If
  Next i
End Function

Public Sub ShowFilterWizard()
  Call ECASE("Not implemented")
End Sub

Private Function IsDataUpdateable() As Boolean
  Dim cn As Connection
  
  On Error Resume Next
  Set cn = ADODataGrid.Recordset.ActiveConnection
  If cn.Mode = adModeUnknown Then
    IsDataUpdateable = True
  Else
    IsDataUpdateable = (cn.Mode And (adModeWrite Or adModeShareDenyNone Or adModeShareExclusive)) <> 0
  End If
End Function

Private Sub EditControl(ByVal EnableEdit As Boolean)
  Dim i As Integer, acol As AutoCol
  Dim ColSetTo As TrueOleDBGrid60.Column

  'Debug.Print "Edit Control:" & EnableEdit
  If m_LastEditControl = CLng(EnableEdit) Then Exit Sub
  m_LastEditControl = CLng(EnableEdit)
  For i = 1 To ParentAC.Count
    Set acol = ParentAC.Item(i)
    If acol.GridColumn >= 0 Then
      Set ColSetTo = grid.Columns(acol.GridColumn)
      If EnableEdit Then
        ColSetTo.Locked = acol.NoAddNew Or (Not IsFieldUpdateable(acol))
      Else
        ColSetTo.Locked = acol.NoEdit
      End If
    End If
  Next i
End Sub

Private Sub Grid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  Dim acol As AutoCol, icol As Integer
  
  On Error GoTo RowColChange_err
  m_CopyAllColumnValues = False
  
  If m_GridActive Then
    'Debug.Print "RowColChange Event"
    icol = grid.Col
    If GetAColByGridIndex(icol, acol) Then
      If IsDataUpdateable() And acol.RefreshRequired Then
        Call RefreshSingleDropDown(acol, True)
      End If
      'Debug.Print "RCC"
      acol.FastKey = ""
      acol.FastKeyPress = ""
      Call DisplayFastKey(acol.FastKey)
      If (grid.AddNewMode = dbgAddNewCurrent) Or (grid.AddNewMode = dbgAddNewPending) Then
        Call EditControl(True)
      ElseIf grid.AddNewMode = dbgNoAddNew Then
        Call EditControl(False)
      End If
    End If
  End If
RowColChange_end:
  Exit Sub
  
RowColChange_err:
  Call ErrorMessage(ERR_ERROR + ERR_ALLOWIGNORE, Err, "RowColChange", "Non Fatal error moving between columns", "Error in RowColChange")
  Resume RowColChange_end
End Sub

Private Sub ShowColumnDebug(ByVal Tag As Long)
  Dim acol As AutoCol, sql As String
  
  If (Tag >= 0) And GetAColByGridIndex(grid.Col, acol) Then
    If Not DisplayMessage(m_AutoGrid_ADO.ContainerForm, "Formats for field [" & acol.DataField & "]" & vbCrLf & acol.FormatString, "Debug formats [" & acol.GridCaptionClean & "]", "Ok", "Copy to Clipboard") Then
      Call SetAnyClipboardData(acol.FormatString)
    End If
  Else
    If Not ADODataGrid.Recordset Is Nothing Then
      sql = ADODataGrid.Recordset.Source
    Else
      sql = "No Source SQL found"
    End If
    If Not DisplayMessage(m_AutoGrid_ADO.ContainerForm, vbCrLf & sql & vbCrLf, "Debug SQL for [" & ParentAC.AutoName & "]", "Ok", "Copy to Clipboard") Then
      Call SetAnyClipboardData(sql)
    End If
  End If
End Sub

Private Sub ClearAudit()
  'On Error GoTo ClearAudit_err
  On Error Resume Next
  mAuditType = GRID_NOEDIT
  If mAuditRS Is Nothing Then
    Set mAuditRS = New Recordset
    'Debug.Print "ClearAudit - createRS"
    ReDim mAuditRSFields(0 To 4) As Variant
    ReDim mAuditRSValues(0 To 4) As Variant
    mAuditRSFields(0) = "SourceField"
    mAuditRSFields(1) = "SourceTable"
    mAuditRSFields(2) = "DataType"
    mAuditRSFields(3) = "OldValue"
    mAuditRSFields(4) = "NewValue"
    mAuditRS.Fields.Append mAuditRSFields(0), adVarChar, 64
    mAuditRS.Fields.Append mAuditRSFields(1), adVarChar, 64
    mAuditRS.Fields.Append mAuditRSFields(2), adInteger
    mAuditRS.Fields.Append mAuditRSFields(3), adVarChar, 255, adFldMayBeNull + adFldIsNullable
    mAuditRS.Fields.Append mAuditRSFields(4), adVarChar, 255, adFldMayBeNull + adFldIsNullable
    Call mAuditRS.Open
  ElseIf Not (mAuditRS.BOF And mAuditRS.EOF) Then
    mAuditRS.MoveFirst
    Do
      mAuditRS.Delete
      mAuditRS.MoveNext
    Loop Until mAuditRS.EOF
  End If
  Exit Sub
  
ClearAudit_err:
  Err.Raise Err.Number, ErrorSource(Err, "ClearAudit"), Err.Description
  Resume
End Sub

Private Sub FireAuditRecordUpdated(ByVal aLevel As GRIDAUDIT_LEVELS)
  On Error Resume Next
  If aLevel <> GRIDAUDIT_NONE Then
    If Not (mAuditRS.EOF And mAuditRS.BOF) Then mAuditRS.MoveFirst
    Call mAuditIF.RecordUpdated(GetNetUser, Now, mAuditType, mAuditRS)
  End If
End Sub

Private Sub RecordAuditValuesAll(ByVal rs As Recordset, ByVal bBeforeDelete As Boolean)
  On Error Resume Next
  If Not (rs.EOF And rs.BOF) Then
    rs.MoveFirst
    Call RecordAuditValues(rs, bBeforeDelete, False)
    rs.MoveNext
  End If
  Do While Not rs.EOF
    Call RecordAuditValues(rs, bBeforeDelete, False, False)
    rs.MoveNext
  Loop
End Sub

Private Sub RecordAuditValues(ByVal rs As Recordset, ByVal bBeforeDelete As Boolean, ByVal bAfterUpdate As Boolean, Optional ByVal bClearAudit As Boolean = True)
  Dim acol As AutoCol, v As Variant, DataField As String
  Dim fld As field, aLevel As GRIDAUDIT_LEVELS
  
  On Error GoTo RecordAuditValues_err
  If mAuditIF Is Nothing Then GoTo RecordAuditValues_end
  aLevel = mAuditIF.AuditLevel
  If aLevel = GRIDAUDIT_NONE Then GoTo RecordAuditValues_end
  If bAfterUpdate Then
    If mAuditType = GRID_ADDNEW Then GoTo RecordAuditValues_end
    'Debug.Print "RecordAuditValues"
    If mAuditRS Is Nothing Then Err.Raise ERR_AUDITSET, "RecordAuditValues", "In RecordAuditValues, however Audit recordset is nothing"
    If Not (mAuditRS.BOF And mAuditRS.EOF) Then mAuditRS.MoveFirst
    'If mAuditRS.BOF And mAuditRS.EOF Then Err.Raise ERR_AUDITSET, "RecordAuditValues", "AfterUpdate flag set but no records to Audit"
    Do While Not mAuditRS.EOF
      DataField = mAuditRS.Fields(mAuditRSFields(0))
      v = rs.Fields(DataField)
      If aLevel = GRIDAUDIT_MINIMUM Then
        Set acol = ParentAC.Item(DataField)
        If IsValueEqual(v, mAuditRS.Fields(mAuditRSFields(3))) And Not acol.PrimaryKey Then GoTo next_auditrecord
      End If
      mAuditRS.Update mAuditRSFields(4), v
next_auditrecord:
      mAuditRS.MoveNext
    Loop
  Else
    If bClearAudit Then Call ClearAudit
    If aLevel = GRIDAUDIT_NONE Then Exit Sub
    If bBeforeDelete Then
      mAuditType = GRID_DELETE
    ElseIf rs.EditMode = adEditAdd Then
      mAuditType = GRID_ADDNEW
    Else
      mAuditType = GRID_EDIT
    End If
    If aLevel = GRIDAUDIT_STATUSONLY Then Exit Sub
    
    For Each fld In rs.Fields
      Set acol = ParentAC.Item(fld.Name)
      If aLevel = GRIDAUDIT_MINIMUM Then
        If ((mAuditType = GRID_ADDNEW) Or (mAuditType = GRID_DELETE)) And Not acol.PrimaryKey Then GoTo next_record
      End If
      mAuditRSValues(0) = acol.SourceField
      mAuditRSValues(1) = acol.SourceTable
      mAuditRSValues(2) = acol.dbDataType
      mAuditRSValues(3) = Null
      mAuditRSValues(4) = Null
      If mAuditType = GRID_DELETE Then
        mAuditRSValues(3) = fld.Value
      ElseIf mAuditType = GRID_EDIT Then
        mAuditRSValues(3) = fld.OriginalValue
      End If
      If mAuditType = GRID_ADDNEW Then
        mAuditRSValues(4) = fld.Value
      End If
      mAuditRS.AddNew mAuditRSFields, mAuditRSValues
next_record:
    Next fld
    
  End If
  
RecordAuditValues_end:
  Exit Sub
  
RecordAuditValues_err:
  Call ErrorMessage(ERR_ERROR, Err, "RecordAuditValues", "Record Audit Values", "Failure recording Audit trail")
  Resume RecordAuditValues_end
  Resume
End Sub

