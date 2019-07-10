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
Private m_AutoGrid_DAO As Object
Private m_AutoGrid_RDO As Object
Private WithEvents DAODataGrid As Data
Attribute DAODataGrid.VB_VarHelpID = -1
Private WithEvents RDODataGrid As MSRDC.MSRDC
Attribute RDODataGrid.VB_VarHelpID = -1
Private WithEvents grid As TDBGrid
Attribute grid.VB_VarHelpID = -1

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

Private m_RDOError As Boolean
Private m_RDOErrorDescription As String

Private m_Popped As Boolean

Private mDeleteMultipleRows As Long
Private mDeleteMultiCancel As Boolean
Private m_SaveComboValue As String

' Cut/Paste column ReOrder
Private StartCol As Long
Private EndCol As Long
Private m_DisplayOrder() As Long
Private m_PastePrivateFormat As Boolean

Private mNotifyIF As IBaseNotify
Private mNotifyDisable As Boolean

'Auto Width helpers
Private mDropDownWidth  As Single
Private mCheckBoxWidth As Single
Private mAutoWidthSettings As AUTOWIDTH_SETTINGS

'Support for Audit interface
Private mAuditIF As iAutoAuditRecord
Private mAuditFields As Collection
Private mAuditType  As GRIDEDIT_TYPE

Private m_CopyAllColumnValues As Boolean

Private Sub Class_Initialize()
  m_InGridRefresh = False
  m_ShowValidateErrors = True
  mNotifyDisable = False
  m_DelayedKeyPressEvent = True
  Debug.Print "  AutoGrid init"
End Sub

Private Sub Class_Terminate()
  m_GridActive = False
  If Not ParentAC Is Nothing Then ParentAC.Kill
  Set ParentAC = Nothing
  Set m_AutoGrid_DAO = Nothing
  Set m_AutoGrid_RDO = Nothing
  Set grid = Nothing
  Debug.Print "  AutoGrid terminate"
End Sub

Friend Sub Kill()
  Set ParentAC = Nothing
  Set GridControl = Nothing
  Set AuditInterface = Nothing
  Set NotifyInterface = Nothing

  Set ParentAC = Nothing
  Set m_AutoGrid_DAO = Nothing
  Set m_AutoGrid_RDO = Nothing
  Set DAODataGrid = Nothing
  Set RDODataGrid = Nothing
  Set grid = Nothing
  Set mNotifyIF = Nothing

  ' Support for Audit interface
  Set mAuditIF = Nothing
  Set mAuditFields = Nothing
  m_GridActive = False
End Sub


Private Sub FillGridRowsFromClipboard(cString As String, ByVal nClipRows As Long, ByVal nClipColumns As Long)
  Dim vbmk As Variant
  Dim p0 As Long, p1 As Long
  Dim sRow As String, vRowColumns As Variant
  Dim i As Long, j As Long, k As Long
  Dim aCol As AutoCol, rs As Recordset
  
  On Error GoTo FillGridRowsFromClipboard_ERR
  If Not m_AutoGrid_RDO Is Nothing Then Err.Raise ERR_PASTE, "FillGridRowsFromClipboard", "Column based copy not implemented for RDO connections"
  If m_AutoGrid_DAO Is Nothing Then GoTo FillGridRowsFromClipboard_END
  grid.Enabled = False
  p0 = 1
  p1 = InStr(p0, cString, vbCrLf, vbBinaryCompare)
  If p1 = 0 Then p1 = Len(cString) + 1
  
  Set rs = DAODataGrid.Recordset
  If Not (rs.BOF And rs.EOF) Then vbmk = rs.Bookmark
  For i = 1 To grid.SelBookmarks.Count
    rs.Bookmark = grid.SelBookmarks(i - 1)
    rs.Edit
    sRow = Mid$(cString, p0, p1 - p0)
    
    k = 0
    vRowColumns = Split(sRow, vbTab)
    For j = 1 To ParentAC.Count
      Set aCol = ParentAC.Item(m_DisplayOrder(j))
      If aCol.ClipboardColumn Then
        rs.Fields(aCol.DataField).Value = GetTypedValueDefault(vRowColumns(k), aCol.dbDataType, Null)
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
  grid.Enabled = True
  Exit Sub
  
FillGridRowsFromClipboard_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "FillGridRowsFromClipboard", "Fill Grid Rows From Clipboard", "Error pasting column values from clipboard")
  Resume FillGridRowsFromClipboard_END
  Resume
End Sub

Private Sub UpdateFieldValue(ByVal rsRDO As RDOResultset, ByVal rs As Recordset, FieldName As Variant, v As Variant)
  Dim dType As DATABASE_FIELD_TYPES
  
  On Error GoTo UpdateFieldValue_ERR
  
  If Not rsRDO Is Nothing Then
    If FieldPresentRDO(rsRDO.rdoColumns, FieldName) Then
      If rsRDO.rdoColumns(FieldName).Updatable Then
        dType = RDOtoDatatype(rsRDO.rdoColumns(FieldName).Type)
        If (dType = TYPE_STR) And (Not rsRDO.rdoColumns(FieldName).AllowZeroLength) And (Len(v) = 0) Then v = Null
        rsRDO.rdoColumns(FieldName).Value = GetGridTypedValueDefault(v, dType, Null)
      End If
    End If
  End If
  If Not rs Is Nothing Then
    If FieldPresent(rs.Fields, FieldName) Then
      If IsFieldUpdateableDAO(rs, FieldName, rs.Fields(FieldName).SourceField) Then
        dType = DAOtoDatatype(rs.Fields(FieldName).Type)
        If (dType = TYPE_STR) And (Not rs.Fields(FieldName).AllowZeroLength) And (Len(v) = 0) Then v = Null
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

Private Sub TryGridSetFocus()
  On Error Resume Next
  Call grid.SetFocus
End Sub

Friend Property Get AuditInterface() As iAutoAuditRecord
  Set AuditInterface = mAuditIF
End Property

Friend Property Set AuditInterface(ByVal NewValue As iAutoAuditRecord)
  If Not NewValue Is Nothing Then
    If NewValue.AuditLevel = GRIDAUDIT_MINIMUM Then Err.Raise 380, "AuditInterface", "Audo Grid does not currently implement support for GRIDAUDIT_MINIMUM"
  End If
  Set mAuditIF = NewValue
End Property

Friend Property Get NotifyInterface() As IBaseNotify
  Set NotifyInterface = mNotifyIF
End Property

Friend Property Set NotifyInterface(ByVal NewValue As IBaseNotify)
  Set mNotifyIF = NewValue
End Property

Public Property Set GridControl(NewValue As Object) ''apf Byval
  Set m_AutoGrid_DAO = Nothing
  Set m_AutoGrid_RDO = Nothing
  If Not NewValue Is Nothing Then
    If StrComp(TypeName(NewValue), "AutoGridCtrl_DAO", vbTextCompare) = 0 Then Set m_AutoGrid_DAO = NewValue
    If StrComp(TypeName(NewValue), "AutoGridCtrl_RDO", vbTextCompare) = 0 Then Set m_AutoGrid_RDO = NewValue
  End If
  If (m_AutoGrid_DAO Is Nothing) And (m_AutoGrid_RDO Is Nothing) Then
    Set DAODataGrid = Nothing
    Set RDODataGrid = Nothing
    Set grid = Nothing
  Else
    Call SetGrid
  End If
End Property

Public Property Get GridControl() As Object
  If Not m_AutoGrid_DAO Is Nothing Then Set GridControl = m_AutoGrid_DAO
  If Not m_AutoGrid_RDO Is Nothing Then Set GridControl = m_AutoGrid_RDO
End Property

Public Property Get TDBGrid() As Object
  Set TDBGrid = grid
End Property

Public Property Get DAORecordset() As Recordset
  If Not m_AutoGrid_DAO Is Nothing Then Set DAORecordset = DAODataGrid.Recordset
End Property

Public Property Get RDOResultset() As RDOResultset
  If Not m_AutoGrid_RDO Is Nothing Then Set RDOResultset = RDODataGrid.Resultset
End Property

Public Property Get ShowValidationErrors() As Boolean
  ShowValidationErrors = m_ShowValidateErrors
End Property

Public Property Let ShowValidationErrors(ByVal NewValue As Boolean)
  m_ShowValidateErrors = NewValue
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
  Dim aCol As AutoCol, dUpdateable As Boolean
  Dim i As Long, ComboCount As Long
  
  On Error GoTo InitGrid_err
  StartCol = -1
  m_GridActive = False
  m_UnboundColumns = False
  If ParentAC.CheckGridDate Then Call CheckGridDate
  dUpdateable = IsDataUpdateable()
  If Not SetGrid Then Err.Raise ERR_SETGRID, "InitGrid", "Auto grid type is not recognised. TCSAUTO and TCSACTRL are out of sync"
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
    Set aCol = ParentAC.Item(i)
    m_DisplayOrder(i) = CLng(GetIniEntry("Order_" & ParentAC.AutoName, CStr(i), "0"))
    If (m_DisplayOrder(i) < 1) Or (m_DisplayOrder(i) > ParentAC.Count) Then m_DisplayOrder(i) = 0
    aCol.ColWidth = CLng(GetIniEntry(ParentAC.AutoName, aCol.DataField, "-1"))
    If (aCol.DropType And DROPDOWN_BOUND) = DROPDOWN_BOUND Then
      If Not m_AutoGrid_DAO Is Nothing Then aCol.QueryCombo = m_AutoGrid_DAO.AddDataCombo
      If Not m_AutoGrid_RDO Is Nothing Then aCol.QueryCombo = m_AutoGrid_RDO.AddDataCombo
    End If
    If aCol.UnboundColumn Then m_UnboundColumns = True
  Next i
  Call EnsureDisplayOrderValid
  Call ReBindColumns
  Call SetAutoWidthDefaults
  ' apf removed PK stuff
  Call TryGridSetFocus
  Exit Sub
  
InitGrid_err:
  ParentAC.SaveColumns = False
  Err.Raise Err.Number, ErrorSource(Err, "InitGrid"), Err.Description
  Resume
End Sub

Public Property Let AutoWidthSetting(ByVal NewValue As AUTOWIDTH_SETTINGS)
  mAutoWidthSettings = NewValue
End Property

Private Sub SetAutoWidthDefaults()
  Dim pic As PictureBox
  
  If Not m_AutoGrid_DAO Is Nothing Then Set pic = m_AutoGrid_DAO.picTest
  If Not m_AutoGrid_RDO Is Nothing Then Set pic = m_AutoGrid_RDO.picTest
  mDropDownWidth = pic.ScaleX(GetSystemMetrics(SM_CXHSCROLL), vbPixels, pic.ScaleMode)
  mCheckBoxWidth = 270
  If mAutoWidthSettings = 0 Then mAutoWidthSettings = AW_INCLUDE_CAPTIONS
End Sub

Public Sub AutoWidth(Optional ByVal ColIndex As Integer = -1)
  Dim pic As PictureBox
  Dim ColWidths() As Single, cWidth As Single
  Dim ColSet As Column, aCol As AutoCol
  Dim vbmk As Variant
  Dim iRow As Long, j As Long
  Const INVALID_ROW_NUM As Long = -2 ^ 31
  Const COLWIDTH_OFFSET As Single = (GRID_MINCOLWIDTH * 2)
  
  On Error GoTo AutoWidth_err
  If grid.Columns.Count = 0 Then GoTo AutoWidth_end
  If Not m_AutoGrid_DAO Is Nothing Then Set pic = m_AutoGrid_DAO.picTest
  If Not m_AutoGrid_RDO Is Nothing Then Set pic = m_AutoGrid_RDO.picTest
  
  If ColIndex >= 0 Then
    ReDim ColWidths(ColIndex To ColIndex)
  Else
    ReDim ColWidths(0 To (grid.Columns.Count - 1))
  End If
  
  ' Initialise ColWidths to the column header widths
  For j = 0 To grid.Columns.Count - 1
    Set ColSet = grid.Columns(j)
    If GetAColByGridIndex(j, aCol) Then
      If Not aCol.Hide Then
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
  Next j
  
  Set ColSet = Nothing
  If ColIndex >= 0 Then Set ColSet = grid.Columns(ColIndex)
      
  If Not ColSet Is Nothing Then Set pic.Font = ColSet.Font
  ' retrieve First Grid bookmark
  iRow = 0
  Do While iRow < grid.VisibleRows
    vbmk = grid.RowBookmark(iRow)
    If Len(vbmk) = 0 Then GoTo NEXT_ROW
    If ColIndex >= 0 Then
      cWidth = pic.TextWidth(Trim$(ColSet.CellText(vbmk)))
      If cWidth > ColWidths(ColIndex) Then ColWidths(ColIndex) = cWidth
    Else
      For j = 0 To grid.Columns.Count - 1
        Set ColSet = grid.Columns(j)
        If GetAColByGridIndex(j, aCol) Then
          If Not aCol.Hide Then
            Set pic.Font = ColSet.Font
            If aCol.IsCheckBox Then
              cWidth = mCheckBoxWidth
            Else
              cWidth = pic.TextWidth(Trim$(ColSet.CellText(vbmk)))
              If aCol.DropType <> DROPDOWN_NONE Then cWidth = cWidth + mDropDownWidth
            End If
            If cWidth > ColWidths(j) Then ColWidths(j) = cWidth
          End If
        End If
      Next j
    End If
NEXT_ROW:
    iRow = iRow + 1
  Loop
  If ColIndex >= 0 Then
    ColSet.Width = ColWidths(ColIndex) + COLWIDTH_OFFSET
  Else
    For j = 0 To grid.Columns.Count - 1
      Set ColSet = grid.Columns(j)
      If ColWidths(j) > 0 Then
        ColSet.Width = ColWidths(j) + COLWIDTH_OFFSET
        ColSet.visible = True
      End If
    Next j
  End If
AutoWidth_end:
  Call CreateGridSplit(False, True)
  Exit Sub
  
AutoWidth_err:
  Resume AutoWidth_end
  Resume
End Sub

Private Function SetGrid() As Boolean
  On Error Resume Next
  If grid Is Nothing Then
    If Not m_AutoGrid_DAO Is Nothing Then
      Set grid = m_AutoGrid_DAO.grid
      Set DAODataGrid = m_AutoGrid_DAO.GridDataControl
      If Not AutoControlRegistered Then Call RegisterLibrary(m_AutoGrid_DAO)
      AutoControlRegistered = True
    End If
    If Not m_AutoGrid_RDO Is Nothing Then
      Set grid = m_AutoGrid_RDO.grid
      Set RDODataGrid = m_AutoGrid_RDO.GridDataControl
      If Not AutoControlRegistered Then Call RegisterLibrary(m_AutoGrid_RDO)
      AutoControlRegistered = True
    End If
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

Private Sub DAODataGrid_Reposition()
  Dim cur As Long, max As Long
  On Error Resume Next
  cur = DAODataGrid.Recordset.AbsolutePosition
  max = DAODataGrid.Recordset.RecordCount
  DAODataGrid.Caption = CStr(cur + 1) & " of " & CStr(max)
  If (Not mNotifyIF Is Nothing) And (Not mNotifyDisable) Then Call mNotifyIF.Notify(cur + 1, max, "Reposition")
End Sub

'Private Sub RestoreColumnFormat(ByVal ColIndex As Integer)
'  Dim aCol As AutoCol, ColSet As TrueDBGrid60.Column
'
'  If GetAColByGridIndex(ColIndex, aCol) Then
'    If (aCol.dbDataType = TYPE_DOUBLE) Or (aCol.dbDataType = TYPE_LONG) Then
'      Set ColSet = Grid.Columns(ColIndex)
'      ColSet.NumberFormat = aCol.Format
'    End If
'  End If
'End Sub

Private Sub Grid_AfterColUpdate(ByVal ColIndex As Integer)
  Dim aCol As AutoCol
  
  'Call RestoreColumnFormat(ColIndex)
  If GetAColByGridIndex(ColIndex, aCol) Then
    If aCol.OnChangeEvent Then
      RaiseEvent OnColumnChange(ParentAC.AutoName, aCol.DataField)
    End If
  End If
End Sub

Private Sub CheckDropDown(ByVal aCol As AutoCol)
  Dim sField As String
  On Error GoTo CheckDropDown_err
  
  If aCol Is Nothing Then Call GetAColByGridIndex(grid.Col, aCol)
  sField = "Unknown"
  If IsDataUpdateable() And aCol.RefreshRequired Then
    sField = aCol.GridCaptionClean
    aCol.RefreshRequired = Not RefreshSingleDropDown(aCol, False)
  End If
  Exit Sub
  
CheckDropDown_err:
  Call ErrorMessage(ERR_ERROR, Err, "CheckDropDown", "Failed to refresh field dropdown", "Unable to refresh dropdown in field " & sField)
  If Not aCol Is Nothing Then aCol.RefreshRequired = True
End Sub

Private Sub Grid_ButtonClick(ByVal ColIndex As Integer)
  Dim aCol As AutoCol
  
  If m_GridActive Then
    If GetAColByGridIndex(ColIndex, aCol) Then
      Call CheckDropDown(aCol)
      If aCol.HasButton Then RaiseEvent ButtonClick(ParentAC.AutoName, aCol.DataField)
    End If
  End If
End Sub


'Private Sub Grid_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
'  Dim aCol As AutoCol, ColSet As TrueDBGrid60.Column
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
  Call DisplayGridError("UpdateGrid")
End Function

Private Sub Grid_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
  Dim aCol As AutoCol
  
  On Error GoTo Grid_FormatText_err
  If GetAColByGridIndex(ColIndex, aCol) Then Value = aCol.GetFormattedValue(Value)
  Exit Sub
  
Grid_FormatText_err:
  Value = "Error in Grid FormatText " & Err.Description
End Sub

Private Sub Grid_LostFocus()
  Static InLoseFocus As Boolean
  
  If InLoseFocus Then Exit Sub
  InLoseFocus = True
  Call UpdateGrid(True)
  
  InLoseFocus = False
  Exit Sub
End Sub

Private Sub Grid_OnAddNew()
  Call SetGridDefaults   'Call Grid.Refresh causes ghost dates to appear 31/12/99
End Sub

Private Sub Grid_PostEvent(ByVal MsgId As Integer)
  Dim aCol As AutoCol
  
  Select Case MsgId
    Case MSG_REFRESH_GRID
      Call RefreshGrid(False, True, False)
    Case MSG_KEYPRESS_DELAY
      m_DelayedKeyPressEvent = False
      Call GridSpeedKey(0, True)
      m_DelayedKeyPressEvent = True
    Case MSG_COMBO_RESELECT
      grid.Text = m_SaveComboValue
      SendKeys "{ENTER}"
  End Select
End Sub

Private Sub Grid_UnboundColumnFetch(Bookmark As Variant, ByVal Col As Integer, Value As Variant)
  Dim sColName As String
  Dim aCol As AutoCol, v As Variant
    
  On Error GoTo UnboundColumnFetch_err
  If GetAColByGridIndex(Col, aCol) Then
    sColName = aCol.GridCaptionClean
    Value = GetCalculatedValue(Me, Nothing, aCol, aCol.OnUpdateCalcValue, aCol.OnUpdateCalcValueType, Bookmark, True)
  End If
  Exit Sub
  
UnboundColumnFetch_err:
  Err.Raise Err.Number, ErrorSource(Err, "UnboundColumnFetch"), Err.Description
End Sub

Private Sub RDODataGrid_Error(ByVal Number As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  m_RDOErrorDescription = m_RDOErrorDescription & Description & vbCrLf
  m_RDOError = True
  CancelDisplay = True
End Sub

Private Sub RDODataGrid_Reposition()
  Dim cur As Long, max As Long
  On Error Resume Next
  cur = RDODataGrid.Resultset.AbsolutePosition
  max = RDODataGrid.Resultset.RowCount
  RDODataGrid.Caption = CStr(cur) & " of " & CStr(max)
  If (Not mNotifyIF Is Nothing) And (Not mNotifyDisable) Then Call mNotifyIF.Notify(cur, max, "Reposition")
End Sub

Private Sub Grid_FetchCellTips(ByVal SplitIndex As Integer, ByVal ColIndex As Integer, ByVal RowIndex As Long, CellTip As String, ByVal FullyDisplayed As Boolean, ByVal TipStyle As TrueDBGrid60.StyleDisp)
  Dim aCol As AutoCol
  
  CellTip = ""
  If GetAColByGridIndex(ColIndex, aCol) Then
    If Len(aCol.ToolTip) Then CellTip = aCol.ToolTip
  End If
End Sub

Private Sub SetGridDefaults()
  Static InSetDefault As Boolean
  Dim i As Integer, aCol As AutoCol, v As Variant
  Dim ColSetTo As TrueDBGrid60.Column

  On Error GoTo SetGridDefaults_err
  If InSetDefault Then Exit Sub
  InSetDefault = True
  For i = 1 To ParentAC.Count
    Set aCol = ParentAC.Item(i)
    If aCol.GridColumn >= 0 Then
      Set ColSetTo = grid.Columns(aCol.GridColumn)
      If Not IsEmpty(aCol.OnAddNewDefault) And ((Len(ColSetTo.Text) = 0) Or IsEmpty(ColSetTo.Value)) Then
        v = GetCalculatedValue(Me, Nothing, aCol, aCol.OnAddNewDefault, aCol.OnAddNewDefaultType)
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

Private Sub SetAColCellValues(ByVal rs As Recordset, ByVal rsRDO As RDOResultset)
  Dim errstring As String
  Dim v As Variant
  Dim i As Long, ColSet As TrueDBGrid60.Column, aCol As AutoCol
  
  On Error GoTo SetAColCellValues_err
  For i = 1 To ParentAC.Count
    Set aCol = ParentAC.Item(i)
    aCol.CellValue = Empty
    If aCol.UnboundColumn Then GoTo next_column
    If (aCol.GridColumn >= 0) And Not m_PastePrivateFormat Then
      Set ColSet = grid.Columns(aCol.GridColumn)
      If (Not IsEmpty(ColSet.Value)) And (Len(ColSet.Value) > 0) Then
        aCol.CellValue = GetGridTypedValueDefault(ColSet.Value, aCol.dbDataType, Empty)
        If IsEmpty(aCol.CellValue) And (((2 ^ (aCol.dbDataType + 1)) And ParentAC.ValidateDataTypes) <> 0) Then
          Err.Raise ERR_INVALID_CELLVALUE, "SetAColCellValues", "The value '" & ColSet.Value & "' is not a valid entry in the column " & aCol.GridCaptionClean & "." & vbCrLf & "The value must be of type " & DataTypeName(aCol.dbDataType)
        End If
      End If
    Else
      If Not rs Is Nothing Then aCol.CellValue = rs.Fields(aCol.DataField)
      If Not rsRDO Is Nothing Then aCol.CellValue = rsRDO.rdoColumns(aCol.DataField)
    End If
next_column:
    If IsNull(aCol.CellValue) Then aCol.CellValue = Empty
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
  Dim i As Integer, aCol As AutoCol, aColSrc As AutoCol
  Dim ColSet As TrueDBGrid60.Column
  Dim FieldName As String
  Dim rs As Recordset, rsRDO As RDOResultset, v As Variant
  Static iLevel As Long
     
  On Error GoTo Grid_BeforeUpdate_err
  iLevel = iLevel + 1
  Call AutoLog("BEGIN BF " & iLevel)
  m_SkipError = False
  m_ForcePopError = False
  Set mAuditFields = Nothing
  mAuditType = GRID_EDIT
  
  ' bmk = Grid.Bookmark note grid bookmark does not!! refer to the current record in addnew buffer
  ' If IsEmpty(bmk) Then Err.Raise ERR_GRID_BEFOREUPDATE, "Grid_BeforeUpdate", "Unable to obtain a bookmark for the current row"
  If Not m_AutoGrid_DAO Is Nothing Then
    Set rs = DAODataGrid.Recordset
    If rs.EditMode = dbEditAdd Then
      mUpdateType = GRID_ADDNEW
    Else
      mUpdateType = GRID_EDIT
    End If
  End If
  If Not m_AutoGrid_RDO Is Nothing Then
    Set rsRDO = RDODataGrid.Resultset
    If rsRDO.EditMode = rdEditAdd Then
      mUpdateType = GRID_ADDNEW
    Else
      mUpdateType = GRID_EDIT
    End If
  End If
  Call SetAColCellValues(rs, rsRDO)
    
  ' Audit Updates - prepare record current values of recordset
  If Not mAuditIF Is Nothing Then Call RecordAuditValues(rs, rsRDO, False)
  
  ' calculated values not derived + booleans
  For i = 1 To ParentAC.Count
    Set aCol = ParentAC.Item(i)
    v = Empty
    If Not aCol.UnboundColumn Then
      If mUpdateType = GRID_ADDNEW Then
        If Not IsEmpty(aCol.OnAddNewCalcValue) Then
          If IsSimpleCalc(aCol.OnAddNewCalcValueType) Then v = GetCalculatedValue(Me, Nothing, aCol, aCol.OnAddNewCalcValue, aCol.OnAddNewCalcValueType)
        End If
      ElseIf mUpdateType = GRID_EDIT Then
        If Not IsEmpty(aCol.OnUpdateCalcValue) Then
          If IsSimpleCalc(aCol.OnUpdateCalcValueType) Then v = GetCalculatedValue(Me, Nothing, aCol, aCol.OnUpdateCalcValue, aCol.OnUpdateCalcValueType)
        End If
      End If
      If Not IsEmpty(v) Then
        If Not m_AutoGrid_DAO Is Nothing Then rs.Fields(aCol.DataField) = v
        If Not m_AutoGrid_RDO Is Nothing Then rsRDO.rdoColumns(aCol.DataField) = v
        aCol.CellValue = v
      End If
      
      If aCol.GridColumn >= 0 Then
        v = Empty
        If aCol.dbDataType = TYPE_BOOL Then
          If IsEmpty(aCol.CellValue) Then
            v = aCol.OnNullValue
          Else
            v = CBoolean(aCol.CellValue)
          End If
        End If
        If (aCol.dbDataType = TYPE_DOUBLE) Or (aCol.dbDataType = TYPE_LONG) Or (aCol.dbDataType = TYPE_DATE) Then
          If IsEmpty(aCol.CellValue) Then v = aCol.OnNullValue
        End If
        If Not IsEmpty(v) Then
          Set ColSet = grid.Columns(aCol.GridColumn)
          ColSet.Value = v
          ColSet.Text = v
          If Not m_AutoGrid_DAO Is Nothing Then rs.Fields(aCol.DataField) = v
          If Not m_AutoGrid_RDO Is Nothing Then rsRDO.rdoColumns(aCol.DataField) = v
          aCol.CellValue = v
        End If
      End If
    End If
  Next i

  ' derived calculated values
  For i = 1 To ParentAC.Count
    Set aCol = ParentAC.Item(i)
    v = Empty
    If Not aCol.UnboundColumn Then
      If mUpdateType = GRID_ADDNEW Then
        If Not IsEmpty(aCol.OnAddNewCalcValue) Then
          If Not IsSimpleCalc(aCol.OnAddNewCalcValueType) Then
            v = GetCalculatedValue(Me, Nothing, aCol, aCol.OnAddNewCalcValue, aCol.OnAddNewCalcValueType)
          End If
        End If
      ElseIf mUpdateType = GRID_EDIT Then
        If Not IsEmpty(aCol.OnUpdateCalcValue) Then
          If Not IsSimpleCalc(aCol.OnUpdateCalcValueType) Then v = GetCalculatedValue(Me, Nothing, aCol, aCol.OnUpdateCalcValue, aCol.OnUpdateCalcValueType)
        End If
      End If
      If Not IsEmpty(v) Then
        If Not m_AutoGrid_DAO Is Nothing Then rs.Fields(aCol.DataField) = v
        If Not m_AutoGrid_RDO Is Nothing Then rsRDO.rdoColumns(aCol.DataField) = v
        aCol.CellValue = v
      End If
    End If
  Next i

  ' CopyFromField - set the value of another field get Source Field from grid first
  If mUpdateType = GRID_ADDNEW Then
    For i = 1 To ParentAC.Count
      Set aCol = ParentAC.Item(i)
      If Len(aCol.CopyFromField) > 0 Then
        Set aColSrc = GetAColByKey(aCol.CopyFromField)
        If Not aColSrc Is Nothing Then
          v = aColSrc.CellValue
        Else ' old behaviour - can be static database field apf see Bug entry 26/2/99
          If Not m_AutoGrid_DAO Is Nothing Then v = rs.Fields(aCol.CopyFromField)
          If Not m_AutoGrid_RDO Is Nothing Then v = rsRDO.rdoColumns(aCol.CopyFromField)
        End If
        If Not m_AutoGrid_DAO Is Nothing Then rs.Fields(aCol.DataField) = v
        If Not m_AutoGrid_RDO Is Nothing Then rsRDO.rdoColumns(aCol.DataField) = v
        aCol.CellValue = v
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
  Dim aCol As AutoCol
  Dim dType As DATABASE_FIELD_TYPES, v As Variant, i As Long
  Dim errstring  As String, dbType As DATABASE_FIELD_TYPES
  
  On Error GoTo ValidRow_DataTypes_err
  
  For i = 1 To ParentAC.Count
    Set aCol = ParentAC.Item(i)
    If (aCol.GridColumn >= 0) And Not aCol.UnboundColumn Then
      v = aCol.CellValue
      dbType = aCol.dbDataType
      If Not ((dbType <> TYPE_STR) And (Len(v) = 0)) Then Call GetTypedValue(v, dbType)
      If Not IsEmpty(aCol.MinValidValue) Then
        If v < aCol.MinValidValue Then Err.Raise ERR_GRIDVALIDATE, "ValidateGridRow", "The Value in Column '" & aCol.GridCaptionClean & _
                                                 "' must be equal to or greater than " & GetValueAsString(aCol.MinValidValue, dbType)
      End If
      If Not IsEmpty(aCol.MaxValidValue) Then
        If v > aCol.MaxValidValue Then Err.Raise ERR_GRIDVALIDATE, "ValidateGridRow", "The Value in Column '" & aCol.GridCaptionClean & _
                                                 "' must be equal to or less than " & GetValueAsString(aCol.MaxValidValue, dbType)
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
  If Not aCol Is Nothing Then
    errstring = "The Column '" & aCol.GridCaptionClean & "' cannot contain the value: " & IsNullEx(v, "(Null)") & vbCrLf & _
                "Column is of type " & DataTypeName(dbType)
  Else
    errstring = "Grid Row could not be validated"
  End If
  Err.Raise ERR_GRIDVALIDATE, "ValidateGridRow", errstring
End Sub


Private Sub Grid_BeforeDelete(Cancel As Integer)
  Dim frm As Form
  Dim rs As Recordset, rsRDO As RDOResultset
  
  If Len(ParentAC.ConfirmDeleteText) > 0 Then
    If mDeleteMultipleRows > 0 Then Cancel = mDeleteMultiCancel
    If Not Cancel Then
      If (mDeleteMultipleRows = 0) Then
        If Not m_AutoGrid_DAO Is Nothing Then Set frm = m_AutoGrid_DAO.ContainerForm
        If Not m_AutoGrid_RDO Is Nothing Then Set frm = m_AutoGrid_RDO.ContainerForm
        mDeleteMultiCancel = Not DisplayMessage(frm, ParentAC.ConfirmDeleteText, "Confirm Delete", "Yes", "No")
      End If
      Cancel = mDeleteMultiCancel
    End If
    If Cancel Then
      m_SkipError = True
      Exit Sub
    End If
  End If
  ' record current values of recordset
  If Not mAuditIF Is Nothing Then
    If Not m_AutoGrid_DAO Is Nothing Then Set rs = DAODataGrid.Recordset
    If Not m_AutoGrid_RDO Is Nothing Then Set rsRDO = RDODataGrid.Resultset
    Call RecordAuditValues(rs, rsRDO, True)
  End If
End Sub

Private Sub Grid_AfterDelete()
  Call DoDBEvents
  ' record current values of recordset
  If Not mAuditIF Is Nothing Then
    If mAuditIF.AuditLevel <> GRIDAUDIT_NONE Then
      Call mAuditIF.RecordUpdated(GetNetUser, Now, mAuditType, mAuditFields)
    End If
    Set mAuditFields = Nothing
  End If
End Sub

Private Sub Grid_AfterUpdate()
  Dim i As Long, ColSet As TrueDBGrid60.Column, aCol As AutoCol
  Dim iFld As FieldDetails, rbmk As Variant
  Dim rs As Recordset, rsRDO As RDOResultset
  Dim OldNotifyDisable  As Boolean
  
  On Error GoTo AfterUpdate_err
  OldNotifyDisable = mNotifyDisable
  mNotifyDisable = True
  Call DoDBEvents
  If Not mAuditIF Is Nothing Then ' record current values of recordset
    If (mAuditIF.AuditLevel = GRIDAUDIT_FULL) Or (mAuditIF.AuditLevel = GRIDAUDIT_MINIMUM) Then
      If Not m_AutoGrid_DAO Is Nothing Then
        Set rs = DAODataGrid.Recordset
        If Not (rs.EOF Or rs.BOF) Then rbmk = rs.Bookmark
        rs.Bookmark = rs.LastModified
      End If
      If Not m_AutoGrid_RDO Is Nothing Then
        Set rsRDO = RDODataGrid.Resultset
        ' RDO has not commited update at this stage - cannot use LastModified property
      End If
      For Each iFld In mAuditFields
        If Not m_AutoGrid_DAO Is Nothing Then iFld.Value = rs.Fields(iFld.Name).Value
        If Not m_AutoGrid_RDO Is Nothing Then iFld.Value = rsRDO.rdoColumns(iFld.Name).Value
      Next iFld
    End If
    If mAuditIF.AuditLevel <> GRIDAUDIT_NONE Then
      Call mAuditIF.RecordUpdated(GetNetUser, Now, mAuditType, mAuditFields)
    End If
    Set mAuditFields = Nothing
    If Not IsEmpty(rbmk) Then
      If Not m_AutoGrid_DAO Is Nothing Then rs.Bookmark = rbmk
      If Not m_AutoGrid_RDO Is Nothing Then rsRDO.Bookmark = rbmk
    End If
  End If
  
  If (Not m_AutoGrid_RDO Is Nothing) And (mUpdateType = GRID_ADDNEW) Then Call grid.PostMsg(MSG_REFRESH_GRID)
  mUpdateType = GRID_NOEDIT
  mNotifyDisable = OldNotifyDisable
  Exit Sub
  
AfterUpdate_err:
  mNotifyDisable = OldNotifyDisable
  Err.Raise Err.Number, ErrorSource(Err, "AfterUpdate"), Err.Description
  Resume
End Sub

Private Sub Grid_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
  Dim aCol As AutoCol
  On Error GoTo ColResize_err
    
  ParentAC.SaveColumns = True
  If GetAColByGridIndex(ColIndex, aCol) Then Call AutoSizeSingleDropDown(aCol)
  Exit Sub
  
ColResize_err:
  Call ErrorMessage(ERR_ERROR + ERR_ALLOWIGNORE, Err, "ColResize", "Non Fatal error resizing column", "Error in Column resize")
End Sub

Private Function GetColumnName(ByVal table As String, ByVal field As String) As String
  Dim aCol As AutoCol
  Dim i As Long
  On Error GoTo GetColumnName_err
    
  GetColumnName = field
  For i = 1 To ParentAC.Count
    Set aCol = ParentAC.Item(i)
    If StrComp(aCol.SourceTable, table, vbTextCompare) = 0 Then
      If StrComp(aCol.SourceField, field, vbTextCompare) = 0 Then
        field = aCol.GridCaptionClean
        Exit For
      End If
    End If
  Next i
  GetColumnName = field
  
GetColumnName_end:
  Exit Function
  
GetColumnName_err:
  GetColumnName = field
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

Private Sub DisplayGridError(ByVal SourceFunction As String)
  Const FieldString As String = "field '"
  Const SetForString As String = "set for '"
  Const PK_ValidationError As String = "Index or primary key can't contain a Null value."
  Const PKdup_ValidationError As String = "The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship.  Change the data in the field or fields that contain duplicate data, remove the index, or redefine the index to permit duplicate entries and try again."
  Const Grid_UpdateFailedMessage As String = "Unable to complete update of Grid row. The changes you have made will not be saved without further input from you. Choose either abort edit, in which case your changes to this row will be lost, or choose Ok, in which case you can complete this row"
  Dim PKErr As Boolean, PKDupErr As Boolean
  Dim fLen As Long
  Dim errstring As String, field As String, table As String
  Dim pkcolumns As String
  Dim p0 As Long, p1 As Long, q As Long
    
  If m_InDisplayError Then Exit Sub
  m_InDisplayError = True
  Err.Source = ErrorSource(Err, SourceFunction)
  If Not m_ForcePopError Then Call ErrorMessagePush(Err)
  On Error Resume Next
  If Not m_SkipError Then
    If m_ForcePopError Then
      RetryCancelCaption = "&Abort Edit"
      If ErrorMessagePop(ERR_ERROR + ERR_ALLOWCANCEL, Err, "Grid_Error", "Unable to complete update of row", Grid_UpdateFailedMessage) Then
        Call AbortGridEdit
      End If
      RetryCancelCaption = ""
    Else
      If Not m_AutoGrid_RDO Is Nothing Then
        Do
          m_RDOError = False
          DoEvents
        Loop Until Not m_RDOError
      End If
      errstring = grid.ErrorText
      If Len(m_RDOErrorDescription) > 0 Then errstring = errstring & vbCrLf & m_RDOErrorDescription
      PKErr = (StrComp(Left$(errstring, Len(PK_ValidationError)), PK_ValidationError, vbTextCompare) = 0)
      PKDupErr = (StrComp(Left$(errstring, Len(PKdup_ValidationError)), PKdup_ValidationError, vbTextCompare) = 0)
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
        Call ErrorMessagePop(ERR_ERROR, Err, "Grid_Error", "Unable to complete update of row", "An error has occurred updating the current row." & vbCr & vbCr & errstring)
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
          Call ErrorMessagePop(ERR_ERROR, Err, "Grid_Error", "Validation failure", "Unable to validate the current row." & vbCr & errstring)
        Else
          If GetValidationString(errstring) Then
            Call ErrorMessagePop(ERR_ERROR, Err, "Grid_Error", "Validation failure", "Unable to validate the current row." & vbCr & errstring)
          Else
            Call ErrorMessagePop(ERR_ERROR, Err, "Grid_Error", "Grid Error", "An error has occurred in the grid." & vbCr & errstring)
          End If
        End If
      End If
    End If
    m_RDOErrorDescription = ""
  End If
  m_SkipError = False
  m_ForcePopError = False
  m_InDisplayError = False
End Sub

Private Sub Grid_Error(ByVal DataError As Integer, Response As Integer)
  Response = False
  Call DisplayGridError("Grid_Error")
End Sub

Private Sub Grid_HeadClick(ByVal ColIndex As Integer)
  ParentAC.SaveColumns = True
End Sub

Private Sub DisplayFastKey(ByVal FastKey As String)
  If Len(FastKey) > 0 Then FastKey = "Fast Key: " & FastKey
  If Not m_AutoGrid_DAO Is Nothing Then
    m_AutoGrid_DAO.FastKeyLabel = FastKey
  End If
  If Not m_AutoGrid_RDO Is Nothing Then
    m_AutoGrid_RDO.FastKeyLabel = FastKey
  End If
End Sub

Private Function IsValidKeyCode(ByVal KeyCode As Integer, ByVal Shift As Integer) As Boolean
  If (KeyCode = 0) Or ((Shift And vbAltMask) = vbAltMask) Or ((Shift And vbCtrlMask) = vbCtrlMask) Then Exit Function
  If (KeyCode = vbKeyReturn) Or (KeyCode = vbKeyTab) Or (KeyCode = vbKeyUp) Or (KeyCode = vbKeyDown) Or (KeyCode = vbKeyLeft) Or (KeyCode = vbKeyRight) Then Exit Function
  IsValidKeyCode = True
End Function
Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
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
    If KeyCode <> vbKeyEscape Then KeyCode = 0
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
  If (Shift = 0) And (KeyCode = vbKeyF4) And ParentAC.EnableF4Copy Then
    If UpdateGrid Then Call CopyPasteCurrentGridRow(grid, ParentAC.AutoColumns)
    KeyCode = 0
  End If
  If (Shift = 0) And (KeyCode = vbKeyF5) And ParentAC.EnableF5Copy Then
    If UpdateGrid Then Call CopyPasteGridValue(grid, ParentAC.AutoColumns)
    KeyCode = 0
  End If
  If (Shift = 0) And (KeyCode = vbKeyF9) And ParentAC.EnableF9Refresh Then
    Call RefreshGrid(False, True, False)
    KeyCode = 0
  End If
  If IsValidKeyCode(KeyCode, Shift) Then Call CheckDropDown(Nothing)
  inKeyDown = False
End Sub

Private Sub Grid_ComboSelect(ByVal ColIndex As Integer)
  Dim aCol As AutoCol
  
  'Debug.Print "CS"
  If GetAColByGridIndex(ColIndex, aCol) Then
    aCol.FastKey = ""
    aCol.FastKeyPress = ""
    Call DisplayFastKey(aCol.FastKey)
  End If
End Sub

Private Sub GridSpeedKey(KeyAscii As Integer, Optional ByVal bGridKeyPressDelay As Boolean = False)
  Dim SetSelText As Boolean, uKeyAscii As Integer, vValue As Variant
  Dim vItem As ValueItem, sValue As String, fLen As Long, s As String
  Dim aCol As AutoCol, ColSet As TrueDBGrid60.Column, NewValue As Integer
  Dim cCtrl As TrueDBGrid60.TDBDropDown, dCtrl As Data, rs As Recordset, vbmk As Variant
    
  On Error GoTo GridSpeedKey_err
  If GetAColByGridIndex(grid.Col, aCol) Then
    If KeyAscii = vbKeyReturn Then
      If IsComboOpen(grid.hwnd) And (Len(m_SaveComboValue) > 0) Then Call grid.PostMsg(MSG_COMBO_RESELECT)
      Exit Sub
    End If
    If aCol.IsCheckBox Then
      If (grid.AddNewMode = dbgAddNewCurrent) And (Not m_AutoGrid_RDO Is Nothing) Then
        If m_DelayedKeyPressEvent Then Call grid.PostMsg(MSG_KEYPRESS_DELAY)
        Exit Sub
      End If
      If (KeyAscii = vbKeySpace) Or (KeyAscii = 0) Then
        Set ColSet = grid.Columns(grid.Col)
        If ColSet.Locked Or (Not grid.AllowUpdate) Then Exit Sub
        NewValue = Not (ColSet.Value = -1)
        ColSet.Text = ""
        ColSet.Value = NewValue
      End If
      KeyAscii = 0
      Exit Sub
    End If
    If ((aCol.DropType And DROPDOWN_LIST) = DROPDOWN_LIST) Or _
       ((aCol.DropType And DROPDOWN_COMBO) = DROPDOWN_COMBO) Then
      m_SaveComboValue = ""
      Set ColSet = grid.Columns(grid.Col)
      If ColSet.Locked Or (Not grid.AllowUpdate) Then Exit Sub
      If Not bGridKeyPressDelay Then
        If (aCol.DropType And DROPDOWN_COMBO) = DROPDOWN_COMBO Then
          aCol.FastKey = Left$(grid.Text, grid.SelStart)
        End If
        If KeyAscii = vbKeyBack Then
          If Len(aCol.FastKeyPress) > 0 Then aCol.FastKeyPress = Left$(aCol.FastKeyPress, Len(aCol.FastKeyPress) - 1)
          If (aCol.DropType And DROPDOWN_COMBO) = DROPDOWN_COMBO Then Exit Sub
          If Len(aCol.FastKey) > 0 Then aCol.FastKey = Left$(aCol.FastKey, Len(aCol.FastKey) - 1)
        ElseIf IsPrint(KeyAscii) Then
          aCol.FastKey = aCol.FastKey & Chr$(ToUpper(KeyAscii))
          aCol.FastKeyPress = aCol.FastKeyPress & Chr$(KeyAscii)
        ElseIf KeyAscii = vbKeyEscape Then
          aCol.FastKey = ""
          aCol.FastKeyPress = ""
        End If
        Debug.Print "BEGIN " & aCol.FastKey & ":" & aCol.FastKeyPress; ":" & grid.Text
        If (aCol.DropType And DROPDOWN_LIST) = DROPDOWN_LIST Then
          Call DisplayFastKey(aCol.FastKey)
          KeyAscii = 0
        End If
      End If
      If (grid.AddNewMode = dbgAddNewCurrent) And (Not m_AutoGrid_RDO Is Nothing) Then
        If m_DelayedKeyPressEvent Then Call grid.PostMsg(MSG_KEYPRESS_DELAY)
        Exit Sub
      End If
      fLen = Len(aCol.FastKey)
      If fLen = 0 Then Exit Sub
      If (aCol.QueryCombo > 0) And (ColSet.ValueItems.Count = 0) Then
        If Not m_AutoGrid_DAO Is Nothing Then
          Set cCtrl = m_AutoGrid_DAO.Combo(aCol.QueryCombo)
          Set dCtrl = m_AutoGrid_DAO.DataCombo(aCol.QueryCombo)
          If dCtrl.Recordset Is Nothing Then Exit Sub
          vbmk = cCtrl.Bookmark
          Set rs = dCtrl.Recordset
          rs.FindFirst rs.Fields(0).Name & " like " & StrSQL(aCol.FastKey & "*")
          If Not rs.NoMatch Then
            cCtrl.FirstRow = rs.Bookmark
            ColSet.Text = cCtrl.Text
          End If
        End If
      Else
        For Each vItem In ColSet.ValueItems
          If ((aCol.DropType And DROPDOWN_LIST) = DROPDOWN_LIST) And ColSet.ValueItems.Translate Then
            vValue = vItem.DisplayValue
          Else
            vValue = vItem.Value
          End If
          s = Left$(vValue, fLen)
          If StrComp(s, aCol.FastKey, vbTextCompare) = 0 Then
            sValue = vItem.Value
            If Not m_AutoGrid_RDO Is Nothing Then sValue = Trim$(sValue)
            If IsComboOpen(grid.hwnd) And ((aCol.DropType And DROPDOWN_LIST) = DROPDOWN_LIST) Then
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
      If (aCol.DropType And DROPDOWN_LIST) = DROPDOWN_LIST Then
        If IsComboOpen(grid.hwnd) Then
          Debug.Print "Combo Open: " & ColSet.Text & " : " & grid.Text
        End If
      ElseIf (aCol.DropType And DROPDOWN_COMBO) = DROPDOWN_COMBO Then
        If Not SetSelText Then
          If IsPrint(KeyAscii) Then
            ColSet.Text = Left$(aCol.FastKeyPress, Len(aCol.FastKeyPress) - 1)
          Else
            ColSet.Text = aCol.FastKeyPress
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
          Call DisplayFastKey(aCol.FastKey)
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
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
  Call GridSpeedKey(KeyAscii)
End Sub

Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim aCol As AutoCol, ColCheckBox As Boolean
  Dim MouseCol As Long, CurRow As Long, HeaderHeight As Single
  Dim UMenus As Collection, um As UserPopupMenu, NextMenu As Long
  Dim GridState As POPUP_MENU_SHOWFLAGS, ActiveMenu As Boolean
  On Error GoTo Grid_MouseDown_err
  
  If CheckMouseButton(VK_RBUTTON) Then
    MouseCol = grid.ColContaining(X)
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
      Call GetAColByGridIndex(grid.Col, aCol)
      If Not aCol Is Nothing Then ColCheckBox = aCol.IsCheckBox
      grid.Row = CurRow
      frmPopupMenu.mnuFilterSelectionEx.visible = True
      frmPopupMenu.mnuFilterSelectionInc.visible = True
      frmPopupMenu.mnuFilterSelectionEx.Enabled = True ' (Len(grid.SelText) > 0) Or ColCheckBox
      frmPopupMenu.mnuFilterSelectionInc.Enabled = True ' (Len(grid.SelText) > 0) Or ColCheckBox
      frmPopupMenu.mnuFilterS.visible = True
      frmPopupMenu.mnuFilterS.Enabled = Not (m_AutoGrid_DAO Is Nothing)
        
      frmPopupMenu.mnuCopyCol.visible = False
      frmPopupMenu.mnuInsCol.visible = False
      frmPopupMenu.mnuAutoSize.visible = False
      frmPopupMenu.mnuAutoSizeAll.visible = False
          
      Call frmFilter.FillRCMenu(Me)
      frmPopupMenu.sep3.visible = False
      
      frmPopupMenu.mnuFind.visible = True
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
      frmPopupMenu.mnuCopyRow.visible = (MouseCol <> -1)
      frmPopupMenu.mnuCutRow.visible = False
      frmPopupMenu.mnuPasteRow.visible = False
      frmPopupMenu.sep4.visible = False
      If (grid.VisibleRows = 0) Or (MouseCol = -1) Then
        frmPopupMenu.mnuFilterS.visible = False
        frmPopupMenu.mnuSortAsc.Enabled = False
        frmPopupMenu.mnuSortDesc.Enabled = False
        frmPopupMenu.mnuFind.visible = False
        frmPopupMenu.sep5.visible = False
      Else
        frmPopupMenu.mnuFilterS.visible = True
        frmPopupMenu.mnuFilterS.Enabled = Not (m_AutoGrid_DAO Is Nothing)
        frmPopupMenu.mnuFind.visible = True
        frmPopupMenu.mnuFind.Enabled = (Not (m_AutoGrid_DAO Is Nothing)) And (Not grid.Columns(MouseCol).ValueItems.Translate) Or True 'apf
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
  Dim aCol As AutoCol, nGridCols As Long, bRebaseDone(1 To 3) As Boolean
  Dim i As Long, j As Long, dTemp As Long
  
  On Error GoTo MoveColumns_Err
  Call xSet("MoveColumns")
  If InsertAtColumn > StartCol Then InsertAtColumn = InsertAtColumn - 1
  If (InsertAtColumn >= StartCol) And (InsertAtColumn <= EndCol) Then Err.Raise ERR_MOVECOLUMNS, "MoveColumns", "Cannot move into the selected column range"
  ' rebase column references InsertAtColumn, StartCol, EndCol to columns in m_DisplayOrder
   nGridCols = 0
   For i = 1 To ParentAC.Count
    Set aCol = ParentAC.Item(i)
    If Not aCol.Hide Then
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
        Call CopyPasteCurrentGridRow(grid, ParentAC.AutoColumns)
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

Private Function IsFieldUpdateable(ByVal aCol As AutoCol) As Boolean
  Dim rsRDO As RDOResultset
  On Error Resume Next

  IsFieldUpdateable = False
  If Not m_AutoGrid_DAO Is Nothing Then
    IsFieldUpdateable = IsFieldUpdateableDAO(DAODataGrid.Recordset, aCol.DataField, aCol.SourceField)
  End If
  If Not m_AutoGrid_RDO Is Nothing Then
    Set rsRDO = RDODataGrid.Resultset
    IsFieldUpdateable = rsRDO.rdoColumns(aCol.DataField).Updatable
  End If
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
  If ((Len(ParentAC.Filter) > 0) Or (Len(ParentAC.Sort) > 0)) And Not (m_AutoGrid_RDO Is Nothing) Then
    grid.AllowAddNew = False
    grid.AllowUpdate = False
    grid.AllowDelete = False
  Else
    grid.AllowAddNew = m_GridAllowAddNew
    grid.AllowUpdate = m_GridAllowUpdate
    grid.AllowDelete = m_GridAllowDelete
    If ForceDDRefresh Then Call RefreshDropDowns
  End If
End Sub

Public Sub ReDisplayColumns()
  Call xSet("ReDisplayColumns")
  Call ReBindColumns
  Call xReturn("ReDisplayColumns")
End Sub

Private Sub CreateGridSplit(ByVal FullRefresh As Boolean, ByVal InResize As Boolean)
  Dim lSplit As TrueDBGrid60.Split, rSplit As TrueDBGrid60.Split, SplitCol As Long
  Dim aCol As AutoCol, ColSet As TrueDBGrid60.Column, i As Long
    
  On Error GoTo CreateGridSplit_err
  Set lSplit = grid.Splits.Item(0)
  If FullRefresh Then
    Do While grid.Splits.Count > 1
      Call grid.Splits.Remove(0)
    Loop
    For i = 1 To ParentAC.Count
      Set aCol = ParentAC.Item(i)
      If aCol.GridColumn >= 0 Then
        If aCol.SplitColumn Then
          Set rSplit = grid.Splits.Add(1)
          SplitCol = aCol.GridColumn + 1
          grid.Split = (grid.Splits.Count - 1)
          Exit For
        End If
      End If
    Next i
  Else
    If grid.Splits.Count > 1 Then
      For i = 1 To ParentAC.Count
        Set aCol = ParentAC.Item(i)
        If aCol.GridColumn >= 0 Then
          If aCol.SplitColumn Then
            Set rSplit = grid.Splits.Item(1)
            SplitCol = aCol.GridColumn + 1
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
    lSplit.AllowSizing = False
    For i = 0 To (lSplit.Columns.Count - 1)
      Set ColSet = lSplit.Columns(i)
      If InResize Then ColSet.Width = grid.Columns.Item(ColSet.ColIndex).Width
      ColSet.visible = (i < SplitCol)
    Next i
    For i = 0 To (rSplit.Columns.Count - 1)
      Set ColSet = rSplit.Columns(i)
      If InResize Then ColSet.Width = grid.Columns.Item(ColSet.ColIndex).Width
      ColSet.visible = (i >= SplitCol)
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
  Dim aCol As AutoCol, ColSet As TrueDBGrid60.Column

  On Error GoTo ReBindColumns_err
  m_GridActive = False
  Call ClearGridColumns
  grid.CellTips = dbgNoCellTips
  
  'Get Display order
  For i = 1 To ParentAC.Count
    Set aCol = ParentAC.Item(m_DisplayOrder(i))
    aCol.GridColumn = -1
    If aCol.ForceCreateColumn Or (Not aCol.Hide) Then
      id = grid.Columns.Count
      Set ColSet = grid.Columns.Add(id)
      aCol.GridColumn = id
      CharCount = CountChar(aCol.GridCaption, vbCr, vbBinaryCompare) + 1
      If CharCount > grid.HeadLines Then grid.HeadLines = CharCount
      ColSet.Caption = aCol.GridCaption
      If Not aCol.UnboundColumn Then ColSet.DataField = aCol.DataField
      ColSet.Alignment = aCol.Alignment
      ColSet.Locked = (Not IsFieldUpdateable(aCol)) Or aCol.NoEdit
      ColSet.visible = True
      If Len(aCol.ToolTip) > 0 Then
        grid.CellTipsDelay = 500
        grid.CellTips = dbgFloating
      End If
      If aCol.Hide Then aCol.ColWidth = 0
      If aCol.ColWidth > 0 Then
        ColSet.Width = aCol.ColWidth
      ElseIf aCol.ColWidth = 0 Then
        ColSet.visible = False
      End If
      If aCol.CustomFormat Then
        ColSet.NumberFormat = "FormatText Event"
      Else
        ColSet.NumberFormat = aCol.Format
      End If
      ColSet.Button = aCol.HasButton
      If aCol.HasButton And aCol.IsCheckBox Then Err.Raise ERR_GRIDFORMAT, "RebindColumns", "Cannot have the CheckBox and Button format on the same column" & vbCrLf & "column [" & aCol.GridCaptionClean & "]" & vbCrLf
      If aCol.HasButton Then ColSet.ButtonPicture = LoadResPicture(aCol.ButtonPictureIndex, vbResBitmap)
      
      'FIX CAD1
      ColSet.ForeColor = aCol.ForeColor
      ColSet.BackColor = aCol.BackColor
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

Private Sub RefreshDropDowns()
  Dim i As Long
  Dim aCol As AutoCol
 
  On Error GoTo RefreshDropDowns_err
  For i = 1 To ParentAC.Count
    Set aCol = ParentAC.Item(i)
    If aCol.IsCheckBox Or (((aCol.DropType And DROPDOWN_LIST) = DROPDOWN_LIST) And (aCol.TranslateDropDown <> 0)) Then
      Call RefreshSingleDropDown(aCol, False)
    Else
      aCol.RefreshRequired = True
    End If
  Next i
  Set aCol = Nothing
  Exit Sub
  
RefreshDropDowns_err:
  Err.Raise Err.Number, ErrorSource(Err, "RefreshDropDowns"), Err.Description
  Resume
End Sub

Private Function RefreshSingleDropDown(ByVal aCol As AutoCol, ByVal ShowButtonOnly As Boolean) As Boolean
  Dim db As Database, cn As rdoConnection
  Dim rs As Recordset
  Dim rsRDO As RDOResultset
  Dim i As Long, j As Long, dItems As Long, cCount As Long
  Dim ColSet As TrueDBGrid60.Column
  Dim vItem As ValueItem, doResize As Boolean
  Dim errstring As String
 
  On Error GoTo RefreshSingleDropDown_err
  If m_InGridRefresh Then Exit Function
  doResize = False
  aCol.RefreshRequired = False
  If (aCol.GridColumn >= 0) And ((aCol.DropType <> DROPDOWN_NONE) Or aCol.IsCheckBox) Then
    Set ColSet = grid.Columns(aCol.GridColumn)
    If aCol.IsCheckBox Then
      Call ColSet.ValueItems.Clear
      Set vItem = New ValueItem
      vItem.DisplayValue = LoadResPicture(TICK_BMP, vbResBitmap)
      vItem.Value = -1
      Call ColSet.ValueItems.Add(vItem)
      
      Set vItem = New ValueItem
      If ParentAC.CheckBoxCross Then
        vItem.DisplayValue = LoadResPicture(CROSS_BMP, vbResBitmap)
      Else
        vItem.DisplayValue = LoadResPicture(CROSS_BLANK_BMP, vbResBitmap)
      End If
      vItem.Value = 0
      Call ColSet.ValueItems.Add(vItem)
          
      ColSet.ValueItems.Translate = True
      ColSet.ValueItems.CycleOnClick = True
      ColSet.ValueItems.Presentation = dbgNormal
      ColSet.Alignment = dbgCenter
      aCol.TranslateDropDown = -1
    Else
      If Not m_AutoGrid_DAO Is Nothing Then Set db = ParentAC.WorkSpace.Databases(ParentAC.DBDAOPath)
      If Not m_AutoGrid_RDO Is Nothing Then Set cn = RDODataGrid.Resultset.ActiveConnection
            
      If ((aCol.DropType And DROPDOWN_LIST) = DROPDOWN_LIST) Or _
         ((aCol.DropType And DROPDOWN_COMBO) = DROPDOWN_COMBO) Then
        doResize = True
        Call ColSet.ValueItems.Clear
        If ShowButtonOnly Then
          Set vItem = New ValueItem
          Call ColSet.ValueItems.Add(vItem)
          ColSet.ValueItems.Presentation = dbgComboBox
          aCol.RefreshRequired = True
        Else
          If (aCol.DropType And DROPDOWN_QUERY) = 0 Then  ' values only
            j = 1
            Do While j <= aCol.NoParams
              Set vItem = New ValueItem
              If aCol.DropType = DROPDOWN_COMBO Then
                vItem.Value = aCol.GetParam(j)
                j = j + 1
              Else
                vItem.DisplayValue = aCol.GetParam(j)
                vItem.Value = aCol.GetParam(j + 1)
                j = j + 2
              End If
              Call ColSet.ValueItems.Add(vItem)
            Loop
            cCount = 2  ' translate if appropriate
          Else   ' query to non bound drop down
            If Not cn Is Nothing Then
              Set rsRDO = GetComboRecordsetRDO(cn, aCol)
              If rsRDO Is Nothing Then GoTo RefreshSingleDropDown_end
              cCount = rsRDO.rdoColumns.Count
              If Not rsRDO Is Nothing Then
                Do While Not rsRDO.EOF
                  Set vItem = New ValueItem
                  If (aCol.DropType And DROPDOWN_COMBO) = DROPDOWN_COMBO Then
                    vItem.Value = rsRDO.rdoColumns(0).Value
                  Else
                    If cCount > 1 Then
                      vItem.DisplayValue = rsRDO.rdoColumns(0).Value
                      vItem.Value = rsRDO.rdoColumns(1).Value
                    Else
                      vItem.Value = rsRDO.rdoColumns(0).Value
                    End If
                  End If
                  Call ColSet.ValueItems.Add(vItem)
                  rsRDO.MoveNext
                Loop
              End If
            End If
            Set rsRDO = Nothing
            If Not db Is Nothing Then
              Set rs = GetComboRecordsetDAO(db, aCol)
              If rs Is Nothing Then GoTo RefreshSingleDropDown_end
              cCount = rs.Fields.Count
              If Not rs Is Nothing Then
                Do While Not rs.EOF
                  Set vItem = New ValueItem
                  If (aCol.DropType And DROPDOWN_COMBO) = DROPDOWN_COMBO Then
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
              End If
            End If
            Set rs = Nothing
          End If
          'ColSet.ValueItems.Validate = True
          If ((aCol.DropType And DROPDOWN_LIST) = DROPDOWN_LIST) And (cCount > 1) Then
            ColSet.ValueItems.Translate = True
            aCol.TranslateDropDown = -1
          Else
            ColSet.ValueItems.Translate = False
            aCol.TranslateDropDown = 0
          End If
          ColSet.ValueItems.Presentation = dbgComboBox
          
          If aCol.MaxDropDownItems > 0 Then
            dItems = aCol.MaxDropDownItems
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
        Call DoQueryCombos(db, aCol, ColSet)
      End If
    End If
  End If
  If doResize Then Call AutoSizeSingleDropDown(aCol)
  RefreshSingleDropDown = True
  
RefreshSingleDropDown_end:
  Set aCol = Nothing
  Exit Function
  
RefreshSingleDropDown_err:
  If Not aCol Is Nothing Then
    If aCol.IsCheckBox Then
      errstring = "Error displaying the CheckBox for column [" & aCol.GridCaptionClean & "]" & vbCrLf
    Else
      errstring = "Error refreshing the drop down for column [" & aCol.GridCaptionClean & "]" & vbCrLf
      If (aCol.DropType And DROPDOWN_QUERY) = 0 Then
        errstring = errstring & "Static values based drop down" & vbCrLf
      Else
        errstring = errstring & "Query based drop down: [" & aCol.QueryName & "]" & vbCrLf
      End If
      errstring = errstring & "Drop down type: " & GetDropDownTypeAsString(aCol.DropType) & vbCrLf
    End If
  End If
  Err.Raise Err.Number, ErrorSource(Err, "RefreshSingleDropDown"), errstring & Err.Description
  Resume
End Function

'Private Sub AutoSizeDropDowns()
'  Dim aCol As AutoCol
'  Dim i As Long
'
'  On Error GoTo AutoSizeDropDowns_err
'  For i = 1 To ParentAC.Count
'    Set aCol = ParentAC.Item(i)
'    Call AutoSizeSingleDropDown(aCol)
'  Next i
'  Exit Sub
'
'AutoSizeDropDowns_err:
'  Err.Raise Err.Number, ErrorSource(Err, "AutoSizeDropDowns"), Err.Description
'End Sub

Private Sub AutoSizeSingleDropDown(ByVal aCol As AutoCol)
  Dim ColSet As TrueDBGrid60.Column
  Dim cCtrl As TrueDBGrid60.TDBDropDown
  Dim tWidth As Single, FirstColumn As Boolean
  Dim j As Long
  
  On Error GoTo AutoSizeSingleDropDown_err
  If aCol.GridColumn >= 0 Then
    Set ColSet = grid.Columns(aCol.GridColumn)
    If Len(ColSet.DropDown) > 0 Then
      Set cCtrl = m_AutoGrid_DAO.Combo(aCol.QueryCombo)
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

Private Sub DoQueryCombos(ByVal db As Database, aCol As AutoCol, ColSet As TrueDBGrid60.Column)
  Dim FirstField As Long, VisibleFields As Long
  Dim daoCtrl As Data, rdoCtrl As MSRDC.MSRDC
  Dim cCtrl As TrueDBGrid60.TDBDropDown, rs As Recordset
  Dim dItems As Long, j As Long
    
  ' apf calc NoItems Properly
  On Error GoTo DoQueryCombos_err
  If aCol.MaxDropDownItems > 0 Then
    dItems = aCol.MaxDropDownItems
  Else
    dItems = ParentAC.MaxDropDownItems
  End If
  If Not m_AutoGrid_DAO Is Nothing Then
    If db Is Nothing Then Set db = ParentAC.WorkSpace.Databases(ParentAC.DBDAOPath)
    Set daoCtrl = m_AutoGrid_DAO.DataCombo(aCol.QueryCombo)
    Set cCtrl = m_AutoGrid_DAO.Combo(aCol.QueryCombo)
    cCtrl.ColumnHeaders = False
    daoCtrl.DatabaseName = ParentAC.DBDAOPath
    Set rs = GetComboRecordsetDAO(db, aCol)
    Set daoCtrl.Recordset = rs
    cCtrl.DataField = aCol.BoundDataField
    cCtrl.Font = grid.Font
    If Not rs Is Nothing Then
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
    Set db = Nothing
  End If
  If Not m_AutoGrid_RDO Is Nothing Then
    Set rdoCtrl = m_AutoGrid_RDO.DataCombo(aCol.QueryCombo)
    Set cCtrl = m_AutoGrid_RDO.Combo(aCol.QueryCombo)
    If Len(RDODataGrid.Connect) > 0 Then
      rdoCtrl.Connect = RDODataGrid.Connect
      rdoCtrl.sql = GetComboSQL(aCol)
      rdoCtrl.Refresh
      cCtrl.DataField = aCol.BoundDataField
      cCtrl.Font = grid.Font
      If Not rdoCtrl.Resultset Is Nothing Then
        If rdoCtrl.Resultset.rdoColumns.Count > 1 Then
          cCtrl.Columns(0).visible = False
          FirstField = 1
          VisibleFields = (rs.Fields.Count - 1)
        Else
          FirstField = 0
          VisibleFields = 1
        End If
        cCtrl.ListField = rdoCtrl.Resultset.rdoColumns(FirstField).Name
        cCtrl.Width = ColSet.Width * VisibleFields
        If dItems > rdoCtrl.Resultset.RowCount Then
          dItems = rdoCtrl.Resultset.RowCount
          If dItems <= 0 Then dItems = 1
        End If
        cCtrl.Height = (grid.RowHeight * dItems) + 25
        ColSet.DropDown = cCtrl.Name
      End If
    End If
  End If
  Exit Sub
  
DoQueryCombos_err:
  Err.Raise Err.Number, ErrorSource(Err, "DoQueryCombos"), Err.Description
End Sub
          
Public Function GetDropRecordset(ByVal DataField As String) As Recordset
  Dim aCol As AutoCol, dCtrl As Data, vbmk As Variant
  Dim tdbgDrop As TrueDBGrid60.TDBDropDown
  
  Set aCol = GetAColByKey(DataField)
  If Not aCol Is Nothing Then
    Set tdbgDrop = m_AutoGrid_DAO.Combo(aCol.QueryCombo)
    Set dCtrl = m_AutoGrid_DAO.DataCombo(aCol.QueryCombo)
    vbmk = tdbgDrop.Bookmark
    Set GetDropRecordset = dCtrl.Recordset
    If Not IsEmpty(vbmk) Then GetDropRecordset.Bookmark = vbmk
  End If
End Function

Public Function GetDropResultset(ByVal DataField As String) As RDOResultset
  Dim aCol As AutoCol, dCtrl As MSRDC.MSRDC, vbmk As Variant
  Dim tdbgDrop As TrueDBGrid60.TDBDropDown
  
  Set aCol = GetAColByKey(DataField)
  If Not aCol Is Nothing Then
    Set tdbgDrop = m_AutoGrid_RDO.Combo(aCol.QueryCombo)
    Set dCtrl = m_AutoGrid_RDO.DataCombo(aCol.QueryCombo)
    vbmk = tdbgDrop.Bookmark
    Set GetDropResultset = dCtrl.Resultset
    If Not IsEmpty(vbmk) Then GetDropResultset.Bookmark = vbmk
  End If
End Function


Private Function GetComboRecordsetRDO(cn As rdoConnection, aCol As AutoCol) As RDOResultset
  Dim rQry As rdoQuery, rsRDOGrid As RDOResultset
  Dim DynamicParams As Boolean
  Dim i As Long, p As Long
  Dim ColIndex As Integer
  Dim rsEmpty As Boolean, rsDoQuery As Boolean
  Dim param As String, paramValue As Variant
  Dim sql As String
  
  On Error GoTo GetComboRecordsetRDO_err
  ' apf should be dynamic ReadOnly Cursor
  Set rsRDOGrid = RDODataGrid.Resultset
  rsEmpty = rsRDOGrid.EOF And rsRDOGrid.BOF
  rsDoQuery = True

  If aCol.NoParams > 0 And Not aCol.IsSQLString Then
    
    Set rQry = cn.CreateQuery("Col_Query" & aCol.DataField, aCol.QueryName)
    
    
    For i = 1 To aCol.NoParams
      param = aCol.GetParam(i)
      p = InStr(1, param, "<")
      If p = 1 Then
        paramValue = Mid$(param, 2, Len(param) - 2)
      Else
        ColIndex = GetGridColIndex(param)
        If ColIndex <> -1 Then
          paramValue = grid.Columns(ColIndex)
          DynamicParams = True
        Else
          rsDoQuery = Not rsEmpty
          If Not rsEmpty Then paramValue = rsRDOGrid.rdoColumns(param)
        End If
      End If
      
      rQry.rdoParameters(i - 1) = paramValue
      
    Next i
    If rsDoQuery Then Set GetComboRecordsetRDO = rQry.OpenResultset(rdOpenDynamic, rdConcurReadOnly)
  Else
    sql = aCol.QueryName
    
    For i = 1 To aCol.NoParams
      param = aCol.GetParam(i)
      p = InStr(1, param, "<")
      If p = 1 Then
        paramValue = Mid$(param, 2, Len(param) - 2)
      Else
        ColIndex = GetGridColIndex(param)
        If ColIndex <> -1 Then
          paramValue = grid.Columns(ColIndex)
          DynamicParams = True
        Else
          rsDoQuery = Not rsEmpty
          If Not rsEmpty Then paramValue = rsRDOGrid.rdoColumns(param)
        End If
      End If
      sql = ReplaceString(sql, BEGIN_PARAM & param & END_PARAM, GetSQLValueAsString(paramValue, aCol.dbDataType), vbTextCompare)
    Next i
    If rsDoQuery Then Set GetComboRecordsetRDO = cn.OpenResultset(sql, rdOpenDynamic, rdConcurReadOnly)
  End If
  If DynamicParams Then
     aCol.DropType = aCol.DropType Or DROPDOWN_DYNAMIC
   Else
     aCol.DropType = aCol.DropType Or DROPDOWN_STATIC
  End If
  
GetComboRecordsetRDO_end:
  Set rQry = Nothing
  Exit Function
GetComboRecordsetRDO_err:
  Set GetComboRecordsetRDO = Nothing
  Call ErrorMessage(ERR_ERROR, Err, "GetComboRecordsetRDO", "Initialising drop down", "Unable to initailise drop down on column '" & aCol.GridCaptionClean & "'." & vbCrLf & "The resultset to support the drop down could not be created.")
  Resume GetComboRecordsetRDO_end
  Resume
End Function

Private Function GetComboRecordsetDAO(db As Database, aCol As AutoCol) As Recordset
  Dim qry As QueryDef, DynamicParams As Boolean, rsGrid As Recordset
  Dim i As Long, p As Long
  Dim ColIndex As Integer
  Dim rsEmpty As Boolean, rsDoQuery As Boolean
  Dim param As String, paramValue As Variant
  Dim sql As String
  
  On Error GoTo GetComboRecordsetDAO_err

  Set rsGrid = DAODataGrid.Recordset
  rsEmpty = rsGrid.EOF And rsGrid.BOF
  rsDoQuery = True
  
  If Not aCol.IsSQLString Then
    If QueryPresent(db.QueryDefs, aCol.QueryName) Then
      Set rsGrid = DAODataGrid.Recordset
      Set qry = db.QueryDefs(aCol.QueryName)
        
      For i = 1 To aCol.NoParams
        param = aCol.GetParam(i)
        p = InStr(1, param, "<")
        If p = 1 Then
          paramValue = Mid$(param, 2, Len(param) - 2)
        Else
          ColIndex = GetGridColIndex(param)
          If ColIndex <> -1 Then
            paramValue = grid.Columns(ColIndex)
            DynamicParams = True
          Else
            rsDoQuery = Not rsEmpty
            If Not rsEmpty Then paramValue = rsGrid.Fields(param)
          End If
        End If
        qry.Parameters(i - 1) = paramValue
      Next i
      
      If rsDoQuery Then Set GetComboRecordsetDAO = qry.OpenRecordset(dbOpenSnapshot, dbFailOnError)
    Else
      If rsDoQuery Then Set GetComboRecordsetDAO = db.OpenRecordset(aCol.QueryName, dbOpenSnapshot, dbFailOnError)
    End If
  Else
    'we have a standard piece of sql to execute
    sql = aCol.QueryName
    
    For i = 1 To aCol.NoParams
     param = aCol.GetParam(i)
     p = InStr(1, param, "<")
     If p = 1 Then
       paramValue = Mid$(param, 2, Len(param) - 2)
     Else
       ColIndex = GetGridColIndex(param)
       If ColIndex <> -1 Then
         paramValue = grid.Columns(ColIndex)
         DynamicParams = True
       Else
         rsDoQuery = Not rsEmpty
         If Not rsEmpty Then paramValue = rsGrid.Fields(param)
       End If
     End If
      sql = ReplaceString(sql, BEGIN_PARAM & param & END_PARAM, GetSQLValueAsString(paramValue, aCol.dbDataType), vbTextCompare)
    Next i
    If rsDoQuery Then Set GetComboRecordsetDAO = db.OpenRecordset(sql, dbOpenSnapshot, dbFailOnError)
  End If
  
  If DynamicParams Then
    aCol.DropType = aCol.DropType Or DROPDOWN_DYNAMIC
  Else
    aCol.DropType = aCol.DropType Or DROPDOWN_STATIC
  End If
  
GetComboRecordsetDAO_end:
  Set qry = Nothing
  Exit Function
  
GetComboRecordsetDAO_err:
  Set GetComboRecordsetDAO = Nothing
  Call ErrorMessage(ERR_ERROR, Err, "GetComboRecordsetDAO", "Initialising drop down", "Unable to initailise drop down on column '" & aCol.GridCaptionClean & "'." & vbCrLf & "The recordset to support the drop down could not be created.")
  Resume GetComboRecordsetDAO_end
  Resume
End Function

' apf should be dynamic ReadOnly Cursor
Private Function GetComboSQL(ByVal aCol As AutoCol) As String
  GetComboSQL = aCol.QueryName
End Function

Public Sub Refresh(Optional ByVal ForceRefresh As Boolean = False)
  Call xSet("Refresh")
  Call RefreshGrid(True, ForceRefresh, ForceRefresh)
  Call xReturn("Refresh")
End Sub

Friend Sub RefreshGrid(ByVal SetFocus As Boolean, ByVal ForceDataRefresh As Boolean, ByVal ForceDataRebind As Boolean)
  Dim OldNotifyDisable As Boolean
  Dim RowOffset As Long, ColOffset As Long
    
  On Error GoTo RefreshGrid_err
  If m_InGridRefresh Then Exit Sub
  OldNotifyDisable = mNotifyDisable
  mNotifyDisable = True
  m_InGridRefresh = True
  If Not grid Is Nothing Then
    If grid.visible Then
      Call UpdateGrid
      If ForceDataRefresh Then
        Call GetCurrentRowColOffset(ColOffset, RowOffset)
        ParentAC.RequeryOnRefresh = True
        ParentAC.SortFilterRefresh = True
        Call ParentAC.SetFilterSort
      End If
      If Not RDODataGrid Is Nothing Then
        RDODataGrid.Refresh
        If RDODataGrid.Resultset.BOF And Not RDODataGrid.Resultset.EOF Then RDODataGrid.Resultset.MoveNext
      End If
      If grid.VisibleRows > 1 Then
        grid.Row = 1
        grid.Row = 0
      End If
      If grid.VisibleCols > 0 Then grid.Col = 0
      grid.Refresh
      grid.CurrentCellVisible = True
      If SetFocus Then Call TryGridSetFocus
      If ForceDataRefresh Then
        Call SetCurrentPosition(ColOffset, RowOffset, grid.Bookmark)
        If ForceDataRebind Then Call ReBindColumns(False)
      End If
    End If
  End If
  
RefreshGrid_end:
  m_InGridRefresh = False
  Call RefreshDropDowns
  mNotifyDisable = OldNotifyDisable
  Exit Sub
  
RefreshGrid_err:
  Call DisplayGridError("RefreshGrid")
  Resume RefreshGrid_end
  Resume
End Sub

Private Sub GetCurrentRowColOffset(ColOffset As Long, RowOffset As Long)
  ColOffset = 0
  RowOffset = 0
  If Not grid.BOF Then
    ColOffset = grid.Col
    If Not RDODataGrid Is Nothing Then
      RowOffset = (RDODataGrid.Resultset.AbsolutePosition) * -1
    ElseIf Not DAODataGrid Is Nothing Then
      RowOffset = (DAODataGrid.Recordset.AbsolutePosition + 1) * -1
    End If
    If RowOffset < 0 Then
      Do
        RowOffset = RowOffset + 1
      Loop Until (Not IsNull(grid.GetBookmark(RowOffset))) Or (RowOffset >= 0)
    End If
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
  If Not DAODataGrid Is Nothing Then DAODataGrid.FontName = NewValue
  If Not RDODataGrid Is Nothing Then RDODataGrid.Font.Name = NewValue
End Property

Property Let FontSize(ByVal NewValue As Long)
  On Error Resume Next
  grid.Font.Size = NewValue
  grid.HeadFont.Size = NewValue
  grid.RowHeight = 0
  If Not DAODataGrid Is Nothing Then DAODataGrid.FontSize = NewValue
  If Not RDODataGrid Is Nothing Then RDODataGrid.Font.Size = NewValue
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
        If DatabaseTarget = DB_TARGET_JET Then
          sSort = sSort & "[" & DataField & "] " & IIf(Asc, "ASC", "DESC")
        Else
          sSort = sSort & GEN_TABLE_NAME & "." & DataField & " " & IIf(Asc, "ASC", "DESC")
        End If
        If i <> grid.SelEndCol Then sSort = sSort & ","
      End If
    Next i
    ParentAC.Sort = sSort
  Else
    DataField = ColKey(grid.Col)
    If Len(DataField) > 0 Then
      If DatabaseTarget = DB_TARGET_JET Then
        ParentAC.Sort = "[" & DataField & "] " & IIf(Asc, "ASC", "DESC")
      Else
        ParentAC.Sort = GEN_TABLE_NAME & "." & DataField & " " & IIf(Asc, "ASC", "DESC")
      End If
    End If
  End If
End Sub

Public Sub FindInCol()
  Dim aCol As AutoCol
  If GetAColByGridIndex(grid.Col, aCol) Then Call Find(aCol.DataField)
End Sub

Private Function GetAColByGridIndex(ByVal GridColumn As Integer, aCol As AutoCol) As Boolean
  GetAColByGridIndex = GetAColByGridIndexEx(GridColumn, aCol, ParentAC)
End Function

Friend Function GetAColByKey(ByVal ColKey As Variant) As AutoCol
  Dim aCol As AutoCol
  
  On Error Resume Next
  Set aCol = ParentAC.Item(ColKey)
  Set GetAColByKey = aCol
End Function

Public Function SetColumnValue(ByVal ColKey As Variant, ByVal ValueItem As COLUMN_VALUES, ByVal NewValue As Variant) As Boolean
  Dim aCol As AutoCol
  
  Set aCol = GetAColByKey(ColKey)
  If Not aCol Is Nothing Then
    Select Case ValueItem
      Case COLUMN_HIDE
        aCol.Hide = CBool(NewValue)
      Case Else
        Call ECASE("Unknown Set Value item: " & CStr(ValueItem))
    End Select
  End If
End Function

Public Function GetColumnValue(ByVal ColKey As Variant, ByVal ValueItem As COLUMN_VALUES) As Variant
  Dim aCol As AutoCol
  
  GetColumnValue = Empty
  Set aCol = GetAColByKey(ColKey)
  If Not aCol Is Nothing Then
    Select Case ValueItem
      Case COLUMN_CAPTION
        GetColumnValue = aCol.GridCaption
      Case COLUMN_CAPTION
        GetColumnValue = aCol.GridCaption
      Case COLUMN_CAPTION_CLEAN
        GetColumnValue = aCol.GridCaptionClean
      Case COLUMN_DATAFIELD
        GetColumnValue = aCol.DataField
      Case COLUMN_HIDE
        GetColumnValue = aCol.Hide
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
  Dim aCol As AutoCol, ColName As String, ColIndex As Integer
  Dim FilterStr As String
  
  On Error GoTo FilterBySelection_err
  ColIndex = grid.Col
  If Not GetAColByGridIndex(ColIndex, aCol) Then Err.Raise ERR_NOFILTER, "FilterBySelection", "Invalid Column selected"
  ColName = aCol.GridCaptionClean
  If Not IncludeSelection Then FilterStr = "NOT "
  If (Len(grid.SelText) <> Len(grid.Columns(ColIndex).Text)) And (aCol.dbDataType = TYPE_STR) Then
    FilterStr = FilterStr & WhereClause(aCol.DataField, grid.SelText, TYPE_STR, True)
  Else
    FilterStr = FilterStr & WhereClause(aCol.DataField, grid.Columns(ColIndex).Value, aCol.dbDataType, False, grid.Columns(ColIndex).Text)
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
    Call SetClipboardColumn(ParentAC.AutoColumns, True)
    Call CopyClipBoardRow
    m_CopyAllColumnValues = True
    Call SetClipboardColumn(ParentAC.AutoColumns, False)
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


Private Sub PasteSelection()
  Dim cString  As String, ColSet As TrueDBGrid60.Column
  Dim aCol As AutoCol, p As Long
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
  Dim ColSet As TrueDBGrid60.Column, i As Long, k As Long, cSplit As Integer
  Dim aCol As AutoCol
  Dim InLastRow As Boolean
  Dim rs As Recordset, rsRDO As RDOResultset, FieldName As String
  Dim Headers() As String, FieldValues() As String, sRow As String, sValue As String
  Dim icol As Long, iRow As Long, MaxCol As Long, maxrow As Long
  Dim p0 As Long, p1 As Long, p As Long, pend As Long, q0 As Long, q1 As Long
  Dim v As Variant, dType As DATABASE_FIELD_TYPES, vbmk As Variant
  Dim rCount As Long
  
  On Error GoTo SetRow_err
  Call SetCursor
  grid.Enabled = False
  If grid.AddNewMode = dbgAddNewCurrent Then
    InLastRow = True
    If (grid.Row > 0) And (grid.VisibleRows > 0) Then
      grid.Row = grid.Row - 1
      InLastRow = False
    End If
  End If
  
  If PrivateFormat Then
    m_PastePrivateFormat = True
    If Not m_AutoGrid_DAO Is Nothing Then Set rs = DAODataGrid.Recordset
    If Not m_AutoGrid_RDO Is Nothing Then
      Set rsRDO = RDODataGrid.Resultset
      Call grid.Close(False)
      If Not (rsRDO.BOF And rsRDO.EOF) Then vbmk = rsRDO.Bookmark
    End If
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
      
      For iRow = 1 To maxrow
        If Not rs Is Nothing Then rs.AddNew
        If Not rsRDO Is Nothing Then rsRDO.AddNew
        For icol = 1 To MaxCol
          FieldName = Headers(icol)
          v = FieldValues(((iRow - 1) * MaxCol) + icol)
          Call UpdateFieldValue(rsRDO, rs, FieldName, v)
        Next icol
        If Not rs Is Nothing Then rs.Update
        If Not rsRDO Is Nothing Then rsRDO.Update
      Next iRow
    End If
  Else ' standard format from excel etc..
    ' vbmk = Grid.RowBookmark(cRow)
    If (Not m_AutoGrid_DAO Is Nothing) And Not InLastRow Then DAODataGrid.Recordset.MoveLast
    If (Not m_AutoGrid_RDO Is Nothing) And Not InLastRow Then RDODataGrid.Resultset.MoveLast
    If Not InLastRow Then grid.Row = grid.Row + 1
  
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
              
              If IsPasteField(ParentAC.AutoColumns, ColSet.DataField) Then grid.Columns(i).Value = sValue
              If q1 = 0 Then GoTo NEXT_ROW
              q0 = q1 + 1
            End If
          Next i
        Next k
NEXT_ROW:
        If (p1 <> 0) Or (rCount <> 1) Then
          Call UpdateGrid
          grid.Row = grid.Row + 1
        End If
      End If
      p0 = p1 + 2
    Loop Until p1 = 0
  End If
  
SetRow_end:
  m_PastePrivateFormat = False
  If Not rsRDO Is Nothing Then
    If Not IsEmpty(vbmk) Then rsRDO.Bookmark = vbmk
    Call grid.ReOpen
    Call RefreshGrid(False, True, False)
  End If
  grid.Enabled = True
  Call ClearCursor
  Exit Sub
  
SetRow_err:
  If (Err.Number = 3164) Or (Err.Number = 3421) Then Resume Next
  Call ErrorMessagePush(Err)
  If Not rs Is Nothing Then
    If rs.EditMode <> dbEditNone Then rs.CancelUpdate
  End If
  If Not rsRDO Is Nothing Then
    If rsRDO.EditMode <> rdEditNone Then rsRDO.CancelUpdate
  End If
  Call AbortGridEdit
  Call ErrorMessagePop(ERR_ERROR, Err, "SetRow", "Paste into row", "Row could not be pasted")
  Resume SetRow_end
  Resume
End Sub
Private Function GetColumnsFromArray(vColumns As Variant) As Long
  Dim i As Long
  Dim j As Long
  Dim aCol As AutoCol
  
  On Error GoTo GetColumnsFromArray_ERR
  
  GetColumnsFromArray = -1
  If Not IsArrayEx(vColumns) Then GoTo GetColumnsFromArray_END
  
  For i = LBound(vColumns) To UBound(vColumns)
    j = j + 1
    Select Case VarType(vColumns(i))
      Case vbString
        Set aCol = GetAColByKey(vColumns(i))
        aCol.ClipboardColumn = True
      Case vbLong, vbInteger
        Call GetAColByGridIndex(grid.Col, aCol)
        aCol.ClipboardColumn = True
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

Public Sub SelectAll()
  If Not m_AutoGrid_DAO Is Nothing Then Call m_AutoGrid_DAO.SelectAll
  If Not m_AutoGrid_RDO Is Nothing Then Call m_AutoGrid_RDO.SelectAll
End Sub

Public Sub PasteColumnValue(ParamArray Columns())
  Dim i As Integer
  Dim vColumns As Variant
  Dim aCol As AutoCol
  
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

Public Function CopyColumnValue(ParamArray Columns()) As Boolean
  Dim i As Integer
  Dim vColumns As Variant
  Dim aCol As AutoCol
  
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
  Call CopyClipBoardRow
  CopyColumnValue = True
  
CopyColumnValue_END:
  Call ClearCursor
  Exit Function
  
CopyColumnValue_ERR:
  Call ErrorMessage(ERR_ERROR, Err, ErrorSource(Err, "CopyColumnValue"), "Copy Column Value", "Error in CopyColumnValue")
  Resume CopyColumnValue_END
  Resume
End Function

Private Sub CopyClipBoardRow()
  Dim rs As Recordset, rsRDO As RDOResultset
  Dim fbmk As Variant, ibmk As Variant

  On Error GoTo CopyClipBoardRow_err
  Call SetCursor
  If Not m_AutoGrid_DAO Is Nothing Then Set rs = DAODataGrid.Recordset.Clone
  If Not m_AutoGrid_RDO Is Nothing Then
    Set rsRDO = RDODataGrid.Resultset
    fbmk = grid.FirstRow
    ibmk = grid.Bookmark
  End If
  Call CopyClipboardRowEx(rs, rsRDO, grid, ParentAC.AutoColumns)
  
CopyClipBoardRow_end:
  On Error Resume Next
  If Len(fbmk) > 0 Then
    grid.FirstRow = fbmk
    grid.Bookmark = ibmk
  End If
  Call ClearCursor
  Exit Sub
  
CopyClipBoardRow_err:
  Resume CopyClipBoardRow_end
End Sub

Private Function CopyClipBoardColumn()
  Dim rs As Recordset, rsRDO As RDOResultset
  Dim bmk As Variant
  Dim i As Integer, FieldKey As String, s As String

  On Error Resume Next
  Call SetCursor
  Call LockWindowUpdate(Me.GridControl.ContainerForm.hwnd)
  If Not m_AutoGrid_DAO Is Nothing Then
    Set rs = DAODataGrid.Recordset
    If rs.EOF And rs.BOF Then Exit Function
    bmk = rs.Bookmark
    rs.MoveFirst
  End If
  If Not m_AutoGrid_RDO Is Nothing Then
    Set rsRDO = RDODataGrid.Resultset
    If rsRDO.EOF And rsRDO.BOF Then Exit Function
    bmk = rsRDO.Bookmark
    rsRDO.MoveFirst
  End If
  Do
    For i = grid.SelStartCol To grid.SelEndCol
      FieldKey = ColKey(i)
      If Len(FieldKey) > 0 Then
        If Not rs Is Nothing Then s = s & rs.Fields(FieldKey) & vbTab
        If Not rsRDO Is Nothing Then s = s & rsRDO.rdoColumns(FieldKey) & vbTab
      End If
    Next i
    s = s & vbCrLf
    If Not rs Is Nothing Then
      rs.MoveNext
      If rs.EOF Then Exit Do
    End If
    If Not rsRDO Is Nothing Then
      rsRDO.MoveNext
      If rsRDO.EOF Then Exit Do
    End If
  Loop
  If Not rs Is Nothing Then rs.Bookmark = bmk
  If Not rsRDO Is Nothing Then rsRDO.Bookmark = bmk
  Call SetAnyClipboardData(s)
  Call LockWindowUpdate(0)
  Call ClearCursor
End Function

Public Sub TotalColumnDisplay()
  Dim v As Variant, aCol As AutoCol
  Dim icol As Integer
    
  On Error GoTo TotalColumnDisplay_err
  icol = grid.Col
  
  If GetAColByGridIndex(icol, aCol) Then
    If IsNumberField(aCol.dbDataType) Then
      v = TotalColumn(icol)
      Call DisplayMessage(Nothing, "Sum = " & FormatNumber(v, 2, vbUseDefault, vbUseDefault, vbTrue), ParentAC.AutoName & " - " & aCol.GridCaptionClean, "Ok", "")
    End If
  End If
  Exit Sub
  
TotalColumnDisplay_err:
  Call ErrorMessage(ERR_ERROR, Err, "TotalColumnDisplay", "Column Total", "Error calculation column total")
End Sub

Public Function TotalColumn(ByVal ColIndex As Integer) As Variant
  Dim vbmkRow As Variant, vbmkFirstRow As Variant, Col As Integer
  Dim rs As Recordset, rsRDO As RDOResultset
  Dim FieldName As String, aCol As AutoCol
  
  On Error GoTo TotalColumn_err
  TotalColumn = Empty
  If ColIndex < 0 Then ColIndex = grid.Col
  If GetAColByGridIndex(ColIndex, aCol) Then
    If IsNumberField(aCol.dbDataType) Then
      FieldName = aCol.DataField
      If Not m_AutoGrid_RDO Is Nothing Then
        Set rsRDO = RDOResultset
        vbmkFirstRow = grid.FirstRow
        vbmkRow = grid.Bookmark
        Col = grid.Col
        Call grid.Close(False)
        If Not (rsRDO.EOF And rsRDO.BOF) Then
          rsRDO.MoveFirst
          Do
            TotalColumn = TotalColumn + IsNullEx(rsRDO.rdoColumns(FieldName), 0)
            rsRDO.MoveNext
          Loop Until rsRDO.EOF
        End If
      Else
        Set rs = DAORecordset.Clone
        If Not (rs.BOF And rs.EOF) Then
          rs.MoveFirst
          Do
            TotalColumn = TotalColumn + IsNullEx(rs.Fields(FieldName), 0)
            rs.MoveNext
          Loop Until rs.EOF
        End If
      End If
    End If
  End If
  If Not IsEmpty(vbmkRow) Then
    Call grid.ReOpen(vbmkFirstRow)
    grid.Bookmark = vbmkRow
    grid.Col = Col
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
  Dim i As Long, vRecs() As Variant
  Dim bMax As Long
  
  On Error GoTo DeleteSelectedRows_err
  Call SetCursor
  mDeleteMultiCancel = True
  If grid.SelBookmarks.Count > 0 Then
    If grid.AllowDelete And IsDataUpdateable() Then
      Call DoDBEvents(True)
      bMax = grid.SelBookmarks.Count - 1
      ReDim vRecs(0 To bMax)
      For i = 0 To bMax
        vRecs(i) = grid.SelBookmarks(i)
      Next i
      For i = 0 To bMax
        mDeleteMultipleRows = i
        grid.Bookmark = vRecs(i)
        grid.Delete
      Next i
    End If
  Else
    If grid.AllowUpdate Then
      If grid.SelLength = 0 Then grid.SelLength = 1
      If Len(grid.SelText) > 0 Then grid.SelText = ""
    End If
  End If
  
DeleteSelectedRows_end:
  If Not mDeleteMultiCancel Then grid.Refresh
  mDeleteMultipleRows = 0
  Call ClearCursor
  Exit Sub

DeleteSelectedRows_err:
  If Not m_SkipError Then
    Call ErrorMessage(ERR_ERROR, Err, "DeleteSelectedRows", "Deleting rows selected", "The program has encountered an error while deleting the selected rows.")
    m_SkipError = False
  End If
  Resume DeleteSelectedRows_end
  Resume
End Sub

Public Sub SaveWidths()
  Dim i As Integer
  Dim FieldName As String
  
  On Error GoTo SaveWidths_err
  For i = 0 To grid.Columns.Count - 1
    FieldName = ColKey(i)
    If Len(FieldName) Then
      If grid.Columns(i).visible Then
        ParentAC.Item(FieldName).ColWidth = grid.Columns(i).Width
      Else
        ParentAC.Item(FieldName).ColWidth = 0
      End If
    End If
  Next i
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

Public Sub SaveRecord()
  Call ECASE("Not implemented")
End Sub

Public Sub ShowFilterWizard()
  If Not m_AutoGrid_RDO Is Nothing Then
    Call ECASE("not implemented")
    Exit Sub
  End If
  Load frmFilter
  Set frmFilter.QBEFilter.rsFilter = DAODataGrid.Recordset
    
ShowFilter:
  frmFilter.Show vbModal
  If frmFilter.ExitMode > CANCEL_BUTTON Then
    ParentAC.SortFilterRefresh = True
    Call ParentAC.SetFilterSort(frmFilter.QBEFilter.sSort, frmFilter.QBEFilter.sFilter)
    If frmFilter.ExitMode = APPLY_BUTTON Then GoTo ShowFilter
  End If
  Set frmFilter.QBEFilter.rsFilter = Nothing
  Unload frmFilter
End Sub

Private Function IsDataUpdateable() As Boolean
  IsDataUpdateable = True
  If Not m_AutoGrid_DAO Is Nothing Then
    If Not DAODataGrid.Recordset Is Nothing Then
      IsDataUpdateable = DAODataGrid.Recordset.Updatable
    End If
  End If
  If Not m_AutoGrid_RDO Is Nothing Then
    If Not RDODataGrid.Resultset Is Nothing Then
      IsDataUpdateable = RDODataGrid.Resultset.Updatable
    End If
  End If
End Function

Private Sub DoDBEvents(Optional ByVal FreeAndRefresh As Boolean = False)
  Dim Flag As Long
  
  Flag = dbFreeLocks
  If FreeAndRefresh Then Flag = Flag + dbRefreshCache
  If Not m_AutoGrid_DAO Is Nothing Then
    Call DBEngine.Idle(Flag)
  End If
  If Not m_AutoGrid_RDO Is Nothing Then
    
  End If
End Sub

Private Sub EditControl(ByVal EnableEdit As Boolean)
  Dim i As Integer, aCol As AutoCol
  Dim ColSetTo As TrueDBGrid60.Column

  'Debug.Print "Edit Control:" & EnableEdit
  For i = 1 To ParentAC.Count
    Set aCol = ParentAC.Item(i)
    If aCol.GridColumn >= 0 Then
      Set ColSetTo = grid.Columns(aCol.GridColumn)
      If EnableEdit Then
        ColSetTo.Locked = aCol.NoAddNew Or (Not IsFieldUpdateable(aCol))
      Else
        ColSetTo.Locked = aCol.NoEdit
      End If
    End If
  Next i
End Sub

Private Sub Grid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  Dim aCol As AutoCol, icol As Integer
  
  On Error GoTo RowColChange_err
  m_CopyAllColumnValues = False
  If m_GridActive Then
    'Debug.Print "RowColChange Event"
    icol = grid.Col
    If GetAColByGridIndex(icol, aCol) Then
      If IsDataUpdateable() And aCol.RefreshRequired Then
        Call RefreshSingleDropDown(aCol, True)
      End If
      'Debug.Print "RCC"
      aCol.FastKey = ""
      aCol.FastKeyPress = ""
      Call DisplayFastKey(aCol.FastKey)
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
  Dim aCol As AutoCol, sql As String, frm As Form
  
  If Not m_AutoGrid_DAO Is Nothing Then Set frm = m_AutoGrid_DAO.ContainerForm
  If Not m_AutoGrid_RDO Is Nothing Then Set frm = m_AutoGrid_RDO.ContainerForm
  If (Tag >= 0) And GetAColByGridIndex(grid.Col, aCol) Then
    If Not DisplayMessage(frm, "Formats for field [" & aCol.DataField & "]" & vbCrLf & aCol.FormatString, "Debug formats [" & aCol.GridCaptionClean & "]", "Ok", "Copy to Clipboard") Then
      Call SetAnyClipboardData(aCol.FormatString)
    End If
  Else
    sql = ParentAC.sql
    If Not DisplayMessage(frm, vbCrLf & sql & vbCrLf, "Debug SQL for [" & ParentAC.AutoName & "]", "Ok", "Copy to Clipboard") Then
      Call SetAnyClipboardData(sql)
    End If
  End If
End Sub

Private Sub RecordAuditValues(rs As Recordset, rsRDO As RDOResultset, ByVal inDelete As Boolean)
  Dim rFld As field, iFld As FieldDetails, rCol As rdoColumn
  
  If mAuditIF.AuditLevel = GRIDAUDIT_NONE Then Exit Sub
  Set mAuditFields = New Collection
  If Not m_AutoGrid_DAO Is Nothing Then
    If inDelete Then
      mAuditType = GRID_DELETE
    ElseIf rs.EditMode = dbEditAdd Then
      mAuditType = GRID_ADDNEW
    Else
      mAuditType = GRID_EDIT
    End If
    If mAuditIF.AuditLevel = GRIDAUDIT_STATUSONLY Then Exit Sub
    For Each rFld In rs.Fields
      Set iFld = New FieldDetails
      iFld.Name = rFld.Name
      iFld.Value = Null
      If mAuditType = GRID_ADDNEW Then
        iFld.Tag = Null
      Else
        iFld.Tag = rFld.Value
      End If
      iFld.DataType = DAOtoDatatype(rs.Fields(iFld.Name).Type)
      iFld.Description = rs.Fields(iFld.Name).SourceTable
      If Len(iFld.Description) > 0 Then Call mAuditFields.Add(iFld, iFld.Name)
    Next rFld
  End If
  If Not m_AutoGrid_RDO Is Nothing Then
    If inDelete Then
      mAuditType = GRID_DELETE
    ElseIf rsRDO.EditMode = rdEditAdd Then
      mAuditType = GRID_ADDNEW
    Else
      mAuditType = GRID_EDIT
    End If
    If mAuditIF.AuditLevel = GRIDAUDIT_STATUSONLY Then Exit Sub
    For Each rCol In rsRDO.rdoColumns
      Set iFld = New FieldDetails
      iFld.Name = rCol.Name
      iFld.Value = Null
      If mAuditType = GRID_ADDNEW Then
        iFld.Tag = Null
      Else
        iFld.Tag = rCol.Value
      End If
      iFld.DataType = RDOtoDatatype(rsRDO.rdoColumns(iFld.Name).Type)
      iFld.Description = rsRDO.rdoColumns(iFld.Name).SourceTable
      Call mAuditFields.Add(iFld, iFld.Name)
    Next rCol
  End If
End Sub


