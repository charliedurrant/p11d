VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AutoClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements ILibraryVersion
Option Compare Text

Public Enum FIND_TYPES
  FT_NONE = 0
  FT_FINDFIRST
  FT_FINDNEXT
  FT_FINDPREVIOUS
  FT_FINDLAST
End Enum

Public Enum ENDAUTOLINE_OPTIONS
  AUTO_ENDREPORT = 0
  AUTO_NEWSECTION
  AUTO_PAGEBREAK
End Enum

Public Enum AUTO_SAVECOLS
  SAVE_WIDTH = 2
  SAVE_ORDER = 4
  SAVE_SORT = 8
  SAVE_FILTER = 16
  SAVE_CLEAR = 1024
End Enum

Private Enum AUTO_PRINT
  PRINT_USER = 0
  PRINT_RS
  PRINT_RDO
End Enum

Public Enum TEMPLATE_PRINT
  TEMPLATE_NONE
  TEMPLATE_PROFORMA_ONLY
  TEMPLATE_PROFORMA_AND_DATA
End Enum

Private m_AutoName  As String
Private m_AutoCountIndex As Long
Private m_AutoInitialised As Boolean
Private m_RSDefault As Recordset
Private m_rdoRSDefault As RDOResultset
Private m_DefaultSQL As String

Private m_DebugMode  As Boolean

Private m_AutoGrid_DAO As Object
Private m_DBDAOPath As String
Private m_AutoGrid_RDO As Object
Private m_Grid As AutoGrid
Private m_SavedGrid As AutoGrid
Private m_AutoCols As Collection
Private m_PersistAutoCols As Collection

Private m_CheckGridDate As Boolean
Private m_CheckDataType As Boolean

Private m_Filter As String
Private m_Sort As String
Private m_sfRefresh As Boolean
Private m_RequeryOnRefresh  As Boolean

Private m_SaveColumns As Boolean
Private m_BooleanCheck As Boolean
Private m_CheckBoxCross As Boolean

'Grid properties exposed through Auto
Private m_EnableF9Refresh As Boolean
Private m_EnableF4Copy As Boolean
Private m_EnableF5Copy As Boolean
Private m_UpdateOnLoseFocus As Boolean


Private m_DefaultMaxDropDownItems As Long

Private m_ValidateDataTypes As DATABASE_FIELD_TYPES

' Auto report constants
Private m_Report As Reporter
Private m_ExportReport As Boolean
Private m_FitToPage As Boolean
Private m_IncludeColHeaders As Boolean
Private m_TrimColHeaders As Boolean
Private m_AlignColHeaders As Boolean
Private m_PrintRecordCount As Boolean
Private m_PrintTotalTitle As Boolean
Private m_PrintTotalAsValue As Boolean
Private m_PrintTotalText As String
Private m_GroupHeaderSeparator As String
Private m_GroupHeaderFieldSuffix As String
Private m_PrintFieldSuffix As Boolean
Private m_TemplatePrint As TEMPLATE_PRINT
Private m_SpanMultiplePages  As Boolean
Private m_MaxSpanPage  As Long

Private m_PrintGroupHeader As Boolean
Private m_CollapseAllGroups As Boolean
Private m_CollapseThisGroup As Boolean
Private m_CollapseLastGroup As Boolean

Private m_GroupHeaderWidth  As Single

Private mLineSpacing As Long
Private m_ReportFormat As String
Private m_ReportNonZeroRowsOnly As Boolean
Private m_ReportHeader As String
Private m_ReportHeaderOnAllPages As Boolean
Private m_PageFooter As String
Private m_PrePageHeader As String
Private m_PostPageHeader As String
Private m_DoHeaders As Boolean
Private m_LastGroupHeader As String
Private m_MaxReportLevel As Long
Private m_DoTotals As Boolean
Private m_DoFinalTotal As Boolean
Private m_OffsetsCalculated As Boolean
Private m_FreeSpaceX  As Single

Private m_OnFirstPage As Boolean
Private m_PrintType As AUTO_PRINT
Private m_RowCount As Long
Private m_GroupRowCount As Long
Private m_LineCount As Long
Private m_PreviewRows As Long
Private m_PrevLine As String
Private m_ForceNewPage As Boolean

' Printing AutoCols
Private m_cols() As AutoCol
Private m_lmin As Long
Private m_lmax As Long
Private m_lfirstubcolumn As Long

Private m_ConfirmDeleteText As String

Private m_DefaultIntegerFormat As String
Private m_DefaultFloatFormat As String
Private m_DefaultDateFormat As String
Private m_DefaultBoolFormat As String
Private Const DEFAULT_FLOAT_FORMAT = "###0.00"
Private Const DEFAULT_INTEGER_FORMAT = "###0"
Private Const DEFAULT_DATE_FORMAT = "dd/mm/yyyy"
Private Const DEFAULT_BOOL_FORMAT = "True/False"

'Handle bespoke menu items
Private mUserPopupHandler As IAutoPopupHandler
Private mUserPopupMenus As Collection

'Handle Grid Audit interface
Private mGridAuditIF As iAutoAuditRecord
Private mGridNotifyIF As IBaseNotify
Private m_WorkSpace As WorkSpace
Private m_AlternateBoxedLines As Boolean

Private Sub Class_Initialize()
  AutoCount = AutoCount + 1
  m_AutoCountIndex = AutoCount
  m_AutoInitialised = True
  Set m_WorkSpace = gwsMain
  If AutoNames Is Nothing Then Set AutoNames = New StringList
  If AutoParser Is Nothing Then
    Set AutoParser = New Parser
    Call SetUpParser
    AutoClipHandle = RegisterPrivateClipboard("TCSAuto Clipboard")
    Call RegisterLibrary(Me)
  End If
  
  Set m_AutoCols = New Collection
  Set m_PersistAutoCols = New Collection
  m_ConfirmDeleteText = ""
  
  m_DefaultFloatFormat = DEFAULT_FLOAT_FORMAT
  m_DefaultIntegerFormat = DEFAULT_INTEGER_FORMAT
  m_DefaultDateFormat = DEFAULT_DATE_FORMAT
  m_DefaultBoolFormat = DEFAULT_BOOL_FORMAT
  m_ReportFormat = "{TIMES=10,N}"
  m_PageFooter = "{x=2}{Times=8}{DATE} {TIME}  " & GetStatic("ApplicationName") & " Version " & GetStatic("Version") & "{X=85}Page {PAGE}"
  m_PrePageHeader = ""
  m_PostPageHeader = ""
  m_PreviewRows = PREVIEWROWS
  m_ExportReport = False
  m_IncludeColHeaders = True
  m_PrintRecordCount = True
  m_BooleanCheck = False
  m_FreeSpaceX = 0
  m_PrintTotalTitle = True
  m_PrintGroupHeader = True
  m_CollapseAllGroups = False
  m_PrintTotalAsValue = False
  m_PrintTotalText = "Total"
  m_GroupHeaderFieldSuffix = " = "
  m_GroupHeaderSeparator = ";  "
  m_PrintFieldSuffix = True
  m_FitToPage = False
  m_ReportHeaderOnAllPages = False
  m_DoFinalTotal = True
  m_CheckGridDate = True
  m_CheckDataType = True
  m_CheckBoxCross = True
  m_ReportNonZeroRowsOnly = False
  m_EnableF9Refresh = True
  m_EnableF4Copy = True
  m_EnableF5Copy = True
  m_DefaultMaxDropDownItems = 5
  m_UpdateOnLoseFocus = False
  Set mUserPopupMenus = New Collection
  ReDim m_cols(0 To 0) As AutoCol
  m_sfRefresh = True
  m_TemplatePrint = TEMPLATE_NONE
  Call InitAutoLog
  Debug.Print "AutoClass init: " & m_AutoCountIndex
End Sub

Private Sub Class_Terminate()
  Call CloseAutoLog
  Debug.Print "AutoClass " & m_AutoName & " terminate: " & m_AutoCountIndex
End Sub

Public Property Get ExportReport() As Boolean
  ExportReport = m_ExportReport
End Property

Public Property Let ExportReport(ByVal NewValue As Boolean)
  m_ExportReport = NewValue
End Property

Public Property Get DebugMode() As Boolean
  DebugMode = m_DebugMode Or IsRunningInIDE()
End Property

Public Property Let DebugMode(ByVal NewValue As Boolean)
  m_DebugMode = NewValue
End Property

Public Property Get sql() As String
  sql = m_DefaultSQL
End Property

Public Property Let BooleanCheckBox(ByVal NewValue As Boolean)
  m_BooleanCheck = NewValue
End Property

Public Property Get BooleanCheckBox() As Boolean
  BooleanCheckBox = m_BooleanCheck
End Property

Public Property Let CheckBoxCross(ByVal NewValue As Boolean)
  m_CheckBoxCross = NewValue
End Property

Public Property Get CheckBoxCross() As Boolean
  CheckBoxCross = m_CheckBoxCross
End Property

Public Property Let CheckGridDate(ByVal NewValue As Boolean)
  m_CheckGridDate = NewValue
End Property

Public Property Get CheckGridDate() As Boolean
  CheckGridDate = m_CheckGridDate
End Property

Public Property Let CheckDataType(ByVal NewValue As Boolean)
  m_CheckDataType = NewValue
End Property

Public Property Get CheckDataType() As Boolean
  CheckDataType = m_CheckDataType
End Property

Friend Property Let SortFilterRefresh(ByVal NewValue As Boolean)
  m_sfRefresh = NewValue
End Property

Friend Property Let RequeryOnRefresh(ByVal NewValue As Boolean)
  m_RequeryOnRefresh = NewValue
End Property

Public Property Let ValidateDataTypes(ByVal NewValue As DATABASE_FIELD_TYPES)
  If NewValue = 0 Then
    m_ValidateDataTypes = 0
  Else
    If Not ((NewValue = TYPE_BOOL) Or (NewValue = TYPE_DATE) Or (NewValue = TYPE_DOUBLE) Or (NewValue = TYPE_LONG) Or (NewValue = TYPE_STR)) Then Err.Raise 380, "ValidateDataTypes", "ValidateDataTypes must be set to a valid datatype"
    m_ValidateDataTypes = m_ValidateDataTypes Or (2 ^ (NewValue + 1))
  End If
End Property

Friend Property Get ValidateDataTypes() As DATABASE_FIELD_TYPES
  ValidateDataTypes = m_ValidateDataTypes
End Property

Public Property Let ConfirmDeleteText(ByVal NewValue As String)
  m_ConfirmDeleteText = NewValue
End Property

Public Property Get ConfirmDeleteText() As String
  ConfirmDeleteText = m_ConfirmDeleteText
End Property

Public Property Let SaveColumns(ByVal NewValue As Boolean)
  m_SaveColumns = NewValue
End Property

Public Sub FilterBySelection(ByVal IncludeSelection As Boolean)
  If m_Grid Is Nothing Then Err.Raise 380, "FilterBySelection", "Unable to execute filter as there is no grid available"
  Call m_Grid.FilterBySelection(IncludeSelection)
End Sub

Public Sub SortColumn(ByVal Asc As Boolean)
  If m_Grid Is Nothing Then Err.Raise 380, "SortColumn", "Unable to Sort columns as there is no grid available"
  Call m_Grid.SortColumn(Asc)
End Sub

Public Property Get LineSpacing() As Long
  LineSpacing = mLineSpacing
End Property

Public Property Let LineSpacing(ByVal NewValue As Long)
  If (NewValue < 1) Or (NewValue > 10) Then Err.Raise 380, "LineSpacing", "Valid values are between 1 and 10"
  mLineSpacing = NewValue
End Property

Public Property Get WorkSpace() As WorkSpace
  Set WorkSpace = m_WorkSpace
End Property

Public Property Set WorkSpace(NewValue As WorkSpace)
  Set m_WorkSpace = NewValue
End Property

Public Property Get Filter() As String
  Filter = m_Filter
End Property

Public Property Let Filter(ByVal NewValue As String)
  m_sfRefresh = True
  Call SetFilterSort(, NewValue)
End Property

Public Property Get Sort() As String
  Sort = m_Sort
End Property

Public Property Let Sort(ByVal NewValue As String)
  m_sfRefresh = True
  Call SetFilterSort(NewValue)
End Property

Public Property Let Font(ByVal NewValue As String)
  If m_Grid Is Nothing Then Err.Raise 380, "Font", "Unable to set property value as there is no grid available"
  m_Grid.Font = NewValue
End Property

Public Property Let FontSize(ByVal NewValue As Long)
  If m_Grid Is Nothing Then Err.Raise 380, "FontSize", "Unable to set property value as there is no grid available"
  m_Grid.FontSize = NewValue
End Property

Public Property Get DefaultFloatFormat() As String
  DefaultFloatFormat = m_DefaultFloatFormat
End Property

Public Property Get DefaultIntegerFormat() As String
  DefaultIntegerFormat = m_DefaultIntegerFormat
End Property

Public Property Get DefaultDateFormat() As String
  DefaultDateFormat = m_DefaultDateFormat
End Property

Public Property Get DefaultBooleanFormat() As String
  DefaultBooleanFormat = m_DefaultBoolFormat
End Property

Public Property Let NumberFormat(ByVal NewValue As String)
  Dim p As Long
  m_DefaultFloatFormat = NewValue
  m_DefaultIntegerFormat = NewValue
  p = InStr(NewValue, ".")
  If p > 0 Then m_DefaultIntegerFormat = Left$(NewValue, p - 1)
End Property

Public Property Let DateFormat(ByVal NewValue As String)
  m_DefaultDateFormat = NewValue
End Property

Public Property Let BooleanFormat(ByVal NewValue As String)
  If Not ((StrComp(NewValue, "True/False", vbTextCompare) = 0) Or _
         (StrComp(NewValue, "Yes/No", vbTextCompare) = 0) Or _
         (StrComp(NewValue, "On/Off", vbTextCompare) = 0)) Then Err.Raise 380, "BooleanFormat", "Not a valid Boolean format"
  m_DefaultBoolFormat = NewValue
End Property

' Report Properties
Public Property Let ReportFormat(ByVal NewValue As String)
  If Len(NewValue) = 0 Then Err.Raise 380, "ReportFormat", "Cannot set the global report format to the empty string"
  m_ReportFormat = NewValue
End Property

Public Property Get ReportNonZeroRowsOnly() As Boolean
  ReportNonZeroRowsOnly = m_ReportNonZeroRowsOnly
End Property

Public Property Let ReportNonZeroRowsOnly(ByVal NewValue As Boolean)
  m_ReportNonZeroRowsOnly = NewValue
End Property

Public Property Get ReportFormat() As String
  ReportFormat = m_ReportFormat
End Property

Public Property Get ReportPreviewRows() As Long
  ReportPreviewRows = m_PreviewRows
End Property

Public Property Let ReportPreviewRows(ByVal NewValue As Long)
  m_PreviewRows = NewValue
End Property

Public Property Get PrintRecordCount() As Boolean
  PrintRecordCount = m_PrintRecordCount
End Property

Public Property Let PrintRecordCount(ByVal NewValue As Boolean)
  m_PrintRecordCount = NewValue
End Property

Public Property Get IncludeColHeaders() As Boolean
  IncludeColHeaders = m_IncludeColHeaders
End Property

Public Property Let IncludeColHeaders(ByVal NewValue As Boolean)
  m_IncludeColHeaders = NewValue
End Property
  
Public Property Get SpanMultiplePages() As Boolean
  SpanMultiplePages = m_SpanMultiplePages
End Property
  
Public Property Let SpanMultiplePages(ByVal NewValue As Boolean)
 m_SpanMultiplePages = NewValue
End Property
  
Public Property Get TemplatePrint() As TEMPLATE_PRINT
  TemplatePrint = m_TemplatePrint
End Property

Public Property Let TemplatePrint(ByVal NewValue As TEMPLATE_PRINT)
  m_TemplatePrint = NewValue
End Property

Public Property Get TrimColHeaders() As Boolean
  TrimColHeaders = m_TrimColHeaders
End Property

Public Property Let TrimColHeaders(ByVal NewValue As Boolean)
  m_TrimColHeaders = NewValue
End Property

Public Property Get AlignColHeaders() As Boolean
  AlignColHeaders = m_AlignColHeaders
End Property

Public Property Let AlignColHeaders(ByVal NewValue As Boolean)
  m_AlignColHeaders = NewValue
End Property

Public Property Get FitToPage() As Boolean
  FitToPage = m_FitToPage
End Property

Public Property Let FitToPage(ByVal NewValue As Boolean)
  m_FitToPage = NewValue
End Property

Public Property Get ReportHeader() As String
  ReportHeader = m_ReportHeader
End Property

Public Property Let ReportHeader(ByVal NewValue As String)
  m_ReportHeader = NewValue
End Property

Public Property Get ReportHeaderOnAllPages() As Boolean
  ReportHeaderOnAllPages = m_ReportHeaderOnAllPages
End Property

Public Property Let ReportHeaderOnAllPages(ByVal NewValue As Boolean)
  m_ReportHeaderOnAllPages = NewValue
End Property

Public Property Get GroupHeaderFieldSuffix() As String
  GroupHeaderFieldSuffix = m_GroupHeaderFieldSuffix
End Property

Public Property Let GroupHeaderFieldSuffix(ByVal NewValue As String)
  m_GroupHeaderFieldSuffix = NewValue
End Property

Public Property Get PrintFieldSuffix() As Boolean
  PrintFieldSuffix = m_PrintFieldSuffix
End Property

Public Property Let PrintFieldSuffix(ByVal NewValue As Boolean)
  m_PrintFieldSuffix = NewValue
End Property

Public Property Get CollapseAllGroups() As Boolean
  CollapseAllGroups = m_CollapseAllGroups
End Property

Public Property Let CollapseAllGroups(ByVal NewValue As Boolean)
  m_CollapseAllGroups = NewValue
End Property

Public Function CollapseGroup() As Boolean
  CollapseGroup = m_CollapseThisGroup Or m_CollapseAllGroups
End Function

Public Property Get GroupHeaderSeparator() As String
  PrintTotalText = m_GroupHeaderSeparator
End Property

Public Property Let GroupHeaderSeparator(ByVal NewValue As String)
  m_GroupHeaderSeparator = NewValue
End Property

Public Property Get PrintTotalText() As String
  PrintTotalText = m_PrintTotalText
End Property

Public Property Let PrintTotalText(ByVal NewValue As String)
  m_PrintTotalText = NewValue
End Property

Public Property Get PrintTotalTitleAsValue() As Boolean
  PrintTotalTitleAsValue = m_PrintTotalAsValue
End Property

Public Property Let PrintTotalTitleAsValue(ByVal NewValue As Boolean)
  m_PrintTotalAsValue = NewValue
End Property

Public Property Get PrintTotalTitle() As Boolean
  PrintTotalTitle = m_PrintTotalTitle
End Property

Public Property Let PrintTotalTitle(ByVal NewValue As Boolean)
  m_PrintTotalTitle = NewValue
End Property

Public Property Get PrintGroupHeader() As Boolean
  PrintGroupHeader = m_PrintGroupHeader
End Property

Public Property Let PrintGroupHeader(ByVal NewValue As Boolean)
  m_PrintGroupHeader = NewValue
End Property

Public Property Get PrintFinalTotal() As Boolean
  PrintFinalTotal = m_DoFinalTotal
End Property

Public Property Let PrintFinalTotal(ByVal NewValue As Boolean)
  m_DoFinalTotal = NewValue
End Property

Public Property Get PostPageHeader() As String
  PostPageHeader = m_PostPageHeader
End Property

Public Property Let PostPageHeader(ByVal NewValue As String)
  m_PostPageHeader = NewValue
End Property

Public Property Get PageHeader() As String
  PageHeader = m_PrePageHeader
End Property

Public Property Let PageHeader(ByVal NewValue As String)
  m_PrePageHeader = NewValue
End Property

Public Property Get EnableF9Refresh() As Boolean
  EnableF9Refresh = m_EnableF9Refresh
End Property

Public Property Let EnableF9Refresh(ByVal NewValue As Boolean)
  m_EnableF9Refresh = NewValue
End Property

Public Property Get EnableF4Copy() As Boolean
  EnableF4Copy = m_EnableF4Copy
End Property

Public Property Let EnableF4Copy(ByVal NewValue As Boolean)
  m_EnableF4Copy = NewValue
End Property

Public Property Get EnableF5Copy() As Boolean
  EnableF5Copy = m_EnableF5Copy
End Property

Public Property Let EnableF5Copy(ByVal NewValue As Boolean)
  m_EnableF5Copy = NewValue
End Property

Public Property Let MaxDropDownItems(ByVal NewValue As Long)
  m_DefaultMaxDropDownItems = NewValue
End Property

Public Property Get MaxDropDownItems() As Long
  MaxDropDownItems = m_DefaultMaxDropDownItems
End Property

Public Property Get UpdateOnLoseFocus() As Boolean
  UpdateOnLoseFocus = m_UpdateOnLoseFocus
End Property

Public Property Let UpdateOnLoseFocus(ByVal NewValue As Boolean)
  m_UpdateOnLoseFocus = NewValue
End Property

Public Property Get PageFooter() As String
  PageFooter = m_PageFooter
End Property

Public Property Let PageFooter(ByVal NewValue As String)
  m_PageFooter = NewValue
End Property

Public Sub ClearPreview()
  Call ClearPreviewLines(m_AutoCols)
End Sub

Public Property Get grid() As AutoGrid
  Set grid = m_Grid
End Property

Public Property Set UserPopupMenuHandler(NewValue As IAutoPopupHandler)
  Set mUserPopupHandler = NewValue
End Property

Public Property Get UserPopupMenuHandler() As IAutoPopupHandler
  Set UserPopupMenuHandler = mUserPopupHandler
End Property

Public Property Get AuditInterface() As iAutoAuditRecord
  Set AuditInterface = mGridAuditIF
End Property

Public Property Set AuditInterface(NewValue As iAutoAuditRecord)
  Set mGridAuditIF = NewValue
  If Not m_Grid Is Nothing Then Set m_Grid.AuditInterface = NewValue
End Property

Public Property Get NotifyInterface() As IBaseNotify
  Set NotifyInterface = mGridNotifyIF
End Property

Public Property Set NotifyInterface(NewValue As IBaseNotify)
  Set mGridNotifyIF = NewValue
  If Not m_Grid Is Nothing Then Set m_Grid.NotifyInterface = NewValue
End Property

Public Sub ClearColumns()
  If Not m_AutoCols Is Nothing Then
    Do While m_AutoCols.Count > 0
      Call m_AutoCols.Remove(1)
    Loop
  End If
End Sub

' Properties Used in AutoGrid - need to be available internally
Friend Property Get AutoName() As String
  AutoName = m_AutoName
End Property

Friend Property Get UserPopupMenus() As Collection
  Set UserPopupMenus = mUserPopupMenus
End Property

Friend Property Get DBDAOPath() As String
  DBDAOPath = m_DBDAOPath
End Property

Public Property Get Count() As Long
  Count = m_AutoCols.Count
End Property

Public Property Get Item(ByVal Index As Variant) As AutoCol
  Set Item = m_AutoCols.Item(Index)
End Property

Friend Property Get AutoColumns() As Collection
  Set AutoColumns = m_AutoCols
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = m_AutoCols.[_NewEnum]
End Function

Public Function InitAutoData(ByVal AutoName As String, rs As Recordset, Optional ByVal DatabaseName As String = "", Optional gridctrl As Object = Nothing) As Boolean
  InitAutoData = InitAutoDataEx(AutoName, rs, Nothing, DatabaseName, "", False, gridctrl)
End Function

Public Function InitAutoDataSQL(ByVal AutoName As String, rs As Recordset, sql As String, Optional ByVal DatabaseName As String = "", Optional gridctrl As Object = Nothing) As Boolean
  InitAutoDataSQL = InitAutoDataEx(AutoName, rs, Nothing, DatabaseName, sql, False, gridctrl)
End Function

Public Function InitAutoDataRDO(ByVal AutoName As String, rsRDO As RDOResultset, Optional gridctrl As Object = Nothing) As Boolean
  InitAutoDataRDO = InitAutoDataEx(AutoName, Nothing, rsRDO, "", "", False, gridctrl)
End Function

Private Function InitAutoDataEx(ByVal AutoName As String, rs As Recordset, rsRDO As RDOResultset, ByVal DatabasePath As String, ByVal sql As String, ByVal ReInit As Boolean, gridctrl As Object) As Boolean
  Dim colcurrent As AutoCol
    
  Call xSet("InitAutoData")
  Call SetCursor
  On Error GoTo InitAutoDataEx_err
  InitAutoDataEx = True
  Call Me.KillInternal
  Call ClearColumns
  
  If Not ReInit Then
    AutoName = Trim$(AutoName)
    If Len(AutoName) = 0 Then Err.Raise ERR_INITAUTODATA, "InitAutoDataEx", "The AutoClass must have a valid Name."
    If AutoNames.IsPresent(AutoName) Then Err.Raise ERR_INITAUTODATA, "InitAutoDataEx", "The AutoClass " & AutoName & " already exists."
  End If
  m_AutoName = AutoName
  m_DBDAOPath = Trim$(DatabasePath)
  m_ReportHeader = AutoName
  
  If Not rs Is Nothing Then
    Set m_RSDefault = rs
  ElseIf Not rsRDO Is Nothing Then
    Set m_rdoRSDefault = rsRDO
  End If
  m_DefaultSQL = sql
  
  If Not gridctrl Is Nothing Then
    If StrComp(TypeName(gridctrl), "AutoGridCtrl_DAO", vbTextCompare) = 0 Then
      Set m_AutoGrid_DAO = gridctrl
      If Len(m_DBDAOPath) = 0 Then Err.Raise ERR_INITAUTODATA, "InitAutoDataEx", "The Database Name/Path must be supplied when using a DAO AutoControl."
    ElseIf StrComp(TypeName(gridctrl), "AutoGridCtrl_RDO", vbTextCompare) = 0 Then
      Set m_AutoGrid_RDO = gridctrl
    Else
      Err.Raise ERR_INITAUTODATA, "InitAutoDataEx", "InitAutoData supports only DAO/RDO AutoControls."
    End If
  End If
  
  If Not m_RSDefault Is Nothing Then
    Dim fld As field
    
    For Each fld In m_RSDefault.Fields
      If InCollection(m_PersistAutoCols, fld.Name) Then
        Set colcurrent = m_PersistAutoCols(fld.Name)
        If colcurrent.UnboundColumn Then Err.Raise ERR_INITAUTODATA, "InitAutoDataEx", "Unbound Column [" & colcurrent.DataField & "] found in recordset"
        If (Not ReInit) Or (colcurrent.dbDataType = [_TYPE_INVALID]) Then colcurrent.dbDataType = DAOtoDatatype(fld.Type)
        m_AutoCols.Add colcurrent, colcurrent.DataField
      Else
        Set colcurrent = AddEx(fld.Name, DAOtoDatatype(fld.Type), False)
      End If
    Next fld
  ElseIf Not m_rdoRSDefault Is Nothing Then
    Dim Col As rdoColumn
    
    For Each Col In m_rdoRSDefault.rdoColumns
      If InCollection(m_PersistAutoCols, Col.Name) Then
        Set colcurrent = m_PersistAutoCols(Col.Name)
        If colcurrent.UnboundColumn Then Err.Raise ERR_INITAUTODATA, "InitAutoDataEx", "Unbound Column [" & colcurrent.DataField & "] found in recordset"
        If (Not ReInit) Or (colcurrent.dbDataType = [_TYPE_INVALID]) Then colcurrent.dbDataType = RDOtoDatatype(Col.Type)
        m_AutoCols.Add colcurrent, colcurrent.DataField
      Else
        Set colcurrent = AddEx(Col.Name, RDOtoDatatype(Col.Type), False)
      End If
    Next Col
  End If
  Set colcurrent = Nothing
  
InitAutoDataEx_end:
  m_RequeryOnRefresh = False
  Call ClearCursor
  Call xReturn("InitAutoData")
  Exit Function
  
InitAutoDataEx_err:
  InitAutoDataEx = False
  Resume InitAutoDataEx_end
End Function

Public Function Add(FieldName As String, Optional ByVal DataType As DATABASE_FIELD_TYPES = TYPE_STR) As AutoCol
  Set Add = AddEx(FieldName, DataType, False)
End Function

Public Function AddUnboundColumn(FieldName As String, Optional ByVal DataType As DATABASE_FIELD_TYPES = TYPE_STR) As AutoCol
  Set AddUnboundColumn = AddEx(FieldName, DataType, False)
  m_PersistAutoCols.Add AddUnboundColumn, FieldName
  AddUnboundColumn.UnboundColumn = True
End Function

Private Sub AddUnboundColumns()
  Dim aCol As AutoCol
  
  For Each aCol In m_PersistAutoCols
    If aCol.UnboundColumn Then m_AutoCols.Add aCol, aCol.DataField
  Next aCol
End Sub

Private Function AddEx(FieldName As String, ByVal DataType As DATABASE_FIELD_TYPES, ByVal PersistOnly As Boolean) As AutoCol
  Dim NewCol As AutoCol
    
  Set NewCol = New AutoCol
  If PersistOnly Then
    m_PersistAutoCols.Add NewCol, FieldName
  Else
    m_AutoCols.Add NewCol, FieldName
  End If
  NewCol.DataField = FieldName
  NewCol.GridCaption = FieldName
  NewCol.PrintCaption = FieldName
  NewCol.dbDataType = DataType
  Set NewCol.Parent = Me
  Set AddEx = NewCol
End Function

Private Function GetAutoColumn(ByVal FieldName As String, ByVal Persist As Boolean) As AutoCol
  Dim aCol As AutoCol, dType As DATABASE_FIELD_TYPES
  
  If Persist Then
    dType = [_TYPE_INVALID]
    Set aCol = Nothing
    If InCollection(m_AutoCols, FieldName) Then
      Set aCol = m_AutoCols(FieldName)
      m_PersistAutoCols.Add aCol, FieldName
    Else
      Set aCol = AddEx(FieldName, dType, True)
    End If
  End If
  If aCol Is Nothing Then Set aCol = m_AutoCols(FieldName)
  Set GetAutoColumn = aCol
End Function

Public Sub ClearFieldFormat(ByVal FieldName As String, Optional ByVal FindColumnError As Boolean = True)
  Dim aCol As AutoCol
    
  Call xSet("ClearFieldFormat")
  On Error GoTo ClearFieldFormat_Err
  Set aCol = m_AutoCols(FieldName)
  aCol.ClearFormat
  
ClearFieldFormat_End:
  Call xReturn("ClearFieldFormat")
  Exit Sub
 
ClearFieldFormat_Err:
  If FindColumnError Then
    If Err.Number = 5 Then Err.Description = "Auto Column name not found"
    Call ErrorMessage(ERR_ERROR, Err, "ClearFieldFormat", "Error clearing Field format", "Unable to clear the format from field " & FieldName)
  End If
  Resume ClearFieldFormat_End
End Sub

Public Function AddFieldFormat(ByVal FieldName As String, vFormat As Variant, Optional ByVal FindColumnError As Boolean = True, Optional ByVal Persist As Boolean = False) As Boolean
  Dim aCol As AutoCol, sFormat As String
    
  Call xSet("AddFieldFormat")
  On Error GoTo AddFieldFormat_Err
  If Not IsNull(vFormat) Then sFormat = vFormat
  Set aCol = GetAutoColumn(FieldName, Persist)
  
  FindColumnError = True
  If Len(sFormat) > 0 Then
    FormatRemove = False
    Set AutoParser.ParseSettings.ParseParameter(1) = m_AutoCols
    Set AutoParser.ParseSettings.ParseParameter(2) = Me
    Set AutoParser.ParseSettings.ParseParameter(3) = aCol
    Call AutoParser.ParseLine(sFormat)
    If DebugMode Then
      ''apf next If ClearFormat Then aCol.FormatString = ""
      aCol.FormatString = aCol.FormatString & sFormat
    End If
  End If
  
AddFieldFormat_End:
  Call xReturn("AddFieldFormat")
  Exit Function
 
AddFieldFormat_Err:
  If FindColumnError Then
    If Err.Number = 5 Then Err.Description = "Auto Column name not found"
    Call ErrorMessage(ERR_ERROR, Err, "AddFieldFormat", "Error adding Field format", "Unable to add the format """ & sFormat & """ to field " & FieldName)
  End If
  Resume AddFieldFormat_End
  Resume
End Function

Public Function RemoveFieldFormat(ByVal FieldName As String, vFormat As Variant, Optional ByVal FindColumnError As Boolean = True) As Boolean
  Dim aCol As AutoCol, sFormat As String
    
  Call xSet("RemoveFieldFormat")
  On Error GoTo RemoveFieldFormat_Err
  If Not IsNull(vFormat) Then sFormat = vFormat
  Set aCol = GetAutoColumn(FieldName, False)
  
  FindColumnError = True
  If Len(sFormat) > 0 Then
    FormatRemove = True
    Set AutoParser.ParseSettings.ParseParameter(1) = m_AutoCols
    Set AutoParser.ParseSettings.ParseParameter(2) = Me
    Set AutoParser.ParseSettings.ParseParameter(3) = aCol
    Call AutoParser.ParseLine(sFormat)
    If DebugMode Then aCol.FormatString = aCol.FormatString & "[-" & sFormat & "]"
  End If
  
RemoveFieldFormat_End:
  Call xReturn("RemoveFieldFormat")
  Exit Function
 
RemoveFieldFormat_Err:
  If FindColumnError Then
    If Err.Number = 5 Then Err.Description = "Auto Column name not found"
    Call ErrorMessage(ERR_ERROR, Err, "RemoveFieldFormat", "Error removing Field format", "Unable to remove the format """ & sFormat & """ to field " & FieldName)
  End If
  Resume RemoveFieldFormat_End
End Function

Public Function AddReportFormat(ByVal sFormat As String) As Boolean
  
  On Error GoTo AddReportFormat_Err
  Call xSet("AddReportFormat")
  If Len(sFormat) > 0 Then
    FormatRemove = False
    Set AutoParser.ParseSettings.ParseParameter(1) = Nothing
    Set AutoParser.ParseSettings.ParseParameter(2) = Nothing
    Set AutoParser.ParseSettings.ParseParameter(3) = Me
    Call AutoParser.ParseLine(sFormat)
  End If
  
AddReportFormat_End:
  Call xReturn("AddReportFormat")
  Exit Function

AddReportFormat_Err:
  Call ErrorMessage(ERR_ERROR, Err, "AddReportFormat", "Error setting report Format", "Format not a valid ReportFormat")
  Resume AddReportFormat_End
End Function


Public Sub AddUserPopupMenuItem(ByVal Caption As String, ByVal ShowFlags As POPUP_MENU_SHOWFLAGS, Tag As Variant)
  Dim um As UserPopupMenu
    
  On Error GoTo AddUserPopupMenuItem_err
  Set um = New UserPopupMenu
  um.Caption = Caption
  um.ShowFlags = ShowFlags
  um.Tag = Tag
  Call mUserPopupMenus.Add(um, Caption)
AddUserPopupMenuItem_end:
  Exit Sub
  
AddUserPopupMenuItem_err:
  Call ErrorMessage(ERR_ERROR, Err, "AddUserPopupMenuItem", "Add User Defined Popup Menu", "Unable to add the user defined popup menu '" & Caption & "' to the auto display [" & m_AutoName & "]")
  Resume AddUserPopupMenuItem_end
End Sub

Public Function ShowGrid(Optional gridctrl As Object = Nothing, Optional ByVal ForceRefresh As Boolean = False) As Boolean
  Dim db As Database
  Dim rdc As MSRDC.MSRDC
  Dim errstring As String
  On Error GoTo ShowGrid_Err
  Call xSet("ShowGrid")
      
  If InErrorMessage Then
    ShowGrid = Not (m_Grid Is Nothing)
    GoTo ShowGrid_End
  End If
  If (m_AutoGrid_DAO Is Nothing) And (m_AutoGrid_RDO Is Nothing) Then
    If gridctrl Is Nothing Then Err.Raise ERR_SHOWGRID, "ShowGrid", "No grid control specified."
    If StrComp(TypeName(gridctrl), "AutoGridCtrl_RDO", vbTextCompare) = 0 Then Set m_AutoGrid_RDO = gridctrl
    If StrComp(TypeName(gridctrl), "AutoGridCtrl_DAO", vbTextCompare) = 0 Then Set m_AutoGrid_DAO = gridctrl
  End If
  
  If m_Grid Is Nothing Then
    If Not m_SavedGrid Is Nothing Then
      Set m_Grid = m_SavedGrid
    Else
      Set m_Grid = New AutoGrid
    End If
    Call InitScriptControl
    
    If Not m_AutoGrid_DAO Is Nothing Then
      Set m_Grid.GridControl = m_AutoGrid_DAO
      m_AutoGrid_DAO.GridDataControl.DatabaseName = m_DBDAOPath
    
      If m_RSDefault.Updatable Then
        Set db = m_WorkSpace.Databases(m_DBDAOPath)
        Call GetPrimaryKeyDAO(m_AutoCols, db, m_RSDefault)
        Set db = Nothing
      End If
    End If
    If Not m_AutoGrid_RDO Is Nothing Then
      Set m_Grid.GridControl = m_AutoGrid_RDO
      Set rdc = m_AutoGrid_RDO.GridDataControl
      If m_rdoRSDefault.Updatable Then Call GetPrimaryKeyRDO(m_AutoCols, m_rdoRSDefault)
    End If
    'Sets the Data Controls Record Set
    Set m_Grid.ParentAC = Me
    m_sfRefresh = False
    If Not SetFilterSort(GetIniEntry(m_AutoName, "Sort", ""), GetIniEntry(m_AutoName, "Filter", "")) Then GoTo ShowGrid_End
    Call m_Grid.InitGrid
    Set m_Grid.AuditInterface = mGridAuditIF
    Set m_Grid.NotifyInterface = mGridNotifyIF
  Else
    Call m_Grid.RefreshGrid(True, ForceRefresh, ForceRefresh)
    If Not m_AutoGrid_DAO Is Nothing Then m_AutoGrid_DAO.ContainerForm.ZOrder
    If Not m_AutoGrid_RDO Is Nothing Then m_AutoGrid_RDO.ContainerForm.ZOrder
  End If
  ShowGrid = True
  
ShowGrid_End:
  Set m_SavedGrid = Nothing
  Call xReturn("ShowGrid")
  Exit Function

ShowGrid_Err:
  ShowGrid = False
  Call ErrorMessage(ERR_ERROR, Err, "ShowGrid", "Show grid", "Unable to display the AutoGrid [" & m_AutoName & "]")
  Resume ShowGrid_End
  Resume
End Function

Friend Sub FindEx(ByVal aCol As AutoCol, ByVal FieldName As String, ByVal FindType As FIND_TYPES)
  Dim FindString As String
  
  
  On Error GoTo FindEx_ERR
  FindString = WhereClause(FieldName, frmFind.txtFind, aCol.dbDataType, True, , (Not m_AutoGrid_RDO Is Nothing))
  If FindType <> FT_NONE Then
    If Not m_AutoGrid_DAO Is Nothing Then
       Call FindRecordEx(m_Grid, m_AutoGrid_DAO.GridDataControl.Recordset, Nothing, FindString, FindType, aCol.dbDataType)
    End If
    If Not m_AutoGrid_RDO Is Nothing Then
      Call FindRecordEx(m_Grid, Nothing, m_AutoGrid_RDO.GridDataControl.Resultset, FindString, FindType, aCol.dbDataType)
    End If
  End If
  
FindEx_END:
  Exit Sub
FindEx_ERR:
  Call ErrorMessage(ERR_ERROR, Err, ErrorSource(Err, "FindEx"), "Find Ex", "Error in FindEx")
  Resume FindEx_END
End Sub

Public Sub Find(ByVal FieldName As String)
  Dim aCol As AutoCol
  Dim FindCaption As String
  
  On Error GoTo Find_err
  FindCaption = frmFind.Caption
  Set aCol = m_AutoCols.Item(FieldName)
  frmFind.Caption = frmFind.Caption & FieldName
  Call frmFind.Start(Me, aCol, FieldName)
  
Find_end:
  Call frmFind.KillReferences
  frmFind.Caption = FindCaption
  Exit Sub
  
Find_err:
  Call ErrorMessage(ERR_ERROR, Err, "Find", "Find value in grid column", "Error finding value in grid column.")
  Resume Find_end
End Sub

Public Sub SaveCol(Optional ByVal SaveOptions As AUTO_SAVECOLS = 0, Optional ByVal PromptUser = False)
  Dim i As Long
  Dim SaveOptionsChk(1 To 4) As Integer
  
  On Error GoTo savecol_err
  Call SetCursor(vbArrow)
  If m_SaveColumns Then
    SaveOptionsChk(1) = IIf((SaveOptions And SAVE_SORT) = SAVE_SORT, vbChecked, vbUnchecked)
    SaveOptionsChk(2) = IIf((SaveOptions And SAVE_WIDTH) = SAVE_WIDTH, vbChecked, vbUnchecked)
    SaveOptionsChk(3) = IIf((SaveOptions And SAVE_ORDER) = SAVE_ORDER, vbChecked, vbUnchecked)
    SaveOptionsChk(4) = IIf((SaveOptions And SAVE_FILTER) = SAVE_FILTER, vbChecked, vbUnchecked)
        
    If PromptUser Then
      frmSaveOptions.chkSort = SaveOptionsChk(1)
      frmSaveOptions.chkWidth = SaveOptionsChk(2)
      frmSaveOptions.chkOrder = SaveOptionsChk(3)
      frmSaveOptions.chkFilter = SaveOptionsChk(4)
      frmSaveOptions.Caption = m_AutoName
      frmSaveOptions.Show vbModal
      SaveOptionsChk(1) = frmSaveOptions.chkSort
      SaveOptionsChk(2) = frmSaveOptions.chkWidth
      SaveOptionsChk(3) = frmSaveOptions.chkOrder
      SaveOptionsChk(4) = frmSaveOptions.chkFilter
      Unload frmSaveOptions
    End If
    If (SaveOptions And SAVE_CLEAR) = SAVE_CLEAR Then
      Call ClearIniEntry(m_AutoName)
      Call ClearIniEntry("Order_" & m_AutoName)
    End If
    If SaveOptionsChk(1) Then
      Call WriteIniEntry(m_AutoName, "Sort", m_Sort)
    End If
    If SaveOptionsChk(2) Then
      If Not m_Grid Is Nothing Then
        m_Grid.SaveWidths
        For i = 1 To Me.Count
          Call WriteIniEntry(m_AutoName, Me.Item(i).DataField, CStr(Me.Item(i).ColWidth))
        Next i
      End If
    End If
    If SaveOptionsChk(3) Then
      For i = 1 To Me.Count
        Call WriteIniEntry("Order_" & m_AutoName, CStr(i), CStr(m_Grid.Display(i)))
      Next i
    End If
    If SaveOptionsChk(4) Then
      Call WriteIniEntry(m_AutoName, "Filter", m_Filter)
    End If
  End If
  
savecol_end:
  Call ClearCursor
  Exit Sub
  
savecol_err:
  Call ErrorMessage(ERR_ERROR, Err, "SaveCol", "Saving column information", "Error saving autogrid column info")
  Resume savecol_end
End Sub

Public Function SetNewRS(rs As Recordset, Optional ByVal PersistColumns As Boolean = True, Optional ByVal sql As String) As Boolean
  Dim aCol As AutoCol
  Dim savGridCtrl As Object
  
  On Error GoTo SetNewRS_err
  Call SetCursor
  Set m_SavedGrid = Nothing
  If PersistColumns Then
    For Each aCol In m_AutoCols
      If Not InCollection(m_PersistAutoCols, aCol.DataField) Then
        m_PersistAutoCols.Add aCol, aCol.DataField
      End If
    Next aCol
  End If
  Set m_RSDefault = rs
  If Not m_AutoGrid_DAO Is Nothing Then
    Set savGridCtrl = m_AutoGrid_DAO
    If Not m_Grid Is Nothing Then
      m_Grid.SaveWidths
      Set m_SavedGrid = m_Grid
    End If
    Call InitAutoDataEx(m_AutoName, rs, Nothing, m_DBDAOPath, sql, True, savGridCtrl)
    Call AddUnboundColumns
    Me.ShowGrid
  End If
  SetNewRS = True
  
SetNewRS_end:
  Set m_SavedGrid = Nothing
  Call ClearCursor
  Exit Function
  
SetNewRS_err:
  Resume SetNewRS_end
End Function

Public Function SetNewRDO(rsRDO As RDOResultset, Optional ByVal PersistColumns As Boolean = True) As Boolean
  Dim aCol As AutoCol
  Dim tmpGrid As Object
  
  On Error GoTo SetNewRDO_err
  Call SetCursor
  If PersistColumns Then
    For Each aCol In m_AutoCols
      If Not InCollection(m_PersistAutoCols, aCol.DataField) Then
        m_PersistAutoCols.Add aCol, aCol.DataField
      End If
    Next aCol
  End If
  Set m_rdoRSDefault = rsRDO
  m_DefaultSQL = ""
  If Not m_AutoGrid_RDO Is Nothing Then
    Set tmpGrid = m_AutoGrid_RDO
    If Not m_Grid Is Nothing Then m_Grid.SaveWidths
    Call InitAutoDataEx(m_AutoName, Nothing, rsRDO, "", "", True, tmpGrid)
    Call AddUnboundColumns
    Me.ShowGrid
  End If
  SetNewRDO = True
  
SetNewRDO_end:
  Call ClearCursor
  Exit Function
  
SetNewRDO_err:
  Resume SetNewRDO_end
End Function

Private Function GetUnBoundColumn(ByVal OFFSET As Long) As AutoCol
  Dim ac As AutoCol
  Dim i As Long
  
  If OFFSET < 0 Then OFFSET = 0
  For i = 1 To m_AutoCols.Count
    Set ac = m_AutoCols.Item(i)
    If ac.UnboundColumn Then
      If OFFSET = 0 Then
        Set GetUnBoundColumn = ac
        Exit Function
      End If
      OFFSET = OFFSET - 1
    End If
  Next i
End Function

' Auto Report
Public Sub InitAutoReport(ObjArray As Variant, rep As Reporter)
  Dim ac As AutoCol, i As Long, bFirstHeader As Boolean
  Dim ubColCount As Long
  
  Set m_Report = rep
  Call ClearPreviewLines(m_AutoCols)
  m_MaxSpanPage = 1
  m_GroupHeaderWidth = 0
  m_FreeSpaceX = 0
  m_DoHeaders = False
  m_DoTotals = False
  bFirstHeader = True
  For Each ac In m_AutoCols
    ac.PrintPageIndex = m_MaxSpanPage
    ac.PrintCaption_Returns = CountChar(ac.PrintCaption, vbCrLf, vbBinaryCompare)
    ac.PrintCaption_Clean = ReplaceChar(ac.PrintCaption, vbCrLf, " ")
    If ac.bGroupHeader Then
      m_DoHeaders = True
      If bFirstHeader And Not ac.bGroupPageBreak Then
        bFirstHeader = False
        ac.FirstHeader = True
      End If
    End If
    If ac.UnboundColumn Then ubColCount = ubColCount + 1
    If ac.SumType <> TYPE_NOSUM Then m_DoTotals = True
  Next ac
  Call SetColumnNoPrint(m_AutoCols, m_Grid)
  If TypeOf ObjArray Is Recordset Then
    m_PrintType = PRINT_RS
    m_lmin = 0
    m_lmax = (ObjArray.Fields.Count - 1) + ubColCount
  ElseIf TypeOf ObjArray Is RDOResultset Then
    m_PrintType = PRINT_RDO
    m_lmin = 0
    m_lmax = (ObjArray.rdoColumns.Count - 1) + ubColCount
  Else
    m_PrintType = PRINT_USER
    m_lmin = LBound(ObjArray)
    m_lmax = UBound(ObjArray)
  End If
  m_lfirstubcolumn = -1
  If ubColCount > 0 Then m_lfirstubcolumn = m_lmax - ubColCount + 1
  ReDim m_cols(m_lmin To m_lmax) As AutoCol
  For i = m_lmin To m_lmax
    If (m_lfirstubcolumn >= 0) And (i >= m_lfirstubcolumn) Then
      Set ac = GetUnBoundColumn(i - m_lfirstubcolumn)
    Else
      If m_PrintType = PRINT_RS Then
        Set ac = m_AutoCols(ObjArray.Fields(i).Name)
      ElseIf m_PrintType = PRINT_RDO Then
        Set ac = m_AutoCols(ObjArray.rdoColumns(i).Name)
      Else
        Set ac = m_AutoCols(i)
      End If
    End If
    Set m_cols(i) = ac
  Next i
  m_OnFirstPage = True
  m_CollapseThisGroup = False
  m_MaxReportLevel = RebaseReportLevels
  m_Report.PageHeader = ""
  m_Report.PageFooter = m_PageFooter
  Call m_Report.PrinterValid
  m_Report.IgnoreExportCR = True
  Call m_Report.Out(m_ReportFormat & m_ReportHeader & "{EOLN}" & vbCrLf)
  Call ResetReportSection(False)
End Sub

Public Sub ShowReport(rep As Reporter)
  On Error GoTo ShowReport_Err
  Call xSet("ShowReport")
    
  If Not rep.ReportActive() Then Call Err.Raise(ERR_REPORTACTIVE, "ShowReport", "Cannot Show report outside an InitReport/EndReport pair ")
  If Not m_rdoRSDefault Is Nothing Then
    Call PrintRDORecordsetEx(rep)
  Else
    Call PrintRecordsetEx(rep)
  End If
  
ShowReport_End:
  Call xReturn("ShowReport")
  Exit Sub

ShowReport_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ShowReport", "Error displaying AutoReport", "Error executing report.")
  Resume ShowReport_End
End Sub

Private Function FillUnboundLine(ByRef LineObj As Variant, ByVal rs As Recordset) As Boolean
  Dim aCol As AutoCol, i As Long
  
  On Error GoTo FillUnboundLine_err
  If m_lfirstubcolumn >= 0 Then
    For i = m_lmin To m_lmax
      If i < m_lfirstubcolumn Then
        LineObj(i) = rs.Fields(i).Value
      Else
        Set aCol = m_cols(i)
        LineObj(i) = GetCalculatedValue(Nothing, rs, aCol, aCol.OnUpdateCalcValue, aCol.OnUpdateCalcValueType, rs.Bookmark, True)
      End If
    Next i
    FillUnboundLine = True
    Exit Function
  End If
  FillUnboundLine = False
  Exit Function
  
FillUnboundLine_err:
  Err.Raise Err.Number, ErrorSource(Err, "FillUnboundLine"), Err.Description
End Function

Private Sub PrintRecordsetEx(ByVal rep As Reporter)
  Dim RecCount As Long, i As Long
  Dim rs As Recordset
  Dim rLine As Variant
  
  On Error GoTo PrintrecordsetEx_err
  Call xSet("PrintRecordsetEx")
  If Not m_Grid Is Nothing Then Set rs = m_Grid.DAORecordset
  If rs Is Nothing Then Set rs = m_RSDefault
  Set rs = rs.Clone
  Call InitAutoReport(rs, rep)
  ReDim rLine(m_lmin To m_lmax)
  
  If (Not (rs.EOF And rs.BOF)) Or (m_TemplatePrint <> TEMPLATE_NONE) Then
    ' get spacing use first  PreviewRows rows as examples
    RecCount = 0
    If m_IncludeColHeaders Or (m_TemplatePrint <> TEMPLATE_NONE) Then RecCount = -1
    If Not rs.BOF Then rs.MoveFirst
      If Not m_ExportReport Then
      Do
        If FillUnboundLine(rLine, rs) Then
          Call PreviewSingleLine(RecCount, rLine)
        Else
          Call PreviewSingleLine(RecCount, rs)
        End If
        If RecCount >= 0 Then rs.MoveNext
        RecCount = RecCount + 1
      Loop Until rs.EOF Or (RecCount > m_PreviewRows) Or (m_TemplatePrint = TEMPLATE_PROFORMA_ONLY)
    End If
    
    ' deal with spacing issues - get total space required
    Call DoColumnSpacing
    For i = 1 To m_MaxSpanPage
      If (i = 1) And (m_MaxSpanPage > 1) Then m_Report.PostPageNumber = "A"
      If i > 1 Then Call ResetReportSection(True)
      If Not rs.BOF Then rs.MoveFirst
      If (m_TemplatePrint = TEMPLATE_PROFORMA_ONLY) Or (rs.EOF And (m_TemplatePrint = TEMPLATE_PROFORMA_AND_DATA)) Then
        Call PrintAutoLineEx(rLine, i)
      Else
        Do
          If FillUnboundLine(rLine, rs) Then
            Call PrintAutoLineEx(rLine, i)
          Else
            Call PrintAutoLineEx(rs.Fields, i)
          End If
          rs.MoveNext
        Loop While Not rs.EOF
      End If
      Call EndAutoReport(AUTO_ENDREPORT, i)
    Next i
  Else
    Call EndAutoReport(AUTO_ENDREPORT, 1)
  End If
  
PrintrecordsetEx_end:
  Call ClearCursor
  Call xReturn("PrintRecordsetEx")
  Exit Sub
  
PrintrecordsetEx_err:
  Call ErrorMessage(ERR_ERROR + ERR_ALLOWIGNORE, Err, "printrecordset", "ERR_PRINTRECORDSET", "Error in printrecordset")
  Resume PrintrecordsetEx_end
  Resume
End Sub
      
Private Sub PrintRDORecordsetEx(rep As Reporter)
  Dim RecCount As Long
  Dim rsRDO As RDOResultset
        
  On Error GoTo PrintRDORecordsetEx_err
  If Not m_Grid Is Nothing Then
    Set rsRDO = m_Grid.RDOResultset
    Call m_Grid.CloseGrid(False)
  End If
  If rsRDO Is Nothing Then Set rsRDO = m_rdoRSDefault
  Call InitAutoReport(rsRDO, rep)
  
  If Not (rsRDO.EOF And rsRDO.BOF) Then
    ' get spacing use first  PreviewRows rows as examples
    If Not rsRDO.BOF Then Call rsRDO.Requery
    RecCount = 0
    If Not m_ExportReport Then
      If m_IncludeColHeaders Then RecCount = -1
      Do
        Call PreviewSingleLine(RecCount, rsRDO)
        If RecCount >= 0 Then rsRDO.MoveNext
        RecCount = RecCount + 1
      Loop Until rsRDO.EOF Or (RecCount > m_PreviewRows)
    End If
    
    ' deal with spacing issues - get total space required
    Call DoColumnSpacing
    If Not rsRDO.BOF Then Call rsRDO.Requery
    Do
      Call PrintAutoLine(rsRDO.rdoColumns)
      rsRDO.MoveNext
    Loop Until rsRDO.EOF
  End If
  Call EndAutoReport(AUTO_ENDREPORT, 1)
  
PrintRDORecordsetEx_end:
  If Not m_Grid Is Nothing Then
    If Not m_Grid.TDBGrid Is Nothing Then Call m_Grid.TDBGrid.ReOpen
  End If
  Call ClearCursor
  Exit Sub
  
PrintRDORecordsetEx_err:
  Call ErrorMessage(ERR_ERROR + ERR_ALLOWIGNORE, Err, "printrecordset", "ERR_PRINTRECORDSET", "Error in printrecordset")
  Resume PrintRDORecordsetEx_end
  Resume
End Sub
      
Public Sub PreviewSingleLine(ByVal RecCount As Long, LineObj As Variant)
  Dim ac As AutoCol, vValue As Variant, sValue As String
  Dim i As Long, X As Single
  
  ' get the width of the Group Headers
  If CollapseGroup Then Call GetGroupHeaderLine(LineObj, , True)
  For i = m_lmin To m_lmax
    Set ac = m_cols(i)
    If Not (ac.Hide Or ac.NoPrint Or ac.bGroupHeader) Then
      If (Not CollapseGroup) Or ac.bGroup Or (ac.SumType <> TYPE_NOSUM) Then
        If RecCount < 0 Then
          vValue = ac.PrintCaption
        Else
          vValue = LineObj(i)
        End If
        If Not IsNull(vValue) Then
          If RecCount < 0 Then
            sValue = vValue
          Else
            sValue = ac.GetFormattedValue(vValue)
          End If
          If ac.Wrap And ac.MinWidth > 0 Then
            ac.MaxWidth = ac.MinWidth
          Else
            m_Report.PreviewOnly = True
            Call m_Report.Out(m_ReportFormat & ac.CaptionFormat)
            X = m_Report.ReportTextWidth(sValue) + ADJSPACING
            If ac.Alignment = dbgRight Then X = X + ADJRIGHTALIGN
            If X < ac.MinWidth Then X = ac.MinWidth
            If X > ac.MaxWidth Then ac.MaxWidth = X
          End If
        End If
      End If
    End If
  Next i
End Sub
      
Public Sub DoColumnSpacing()
  Dim i As Long, ac As AutoCol, cCount  As Long
  Dim LastVisible As Long, VisibleFields As Long
  Dim TotalFieldWidth As Single, SquashFieldWidth As Single, PageWidth As Single
  Dim TotalInterColWidth As Single, MinFieldWidth As Single, MinUsed As Boolean, OrgFieldWidth As Single

  If m_ExportReport Then Exit Sub
  m_GroupHeaderWidth = 0
  If CollapseGroup Then
    For i = m_lmin To m_lmax
      Set ac = m_cols(i)
      If (Not (ac.Hide Or ac.NoPrint)) And ac.bGroupHeader Then
        m_GroupHeaderWidth = m_GroupHeaderWidth + ac.MaxWidth
      End If
    Next i
    If m_GroupHeaderWidth > 0 Then m_GroupHeaderWidth = m_GroupHeaderWidth + (MININTERCOLSPACE * 2)
  End If
  
redo_columnspacing:
  TotalInterColWidth = 0
  TotalFieldWidth = m_GroupHeaderWidth ' 0 if not in CollapseGroup
  SquashFieldWidth = 0: VisibleFields = 0: LastVisible = -1
  For i = m_lmin To m_lmax
    Set ac = m_cols(i)
    If Not (ac.Hide Or ac.NoPrint Or ac.bGroupHeader) Then
      VisibleFields = VisibleFields + 1
      TotalInterColWidth = TotalInterColWidth + MININTERCOLSPACE
      If ac.MaxWidth < MINCOLWIDTH Then
        If ac.MaxWidth = 0 Then
          ac.MaxWidth = MINEMPTYCOLWIDTH
        Else
          ac.MaxWidth = MINCOLWIDTH
        End If
      End If
      LastVisible = i
      If IsTrimField(ac) Then SquashFieldWidth = SquashFieldWidth + ac.MaxWidth
      If Not MinUsed Then ac.MaxWidth = ac.MaxWidth + MININTERCOLSPACE
      TotalFieldWidth = TotalFieldWidth + ac.MaxWidth
    End If
  Next i
  
  ' apportion unused/overused space
  If LastVisible >= 0 Then
    TotalFieldWidth = TotalFieldWidth + MINLEFTMARGIN
    ' Right margin is greater of MINRIGHTMARGIN, MININTERCOLSPACE
    If MINRIGHTMARGIN > MININTERCOLSPACE Then TotalFieldWidth = TotalFieldWidth + MINRIGHTMARGIN
    
    PageWidth = m_Report.ReportPageWidth
    m_FreeSpaceX = PageWidth - TotalFieldWidth
    If m_FreeSpaceX < 0 Then
      If m_SpanMultiplePages Then
        TotalFieldWidth = MINLEFTMARGIN
        m_MaxSpanPage = 1: cCount = 0
        For i = m_lmin To m_lmax
          Set ac = m_cols(i)
          If Not (ac.Hide Or ac.NoPrint Or ac.bGroupHeader) Then
            TotalFieldWidth = TotalFieldWidth + ac.MaxWidth
            If (TotalFieldWidth > PageWidth) And (cCount > 0) Then
              cCount = 0
              m_MaxSpanPage = m_MaxSpanPage + 1
              TotalFieldWidth = ac.MaxWidth + MINLEFTMARGIN
            End If
            ac.PrintPageIndex = m_MaxSpanPage
            cCount = cCount + 1
          End If
        Next i
      Else
        ' Do not include inter column width in squash calcs.
        TotalFieldWidth = TotalFieldWidth - TotalInterColWidth
        PageWidth = PageWidth - TotalInterColWidth
        
        MinUsed = False
        ' apf expand onto multiple pages
        If SquashFieldWidth <= 0 Then Call ECASE_DEBUG("Unable to compress report horizontally." & vbCrLf & "Non Text fields will be trimmed")
        If (TotalFieldWidth - SquashFieldWidth) > PageWidth Then
          Call ECASE_DEBUG("Unable to compress report horizontally as the non variable fields are too wide." & vbCrLf & "Non Text fields will be trimmed")
          SquashFieldWidth = 0
        End If
        For i = m_lmin To m_lmax
          Set ac = m_cols(i)
          If Not (ac.Hide Or ac.NoPrint Or ac.bGroupHeader) Then
            ac.MaxWidth = ac.MaxWidth - MININTERCOLSPACE
            
            If SquashFieldWidth <= 0 Then
              ac.MaxWidth = ac.MaxWidth * (PageWidth / TotalFieldWidth)
            ElseIf IsTrimField(ac) Then
              OrgFieldWidth = ac.MaxWidth
              MinFieldWidth = MINCOLWIDTH
              If OrgFieldWidth < MINCOLWIDTH Then MinFieldWidth = OrgFieldWidth
                          
              ac.MaxWidth = ac.MaxWidth * ((PageWidth - (TotalFieldWidth - SquashFieldWidth)) / SquashFieldWidth)
              
              If ac.MaxWidth < MinFieldWidth Then
                ac.MaxWidth = MinFieldWidth
                ac.TrimField = False
                MinUsed = True
              End If
            End If
            ac.MaxWidth = ac.MaxWidth + MININTERCOLSPACE
          End If
        Next i
        If MinUsed Then GoTo redo_columnspacing
      End If
    ElseIf (TotalFieldWidth > 0) And m_FitToPage Then
      TotalFieldWidth = Int((PageWidth - TotalFieldWidth) / VisibleFields)
      For i = m_lmin To m_lmax
        Set ac = m_cols(i)
        If Not (ac.Hide Or ac.NoPrint Or ac.bGroupHeader) Then ac.MaxWidth = ac.MaxWidth + TotalFieldWidth
      Next i
    End If
  End If
End Sub

Private Function RebaseReportLevels() As Long
  Dim ac As AutoCol, i As Long, CurLevel As Long
  Dim MaxReportLevel As Long, LevelPopulated As Boolean
  
  MaxReportLevel = 0
  For i = m_lmin To m_lmax
    Set ac = m_cols(i)
    ac.RebasedSumLevel = ac.SumLevel
    If ac.SumLevel > MaxReportLevel Then MaxReportLevel = ac.SumLevel
  Next i
  CurLevel = 1
  Do While CurLevel <= MaxReportLevel
    For i = m_lmin To m_lmax
      Set ac = m_cols(i)
      If ac.RebasedSumLevel = CurLevel Then
        CurLevel = CurLevel + 1
        GoTo next_level
      End If
    Next i
    ' Current level not populated
    For i = m_lmin To m_lmax
      Set ac = m_cols(i)
      If ac.RebasedSumLevel > CurLevel Then
        ac.RebasedSumLevel = ac.RebasedSumLevel - 1
      End If
    Next i
    MaxReportLevel = MaxReportLevel - 1
next_level:
  Loop
  ' levels 0 .. MaxReportLevel required
  For i = m_lmin To m_lmax
    Set ac = m_cols(i)
    Call ac.RedimSumLevels(MaxReportLevel)
  Next i
  ' redim array in AutoCols
  RebaseReportLevels = MaxReportLevel
End Function

Private Sub ResetReportSection(ByVal ForceNewPage As Boolean)
  m_OffsetsCalculated = False
  m_LastGroupHeader = ""
  m_PrevLine = ""
  m_RowCount = 0
  m_GroupRowCount = 0
  m_LineCount = 0
  m_ForceNewPage = ForceNewPage
End Sub

Public Sub PreviewAutoLine(LineObj As Variant)
  On Error GoTo PreviewAutoLine_Err
  Call xSet("PreviewAutoLine")
  
  Call AddPreviewLine(LineObj)
  
PreviewAutoLine_End:
  m_LineCount = m_LineCount + 1
  Call xReturn("PreviewAutoLine")
  Exit Sub
  
PreviewAutoLine_Err:
  Call ErrorMessage(ERR_ERROR + ERR_ALLOWIGNORE, Err, "PreviewAutoLine", "Auto Report Preview print Line", "Unable to preview print line: " & CStr(m_LineCount + 1))
  Resume PreviewAutoLine_End
End Sub

Public Sub PreviewDoColumnSpacing()
  Dim MaxLine As Long, RecCount As Long
  Dim rLine As Variant, i As Long
  On Error GoTo PreviewDoColumnSpacing_Err
  Call xSet("PreviewDoColumnSpacing")
  
  m_LineCount = 0
  RecCount = 0
  ReDim rLine(m_lmin To m_lmax)
  If m_IncludeColHeaders Then RecCount = -1
  MaxLine = Min(Preview.MaxPreviewLine, m_PreviewRows)
  For i = 1 To MaxLine
    If RecCount >= 0 Then Call GetPreviewLine(rLine, i)
    Call PreviewSingleLine(RecCount, rLine)
    RecCount = RecCount + 1
    m_LineCount = m_LineCount + 1
  Next i
  Call DoColumnSpacing
  
PreviewDoColumnSpacing_End:
  m_LineCount = m_LineCount + 1
  Exit Sub
  
PreviewDoColumnSpacing_Err:
  Call ErrorMessage(ERR_ERROR + ERR_ALLOWIGNORE, Err, "PreviewDoColumnSpacing", "Auto Report Space Preview Lines", "Unable to output preview line: " & CStr(m_LineCount + 1))
  Resume PreviewDoColumnSpacing_End
End Sub

Public Sub PreviewOutput()
  Dim ac As AutoCol
  Dim MaxLine As Long, rLine As Variant, i As Long
  On Error GoTo PreviewOutput_Err
  Call xSet("PreviewOutput")
  
  'apf does not work with m_SpanMultiplePages
  Call SortPreviewLines(m_cols, m_lmin, m_lmax)
  MaxLine = MaxPreviewLine
  For i = 1 To MaxLine
    rLine = GetForwardOnlyPreviewLine(i)
    Call PrintAutoLine(rLine)
  Next i
  
PreviewOutput_End:
  Exit Sub
  
PreviewOutput_Err:
  Call ErrorMessage(ERR_ERROR + ERR_ALLOWIGNORE, Err, "PreviewOutput", "Auto Report output preview lines", "Unable to output preview line: " & CStr(m_LineCount + 1))
  Resume PreviewOutput_End
End Sub

Public Sub PrintAutoLine(LineObj As Variant)
  Call PrintAutoLineEx(LineObj, -1)
End Sub

Private Function IsCurrentSpan(ByVal ac As AutoCol, ByVal cSpan As Long) As Boolean
  IsCurrentSpan = (cSpan < 0) Or (cSpan = ac.PrintPageIndex)
End Function

Private Sub PrintAutoLineEx(LineObj As Variant, ByVal CurrentSpan As Long)
  Static bAlternateBoxedLines As Boolean 'cad
  
  Dim i As Long, j As Long
  Dim bRowNonZero As Boolean, bHasNumberFields As Boolean
  Dim ac As AutoCol
  Dim x1 As Single, diff As Single, nWrapReturns As Long, nRet As Long
  Dim bDoGroup As Boolean, bDoGroupSum As Boolean, bDoPageBreak As Boolean
  Dim vValue As Variant
  Dim qsline As QString, sline As String, sValue As String, hline As String
  Dim sumline As String, sumrets As Long, bUseEscape As Boolean
  
  On Error GoTo PrintAutoLineEx_Err
  Call xSet("PrintAutoLineEx")
  ' change cols array into offsets (each offset is sum of prev col array values)
  ' Headers first
  If Not m_OffsetsCalculated Then
    ' m_FreeSpaceX free column space avail
    x1 = MINLEFTMARGIN + m_GroupHeaderWidth
    For i = m_lmin To m_lmax
      Set ac = m_cols(i)
      If Not (ac.Hide Or ac.NoPrint) And IsCurrentSpan(ac, CurrentSpan) Then
        If (Not CollapseGroup) Or ac.bGroup Or (ac.SumType <> TYPE_NOSUM) Then
          If ac.bGroupHeader Or (ac.bGroupHeader And ac.bGroupPageBreak) Then
            ac.pOffset = -1 ' invalid
            If ac.FirstHeader Or ac.bGroupPageBreak Then ac.pOffset = MINLEFTMARGIN
          Else
            ac.pOffset = x1
            x1 = ac.pOffset + ac.MaxWidth
          End If
          If (ac.pOffsetUser > 0) And (ac.pOffsetUser > ac.pOffset) Then
            diff = ac.pOffsetUser - ac.pOffset
            If diff > m_FreeSpaceX Then
              ac.pOffset = ac.pOffset + diff
              x1 = x1 + diff
              m_FreeSpaceX = m_FreeSpaceX - diff
            End If
          End If
        End If
      End If
    Next i
    
    ' Create PageHeader do column headers
    For i = m_lmin To m_lmax
      Set ac = m_cols(i)
      If Not (ac.Hide Or ac.NoPrint) Then ac.LastValue = LineObj(i)
    Next i
    sline = CreatePageHeader(CurrentSpan)
    m_Report.PageHeader = sline
    If m_ForceNewPage Then
      Call m_Report.Out("{NEWPAGE}")
      If m_MaxSpanPage > 1 Then
        m_Report.StartPageNumber = 1
        m_Report.PostPageNumber = Chr$(Asc("A") + CurrentSpan - 1)
      End If
    Else
      Call m_Report.Out(sline)
    End If
    If Len(sline) > 0 Then
      If (Me.AlternateBoxedLines) Then bAlternateBoxedLines = False
    End If
    m_OffsetsCalculated = True
  End If

  If m_ReportNonZeroRowsOnly Then
    bRowNonZero = False
    bHasNumberFields = False
    For i = m_lmin To m_lmax
      Set ac = m_cols(i)
      vValue = LineObj(i)
      If (Not (ac.Hide Or ac.NoPrint)) And (Not IsNull(vValue)) Then
        If Not ac.bGroupHeader And ((Not CollapseGroup) Or ac.bGroup) Then
          sValue = ac.GetFormattedValue(vValue)
          ' Does this row contain only zero value numberic fields
          If IsNumberField(ac.dbDataType) Then
            bHasNumberFields = True
            bRowNonZero = Not IsZero(ac.GetValue(vValue))
            If bRowNonZero Then Exit For
          End If
        End If
      End If
    Next i
    ' if non zero skip line
    If bHasNumberFields And (Not bRowNonZero) Then GoTo PrintAutoLineEx_End
  End If
  
  Set qsline = New QString
  If Me.AlternateBoxedLines And Not bAlternateBoxedLines Then
    Call qsline.Append("{PUSH}{COLORTBFONTHEIGHT=100,,12632256,12632256,L}{POP}")
  End If
  bAlternateBoxedLines = Not bAlternateBoxedLines
'

  nWrapReturns = 0
  bDoGroup = False:   bDoGroupSum = False: bDoPageBreak = False
  m_CollapseLastGroup = CollapseGroup
  m_CollapseThisGroup = False
  For i = m_lmin To m_lmax
    Set ac = m_cols(i)
    vValue = LineObj(i)
    If Not (ac.Hide Or ac.NoPrint) Then
      If (Not IsNull(vValue)) And IsCurrentSpan(ac, CurrentSpan) Then
        If Not ac.bGroupHeader Then
          If (Not CollapseGroup) Or ac.bGroup Then
            sValue = ac.GetFormattedValue(vValue)
            If ac.FixedWidth > 0 Then
              If ac.FW_PadLeft Then
                qsline.Append "{FIXEDL=" & CStr(ac.FixedWidth) & "}"
              Else
                qsline.Append "{FIXEDR=" & CStr(ac.FixedWidth) & "}"
              End If
            End If
            qsline.Append m_ReportFormat & ac.DataFormat
            If m_ExportReport Then
              ' do nothing no formatting required
            ElseIf ac.Alignment = ALIGN_RIGHT Then
              qsline.Append AbsXCoord(ac.pOffset + ac.MaxWidth - MININTERCOLSPACE) & "{RIGHT}"
            ElseIf ac.Alignment = ALIGN_CENTER Then
              qsline.Append AbsXCoord(ac.pOffset) & "{CENTER=" & XCoord(ac.MaxWidth - MININTERCOLSPACE) & "}"
            ElseIf Not IsTrimField(ac) Or (Len(sValue) = 0) Or ac.Wrap Then
              qsline.Append AbsXCoord(ac.pOffset)
            Else
              qsline.Append AbsXCoord(ac.pOffset) & "{" & ac.TrimType & "=" & XCoord(ac.MaxWidth - MININTERCOLSPACE) & "}"
            End If
            If (ac.Alignment = ALIGN_LEFT) And Not m_ExportReport Then qsline.Append "{LEFT}"
            nRet = 0
            If ac.Wrap Then
              nRet = GetWrapText(sValue, m_ReportFormat & ac.DataFormat, ac.MaxWidth - MININTERCOLSPACE)
              If nRet > nWrapReturns Then nWrapReturns = nRet
            End If
            bUseEscape = (InStr(1, sValue, "{", vbBinaryCompare) > 0)
            If bUseEscape Then qsline.Append "{ESCAPE}"
            If nRet > 0 Then
              qsline.Append "{PUSH}{STATICX}" & sValue & "{POP}"
            Else
              qsline.Append sValue
            End If
            If bUseEscape Then qsline.Append "{ESCAPE}"
          End If
        End If
      End If
      ' Sums track one value in arrears ...
      If (ac.SumType <> TYPE_NOSUM) And (m_TemplatePrint <> TEMPLATE_PROFORMA_ONLY) Then
        If IsCurrentSpan(ac, CurrentSpan) Then
          For j = 0 To m_MaxReportLevel
            ac.SumGroupCount(j) = ac.SumGroupCount(j) + 1
            ac.SumGroup(j) = ac.SumGroup(j) + ac.SumLast
          Next j
          ac.SumTotal = ac.SumTotal + ac.SumLast
          ac.SumLast = ac.GetValue(vValue)
        End If
      End If
      If ac.bGroup Then
        If Not IsEmpty(ac.CollapseLike) Then
          For j = 1 To UBound(ac.CollapseLike)
            If ac.CollapseLike(j) Like vValue Then m_CollapseThisGroup = True
          Next j
        End If
        If Not IsValueEqual(ac.LastValue, vValue) Then
          bDoGroup = True
          ac.DoGroup = True
          If ac.bGroupTotal Then bDoGroupSum = True
          If ac.bGroupPageBreak Then bDoPageBreak = True
          ac.LastGroupValue = ac.LastValue
        End If
        ac.LastValue = vValue
      End If
      If i <> m_lmax And Not ac.bGroupHeader Then qsline.Append "{COMMA}"
    End If
  Next i
      
  hline = ""
  If (bDoGroup Or (m_RowCount = 0)) And m_DoHeaders Then
    Call GetGroupHeaderLine(LineObj, hline)
    If Not CollapseGroup And m_PrintGroupHeader Then hline = hline & "{EOLN}" & vbCrLf & vbCrLf
    ' Header has not changed
    If Not bDoPageBreak And (StrComp(m_LastGroupHeader, hline, vbTextCompare) = 0) Then hline = ""
  End If
  If bDoGroup Then
    If bDoGroupSum Then
      Call OutputGroup(sumline, sumrets, False, -1, CurrentSpan)
      Call m_Report.Out("{KEEPTOGETHER=" & CStr(sumrets + nWrapReturns + 1 + mLineSpacing - 1) & "}")
      If Len(m_PrevLine) > 0 And (Not m_CollapseLastGroup) Then
        Call m_Report.Out(m_PrevLine)
      End If
      m_PrevLine = ""
      Call m_Report.Out(sumline)
    Else
      Call m_Report.Out("{KEEPTOGETHER=" & CStr(nWrapReturns + 1 + mLineSpacing - 1) & "}")
      If Len(m_PrevLine) > 0 And (Not m_CollapseLastGroup) Then
        Call m_Report.Out(m_PrevLine)
      End If
      m_PrevLine = ""
    End If
    If m_CollapseLastGroup Then
      Call OutputCollapseText(sumline, CurrentSpan)
      If Len(sumline) > 0 Then
        Call m_Report.Out(sumline)
      End If
    End If
            
    ' don't print if Page Break would result
    If Not m_Report.TestPageBreak(1) Then
      Call m_Report.Out("{EOLN}" & vbCrLf)
    End If
    m_GroupRowCount = 0
  End If
  
  If Len(hline) > 0 Then ' need to recalc header line as it is different
    If (m_RowCount = 0) Or bDoGroup Then
      If m_RowCount = 0 Then
        Call m_Report.Out(hline)
      End If
      m_LastGroupHeader = hline
      m_Report.PageHeader = CreatePageHeader(CurrentSpan)
    End If
  End If
  If bDoGroup Then
    If bDoPageBreak Then
      Call m_Report.Out("{NEWPAGE}")
    ElseIf Len(hline) > 0 Then
      If Not m_Report.Out("{KEEPTOGETHER=4}") Then Call m_Report.Out(m_LastGroupHeader)
    End If
  End If
  ' Delay output to allow correct grouping apf
  If Len(m_PrevLine) > 0 And (Not m_CollapseLastGroup) Then
    Call m_Report.Out(m_PrevLine)
  End If
  If nWrapReturns > 0 Then Call m_Report.Out("{KEEPTOGETHER=" & (nWrapReturns + 1 + mLineSpacing - 1) & "}")
  m_PrevLine = qsline
  If Not CollapseGroup Then m_PrevLine = m_PrevLine & "{EOLN}" & vbCrLf
  Do While nWrapReturns > 0
    m_PrevLine = m_PrevLine & vbCrLf
    nWrapReturns = nWrapReturns - 1
  Loop
  If (Not CollapseGroup) And (mLineSpacing > 1) Then
    ' add extra returns after each line
    nWrapReturns = mLineSpacing - 1
    Do While nWrapReturns > 0
      m_PrevLine = m_PrevLine & vbCrLf
      nWrapReturns = nWrapReturns - 1
    Loop
  End If
  
  m_RowCount = m_RowCount + 1
  m_GroupRowCount = m_GroupRowCount + 1
  
PrintAutoLineEx_End:
  m_LineCount = m_LineCount + 1
  Call xReturn("PrintAutoLineEx")
  Exit Sub

PrintAutoLineEx_Err:
  Call ErrorMessage(ERR_ERROR + ERR_ALLOWIGNORE, Err, "PrintAutoLineEx", "Auto Report Print Line", "Unable to print line: " & CStr(m_LineCount + 1))
  Resume PrintAutoLineEx_End
  Resume
End Sub

Private Function GetWrapText(sValue As String, ByVal FormatString As String, ByVal ColumnWidth As Single)
  Dim X As Single
  
  m_Report.PreviewOnly = True
  Call m_Report.Out(FormatString)
  GetWrapText = m_Report.WrapTextToWidth(sValue, ColumnWidth, " ,;:-=")
End Function

Private Sub GetGroupHeaderLine(LineObj As Variant, Optional hline As String, Optional ByVal WidthOnly As Boolean = False)
  Dim i As Long
  Dim ac As AutoCol, X As Single
  Dim vValue As Variant, sValue As String

  If Not m_PrintGroupHeader Then
    hline = ""
    Exit Sub
  End If
  If WidthOnly Then
    For i = m_lmin To m_lmax
      Set ac = m_cols(i)
      vValue = LineObj(i)
      If Not (ac.Hide Or ac.NoPrint) Then
        If Not IsNull(vValue) Then
          If ac.bGroupHeader And Not ac.bGroupPageBreak Then
            m_Report.PreviewOnly = True
            Call m_Report.Out(m_ReportFormat & ac.DataFormat)
            If ac.FirstHeader Then
              If (Len(ac.PrintCaption_Clean) > 0) Or m_PrintFieldSuffix Then sValue = ac.PrintCaption_Clean & m_GroupHeaderFieldSuffix
            Else
              sValue = m_GroupHeaderSeparator
              If (Len(ac.PrintCaption_Clean) > 0) Or m_PrintFieldSuffix Then sValue = sValue & ac.PrintCaption_Clean & m_GroupHeaderFieldSuffix
            End If
            sValue = sValue & ac.GetFormattedValue(vValue)
            X = m_Report.ReportTextWidth(sValue) + ADJSPACING
            If X < ac.MinWidth Then X = ac.MinWidth
            If X > ac.MaxWidth Then ac.MaxWidth = X
          End If
        End If
      End If
    Next i
  Else
    For i = m_lmin To m_lmax
      Set ac = m_cols(i)
      vValue = LineObj(i)
      If Not (ac.Hide Or ac.NoPrint) Then
        If ac.bGroupHeader And Not ac.bGroupPageBreak Then
          hline = hline & m_ReportFormat & ac.DataFormat
          If ac.FirstHeader Then
            hline = hline & "{LEFT}" & AbsXCoord(ac.pOffset)
            If (Len(ac.PrintCaption_Clean) > 0) Or m_PrintFieldSuffix Then hline = hline & ac.PrintCaption_Clean & m_GroupHeaderFieldSuffix
          Else
            hline = hline & m_GroupHeaderSeparator
            If (Len(ac.PrintCaption_Clean) > 0) Or m_PrintFieldSuffix Then hline = hline & ac.PrintCaption_Clean & m_GroupHeaderFieldSuffix
          End If
          If IsNull(vValue) Then
            sValue = "(Null)"
          Else
            sValue = ac.GetFormattedValue(vValue)
          End If
          hline = hline & "{B+}" & sValue & "{B-}{COMMA}"
        End If
      End If
    Next i
  End If
End Sub

Private Sub OutputCollapseText(CollapseText As String, ByVal CurrentSpan As Long)
  Dim ac As AutoCol, i As Long, nRet As Long
  Dim sValue As String, nWrapWidth As Single
  
  Call xSet("OutputCollapseText")
  CollapseText = ""
  nWrapWidth = m_Report.ReportPageWidth
  For i = m_lmin To m_lmax
    Set ac = m_cols(i)
    If IsCurrentSpan(ac, CurrentSpan) Then
      If ac.SumType <> TYPE_NOSUM Then
        nWrapWidth = ac.pOffset
        Exit For
      End If
    End If
  Next i
  For i = m_lmin To m_lmax
    Set ac = m_cols(i)
    If IsCurrentSpan(ac, CurrentSpan) Then
      sValue = ac.CollapseText
      If Len(sValue) > 0 Then
        nRet = GetWrapText(sValue, m_ReportFormat & ac.DataFormat, nWrapWidth - MININTERCOLSPACE) + 1
        Call m_Report.Out("{KEEPTOGETHER=" & CStr(nRet + 1) & "}")
        CollapseText = m_ReportFormat & ac.DataFormat & "{PUSH}" & AbsXCoord(ac.pOffset) & "{STATICX}" & sValue & "{POP}"
        Do While nRet > 0
          CollapseText = CollapseText & vbCrLf
          nRet = nRet - 1
        Loop
        Exit For
      End If
    End If
  Next i
  Call xReturn("OutputCollapseText")
End Sub

Private Sub OutputGroup(GroupString As String, nReturns As Long, ByVal bFinalTotal As Boolean, ByVal TotalItems As Long, ByVal CurrentSpan As Long)
  Dim ac As AutoCol
  Dim ac2 As AutoCol, j As Long
  Dim i As Long, CurLevel As Long, LevelFired As Boolean, HigherLevelFired As Boolean
  Dim LowerLevelFired As Boolean
    
  LowerLevelFired = False
  For CurLevel = 0 To m_MaxReportLevel
    For i = m_lmin To m_lmax
      Set ac = m_cols(i)
      If (ac.RebasedSumLevel = CurLevel) And ac.bGroupTotal Then
        If LowerLevelFired Then
          ac.DoGroup = True
        ElseIf ac.DoGroup Then
          LowerLevelFired = True
        End If
      End If
    Next i
  Next CurLevel
  
  nReturns = 0: GroupString = "": HigherLevelFired = False
  For CurLevel = m_MaxReportLevel To 0 Step -1
    LevelFired = False
    For i = m_lmin To m_lmax
      Set ac = m_cols(i)
      If bFinalTotal Then
        Call OutputGroupLevel(GroupString, nReturns, Nothing, TotalItems, HigherLevelFired, CurrentSpan)
        GoTo OutputGroup_end
      End If
      If ac.DoGroup And (ac.RebasedSumLevel = CurLevel) Then
        If Not LevelFired Then
          Call OutputGroupLevel(GroupString, nReturns, ac, TotalItems, HigherLevelFired, CurrentSpan)
          LevelFired = True
          HigherLevelFired = True
          ' Clear sum/counts at this level
          For j = m_lmin To m_lmax
            Set ac2 = m_cols(j)
            ac2.SumGroup(CurLevel) = 0
            ac2.SumGroupCount(CurLevel) = 0
          Next j
        End If
        ac.DoGroup = False
      End If
    Next i
  Next CurLevel
OutputGroup_end:
End Sub

Private Sub OutputGroupLevel(GroupString As String, nReturns As Long, ByVal GroupAC As AutoCol, ByVal TotalItems As Long, ByVal HigherLevelFired As Boolean, ByVal CurrentSpan As Long)
  Dim ac As AutoCol, GroupLevel As Long
  Dim i As Long, sValue As String
  Dim sout0 As String, sout1 As String, sout2 As String
  Dim dsum As Double
  
  sout0 = "": sout1 = "": sout2 = ""
  GroupLevel = -1
  If Not GroupAC Is Nothing Then
    GroupLevel = GroupAC.RebasedSumLevel
    If IsCurrentSpan(GroupAC, CurrentSpan) Then
      If m_CollapseLastGroup And Not HigherLevelFired Then
        sout1 = m_PrevLine
      Else
        If m_PrintTotalTitle Then
          sout1 = m_ReportFormat & GroupAC.DataFormat & AbsXCoord(MINLEFTMARGIN)
          If m_PrintTotalAsValue Then
            sout1 = sout1 & "{B+}" & GroupAC.GetFormattedValue(GroupAC.LastGroupValue) & "{B-} " & m_PrintTotalText
          Else
            sValue = GroupAC.PrintCaption_Clean & " " & m_PrintTotalText
            If GroupAC.Wrap Then Call GetWrapText(sValue, m_ReportFormat & GroupAC.DataFormat, GroupAC.MaxWidth - MININTERCOLSPACE)
            sout1 = sout1 & sValue
          End If
        End If
      End If
    End If
  Else
    If CurrentSpan <= 1 Then sout1 = m_ReportFormat & AbsXCoord(MINLEFTMARGIN) & m_PrintTotalText
  End If
  For i = m_lmin To m_lmax
    Set ac = m_cols(i)
    If (ac.SumType <> TYPE_NOSUM) And (ac.MaxWidth > 0) And IsCurrentSpan(ac, CurrentSpan) Then
      If ac.FixedWidth > 0 Then
        If ac.FW_PadLeft Then
          sout1 = sout1 & "{FIXEDL=" & CStr(ac.FixedWidth) & "}"
        Else
          sout1 = sout1 & "{FIXEDR=" & CStr(ac.FixedWidth) & "}"
        End If
      End If
      'If ((m_GroupRowCount <> 1) Or CollapseGroup Or HigherLevelFired) Or (GroupAC Is Nothing) Then
      If ((m_GroupRowCount <> 1) Or m_CollapseLastGroup Or HigherLevelFired) Or (GroupAC Is Nothing) Then
        sout0 = sout0 & AbsXCoord(ac.pOffset) & "{LINEABS=" & XCoord(ac.MaxWidth - MININTERCOLSPACE) & "}"
        sout2 = sout2 & AbsXCoord(ac.pOffset) & "{LINEABS=" & XCoord(ac.MaxWidth - MININTERCOLSPACE) & ",D}"
        If GroupAC Is Nothing Then
          dsum = ac.SumTotal
        Else
          dsum = ac.SumGroup(GroupLevel)
        End If
        If ac.SumType = TYPE_MEAN Then
          If TotalItems < 0 Then TotalItems = ac.SumGroupCount(GroupLevel)
          dsum = dsum / TotalItems
        End If
      End If
      'If ((m_GroupRowCount <> 1) Or CollapseGroup Or HigherLevelFired) Or (GroupAC Is Nothing) Then
      If ((m_GroupRowCount <> 1) Or m_CollapseLastGroup Or HigherLevelFired) Or (GroupAC Is Nothing) Then
        sout1 = sout1 & m_ReportFormat & ac.DataFormat
        If ac.Alignment = ALIGN_RIGHT Then
          sout1 = sout1 & AbsXCoord(ac.pOffset + ac.MaxWidth - MININTERCOLSPACE) & "{RIGHT}"
        ElseIf ac.Alignment = ALIGN_CENTER Then
          sout1 = sout1 & AbsXCoord(ac.pOffset) & "{CENTER=" & XCoord(ac.MaxWidth - MININTERCOLSPACE) & "}"
        Else ' LEFT
          sout1 = sout1 & AbsXCoord(ac.pOffset)
        End If
        sout1 = sout1 & Format$(dsum, ac.Format)
      End If
      If i <> m_lmax Then sout1 = sout1 & "{COMMA}"
    End If
  Next i
  
  'If CollapseGroup Then
  If m_CollapseLastGroup Then
    GroupString = GroupString & sout1 & "{EOLN}" & vbCrLf
  ElseIf ((m_GroupRowCount <> 1) Or HigherLevelFired) Or (GroupAC Is Nothing) Then
    If Len(sout0) > 0 Then
      GroupString = GroupString & "{YREL=50}{STARTSKIPEXPORT}" & sout0 & m_ReportFormat & "{ENDSKIPEXPORT}{YREL=100}" & sout1 & "{EOLN}" & vbCrLf & "{YREL=100}"
      nReturns = nReturns + 3
      If GroupAC Is Nothing Then
        GroupString = GroupString & "{STARTSKIPEXPORT}" & sout2 & "{ENDSKIPEXPORT}"
        nReturns = nReturns + 1
      End If
    End If
  Else
    GroupString = GroupString & "{YREL=50}"
  End If
End Sub

Private Function CreatePageHeader(ByVal CurrentSpan As Long) As String
  Dim returnsreq As Long
  Dim i As Long, s As String
  Dim ac As AutoCol
  
  On Error GoTo CreatePageHeader_err
  returnsreq = 0
  s = "{LEFT}"
  For i = m_lmin To m_lmax
    Set ac = m_cols(i)
    If Not (ac.Hide Or ac.NoPrint Or ac.bGroupHeader) And IsCurrentSpan(ac, CurrentSpan) Then
      If (Not CollapseGroup) Or ac.bGroup Or (ac.SumType <> TYPE_NOSUM) Then
        If ac.FixedWidth > 0 Then
          If ac.FW_PadLeft Then
            s = s & "{FIXEDL=" & CStr(ac.FixedWidth) & "}"
          Else
            s = s & "{FIXEDR=" & CStr(ac.FixedWidth) & "}"
          End If
        End If
        s = s & m_ReportFormat & ac.CaptionFormat
        If m_TrimColHeaders And (ac.PrintCaption_Returns = 0) Then
          s = s & AbsXCoord(ac.pOffset) & "{TRIM=" & XCoord(ac.MaxWidth - MININTERCOLSPACE) & "}" & ac.PrintCaption
        Else
          If m_AlignColHeaders And ((ac.Alignment = ALIGN_RIGHT) Or (ac.Alignment = ALIGN_CENTER)) Then
            If ac.Alignment = ALIGN_RIGHT Then
              s = s & "{PUSH}" & AbsXCoord(ac.pOffset + ac.MaxWidth - MININTERCOLSPACE) & "{RIGHT}" & "{STATICX}" & ac.PrintCaption & "{POP}"
            Else ' center
              s = s & "{PUSH}" & AbsXCoord(ac.pOffset) & "{STATICX}" & "{CENTER=" & XCoord(ac.MaxWidth - MININTERCOLSPACE) & "}" & ac.PrintCaption & "{POP}"
            End If
          Else
            s = s & "{PUSH}" & AbsXCoord(ac.pOffset) & "{STATICX}" & ac.PrintCaption & "{POP}"
          End If
        End If
        If ac.PrintCaption_Returns > returnsreq Then returnsreq = ac.PrintCaption_Returns
        If i <> m_lmax Then s = s & "{COMMA}"
      End If
    End If
  Next i
  If Len(m_PrePageHeader) > 0 Then s = m_PrePageHeader & vbCrLf & "{EOLN}" & s
  ' MPS changed: If m_ReportHeaderOnAllPages And (m_RowCount > 0) Then s = m_ReportFormat & m_ReportHeader & "{EOLN}" & vbCrLf & s
  If m_ReportHeaderOnAllPages And Not m_OnFirstPage Then s = m_ReportFormat & m_ReportHeader & "{EOLN}" & vbCrLf & s
  m_OnFirstPage = False
  CreatePageHeader = s
  
  s = "{EOLN}" & vbCrLf & "{EOLN}" & vbCrLf
  For i = 1 To returnsreq
    s = s & vbCrLf
  Next i
  CreatePageHeader = CreatePageHeader & s & m_PostPageHeader
  ' Do GroupHeaders on PageBreaks
  s = ""
  For i = m_lmin To m_lmax
    Set ac = m_cols(i)
    If ac.bGroupHeader And ac.bGroupPageBreak And IsCurrentSpan(ac, CurrentSpan) Then
      s = s & m_ReportFormat & ac.DataFormat & AbsXCoord(ac.pOffset) & ac.PrintCaption_Clean & " = {B+}"
      If Not IsNull(ac.LastValue) Then
        s = s & "{ESCAPE}" & ac.GetFormattedValue(ac.LastValue) & "{ESCAPE}"
      Else
        s = s & "(Null)"
      End If
      s = s & "{B-}" & "{EOLN}" & vbCrLf & "{YREL=60}"
    End If
  Next i
  If Len(m_LastGroupHeader) > 0 Then s = "{LEFT}" & s & m_LastGroupHeader

CreatePageHeader_end:
  CreatePageHeader = CreatePageHeader & s
  Exit Function
  
CreatePageHeader_err:
  'CreatePageHeader = "Error in CreatePageHeader " & Err.Description
  Err.Raise Err.Number, ErrorSource(Err, "CreatePageHeader"), Err.Description
End Function

Public Sub EndAutoLine(Optional ByVal EndOption As ENDAUTOLINE_OPTIONS)
  Call EndAutoReport(EndOption, 1)
End Sub

Private Sub EndAutoReport(ByVal EndOption As ENDAUTOLINE_OPTIONS, ByVal CurrentSpan As Long)
  Dim bDoGroupSum As Boolean
  Dim lmin As Long, lmax As Long, i As Long, j As Long
  Dim ac As AutoCol, ExtraReturns As Long
  Dim sumline As String, sumrets As Long

  ExtraReturns = 2
  If Len(m_PrevLine) > 0 Then Call m_Report.Out(m_PrevLine)
  If m_PrintRecordCount Then ExtraReturns = ExtraReturns + 1
  If m_DoTotals And (m_TemplatePrint <> TEMPLATE_PROFORMA_ONLY) Then
    bDoGroupSum = False
    For i = m_lmin To m_lmax
      Set ac = m_cols(i)
      If IsCurrentSpan(ac, CurrentSpan) Then
        If ac.SumType <> TYPE_NOSUM Then
          For j = 0 To m_MaxReportLevel
            ac.SumGroupCount(j) = ac.SumGroupCount(j) + 1
            ac.SumGroup(j) = ac.SumGroup(j) + ac.SumLast
          Next j
          ac.SumTotal = ac.SumTotal + ac.SumLast
        End If
        If ac.bGroup Then ac.DoGroup = True
        If ac.bGroupTotal Then bDoGroupSum = True
      End If
    Next i
    If bDoGroupSum Then
      Call OutputGroup(sumline, sumrets, False, -1, CurrentSpan)
      Call m_Report.Out("{KEEPTOGETHER=" & CStr(sumrets + ExtraReturns) & "}")
      Call m_Report.Out(sumline)
      Call m_Report.Out("{EOLN}" & vbCrLf)
    End If
    If m_CollapseLastGroup Then
      Call OutputCollapseText(sumline, CurrentSpan)
      If Len(sumline) > 0 Then Call m_Report.Out(sumline)
    End If
    If m_DoFinalTotal And bDoGroupSum Then Call m_Report.Out(vbCrLf)
    If m_DoFinalTotal Then
      Call OutputGroup(sumline, sumrets, True, m_RowCount, CurrentSpan)
      Call m_Report.Out("{KEEPTOGETHER=" & CStr(sumrets + ExtraReturns) & "}")
      Call m_Report.Out(sumline)
      If m_PrintRecordCount Then Call m_Report.Out("{EOLN}" & vbCrLf & vbCrLf)
    End If
  End If
  If m_PrintRecordCount And (m_TemplatePrint <> TEMPLATE_PROFORMA_ONLY) Then Call m_Report.Out(m_ReportFormat & "{LEFT}Total record count = " & CStr(m_RowCount) & "{EOLN}")
  If (EndOption = AUTO_ENDREPORT) And (CurrentSpan = m_MaxSpanPage) Then
    For i = m_lmin To m_lmax
      Set m_cols(i) = Nothing
    Next i
    m_lmin = 0: m_lmax = 0
    Set m_Report = Nothing
  Else
    Call ResetReportSection(EndOption = AUTO_PAGEBREAK)
  End If
End Sub
                    
Public Sub RemoveFilterSort()
  
  m_sfRefresh = True
  Call SetFilterSort("", "")
  
End Sub

' apf extra requery is not neccessary
Private Sub RefreshDataSource(Optional ByVal sql As String)
  If Not m_AutoGrid_DAO Is Nothing Then
    If m_RSDefault.Restartable And m_RequeryOnRefresh Then Call m_RSDefault.Requery
    If Len(m_DefaultSQL) = 0 Then
      If Len(m_RSDefault.Name) = 256 Then Err.Raise ERR_SQLTOOLONG, "RefreshDataSource", "Unable to refresh DAO based data source as the SQL string is too long (>256) characters"
      m_DefaultSQL = m_RSDefault.Name
    End If
  End If
  If Not m_AutoGrid_RDO Is Nothing Then
    If m_rdoRSDefault.Restartable And m_RequeryOnRefresh Then m_rdoRSDefault.Requery
    If Len(m_DefaultSQL) = 0 Then
      If (Len(m_rdoRSDefault.Name) = 256) And (Len(sql) = 0) Then Err.Raise ERR_SQLTOOLONG, "RefreshDataSource", "Unable to refresh RDO based data source as the SQL string is too long (>256) characters"
      If Len(sql) = 0 Then sql = m_rdoRSDefault.Name
      m_DefaultSQL = sql
    End If
  End If
  m_RequeryOnRefresh = False
End Sub

Public Function SetFilterSort(Optional ByVal Sort As Variant, Optional ByVal Filter As Variant) As Boolean
  Dim NewValue As String, sql As String
  Dim dat As MSRDC.MSRDC
  Dim grid As TrueDBGrid60.TDBGrid
  Dim bRefreshRDOData As Boolean
  Dim rs As RDO.RDOResultset
  
  On Error GoTo SetFilterSort_err
  Call xSet("SetFilterSort")
  Call SetCursor
  
  If Not IsMissing(Sort) Then
    NewValue = CStr(Sort)
    If Len(NewValue) > 0 Then NewValue = MergeSortString(m_Sort, NewValue)
    If Not m_RSDefault Is Nothing Then m_RSDefault.Sort = NewValue
    m_Sort = NewValue
  End If
  
  If Not IsMissing(Filter) Then
    NewValue = CStr(Filter)
    If Len(NewValue) > 0 Then NewValue = MergeFilterString(m_Filter, NewValue)
    If Not m_RSDefault Is Nothing Then m_RSDefault.Filter = NewValue
    m_Filter = NewValue
  End If
  
  If Not m_Grid Is Nothing Then
    m_Grid.SaveWidths
    If Not m_AutoGrid_DAO Is Nothing Then
      Call RefreshDataSource
      If Not m_RSDefault Is Nothing Then
        Set m_AutoGrid_DAO.GridDataControl.Recordset = m_RSDefault.OpenRecordset
        m_AutoGrid_DAO.GridDataControl.Caption = CStr(m_AutoGrid_DAO.grid.Row + 1) & " of " & CStr(m_AutoGrid_DAO.GridDataControl.Recordset.RecordCount)
      Else
        Set m_AutoGrid_DAO.GridDataControl.Recordset = Nothing
        m_AutoGrid_DAO.GridDataControl.Caption = "0 of 0"
      End If
      m_AutoGrid_DAO.FilterLabel = "Filter: " & StrDupChar(m_Filter, "&")
      m_AutoGrid_DAO.FilterLabel.ToolTipText = m_AutoGrid_DAO.FilterLabel
      m_AutoGrid_DAO.SortLabel = "Sort: " & m_Sort
      m_AutoGrid_DAO.SortLabel.ToolTipText = m_AutoGrid_DAO.SortLabel
    End If
    If Not m_AutoGrid_RDO Is Nothing Then
      Call SetRDOControl(m_AutoGrid_RDO.GridDataControl, m_rdoRSDefault)
      Call RefreshDataSource(m_AutoGrid_RDO.GridDataControl.sql)
      sql = GenerateSortFilterSQL(m_DefaultSQL, m_Filter, m_Sort)
      
      ' ATW 20/11/2000
      If Not IsMissing(Filter) Then
        If Len(Filter) = 0 Then
          Set dat = m_AutoGrid_RDO.GridDataControl
          Set rs = dat.Resultset
          bRefreshRDOData = rs.EOF And rs.BOF
        End If
      End If
      Call SetRDOControlSQL(m_rdoRSDefault.ActiveConnection, m_AutoGrid_RDO.GridDataControl, m_rdoRSDefault, sql, m_DefaultSQL)
      If bRefreshRDOData Then Call dat.Refresh
      
      m_AutoGrid_RDO.GridDataControl.Caption = CStr(m_AutoGrid_RDO.grid.Row + 1) & " of " & CStr(m_AutoGrid_RDO.GridDataControl.Resultset.RowCount)
      m_AutoGrid_RDO.FilterLabel = "Filter: " & ReplaceString(StrDupChar(m_Filter, "&"), "sf_table", "")
      m_AutoGrid_RDO.FilterLabel.ToolTipText = m_AutoGrid_RDO.FilterLabel
      m_AutoGrid_RDO.SortLabel = "Sort: " & m_Sort
      m_AutoGrid_RDO.SortLabel.ToolTipText = m_AutoGrid_RDO.SortLabel
      If m_sfRefresh Then Call m_Grid.RefreshUpdateMode
    End If
    If m_sfRefresh And False Then Call m_Grid.RefreshGrid(True, False, False)
  End If
  m_SaveColumns = True
  SetFilterSort = True
      
SetFilterSort_end:
  Call ClearCursor
  Call xReturn("SetFilterSort")
  Exit Function
  
SetFilterSort_err:
  SetFilterSort = False
  If Err.Number <> ERR_RDOSETUP Then
    Call ErrorMessage(ERR_ERROR, Err, "SetFilterSort", "Set the Filter and Sort", "Error setting Filter/Sort")
  End If
  Resume SetFilterSort_end
  Resume
End Function


Public Sub Kill()
  Static inkill As Boolean
  Dim i As Long
  
  On Error Resume Next
  If inkill Then Exit Sub
  inkill = True
  Call Me.KillInternal
  Set mGridAuditIF = Nothing
  Set mGridNotifyIF = Nothing
  Set m_AutoCols = Nothing
  Set m_PersistAutoCols = Nothing
  Set mUserPopupHandler = Nothing
  Set mUserPopupMenus = Nothing
  If m_AutoInitialised Then
    m_AutoInitialised = False
    If AutoNames.IsPresent(m_AutoName) Then Call AutoNames.Remove(m_AutoName)
    If Not AutoParser Is Nothing Then
      Set AutoParser.ParseSettings.ParseParameter(1) = Nothing
      Set AutoParser.ParseSettings.ParseParameter(2) = Nothing
      Set AutoParser.ParseSettings.ParseParameter(3) = Nothing
    End If
    AutoCount = AutoCount - 1
    If AutoCount = 0 Then
      Set AutoNames = Nothing
      Set AutoParser = Nothing
      Call ResetScriptControl
    End If
  End If
  inkill = False
End Sub

Public Sub KillInternal()
  Dim i As Long
  
  On Error Resume Next
  Set m_Report = Nothing
  For i = LBound(m_cols) To UBound(m_cols)
    Set m_cols(i) = Nothing
  Next i
  Call KillGrid
  Set m_RSDefault = Nothing
  m_ReportNonZeroRowsOnly = False
End Sub

Public Sub KillGrid()
  Static inkill As Boolean
  Dim i As Long
  
  On Error Resume Next
  If inkill Then Exit Sub
  inkill = True
  If Not m_AutoGrid_DAO Is Nothing Then
    m_AutoGrid_DAO.Kill
    Set m_AutoGrid_DAO = Nothing
  End If
  If Not m_AutoGrid_RDO Is Nothing Then
    m_AutoGrid_RDO.Kill
    Set m_AutoGrid_RDO = Nothing
  End If
  If Not m_Grid Is Nothing Then
    Call m_Grid.Kill
    Set m_Grid = Nothing
  End If
  inkill = False
End Sub

Private Property Get ILibraryVersion_Name() As String
  ILibraryVersion_Name = "Auto Display library"
End Property

Private Property Get ILibraryVersion_Version() As String
  ILibraryVersion_Version = App.Major & "." & App.Minor & "." & App.Revision
End Property
Public Property Get AlternateBoxedLines() As Boolean
  AlternateBoxedLines = m_AlternateBoxedLines
End Property
Public Property Let AlternateBoxedLines(ByVal NewValue As Boolean)
  m_AlternateBoxedLines = True
End Property

