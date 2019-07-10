VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AutoCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum SORT_TYPE
  SORT_NONE = 0
  SORT_ASCENDING
  SORT_DESCENDING
End Enum

Public Enum SUM_TYPE
  TYPE_NOSUM = 0
  TYPE_SUM
  TYPE_MEAN
End Enum

Public Enum DROPDOWN_TYPE
  DROPDOWN_NONE = 0
  DROPDOWN_QUERY = 1
  DROPDOWN_BOUND = 2 + DROPDOWN_QUERY ' total limit of these two is 12 (see ACTRL)
  DROPDOWN_STATIC = 4
  DROPDOWN_DYNAMIC = 8                ' requery required on dropdown
  DROPDOWN_LIST = 16
  DROPDOWN_COMBO = 32
End Enum

Public Enum DERIVED_VALUES_CALC
  DERIVED_VALUE_NO_CALC = 0
  DERIVED_VALUE_DEFAULT = 1
  DERIVED_VALUE_STATIC = 2
  DERIVED_VALUE_FIELD = 4
End Enum

Public Parent As AutoClass
Public DataField As String             ' Key For Columns - DataField in recordset

Public PrimaryKey As Boolean     ' KEYCOLUMN
Private mAllowNullValues As Boolean
Public SourceField As String     ' BASECOLUMNNAME
Public SourceTable  As String    ' BASETABLENAME

Public GridColumn As Integer           ' Column in Grid for this AutoCol (-1 if not present)

Private mGridCaption As String
Private mGridCaptionClean As String

Public PrintCaption As String
Public PrintCaption_Clean As String
Public PrintCaption_Returns As Long

Public ToolTip As String

Public FirstHeader As Boolean

Public dbDataType As DATABASE_FIELD_TYPES
Private mIsCheckBox As Boolean
Public HasButton As Boolean
Public ButtonPictureIndex As Long
Public TrimField As Boolean
Public SplitColumn As Boolean

Public Hide As Boolean          'H
Public ForceCreateColumn As Boolean  ' Forces column to exist in the grid - it will then be copied by CopyCurrentRow etc..
Public NoCopy  As Boolean             ' Column can be copied from grid in Copy/CopyCurrentRow

Public NoPrint As Boolean        'Is this a printing column
Public DerivedNoPrint As Boolean ' .. calculated value derived from NoPrint + width
Public Wrap As Boolean

Public CollapseLike As Variant

' DropDowns in Grid
Public DropType As DROPDOWN_TYPE
Private mQueryName As String
Public QueryStyle As Integer    'v#{n}sqlname{p1}{p2}{n}# n=0 List n=1 combo - Required!!
Public QueryCombo As Long
Public BoundDataField As String 'See ParseQuery
Public TranslateDropDown As Long
Public MaxDropDownItems As Long
Private Parameters() As String  'Array of Strings containing all parameters
Private m_ParameterCount As Long
Private m_SQLParameterCount As Long
Private m_RefreshRequired As Boolean

Private mNoEdit As Boolean        'E
Private mSysNoEdit As Boolean
Private mLockEdit As Boolean
Public NoAddNew As Boolean
 
Public CopyFromField As String      'N#fldName#
Public ClipboardColumn As Boolean

Public OnChangeEvent As Boolean
Public NoCalc As Boolean             ' do not evaluate the contents of this column always treat as literal
Public UnboundColumn As Boolean
Public OnAddNewDefault As Variant
Public OnAddNewDefaultType As DERIVED_VALUES_CALC
Public OnAddNewCalcValue As Variant
Public OnAddNewCalcValueType As DERIVED_VALUES_CALC
Public OnUpdateCalcValue As Variant
Public OnUpdateCalcValueType As DERIVED_VALUES_CALC

Public NewColumn As Boolean     'C
Public TabTitle As String       't#tabtitle
Public ColWidth As Long

' Format Control
Public FormatString As String
Private m_Alignment As ALIGNMENT_TYPE
Public SortType As SORT_TYPE
Private m_Format As String
Private nDecimalPlaces As Long
Private m_AlignSet As Boolean
Private m_FormatSet As Boolean
Private m_Width As Long         ' restrict to this many chars
Public CustomFormat As Boolean

' Printer vars
Public CaptionFormat As String
Public DataFormat As String
Public TrimType As String
Public MinWidth As Single
Public MaxWidth As Single
Public PrintPageIndex As Long
Public pOffset As Single
Public pOffsetUser As Single

Public SumType As SUM_TYPE
Public bGroup As Boolean
Public bGroupHeader As Boolean
Public bGroupTotal As Boolean
Public bGroupCount As Boolean
Public bGroupPageBreak As Boolean
Public Cumulative As Boolean

Public ReportFormat As String
Public SumLevel As Long          ' used for nested subtotals 0 default
Public RebasedSumLevel As Long   ' used for nested subtotals 0 default - do not set
Public DoGroup As Boolean        ' used for nested subtotals

Public SumLast As Double
Private mSumGroup() As Double
Private mSumGroupCount() As Long
Public SumTotal As Double
Public LastGroupValue As Variant  ' used for printing subtotal captions
Public LastValue As Variant

'Validation elements
Public FastKey As String
Public FastKeyPress As String
Public CellValue As Variant
Public OnNullValue As Variant
Public MaxValidValue As Variant
Public MinValidValue As Variant

'Export help
Public FixedWidth As Long
Public FW_PadLeft As Boolean ' left for text, right for numbers
'FIX CAD1
Public BackColor As Long
Public ForeColor As Long

Private Sub Class_Initialize()
  Call ClearFormat
End Sub

Public Sub ClearFormat()
  On Error Resume Next
  SumLevel = 0
  DoGroup = False
  MinWidth = 0
  PrimaryKey = False
  SourceField = ""
  SourceTable = ""
  pOffsetUser = -1
  m_FormatSet = False
  CustomFormat = False
  m_Alignment = dbgLeft
  m_AlignSet = False
  TrimField = True
  PrintCaption_Returns = 0
  FixedWidth = 0
  FW_PadLeft = True
  mIsCheckBox = False
  TrimType = "TRIM"
  Hide = False
  ForceCreateColumn = False
  SplitColumn = False
  HasButton = False
  ButtonPictureIndex = BUTTON_BASE_BMP
  NoCopy = False
  NoPrint = False
  DerivedNoPrint = False
  m_Format = ""
  nDecimalPlaces = -1
  m_Width = 0
  NoCalc = False
  OnUpdateCalcValue = Empty
  OnAddNewCalcValue = Empty
  OnAddNewDefault = Empty
  OnUpdateCalcValueType = DERIVED_VALUE_NO_CALC
  OnAddNewDefaultType = DERIVED_VALUE_NO_CALC
  OnAddNewCalcValueType = DERIVED_VALUE_NO_CALC
  MaxValidValue = Empty
  MinValidValue = Empty
  DataField = DataField
  GridCaption = DataField
  PrintCaption = DataField
  ToolTip = ""
  SortType = SORT_NONE
  CollapseLike = Empty
  Cumulative = False
  OnChangeEvent = False
  OnNullValue = Null
  
  'Query info
  DropType = DROPDOWN_NONE
  mQueryName = ""
  QueryCombo = 0
  BoundDataField = ""
  MaxDropDownItems = -1
  m_ParameterCount = 0
  m_SQLParameterCount = 0
  m_RefreshRequired = False
  TranslateDropDown = -2
  BackColor = vbWhite
  ForeColor = vbBlack
End Sub

Private Sub SQLParseParameters()
  Dim p0 As Long, p1 As Long
  Dim sParamName As String
  
  On Error GoTo SQLParseParameters_err
  p0 = 1
  m_SQLParameterCount = 0
  Do
    p1 = InStr(p0, mQueryName, BEGIN_PARAM, vbBinaryCompare)
    If p1 > 0 Then
      p0 = p1 + Len(BEGIN_PARAM)
      p1 = InStr(p0, mQueryName, END_PARAM, vbBinaryCompare)
      If p1 = 0 Then Err.Raise ERR_INVALIDSQL, "SQLParseParameters", "Invalid SQL string specified, could not determine parameter start/end" & vbCrLf & mQueryName
      If p0 >= p1 Then Err.Raise ERR_INVALIDSQL, "SQLParseParameters", "Invalid SQL string specified, empty parameter found" & vbCrLf & mQueryName
      sParamName = Mid$(mQueryName, p0, p1 - p0)
      
      m_SQLParameterCount = m_SQLParameterCount + 1
      ReDim Preserve Parameters(1 To m_SQLParameterCount)
      Parameters(m_SQLParameterCount) = sParamName
      p0 = p1 + Len(END_PARAM)
    End If
  Loop Until p1 = 0
SQLParseParameters_end:
  Exit Sub
  
SQLParseParameters_err:
  Err.Raise Err.Number, ErrorSource(Err, "SQLParseParameters"), Err.Description
End Sub
Public Property Get QueryName() As String
  QueryName = mQueryName
End Property

Public Property Let QueryName(ByVal NewValue As String)
  mQueryName = NewValue
  Call SQLParseParameters
End Property


Public Sub AddParam(ByVal NewValue As String)
  If m_SQLParameterCount > 0 Then Err.Raise ERR_AUTOPARSE, "AddParam", "Unable to add a parameter where the sql contains embedded parameters" & vbCrLf & mQueryName
  m_ParameterCount = m_ParameterCount + 1
  ReDim Preserve Parameters(1 To m_ParameterCount) As String
  Parameters(m_ParameterCount) = NewValue
End Sub

Public Property Let RefreshRequired(ByVal NewValue As Boolean)
  m_RefreshRequired = NewValue
End Property

Public Property Get RefreshRequired() As Boolean
  RefreshRequired = m_RefreshRequired Or ((DropType And DROPDOWN_DYNAMIC) = DROPDOWN_DYNAMIC)
End Property

Public Function GetParam(ByVal Index As Long) As String
  If Index < 1 Then Err.Raise ERR_AUTOCOL, "GetParam", "Parameter Index must be > 0"
  If m_SQLParameterCount > 0 Then
    If Index > m_SQLParameterCount Then Err.Raise ERR_AUTOCOL, "GetParam", "SQL Parameter Index must be <= " & m_SQLParameterCount
  Else
    If Index > m_ParameterCount Then Err.Raise ERR_AUTOCOL, "GetParam", "Parameter Index must be <= " & m_ParameterCount
  End If
  GetParam = Parameters(Index)
End Function

Public Property Get IsSQLString() As Boolean
  IsSQLString = (m_SQLParameterCount > 0)
End Property

Public Property Get NoParams() As Long
  If m_SQLParameterCount > 0 Then
    NoParams = m_SQLParameterCount
  Else
    NoParams = m_ParameterCount
  End If
End Property

Public Sub RedimSumLevels(ByVal MaxLevel As Long)
  ReDim mSumGroupCount(0 To MaxLevel) As Long
  ReDim mSumGroup(0 To MaxLevel) As Double
End Sub

Public Property Get SumGroupCount(ByVal Index As Long) As Long
  SumGroupCount = mSumGroupCount(Index)
End Property

Public Property Let SumGroupCount(ByVal Index As Long, ByVal NewValue As Long)
  mSumGroupCount(Index) = NewValue
End Property

Public Property Let LockEdit(ByVal NewValue As Boolean)
  mLockEdit = NewValue
End Property

Public Property Get NoEdit() As Boolean
  NoEdit = mNoEdit Or mSysNoEdit Or mLockEdit
End Property

Public Property Let NoEdit(ByVal NewValue As Boolean)
  mNoEdit = NewValue
End Property

Friend Property Let SysNoEdit(ByVal NewValue As Boolean)
  mSysNoEdit = NewValue
End Property

Public Property Get IsCheckBox() As Boolean
  IsCheckBox = mIsCheckBox Or ((dbDataType = TYPE_BOOL) And (Parent.BooleanCheckBox) And (Not CustomFormat))
End Property

Public Property Let IsCheckBox(ByVal NewValue As Boolean)
  mIsCheckBox = NewValue
End Property

Public Property Get SumGroup(ByVal Index As Long) As Double
  SumGroup = mSumGroup(Index)
End Property

Public Property Let SumGroup(ByVal Index As Long, ByVal NewValue As Double)
  mSumGroup(Index) = NewValue
End Property

Public Property Get GridCaption() As String
  GridCaption = mGridCaption
End Property

Public Property Get GridCaptionClean() As String
  GridCaptionClean = mGridCaptionClean
End Property

Public Property Let GridCaption(ByVal NewValue As String)
  mGridCaption = NewValue
  mGridCaptionClean = ReplaceChar(NewValue, vbCrLf, " ")
End Property

Public Property Get Alignment() As ALIGNMENT_TYPE
  If Not m_AlignSet Then
    If IsNumberField(dbDataType) Then
      m_Alignment = dbgRight
    End If
    m_AlignSet = True
  End If
  Alignment = m_Alignment
End Property

Public Property Let Alignment(ByVal NewValue As ALIGNMENT_TYPE)
  m_Alignment = NewValue
  m_AlignSet = True
End Property

Public Property Let Format(ByVal NewValue As String)
  m_FormatSet = True
  nDecimalPlaces = -1
  If Me.IsCheckBox Then
    m_Format = "" ' ignore format
  Else
    m_Format = NewValue
  End If
End Property

Public Property Let Width(ByVal NewValue As Long)
  m_Width = NewValue
End Property

Public Property Get Format() As String
  If Not m_FormatSet Then
    If dbDataType = TYPE_DOUBLE Then
      m_Format = Parent.DefaultFloatFormat
    ElseIf dbDataType = TYPE_LONG Then
      m_Format = Parent.DefaultIntegerFormat
    ElseIf dbDataType = TYPE_DATE Then
      m_Format = Parent.DefaultDateFormat
    ElseIf dbDataType = TYPE_BOOL Then
      If Me.IsCheckBox Then
        m_Format = ""
      Else
        m_Format = Parent.DefaultBooleanFormat
      End If
    End If
    m_FormatSet = True
  End If
  If nDecimalPlaces < 0 Then
    nDecimalPlaces = 0
    If IsNumberField(dbDataType) Then nDecimalPlaces = GetDecimalPrecision(m_Format)
  End If
  Format = m_Format
End Property

Public Function GetFormattedValue(v As Variant) As String
  Dim bVal As Boolean, p As Long
  
  If Me.CustomFormat Then
    If Me.dbDataType = TYPE_LONG Then
      p = InStr(1, Me.Format, "/")
      GetFormattedValue = IsNullEx(v, False)
      If IsNumeric(GetFormattedValue) Then bVal = CLng(GetFormattedValue)
      If CBoolean(bVal) Then
        GetFormattedValue = "Unknown true"
        If p > 0 Then GetFormattedValue = Mid$(Me.Format, 1, p - 1)
      Else
        GetFormattedValue = "Unknown false"
        If p > 0 Then GetFormattedValue = Mid$(Me.Format, p + 1)
      End If
    Else
      GetFormattedValue = "Unknown Custom Format"
    End If
  Else
    If Len(Me.Format) = 0 Then
      If m_Width <= 0 Then
        GetFormattedValue = v
      Else
        GetFormattedValue = Left$(v, m_Width)
      End If
    Else
      If m_Width <= 0 Then
        GetFormattedValue = VBA.Format$(v, Me.Format)
      Else
        GetFormattedValue = VBA.Format$(Left$(v, m_Width), Me.Format)
      End If
    End If
    If (dbDataType = TYPE_DATE) And (v = UNDATED) Then GetFormattedValue = ""
  End If
End Function

Public Function GetValue(v As Variant) As Double
  If Not IsNumberField(dbDataType) Then Err.Raise ERR_GETVALUE, "GetValue", "Auto Column " & Me.PrintCaption_Clean & " is not a numeric column"
  If Not IsNumeric(IsNullEx(v, "")) Then
    GetValue = 0#
    Exit Function
  End If
  If dbDataType = TYPE_LONG Then
    GetValue = CDbl(v)
  Else
    GetValue = Round(v, nDecimalPlaces)   ' using VB Round function to match with GetFormattedValue
  End If
End Function

