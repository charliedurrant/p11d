VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestUDM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mFields As Long, mCurField As Long
Private mRows As Long, mCurRow As Long
Private UDMLines() As Variant
Private FieldDetails() As FieldDetails
Private FieldNames() As String
Implements ITCSUDM
Private Const NUM_ROWS = 1500
Private Const NUM_COLUMNS = 10

Private Function Initialise(ByVal NumFields As Long, ByVal NumRows As Long)
  Dim i As Long, j As Long
  Dim aRow() As String
  
  ReDim UDMLines(1 To NumRows)
  ReDim aRow(1 To NumFields)
  For i = 1 To NumRows
    'ReDim aRow(1 To NumFields)
    For j = 1 To NumFields
      aRow(j) = "Row " & CStr(i) & " Field " & CStr(j)
    Next j
    UDMLines(i) = aRow
  Next i
  ReDim FieldDetails(1 To NumFields)
  ReDim FieldNames(1 To NumFields)
  For j = 1 To NumFields
    Set FieldDetails(j) = New FieldDetails
    FieldNames(j) = "Field " & CStr(j)
    FieldDetails(j).Name = "Field " & CStr(j)
    FieldDetails(j).DataType = TYPE_STR
  Next j
  mRows = NumRows
  mFields = NumFields
End Function

Private Property Get ITCSUDM_Name() As String
  ITCSUDM_Name = "Test UDM"
End Property

Private Property Get ITCSUDM_Description() As String
  ITCSUDM_Description = "This is a test UDM"
End Property

Private Property Get ITCSUDM_EOF() As Boolean
  ITCSUDM_EOF = (mCurRow = mRows)
End Property

Private Property Get ITCSUDM_BOF() As Boolean
  ITCSUDM_BOF = (mCurRow = 0)
End Property

Private Sub ITCSUDM_CloseSession()
  mCurField = 1
  mCurRow = 0
End Sub

Private Property Let ITCSUDM_CurrentDataSet(RHS As Variant)
  If RHS(1) <> ITCSUDM_Name Then Call ECASE("Error in Let CurrentDataset")
End Property

Private Property Get ITCSUDM_CurrentDataSet() As Variant
  Dim v() As Variant
  ReDim v(1 To 1)
  v(1) = ITCSUDM_Name
  ITCSUDM_CurrentDataSet = v
End Property

Private Function ITCSUDM_GetFieldDetails(FieldDetailsArray As Variant) As TCSBASE.OLE_ERRORTYPE
  FieldDetailsArray = FieldDetails
  ITCSUDM_GetFieldDetails = S_OK
End Function

Private Function ITCSUDM_GetValues(SourceFields As Variant, SourceValues As Variant, Optional KeyID As Variant) As TCSBASE.OLE_ERRORTYPE
  SourceFields = FieldNames
  SourceValues = UDMLines(mCurRow)
  ITCSUDM_GetValues = S_OK
End Function

Private Sub ITCSUDM_MoveFirst()
  mCurRow = 0
End Sub

Private Sub ITCSUDM_MoveNext()
  mCurRow = mCurRow + 1
End Sub

Private Function ITCSUDM_OpenSession(FileName As String, Optional ByVal DestSelectOnly As Boolean = False) As Boolean
  mCurField = 1
  mCurRow = 0
  Call Initialise(NUM_COLUMNS, NUM_ROWS)
End Function

Private Function ITCSUDM_QueryAllSubDataSets(DataSetArray As Variant) As Variant
  ' do nothing
End Function

Private Function ITCSUDM_QueryAllTopLevelDataSets() As Variant
  Dim v() As Variant
  ReDim v(1 To 1)
  v(1) = ITCSUDM_Name
  ITCSUDM_QueryAllTopLevelDataSets = v
End Function

Private Function ITCSUDM_IsGetDetailsAvailable() As Boolean
  ITCSUDM_IsGetDetailsAvailable = True
End Function

Private Function ITCSUDM_GetErrorDescription(ByVal ErrorCode As TCSBASE.OLE_ERRORTYPE) As String
  ITCSUDM_GetErrorDescription = "GetErrorDescription Not Implemented"
End Function

Private Property Get ITCSUDM_RecordCount() As Long
  ITCSUDM_RecordCount = mRows
End Property

Private Property Get ITCSUDM_Version() As String
  ITCSUDM_Version = "1.0"
End Property

Private Function ITCSUDM_IsSetDetailsAvailable() As Boolean
  ITCSUDM_IsSetDetailsAvailable = False
End Function

Private Function ITCSUDM_SetValues(ByVal DestFields As Variant, ByVal DestValues As Variant, Optional KeyID As Variant) As TCSBASE.OLE_ERRORTYPE

End Function

Private Function ITCSUDM_AppendValues(ByVal DestFields As Variant, ByVal DestValues As Variant) As TCSBASE.OLE_ERRORTYPE

End Function

Private Property Set ITCSUDM_Notify(RHS As TCSBASE.IBaseNotify)

End Property

Private Property Get ITCSUDM_Notify() As TCSBASE.IBaseNotify

End Property


