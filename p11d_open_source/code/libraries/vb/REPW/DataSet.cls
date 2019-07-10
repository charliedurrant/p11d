VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReportDataSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Name As String
Public Parent As ReportDataSet
Public Depth As Long                ' Depth of tree 1 to MAX_DEPTH
Private mCurrentDataSet As Variant  ' fully qualified data set path
Private mCurrentDataSetString As String
Private mLeafNode As Boolean

Public cFields As Collection

' printing only just used to simplify NodeCount, Selection etc
Public NodeUsed As Boolean

' printing only calculated in Prepare
Public FieldsSelected As Collection
Private mFieldOffset As Long

Public MinFieldIndex As Long
Public MaxFieldIndex As Long
Public Children As DataSetCollection

Private Sub Class_Initialize()
  On Error Resume Next
  Set Children = New DataSetCollection
  Set cFields = New Collection
  Set FieldsSelected = New Collection
  Depth = 1
End Sub

Public Property Get IsLeafNode() As Boolean
  IsLeafNode = mLeafNode
End Property

Public Property Get CurrentDataSetString() As String
  CurrentDataSetString = UCase$(mCurrentDataSetString)
End Property

Public Property Get CurrentDataSet() As Variant
  CurrentDataSet = mCurrentDataSet
End Property

Public Property Let CurrentDataSet(NewValue As Variant)
  Dim i As Long

  mCurrentDataSet = NewValue
  mCurrentDataSetString = ""
  For i = LBound(mCurrentDataSet) To UBound(mCurrentDataSet)
    If i = LBound(mCurrentDataSet) Then
      mCurrentDataSetString = CStr(mCurrentDataSet(i))
    Else
      mCurrentDataSetString = mCurrentDataSetString & DATASET_SEP & CStr(mCurrentDataSet(i))
    End If
  Next i
End Property

Private Function GetUDMFieldDetails(udmIF As ITCSUDM) As Variant
  Dim Fields As Variant
  Dim oleerr As OLE_ERRORTYPE
    
  udmIF.CurrentDataSet = CurrentDataSet
  oleerr = udmIF.GetFieldDetails(Fields)
  If oleerr <> S_OK Then Call Err.Raise(ERR_GETFIELDDETAILS, "GetUDMFieldDetails", "Error calling interface implementation of GetFieldDetails")
  GetUDMFieldDetails = Fields
End Function

Public Function GetFieldDetails(udmIF As ITCSUDM) As Boolean
  Dim Fields As Variant, fldBase As FieldDetails, rFld As ReportField
  Dim i As Long, oleerr As OLE_ERRORTYPE
  Dim fielderror As String, errno As Long
  
  On Error GoTo GetFieldDetails_Err
  Fields = GetUDMFieldDetails(udmIF)
  
  If IsEmpty(Fields) Then
    GetFieldDetails = True
    GoTo GetFieldDetails_End
  End If
  
  MinFieldIndex = LBound(Fields)
  MaxFieldIndex = UBound(Fields)
  For i = MinFieldIndex To MaxFieldIndex
    Set fldBase = Fields(i)
    Set rFld = New ReportField
    rFld.Name = fldBase.Name
    rFld.Description = fldBase.Description 'cad p11d
    rFld.DataType = fldBase.DataType
    rFld.FixedWidth = fldBase.Length
    rFld.IndexUDM = i
    #If AbacusReporter Then
     rFld.Tag = fldBase.Tag
    #End If
    Set rFld.DataSet = Me
    rFld.DataSetString = Me.CurrentDataSetString

    
    Me.cFields.Add rFld, rFld.KeyString
    If rFld.DataType <> TYPE_BOOL And rFld.DataType <> TYPE_DATE And rFld.DataType <> TYPE_DOUBLE And rFld.DataType <> TYPE_LONG And rFld.DataType <> TYPE_STR Then
      Call ECASE("Field " & rFld.KeyString & " has an invalid data type")
    End If
    Select Case rFld.DataType
      Case TYPE_LONG, TYPE_DOUBLE
        rFld.Alignment = ALIGN_RIGHT
      Case Else
        rFld.Alignment = ALIGN_LEFT
    End Select
  Next i
  GetFieldDetails = True
  
GetFieldDetails_End:
  Exit Function
  
GetFieldDetails_Err:
  GetFieldDetails = False
  errno = Err.Number: fielderror = ""
  If (errno <> ERR_GETFIELDDETAILS) And (i >= MinFieldIndex) Then
      If IsArray(Fields) Then
        Set fldBase = Fields(i)
        If fldBase Is Nothing Then
          fielderror = "Field: UNKNOWN"
        Else
          fielderror = "Field: " & fldBase.Name
        End If
      Else
        fielderror = "Field: UNKNOWN"
      End If
  End If
  Call ErrorMessage(ERR_ERROR, Err, "GetFieldDetails", "Get Field Details", fielderror & vbCrLf & "Report Interface: " & udmIF.Name & vbCrLf & "Current data set: " & Me.CurrentDataSetString)
  Resume GetFieldDetails_End
  Resume
End Function

Public Sub ReadChildren(udmIF As ITCSUDM)
  Dim DataSets As Variant, Fields As Variant
  Dim CurDataSet As Variant, fldBase As FieldDetails
  Dim rFld As ReportField
  Dim i As Long, j As Long
  Dim dSet As ReportDataSet
  
  On Error GoTo ReadChildren_Err
  
  DataSets = udmIF.QueryAllSubDataSets(CurrentDataSet)
  
  udmIF.MoveFirst
  If Not IsArray(DataSets) Then
    mLeafNode = True
    Exit Sub
  End If
  CurDataSet = udmIF.CurrentDataSet
  ReDim Preserve CurDataSet(LBound(CurDataSet) To UBound(CurDataSet) + 1)
  For i = LBound(DataSets) To UBound(DataSets)
    Set dSet = New ReportDataSet
    dSet.Name = DataSets(i)
    Set dSet.Parent = Me
    dSet.Depth = Me.Depth + 1
    
    CurDataSet(UBound(CurDataSet)) = dSet.Name
    dSet.CurrentDataSet = CurDataSet
    udmIF.CurrentDataSet = CurDataSet
    
    If dSet.GetFieldDetails(udmIF) Then
      Call dSet.ReadChildren(udmIF)
      Me.Children.Add dSet, dSet.CurrentDataSetString
    End If
  Next i

ReadChildren_End:
  Exit Sub
  
ReadChildren_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ReadChildren", "Read data set children", "Unable to read children of dataset " & Me.Name & vbCrLf & "Report Interface: " & udmIF.Name)
  Resume ReadChildren_End
  Resume
End Sub

Public Sub PrepareOutput(AllSelectedFields As Collection, IFace As ITCSUDM, FieldOffset As Long, PrintOrder As Variant, ReportFields As Variant, ByVal MaxCriteria As Long)
  Dim i As Long, j As Long, k As Long, rFld As ReportField, Crit As Criterion
  Dim Fields As Variant, fldBase As FieldDetails
  On Error GoTo PrepareOutput_Err
  
  Call ClearCollection(FieldsSelected)
  j = FieldOffset
  For i = 1 To AllSelectedFields.Count
    Set rFld = AllSelectedFields.Item(i)
    If StrComp(mCurrentDataSetString, rFld.DataSet.CurrentDataSetString, vbTextCompare) = 0 Then
      If InCollection(cFields, rFld.DataSetString & DATASET_SEP & rFld.Name) Then
        Call FieldsSelected.Add(rFld, rFld.Name)
        Set ReportFields(rFld.Order) = rFld
        PrintOrder(j) = rFld.Order
        For k = rFld.Criteria.Count To MaxCriteria
          Call rFld.Criteria.Add(Nothing)
        Next k
        j = j + 1
      End If
    End If
  Next i
  mFieldOffset = FieldOffset
  FieldOffset = mFieldOffset + FieldsSelected.Count
  
  ' set required flag on field I require
  Fields = GetUDMFieldDetails(IFace)
 'RK 22/03/05 Addition to deal with DataSets which have no fields defined i.e. used for display purposes
  If Not IsEmpty(Fields) Then
    For i = MinFieldIndex To MaxFieldIndex
      Set fldBase = Fields(i)
      fldBase.Required = False
    Next i
  End If
  For Each rFld In FieldsSelected
    Set fldBase = Fields(rFld.IndexUDM)
    fldBase.Required = True
  Next rFld
  IFace.CurrentDataSet = mCurrentDataSet
  If Not (IFace.EOF And IFace.BOF) Then IFace.MoveFirst
  
PrepareOutput_End:
  Exit Sub
  
PrepareOutput_Err:
  Call ErrorMessage(ERR_ERROR, Err, "PrepareOutput", "Prepare Report output columns", "Error preparing report columns")
  Resume PrepareOutput_End
  Resume
End Sub

' Get next available line from UDM
' PrintLine = 1 to FieldSelectedCount
'
'   Fills in the appropriate elements of PrintLine with data from the selected
' fields in this data set.
Public Function GetLineEx(PrintLine As Variant, curid As Variant, IFace As ITCSUDM, LastId As Variant, PrintOrder As Variant, ByVal IgnoreUserError As Boolean) As Boolean
  Dim errno As OLE_ERRORTYPE
  Dim CurFields As Variant, CurValues As Variant
  Dim i As Long, j As Long, rFld As ReportField, fCount As Long
  
  On Error GoTo GetLineEx_Err
  Call xSet("GetLineEx")
  GetLineEx = False
  ' blank out current values
  For i = mFieldOffset To (mFieldOffset + FieldsSelected.Count - 1)
    PrintLine(PrintOrder(i)) = Null
  Next i
  
  IFace.CurrentDataSet = GetRestrictedDataSet(LastId) ' UDM does not change current record
  ' Check EOF, Move then recheck
redo_getrecord:
  If IFace.EOF Then GoTo GetLineEx_End
  IFace.MoveNext
  If IFace.EOF Then GoTo GetLineEx_End
    
  errno = IFace.GetValues(CurFields, CurValues, curid)
  If (errno >= E_USERERRORMIN) And (errno <= E_USERERRORMAX) And IgnoreUserError Then GoTo redo_getrecord
  If errno <> S_OK Then GoTo GetLineEx_End
  fCount = mFieldOffset
  For i = 1 To FieldsSelected.Count
    Set rFld = FieldsSelected.Item(i)
    If rFld.IndexUDM = -1 Then Call ECASE("Invalid UDM Index")
    PrintLine(PrintOrder(fCount)) = CurValues(rFld.IndexUDM)
    fCount = fCount + 1
  Next i
  GetLineEx = ((fCount - mFieldOffset) = FieldsSelected.Count)

GetLineEx_End:
  Call xReturn("GetLineEx")
  Exit Function
  
GetLineEx_Err:
  Call ErrorMessage(ERR_ERROR, Err, "GetLine", "Get Next Valid Line", "Error retrieving Data")
  Resume GetLineEx_End
  Resume
End Function

Private Function GetRestrictedDataSet(LastId As Variant) As Variant
  Dim i As Long, j As Long
  Dim CurDataSet As Variant
  
  CurDataSet = mCurrentDataSet
  If IsArray(LastId) Then
    j = LBound(mCurrentDataSet)
    For i = LBound(LastId) To UBound(LastId)
      If Len(LastId(i)) > 0 Then
        CurDataSet(j) = CurDataSet(j) & "=" & LastId(i)
      End If
      j = j + 1
    Next i
  End If
  GetRestrictedDataSet = CurDataSet
End Function


Public Property Get Selected(Optional ByVal ThisDataSetOnly As Boolean = False) As Boolean
  Dim Fld As ReportField, dSet As ReportDataSet
  
  Selected = False
  For Each Fld In cFields
    If Fld.Selected Then
      Selected = True
      Exit For
    End If
  Next Fld
  If (Not Selected) And (Not ThisDataSetOnly) Then
    For Each dSet In Children
      Selected = dSet.Selected
      If Selected Then Exit For
    Next dSet
  End If
End Property

Public Property Get CriteriaSQL() As String
  CriteriaSQL = CriteriaSQLEx(FieldsSelected)
End Property
