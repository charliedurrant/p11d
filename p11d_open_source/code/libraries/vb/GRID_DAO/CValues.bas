Attribute VB_Name = "CalcValues"
Option Explicit
Public Const DELIMITER_FIELD_START As String = "["
Public Const DELIMITER_FIELD_END As String = "]"
Public Const DELIMITER_STATIC_START As String = "<"
Public Const DELIMITER_STATIC_END As String = ">"

Public Function IsSimpleCalc(ByVal CalcType As DERIVED_VALUES_CALC) As Boolean
  IsSimpleCalc = ((CalcType And DERIVED_VALUE_DEFAULT) = DERIVED_VALUE_DEFAULT) Or ((CalcType And DERIVED_VALUE_STATIC) = DERIVED_VALUE_STATIC)
End Function

Public Function DerivedValueType(ByVal CalcValue As String) As DERIVED_VALUES_CALC
  Dim p0 As Long, p1 As Long, FieldName As String
  
  ' substitute in grid field values
  DerivedValueType = DERIVED_VALUE_DEFAULT
  p0 = InStr(1, CalcValue, DELIMITER_STATIC_START)
  If p0 > 0 Then
    p1 = InStr(p0 + 1, CalcValue, DELIMITER_STATIC_END)
    If p1 > 0 And ((p1 - p0 - 1) > 0) Then
      DerivedValueType = DERIVED_VALUE_STATIC
    End If
  End If
  
  p0 = InStr(1, CalcValue, DELIMITER_FIELD_START)
  If p0 > 0 Then
    p1 = InStr(p0 + 1, CalcValue, DELIMITER_FIELD_END)
    If p1 > 0 And ((p1 - p0 - 1) > 0) Then
      If (DerivedValueType And DERIVED_VALUE_STATIC) = DERIVED_VALUE_STATIC Then
        DerivedValueType = DERIVED_VALUE_STATIC
      Else
        DerivedValueType = DERIVED_VALUE_NO_CALC
      End If
      DerivedValueType = DerivedValueType + DERIVED_VALUE_FIELD
    End If
  End If
  
End Function

Public Function GetDerivedPKColumnCaptionsEx(ByVal ac As AutoClass, ByVal AGrid As AutoGrid, ByVal UpdateType As GRIDEDIT_TYPE) As String
  Dim i As Long, pkderivedcolumns As String, flist As String
  Dim CalcValue As String, CalcType As DERIVED_VALUES_CALC
  Dim aCol As AutoCol
  
  On Error GoTo GetDerivedPKColumnCaptionsEx_err
  For i = 1 To ac.Count
    Set aCol = ac.Item(i)
    If aCol.PrimaryKey Then
      If Not aCol.Hide Then
        pkderivedcolumns = pkderivedcolumns & "[" & aCol.GridCaptionClean & "]" & vbCrLf
      Else
        CalcType = DERIVED_VALUE_NO_CALC
        If UpdateType = GRID_EDIT Then
          CalcValue = aCol.OnUpdateCalcValue
          CalcType = aCol.OnUpdateCalcValueType
        ElseIf UpdateType = GRID_ADDNEW Then
          CalcValue = aCol.OnAddNewCalcValue
          CalcType = aCol.OnAddNewCalcValueType
        End If
        If Not ((CalcType And DERIVED_VALUE_NO_CALC) = DERIVED_VALUE_NO_CALC) Or ((CalcType And DERIVED_VALUE_DEFAULT) = DERIVED_VALUE_DEFAULT) Then
          Call ReplaceFieldNames(AGrid, CalcValue, flist, vbCrLf, False)
          If Len(flist) > 0 Then pkderivedcolumns = pkderivedcolumns & flist
        End If
      End If
    End If
  Next i
GetDerivedPKColumnCaptionsEx_end:
  GetDerivedPKColumnCaptionsEx = pkderivedcolumns
  Exit Function
  
GetDerivedPKColumnCaptionsEx_err:
  pkderivedcolumns = "Error obtaining primary key columns" & vbCrLf & Err.Description
  Resume GetDerivedPKColumnCaptionsEx_end
End Function

Public Function ReplaceFieldNames(ByVal AGrid As AutoGrid, ByVal CalcValueFormulae As String, Optional FieldNameList As String, Optional ByVal FieldSeparator As String = ", ", Optional ByVal TrimSeparator As Boolean = True) As String
  Dim aColSrc As AutoCol
  Dim p0 As Long, p1 As Long, FieldName As String, CalcLen As Long
  Dim l As Long
  
  On Error GoTo ReplaceFieldNames_err
  FieldNameList = ""
  p0 = 1
  Do
    p0 = InStr(p0, CalcValueFormulae, DELIMITER_FIELD_START)
    If p0 > 0 Then
      p1 = InStr(p0 + 1, CalcValueFormulae, DELIMITER_FIELD_END)
      If p1 > 0 Then
        FieldName = Mid$(CalcValueFormulae, p0 + 1, p1 - p0 - 1)
        Set aColSrc = AGrid.GetAColByKey(FieldName)
        If Not aColSrc Is Nothing Then
          FieldNameList = FieldNameList & "[" & aColSrc.GridCaptionClean & "]" & FieldSeparator
          CalcValueFormulae = Left$(CalcValueFormulae, p0) & aColSrc.GridCaptionClean & Mid$(CalcValueFormulae, p1)
        End If
        p0 = p0 + Len(aColSrc.GridCaptionClean) + 2
      End If
    End If
  Loop Until (p0 = 0) Or (p1 = 0)

ReplaceFieldNames_end:
  If (Len(FieldNameList) > 0) And TrimSeparator Then FieldNameList = Left$(FieldNameList, Len(FieldNameList) - Len(FieldSeparator))
  ReplaceFieldNames = CalcValueFormulae
  Exit Function
  
ReplaceFieldNames_err:
  Resume ReplaceFieldNames_end
End Function

Private Function GetFieldValueRS(ByVal rs As Recordset, ByVal FieldName As String, ByVal vbmk As Variant, ByRef CalcValue, ByVal p0 As Long, ByVal p1 As Long) As Variant
  Dim fld As field
  
  On Error GoTo GetFieldValueRS_err
  Set fld = rs.Fields(FieldName)
  If Not IsNull(vbmk) Then rs.Bookmark = vbmk
  If IsNull(fld.Value) Then
    CalcValue = Left$(CalcValue, p0 - 1) & "(Null)" & Mid$(CalcValue, p1 + 1)
    Err.Raise ERR_GETCALCVALUE, "GetFieldValueRS", "Fieldname '" & FieldName & "' does not have a value that can be evaluated."
  End If
  GetFieldValueRS = fld.Value
  Exit Function
  
GetFieldValueRS_err:
  Err.Raise Err.Number, ErrorSource(Err, "GetFieldValueRS"), Err.Description
  Resume
End Function


Private Function GetFieldValue(ByVal AGrid As AutoGrid, ByVal FieldName As String, ByVal vbmk As Variant, ByRef CalcValue, ByVal p0 As Long, ByVal p1 As Long) As Variant
  Dim aColSrc As AutoCol
  Dim ColSet As TrueDBGrid60.Column
  
  On Error GoTo GetFieldValue_err
  Set aColSrc = AGrid.GetAColByKey(FieldName)
  If aColSrc Is Nothing Then Err.Raise ERR_GETCALCVALUE, "GetFieldValue", "Unable to lookup fieldname '" & FieldName & "' in the Auto Columns for the Grid."
  If IsNull(vbmk) Then
    If IsEmpty(aColSrc.CellValue) Then
      CalcValue = Left$(CalcValue, p0 - 1) & "(Empty)" & Mid$(CalcValue, p1 + 1)
      Err.Raise ERR_GETCALCVALUE, "GetFieldValue", "Fieldname '" & FieldName & "' does not have a value that can be evaluated."
    End If
    GetFieldValue = aColSrc.CellValue
  Else
    If aColSrc.GridColumn < 0 Then Err.Raise ERR_GETCALCVALUE, "GetFieldValue", "Unable to find fieldname '" & FieldName & "' in the Grid Columns."
    If aColSrc.UnboundColumn Then Err.Raise ERR_GETCALCVALUE, "GetFieldValue", "Unable to use the unbound column '" & FieldName & "' in the calculation of other columns."
    Set ColSet = AGrid.TDBGrid.Columns.Item(aColSrc.GridColumn)
    GetFieldValue = GetTypedValueNull(ColSet.CellValue(vbmk), aColSrc.dbDataType)
  End If
  Exit Function
  
GetFieldValue_err:
  Err.Raise Err.Number, ErrorSource(Err, "GetFieldValue"), Err.Description
  Resume
End Function

Public Function GetCalculatedValue(ByVal AGrid As AutoGrid, ByVal rs As Recordset, ByVal aCol As AutoCol, ByVal CalcValue As String, ByVal CalcType As DERIVED_VALUES_CALC, Optional ByVal vbmk As Variant = Null, Optional ByVal SuppressErrors As Boolean = False) As Variant
  Dim p0 As Long, p1 As Long, FieldName As String, CalcLen As Long
  Dim Col As Long, IsCalculation As Boolean
  Dim doEval As Boolean, Value As Variant
  Dim errstring As String, CalcValueFormulae As String, CalcValueList As String
  
  On Error GoTo GetCalculatedValue_err
  ' substitute in grid field values
  If Not aCol.NoCalc Then
    If (aCol.dbDataType = TYPE_STR) And (InStr(CalcValue, "&") > 0) Then IsCalculation = True
    If IsNumberField(aCol.dbDataType) And (InStrAny(CalcValue, "+-/*") > 0) Then IsCalculation = True
  End If
  CalcLen = Len(CalcValue)
  CalcValueFormulae = CalcValue
  If (CalcType And DERIVED_VALUE_FIELD) = DERIVED_VALUE_FIELD Then
    p0 = 1
    Do
      p0 = InStr(p0, CalcValue, DELIMITER_FIELD_START)
      If p0 > 0 Then
        p1 = InStr(p0 + 1, CalcValue, DELIMITER_FIELD_END)
        If p1 > 0 Then
          FieldName = Mid$(CalcValue, p0 + 1, p1 - p0 - 1)
          If Not AGrid Is Nothing Then
            Value = GetFieldValue(AGrid, FieldName, vbmk, CalcValue, p0, p1)
          ElseIf Not rs Is Nothing Then
            Value = GetFieldValueRS(rs, FieldName, vbmk, CalcValue, p0, p1)
          Else
            Err.Raise ERR_GETCALCVALUE, "GetCalculatedValue ", "No grid or recordset specified for calculation"
          End If
          If (p0 = 1) And (p1 = CalcLen) Then
            GetCalculatedValue = Value
            Exit Function
          End If
          If VarType(Value) = vbString Then Value = """" & Value & """"
          Value = CStr(Value)
          CalcValue = Left$(CalcValue, p0 - 1) & Value & Mid$(CalcValue, p1 + 1)
          p0 = p0 + Len(Value)
        End If
      End If
    Loop Until (p0 = 0) Or (p1 = 0)
  End If
  
  If (CalcType And DERIVED_VALUE_STATIC) = DERIVED_VALUE_STATIC Then
    p0 = 1
    Do
      p0 = InStr(p0, CalcValue, DELIMITER_STATIC_START)
      If p0 > 0 Then
        p1 = InStr(p0 + 1, CalcValue, DELIMITER_STATIC_END)
        If p1 > 0 Then
          FieldName = Mid$(CalcValue, p0 + 1, p1 - p0 - 1)
          Value = GetTypedValue(AutoEvaluate(FieldName), aCol.dbDataType)
          If (p0 = 1) And (p1 = CalcLen) Then
            GetCalculatedValue = Value
            Exit Function
          End If
          If VarType(Value) = vbString Then Value = """" & Value & """"
          Value = CStr(Value)
          CalcValue = Left$(CalcValue, p0 - 1) & Value & Mid$(CalcValue, p1 + 1)
          p0 = p0 + Len(Value)
        End If
      End If
    Loop Until (p0 = 0) Or (p1 = 0)
  End If
  If IsCalculation Then
    GetCalculatedValue = GetTypedValue(AutoEvaluate(CalcValue), aCol.dbDataType)
  Else
    GetCalculatedValue = GetTypedValue(CalcValue, aCol.dbDataType)
  End If
  
GetCalculatedValue_end:
  Exit Function
  
GetCalculatedValue_err:
  GetCalculatedValue = Empty
  If SuppressErrors Then Resume GetCalculatedValue_end
  Call ErrorMessagePush(Err)
  If aCol.Hide Then
    errstring = "Error calculating value for hidden column '"
  Else
    errstring = "Error calculating value for column '"
  End If
  CalcValueFormulae = ReplaceFieldNames(AGrid, StrDupChar(CalcValueFormulae, "&"), CalcValueList)
  Call ErrorMessagePopErr(Err)
  errstring = errstring & aCol.GridCaptionClean & "'" & vbCrLf & _
              Err.Description & vbCrLf & _
              "Please ensure that all the columns " & CalcValueList & " contain values" & vbCrLf & vbCrLf & _
              "Formula: " & CalcValueFormulae & vbCrLf & _
              "Value:    " & ReplaceFieldNames(AGrid, StrDupChar(CalcValue, "&"))
  Err.Raise ERR_GRIDCALCFAIL, ErrorSource(Err, "GetCalculatedValue"), errstring
  'Call ErrorMessage(ERR_ERROR, Err, "GetCalculatedValue", "Unable to derive calculated value", "Error calculating value for column '" & aCol.GridCaptionClean & "'" & vbCrLf & vbCrLf & "Formula: " & StrDupChar(CalcValue, "&"))
  Resume GetCalculatedValue_end
  Resume
End Function


