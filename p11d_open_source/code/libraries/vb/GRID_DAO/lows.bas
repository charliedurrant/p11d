Attribute VB_Name = "lows"
Option Explicit
Public Enum FILTER_DELIMITER
  FILTER_START
  FILTER_END
End Enum
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const WM_USER As Long = &H400

Public Const VK_LBUTTON As Long = &H1
Public Const VK_RBUTTON As Long = &H2

Public Const SM_CXHSCROLL As Long = 21
Public Const SM_CYHSCROLL As Long = 3

Public Function GetAColByGridIndexEx(ByVal GridColumn As Integer, aCol As AutoCol, ByVal ac As AutoClass) As Boolean
  
  Dim i As Long
  
  If GridColumn >= 0 Then
    For i = 1 To ac.Count
      Set aCol = ac.Item(i)
      If aCol.GridColumn = GridColumn Then
        GetAColByGridIndexEx = True
        Exit Function
      End If
    Next i
    Set aCol = Nothing
  End If
End Function

Public Function LikeFilterDelimit(ByVal FilterType As FILTER_DELIMITER, ByVal ForceVBExpressions As Boolean) As String
  If FilterType = FILTER_END Then
    If ForceVBExpressions Then
      LikeFilterDelimit = "*"
    ElseIf DatabaseTarget = DB_TARGET_JET Then
      LikeFilterDelimit = "*'"
    Else
      LikeFilterDelimit = "%'"
    End If
  Else
    If ForceVBExpressions Then
      LikeFilterDelimit = "LIKE *"
    ElseIf DatabaseTarget = DB_TARGET_JET Then
      LikeFilterDelimit = "LIKE '*"
    Else
      LikeFilterDelimit = "LIKE '%"
    End If
  End If
End Function

Public Function IsTrimField(ac As AutoCol) As Boolean
  IsTrimField = (ac.dbDataType = TYPE_STR) And ac.TrimField
End Function

Private Function EvaluateFindExpression(ByVal Value As Variant, ByVal Expression As String, ByVal dbType As DATABASE_FIELD_TYPES) As Boolean
  Dim v1 As Variant
  Dim IsLike As String, p As Long
  
  IsLike = (InStr(1, Expression, "like", vbTextCompare) = 1)
  If IsLike Then
    Expression = Mid$(Expression, Len("Like") + 2)
    EvaluateFindExpression = UCase$(Value) Like Expression
  Else
    ' Value based
    If (InStr(1, Expression, "=", vbTextCompare) = 1) Then Expression = Mid$(Expression, 2)
    v1 = GetSQLValue(Value, dbType)
    EvaluateFindExpression = (v1 = Expression)
  End If
End Function


Public Function FindRecordEx(ByVal AGrid As AutoGrid, ByVal rs As Recordset, ByVal rsRDO As RDOResultset, FindString As String, Optional ByVal FindType As FIND_TYPES = FT_FINDFIRST, Optional ByVal dbType As DATABASE_FIELD_TYPES = TYPE_STR) As Boolean
  Dim FieldName As String
  Dim p As Long, Value As Variant
  Dim vbmk As Variant
  
  On Error GoTo FindEecordEx_Err
  If Not rs Is Nothing Then
    vbmk = rs.Bookmark
    Select Case FindType
      Case FT_FINDFIRST
        Call rs.FindFirst(FindString)
      Case FT_FINDNEXT
        Call rs.FindNext(FindString)
      Case FT_FINDPREVIOUS
        Call rs.FindPrevious(FindString)
      Case FT_FINDLAST
        Call rs.FindLast(FindString)
      Case Else
        Call ECASE("FindRecordEx")
    End Select
    FindRecordEx = Not rs.NoMatch
  End If
  If Not rsRDO Is Nothing Then
    Call AGrid.TDBGrid.Close(False)
    vbmk = rsRDO.Bookmark
    If FindType = FT_FINDFIRST Then rsRDO.MoveFirst
    If FindType = FT_FINDLAST Then rsRDO.MoveLast
    If (FindType = FT_FINDNEXT) And Not rsRDO.EOF Then rsRDO.MoveNext
    If (FindType = FT_FINDPREVIOUS) And Not rsRDO.BOF Then rsRDO.MovePrevious
    p = InStr(FindString, " ")
    If p = 0 Then p = InStr(FindString, "=")
    If p = 0 Then Err.Raise ERR_FINDRECORD, "FindRecord", "Unable to Parse FindString " & FindString
    FieldName = Left$(FindString, p - 1)
    FindString = Mid$(FindString, p + 1)
    p = InStr(FieldName, ".")
    If p > 0 Then FieldName = Mid$(FieldName, p + 1)
    FindString = Trim$(UCase$(FindString))
    Do
      Value = rsRDO.rdoColumns(FieldName)
      If Not IsNull(Value) Then
        If EvaluateFindExpression(Value, FindString, dbType) Then
          FindRecordEx = True
          Exit Do
        End If
      End If
      If (FindType = FT_FINDFIRST) Or (FindType = FT_FINDNEXT) Then Call rsRDO.MoveNext
      If (FindType = FT_FINDLAST) Or (FindType = FT_FINDPREVIOUS) Then Call rsRDO.MovePrevious
    Loop Until rsRDO.EOF Or rsRDO.BOF
  End If
    
  
FindEecordEx_End:
  If (Not FindRecordEx) And (Len(vbmk) > 0) Then
    If Not rs Is Nothing Then rs.Bookmark = vbmk
    If Not rsRDO Is Nothing Then rsRDO.Bookmark = vbmk
    Beep
  End If
  If Not rsRDO Is Nothing Then Call AGrid.TDBGrid.ReOpen
  Exit Function
  
FindEecordEx_Err:
  FindRecordEx = False
  Resume FindEecordEx_End
End Function


' Is this column NoPrint in the Grid
Public Sub SetColumnNoPrint(aCols As Collection, AGrid As AutoGrid)
  Dim ac As AutoCol, dGrid As TrueDBGrid60.TDBGrid
  Dim ctrlDAO As Object, ctrlRDO As Object
  Dim ColSet As TrueDBGrid60.Column
    
  If Not AGrid Is Nothing Then Set dGrid = AGrid.TDBGrid
  For Each ac In aCols
    ac.NoPrint = False
    If Not ac.Hide And Not (dGrid Is Nothing) Then
      If ac.GridColumn >= 0 Then
        Set ColSet = dGrid.Columns.Item(ac.GridColumn)
        ac.NoPrint = (ColSet.Width < GRID_MINCOLWIDTH) Or Not ColSet.visible
      End If
    End If
  Next ac
End Sub

Public Function SetRDOControl(rdc As MSRDC.MSRDC, ByVal rsRDO As RDOResultset) As Boolean
  On Error GoTo SetRDOControl_err
  
  Set rdc.Resultset = rsRDO
  rdc.Refresh
  If rsRDO.Type <> rdOpenStatic Then Err.Raise ERR_RDOSETUP, "InitAutoDataEX", "The resultset must use a Static Cursor"
  SetRDOControl = True
  
SetRDOControl_end:
  Call rdoEngine.rdoErrors.Clear
  Exit Function
  
SetRDOControl_err:
  Call ErrorMessage(ERR_ERROR, Err, "SetRDOControl", "Set Remote Data Control", "Unable to setup Remote data control.")
  Resume SetRDOControl_end
End Function

Public Function SetRDOControlSQL(ByVal cn As rdoConnection, rdc As MSRDC.MSRDC, rsRDO As RDOResultset, ByVal DerivedSQL As String, ByVal DefaultSQL As String) As Boolean
  Dim rsDerived As RDOResultset
  Dim UseDerived As Boolean
  Dim s As String
  Dim vType, vLockType

  On Error GoTo SetRDOControl_err
  UseDerived = True
  vType = rsRDO.Type
  vLockType = rsRDO.LockType
  If Not rsRDO Is Nothing Then
    If StrComp(rsRDO.Name, Right$(DerivedSQL, Len(rsRDO.Name)), vbTextCompare) = 0 Then GoTo SetRDO_ValidRS
  End If
  Set rsDerived = cn.OpenResultset(DerivedSQL, vType, vLockType)
  Set rsRDO = rsDerived
  GoTo SetRDO_ValidRS
  
SetRDO_UseDefault:
  UseDerived = False
  If Not rsRDO Is Nothing Then
    If StrComp(rsRDO.Name, Right$(DefaultSQL, Len(rsRDO.Name)), vbTextCompare) = 0 Then GoTo SetRDO_ValidRS
  End If
  Set rsRDO = Nothing
  Set rsRDO = cn.OpenResultset(DefaultSQL, vType, vLockType)

SetRDO_ValidRS:
  UseDerived = False
  Set rdc.Resultset = rsRDO
  'apf rdc.Refresh
  SetRDOControlSQL = True
  
SetRDOControl_end:
  Exit Function
  
SetRDOControl_err:
  SetRDOControlSQL = False
  If (vType = rdOpenForwardOnly) Or (vType = rdOpenStatic) Then s = "The resultset must use a Dynamic or Keyset Cursor"
  Call ErrorMessage(ERR_ERROR, Err, "SetRDOControl", "Set Remote Data Control", "Unable to setup Remote data control." & vbCrLf & s)
  If UseDerived Then Resume SetRDO_UseDefault
  Resume SetRDOControl_end
  Resume
End Function

Public Function AbsXCoord(ByVal X As Single) As String
  AbsXCoord = "{xabs=" & FormatNumber(X, 1, vbTrue, , vbFalse) & "}"
End Function

Public Function XCoord(ByVal X As Single) As String
  XCoord = FormatNumber(X, 0, vbTrue, , vbFalse)
End Function

Public Function GenerateSortFilterSQL(ByVal sql As String, ByVal sFilter As String, ByVal sSort As String) As String
  Dim p2 As Long, p1 As Long, p0 As Long
  
  If Len(sFilter) > 0 Or Len(sSort) > 0 Then
    p0 = InStr(1, sql, "ORDER", vbTextCompare)
    If p0 > 0 Then
      p2 = -1
      p1 = NotInStr(sql, " ", p0 + Len("ORDER"))
      If p1 > 0 Then p2 = InStr(p1, sql, "BY", vbTextCompare)
      If p2 = p1 Then sql = Left$(sql, p0 - 1)
    End If
    GenerateSortFilterSQL = "SELECT * from (" & sql & ") " & GEN_TABLE_NAME
    If Len(sFilter) > 0 Then GenerateSortFilterSQL = GenerateSortFilterSQL & " WHERE " & sFilter
    If Len(sSort) > 0 Then GenerateSortFilterSQL = GenerateSortFilterSQL & " ORDER BY " & sSort
  Else
    GenerateSortFilterSQL = sql
  End If
End Function

Public Function MergeFilterString(FilterString As String, AddFilterString As String) As String
  Dim s1 As String
  
  If Len(FilterString) > 0 Then
    s1 = "(" & FilterString & ") AND " & AddFilterString
  Else
    s1 = AddFilterString
  End If
  MergeFilterString = s1
End Function

Public Function MergeSortString(SortString As String, AddSortString As String) As String
  Dim s1 As String, sElement As String
  Dim p As Long, p0 As Long, p1 As Long, bAsc As Boolean
  
  If Len(SortString) > 0 Then
    p = 1
    Do
      p0 = InStr(p, SortString, ",")
      If p0 = 0 Then
        sElement = Mid$(SortString, p)
      Else
        sElement = Mid$(SortString, p, p0 - p)
      End If
      p1 = InStr(1, sElement, " ASC", vbTextCompare)
      bAsc = p1 <> 0
      sElement = Trim$(Left$(sElement, Len(sElement) - 4)) ' remove ASC/DESC
      If InStr(1, AddSortString, sElement, vbTextCompare) = 0 Then
        s1 = s1 & sElement & IIf(bAsc, " ASC", " DESC") & ","
      End If
      p = p0 + 1
    Loop Until p0 = 0
    s1 = s1 & AddSortString
  Else
    s1 = AddSortString
  End If
  MergeSortString = s1
End Function

Public Function GetGridTypedValueDefault(v As Variant, ByVal dType As DATABASE_FIELD_TYPES, DefaultValue As Variant) As Variant
  If (dType = TYPE_DATE) And IsDate(v) Then
    GetGridTypedValueDefault = CDate(v)
  Else
    GetGridTypedValueDefault = GetTypedValueDefault(v, dType, DefaultValue)
  End If
End Function

Public Function IsZero(ByVal d As Double) As Boolean
  Const ALMOST_ZERO As Double = 0.00000001
  
  d = Abs(d)
  IsZero = (d <= ALMOST_ZERO)
End Function

Public Function WhereClause(ByVal FieldName As String, ByVal Value As Variant, ByVal dType As DATABASE_FIELD_TYPES, ByVal UseLike As Boolean, Optional ByVal TextValue As Variant, Optional ByVal ForceVBExpressions As Boolean = False)
  Dim sFilter As String, UseTextValue As Boolean
  
  On Error GoTo WhereClause_err
  UseTextValue = False
  
redo_usingText:
  If UseTextValue Then Value = TextValue
  If DatabaseTarget = DB_TARGET_JET Or DatabaseTarget = DB_TARGET_SQLSERVER Then
    sFilter = "[" & FieldName & "]"
  Else
    sFilter = GEN_TABLE_NAME & "." & FieldName
  End If
  If UseLike And (dType = TYPE_STR) Then
    sFilter = sFilter & " " & LikeFilterDelimit(FILTER_START, ForceVBExpressions) & Trim$(Value) & LikeFilterDelimit(FILTER_END, ForceVBExpressions)
  Else
    If Len(Value) = 0 Then
      If dType = TYPE_STR Then
        sFilter = "((" & sFilter & "=" & GetSQLValue(Value, dType) & ") OR (" & sFilter & " is Null))"
      Else
        sFilter = sFilter & " is Null"
      End If
    Else
      sFilter = sFilter & "=" & GetSQLValue(Value, dType)
    End If
  End If
  WhereClause = sFilter
  Exit Function
  
WhereClause_err:
  If IsMissing(TextValue) Or UseTextValue Then Err.Raise Err.Number, ErrorSource(Err, "WhereClause"), Err.Description
  UseTextValue = True
  Resume redo_usingText
End Function

Public Function GetDropDownTypeAsString(ByVal dType As DROPDOWN_TYPE) As String
  GetDropDownTypeAsString = ""
  If (dType And DROPDOWN_QUERY) = DROPDOWN_QUERY Then GetDropDownTypeAsString = GetDropDownTypeAsString & "QUERY+"
  If (dType And DROPDOWN_BOUND) = DROPDOWN_BOUND Then GetDropDownTypeAsString = GetDropDownTypeAsString & "BOUND+"
  If (dType And DROPDOWN_STATIC) = DROPDOWN_STATIC Then GetDropDownTypeAsString = GetDropDownTypeAsString & "STATIC+"
  If (dType And DROPDOWN_DYNAMIC) = DROPDOWN_DYNAMIC Then GetDropDownTypeAsString = GetDropDownTypeAsString & "DYNAMIC+"
  If (dType And DROPDOWN_LIST) = DROPDOWN_LIST Then GetDropDownTypeAsString = GetDropDownTypeAsString & "LIST+"
  If (dType And DROPDOWN_COMBO) = DROPDOWN_COMBO Then GetDropDownTypeAsString = GetDropDownTypeAsString & "COMBO+"
  If Len(GetDropDownTypeAsString) > 0 Then GetDropDownTypeAsString = Left$(GetDropDownTypeAsString, Len(GetDropDownTypeAsString) - 1)
End Function

Public Function IsComboOpen(ByVal hwnd As Long) As Boolean
  Const WM_ISCOMBOOPEN As Long = WM_USER + 203
  IsComboOpen = SendMessage(hwnd, WM_ISCOMBOOPEN, 0, 0)
End Function

Public Function IsValueEqual(ByVal v0 As Variant, ByVal v1 As Variant) As Boolean
  If IsNull(v0) Then
    IsValueEqual = IsNull(v1)
  ElseIf IsNull(v1) Then
    IsValueEqual = IsNull(v0)
  Else
    IsValueEqual = (v0 = v1)
  End If
End Function
        
Public Function GetTypedValueNull(ByVal v0 As Variant, ByVal dType As DATABASE_FIELD_TYPES) As Variant
  Dim vDefault As Variant
  
  If (dType = TYPE_BOOL) Or (dType = TYPE_LONG) Or (dType = TYPE_DOUBLE) Then
    vDefault = 0
  ElseIf (dType = TYPE_DATE) Then
    vDefault = UNDATED
  Else
    vDefault = ""
  End If
  GetTypedValueNull = GetTypedValueDefault(v0, dType, vDefault)
End Function

'FIX CAD1
Public Function ColorHexToLong(ByVal shexcolor As String, ByVal DefaultValue As Long) As Long
  On Error GoTo ColorHexToLong_ERR
  If Len(shexcolor) = 0 Then Exit Function
  shexcolor = Replace(shexcolor, "#", "")
  ColorHexToLong = CLng("&H" & shexcolor)
  Exit Function
  
ColorHexToLong_ERR:
  ColorHexToLong = DefaultValue
  Exit Function
End Function



