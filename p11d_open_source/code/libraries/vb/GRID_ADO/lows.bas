Attribute VB_Name = "lows"
Option Explicit

Public Enum FILTER_DELIMIT_TYPE
  FILTER_NONE = 0
  FILTER_LIKE = 2
  FILTER_LIKE_ALL = 4
  FILTER_START = 128
  FILTER_END = 256
End Enum

Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const WM_USER As Long = &H400

Public Const SM_CXHSCROLL As Long = 21
Public Const SM_CYHSCROLL As Long = 3

Public Function GetAColByGridIndexEx(ByVal GridColumn As Integer, acol As AutoCol, ByVal ac As AutoClass) As Boolean
  Dim i As Long
  
  If GridColumn >= 0 Then
    For i = 1 To ac.Count
      Set acol = ac.Item(i)
      If acol.GridColumn = GridColumn Then
        GetAColByGridIndexEx = True
        Exit Function
      End If
    Next i
    Set acol = Nothing
  End If
End Function

Public Function LikeFilterDelimit(ByVal FilterType As FILTER_DELIMIT_TYPE) As String
  If (FilterType = FILTER_NONE) Or (((FilterType And FILTER_START) = FILTER_START) And ((FilterType And FILTER_END) = FILTER_END)) Then Err.Raise ERR_FILTER, "LikeFilterDelimit", "Invalid filter type [" & FilterType & "]"
  If (FilterType And FILTER_START) = FILTER_START Then
    If DatabaseTarget = DB_TARGET_JET Then
      LikeFilterDelimit = "LIKE '"
      If (FilterType And FILTER_LIKE_ALL) = FILTER_LIKE_ALL Then LikeFilterDelimit = LikeFilterDelimit & "*"
    Else
      LikeFilterDelimit = "LIKE '"
      If (FilterType And FILTER_LIKE_ALL) = FILTER_LIKE_ALL Then LikeFilterDelimit = LikeFilterDelimit & "%"
    End If
  Else
    If DatabaseTarget = DB_TARGET_JET Then
      LikeFilterDelimit = "*'"
    Else
      LikeFilterDelimit = "%'"
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

Public Function FindRecordEx(ByVal AGrid As AutoGrid, ByVal rs As Recordset, ByVal FindString As String, Optional ByVal FindType As FIND_TYPES = FT_FINDFIRST, Optional ByVal dbType As DATABASE_FIELD_TYPES = TYPE_STR) As Boolean
  Dim FieldName As String
  Dim vbmk As Variant
  
  On Error GoTo FindEecordEx_Err
  If Not rs Is Nothing Then
    If Not (rs.EOF Or rs.BOF) Then vbmk = rs.Bookmark
    Select Case FindType
      Case FT_FINDFIRST
        rs.MoveFirst
        rs.Find FindString, , adSearchForward
      Case FT_FINDNEXT
        rs.Find FindString, 1, adSearchForward
      Case FT_FINDPREVIOUS
        rs.Find FindString, -1, adSearchBackward
      Case FT_FINDLAST
        rs.MoveLast
        rs.Find FindString, , adSearchBackward
      Case Else
        Call ECASE("FindRecordEx")
    End Select
    FindRecordEx = Not (rs.EOF Or rs.BOF)
  End If
  
FindEecordEx_End:
  If (Not FindRecordEx) And (Len(vbmk) > 0) Then
    rs.Bookmark = vbmk
    Beep
  End If
  Exit Function
  
FindEecordEx_Err:
  FindRecordEx = False
  Resume FindEecordEx_End
End Function

' Is this column NoPrint in the Grid
' apf fixed printing across splits
Public Sub SetColumnNoPrint(ByVal aCols As Collection, ByVal AGrid As AutoGrid)
  Dim ac As AutoCol, dGrid As TrueOleDBGrid60.TDBGrid
  Dim ColSet As TrueOleDBGrid60.Column
  Dim k As Long
    
  If Not AGrid Is Nothing Then Set dGrid = AGrid.TDBGrid
  For Each ac In aCols
    ac.DerivedNoPrint = ac.NoPrint
    If Not ac.NoPrint And Not ac.Hide And Not (dGrid Is Nothing) Then
      If ac.GridColumn >= 0 Then
        For k = 0 To (dGrid.Splits.Count - 1)
          dGrid.Split = k
          Set ColSet = dGrid.Columns.Item(ac.GridColumn)
          If ColSet.Visible Then Exit For
        Next k
        ac.DerivedNoPrint = (ColSet.Width < GRID_MINCOLWIDTH) Or Not ColSet.Visible
      End If
    End If
  Next ac
End Sub

Public Function AbsXCoord(ByVal x As Single) As String
  AbsXCoord = "{xabs=" & FormatNumber(x, 1, vbTrue, , vbFalse) & "}"
End Function

Public Function XCoord(ByVal x As Single) As String
  XCoord = FormatNumber(x, 0, vbTrue, , vbFalse)
End Function

Public Function MergeFilterString(FilterString As String, AddFilterString As String) As String
  Dim s1 As String, p0 As Long
  
  If Len(FilterString) > 0 Then
    p0 = InStr(1, FilterString, " OR ", vbTextCompare)
    If p0 > 0 Then
      If InStr(1, AddFilterString, " OR ", vbTextCompare) > 0 Then Err.Raise ERR_FILTER, "MergeFilterString", "Cannot use exclude filter twice in this context" & vbCrLf & "Original filter " & FilterString & vbCrLf & "Additional filter " & AddFilterString
      s1 = "(" & Mid$(FilterString, 1, p0 - 1) & " AND (" & AddFilterString & ")) OR (" & Mid$(FilterString, p0 + Len(" OR ")) & " AND (" & AddFilterString & "))"
    Else
      s1 = "(" & FilterString & ") AND " & AddFilterString
    End If
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

Public Function WhereClause(ByVal FieldName As String, ByVal Value As Variant, ByVal dType As DATABASE_FIELD_TYPES, ByVal FilterType As FILTER_DELIMIT_TYPE, Optional ByVal TextValue As Variant, Optional ByVal Equal As Boolean = True)
  Dim sFilter As String, UseTextValue As Boolean
  
  On Error GoTo WhereClause_err
  UseTextValue = False
  
redo_usingText:
  If UseTextValue Then Value = TextValue
  sFilter = "[" & FieldName & "]"
  If (FilterType <> FILTER_NONE) And (dType = TYPE_STR) Then
    If Not Equal Then sFilter = "NOT " & sFilter
    sFilter = sFilter & " " & LikeFilterDelimit(FilterType + FILTER_START) & Trim$(Value) & LikeFilterDelimit(FilterType + FILTER_END)
  Else
    If Equal Then
      sFilter = sFilter & "="
    Else
      sFilter = sFilter & "<>"
    End If
    If Len(Value) = 0 Then
      sFilter = sFilter & "NULL"
    Else
      If DatabaseTarget = DB_TARGET_ORACLE And dType = TYPE_DATE Then
        sFilter = sFilter & "#" & Format$(Value, "DD/MM/YYYY") & "#"
      Else
        sFilter = sFilter & GetSQLValue(Value, dType)
      End If
    End If
  End If
  If (Not Equal) And (Len(Value) > 0) Then sFilter = "([" & FieldName & "] = NULL)" & " OR (" & sFilter & ")"
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




