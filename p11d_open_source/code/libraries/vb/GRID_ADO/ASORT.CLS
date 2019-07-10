VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AutoSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Columns As Variant
Public MinCol As Long
Public MaxCol As Long
Implements ISortFunction

'   Compare 2 sort items, return -1 if v0 < v1, 1 if v0 > v1, 0 if the same
Private Function ISortFunction_CompareItems(v0 As Variant, v1 As Variant) As Long
  Dim aCol As AutoCol, i As Long
  
  ISortFunction_CompareItems = 0
  For i = MinCol To MaxCol
    Set aCol = Columns(i)
    Select Case aCol.SortType
      Case SORT_NONE
      Case SORT_ASCENDING, SORT_DESCENDING
        ISortFunction_CompareItems = SortElement(v0(i), v1(i), aCol.dbDataType, aCol.SortType = SORT_ASCENDING)
      Case Else
        Call ECASE("AutoSort Invalid sort type")
    End Select
    If ISortFunction_CompareItems <> 0 Then Exit Function
  Next i
End Function

Private Function SortElement(ByVal v0 As Variant, ByVal v1 As Variant, ByVal dType As DATABASE_FIELD_TYPES, ByVal SortAsc As Boolean) As Long
  Select Case dType
    Case TYPE_BOOL, TYPE_LONG, TYPE_DOUBLE, TYPE_DATE
      SortElement = (v0 - v1)
    Case TYPE_STR
      SortElement = StrComp(v0, v1, vbTextCompare)
    Case Else
      Call ECASE("SortElement Invalid data type")
  End Select
  If Not SortAsc Then SortElement = SortElement * -1
End Function

