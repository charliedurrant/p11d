Attribute VB_Name = "OList1"
Option Explicit
'Public Event UpdateTime( ByVal dblJump As Double )
'
'Public Sub UnboundObjectListRead(oList As ObjectList, RowBuf As TrueDBGrid50.RowBuffer, StartLocation As Variant, ByVal Offset As Long, ApproximatePosition As Long)
'  Dim Index As Long, rCount As Long
'  Dim i As Long
'
'  'this gets the fist object list index
'  Index = GetObjectListIndex(oList, RowBuf, StartLocation, Offset)
'  If Index < 0 Then Exit Sub
'
'  rCount = 0
'  Do While (Index > 0) And (rCount < RowBuf.RowCount)
'    ' Fill Row Buffer
'    RowBuf.Bookmark(rCount) = CStr(Index)
'    ' (RowBuf.ColumnCount = 0) And (RowBuf.RowCount = 1) nothing else todo
'
'    rCount = rCount + 1
'    Index = GetObjectListIndex(oList, RowBuf, Index, 1)
'  Loop
'  RowBuf.RowCount = rCount
'  'ApproximatePosition
'End Sub
'
'
'Public Function GetObjectListIndex(oList As ObjectList, RowBuf As TrueDBGrid50.RowBuffer, StartLocation As Variant, ByVal Offset As Long) As Long
'  Dim FirstObject As Long
'  On Error GoTo GetObjectListIndex_Err
'
'  Call xSet("GetObjectListIndex")
'  If IsNull(StartLocation) Then
'    If Offset > 0 Then
'      FirstObject = GetObjectListOffset(oList, 0, 1)
'      Offset = Offset - 1
'      If FirstObject < 0 Then RowBuf.RowCount = 0
'    ElseIf Offset < 0 Then
'      FirstObject = GetObjectListOffset(oList, oList.count + 1, -1)
'    Else
'      GetObjectListIndex = -1
'      Exit Function
'    End If
'    If FirstObject < 0 Then Exit Function
'  Else
'    FirstObject = StartLocation
'  End If
'  GetObjectListIndex = GetObjectListOffset(oList, FirstObject, Offset)
'
'GetObjectListIndex_End:
'  Call xReturn("GetObjectListIndex")
'  Exit Function
'
'GetObjectListIndex_Err:
'  GetObjectListIndex = -1
'  Resume GetObjectListIndex_End
'End Function
'
'Public Function GetObjectListOffset(oList As ObjectList, ByVal Start As Long, ByVal Offset As Long) As Long
'  Dim i As Long, xOffset As Long
'
'  GetObjectListOffset = Start
'  If Offset > 0 Then
'    xOffset = 0: GetObjectListOffset = -1
'    For i = Start + 1 To oList.count
'      If Not oList(i) Is Nothing Then
'        xOffset = xOffset + 1
'        If Offset = xOffset Then
'          GetObjectListOffset = i
'          Exit For
'        End If
'      End If
'    Next i
'  ElseIf Offset < 0 Then
'    Offset = Offset * -1
'    xOffset = 0: GetObjectListOffset = -1
'    For i = (Start - 1) To 1 Step -1
'      If Not oList(i) Is Nothing Then
'        xOffset = xOffset + 1
'        If Offset = xOffset Then
'          GetObjectListOffset = i
'          Exit For
'        End If
'      End If
'    Next i
'  End If
'End Function
'
